(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 无进展守卫（移植自 Office-AI-addin 本轮加固）：相同请求且相同结果才算重复；
  // 成功写入会开启新 epoch，读去重不跨写入继承。成功写入永不因相同参数被杀掉。
  // 分页、不同目标、新证据继续。公开版没有 host.js 的统一执行结果合同
  // （App.writeOutcome），isFailed 自动退回本地判定（success:false / blocked）。
  const READ_RE = /^(get_|search_|read_|web_search$|check_consistency$|extract_images$)/;
  // 提案装饰字段：模型改写它们重试同一操作时指纹必须不变（改 changes[].label
  // 重发同一 find/replace，指纹变化会绕过守卫）。find/replace/operation/target 等实质字段
  // 绝不进名单——它们变了就是新参数。忽略在 stable() 的对象分支做，任意深度生效。
  const IGNORE_KEYS = { explanation: true, label: true, summary: true };
  const PAGINATION_KEYS = ['start', 'count', 'cursor', 'nextCursor', 'offset', 'limit', 'maxResults', 'maxParagraphs', 'nextStart'];

  function stable(value) {
    if (value == null || typeof value !== 'object') return JSON.stringify(value);
    if (Array.isArray(value)) return '[' + value.map(stable).join(',') + ']';
    const keys = Object.keys(value).sort().filter(key => !IGNORE_KEYS[key]);
    return '{' + keys.map(k => JSON.stringify(k) + ':' + stable(value[k])).join(',') + '}';
  }

  function fingerprint(name, args) {
    const src = args && typeof args === 'object' && !Array.isArray(args) ? args : {};
    return String(name || '') + ':' + stable(src);
  }

  function resultFingerprint(result) {
    if (result == null || typeof result !== 'object') return stable(result);
    const semantic = Object.assign({}, result);
    // 这些字段由重试循环挂上，不是新的工具证据；否则每次重试的结果都"变了"，守卫永远数不到阈值。
    for (const key of ['consecutiveFailures', 'guidance', 'stopRetrying']) delete semantic[key];
    return stable(semantic);
  }

  // 宿主若有统一执行结果合同（App.writeOutcome）则共用；没有就退回本地判定。
  function sharedOutcome(name, result) {
    return (name && typeof App.writeOutcome === 'function') ? App.writeOutcome(name, result) : null;
  }
  function isFailed(result, name) {
    if (!result || typeof result !== 'object') return false;
    const outcome = sharedOutcome(name, result);
    if (outcome) return outcome === 'failed' || outcome === 'mismatch' || outcome === 'unverified' || outcome === 'not_written';
    return result.success === false || result.blocked === true;
  }

  function isReadTool(name) {
    return READ_RE.test(String(name || ''));
  }

  function isPaginating(prevArgs, args) {
    if (!prevArgs || !args) return false;
    for (let i = 0; i < PAGINATION_KEYS.length; i++) {
      const key = PAGINATION_KEYS[i];
      if (args[key] == null || prevArgs[key] == null) continue;
      if (key === 'cursor' || key === 'nextCursor') {
        if (String(args[key]) !== String(prevArgs[key])) return true;
        continue;
      }
      const next = Number(args[key]);
      const prev = Number(prevArgs[key]);
      if (Number.isFinite(next) && Number.isFinite(prev) && next !== prev) return true;
    }
    return false;
  }

  function isSuccessfulWrite(name, result) {
    if (!result || result.success !== true || isFailed(result, name) || isReadTool(name)) return false;
    if (result.declined || result.cancelled || result.verified === false || result.verified === null) return false;
    if (name === 'propose_edits') {
      return Number(result.applied) > 0 || (Array.isArray(result.results) && result.results.some(item => item && item.applied === true));
    }
    for (const key of ['replaced', 'applied', 'inserted', 'updated', 'deleted', 'matched']) {
      if (typeof result[key] === 'number' && result[key] === 0) return false;
    }
    if (typeof App.isDirectWriteTool === 'function') return App.isDirectWriteTool(name);
    return /^(?:insert_|add_|set_|replace_|delete_|update_|apply_|clear_|fill_|resize_|remove_|duplicate_|move_|merge_|sort_|format_|create_|write_range$|undo_last_write$)/.test(String(name || ''));
  }

  function create(options) {
    const identicalCorrectAt = 2;
    const identicalStopAt = 3;
    const counts = new Map();
    // 失败记忆（防「写成功→同一失败读」交替空转）：counts 以 writeEpoch 为前缀，写成功即清零，
    // 交替序列每个 epoch 都从 1 重数、永远到不了阈值。记忆按指纹跨 epoch 结转：成功写入时
    // 旧记忆 ceil(n/2) 衰减、本 epoch 失败全额并入；verdict 取 counts 与记忆之和。仅本 run 可见。
    const failureMemory = new Map();  // fp → 已结转的失败次数
    const epochFailures = new Map();  // fp → 本 epoch（自上次成功写入以来）的失败次数
    let writeEpoch = 0;

    function bump(key) {
      const n = (counts.get(key) || 0) + 1;
      counts.set(key, n);
      return n;
    }

    function foldFailureMemory() {
      const fps = new Set([...failureMemory.keys(), ...epochFailures.keys()]);
      for (const fp of fps) {
        const carried = Math.ceil((failureMemory.get(fp) || 0) / 2) + (epochFailures.get(fp) || 0);
        if (carried > 0) failureMemory.set(fp, carried); else failureMemory.delete(fp);
      }
      epochFailures.clear();
    }

    function repeatVerdict(name, n, kind) {
      const guidance = kind === 'failed'
        ? (name + ' 用相同参数连续失败且结果没有变化。不要重复同一调用，换路径或停下来告诉用户。')
        : kind === 'replay'
          ? (name + ' 上一次写入的结果没有确认（可能已经写入），同一调用被拦下了。不要原样重发：先用读取工具核对目标当前状态，确认确实没写入再决定下一步。')
          : (name + ' 用相同参数得到了相同结果，没有新证据。不要重复同一读取；若目标未完成，换段落、出现位置、分页游标，或执行写入。');
      const reason = kind === 'same' ? 'repeat' : (kind === 'replay' ? 'replay' : 'repeat-fail');
      if (n >= identicalStopAt) return { action: 'stop', reason, guidance };
      if (n === identicalCorrectAt) return { action: 'correct', reason, guidance };
      return { action: 'continue' };
    }

    function observe(name, args, result) {
      const failed = isFailed(result, name);
      if (isSuccessfulWrite(name, result)) {
        writeEpoch += 1;
        counts.clear();
        foldFailureMemory();
        return { action: 'continue', reason: 'write', writeEpoch };
      }
      const fp = fingerprint(name, args);
      if (failed) epochFailures.set(fp, (epochFailures.get(fp) || 0) + 1);
      const rfp = resultFingerprint(result);
      const key = writeEpoch + '|' + fp + '|' + rfp;
      // 阈值判定用本 epoch 计数 + 结转记忆之和：记忆只含成功写入之前的失败，
      // 与 counts 不重叠，相加即该指纹的累计失败量。
      const n = bump(key) + (failureMemory.get(fp) || 0);
      if (result && result.replayBlocked) return Object.assign({ fingerprint: fp, failed: true, replayBlocked: true }, repeatVerdict(name, n, 'replay'));
      if (failed) return Object.assign({ fingerprint: fp, failed: true }, repeatVerdict(name, n, 'failed'));
      return Object.assign({ fingerprint: fp }, repeatVerdict(name, n, 'same'));
    }

    function observeBatch(calls) {
      const list = Array.isArray(calls) ? calls : [];
      let worst = { action: 'continue' };
      for (let i = 0; i < list.length; i++) {
        const item = list[i] || {};
        const verdict = observe(item.name, item.args, item.result);
        if (verdict.action === 'stop') worst = verdict;
        else if (verdict.action === 'correct' && worst.action === 'continue') worst = verdict;
      }
      const wrote = list.some(item => isSuccessfulWrite(item && item.name, item && item.result));
      if (wrote && worst.reason === 'repeat') return { action: 'continue' };
      return worst;
    }

    return { observe, observeBatch, fingerprint, resultFingerprint };
  }

  App.ProgressGuard = {
    create,
    fingerprint,
    resultFingerprint,
    isFailed,
    isReadTool,
    isPaginating,
    isSuccessfulWrite
  };
})();
