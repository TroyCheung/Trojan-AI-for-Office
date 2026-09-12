(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 宿主提供者注册表。各 host-*.js 在加载时把自己注册进来。
  App.HOSTS = App.HOSTS || {};

  // ---- 通用 Office 文档设置存取（宿主无关，Excel/Word/PowerPoint 通用）----
  function saveDocSetting(key, value) {
    return new Promise((resolve, reject) => {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message || 'Failed to save document setting')));
    });
  }
  function loadDocSetting(key, fallback) {
    try { return Office.context.document.settings.get(key) || fallback; } catch { return fallback; }
  }
  async function getDocumentId() {
    if (!App.hasOffice()) return 'browser';
    return new Promise((resolve, reject) => {
      const settings = Office.context.document.settings;
      let v = settings.get(App.STORAGE_KEYS.workbookId);
      if (v) return resolve(v);
      v = App.id();
      settings.set(App.STORAGE_KEYS.workbookId, v);
      settings.saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(v) : reject(new Error(r.error?.message || 'Failed to save document ID')));
    });
  }
  function requireOffice() { if (!App.hasOffice()) throw new Error(App.t('demo')); }

  // ---- 【R1/30.2】文档写入操作生命周期（宿主共享）：取消代次 + 在途计数 ----
  // 此前只在 host-word.js 的 applyEdit 内实现；Excel/PPT 写入不受停止保护（Astra
  // 独立探针实证：Excel applyEdit 读回挂起期间停止，恢复后照样写入）。现在三宿主共用：
  //   - App.markDocWritesCancelled()：停止入口（ui stopActiveRequest）使全局代次 +1，
  //     单调递增、永不复位——新任务的发送/重置不能清除旧操作的取消事实（25.2 语义）。
  //   - App.docWriteBegin(label)：操作进入时捕获代次快照并在 App.state.pendingDocWrites
  //     计数（桥的停止协议据此等在途写入收敛，不提前 confirmed）；返回 { token, end }。
  //     token.throwIfCancelled() 由宿主在「预检后/真正提交前」调用（Excel 赋值前、
  //     PPT 定位后写入前）；已提交宿主的写入不假装撤回，按实际结果完成、finally 归零。
  // 运行归属由桥命令层保证（executeCommand 捕获 ctx.runId + sideEffectAllowed）；
  // 宿主层负责的是文档写入的代次与计数。
  let docWriteCancelGen = 0;
  App.markDocWritesCancelled = () => { docWriteCancelGen += 1; };
  App.docWriteCancelledReset = () => { /* 【25.2】已废弃：取消事实不可复位（保留空实现兼容既有调用） */ };
  App.docWriteBegin = function (label) {
    const startGen = docWriteCancelGen;
    const startStop = !!(App.state && App.state.stopRequested);
    let cancelled = false;
    const token = {
      label: String(label || 'doc-write'),
      throwIfCancelled() {
        // 粘性取消：本操作的标记优先；代次比较为主（stopRequested 布尔可被新任务
        // 重置，只作本操作开始后新发生停止的补充证据）。
        if (cancelled || docWriteCancelGen !== startGen || ((App.state && App.state.stopRequested) && !startStop)) {
          cancelled = true;
          const err = new Error('应用操作已被停止取消：本次写入不再执行。');
          err.code = 'WRITE_CANCELLED';
          throw err;
        }
      }
    };
    App.state.pendingDocWrites = (App.state.pendingDocWrites || 0) + 1;
    let ended = false;
    return {
      token,
      end() {
        if (ended) return;
        ended = true;
        App.state.pendingDocWrites = Math.max(0, (App.state.pendingDocWrites || 0) - 1);
      }
    };
  };


  // ---- 非 Office 环境 / 未识别宿主时的占位提供者 ----
  function makeDemoHost(hostType) {
    const unavailable = async () => { throw new Error(App.t('demo')); };
    return {
      hostType: hostType || 'none',
      available: false,
      metadataLabel: 'Document metadata',
      systemPrompt: 'You are an AI assistant integrated into Microsoft Office.',
      toolDefinitions: [],
      toolExecutors: {},
      defaultArgsForTool() { return '{}'; },
      evalToolName: 'eval_officejs',
      getMetadata: unavailable,
      navigateCitation: unavailable,
      follow: async () => {},
      i18n: { zh: {}, en: {} }
    };
  }

  // ---- 宿主检测与选择 ----
  function detectHostType() {
    if (typeof Excel !== 'undefined' && Excel.run) return 'excel';
    if (typeof Word !== 'undefined' && Word.run) return 'word';
    if (typeof PowerPoint !== 'undefined' && PowerPoint.run) return 'powerpoint';
    return null;
  }
  function hostTypeFromOffice(officeHost) {
    if (typeof Office === 'undefined' || !Office.HostType) return null;
    switch (officeHost) {
      case Office.HostType.Excel: return 'excel';
      case Office.HostType.Word: return 'word';
      case Office.HostType.PowerPoint: return 'powerpoint';
      default: return null;
    }
  }
  function selectHost(type) {
    const provider = type && App.HOSTS[type];
    App.host = provider || makeDemoHost(type);
    return App.host;
  }

  // ---- 稳定的委托封装：api.js / ui.js 在加载时取这些别名，调用时再转发到当前宿主 ----
  function executeToolByName(name, args) {
    const ex = App.host && App.host.toolExecutors;
    const fn = ex && ex[name];
    if (fn) return fn(args || {});
    // 宿主工具没有时回落到内置工具（如 web_search）
    const builtin = App.builtinTools && App.builtinTools[name];
    if (builtin) return builtin(args || {});
    throw new Error(`Tool ${name} not found`);
  }
  function maybeFollow(result) { return App.host && App.host.follow ? App.host.follow(result) : Promise.resolve(); }
  function navigateCitation(ref) { return App.host && App.host.navigateCitation ? App.host.navigateCitation(ref) : Promise.resolve(); }
  function getWorkbookMetadata() { return App.host.getMetadata(); }

  // ---- Diff 提案的跨宿主安全匹配：只兼容排版等价差异，不做可能误改内容的模糊匹配 ----
  const DOUBLE_QUOTES = /[“”„‟＂]/;
  const SINGLE_QUOTES = /[‘’‚‛＇]/;
  const ZERO_WIDTH = /[\u200B-\u200D\uFEFF]/;
  function normalizedTextWithMap(value) {
    const input = String(value == null ? '' : value);
    let text = '';
    const starts = [], ends = [];
    for (let i = 0; i < input.length; i++) {
      const ch = input[i];
      if (ZERO_WIDTH.test(ch)) continue;
      let mapped = DOUBLE_QUOTES.test(ch) ? '"' : SINGLE_QUOTES.test(ch) ? "'" : /\s|\u00A0|\u3000/.test(ch) ? ' ' : ch;
      if (mapped === ' ' && text.endsWith(' ')) { ends[ends.length - 1] = i + 1; continue; }
      text += mapped;
      starts.push(i);
      ends.push(i + 1);
    }
    return { text: text.trim(), starts, ends, leadingTrim: text.length - text.trimStart().length };
  }
  function normalizeComparableText(value) { return normalizedTextWithMap(value).text; }
  function findSafeTextMatch(haystack, needle) {
    const source = String(haystack == null ? '' : haystack);
    const query = String(needle == null ? '' : needle);
    if (!query) return { status: 'missing' };
    const exact = [];
    for (let at = source.indexOf(query); at >= 0; at = source.indexOf(query, at + Math.max(1, query.length))) exact.push({ start: at, end: at + query.length, text: query, exact: true });
    if (exact.length === 1) return { status: 'matched', ...exact[0] };
    if (exact.length > 1) return { status: 'ambiguous', count: exact.length };
    const h = normalizedTextWithMap(source);
    const q = normalizeComparableText(query);
    if (!q) return { status: 'missing' };
    const matches = [];
    for (let at = h.text.indexOf(q); at >= 0; at = h.text.indexOf(q, at + Math.max(1, q.length))) {
      const mappedAt = at + h.leadingTrim;
      const start = h.starts[mappedAt];
      const end = h.ends[mappedAt + q.length - 1];
      if (start != null && end != null) matches.push({ start, end, text: source.slice(start, end), exact: false });
    }
    if (matches.length === 1) return { status: 'matched', ...matches[0] };
    return matches.length ? { status: 'ambiguous', count: matches.length } : { status: 'missing' };
  }
  function makeStaleEditError(message, details = {}) {
    const error = new Error(message || '原文已变化，请重新读取后生成此项。');
    error.code = 'STALE_EDIT';
    Object.assign(error, details);
    return error;
  }
  function valuesEquivalent(a, b) {
    if (a == null || a === '') return b == null || b === '';
    if (b == null || b === '') return false;
    if (typeof a === 'string' || typeof b === 'string') return normalizeComparableText(a) === normalizeComparableText(b);
    return Object.is(a, b);
  }

  App.saveDocSetting = saveDocSetting;
  App.loadDocSetting = loadDocSetting;
  App.getDocumentId = getDocumentId;
  App.getWorkbookId = getDocumentId; // 向后兼容旧调用名
  App.requireOffice = requireOffice;
  App.makeDemoHost = makeDemoHost;
  App.detectHostType = detectHostType;
  App.hostTypeFromOffice = hostTypeFromOffice;
  App.selectHost = selectHost;
  App.executeToolByName = executeToolByName;
  App.maybeFollow = maybeFollow;
  App.navigateCitation = navigateCitation;
  App.getWorkbookMetadata = getWorkbookMetadata;
  App.normalizeComparableText = normalizeComparableText;
  App.findSafeTextMatch = findSafeTextMatch;
  App.makeStaleEditError = makeStaleEditError;
  App.valuesEquivalent = valuesEquivalent;

  // 默认先挂占位提供者，保证 App.host 在首屏渲染前一定存在；app.js 会在 Office.onReady 后重新选定。
  App.host = makeDemoHost(detectHostType());
})();
