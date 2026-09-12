(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 这是任务范围的事实账本，不是模型的自我报告。页面始终按 Office 返回的 id 归档，
  // index 只用于定位和展示，避免删页/插页后的页序变化把状态写到另一页。
  const WHOLE_DECK = /(?:整(?:份|套|个)|整稿|全稿|整个\s*(?:deck|ppt|演示文稿)|全部(?:页面|幻灯片)|所有(?:页面|页|幻灯片)|整体(?:整理|重排|改版|重组))/i;
  const DESIGN = /排版|版式|布局|视觉|设计|重排|重组|重构|改版|整理|层级|重点|阅读(?:路径|结构)?/i;
  const EXPLICIT_PAGES = /第\s*\d+\s*(?:页|张|slide)|\d+\s*(?:[-~到至]\s*\d+\s*)?(?:页|张|slides?)/ig;
  const NAMED_PAGE = /第\s*(\d+)\s*(?:页|张|slide).{0,20}?(?:必须|一定|重点|优先|要)(?:改|修改|重排|调整|处理)/ig;
  const LOCAL_ONLY = /(?:只|仅|只是|单独).{0,10}(?:改|修改|调整|处理|重排).{0,16}(?:第\s*\d+\s*(?:页|张|slide)|\d+\s*(?:[-~到至]\s*\d+\s*)?(?:页|张|slides?))/i;
  // 只拦真正的只读/不执行意图。“不要删除 Logo”“不改文案”是执行约束，不能把整稿任务降成只读。
  const READ_ONLY = /(?:只读|仅(?:做|给)?(?:建议|分析|方案)|只(?:给|出).{0,4}建议|先(?:给)?(?:建议|分析)|暂不执行|(?:不要|不)执行|(?:不要|不)写入|(?:不要|不)做任何修改|(?:不要|不)修改任何内容)/i;
  const PROPOSAL = /(?:修改)?提案|确认卡|diff\s*卡|修改卡/i;

  function textOf(value) { return String(value || '').trim(); }
  function wholeDeckRequest(value) {
    const text = textOf(value);
    const whole = WHOLE_DECK.test(text);
    const design = DESIGN.test(text);
    const local = EXPLICIT_PAGES.test(text) || LOCAL_ONLY.test(text);
    EXPLICIT_PAGES.lastIndex = 0;
    // “第 2 页”一类局部请求不能被全稿词误扩张；但明确说全稿且点名某页时仍是整稿。
    return { whole: whole && design, localOnly: LOCAL_ONLY.test(text) || (local && !whole), active: whole && design && !READ_ONLY.test(text), proposal: PROPOSAL.test(text), text };
  }

  function namedRequiredIndexes(value) {
    const indexes = [];
    const text = textOf(value);
    let match;
    while ((match = NAMED_PAGE.exec(text))) indexes.push(Number(match[1]) - 1);
    NAMED_PAGE.lastIndex = 0;
    return [...new Set(indexes.filter(index => Number.isInteger(index) && index >= 0))];
  }

  function safeSlides(outline) {
    if (!outline || outline.success === false || outline.hasMore || !Array.isArray(outline.slides)) return null;
    const slideCount = Number(outline.slideCount);
    if (!Number.isInteger(slideCount) || slideCount < 0 || outline.slides.length !== slideCount) return null;
    const ids = new Set();
    for (const slide of outline.slides) {
      if (!slide || slide.id == null || !Number.isInteger(Number(slide.index))) return null;
      const id = String(slide.id);
      if (ids.has(id)) return null;
      ids.add(id);
    }
    return outline.slides;
  }

  function makeLedger(request, outline) {
    const scope = wholeDeckRequest(request);
    if (!scope.whole || scope.localOnly) return null;
    const slides = safeSlides(outline);
    const ledger = {
      version: 1,
      kind: 'powerpoint-whole-deck',
      requestedAt: Date.now(),
      request: scope.text.slice(0, 300),
      status: slides ? 'active' : 'uncertain',
      scopeConfirmed: Boolean(slides),
      scopeError: slides ? '' : '无法取得完整且未截断的演示文稿大纲，不能验证整稿覆盖。',
      slideCount: slides ? slides.length : null,
      slides: [],
      failures: [],
      completion: null
    };
    const named = new Set(namedRequiredIndexes(scope.text));
    for (const slide of slides || []) {
      ledger.slides.push({
        id: String(slide.id), index: Number(slide.index), initialIndex: Number(slide.index),
        read: false, previewed: false, applied: false, verified: false, retainedReason: '',
        namedRequired: named.has(Number(slide.index)), failures: []
      });
    }
    ledger.indexMapCurrent = true;
    return ledger;
  }

  function find(ledger, id) {
    if (!ledger || id == null) return null;
    return (ledger.slides || []).find(slide => String(slide.id) === String(id)) || null;
  }
  function byIndex(ledger, index) {
    const n = Number(index);
    return Number.isInteger(n) ? (ledger.slides || []).find(slide => Number(slide.index) === n) || null : null;
  }
  function refreshOutline(ledger, outline) {
    const slides = safeSlides(outline);
    if (!ledger || !slides) {
      if (ledger) { ledger.status = 'uncertain'; ledger.scopeConfirmed = false; ledger.indexMapCurrent = false; ledger.scopeError = '后续大纲不完整或已截断，不能继续证明整稿覆盖。'; }
      return false;
    }
    const seen = new Set();
    for (const item of slides) {
      const record = find(ledger, item.id);
      if (record) { record.index = Number(item.index); seen.add(record.id); }
    }
    // 只可恢复曾由完整大纲确认过的同一组真实 ID；未知初始范围或增删页都不能借此假造覆盖。
    const hadConfirmedSet = Number.isInteger(ledger.slideCount) && ledger.slideCount === ledger.slides.length;
    if (!hadConfirmedSet || slides.length !== ledger.slides.length || seen.size !== ledger.slides.length) {
      ledger.status = 'uncertain'; ledger.scopeConfirmed = false;
      ledger.indexMapCurrent = false;
      ledger.scopeError = '演示文稿页面集合已变化，原整稿范围无法完整对应。';
      return false;
    }
    ledger.indexMapCurrent = true;
    ledger.scopeConfirmed = true;
    if (ledger.status === 'uncertain') ledger.status = 'active';
    ledger.scopeError = '';
    // 成功重取完整大纲只能解决此前同一读取动作的失败；绝不能冲掉写入或提案失败。
    ledger.failures = ledger.failures.filter(entry => !(entry.kind === 'read' && entry.name === 'get_presentation_outline'));
    return true;
  }
  function markStructureChanged(ledger) { if (ledger) ledger.indexMapCurrent = false; }
  function normalizedActionValue(value) {
    if (value == null || typeof value !== 'object') return value;
    if (Array.isArray(value)) return value.map(normalizedActionValue);
    const copy = {};
    for (const key of Object.keys(value).sort()) {
      // expected 是随实时文档读取而刷新的前置快照，不是模型请求的修改内容。
      if (key !== 'expected') copy[key] = normalizedActionValue(value[key]);
    }
    return copy;
  }
  function actionKey(args, name) {
    const target = (args && args.target) || {};
    const slide = target.slideId ?? target.index ?? (args && (args.slideId ?? args.index));
    const shape = target.shapeId ?? (args && args.shapeId);
    const operation = args && args.operation;
    const findText = textOf(args && args.find);
    if (slide == null || (!shape && !operation && !findText)) return '';
    const desired = JSON.stringify(normalizedActionValue({
      operation: operation || null,
      find: args && args.find,
      replacement: args && (args.replacement ?? args.replace),
      placement: args && args.placement
    }));
    return [name || '', slide, shape || '', desired].join('|');
  }
  function recordRead(ledger, result) {
    if (!ledger || !result || result.success === false || result.id == null) return null;
    const slide = find(ledger, result.id);
    if (!slide) return null;
    slide.read = true;
    // 成功读取只能证明现在能读，不能证明此前写入已经成功；只解决读取失败。
    slide.failures = slide.failures.filter(entry => entry.kind !== 'read');
    if (slide.applied) slide.verified = true;
    if (Number.isInteger(Number(result.index))) slide.index = Number(result.index);
    return slide;
  }
  function recordWrite(ledger, args, result, name) {
    if (!ledger || !result || result.success === false) return null;
    const target = (args && args.target) || {};
    const id = result.id ?? result.slideId ?? target.slideId ?? (args && args.slideId);
    // 没有真实 id 时只接受当前已由真实 read/outline 同步过的 index；结构已不确定时不猜。
    const slide = id != null ? find(ledger, id) : (ledger.scopeConfirmed && ledger.indexMapCurrent ? byIndex(ledger, result.index ?? target.index ?? (args && args.index)) : null);
    if (!slide) return null;
    slide.applied = true;
    slide.verified = false; // 每次新写入都要重新读回，不能沿用上一次核验。
    slide.retainedReason = '';
    // 只有同一动作的成功重试才能解决对应写入/提案失败；未知动作宁可保留受限证据。
    const key = actionKey(args, name);
    if (key) slide.failures = slide.failures.filter(entry => !(entry.kind !== 'read' && entry.actionKey === key));
    return slide;
  }
  function recordVerification(ledger, snapshot) {
    const slide = recordRead(ledger, snapshot);
    if (slide && slide.applied) slide.verified = true;
    return slide;
  }
  function recordPreview(ledger, result) {
    if (!ledger || !result || result.success === false) return null;
    const slide = find(ledger, result.id);
    if (!slide) return null;
    slide.previewed = true;
    if (Number.isInteger(Number(result.index))) slide.index = Number(result.index);
    return slide;
  }
  function recordFailure(ledger, args, result, name, kind) {
    if (!ledger || !result || result.success !== false) return;
    const target = (args && args.target) || {};
    const slide = find(ledger, result.id ?? result.slideId ?? target.slideId ?? (args && args.slideId))
      || (ledger.scopeConfirmed && ledger.indexMapCurrent ? byIndex(ledger, target.index ?? (args && args.index)) : null);
    const entry = { at: Date.now(), name: String(name || ''), kind: kind || (String(name || '').startsWith('get_') ? 'read' : 'write'), actionKey: actionKey(args, name), error: textOf(result.error).slice(0, 240) || '工具未完成' };
    if (slide) slide.failures.push(entry); else ledger.failures.push(entry);
  }
  function retain(ledger, args) {
    if (!ledger || !['active', 'unresolved'].includes(ledger.status) || !ledger.scopeConfirmed) return { success: false, error: '整稿范围尚未确认，不能记录保留页。' };
    if (!args || args.decision !== 'retain') return { success: false, error: '只能记录 retain 决定。' };
    const slide = find(ledger, args && args.slideId);
    const reason = textOf(args && args.reason);
    if (!slide) return { success: false, error: 'slideId 不在本次真实整稿范围内。' };
    if (!slide.read) return { success: false, error: '记录保留前必须先实际读取该页。' };
    if (slide.namedRequired) return { success: false, error: '用户点名要求修改这一页，不能记录为保留。' };
    if (reason.length < 8) return { success: false, error: '保留理由过短，请说明该页为何无需改动。' };
    slide.retainedReason = reason.slice(0, 400);
    return { success: true, slideId: slide.id, retained: true };
  }
  function summary(ledger) {
    if (!ledger) return null;
    if (!ledger.scopeConfirmed) return { complete: false, uncertain: true, message: ledger.scopeError || '整稿范围未确认。', slides: [] };
    const slides = ledger.slides.map(slide => ({ ...slide }));
    const unhandled = slides.filter(slide => !slide.applied && !slide.retainedReason);
    const unverified = slides.filter(slide => slide.applied && !slide.verified);
    const failed = slides.filter(slide => slide.failures.length);
    const coverageComplete = !unhandled.length && !unverified.length && !failed.length && !ledger.failures.length;
    return {
      complete: coverageComplete, coverageComplete,
      uncertain: false, unhandled, unverified, failed, failures: ledger.failures.slice(), slides
    };
  }
  function finalInstruction(ledger) {
    const data = summary(ledger);
    if (!data) return '';
    if (data.uncertain) return `【整稿覆盖状态】${data.message}。不能声称已完整处理；请如实说明受限范围。`;
    const labels = list => list.map(slide => `第 ${Number(slide.index) + 1} 页`).join('、');
    return `【整稿覆盖状态，程序记录】已应用：${labels(data.slides.filter(s => s.applied)) || '无'}；已读且保留：${data.slides.filter(s => s.retainedReason).map(s => `第 ${Number(s.index) + 1} 页（${s.retainedReason}）`).join('；') || '无'}；已读回核对：${labels(data.slides.filter(s => s.verified)) || '无'}；已取得预览：${labels(data.slides.filter(s => s.previewed)) || '无'}；待处理：${labels(data.unhandled) || '无'}；待读回核对：${labels(data.unverified) || '无'}；失败/受限：${labels(data.failed) || (data.failures.length ? '有未定位失败' : '无')}。${data.coverageComplete ? '以上仅证明程序范围覆盖已闭合；预览可取得或几何无 high 问题都不等于视觉判断已通过。' : '程序范围覆盖未闭合，不能声称整稿完成，请补齐或如实总结。'}`;
  }
  function userSummary(ledger) {
    const data = summary(ledger);
    if (!data) return '';
    if (data.uncertain) return `本轮无法确认整稿范围，已保留已完成部分。${data.message}`;
    const labels = list => list.map(slide => `第 ${Number(slide.index) + 1} 页`).join('、');
    const applied = data.slides.filter(slide => slide.applied);
    const retained = data.slides.filter(slide => slide.retainedReason);
    if (!data.coverageComplete) {
      const parts = [`本轮未完成整稿。已处理 ${applied.length} 页`];
      if (retained.length) parts.push(`保留 ${labels(retained)}`);
      if (data.unhandled.length) parts.push(`未处理 ${labels(data.unhandled)}`);
      if (data.unverified.length) parts.push(`待读回核对 ${labels(data.unverified)}`);
      if (data.failed.length || data.failures.length) parts.push(`有失败或受限页面`);
      return `${parts.join('；')}。已保留当前结果，可继续处理剩余范围。`;
    }
    return `整稿范围已逐页处理或记录保留，并已读回核对。该记录只说明处理覆盖，仍需人工确认视觉效果。`;
  }

  const toolDefinition = {
    type: 'function', function: {
      name: 'record_ppt_scope_decision',
      description: 'Record a reasoned decision to retain one already-read slide in the current whole-deck redesign. This does not write, apply, or verify anything. It cannot mark a user-named required page as retained.',
      parameters: { type: 'object', properties: {
        slideId: { type: 'string', description: 'Exact slide id returned by get_slide.' },
        decision: { type: 'string', enum: ['retain'], description: 'Only retain is accepted.' },
        reason: { type: 'string', description: 'Specific reason this read slide should remain unchanged.' }
      }, required: ['slideId', 'decision', 'reason'] }
    }
  };

  App.pptTaskScope = { wholeDeckRequest, makeLedger, refreshOutline, markStructureChanged, recordRead, recordPreview, recordWrite, recordVerification, recordFailure, retain, summary, finalInstruction, userSummary, toolDefinition };
})();
