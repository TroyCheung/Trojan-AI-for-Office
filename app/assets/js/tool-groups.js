(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ── 连续同构工具卡折叠（v135，用户 真机反馈：删 6 个元素出 6 张卡把对话撑爆）──
  // 纯逻辑层：分组 + 组状态 + 展开态计算。渲染组装在 ui.js（复用 toolCardHtml 做明细）。
  // 折叠只影响消息流呈现，不改 toolCalls 数据本身；会话级汇总仍走顶部「本次会话的修改记录」。

  const MIN_GROUP = 3; // 连续 3 次起才值得折，两张卡不构成噪音
  // 决策/交互类工具永不折叠：每张卡都是用户要做决定的入口（render/propose_edits 走专用卡，
  // ask_clarification 是拍板卡；它们即使出现在流里也必须单张可见）
  const NEVER_GROUP = new Set(['propose_edits', 'ask_clarification', 'render']);

  function group(list) {
    const items = Array.isArray(list) ? list : [];
    const out = [];
    let i = 0;
    while (i < items.length) {
      const tc = items[i];
      if (NEVER_GROUP.has(tc.name) || !tc.rid) { out.push({ type: 'single', tc }); i++; continue; }
      let j = i;
      while (j < items.length && items[j].name === tc.name && !!items[j].rid) j++;
      const run = items.slice(i, j);
      if (run.length >= MIN_GROUP) out.push({ type: 'group', items: run });
      else for (const t of run) out.push({ type: 'single', tc: t });
      i = j;
    }
    return out;
  }

  // 组状态：running 优先（有进行中就是 running）；有失败是 mixed；全终态成功是 done
  function status(items) {
    let running = 0, error = 0;
    for (const tc of items) {
      const s = tc.status || 'running';
      if (s === 'running') running++;
      else if (s === 'error') error++;
    }
    if (running > 0) return { key: 'running', running, done: items.length - running - error };
    if (error > 0) return { key: 'mixed', done: items.length - error, error };
    return { key: 'done', done: items.length };
  }

  // 展开判定：用户显式意图（_groupOpen）优先；否则有进行中默认展开（能看到当前做到哪），
  // 全部完成默认收起（安静）。点击 toggle 时 ui.js 按 isOpen 的当前值取反写回 _groupOpen。
  function isOpen(items) {
    const first = items[0];
    if (!first) return false;
    if (first._groupOpen === true) return true;
    if (first._groupOpen === false) return false;
    return status(items).key === 'running';
  }

  function toggle(items) {
    const first = items[0];
    if (!first) return;
    first._groupOpen = !isOpen(items);
  }

  App.toolGroups = { group, status, isOpen, toggle, MIN_GROUP, NEVER_GROUP };
})();
