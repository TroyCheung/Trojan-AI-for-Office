(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ── 应用修改后的"指给你看"：目标位置短暂高亮（借鉴 Claude for Excel 的改动高亮）──
  //
  // 设计约束：
  // 1. 只服务 Word。Excel 已有选中跳转（_dirtyRanges follow）、PPT 已有版面复检，不重复做回执。
  // 2. 不阻塞应用流程：ui.js 调 queueEditFlash 后立即继续；内部串行队列防止批量应用时
  //    两次高亮互相把对方的还原值覆盖成黄色。
  // 3. 还原时恢复该范围原本的高亮色，而不是粗暴置空，避免清掉用户已有的标注。
  // 4. 一切失败静默：高亮是锦上添花，绝不能因为它报错打断应用链路。

  const FLASH_MS = 1800;
  const MAX_QUEUE = 5;
  let chain = Promise.resolve();
  let queued = 0;

  // 从提案项里提取可高亮的 Word 段落序号（纯函数，供测试与主流程共用）
  function collectFlashTargets(edit, result) {
    const out = [];
    const push = v => {
      const n = Number(v);
      if (Number.isInteger(n) && n >= 0 && !out.includes(n)) out.push(n);
    };
    push(edit && edit.target && edit.target.paragraphIndex);
    push(result && result.paragraphIndex);
    return out;
  }

  function flashSupported() {
    return Boolean(App.host && App.host.hostType === 'word'
      && typeof Word !== 'undefined' && Word && Word.run
      && typeof App.hasOffice === 'function' && App.hasOffice());
  }

  // 在同一个 Word.run 里完成「取原色 → 上黄 → 等待 → 还原」，
  // 中途不换 context，避免还原阶段重新定位范围。
  async function runFlash(resolveRange) {
    await Word.run(async context => {
      const range = await resolveRange(context);
      if (!range) return;
      range.load('font/highlightColor');
      await context.sync();
      const prev = range.font.highlightColor;
      range.font.highlightColor = 'yellow';
      await context.sync();
      await new Promise(resolve => setTimeout(resolve, FLASH_MS));
      try {
        range.font.highlightColor = prev;
        await context.sync();
      } catch {}
    });
  }

  // 优先高亮当前选区（followMode 开启时 applyEdit 刚用 followSelect 选中了被改文字，最精确）；
  // 否则退回段落序号整段高亮。
  function makeResolver(targets) {
    const useSelection = Boolean(App.state.settings.followMode);
    return async function resolveRange(context) {
      if (useSelection) {
        const sel = context.document.getSelection();
        sel.load('text');
        await context.sync();
        if (String(sel.text || '').trim()) return sel;
      }
      if (!targets.length) return null;
      const paras = context.document.body.paragraphs;
      paras.load('items');
      await context.sync();
      const paragraph = paras.items[targets[0]];
      return paragraph || null;
    };
  }

  function queueEditFlash(edit, result) {
    if (!flashSupported()) return;
    if (!(result && result.success !== false)) return;
    const targets = collectFlashTargets(edit, result);
    if (!targets.length && !App.state.settings.followMode) return;
    if (queued >= MAX_QUEUE) return; // 批量应用时每项本就有状态徽标反馈，超出即放弃
    queued += 1;
    chain = chain
      .then(() => runFlash(makeResolver(targets)))
      .catch(() => {})
      .then(() => { queued -= 1; });
  }

  App.collectFlashTargets = collectFlashTargets;
  App.queueEditFlash = queueEditFlash;
})();
