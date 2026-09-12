(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ── 会话改动汇总卡（借鉴 Claude for Excel 的会话日志）──
  //
  // 设计约束：
  // 1. 纯函数模块：从 state.messages 的 propose_edits 工具调用里汇总全部提案与状态，
  //    不修改任何状态；ui.js 只负责把返回的 HTML 插进聊天面板顶部。
  // 2. 单卡（edits）与批量卡（changes）归一成同一条目结构，状态映射到七种展示态（含「已过期」）。
  // 3. 行内 ⌖ 点击跳转原文：复用各宿主 navigateCitation，定位规则与 Diff 卡一致。

  function normalizeBatchState(state, settled) {
    switch (state) {
      case 'applied': return 'applied';
      case 'skipped': return 'skipped';
      case 'stale': return 'stale';
      case 'error': return 'error';
      case 'refresh_requested': return 'refresh';
      case 'applying': return 'pending';
      case 'interrupted': return 'interrupted';   // 【35-R1】恢复的在途条目：未知写入结果，据实显示已中断
      default: return settled ? 'skipped' : 'pending';
    }
  }

  function normalizeCardDecision(decision) {
    switch (decision) {
      case 'applied': return 'applied';
      case 'declined': return 'skipped';
      case 'stale': return 'stale';
      case 'error': return 'error';
      case 'refresh_requested': return 'refresh';
      case 'interrupted': return 'interrupted';   // 【35-R1】同批量：恢复的在途卡据实显示已中断
      default: return 'pending'; // 含 null 与 'applying'
    }
  }

  // 与 ui.js 卡片的 expired 判定同款：插件重开后 Promise 解析器丢失，
  // 未决条目显示「已过期」而不是「待确认」，与卡片只读态对齐
  function cardExpired(tc, decided) {
    return !decided && !tc._resolveProposal && !(App.state && App.state.isWorking);
  }

  // 汇总本会话所有提案：[{rid, idx, label, status, target}]
  function buildChangeLogItems(messages) {
    const items = [];
    for (const m of Array.isArray(messages) ? messages : []) {
      if (!m || !Array.isArray(m.toolCalls)) continue;
      for (const tc of m.toolCalls) {
        if (!tc || tc.name !== 'propose_edits' || !tc.rid) continue;
        const args = tc.args || {};
        const changes = Array.isArray(args.changes) && args.changes.length ? args.changes : null;
        const edits = !changes && Array.isArray(args.edits) ? args.edits : null;
        if (changes) {
          const b = tc.batch || { items: changes.map(() => 'pending'), settled: false };
          const expired = cardExpired(tc, b.settled);
          changes.forEach((c, i) => {
            let status = normalizeBatchState(b.items[i] || 'pending', b.settled);
            if (expired && status === 'pending') status = 'expired';
            items.push({
              rid: tc.rid,
              idx: i,
              label: (c && c.label) || `#${i + 1}`,
              status,
              target: (c && c.target) || null
            });
          });
        } else if (edits && edits.length) {
          const p = tc.proposal || {};
          let decision = normalizeCardDecision(p.decision);
          // 与 ui.js 一致：stale/applying 不算终态，重开后同样归为已过期
          if (cardExpired(tc, decision !== 'pending' && decision !== 'stale')) decision = 'expired';
          // 定位跟随被应用的版本：拒绝/未决时看当前查看的版本
          const viewIdx = Math.min(Number(p.appliedIdx) || 0, edits.length - 1);
          items.push({
            rid: tc.rid,
            idx: viewIdx,
            label: (edits[viewIdx] && edits[viewIdx].label) || '修改提案',
            status: decision,
            target: (edits[viewIdx] && edits[viewIdx].target) || null
          });
        }
      }
    }
    return items;
  }

  // 定位引用串，规则与 ui.js 的 proposalCitation 一致（Word 段落 / PPT 页 / Excel 区域）
  function changeLogRef(target) {
    const tg = target && typeof target === 'object' ? target : null;
    if (!tg) return '';
    if (Number.isInteger(Number(tg.paragraphIndex))) {
      const index = Number(tg.paragraphIndex);
      return index >= 0 ? `p:${index}` : '';
    }
    if (tg.slideId != null) return `id:${tg.slideId}`;
    if (Number.isInteger(Number(tg.index))) {
      const index = Number(tg.index);
      return index >= 0 ? `s:${index}` : '';
    }
    const sheetId = Number(tg.sheetId);
    const range = String(tg.range || '').trim();
    return Number.isFinite(sheetId) && range ? `${sheetId}!${range}` : '';
  }

  const STATUS_KEYS = {
    applied: 'proposalApplied',
    skipped: 'proposalDeclined',
    stale: 'proposalStale',
    error: 'proposalApplyError',
    refresh: 'proposalRefreshQueued',
    expired: 'proposalExpired',
    interrupted: 'proposalInterruptedBadge',
    pending: 'proposalPending'
  };

  function changeLogHtml(items) {
    const esc = App.escapeHtml;
    const t = App.t;
    const open = Boolean(App.state.changeLogOpen);
    const applied = items.filter(x => x.status === 'applied').length;
    const head = `<button type="button" class="chg-toggle" data-action="toggle-changelog" aria-expanded="${open}">
      <span class="chg-ico">${App.icon ? App.icon('edit', 'sm') : ''}</span>
      <span>${esc(t('chgTitle'))}</span>
      <strong>${applied}/${items.length}</strong>
      <span class="chg-chev">${App.icon ? App.icon(open ? 'chevronDown' : 'chevronRight', 'sm') : ''}</span>
    </button>`;
    const rows = !open ? '' : `<div class="chg-list">${items.map((item, i) => {
      const ref = changeLogRef(item.target);
      const jump = ref ? `<button type="button" class="chg-jump" data-action="changelog-jump" data-idx="${i}" title="${esc(t('proposalLocate'))}">${App.icon ? App.icon('locate', 'sm') : ''}⌖</button>` : '';
      return `<div class="chg-item st-${item.status}">
        <span class="chg-badge st-${item.status}">${esc(t(STATUS_KEYS[item.status] || 'proposalPending'))}</span>
        <span class="chg-label">${esc(item.label)}</span>
        ${jump}
      </div>`;
    }).join('')}</div>`;
    return `<div class="chg-bar">${head}${rows}</div>`;
  }

  App.buildChangeLogItems = buildChangeLogItems;
  App.changeLogRef = changeLogRef;
  App.changeLogHtml = changeLogHtml;
})();
