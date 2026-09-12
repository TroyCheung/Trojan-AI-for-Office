(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ── render 工具：把模型的结构化输出渲染成卡片 ──
  //
  // 设计约束（改动前请先读）：
  // 1. 本文件不修改 ui.js 的任何既有函数。ui.js 只在 toolCardHtml 里加一个分支调用
  //    App.renderBlockHtml(tc)，以及在 click 分发里调一次 App.handleBlockAction(action, el)。
  //    历史上 ui.js 被批量重构误删过 settleProposal / settleBatchIfDone，所以新功能一律进新文件。
  // 2. 序号一律由界面生成（1,2,3…），模型给的任何编号都忽略。这是为了避免段落序号泄漏到用户眼前。
  // 3. target 只用于跳转定位，永远不渲染成可见文本。
  // 4. 五种 type 共享同一套卡片语法：blk-card > blk-head / blk-body / blk-foot。

  const esc = (s) => App.escapeHtml(s == null ? '' : s);
  const ico = (n, o) => (App.icon ? App.icon(n, o) : '');
  const t = (k) => App.t(k);

  const TYPES = ['list', 'compare', 'matrix', 'outline', 'form'];

  // ---------- 校验：坏数据要能被模型看懂并重试，而不是渲染出半个空卡 ----------
  function validateBlock(args) {
    const a = args && typeof args === 'object' ? args : {};
    const type = String(a.type || '').trim();
    if (!TYPES.includes(type)) return `type 必须是 ${TYPES.join(' / ')} 之一，收到的是 ${JSON.stringify(a.type)}。`;
    if (type === 'list') {
      if (!Array.isArray(a.items) || !a.items.length) return 'list 需要非空的 items 数组。';
      const bad = a.items.findIndex(x => !x || !String(x.title || '').trim());
      if (bad >= 0) return `items[${bad}] 缺少 title。每一项都必须有简短标题。`;
    }
    if (type === 'compare') {
      if (!Array.isArray(a.options) || a.options.length < 2) return 'compare 需要至少 2 个 options，只有一个选项不构成比较。';
      if (!Array.isArray(a.fields) || !a.fields.length) return 'compare 需要 fields 数组，用来对齐各选项的同名字段。';
      const bad = a.options.findIndex(x => !x || !String(x.name || '').trim());
      if (bad >= 0) return `options[${bad}] 缺少 name。`;
    }
    if (type === 'matrix') {
      if (!Array.isArray(a.columns) || !a.columns.length) return 'matrix 需要 columns 数组。';
      if (!Array.isArray(a.rows) || !a.rows.length) return 'matrix 需要非空的 rows 数组。';
      const n = a.columns.length;
      const bad = a.rows.findIndex(r => !r || !Array.isArray(r.cells) || r.cells.length !== n);
      if (bad >= 0) return `rows[${bad}].cells 的长度必须等于 columns 的长度（${n}）。`;
    }
    if (type === 'outline') {
      if (!Array.isArray(a.nodes) || !a.nodes.length) return 'outline 需要非空的 nodes 数组。';
    }
    if (type === 'form') {
      if (!Array.isArray(a.fields) || !a.fields.length) return 'form 需要非空的 fields 数组。';
      const bad = a.fields.findIndex(f => !f || !String(f.key || '').trim());
      if (bad >= 0) return `fields[${bad}] 缺少 key。`;
    }
    return '';
  }

  // ---------- 状态：每个 render 调用有自己的一份交互状态 ----------
  function blockState(tc) {
    if (!tc.rid) tc.rid = App.id();
    if (!tc.block) {
      const a = tc.args || {};
      tc.block = {
        settled: false,
        // list：每项 pending / accepted / skipped
        items: Array.isArray(a.items) ? a.items.map(() => 'pending') : [],
        // compare：选中的 option index
        picked: null,
        // outline：展开的节点
        open: Array.isArray(a.nodes) ? a.nodes.map((n, i) => n && n.open === true ? true : i === 0) : [],
        // form：字段当前值
        values: Array.isArray(a.fields) ? a.fields.reduce((m, f) => (m[f.key] = f.value == null ? '' : String(f.value), m), {}) : {},
        // matrix：维度筛选
        filter: null
      };
    }
    return tc.block;
  }

  const SEV = {
    high:   { cls: 'sev-high', key: 'sevHigh' },
    medium: { cls: 'sev-med',  key: 'sevMedium' },
    low:    { cls: 'sev-low',  key: 'sevLow' }
  };

  // 有需要用户决策的按钮才算交互卡（与 presentRenderBlock 的阻塞判定一致）；
  // matrix / outline / 纯陈述 list 是只读形态，重开后按钮（定位、筛选、折叠）仍然可用
  function isInteractive(a) {
    return (a.type === 'list' && (a.items || []).some(x => x && x.actionable !== false))
      || (a.type === 'compare' && a.selectable !== false)
      || a.type === 'form';
  }

  // 插件重开后 Promise 解析器不随会话持久化，交互卡的按钮会静默失效；
  // 这种卡渲染成「已过期」只读态（与 diff 卡的 expired 判定同款）
  function isExpired(tc) {
    const b = blockState(tc);
    return !b.settled && !tc._resolveBlock && !(App.state && App.state.isWorking) && isInteractive(tc.args || {});
  }

  function expiredFootHtml() {
    return `<div class="blk-bar"><span class="blk-count">${esc(t('proposalExpired'))} · ${esc(t('proposalExpiredHint'))}</span></div>`;
  }

  function sevChip(v) {
    const s = SEV[String(v || '').toLowerCase()];
    if (!s) return '';
    return `<span class="blk-sev ${s.cls}">${esc(t(s.key))}</span>`;
  }

  function hasTarget(x) {
    const tg = x && x.target;
    if (!tg || typeof tg !== 'object') return false;
    return Number.isInteger(Number(tg.paragraphIndex)) || tg.slideId != null || Number.isInteger(Number(tg.index)) || tg.tag != null;
  }

  // ---------- list ----------
  function listHtml(tc) {
    const a = tc.args || {};
    const b = blockState(tc);
    const items = a.items || [];
    const expired = isExpired(tc);
    const done = b.settled || expired;
    const actionable = items.some(x => x && x.actionable !== false);

    const cards = items.map((item, i) => {
      const st = b.items[i] || 'pending';
      const badge = st === 'accepted' ? `<span class="blk-badge ok">${ico('check')}${esc(t('blkAccepted'))}</span>`
        : st === 'skipped' ? `<span class="blk-badge no">${ico('close')}${esc(t('blkSkipped'))}</span>`
        : sevChip(item.severity);
      const locatable = hasTarget(item);
      const foot = (done || st !== 'pending' || item.actionable === false) ? '' : `
        <div class="blk-foot">
          ${locatable ? `<button type="button" class="blk-btn" data-action="blk-locate" data-rid="${tc.rid}" data-idx="${i}">${ico('locate')}${esc(t('blkLocate'))}</button>` : ''}
          <button type="button" class="blk-btn primary" data-action="blk-accept" data-rid="${tc.rid}" data-idx="${i}">${esc(t('blkAccept'))}</button>
          <button type="button" class="blk-btn" data-action="blk-skip" data-rid="${tc.rid}" data-idx="${i}">${esc(t('blkSkip'))}</button>
        </div>`;
      return `<div class="blk-card st-${st}">
        <div class="blk-head">
          <span class="blk-num">${i + 1}</span>
          <span class="blk-title">${esc(item.title)}</span>
          ${badge}
        </div>
        <div class="blk-body">
          ${item.quote ? `<div class="blk-quote${locatable ? ' locatable' : ''}"${locatable ? ` data-action="blk-locate" data-rid="${tc.rid}" data-idx="${i}" role="button" tabindex="0"` : ''}><span class="blk-tag">${esc(t('blkQuote'))}${locatable ? ' · ' + esc(t('blkClickLocate')) : ''}</span>${esc(item.quote)}</div>` : ''}
          ${item.problem ? `<div class="blk-note">${esc(item.problem)}</div>` : ''}
          ${item.suggestion ? `<div class="blk-fix"><span class="blk-tag ok">${esc(t('blkSuggest'))}</span>${esc(item.suggestion)}</div>` : ''}
        </div>
        ${foot}
      </div>`;
    }).join('');

    const pending = b.items.filter(s => s === 'pending').length;
    const foot = (done || !actionable || !pending) ? '' : `
      <div class="blk-bar">
        <span class="blk-count">${esc(t('blkCount').replace('{total}', items.length).replace('{done}', items.length - pending))}</span>
        <button type="button" class="blk-btn" data-action="blk-skip-all" data-rid="${tc.rid}">${esc(t('blkSkipAll'))}</button>
        <button type="button" class="blk-btn primary" data-action="blk-accept-all" data-rid="${tc.rid}">${esc(t('blkAcceptAll'))}</button>
      </div>`;

    return summaryHtml(a) + cards + foot + (expired ? expiredFootHtml() : '') + exportsHtml(tc, a);
  }

  // ---------- compare ----------
  function compareHtml(tc) {
    const a = tc.args || {};
    const b = blockState(tc);
    const expired = isExpired(tc);
    const done = b.settled || expired;
    const fields = a.fields || [];
    const options = a.options || [];
    const selectable = a.selectable !== false;

    const cols = options.map((opt, i) => {
      const picked = b.picked === i;
      const rows = fields.map(f => {
        const v = opt.values && opt.values[f];
        if (v == null || v === '') return '';
        return `<div class="cmp-row"><span class="blk-tag">${esc(f)}</span><div class="cmp-val">${esc(v)}</div></div>`;
      }).join('');
      const btn = !selectable ? '' : (done
        ? (picked ? `<span class="blk-badge ok">${ico('check')}${esc(t('blkPicked'))}</span>` : '')
        : `<button type="button" class="blk-btn ${picked ? 'primary' : ''}" data-action="blk-pick" data-rid="${tc.rid}" data-idx="${i}">${esc(picked ? t('blkPicked') : t('blkPick'))}</button>`);
      return `<div class="cmp-col${picked ? ' picked' : ''}">
        <div class="cmp-head">
          ${opt.key ? `<span class="blk-tag">${esc(opt.key)}</span>` : ''}
          <div class="cmp-name">${esc(opt.name)}</div>
        </div>
        ${rows}
        ${btn ? `<div class="cmp-foot">${btn}</div>` : ''}
      </div>`;
    }).join('');

    return summaryHtml(a)
      + `<div class="cmp-scroll"><div class="cmp" style="--cmp-n:${options.length}">${cols}</div></div>`
      + (expired ? expiredFootHtml() : '')
      + exportsHtml(tc, a);
  }

  // ---------- matrix ----------
  function matrixHtml(tc) {
    const a = tc.args || {};
    const b = blockState(tc);
    const cols = a.columns || [];
    const rows = a.rows || [];
    const tags = Array.from(new Set(rows.flatMap(r => Array.isArray(r.tags) ? r.tags : []).filter(Boolean)));
    const shown = b.filter ? rows.filter(r => Array.isArray(r.tags) && r.tags.includes(b.filter)) : rows;

    const filterBar = tags.length < 2 ? '' : `<div class="mx-filters">
      <button type="button" class="mx-chip${b.filter ? '' : ' on'}" data-action="blk-filter" data-rid="${tc.rid}" data-tag="">${esc(t('blkAll'))}</button>
      ${tags.map(g => `<button type="button" class="mx-chip${b.filter === g ? ' on' : ''}" data-action="blk-filter" data-rid="${tc.rid}" data-tag="${esc(g)}">${esc(g)}</button>`).join('')}
    </div>`;

    const body = shown.map(r => {
      const i = rows.indexOf(r);
      const cells = (r.cells || []).map((c, ci) => {
        const isFirst = ci === 0;
        if (isFirst && hasTarget(r)) {
          return `<td><button type="button" class="mx-jump" data-action="blk-locate" data-rid="${tc.rid}" data-idx="${i}">${esc(c)}</button></td>`;
        }
        return `<td class="${isFirst ? 'mx-key' : ''}">${esc(c)}</td>`;
      }).join('');
      const tagCell = tags.length ? `<td>${(r.tags || []).map(g => `<span class="mx-tag">${esc(g)}</span>`).join('')}</td>` : '';
      return `<tr>${cells}${tagCell}</tr>`;
    }).join('');

    return summaryHtml(a) + filterBar
      + `<div class="mx-scroll"><table class="mx">
          <thead><tr>${cols.map(c => `<th>${esc(c)}</th>`).join('')}${tags.length ? `<th>${esc(t('blkTag'))}</th>` : ''}</tr></thead>
          <tbody>${body}</tbody>
        </table></div>`
      + exportsHtml(tc, a);
  }

  // ---------- outline ----------
  function outlineHtml(tc) {
    const a = tc.args || {};
    const b = blockState(tc);
    const root = a.root || null;
    const nodes = a.nodes || [];

    const rootHtml = root ? `<div class="ol-core">${root.label ? `<span class="blk-tag">${esc(root.label)}</span>` : ''}<div>${esc(root.text || '')}</div></div>` : '';
    const body = nodes.map((n, i) => {
      const open = b.open[i] === true;
      const items = Array.isArray(n.items) ? n.items : [];
      return `<div class="ol-node">
        <button type="button" class="ol-sum${open ? ' open' : ''}" data-action="blk-toggle-node" data-rid="${tc.rid}" data-idx="${i}" aria-expanded="${open}">
          ${ico(open ? 'chevronDown' : 'chevronRight')}
          ${n.label ? `<span class="blk-tag">${esc(n.label)}</span>` : ''}
          <span class="ol-title">${esc(n.title || '')}</span>
          <span class="ol-n">${items.length}</span>
        </button>
        ${open && items.length ? `<ul class="ol-items">${items.map(x => `<li>${esc(x)}</li>`).join('')}</ul>` : ''}
      </div>`;
    }).join('');

    return summaryHtml(a) + rootHtml + `<div class="ol-wrap">${body}</div>` + exportsHtml(tc, a);
  }

  // ---------- form ----------
  // 宽版判定：任一字段名超过 8 字（ask_clarification 的整句问题），整个表单切成
  // 左右对半——问题占左半幅铺开，作答区占右半幅。短字段名（Brief 表单）保持 56px
  // 窄标签的紧凑形态不变。
  const WIDE_LABEL_CHARS = 8;

  function formHtml(tc) {
    const a = tc.args || {};
    const b = blockState(tc);
    const expired = isExpired(tc);
    const done = b.settled || expired;
    const fields = a.fields || [];
    const wide = fields.some(f => String((f && f.key) || '').length > WIDE_LABEL_CHARS);
    const missing = fields.filter(f => f.required !== false && !String(b.values[f.key] || '').trim()).length;

    const rows = fields.map(f => {
      const v = String(b.values[f.key] == null ? '' : b.values[f.key]);
      const hasOpts = Array.isArray(f.options) && f.options.length > 0;
      // 选项字段 = 输入框 + 选项 chips（2026-09-03 用户：原生 select 只能选不能填，
      // 和「点击填写」的承诺矛盾；现在点 chip 即填，选项都不合适时直接打字）
      const editor = `<input class="fm-in" type="text" value="${esc(v)}" placeholder="${esc(f.placeholder || t(hasOpts ? 'blkFillOpts' : 'blkFillHere'))}" data-action="blk-field" data-rid="${tc.rid}" data-key="${esc(f.key)}" />`;
      const chips = !hasOpts || done ? '' : `<span class="fm-opts">${f.options.filter(o => String(o).trim()).map(o => {
        const opt = String(o);
        return `<button type="button" class="fm-opt${opt === v ? ' on' : ''}" data-action="blk-option" data-rid="${tc.rid}" data-key="${esc(f.key)}" data-opt="${esc(opt)}">${esc(opt)}</button>`;
      }).join('')}</span>`;
      const body = done
        ? `<span class="fm-val${v ? '' : ' empty'}">${esc(v || '—')}</span>`
        : hasOpts ? `<span class="fm-field">${editor}${chips}</span>` : editor;
      return `<label class="fm-row">
        <span class="fm-lbl">${esc(f.key)}</span>
        ${body}
      </label>`;
    }).join('');

    const foot = b.settled ? `<div class="blk-bar"><span class="blk-count">${esc(t('blkSubmitted'))}</span></div>`
      : expired ? expiredFootHtml() : `
      <div class="blk-bar">
        <span class="blk-count">${missing ? esc(t('blkMissing').replace('{n}', missing)) : esc(t('blkAllSet'))}</span>
        <button type="button" class="blk-btn primary" data-action="blk-submit" data-rid="${tc.rid}">${esc(a.submitLabel || t('blkSubmit'))}</button>
      </div>`;

    return summaryHtml(a) + `<div class="fm${wide ? ' fm-wide' : ''}">${rows}</div>` + (a.note ? `<div class="blk-hint">${esc(a.note)}</div>` : '') + foot;
  }

  // ---------- 公共片段 ----------
  function summaryHtml(a) {
    const s = String(a.summary || '').trim();
    return s ? `<div class="blk-summary">${esc(s)}</div>` : '';
  }

  // 输出即制品：把卡片内容落回 Office 文档
  const EXPORT_LABELS = {
    'word-comment': 'blkToComment',
    'word-table':   'blkToTable',
    'xlsx':         'blkToXlsx',
    'pptx-skeleton':'blkToDeck'
  };
  const EXPORT_ICONS = {
    'word-comment': 'comment', 'word-table': 'table', 'xlsx': 'download', 'pptx-skeleton': 'slide'
  };
  function exportsHtml(tc, a) {
    const list = (Array.isArray(a.exportable) ? a.exportable : []).filter(x => EXPORT_LABELS[x]);
    if (!list.length) return '';
    return `<div class="blk-ships">${list.map(x =>
      `<button type="button" class="blk-ship" data-action="blk-export" data-rid="${tc.rid}" data-kind="${esc(x)}">${ico(EXPORT_ICONS[x])}${esc(t(EXPORT_LABELS[x]))}</button>`
    ).join('')}</div>`;
  }

  const RENDERERS = { list: listHtml, compare: compareHtml, matrix: matrixHtml, outline: outlineHtml, form: formHtml };

  // ---------- 入口：ui.js 的 toolCardHtml 调这个 ----------
  function renderBlockHtml(tc) {
    if (tc.blockError) {
      return `<div class="blk-wrap"><div class="blk-card err"><div class="blk-head">${ico('alert')}<span class="blk-title">${esc(t('blkInvalid'))}</span></div><div class="blk-body"><div class="blk-note">${esc(tc.blockError)}</div></div></div></div>`;
    }
    const type = String((tc.args && tc.args.type) || '').trim();
    const fn = RENDERERS[type];
    if (!fn) return '';
    blockState(tc);
    return `<div class="blk-wrap blk-${type}">${fn(tc)}</div>`;
  }

  // ---------- 交互：ui.js 的 click 分发把 blk-* 开头的 action 转过来 ----------
  function findBlock(rid) {
    const msgs = App.state.messages;
    for (let i = msgs.length - 1; i >= 0; i--) {
      const m = msgs[i];
      if (!m.toolCalls) continue;
      const tc = m.toolCalls.find(x => x.name === 'render' && x.rid === rid);
      if (tc) return tc;
    }
    return null;
  }

  function settle(tc, extra) {
    const b = blockState(tc);
    if (b.settled) return;
    b.settled = true;
    const a = tc.args || {};
    const result = Object.assign({ success: true, type: a.type }, extra || {});
    tc.result = result;
    const r = tc._resolveBlock;
    if (r) { tc._resolveBlock = null; r(result); }
  }

  function maybeSettleList(tc) {
    const b = blockState(tc);
    if (b.items.some(s => s === 'pending')) return;
    const a = tc.args || {};
    const items = a.items || [];
    settle(tc, {
      accepted: b.items.map((s, i) => s === 'accepted' ? { index: i, title: items[i] && items[i].title } : null).filter(Boolean),
      skipped:  b.items.map((s, i) => s === 'skipped'  ? { index: i, title: items[i] && items[i].title } : null).filter(Boolean)
    });
  }

  async function handleBlockAction(action, el) {
    const rid = el.dataset.rid;
    const tc = findBlock(rid);
    if (!tc) return true;
    const b = blockState(tc);
    const a = tc.args || {};
    const idx = Number(el.dataset.idx);

    // 过期卡（重开后解析器丢失）不允许任何会改变结算状态的操作；
    // 定位、折叠、筛选、导出是只读辅助动作，仍然放行
    if (isExpired(tc) && !['blk-locate', 'blk-toggle-node', 'blk-filter', 'blk-export'].includes(action)) return true;

    if (action === 'blk-locate') {
      const src = a.type === 'matrix' ? (a.rows || [])[idx] : (a.items || [])[idx];
      const tg = src && src.target;
      if (tg) {
        try {
          if (Number.isInteger(Number(tg.paragraphIndex))) await App.navigateCitation(`p:${Number(tg.paragraphIndex)}`);
          else if (tg.slideId != null) await App.navigateCitation(`s:${Number(tg.index) || 0}`);
          else if (tg.tag) await App.navigateCitation(`cc:${tg.tag}`);
          else if (Number.isInteger(Number(tg.index))) await App.navigateCitation(`s:${Number(tg.index)}`);
        } catch (e) { App.state.error = e.message || String(e); App.render(); }
      }
      return true;
    }

    if (action === 'blk-accept' || action === 'blk-skip') {
      if (b.settled || b.items[idx] !== 'pending') return true;
      b.items[idx] = action === 'blk-accept' ? 'accepted' : 'skipped';
      App.render();
      maybeSettleList(tc);
      App.persistCurrentSession();
      return true;
    }

    if (action === 'blk-accept-all' || action === 'blk-skip-all') {
      if (b.settled) return true;
      const mark = action === 'blk-accept-all' ? 'accepted' : 'skipped';
      for (let i = 0; i < b.items.length; i++) if (b.items[i] === 'pending') b.items[i] = mark;
      App.render();
      maybeSettleList(tc);
      App.persistCurrentSession();
      return true;
    }

    if (action === 'blk-pick') {
      if (b.settled) return true;
      b.picked = idx;
      App.render();
      const opt = (a.options || [])[idx] || {};
      settle(tc, { selected: { index: idx, key: opt.key, name: opt.name } });
      App.persistCurrentSession();
      return true;
    }

    if (action === 'blk-toggle-node') {
      b.open[idx] = !b.open[idx];
      App.render();
      return true;
    }

    if (action === 'blk-filter') {
      b.filter = el.dataset.tag || null;
      App.render();
      return true;
    }

    if (action === 'blk-option') {
      if (b.settled) return true;
      const key = el.dataset.key;
      // 再点已选中的 chip = 清掉，不打字也能反悔
      b.values[key] = b.values[key] === el.dataset.opt ? '' : el.dataset.opt;
      App.render();
      // 焦点还给这张卡的输入框：点完 chip 常接着要补充打字，别让人重新点一次
      const inputs = document.querySelectorAll(`[data-action="blk-field"][data-rid="${tc.rid}"]`);
      for (const inp of inputs) { if (inp.dataset.key === key) { inp.focus(); break; } }
      return true;
    }

    if (action === 'blk-submit') {
      if (b.settled) return true;
      settle(tc, { fields: Object.assign({}, b.values) });
      App.render();
      App.persistCurrentSession();
      return true;
    }

    if (action === 'blk-export') {
      await runExport(tc, el.dataset.kind);
      return true;
    }

    return false;
  }

  // 输入框是 input 事件，不走 click 分发
  function handleBlockInput(el) {
    if (el.dataset.action !== 'blk-field') return false;
    const tc = findBlock(el.dataset.rid);
    if (!tc) return false;
    const b = blockState(tc);
    b.values[el.dataset.key] = el.value;
    // 不重渲染（保输入焦点），但同 key 的 chip 高亮要跟上，不然视觉上「还选着」而答案已是自拟文本
    const wrap = el.parentElement;
    if (wrap && wrap.classList.contains('fm-field')) {
      wrap.querySelectorAll('.fm-opt').forEach(btn => btn.classList.toggle('on', btn.dataset.opt === el.value));
    }
    return true;
  }

  // ---------- 输出即制品 ----------
  function listToLines(a) {
    return (a.items || []).map((x, i) => `${i + 1}. ${x.title || ''}${x.suggestion ? `\n   建议：${x.suggestion}` : ''}`);
  }

  async function runExport(tc, kind) {
    const a = tc.args || {};
    const state = App.state;
    try {
      if (kind === 'word-comment') {
        const items = a.items || [];
        let ok = 0;
        for (const item of items) {
          const tg = item.target || {};
          if (!Number.isInteger(Number(tg.paragraphIndex))) continue;
          await App.navigateCitation(`p:${Number(tg.paragraphIndex)}`);
          const body = [item.problem, item.suggestion ? `建议：${item.suggestion}` : ''].filter(Boolean).join('\n');
          await App.executeToolByName('manage_comment', { operation: 'add', text: `${item.title}\n${body}` });
          ok++;
        }
        state.error = null;
        flash(tc, t('blkDoneComments').replace('{n}', ok));
        return;
      }
      if (kind === 'word-table') {
        const rows = [(a.columns || []).slice()];
        for (const r of (a.rows || [])) rows.push((r.cells || []).map(x => String(x == null ? '' : x)));
        await App.executeToolByName('insert_table', { rows, location: 'End', headerRow: true });
        flash(tc, t('blkDoneTable'));
        return;
      }
      if (kind === 'xlsx') {
        const rows = [(a.columns || []).slice()].concat((a.rows || []).map(r => (r.cells || []).map(x => String(x == null ? '' : x))));
        const csv = rows.map(r => r.map(c => /[",\n]/.test(c) ? `"${c.replace(/"/g, '""')}"` : c).join(',')).join('\n');
        await copyText('﻿' + csv);
        flash(tc, t('blkDoneCsv'));
        return;
      }
      if (kind === 'pptx-skeleton') {
        const nodes = a.nodes || [];
        for (const n of nodes) {
          const res = await App.executeToolByName('add_slide', {});
          const index = res && Number.isInteger(res.index) ? res.index : null;
          if (index == null) continue;
          await App.executeToolByName('insert_textbox', {
            index, text: n.title || n.label || '', left: 60, top: 60, width: 600, height: 60
          });
          if (Array.isArray(n.items) && n.items.length) {
            await App.executeToolByName('insert_textbox', {
              index, text: n.items.map(x => '· ' + x).join('\n'), left: 60, top: 150, width: 600, height: 260
            });
          }
        }
        flash(tc, t('blkDoneDeck').replace('{n}', nodes.length));
        return;
      }
    } catch (e) {
      flash(tc, (t('blkExportFailed') || '') + (e.message || String(e)));
    }
  }

  function flash(tc, msg) {
    tc.blockFlash = msg;
    App.render();
    setTimeout(() => { if (tc.blockFlash === msg) { tc.blockFlash = null; App.render(); } }, 4000);
  }

  async function copyText(text) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) { await navigator.clipboard.writeText(text); return; }
    } catch {}
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); } finally { document.body.removeChild(ta); }
  }

  // ---------- 工具定义（宿主无关，api.js 会追加到工具列表）----------
  const TARGET_SCHEMA = {
    type: 'object',
    description: '定位锚点，只用于点击跳转，不会显示给用户。',
    properties: {
      paragraphIndex: { type: 'number', description: 'Word：段落序号' },
      slideId: { type: 'string', description: 'PowerPoint：slideId' },
      index: { type: 'number', description: 'PowerPoint：幻灯片索引' },
      tag: { type: 'string', description: 'Word：内容控件 tag' }
    }
  };

  const TOOL = {
    type: 'function',
    function: {
      name: 'render',
      description: [
        '把结构化结果渲染成界面卡片。界面负责编号、对齐、跳转和操作按钮，你只提供数据。',
        '',
        '按场景选 type（这是最常见的对应关系，照着选即可）：',
        '· 逐条审查、修改建议、问题清单、检查清单、待办、风险点、实体档案 → list',
        '· 候选论证线、几个标题备选、几种写法、方案 A/B/C、人选或代言人对比、优缺点对照 → compare',
        '· 观点与事实拆解、维度归类、素材清单、逐项核对表、任何行×列的数据 → matrix',
        '· 叙事策略框架、信息屋、方案结构、大纲、支柱与要点的层级关系 → outline',
        '· Brief 确认、参数确认、需要用户填空或改写后回传的字段 → form',
        '',
        '触发信号（命中任意一条就该用 render，而不是写成文字）：',
        '· 你正准备写「1. …… 2. …… 3. ……」这样的编号列表',
        '· 你正准备写「方案一 / 方案二 / 方案三」或「线 A / 线 B」这样的并列选项',
        '· 你正准备写一个 markdown 表格',
        '· 你正准备用缩进的项目符号表达层级',
        '· 你正准备逐个提问来收集参数',
        '',
        '不要用于：普通问答、解释概念、闲聊、确认、只有一两句话的回答、单个连续段落的叙述或改写。给一段话套上卡片只会让界面变吵。',
        '用了 render 之后，正文里不要再把同样的内容复述一遍，写一句话引出即可。'
      ].join('\n'),
      parameters: {
        type: 'object',
        properties: {
          type: { type: 'string', enum: TYPES, description: 'list=同构列表；compare=并排比较；matrix=行列矩阵；outline=层级大纲；form=可编辑表单' },
          summary: { type: 'string', description: '一两句整体判断，显示在卡片上方。' },
          items: {
            type: 'array',
            description: 'type=list 时使用。序号由界面生成，不要自己写编号。',
            items: {
              type: 'object',
              properties: {
                title: { type: 'string', description: '短标题，一行以内。' },
                severity: { type: 'string', enum: ['high', 'medium', 'low'] },
                quote: { type: 'string', description: '原文摘录，可选。' },
                problem: { type: 'string', description: '判断或说明。' },
                suggestion: { type: 'string', description: '建议改法，可选。' },
                actionable: { type: 'boolean', description: '默认 true，显示采纳/跳过按钮。纯陈述项设为 false。' },
                target: TARGET_SCHEMA
              },
              required: ['title']
            }
          },
          fields: {
            type: 'array',
            description: 'type=compare 时是用于对齐的字段名数组（字符串）；type=form 时是字段对象数组。',
            items: {}
          },
          options: {
            type: 'array',
            description: 'type=compare 时使用，至少 2 项。',
            items: {
              type: 'object',
              properties: {
                key: { type: 'string', description: '短代号，如「线 A」。' },
                name: { type: 'string', description: '选项名称。' },
                values: { type: 'object', description: '键为 fields 里的字段名，值为该选项在该字段下的内容。' }
              },
              required: ['name', 'values']
            }
          },
          selectable: { type: 'boolean', description: 'type=compare，默认 true，显示「选这条」按钮。' },
          columns: { type: 'array', items: { type: 'string' }, description: 'type=matrix 的表头。' },
          rows: {
            type: 'array',
            description: 'type=matrix 的数据行。cells 长度必须等于 columns 长度。',
            items: {
              type: 'object',
              properties: {
                cells: { type: 'array', items: { type: 'string' } },
                tags: { type: 'array', items: { type: 'string' }, description: '维度标签，用于筛选。' },
                target: TARGET_SCHEMA
              },
              required: ['cells']
            }
          },
          root: {
            type: 'object',
            description: 'type=outline 的根节点。',
            properties: { label: { type: 'string' }, text: { type: 'string' } }
          },
          nodes: {
            type: 'array',
            description: 'type=outline 的分支。',
            items: {
              type: 'object',
              properties: {
                label: { type: 'string', description: '如「支柱一」。' },
                title: { type: 'string' },
                items: { type: 'array', items: { type: 'string' } },
                open: { type: 'boolean' }
              },
              required: ['title']
            }
          },
          note: { type: 'string', description: 'type=form 的说明文字。' },
          submitLabel: { type: 'string', description: 'type=form 的提交按钮文案。' },
          exportable: {
            type: 'array',
            items: { type: 'string', enum: ['word-comment', 'word-table', 'xlsx', 'pptx-skeleton'] },
            description: '允许用户把内容落回文档的方式。word-comment 仅 list+Word；word-table/xlsx 仅 matrix；pptx-skeleton 仅 outline+PowerPoint。'
          }
        },
        required: ['type']
      }
    }
  };

  // 各宿主允许的 export 类型不同，工具定义按宿主裁剪，避免模型提出做不到的操作
  function toolForHost(hostType) {
    const tool = JSON.parse(JSON.stringify(TOOL));
    const allowed = hostType === 'word' ? ['word-comment', 'word-table', 'xlsx']
      : hostType === 'powerpoint' ? ['pptx-skeleton', 'xlsx']
      : ['xlsx'];
    tool.function.parameters.properties.exportable.items.enum = allowed;
    return tool;
  }

  App.RENDER_BLOCK_TYPES = TYPES;
  App.renderBlockTool = toolForHost;
  App.validateRenderBlock = validateBlock;
  App.renderBlockHtml = renderBlockHtml;
  App.handleBlockAction = handleBlockAction;
  App.handleBlockInput = handleBlockInput;
  App.blockState = blockState;
  App.renderBlockIsInteractive = isInteractive;
  App.presentRenderBlock = function (uiCall, args) {
    const err = validateBlock(args);
    if (err) {
      uiCall.blockError = err;
      return Promise.resolve({ success: false, error: err, retryable: true });
    }
    uiCall.args = args;
    blockState(uiCall);
    const a = args || {};
    // 只读形态（无按钮）立刻结算，不阻塞 agent loop
    if (!isInteractive(a)) return Promise.resolve({ success: true, type: a.type, rendered: true });
    return new Promise(resolve => { uiCall._resolveBlock = resolve; });
  };

  // ask_clarification（PPT 宿主工具）的呈现通道：把 questions 翻译成 render 的 form 卡，
  // 复用现有渲染/结算管线，不新增卡片类型。uiCall.name 改写成 'render'：
  // 卡片的渲染（ui.js 只认 render）、按钮事件找回（findBlock）、停止时的结算（abortRenderBlock）
  // 全都按 name === 'render' 分发——而这张卡本来就是一张 render form 卡。
  App.presentClarification = function (uiCall, args) {
    const raw = Array.isArray(args && args.questions) ? args.questions : [];
    if (!raw.length) {
      return Promise.resolve({ success: false, retryable: true, error: 'ask_clarification 需要非空的 questions 数组：每项 { id, label, type: "select"|"text", options? }，其中 label 是显示给用户的问题文本。' });
    }
    // 缺 label 必须逐项点名拒绝，不能静默过滤：静默丢题后模型只看到「数组为空」，
    // 不知道哪项缺什么字段，无法自愈（v137 真机事故：两道题全缺 label 被过滤后模型放弃卡片路线）
    const missingLabel = [];
    raw.forEach((q, i) => {
      if (!q || !String(q.label || '').trim()) missingLabel.push(`questions[${i}]${q && q.id ? `（id: ${q.id}）` : ''}`);
    });
    if (missingLabel.length) {
      return Promise.resolve({
        success: false,
        retryable: true,
        error: `${missingLabel.join('、')} 缺少必填的 label 字段。label 是显示给用户的问题文本（例如「「内部资料」放在哪里？」），id 只是内部键不能替代。请补全后原样重发。`
      });
    }
    const questions = raw.slice(0, 8);
    const badSelect = questions.findIndex(q => String(q.type || 'select') !== 'text' && (!Array.isArray(q.options) || q.options.length < 2));
    if (badSelect >= 0) {
      return Promise.resolve({ success: false, retryable: true, error: `questions[${badSelect}]（${String(questions[badSelect].label).slice(0, 40)}）是 select 但 options 少于 2 个：补上 options，或把 type 改成 "text"。` });
    }
    // 用户看到的是字段标签：key 直接用 label（form 卡把 key 当标签渲染）；同名标签加序号防撞键
    const usedKeys = new Set();
    const fields = questions.map(q => {
      let key = String(q.label).slice(0, 120);
      for (let n = 2; usedKeys.has(key); n++) key = `${String(q.label).slice(0, 110)}（${n}）`;
      usedKeys.add(key);
      const isText = String(q.type || 'select') === 'text';
      return {
        key,
        options: isText ? undefined : q.options.map(String).slice(0, 12),
        placeholder: q.placeholder ? String(q.placeholder).slice(0, 120) : undefined,
        value: ''
      };
    });
    uiCall.name = 'render';
    const presented = App.presentRenderBlock(uiCall, {
      type: 'form',
      summary: String(args.summary || '动手前需要你先确认几件事'),
      fields,
      note: args.note ? String(args.note) : undefined,
      submitLabel: args.submitLabel ? String(args.submitLabel) : undefined
    });
    return presented.then(result => {
      // 校验失败 / 用户中止：原样透传，语义保持「没有提交」
      if (!result || result.success === false) return result;
      // 按 fields[i].key 查值，不用 Object.values 下标对齐——纯数字标签会被 JS
      // 当整数型键排到最前，下标映射会把答案塞进错误的 question.id
      const answers = {};
      questions.forEach((q, i) => {
        const v = result.fields ? result.fields[fields[i].key] : null;
        answers[String(q.id || `q${i + 1}`)] = v == null ? '' : v;
      });
      return { success: true, type: 'ask_clarification', answers };
    });
  };

  // 用户点停止时结算交互卡：只释放挂起的 Promise，不改卡片可视状态。
  // 结果明确告知模型是「中止」而不是提交；任务停下后 isWorking=false，
  // 卡片自然落入「已过期」只读态，语义就是「用户没有提交任何内容」。
  App.abortRenderBlock = function (tc) {
    const r = tc && tc._resolveBlock;
    if (!r) return;
    tc._resolveBlock = null;
    const a = tc.args || {};
    const result = { success: false, aborted: true, type: a.type, error: '用户点击了停止，这张卡片没有提交任何选择或表单内容，请勿当作已确认。' };
    tc.result = result;
    r(result);
  };
})();
