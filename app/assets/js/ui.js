(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 本模块在其余脚本之后加载，取别名时这些成员均已挂载到 App。
  const state = App.state;
  const t = App.t;
  const escapeHtml = App.escapeHtml;
  const pretty = App.pretty;
  const now = App.now;
  const hasOffice = App.hasOffice;
  const STORAGE_KEYS = App.STORAGE_KEYS;
  const getProviderModels = App.getProviderModels;
  const sessionsForCurrentWorkbook = App.sessionsForCurrentWorkbook;
  const icon = (n, o) => (App.icon ? App.icon(n, o) : '');
  const renderMarkdown = App.renderMarkdown;
  const trimToolText = App.trimToolText;
  const makeSession = App.makeSession;
  const saveSessions = App.saveSessions;
  const ensureSession = App.ensureSession;
  const persistCurrentSession = App.persistCurrentSession;
  const saveSettings = App.saveSettings;
  const setStoredItem = App.setStoredItem;
  const runAgentLoop = App.runAgentLoop;
  const markStoppedMessage = App.markStoppedMessage;
  const executeToolByName = App.executeToolByName;
  const maybeFollow = App.maybeFollow;
  const navigateCitation = App.navigateCitation;

  // ---- 工具人话化：名称双语映射（三宿主工具名不冲突）----
  const TOOL_LABELS = {
    zh: {
      get_document_outline: '读取文档大纲', get_selection: '读取选区', get_paragraphs: '读取段落', get_tables: '读取表格',
      search_text: '搜索文本', insert_text: '插入文本', replace_text: '替换文本', apply_style: '设置样式',
      set_paragraph_format: '设置段落格式', insert_heading: '插入标题', insert_table: '插入表格', insert_page_break: '插入分页符',
      insert_image: '插入图片', manage_comment: '批注操作', get_comments: '读取批注', get_tracked_changes: '读取修订', manage_content_control: '内容控件', eval_officejs: '执行自定义代码',
      get_all_objects: '读取表格结构', read_range: '读取区域', get_workbook_overview: '工作簿总览', trace_dependencies: '追踪公式依赖', get_cell_ranges: '读取单元格', get_range_as_csv: '读取区域数据', set_cell_range: '写入单元格',
      clear_cell_range: '清除单元格', resize_range: '调整区域大小', copy_to: '复制区域', search_data: '搜索数据',
      insert_delete_rows_columns: '增删行列', hide_unhide_rows_columns: '隐藏行列', freeze_panes: '冻结窗格', view_settings: '视图设置',
      format_cells: '设置单元格格式', comments: '批注', conditional_format: '条件格式', instructions: '持久规则', undo_last_write: '撤销上次写入',
      modify_object: '修改对象', modify_sheet_structure: '修改工作表', modify_workbook_structure: '修改工作簿',
      add_slide: '新增幻灯片', delete_slide: '删除幻灯片', duplicate_slide: '复制幻灯片', get_presentation_outline: '读取演示大纲',
      get_selected_slides: '读取选中幻灯片', get_slide: '读取幻灯片', goto_slide: '跳转幻灯片', insert_textbox: '插入文本框',
      set_text: '设置文本', set_slide_notes: '设置备注', apply_layout: '调整幻灯片布局', list_layouts: '读取母版版式', recolor_slide: '统一配色字体',
      read_memory: '读取记忆', write_memory: '沉淀记忆', ask_clarification: '澄清确认'
    },
    en: {
      get_document_outline: 'Read outline', get_selection: 'Read selection', get_paragraphs: 'Read paragraphs', get_tables: 'Read tables',
      search_text: 'Search text', insert_text: 'Insert text', replace_text: 'Replace text', apply_style: 'Apply style',
      set_paragraph_format: 'Paragraph format', insert_heading: 'Insert heading', insert_table: 'Insert table', insert_page_break: 'Page break',
      insert_image: 'Insert image', manage_comment: 'Comment action', get_comments: 'Read comments', get_tracked_changes: 'Read tracked changes', manage_content_control: 'Content control', eval_officejs: 'Run custom code',
      get_all_objects: 'List sheet objects', read_range: 'Read range', get_workbook_overview: 'Workbook overview', trace_dependencies: 'Trace formula deps', get_cell_ranges: 'Read cells', get_range_as_csv: 'Read range data', set_cell_range: 'Write cells',
      clear_cell_range: 'Clear cells', resize_range: 'Resize range', copy_to: 'Copy range', search_data: 'Search data',
      insert_delete_rows_columns: 'Insert/delete rows/cols', hide_unhide_rows_columns: 'Hide/unhide rows/cols', freeze_panes: 'Freeze panes', view_settings: 'View settings',
      format_cells: 'Format cells', comments: 'Comments', conditional_format: 'Conditional format', instructions: 'Persistent rules', undo_last_write: 'Undo last write',
      modify_object: 'Modify object', modify_sheet_structure: 'Modify sheet', modify_workbook_structure: 'Modify workbook',
      add_slide: 'Add slide', delete_slide: 'Delete slide', duplicate_slide: 'Duplicate slide', get_presentation_outline: 'Read outline',
      get_selected_slides: 'Read selected slides', get_slide: 'Read slide', goto_slide: 'Go to slide', insert_textbox: 'Insert textbox',
      set_text: 'Set text', set_slide_notes: 'Set notes', apply_layout: 'Adjust slide layout', list_layouts: 'List layouts', recolor_slide: 'Recolor slide',
      read_memory: 'Read memory', write_memory: 'Save memory', ask_clarification: 'Clarify'
    }
  };
  function toolLabel(name) {
    const table = TOOL_LABELS[state.locale] || TOOL_LABELS.en;
    return table[name] || name;
  }
  // 工具卡语义短语：在动宾短语上叠加关键参数（“写入单元格 · C3:G12”而非干巴巴“写入单元格”）。
  // 字段名以各 host-*.js 的 parameters schema 为准，已逐一核对；取不到参数就退化为纯动词，不硬造。
  // PPT 的 index 是 0 起始页码（host-powerpoint.js 内部约定），展示时 +1。
  const TOOL_ARG_HINTS = {
    web_search: a => a.query,
    read_local_file: a => String(a.path || '').split(/[\\/]/).pop(),
    load_skill: a => a.name,
    extract_images: a => String(a.url || '').replace(/^https?:\/\//, ''),
    read_range: a => a.range,
    set_cell_range: a => a.range,
    clear_cell_range: a => a.range,
    copy_to: a => (a.sourceRange && a.destinationRange) ? `${a.sourceRange} → ${a.destinationRange}` : (a.sourceRange || ''),
    search_data: a => a.searchTerm,
    search_text: a => a.query,
    replace_text: a => a.query ? (state.locale === 'en' ? `"${String(a.query)}"` : `“${String(a.query)}”`) : '',
    trace_dependencies: a => a.cell,
    get_workbook_overview: a => a.sheet,
    insert_delete_rows_columns: a => dimRef(a),
    hide_unhide_rows_columns: a => dimRef(a),
    get_slide: a => slideRef(a.index),
    goto_slide: a => slideRef(a.index),
    set_text: a => slideRef(a.index),
    set_slide_notes: a => slideRef(a.index),
    apply_layout: a => slideRef(a.index),
    delete_slide: a => slideRef(a.index),
    duplicate_slide: a => slideRef(a.index)
  };
  function dimRef(a) {
    const ref = a.reference || '';
    if (!ref) return '';
    if (a.dimension === 'rows') return state.locale === 'en' ? `row ${ref}` : `第 ${ref} 行`;
    return state.locale === 'en' ? `col ${ref}` : `${ref} 列`;
  }
  function slideRef(i) {
    const n = Number(i);
    if (!Number.isInteger(n)) return '';
    return state.locale === 'en' ? `slide ${n + 1}` : `第 ${n + 1} 页`;
  }
  function toolPhrase(tc) {
    const base = toolLabel(tc.name);
    const hintFn = TOOL_ARG_HINTS[tc.name];
    if (!hintFn || !tc.args) return base;
    try {
      let hint = String(hintFn(tc.args) || '').trim();
      if (!hint) return base;
      if (hint.length > 40) hint = hint.slice(0, 40) + '…';
      return `${base} · ${hint}`;
    } catch { return base; }
  }
  const PHASE_KEYS = { thinking: 'phaseThinking', reasoning: 'phaseReasoning', reading: 'phaseReading', writing: 'phaseWriting', replying: 'phaseReplying', reviewing: 'phaseReviewing' };
  function statusLabel(s) {
    return s === 'complete' ? t('statusComplete') : s === 'error' ? t('statusError') : s === 'stopped' ? t('statusStopped') : t('statusRunning');
  }

  function allSkillsList() { return App.allSkills ? App.allSkills() : (App.SKILLS || []); }

  function skillCommand(skill) {
    return '/' + String(skill?.id || '').trim().replace(/\s+/g, '-');
  }

  function isSlashBoundary(text, index) {
    if (index <= 0) return true;
    return !/[A-Za-z0-9_:/.-]/.test(text[index - 1]);
  }

  function slashFragmentAtCursor(value, cursor) {
    const text = String(value || '');
    const end = Math.max(0, Math.min(Number(cursor) || 0, text.length));
    const prefix = text.slice(0, end);
    const start = prefix.lastIndexOf('/');
    if (start < 0 || !isSlashBoundary(text, start)) return null;
    const token = prefix.slice(start + 1);
    if (/\s|\//.test(token)) return null;
    return { token, start, end };
  }

  function completedSlashFragment(value, cursor) {
    const text = String(value || '');
    const end = Math.max(0, Math.min(Number(cursor) || 0, text.length));
    if (end < 2 || !/[ \t]/.test(text[end - 1])) return null;
    const fragment = slashFragmentAtCursor(text, end - 1);
    if (!fragment || !fragment.token) return null;
    return { token: fragment.token, start: fragment.start, end };
  }

  function removeSlashRange(text, range) {
    if (!range) return String(text || '');
    const before = String(text || '').slice(0, range.start);
    let after = String(text || '').slice(range.end);
    if (/\s$/.test(before) && /^[ \t]/.test(after)) after = after.slice(1);
    return before + after;
  }

  function extractSkillCommand(value) {
    const text = String(value || '');
    for (let start = text.indexOf('/'); start >= 0; start = text.indexOf('/', start + 1)) {
      if (!isSlashBoundary(text, start)) continue;
      let end = start + 1;
      while (end < text.length && !/\s|\//.test(text[end])) end++;
      const skill = findSkillByCommand(text.slice(start + 1, end));
      if (skill) return { skill, text: removeSlashRange(text, { start, end }).replace(/[ \t]{2,}/g, ' ').trim() };
    }
    return null;
  }

  function findSkillByCommand(token) {
    const query = String(token || '').toLowerCase();
    return allSkillsList().find(skill => String(skill.id || '').toLowerCase() === query) || null;
  }

  function slashMatches(query) {
    const q = String(query || '').toLowerCase();
    return allSkillsList().filter(skill =>
      !q
      || String(skill.id || '').toLowerCase().includes(q)
      || String(skill.name || '').toLowerCase().includes(q)
    );
  }

  function selectSlashSkill(skill) {
    if (!skill) return null;
    const range = state.slashRange;
    state.activeSkillId = skill.id;
    state.composerDraft = removeSlashRange(state.composerDraft, range);
    state.slashItems = null;
    state.slashRange = null;
    state.slashIdx = 0;
    return range ? Math.min(range.start, state.composerDraft.length) : state.composerDraft.length;
  }

  function skillManageHtml() {
    const builtin = App.SKILLS || [];
    const custom = state.settings.customSkills || [];
    const item = (s, isCustom) => `<div class="skill-manage-item">
      <div class="sm-info"><strong>${escapeHtml(s.name)}</strong><small>${escapeHtml(s.desc || '')}</small></div>
      <span class="diff-badge no">${isCustom ? t('custom') : t('builtin')}</span>
      ${isCustom ? `<button class="ghost-btn compact-btn danger-ghost" data-action="delete-skill" data-sid="${s.id}" title="${t('removeModel')}">×</button>` : ''}
    </div>`;
    return builtin.map(s => item(s, false)).join('') + custom.map(s => item(s, true)).join('')
      + `<div class="skill-import-actions">
          <label class="skill-import-btn"><span class="skill-action-icon" aria-hidden="true">${icon('upload')}</span><span>${t('importSkillFiles')}</span><input type="file" id="skill-file" accept=".md,.txt,.zip" multiple style="display:none" /></label>
          <label class="skill-import-btn"><span class="skill-action-icon" aria-hidden="true">${icon('folder')}</span><span>${t('importSkillFolder')}</span><input type="file" id="skill-folder" webkitdirectory directory multiple style="display:none" /></label>
          <button class="skill-import-btn skill-scan-btn" data-action="scan-skills"><span class="skill-action-icon" aria-hidden="true">${icon('refresh')}</span><span>${t('scanSkills')}</span></button>
        </div>
        ${state.skillImportNotice ? `<div class="skill-import-notice">${escapeHtml(state.skillImportNotice)}</div>` : ''}`;
  }

  function skillResultText(key, result) {
    return t(key)
      .replace('{added}', String(result.added || 0))
      .replace('{updated}', String(result.updated || 0))
      .replace('{existing}', String(result.existing || 0));
  }

  async function importSkillFiles(files, replaceExisting, resultKey, allowEmpty) {
    if (typeof App.readSkillInputs !== 'function' || typeof App.mergeCustomSkills !== 'function') {
      throw new Error(t('skillImportUnavailable'));
    }
    const entries = await App.readSkillInputs(files);
    if (!entries.length) {
      if (!allowEmpty) throw new Error(t('skillNotFound'));
      const emptyResult = { skills: state.settings.customSkills || [], added: 0, updated: 0, existing: 0 };
      state.skillImportNotice = skillResultText(resultKey, emptyResult);
      render();
      return emptyResult;
    }
    const result = App.mergeCustomSkills(
      state.settings.customSkills || [],
      App.SKILLS || [],
      entries,
      { replaceExisting }
    );
    state.settings.customSkills = result.skills;
    state.skillImportNotice = skillResultText(resultKey, result);
    saveSettings();
    render();
    return result;
  }

  function encodedSkillPath(filePath) {
    return String(filePath || '').split('/').filter(Boolean).map(encodeURIComponent).join('/');
  }

  async function skillFilesFromServer() {
    let paths = [];
    try {
      const response = await fetch(App.localApiUrl('/api/skills'));
      if (response.ok) {
        const payload = await response.json();
        paths = Array.isArray(payload.files) ? payload.files : [];
      }
    } catch {}

    // 兼容没有本地 API 的静态部署。
    if (!paths.length) {
      const response = await fetch('skills/index.json');
      if (!response.ok) throw new Error('skills/index.json ' + response.status);
      const payload = await response.json();
      paths = (Array.isArray(payload) ? payload : []).map(item => typeof item === 'string' ? item : item.file).filter(Boolean);
    }

    const files = [];
    for (const filePath of paths) {
      const response = await fetch('skills/' + encodedSkillPath(filePath));
      if (!response.ok) continue;
      const bytes = await response.arrayBuffer();
      const name = String(filePath).split('/').pop() || String(filePath);
      files.push({
        name,
        relativePath: filePath,
        text: async () => new TextDecoder('utf-8').decode(bytes),
        arrayBuffer: async () => bytes
      });
    }
    return files;
  }

  // ---- 模型服务管理：所有服务使用同一数据结构和交互 ----
  function providerEntries() {
    return (state.settings.serviceProviders || []).map(provider => ({
      id: provider.id,
      name: provider.name,
      baseUrl: provider.baseUrl,
      key: provider.apiKey || '',
      protocol: provider.protocol || 'openai',
      models: provider.models || []
    }));
  }

  function providerEditFormHtml(e) {
    const isNew = !e;
    const v = e || { name: '', baseUrl: '', key: '' };
    return `<div class="prov-edit">
      <div class="pe-title">${isNew ? t('newServiceTitle') : t('editService')}</div>
      <div class="field"><label>${t('serviceName')}</label><input id="pe-name" value="${escapeHtml(v.name || '')}" placeholder="${escapeHtml(t('serviceNameHint'))}" /></div>
      <div class="field"><label>${t('baseUrlLabel')}</label><input id="pe-url" value="${escapeHtml(v.baseUrl || '')}" placeholder="https://.../v1" /></div>
      <div class="field"><label>${t('apiKeyLabel')}</label><input id="pe-key" type="password" value="${escapeHtml(v.key || '')}" placeholder="sk-..." /></div>
      <div class="field"><label>${t('protocolLabel')}</label><select id="pe-protocol">
        <option value="openai" ${(v.protocol || 'openai') === 'openai' ? 'selected' : ''}>${t('protocolOpenai')}</option>
        <option value="anthropic" ${v.protocol === 'anthropic' ? 'selected' : ''}>${t('protocolAnthropic')}</option>
      </select></div>
      <div class="hint">${t('protocolHint')}</div>
      <div class="hint">${t('saveThenManageModels')}</div>
      <div class="prov-edit-actions">
        <button class="primary-btn" data-action="save-provider" data-pid="${e ? escapeHtml(e.id) : ''}">${t('saveService')}</button>
        <button class="ghost-btn" data-action="cancel-edit">${t('cancelEdit')}</button>
        ${isNew ? '' : `<button class="ghost-btn danger-ghost" data-action="delete-provider" data-pid="${escapeHtml(e.id)}">${t('deleteService')}</button>`}
      </div>
    </div>`;
  }

  function providerDomId(providerId) {
    return String(providerId || '').replace(/[^a-zA-Z0-9_-]/g, '_');
  }

  function providerLibraryIds(entry) {
    const values = [];
    for (const model of (entry.models || [])) values.push(model.id || model);
    values.push(...(((state.settings.knownModels || {})[entry.id]) || []));
    values.push(...(((state.settings.providerModels || {})[entry.id]) || []));
    if (state.settings.provider === entry.id && state.settings.model) values.push(state.settings.model);
    return Array.from(new Set(values.map(id => String(id || '').trim()).filter(Boolean)))
      .filter(id => !App.isRetiredModel || !App.isRetiredModel(entry.id, id));
  }

  function providerModelManagerHtml(entry) {
    const library = providerLibraryIds(entry);
    const selected = ((state.settings.providerModels || {})[entry.id] || []);
    const inputId = `provider-model-input-${providerDomId(entry.id)}`;
    return `<div class="provider-model-manager">
      <div class="provider-model-head"><strong>${t('providerModelLibrary')}</strong><span>${selected.length} ${t('selectedCount')}</span></div>
      <div class="add-model-row provider-add-model-row">
        <input id="${inputId}" placeholder="${escapeHtml(t('addModelHint'))}" />
        <button class="ghost-btn" data-action="add-provider-model" data-pid="${escapeHtml(entry.id)}">${icon('plus','sm')} ${t('addModel')}</button>
        <button class="ghost-btn" data-action="fetch-provider-models" data-pid="${escapeHtml(entry.id)}">${icon('download','sm')} ${t('fetchModels')}</button>
      </div>
      <div class="hint provider-model-hint">${t('providerModelHint')}</div>
      ${library.length ? `<div class="settings-model-list provider-library-list">
        ${library.map(id => `<button class="model-item provider-library-item ${selected.includes(id) ? 'active' : ''}" data-action="toggle-provider-model" data-pid="${escapeHtml(entry.id)}" data-model="${escapeHtml(id)}">${App.modelLogoHtml(id)}<span class="mi-name">${escapeHtml(id)}</span>${selected.includes(id) ? `<span class="check">${icon('check','sm')}</span>` : ''}</button>`).join('')}
      </div>` : `<div class="provider-model-empty">${t('noProviderModels')}</div>`}
    </div>`;
  }

  function providerCardsHtml() {
    const cur = state.settings.provider;
    const editing = state.providerEditing;
    const managing = state.providerModelManaging;
    let html = providerEntries().map(e => {
      const active = e.id === cur;
      const host = (e.baseUrl || '').replace(/^https?:\/\//, '').split('/')[0] || '—';
      return `
      <div class="prov-card ${active ? 'active' : ''}">
        <div class="prov-row">
          <button class="prov-main" data-action="use-provider" data-pid="${escapeHtml(e.id)}">
            <span class="prov-dot ${e.key ? 'ok' : ''}"></span>
            <span class="prov-info"><strong>${escapeHtml(e.name)}</strong><small>${escapeHtml(host)} · ${e.key ? t('presetKey') : t('noKey')}</small></span>
            ${active ? `<span class="prov-cur">${icon('check','sm')} ${t('currentService')}</span>` : ''}
          </button>
          <button class="ghost-btn compact-btn icon-only ${managing === e.id ? 'active' : ''}" data-action="manage-provider-models" data-pid="${escapeHtml(e.id)}" title="${t('manageModels')}">${icon('listView','sm')}</button>
          <button class="ghost-btn compact-btn icon-only" data-action="edit-provider" data-pid="${escapeHtml(e.id)}" title="${t('editService')}">${icon('edit','sm')}</button>
        </div>
        ${editing === e.id ? providerEditFormHtml(e) : ''}
        ${managing === e.id ? providerModelManagerHtml(e) : ''}
      </div>`;
    }).join('');
    if (editing === '__new__') html += providerEditFormHtml(null);
    html += editing
      ? `<button class="ghost-btn" data-action="cancel-edit">${t('cancelEdit')}</button>`
      : `<button class="add-service-btn" data-action="new-provider">${icon('plus','sm')} ${t('addService')}</button>`;
    return html;
  }

  const PANEL_SELECTORS = ['.settings-panel'];
  function captureUiState() {
    const snap = { focused: null, atBottom: true, scrollTop: 0, panels: {} };
    try {
      const el = document.activeElement;
      if (el && (el.id === 'chat-input' || el.id === 'manual-args' || el.id === 'skill-search')) {
        snap.focused = { selector: '#' + el.id, value: el.value, start: el.selectionStart, end: el.selectionEnd };
      } else if (el && el.dataset && el.dataset.bind) {
        const canSelect = el.tagName === 'INPUT' || el.tagName === 'TEXTAREA';
        snap.focused = { selector: `[data-bind="${el.dataset.bind}"]`, value: canSelect ? el.value : undefined, start: canSelect ? el.selectionStart : null, end: canSelect ? el.selectionEnd : null };
      }
      const msgEl = document.getElementById('messages');
      if (msgEl) {
        snap.scrollTop = msgEl.scrollTop;
        snap.atBottom = msgEl.scrollHeight - msgEl.scrollTop - msgEl.clientHeight < 40;
      }
      for (const sel of PANEL_SELECTORS) {
        const panel = document.querySelector(sel);
        if (panel) snap.panels[sel] = panel.scrollTop;
      }
    } catch {}
    return snap;
  }
  function restoreUiState(snap) {
    if (!snap) return;
    if (snap.focused) {
      const el = document.querySelector(snap.focused.selector);
      if (el) {
        if (snap.focused.value !== undefined && 'value' in el) el.value = snap.focused.value;
        el.focus();
        if (snap.focused.start != null) { try { el.setSelectionRange(snap.focused.start, snap.focused.end ?? snap.focused.start); } catch {} }
      }
    }
    const msgEl = document.getElementById('messages');
    // 用户已经离开底部时尊重其阅读位置；只有原本贴底时才继续跟随新内容。
    if (msgEl) msgEl.scrollTop = snap.atBottom ? msgEl.scrollHeight : snap.scrollTop;
    if (snap.panels) {
      for (const sel of PANEL_SELECTORS) {
        const panel = document.querySelector(sel);
        if (panel && snap.panels[sel] != null) panel.scrollTop = snap.panels[sel];
      }
    }
  }

  // 品牌 logo 按主题切换：暗色用白色马头剪影（icon-128-dark.png，浅底转透明生成），
  // 亮色用原版（用户 9-05：暗色下黑色马头糊成一块；原「filter 压白」会把浅色圆片底一起变白失去形状）
  function brandLogoSrc() {
    return document.documentElement.dataset.theme === 'dark' ? 'assets/brand/icon-128-dark.png' : 'assets/brand/icon-128.png';
  }

  function render() {
    const snap = captureUiState();
    document.documentElement.dataset.theme = state.theme;
    document.documentElement.dataset.fontsize = String(Number(state.settings.uiFontSize) || 3);
    ensureSession();
    const configured = Boolean(state.settings.apiKey && state.settings.model && state.settings.customPrefixUrl);
    const allModels = getProviderModels(state.settings.provider);
    const currentProviderEntry = providerEntries().find(entry => entry.id === state.settings.provider);
    const currentProviderHost = ((currentProviderEntry && currentProviderEntry.baseUrl) || '').replace(/^https?:\/\//, '').split('/')[0] || '—';
    const selectedModelIds = ((state.settings.providerModels || {})[state.settings.provider] || []);
    const models = selectedModelIds.map(id => allModels.find(model => model.id === id) || { id, name: id });
    const toolDefs = (App.host && App.host.toolDefinitions) || [];
    const hasSelectionTool = Boolean(App.host && App.host.toolExecutors && App.host.toolExecutors.get_selection);
    const changeLogItems = App.buildChangeLogItems ? App.buildChangeLogItems(state.messages) : [];
    const changeLogBarHtml = changeLogItems.length && App.changeLogHtml ? App.changeLogHtml(changeLogItems) : '';
    const defaultArgsForTool = (App.host && App.host.defaultArgsForTool) || (() => '{}');
    const brand = t('brand') === 'brand' ? 'Trojan AI for Office' : t('brand');
    const brandFooter = t('brandFooter') === 'brandFooter' ? 'Trojan AI for Office' : t('brandFooter');
    const visibleSessions = sessionsForCurrentWorkbook();
    const currentModelEntry = allModels.find(m => m.id === state.settings.model);
    const currentModelName = currentModelEntry ? (currentModelEntry.name || currentModelEntry.id) : '';
    const mlogo = App.modelLogoHtml ? App.modelLogoHtml(state.settings.model) : '';
    const activeEditMode = typeof App.currentAccessMode === 'function' ? App.currentAccessMode() : (state.settings.accessMode || 'confirm');
    const advancedAccess = typeof App.currentAdvancedAccess === 'function' ? App.currentAdvancedAccess() : false;
    const memoryStats = (App.memory && typeof App.memory.stats === 'function') ? App.memory.stats() : null;
    const accessModes = [
      { id: 'confirm', label: t('accessConfirm'), icon: 'shield' },
      { id: 'auto', label: t('accessAuto'), icon: 'zap' },
      { id: 'plan', label: t('accessPlan'), icon: 'listView' }
    ];
    const currentAccessMode = accessModes.find(x => x.id === activeEditMode) || accessModes[0];
    const thinkingLevels = [
      { id: 'none', label: t('thinkingOff') },
      { id: 'low', label: t('thinkingLow') },
      { id: 'medium', label: t('thinkingMedium') },
      { id: 'high', label: t('thinkingHigh') }
    ];
    const currentThinking = thinkingLevels.find(x => x.id === (state.settings.thinking || 'medium')) || thinkingLevels[2];
    const activeSkill = allSkillsList().find(s => s.id === state.activeSkillId);
    const activeSkillName = activeSkill ? activeSkill.name : '';
    const skillQ = (state.skillSearch || '').toLowerCase();
    const filteredSkills = allSkillsList().filter(s => !skillQ || (s.name || '').toLowerCase().includes(skillQ) || (s.id || '').toLowerCase().includes(skillQ));
    const sessionsHtml = visibleSessions.length ? visibleSessions.map(s => {
      const pendingDelete = state.pendingDeleteSessionId === s.id;
      return `
      <div class="session-item ${s.id === state.currentSessionId ? 'active' : ''} ${pendingDelete ? 'confirming-delete' : ''}">
        <button data-action="switch-session" data-id="${s.id}" title="${escapeHtml(s.name)}" ${pendingDelete ? 'disabled' : ''}><span>${escapeHtml(s.name)}</span><small>${new Date(s.updatedAt).toLocaleString()}</small></button>
        ${pendingDelete ? `<div class="delete-confirm"><small>${t('confirmDeleteSession')}</small><div><button class="confirm-delete-btn" data-action="confirm-delete-session" data-id="${s.id}">${t('confirm')}</button><button class="cancel-delete-btn" data-action="cancel-delete-session" data-id="${s.id}">${t('cancel')}</button></div></div>` : `<button class="delete-btn" data-action="delete-session" data-id="${s.id}" title="${t('deleteSession')}">×</button>`}
      </div>`;
    }).join('') : `<div class="session-empty">${t('noSessions')}</div>`;
    const messagesHtml = state.messages.length ? state.messages.map(renderMessage).join('') + phaseBarHtml() : `
      <div class="empty">
        <div class="empty-glow"><img class="logo" src="${brandLogoSrc()}" alt="logo" /></div>
        <h2>${t('title')}</h2>
        <p>${t('subtitle')}</p>
        ${App.host && App.host.available === false && (typeof Office !== 'undefined' && Office.context) ? `<div class="host-warning">${escapeHtml(t('hostNotLoaded'))}（office.js 来源：${window.__officeJsSource__ || '未加载'}）</div>` : ''}
        <div class="prompt-grid">
          <button class="prompt-card" data-prompt="${escapeHtml(t('chartPrompt'))}"><strong>${t('chart')}</strong><span>${t('chartDesc')}</span></button>
          <button class="prompt-card" data-prompt="${escapeHtml(t('fixPrompt'))}"><strong>${t('fix')}</strong><span>${t('fixDesc')}</span></button>
          <button class="prompt-card" data-prompt="${escapeHtml(t('analyzePrompt'))}"><strong>${t('analyze')}</strong><span>${t('analyzeDesc')}</span></button>
        </div>
      </div>`;
    document.getElementById('container').innerHTML = `
      <div class="app">
        <header class="header">
          <img class="logo" src="${brandLogoSrc()}" alt="logo" />
          <div class="title"><strong>${escapeHtml(brand)}</strong><span>${escapeHtml(state.workbookLabel || (configured ? t('configuredShort') : t('notConfiguredShort')))}</span></div>
          <div class="session-menu-wrap">
            <button class="session-trigger" data-action="toggle-session-menu" title="${t('sessionHistory')}"><span>${t('sessionHistory')}</span></button>
            ${state.sessionMenuOpen ? `<div class="session-menu">
              <div class="session-menu-head"><strong>${t('sessionHistory')}</strong><button class="secondary-btn compact" data-action="new-chat">${icon('plus','sm')} ${t('newChat')}</button></div>
              <div class="session-list header-session-list">${sessionsHtml}</div>
            </div>` : ''}
          </div>
          <button class="icon-btn" data-action="toggle-theme" title="${state.theme === 'dark' ? t('light') : t('dark')}">${state.theme === 'dark' ? icon('sun') : icon('moon')}</button>
          <button class="icon-btn" data-action="toggle-locale" title="language">${t('lang')}</button>
        </header>
        <nav class="tabs">
          <button class="tab ${state.tab === 'chat' ? 'active' : ''}" data-tab="chat">${t('chat')}</button>
          <button class="tab ${state.tab === 'settings' ? 'active' : ''}" data-tab="settings">${t('settings')}</button>
        </nav>
        <main class="main">
          <section class="panel chat-panel ${state.tab === 'chat' ? 'active' : ''}">
            ${changeLogBarHtml}
            <div class="messages" id="messages">${state.recoveredTaskNotice ? `<div class="recovery-banner"><span>${t('recoveredTaskNotice')}</span><button data-action="dismiss-recovery" title="${t('dismissRecovery')}">×</button></div>` : ''}${messagesHtml}${state.reviewRunning ? `<div class="review-running">${t('reviewRunning')}</div>` : ''}</div>
            <div class="composer">
              <div class="drag-bar" data-drag="composer" title="⇕"></div>
              ${state.storageQuotaWarning ? `<div class="error">${t('storageQuotaWarning')}</div>` : ''}
              ${state.settingsSyncWarning ? `<div class="error">${t('settingsSyncWarning')}</div>` : ''}
              ${state.settingsSyncUnsupported ? `<div class="error">${t('settingsSyncUnsupported')}</div>` : ''}
              ${state.settingsSyncConflict ? `<div class="error">${t('settingsSyncConflict')} <button class="settings-adopt-btn" data-action="adopt-shared-settings">${t('adoptSharedSettings')}</button></div>` : ''}
              ${state.error ? `<div class="error">${escapeHtml(state.error)}</div>` : ''}
              <div class="quickbar">
                <div class="quick-model-wrap">
                  <button class="quick-model" data-action="toggle-model-menu" title="${t('quickModel')}">${mlogo || `<span class="qm-label">${t('quickModel')}</span>`}<strong>${escapeHtml(currentModelName || state.settings.model || '')}</strong><span class="chev">${icon('chevronDown','sm')}</span></button>
                  ${state.modelMenuOpen ? `<div class="model-menu">
                    ${models.length ? models.map(m => `<button class="model-item ${state.settings.model === m.id ? 'active' : ''}" data-action="pick-quick-model" data-model="${escapeHtml(m.id)}">${App.modelLogoHtml(m.id)}<span class="mi-name">${escapeHtml(m.name || m.id)}</span>${state.settings.model === m.id ? `<span class="check">${icon('check','sm')}</span>` : ''}</button>`).join('') : `<div class="provider-model-empty">${t('noSelectedModels')}</div>`}
                    <div class="model-menu-foot">${t('moreModels')}</div>
                  </div>` : ''}
                </div>
                <div class="quick-access-wrap">
                  <button class="quick-model quick-access" data-action="toggle-access-menu" title="${t('accessMode')}"><span class="mode-symbol">${icon(currentAccessMode.icon,'sm')}</span><strong>${escapeHtml(currentAccessMode.label)}</strong><span class="chev">${icon('chevronDown','sm')}</span></button>
                  ${state.accessMenuOpen ? `<div class="model-menu access-menu">
                    ${accessModes.map(m => `<button class="model-item ${m.id === currentAccessMode.id ? 'active' : ''}" data-action="set-access-mode" data-mode="${m.id}"><span>${icon(m.icon,'sm')} ${escapeHtml(m.label)}</span>${m.id === currentAccessMode.id ? `<span class="check">${icon('check','sm')}</span>` : ''}</button>`).join('')}
                  </div>` : ''}
                </div>
                <div class="quick-thinking-wrap">
                  <button class="quick-model quick-thinking" data-action="toggle-thinking-menu" title="${t('thinking')}">${thinkingMeterHtml(currentThinking.id)}<strong>${escapeHtml(currentThinking.label)}</strong><span class="chev">${icon('chevronDown','sm')}</span></button>
                  ${state.thinkingMenuOpen ? `<div class="model-menu access-menu">
                    ${thinkingLevels.map(m => `<button class="model-item thinking-item ${m.id === currentThinking.id ? 'active' : ''}" data-action="set-thinking" data-level="${m.id}"><span>${thinkingMeterHtml(m.id)}${escapeHtml(m.label)}</span>${m.id === currentThinking.id ? `<span class="check">${icon('check','sm')}</span>` : ''}</button>`).join('')}
                  </div>` : ''}
                </div>
              </div>
              <div class="composer-box">
                ${(state.activeSkillId || state.selectionSnapshot || (state.attachments && state.attachments.length) || (state.pendingImages && state.pendingImages.length) || state.fileLoading) ? `<div class="chip-row">
                  ${state.selectionSnapshot ? `<span class="bubble file-bubble sel-bubble" title="${escapeHtml(String(state.selectionSnapshot.text || state.selectionSnapshot.paragraphText || ''))}"><span class="bubble-ico">${icon('locate','sm')}</span>${escapeHtml(promptSummary(state.selectionSnapshot.text || state.selectionSnapshot.paragraphText) || t('selEmpty'))}<button data-action="refresh-selection" title="${t('selRefresh')}">↻</button><button data-action="clear-selection" title="${t('cancelEdit')}">×</button></span>` : ''}
                  ${state.activeSkillId ? `<span class="bubble skill-bubble"><span class="bubble-ico">${icon('sparkle','sm')}</span>${escapeHtml(activeSkillName)}<button data-action="clear-skill" title="${t('cancelEdit')}">×</button></span>` : ''}
                  ${state.fileLoading ? `<span class="bubble">${icon('clock','sm')} ${escapeHtml(state.fileLoading)}…</span>` : ''}
                  ${(state.attachments || []).map((a, i) => `<span class="bubble file-bubble"><span class="bubble-ico">${icon('file','sm')}</span>${escapeHtml(a.name)}<button data-action="remove-attachment" data-idx="${i}" title="${t('removeModel')}">×</button></span>`).join('')}
                  ${(state.pendingImages || []).map((a, i) => `<span class="bubble file-bubble" title="${escapeHtml(a.id)}"><span class="bubble-ico">${icon('image','sm')}</span>${escapeHtml(a.name)}<button data-action="remove-pending-image" data-idx="${i}" title="${t('removeModel')}">×</button></span>`).join('')}
                </div>` : ''}
                ${state.slashItems && state.slashItems.length ? `<div class="slash-menu">
                  ${state.slashItems.map((s, i) => `<button class="slash-item ${i === state.slashIdx ? 'active' : ''}" data-action="slash-pick" data-skill="${s.id}"><span class="slash-command">${escapeHtml(skillCommand(s))}</span><span class="si-name">${escapeHtml(s.name)}</span><span class="skill-desc">${escapeHtml(s.desc || '')}</span></button>`).join('')}
                </div>` : ''}
                ${(state.steerQueue || []).length ? `<div class="held-message-list">${state.steerQueue.map((q, i) => `<div class="held-message"><span>${escapeHtml(promptSummary(q.text))}</span><div><button class="held-send" data-action="steer-held" data-idx="${i}" title="${t('composeSteer')}">${icon('enter')}</button><button data-action="edit-held" data-idx="${i}" title="${t('composeEditHeld')}">×</button></div></div>`).join('')}</div>` : ''}
                <textarea id="chat-input" rows="2" style="${state.settings.composerHeight ? `height:${state.settings.composerHeight}px` : ''}" placeholder="${configured ? t('input') : t('noConfig')}" ${!configured ? 'disabled' : ''}>${escapeHtml(state.composerDraft || '')}</textarea>
                <div class="composer-actions">
                  ${hasSelectionTool ? `<button class="plus-btn sel-toggle ${state.selectionTarget ? 'active' : ''}" data-action="toggle-selection-target" title="${t('selTargetTitle')}">${icon('locate','lg')}</button>` : ''}
                  <div class="plus-wrap">
                    <button class="plus-btn" data-action="toggle-plus" title="+">${icon('plus','lg')}</button>
                    ${state.plusMenuOpen ? `<div class="plus-menu">
                      <button class="plus-item" data-action="plus-skills"><span class="pi-ico">${icon('sparkle')}</span><span>${t('plusSkills')}</span><span class="pi-chev">${state.plusSkillsOpen ? icon('chevronDown','sm') : icon('chevronRight','sm')}</span></button>
                      ${state.plusSkillsOpen ? `<div class="plus-skills-list">
                        <input id="skill-search" placeholder="${escapeHtml(t('skillSearchHint'))}" value="${escapeHtml(state.skillSearch || '')}" />
                        <div class="plus-skill-items">${filteredSkills.map(s => `<button class="plus-skill-item ${state.activeSkillId === s.id ? 'active' : ''}" data-action="pick-skill" data-skill="${s.id}"><span class="si-name">${escapeHtml(s.name)}${state.activeSkillId === s.id ? ' ' + icon('check','sm') : ''}</span><span class="skill-desc">${escapeHtml(s.desc || '')}</span></button>`).join('') || `<div class="session-empty">—</div>`}</div>
                      </div>` : ''}
                      <label class="plus-item"><span class="pi-ico">${icon('file')}</span><span>${t('plusUpload')}</span><input type="file" id="file-input" accept=".pdf,.docx,.pptx,.txt,.md,.csv" multiple style="display:none" /></label>
                      <label class="plus-item"><span class="pi-ico">${icon('image','sm')}</span><span>${t('plusUploadImage')}</span><input type="file" id="image-input" accept="image/png,image/jpeg,image/gif" multiple style="display:none" /></label>
                    </div>` : ''}
                  </div>
                  ${state.isWorking
                    ? `${String(state.composerDraft || '').trim() ? `<button class="send-btn icon-only" data-action="hold-draft" title="${t('composeQueueHint')}">${icon('enter')}</button>` : ''}<button class="stop-btn compact-stop" data-action="stop" title="${t('stop')}">${icon('stop','sm')}</button>`
                    : `<button class="send-btn icon-only" data-action="send" ${!configured ? 'disabled' : ''} title="${t('send')}">${icon('enter')}</button>`}
                </div>
              </div>
            </div>
          </section>
          <section class="panel settings-panel ${state.tab === 'settings' ? 'active' : ''}">
            <div class="section">
              <h3>${t('modelServices')}</h3>
              ${providerCardsHtml()}
            </div>
            <div class="section">
              <h3>${t('advancedSection')}</h3>
              <div class="current-provider-summary">
                <span><strong>${escapeHtml((currentProviderEntry && currentProviderEntry.name) || '')}</strong><small>${escapeHtml(currentProviderHost)}</small></span>
                <span class="prov-cur">${icon('check','sm')} ${t('currentService')}</span>
              </div>
              <label>${t('selectedModels')}</label>
              ${models.length ? `<div class="settings-model-list">
                ${models.map(m => `<button class="model-item ${state.settings.model === m.id ? 'active' : ''}" data-action="pick-quick-model" data-model="${escapeHtml(m.id)}">${App.modelLogoHtml(m.id)}<span class="mi-name">${escapeHtml(m.name || m.id)}</span>${state.settings.model === m.id ? `<span class="check">${icon('check','sm')}</span>` : ''}</button>`).join('')}
              </div>` : `<div class="provider-model-empty">${t('noSelectedModels')}</div>`}
              <div class="hint">${t('selectedModelsHint')}</div>
              <div class="field"><label>${t('thinking')}</label><select data-bind="thinking"><option value="none" ${state.settings.thinking === 'none' ? 'selected' : ''}>${t('thinkingOff')}</option><option value="low" ${state.settings.thinking === 'low' ? 'selected' : ''}>${t('thinkingLow')}</option><option value="medium" ${state.settings.thinking === 'medium' ? 'selected' : ''}>${t('thinkingMedium')}</option><option value="high" ${state.settings.thinking === 'high' ? 'selected' : ''}>${t('thinkingHigh')}</option></select></div>
              <label class="check"><input type="checkbox" data-bind="followMode" ${state.settings.followMode ? 'checked' : ''}/> ${t('followOn')}</label>
              <div class="hint">${configured ? t('configured') : t('notConfigured')}</div>
            </div>
            <div class="section">
              <h3>${t('customInstructions')}</h3>
              <div class="field"><textarea data-bind="customInstructions" rows="8" placeholder="${escapeHtml((App.DEFAULT_INSTRUCTIONS || '').slice(0, 80))}…">${escapeHtml(state.settings.customInstructions || '')}</textarea></div>
              <div class="hint">${t('customInstructionsHint')}</div>
            </div>
            ${memoryStats ? `<div class="section">
              <h3>${t('memorySection')}</h3>
              <div class="hint" style="margin-bottom:10px">${memoryStats.enabled ? t('memoryDesc') : t('memoryDisabled')}</div>
              <label class="check"><input type="checkbox" data-action="toggle-memory" ${memoryStats.enabled ? 'checked' : ''}/> ${t('memoryEnabledLabel')}</label>
              ${memoryStats.enabled ? (memoryStats.count ? `<div class="settings-model-list memory-list">
                ${memoryStats.topics.map(tp => { const preview = (App.memory.readTopicText(tp.key) || '').replace(/\s+/g, ' '); return `<div class="memory-item">
                  <div class="memory-item-head"><strong>${escapeHtml(tp.title)}</strong><small>${escapeHtml(tp.dateLabel || '')} · ${tp.chars}c</small><button class="memory-del" data-action="delete-memory" data-key="${escapeHtml(tp.key)}" title="${escapeHtml(t('memoryDelete'))}">×</button></div>
                  ${tp.summary ? `<div class="memory-sum">${escapeHtml(tp.summary)}</div>` : ''}
                  <div class="memory-preview">${escapeHtml(preview.slice(0, 120))}${preview.length > 120 ? '…' : ''}</div>
                  <div class="memory-triggers"><label>${escapeHtml(t('memoryTriggersLabel'))}</label><input type="text" data-action="edit-memory-triggers" data-key="${escapeHtml(tp.key)}" value="${escapeHtml((tp.triggers || []).join(', '))}" placeholder="${escapeHtml(t('memoryTriggersPlaceholder'))}" /></div>
                </div>`; }).join('')}
                <div class="memory-meta">${t('memoryCount').replace('{count}', memoryStats.count).replace('{max}', memoryStats.maxTopics).replace('{chars}', memoryStats.chars)}</div>
                <button class="memory-clear" data-action="clear-memory">${t('memoryClear')}</button>
              </div>` : `<div class="provider-model-empty">${t('memoryEmpty')}</div>`) : ''}
            </div>` : ''}
            <div class="section">
              <h3>${t('reviewSection')}</h3>
              <div class="field"><label>${t('reviewModelLabel')}</label><input type="text" data-bind="reviewModel" value="${escapeHtml(state.settings.reviewModel || '')}" placeholder="${escapeHtml(t('reviewModelPlaceholder'))}" /></div>
              <div class="hint">${t('reviewModelHint')}</div>
            </div>
            <div class="section">
              <h3>${t('accessMode')}</h3>
              <div class="access-mode-row">
                <button class="access-opt ${activeEditMode === 'confirm' ? 'active' : ''}" data-action="set-access-mode" data-mode="confirm"><strong>${icon('shield','sm')} ${t('accessConfirm')}</strong><small>${t('accessConfirmDesc')}</small></button>
                <button class="access-opt ${activeEditMode === 'auto' ? 'active' : ''}" data-action="set-access-mode" data-mode="auto"><strong>${icon('zap','sm')} ${t('accessAuto')}</strong><small>${t('accessAutoDesc')}</small></button>
                <button class="access-opt ${activeEditMode === 'plan' ? 'active' : ''}" data-action="set-access-mode" data-mode="plan"><strong>${icon('listView','sm')} ${t('accessPlan')}</strong><small>${t('accessPlanDesc')}</small></button>
                <button class="access-opt advanced-access ${advancedAccess ? 'active' : ''}" data-action="toggle-advanced-access"><strong>${icon('bolt','sm')} ${t('advancedAccess')} ${advancedAccess ? icon('check','sm') : ''}</strong><small>${t('advancedAccessDesc')}</small></button>
              </div>
            </div>
            <div class="section">
              <h3>${t('fontSizeSection')}</h3>
              <input type="range" id="font-size-range" min="1" max="5" step="1" value="${Number(state.settings.uiFontSize) || 3}" />
              <div class="fs-labels">${t('fontSizeLabels').split(',').map((x, i) => `<span>${Number(state.settings.uiFontSize) || 3 === i + 1 ? '' : ''}${x.trim()}</span>`).join('')}</div>
            </div>
            <div class="section">
              <h3>${t('skillManage')}</h3>
              <div class="hint" style="margin-bottom:10px">${t('skillManageHint')}</div>
              ${skillManageHtml()}
            </div>
            <div class="section">
              <h3>${t('webSearch')}</h3>
              <label class="check"><input type="checkbox" data-bind="searchEnabled" ${state.settings.searchEnabled ? 'checked' : ''}/> ${t('searchEnabled')}</label>
              <div class="field"><label>${t('tavilyKey')}</label><input data-bind="tavilyKey" value="${escapeHtml(state.settings.tavilyKey || '')}" type="password" autocomplete="new-password" autocapitalize="off" spellcheck="false" placeholder="tvly-..." /></div>
              ${state.settings.searchEnabled && !String(state.settings.tavilyKey || '').trim() ? `<div class="hint warn-hint">${t('searchNeedsKey')}</div>` : ''}
              <div class="hint">${t('webSearchHint')}</div>
            </div>
            <div class="section"><h3>${t('about')}</h3><p class="hint">${t('aboutText')}</p>${hasOffice() ? '' : `<p class="hint">${t('demo')}</p>`}</div>
          </section>

        </main>
        <footer class="footer"><span>${escapeHtml(brandFooter)}</span><span>${escapeHtml(state.settings.model || '')}</span></footer>
      </div>`;
    restoreUiState(snap);
  }

  function phaseBarHtml() {
    if (!state.isWorking || !state.workPhase) return '';
    const key = PHASE_KEYS[state.workPhase] || 'phaseThinking';
    return `<div class="phase-bar" data-phase="${escapeHtml(state.workPhase)}" role="status"><span class="phase-dots" aria-hidden="true"><i></i><i></i><i></i></span><span>${t(key)}…</span></div>`;
  }

  function thinkingMeterHtml(level) {
    const safe = ['none', 'low', 'medium', 'high'].includes(level) ? level : 'medium';
    return `<span class="thinking-meter level-${safe}" aria-hidden="true"><i></i><i></i><i></i></span>`;
  }

  function proposalCitation(edit) {
    const target = edit && edit.target;
    if (!target || typeof target !== 'object') return '';
    const hostType = App.host && App.host.hostType;
    if (hostType === 'word' || Number.isInteger(Number(target.paragraphIndex))) {
      const index = Number(target.paragraphIndex);
      return Number.isInteger(index) && index >= 0 ? `p:${index}` : '';
    }
    if (hostType === 'powerpoint' || target.slideId != null || target.index != null) {
      if (target.slideId) return `id:${target.slideId}`;
      const index = Number(target.index);
      return Number.isInteger(index) && index >= 0 ? `s:${index}` : '';
    }
    if (hostType === 'excel' || target.sheetId != null) {
      const sheetId = Number(target.sheetId);
      const range = String(target.range || '').trim();
      return Number.isFinite(sheetId) && range ? `${sheetId}!${range}` : '';
    }
    return '';
  }

  async function navigateProposalEdit(edit) {
    const ref = proposalCitation(edit);
    if (!ref) return;
    // 宿主返回 success:false（导航没成功但不抛错）时也要给用户反馈，不能静默无事发生
    const result = await navigateCitation(ref);
    if (result && result.success === false) showTransientNotice(t('navFailed') || '跳转失败：目标位置可能已变化，请让 AI 重新读取后再试。');
  }

  // 【34.4】写入结果与跟随导航分离：applyEdit 已确认成功就据实结算卡片，
  // 导航只是辅助——挂起不拖住结算，失败/迟迟未完成只给定位提示，
  // 绝不让已成功的卡回到「可重试写入」。slowMs 供测试缩短慢提示计时。
  function followInBackground(out, slowMs) {
    if (!out || out.success === false) return;
    let finished = false;
    const slowTimer = setTimeout(() => { if (!finished) showTransientNotice(t('followSlowHint')); }, typeof slowMs === 'number' ? slowMs : 10000);
    Promise.resolve().then(() => maybeFollow(out)).then(
      () => { finished = true; clearTimeout(slowTimer); },
      () => { finished = true; clearTimeout(slowTimer); showTransientNotice(t('followFailed')); }
    );
  }

  // ---- Diff 卡：修改提案（学 Claude add-in 的 proposal 交互）----
  function findProposalByRid(rid) {
    for (const m of state.messages) {
      if (m.toolCalls) {
        const tc = m.toolCalls.find(x => x.rid === rid && x.name === 'propose_edits');
        if (tc) return tc;
      }
    }
    return null;
  }

  // 当前正在等待用户决策的提案（agent loop 串行，同一时刻最多一个）。
  // 【34.2】带显式 rid 的用户操作必须精确命中那张卡本身：rid 找不到、或该卡已没有
  // 决策解析器（已结算/已过期/插件重开/会话切换），都不能退回「最近一张有解析器的卡」——
  // 旧 rid 曾因此被路由到新卡（点旧卡应用了新卡）。失败给用户明确提示，不静默退出。
  // 无 rid 的调用（内部兜底）才允许扫描最近未决提案。opts.silent 用于 input 等高频事件。
  function findPendingProposal(rid, opts) {
    if (rid) {
      const byRid = findProposalByRid(rid);
      if (byRid && byRid._resolveProposal) return byRid;
      if (!opts || !opts.silent) showTransientNotice(`${t('proposalExpired')}：${t('proposalExpiredHint')}`);
      return null;
    }
    for (let i = state.messages.length - 1; i >= 0; i--) {
      const m = state.messages[i];
      if (m.toolCalls) {
        const tc = m.toolCalls.find(x => x.name === 'propose_edits' && x._resolveProposal);
        if (tc) return tc;
      }
    }
    return null;
  }

  // 【34.3】「这张卡处于什么阶段、能否操作」的唯一判断，渲染与事件分派共用。
  // 失效只看卡自身：没有决策解析器且不在 enrich 准备期 = 过期；
  // 不再被全局 isWorking 赦免（旧逻辑下别的任务在跑，孤儿卡会恢复成可点击假按钮）。
  // applying/preparing 是独立阶段：显示处理中/准备中并禁用互斥操作，不残留假按钮。
  // 【35-R1】interrupted = 恢复出来的在途卡（持久化丢了解析器）：写入结果未知，
  // 据实终态显示「已中断」，不标成功/失败、不提供重试。applying 无解析器同理
  //（双保险：即使某条路径跳过恢复归一化，阶段判断也不会把死卡显示成处理中）。
  function proposalCardPhase(tc) {
    if (!tc || tc.name !== 'propose_edits') return 'expired';
    if (tc.batch) {
      if (tc.batch.settled) return 'settled';
      if (!tc._resolveProposal) return tc._enriching ? 'preparing' : 'expired';
      return 'pending';
    }
    const d = tc.proposal ? tc.proposal.decision : null;
    if (d === 'applied' || d === 'declined' || d === 'error' || d === 'refresh_requested' || d === 'stopped') return 'settled';
    if (d === 'interrupted') return 'interrupted';
    if (d === 'applying') return tc._resolveProposal ? 'applying' : 'interrupted';
    if (d === 'stale') return tc._resolveProposal ? 'stale' : 'expired';
    if (!tc._resolveProposal) return tc._enriching ? 'preparing' : 'expired';
    return 'pending';
  }

  function normalizeProposalArgs(args) {
    const normalized = args && typeof args === 'object' ? { ...args } : {};
    const edits = Array.isArray(normalized.edits) ? normalized.edits : [];
    // 兼容旧工具定义诱导出的 edits: [{ changes: [...] }]，统一提升为顶层 changes。
    if ((!Array.isArray(normalized.changes) || !normalized.changes.length) && edits.length === 1 && Array.isArray(edits[0].changes)) {
      normalized.changes = edits[0].changes;
      normalized.explanation = normalized.explanation || edits[0].explanation || edits[0].reasoning || '';
      delete normalized.edits;
    }
    return normalized;
  }

  // 提案项校验（v149 逐项点名，v141 提问卡同款标准）：拒绝必须点名「哪一项缺什么字段」
  // 并给语义与出路——笼统的「提案内容不完整」让模型盲重试（2026-09-04 真机：AI 漏发
  // replacement 被笼统打回，多绕一弯才对）。返回空串 = 通过。
  // 【49-R4】审核卡错误的结构化诊断：目标/参数/原 Office 错误/阶段/残形 ID/写入状态。
  // 审核卡路径不经过 api.js 的异常字段拷贝，必须在此处保留——界面可摘要显示，
  // 字段完整性供模型自纠与复核使用。
  function diagnosticsOf(e) {
    if (!e || typeof e !== 'object') return {};
    const out = {};
    for (const key of ['writeState', 'shapeCreated', 'shapeId', 'phase', 'target', 'shape', 'office']) {
      if (e[key] !== undefined) out[key] = e[key];
    }
    return out;
  }

  // 卡上编辑替换文字仅对 Word 文字提案开放：Excel/PPT 的修改值在 target.cells/operation 里，
  // 编辑 replacement 文字不影响实际写入，只会误导（2026-09-05 真机：Excel 编辑 999 应用仍是旧值）
  function canEditReplacement() { return !!(App.host && App.host.hostType === 'word'); }

  function proposalItemError(item, index, listKey) {
    const at = `${listKey}[${index}]${item && item.label ? `（${String(item.label).slice(0, 40)}）` : ''}`;
    if (!item || typeof item !== 'object') return `${at} 不是有效对象。每项必须是 { label, find, replacement } 形态的对象。`;
    if (item.kind === 'layout' || item.operation) {
      if (App.host && typeof App.host.validateEditProposal === 'function') {
        const verdict = App.host.validateEditProposal(item);
        if (verdict !== true) return `${at} ${typeof verdict === 'string' ? verdict : t('proposalInvalid')}`;
      }
      const operation = item.operation;
      const target = item.target || {};
      const validType = operation && ['addShape', 'updateShape', 'deleteShape'].includes(operation.type);
      if (!validType || (!target.slideId && (target.index == null || !Number.isInteger(Number(target.index))))) return `${at} 缺少有效的 operation.type（addShape/updateShape/deleteShape）或目标位置（target.slideId/index）。`;
      if (operation.type === 'addShape') {
        const shape = operation.shape || {};
        if (!['rectangle', 'roundRectangle', 'ellipse', 'textBox', 'line'].includes(shape.shapeType)) return `${at} 的 shapeType 不受支持（收到「${shape.shapeType || ''}」）。可用值：rectangle/roundRectangle/ellipse/textBox/line。`;
        if (!['left', 'top', 'width', 'height'].every(key => Number.isFinite(Number(shape[key])))) return `${at} 的 shape 缺少数值化的 left/top/width/height。`;
      } else if (!target.shapeId || !operation.expected || typeof operation.expected !== 'object' || (operation.type === 'updateShape' && !Object.keys(operation.shape || {}).length)) return `${at} 缺少 target.shapeId 或 operation.expected（目标当前状态快照）。`;
      return '';
    }
    const find = String(item.find == null ? '' : item.find);
    const hasReplacement = Object.prototype.hasOwnProperty.call(item, 'replacement');
    const replacement = String(item.replacement == null ? '' : item.replacement);
    if (!find.trim()) return `${at} 缺 find（要被修改的原文锚点）。find 必须引用文档中真实存在的文字（先用 get_paragraphs 读取原文再引用）；纯新增内容请改用 insert_heading/insert_text 直通，不走提案卡。`;
    if (!hasReplacement) return `${at} 缺 replacement（修改后的完整文字）。现在只有 find 没有改后内容——请把该项改好的文字放进 replacement 字段后原样重发。`;
    if (App.host && App.host.hostType === 'word' && !item.contextIsWholeParagraph && !String(item.contextBefore || '') && !String(item.contextAfter || '')) return `${at} 缺少可辨认的前后文（contextBefore/contextAfter）。Word 卡的上下文必须由工具核实补齐——请确认 find 与文档原文完全一致后重新提交。`;
    if ((item.placement === 'after' || item.placement === 'before') && !replacement) return `${at} 的 placement 是 ${item.placement}（纯插入）但 replacement 为空。把要新增的文字放进 replacement。`;
    if ((!item.placement || item.placement === 'replace') && find === replacement) return `${at} 的 find 与 replacement 相同（没有任何变更）。请修改 replacement 使其与原文不同。`;
    return '';
  }

  function isLayoutProposal(item) { return Boolean(item && (item.kind === 'layout' || item.operation)); }

  function layoutValue(value) {
    const number = Number(value);
    return Number.isFinite(number) ? String(Math.round(number * 10) / 10) : String(value == null ? '' : value);
  }

  function layoutOperationHtml(change) {
    const operation = change.operation || {};
    const shape = operation.shape || {};
    const expected = operation.expected || {};
    const opKey = operation.type === 'addShape' ? 'layoutAddShape' : operation.type === 'deleteShape' ? 'layoutDeleteShape' : 'layoutUpdateShape';
    const typeKey = `layoutShape_${shape.shapeType || expected.type || 'shape'}`;
    const rows = [];
    const row = (label, value, before) => {
      if (value == null || value === '') return;
      const changeHtml = before != null && String(before) !== String(value)
        ? `<span class="layout-before">${escapeHtml(layoutValue(before))}</span><span class="layout-arrow">${icon('arrowRight','sm')}</span>`
        : '';
      rows.push(`<div class="layout-row"><span>${escapeHtml(label)}</span><strong>${changeHtml}${escapeHtml(layoutValue(value))}</strong></div>`);
    };
    if (operation.type === 'deleteShape') {
      row(t('layoutPosition'), `L ${layoutValue(expected.left)} · T ${layoutValue(expected.top)}`);
      row(t('layoutSize'), `${layoutValue(expected.width)} × ${layoutValue(expected.height)} pt`);
      if (expected.text) row(t('layoutText'), clampLayoutText(expected.text));
    } else {
      const position = (hasLayoutValue(shape, 'left') || hasLayoutValue(shape, 'top')) ? `L ${layoutValue(shape.left)} · T ${layoutValue(shape.top)}` : '';
      const oldPosition = (hasLayoutValue(expected, 'left') || hasLayoutValue(expected, 'top')) ? `L ${layoutValue(expected.left)} · T ${layoutValue(expected.top)}` : null;
      row(t('layoutPosition'), position, oldPosition);
      const size = (hasLayoutValue(shape, 'width') || hasLayoutValue(shape, 'height')) ? `${layoutValue(shape.width)} × ${layoutValue(shape.height)} pt` : '';
      const oldSize = (hasLayoutValue(expected, 'width') || hasLayoutValue(expected, 'height')) ? `${layoutValue(expected.width)} × ${layoutValue(expected.height)} pt` : null;
      row(t('layoutSize'), size, oldSize);
      if (shape.fillColor) rows.push(`<div class="layout-row"><span>${t('layoutFill')}</span><strong><i class="layout-swatch" style="--swatch:${escapeHtml(shape.fillColor)}"></i>${escapeHtml(shape.fillColor)}${shape.fillTransparency != null ? ` · ${Math.round(Number(shape.fillTransparency) * 100)}%` : ''}</strong></div>`);
      if (shape.lineVisible === false) row(t('layoutBorder'), t('layoutNone'));
      else if (shape.lineColor || shape.lineWidth) row(t('layoutBorder'), [shape.lineColor, shape.lineWidth ? `${layoutValue(shape.lineWidth)} pt` : ''].filter(Boolean).join(' · '));
      if (hasLayoutValue(shape, 'text')) row(t('layoutText'), clampLayoutText(shape.text), hasLayoutValue(expected, 'text') ? clampLayoutText(expected.text) : null);
      if (shape.fontSize || shape.fontColor || shape.bold != null) row(t('layoutTypography'), [shape.fontName, shape.fontSize ? `${layoutValue(shape.fontSize)} pt` : '', shape.fontColor, shape.bold ? t('layoutBold') : ''].filter(Boolean).join(' · '));
      if (shape.zOrder) row(t('layoutLayer'), t(`layoutZ_${shape.zOrder}`));
    }
    return `<div class="layout-preview">
      <div class="layout-op"><span class="layout-op-icon">${operation.type === 'addShape' ? icon('plus','sm') : operation.type === 'deleteShape' ? icon('minus','sm') : icon('swap','sm')}</span><strong>${escapeHtml(t(opKey))}</strong><span>${escapeHtml(t(typeKey) === typeKey ? (shape.shapeType || expected.type || t('layoutShape_shape')) : t(typeKey))}</span></div>
      ${rows.length ? `<div class="layout-rows">${rows.join('')}</div>` : ''}
    </div>`;
  }

  function hasLayoutValue(object, key) { return Object.prototype.hasOwnProperty.call(object || {}, key); }

  // replace 型提案的删除量警示（真机 2026-09-04：「末尾追加」被模型做成替换整段，原文被删）——
  // 只提示不拦截：大段重写是合法操作，但用户必须在确认前看到删除规模
  function removedCharCount(change) {
    const original = String(change.find || '');
    const replacement = String(change.replacement == null ? '' : change.replacement);
    let prefix = 0;
    const maxPrefix = Math.min(original.length, replacement.length);
    while (prefix < maxPrefix && original[prefix] === replacement[prefix]) prefix++;
    let suffix = 0;
    const maxSuffix = Math.min(original.length - prefix, replacement.length - prefix);
    while (suffix < maxSuffix && original[original.length - 1 - suffix] === replacement[replacement.length - 1 - suffix]) suffix++;
    return original.length - prefix - suffix;
  }

  function deletionWarningHtml(change) {
    if (isLayoutProposal(change)) return '';
    if ((change.placement || 'replace') !== 'replace') return '';
    const removed = removedCharCount(change);
    return removed > 30 ? `<div class="diff-warning">${icon('alert','sm')} ${escapeHtml(t('deleteWarning').replace('{n}', String(removed)))}</div>` : '';
  }
  function clampLayoutText(value) {
    const text = String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
    return text.length > 80 ? text.slice(0, 80) + '…' : text;
  }

  function inlineDiffHtml(change) {
    if (isLayoutProposal(change)) return layoutOperationHtml(change);
    const before = String(change.contextBefore || '');
    const after = String(change.contextAfter || '');
    const original = String(change.find || '');
    const replacement = String(change.replacement == null ? '' : change.replacement);
    const placement = change.placement || 'replace';
    const span = (cls, value) => value ? `<span class="${cls}">${escapeHtml(value)}</span>` : '';

    if (before || after) {
      if (placement === 'after') return span('diff-context', before + original) + span('diff-add', replacement) + span('diff-context', after);
      if (placement === 'before') return span('diff-context', before) + span('diff-add', replacement) + span('diff-context', original + after);
      return span('diff-context', before) + span('diff-remove', original) + span('diff-add', replacement) + span('diff-context', after);
    }
    if (placement === 'after') return span('diff-context', original) + span('diff-add', replacement);
    if (placement === 'before') return span('diff-add', replacement) + span('diff-context', original);

    let prefix = 0;
    const maxPrefix = Math.min(original.length, replacement.length);
    while (prefix < maxPrefix && original[prefix] === replacement[prefix]) prefix++;
    let suffix = 0;
    const maxSuffix = Math.min(original.length - prefix, replacement.length - prefix);
    while (suffix < maxSuffix && original[original.length - 1 - suffix] === replacement[replacement.length - 1 - suffix]) suffix++;
    const oldEnd = suffix ? original.length - suffix : original.length;
    const newEnd = suffix ? replacement.length - suffix : replacement.length;
    const oldMid = original.slice(prefix, oldEnd);
    const newMid = replacement.slice(prefix, newEnd);
    const sharedBefore = original.slice(0, prefix);
    const sharedAfter = suffix ? original.slice(original.length - suffix) : '';
    return span('diff-context', sharedBefore) + span('diff-remove', oldMid) + span('diff-add', newMid) + span('diff-context', sharedAfter);
  }

  function settleBatchIfDone(tc) {
    const b = tc.batch;
    if (!b || b.settled) return;
    if (b.items.some(s => s === 'pending' || s === 'applying' || s === 'stale')) return;
    b.settled = true;
    const changes = (tc.args && tc.args.changes) || [];
    // 【47.5/48-R4/49-R4】失败项的错误原因、写入状态与结构化诊断（目标/参数/原 Office
    // 错误/阶段/残形 ID）随逐项结果回喂模型——审核卡路径不经过 api.js 异常拷贝，这里
    // 是诊断到达工具结果的唯一通道。措辞按写入状态区分：只有确认未写入（not_committed）
    // 才能说「未写入」；其余一律「结果待核验」
    const results = changes.map((c, i) => {
      const base = { index: i, label: c.label, applied: b.items[i] === 'applied', status: b.items[i] };
      if (b.skipReasons && b.skipReasons[i]) base.reason = b.skipReasons[i];
      const itemErr = b.itemErrors ? b.itemErrors[i] : null;
      if (itemErr) {
        const { message, ...diagnostics } = itemErr;
        base.error = String(message || '');
        Object.assign(base, diagnostics);
      }
      return base;
    });
    const errorCount = b.items.filter(s => s === 'error').length;
    const refreshRequested = changes.map((c, i) => b.items[i] === 'refresh_requested' ? { index: i, label: c.label, target: c.target || null, previousFind: c.find || '', error: b.itemErrors && b.itemErrors[i] } : null).filter(Boolean);
    tc.result = { success: errorCount === 0, applied: b.appliedCount, skipped: b.items.filter(s => s === 'skipped').length, errors: errorCount, total: changes.length, declined: b.appliedCount === 0 && errorCount === 0 && !refreshRequested.length, refreshRequested, results };
    // 【47.5/48-R4】部分失败≠完成：observation 直接进工具结果，措辞与各项写入状态一致
    if (errorCount > 0) {
      const failedItems = results.filter(r => r.status === 'error');
      const confirmedNotWritten = failedItems.filter(r => r.writeState === 'not_committed');
      const pendingVerify = failedItems.filter(r => r.writeState !== 'not_committed');
      const labelOf = r => `#${r.index + 1} ${r.label || ''}`.trim();
      const parts = [];
      if (confirmedNotWritten.length) parts.push(`${confirmedNotWritten.map(labelOf).join('、')}已确认未写入`);
      if (pendingVerify.length) parts.push(`${pendingVerify.map(labelOf).join('、')}结果待核验（是否产生部分写入或残留，以各项错误说明为准，需 get_slide 核对后处置）`);
      tc.result.observation = `本批有 ${errorCount}/${changes.length} 项未成功完成：${parts.join('；')}。不要宣称整页或整份任务已完成：向用户逐项列出失败项、写入状态与原因，并在同一轮给出下一步（重新读取目标后出修正提案，或说明该项需要手动处理）。`;
    }
    // A4（v146）：整卡拒绝附结构化理由，observation 回喂模型（只填附言不点 chip 也带上）
    if (tc.result.declined && (b.declineReason || (b.declineNote || '').trim())) {
      const reason = { selected: b.declineReason ? t(b.declineReason) : '', note: (b.declineNote || '').trim() };
      tc.result.reason = reason;
      tc.result.observation = declineObservation(reason);
    }
    delete b.pendingReplacements;
    b.editing = -1;
    b.declineOpen = false;
    const r = tc._resolveProposal;
    if (r) { tc._resolveProposal = null; r(tc.result); }
  }

  async function presentEditProposal(uiCall, args) {
    let normalized = normalizeProposalArgs(args);
    if (App.host && typeof App.host.enrichEditProposal === 'function') {
      // 【34.3】enrich 补读期间卡还没有解析器：标记准备中，渲染成「准备中」禁用态，
      // 既不画可点假按钮，也不误标永久过期
      uiCall._enriching = true;
      try { normalized = await App.host.enrichEditProposal(normalized); }
      finally { uiCall._enriching = false; }
    } else {
      uiCall._enriching = false;   // 【35-R1】无 enrich 的宿主：api.js 的前置准备标记在此收尾
    }
    uiCall.args = normalized;
    const changes = Array.isArray(normalized.changes) ? normalized.changes : [];
    const edits = Array.isArray(normalized.edits) ? normalized.edits : [];
    const items = changes.length ? changes : edits;
    const listKey = changes.length ? 'changes' : 'edits';
    const itemError = !items.length
      ? `propose_edits 的 ${listKey} 数组为空：至少给一项 { label, find, replacement }。没有真实修改就不要调用本工具（纯新增用 insert_heading/insert_text，汇报结论用普通回复）。`
      : items.map((item, i) => proposalItemError(item, i, listKey)).find(Boolean);
    if (itemError) {
      uiCall.proposalError = itemError;
      // 【48-R2】出卡校验失败也要带结构化可重试性：能力类确定错误（_snapshotError.retryable
      // =false）不得与「不可原样重试」的文字矛盾地标 retryable:true；error 字段兼容旧调用
      const badItem = items.find(item => item && item.operation && item.operation._snapshotError);
      const snap = badItem ? badItem.operation._snapshotError : null;
      const result = { success: false, error: itemError, retryable: snap ? snap.retryable !== false : true };
      if (snap) {
        result.object = snap.object;
        result.phase = snap.phase;
        result.reason = snap.reason;
        result.guidance = snap.retryable === false
          ? '此项为确定不可行（如组合内容本宿主不可读写）：不要按原参数原样重试，改为对可操作对象出提案，或向用户说明需要手动处理。'
          : '请先用 get_slide 重新读取目标现状，再重出提案；不要按原参数原样重试。';
      }
      return Promise.resolve(result);
    }
    return new Promise(resolve => {
      uiCall.proposal = { decision: null, appliedIdx: null, activeIdx: 0, reasoningOpen: false };
      if (changes.length) uiCall.batch = { items: changes.map(() => 'pending'), settled: false, appliedCount: 0, itemErrors: {}, itemVerifs: {} };
      uiCall._resolveProposal = resolve;
      // 【34.3】解析器就位即重绘：可操作性以解析器为准后，按钮必须在这一刻才真正出现
      //（loop 在等用户决策期间不会再 render，缺了这次重绘卡会一直停在准备中）
      render();
      // 【38】出卡检查点：卡可见、进入等待用户决策前就把提案快照落盘
      //（真机 D：关窗重开后请求还在、待决卡消失——上一次保存发生在出卡前）。
      persistProposalCheckpoint(uiCall);
    });
  }

  // 【38】按卡的对象归属持久化：enrich 是异步的，完成时用户可能已切走当前会话，
  // 必须找到实际持有这张卡的会话只写它，不能误写当前另一个会话。
  // resolver 函数不持久化（JSON 存不下，恢复后按 35-R1 归一化为只读过期/中断）。
  // 持久化失败沿用 saveSessions 的容量警示（storageQuotaWarning），不宣称已保留。
  function persistProposalCheckpoint(uiCall) {
    const holds = msgs => Array.isArray(msgs) && msgs.some(m => m && Array.isArray(m.toolCalls) && m.toolCalls.includes(uiCall));
    if (holds(state.messages)) { persistCurrentSession(); return; }
    const owner = (state.sessions || []).find(s => holds(s.messages));
    if (owner) saveSessions();
  }

  function batchCardHtml(tc) {
    if (!tc.rid) tc.rid = App.id();
    const changes = (tc.args && tc.args.changes) || [];
    const b = tc.batch || (tc.batch = { items: changes.map(() => 'pending'), settled: false, appliedCount: 0, itemErrors: {} });
    const pending = b.items.filter(s => s === 'pending' || s === 'stale').length;
    const applicable = b.items.filter(s => s === 'pending').length;
    // 【34.3】阶段判断与事件分派共用 proposalCardPhase：插件重开/会话恢复后解析器丢失
    // 的卡渲染成「已过期」只读态，是否失效只看卡自身，不再被全局 isWorking 赦免；
    // enrich 准备期显示「准备中」而非假按钮或误标过期
    const phase = proposalCardPhase(tc);
    const expired = phase === 'expired';
    const preparing = phase === 'preparing';
    // 批量应用进行中同样锁定操作区，逐项徽标（执行中/已应用）实时可见
    const done = b.settled || expired || preparing || b.applyingAll === true;
    const tcRid = tc.rid;
    const items = changes.map((c, i) => {
      const st = b.items[i];
      const badge = (st === 'applied' ? `<span class="diff-badge ok">${icon('check','sm')} ${t('proposalApplied')}</span>`
        : st === 'skipped' ? `<span class="diff-badge no">${icon('close','sm')} ${t('proposalDeclined')}</span>`
        : st === 'applying' ? `<span class="diff-badge pending">${t('statusRunning')}</span>`
        : st === 'interrupted' ? `<span class="diff-badge err">${t('proposalInterruptedBadge')}</span>`
        : st === 'error' ? `<span class="diff-badge err">${t('proposalApplyError')}</span>`
        : st === 'stale' ? `<span class="diff-badge stale">${t('proposalStale')}</span>`
        : st === 'refresh_requested' ? `<span class="diff-badge pending">${t('proposalRefreshQueued')}</span>`
        : '')
        + (b.edited && b.edited[i] ? ` <span class="diff-badge edit">${icon('edit','sm')} ${t('diffEditedBadge')}</span>` : '');
      const actions = done ? '' : st === 'pending' ? `<span class="bi-actions">
          <button class="bi-yes" data-action="apply-batch-item" data-rid="${tcRid}" data-idx="${i}" title="${t('proposalApply')}">${icon('check','sm')}</button>
          ${!isLayoutProposal(c) && canEditReplacement() ? `<button class="bi-edit" data-action="toggle-batch-edit" data-rid="${tcRid}" data-idx="${i}" title="${t('diffEditLabel')}">${icon('edit','sm')}</button>` : ''}
          <button class="bi-no" data-action="skip-batch-item" data-rid="${tcRid}" data-idx="${i}" title="${t('proposalDecline')}">${icon('close','sm')}</button>
        </span>` : st === 'stale' ? `<span class="stale-actions"><button data-action="refresh-batch-item" data-rid="${tcRid}" data-idx="${i}">${t('proposalRefresh')}</button><button data-action="skip-batch-item" data-rid="${tcRid}" data-idx="${i}">${t('proposalSkip')}</button></span>` : st === 'error' ? `<span class="stale-actions"><button data-action="retry-batch-item" data-rid="${tcRid}" data-idx="${i}">${t('proposalRetry')}</button></span>` : '';
      const rationaleSummary = String(c.summary || tc.args.explanation || '').trim();
      const rationaleReason = String(c.reasoning || '').trim();
      const locatable = Boolean(proposalCitation(c));
      return `<div class="batch-item st-${st} ${locatable ? 'is-locatable' : ''}" ${locatable ? `data-action="navigate-proposal-item" data-rid="${tcRid}" data-idx="${i}" role="button" tabindex="0" title="${escapeHtml(t('proposalLocate'))}"` : ''}>
        <div class="bi-head">
          <span class="bi-label">${escapeHtml(c.label || `#${i + 1}`)}</span>
          ${locatable ? `<span class="bi-locate" aria-hidden="true">${icon('locate')}</span>` : ''}
          ${badge}
          ${actions}
        </div>
        <div class="inline-diff ${isLayoutProposal(c) ? 'layout-diff' : ''}">${inlineDiffHtml(c)}</div>
        ${deletionWarningHtml(c)}
        ${b.editing === i && phase === 'pending' && canEditReplacement() && !isLayoutProposal(c) ? diffEditHtml(tcRid, i, b.pendingReplacements && b.pendingReplacements[i] != null ? b.pendingReplacements[i] : String(c.replacement == null ? '' : c.replacement)) : ''}
        ${st === 'interrupted' ? `<div class="stale-note">${t('proposalApplyInterruptedHint')}</div>` : (st === 'stale' || st === 'error') ? `<div class="stale-note">${escapeHtml((b.itemErrors && b.itemErrors[i] && b.itemErrors[i].message) || t('proposalStaleHint'))}</div>` : ''}
        ${verificationNoteHtml(b.itemVerifs && b.itemVerifs[i])}
        ${(rationaleSummary || rationaleReason) ? `<div class="diff-rationale"><span>${t('proposalReasoning')}</span>${rationaleSummary ? `<strong>${escapeHtml(rationaleSummary)}</strong>` : ''}${rationaleReason && rationaleReason !== rationaleSummary ? `<p>${escapeHtml(rationaleReason)}</p>` : ''}</div>` : ''}
      </div>`;
    }).join('');
    const footActions = done ? (expired ? `<div class="diff-actions"><span class="stale-note">${t('proposalExpiredHint')}</span></div>` : preparing ? `<div class="diff-actions"><span class="stale-note">${t('proposalPreparing')}</span></div>` : '') : `<div class="diff-actions">
      ${applicable ? `<button class="ghost-btn primary-ghost" data-action="apply-batch" data-rid="${tcRid}">${icon('check','sm')} ${t('batchApply')}（${applicable}）</button>` : ''}
      <button class="ghost-btn" data-action="decline-batch" data-rid="${tcRid}">${t('batchSkipAll')}</button>
    </div>`;
    return `<div class="diff-card batch" data-rid="${tcRid}">
      <div class="diff-head"><span class="dh-ico">${icon('edit','sm')}</span><span>${t('batchSuggestions').replace('{count}', changes.length)}</span>${expired ? `<span class="diff-badge stale">${t('proposalExpired')}</span>` : preparing ? `<span class="diff-badge pending">${t('proposalPreparing')}</span>` : done ? `<span class="diff-badge ok">${t('batchReviewed')} ${changes.length}/${changes.length}</span>` : `<span class="diff-badge pending">${pending} ${t('batchPending')}</span>`}</div>
      ${tc.args.explanation ? `<div class="batch-intro">${escapeHtml(tc.args.explanation)}</div>` : ''}
      ${items}
      ${footActions}
      ${phase === 'pending' ? declineReasonHtml(tc, 'batch') : ''}
    </div>`;
  }

  // 写后公式验证回执（Excel 批 1）：应用成功但公式算出 #REF!/#DIV/0! 等错误时，在卡上附一行
  // 地址+错误码；验证被跳过（超大写入）或无错误时不渲染，保持回执安静。
  function verificationNoteHtml(v) {
    if (!v || !v.checked || !v.formulaErrorCount) return '';
    const list = (v.formulaErrors || []).slice(0, 5).map(e => `${e.address}: ${e.error}`).join('；');
    const more = v.formulaErrorCount > 5 ? ` +${v.formulaErrorCount - 5}` : '';
    return `<div class="verif-note">${escapeHtml(t('proposalFormulaErrors'))} ${escapeHtml(list + more)}</div>`;
  }

  function editProposalCardHtml(tc) {
    if (!tc.rid) tc.rid = App.id();
    const p = tc.proposal || (tc.proposal = { decision: null, appliedIdx: null, activeIdx: 0, reasoningOpen: false });
    const edits = (tc.args && Array.isArray(tc.args.edits)) ? tc.args.edits : [];
    if (!edits.length) return `<div class="diff-card"><div class="diff-head">${icon('edit','sm')} ${t('proposalTitle')}</div><div class="diff-empty">${escapeHtml(t('statusError'))}</div></div>`;
    // 【34.3】阶段判断与事件分派共用 proposalCardPhase：解析器丢失即过期（不看全局
    // isWorking），applying 显示处理中并禁用，enrich 准备期显示准备中
    const phase = proposalCardPhase(tc);
    const decided = phase === 'settled';
    const expired = phase === 'expired';
    // 决策后锁定到被应用的版本；拒绝则停留在当前查看的版本
    const viewIdx = p.decision === 'applied' && p.appliedIdx != null ? p.appliedIdx : (p.activeIdx || 0);
    const view = edits[Math.min(viewIdx, edits.length - 1)];
    const badge = (expired ? `<span class="diff-badge stale">${t('proposalExpired')}</span>`
      : phase === 'interrupted' ? `<span class="diff-badge err">${t('proposalInterruptedBadge')}</span>`
      : phase === 'applying' ? `<span class="diff-badge pending">${t('proposalApplying')}</span>`
      : phase === 'preparing' ? `<span class="diff-badge pending">${t('proposalPreparing')}</span>`
      : p.decision === 'applied' ? `<span class="diff-badge ok">${icon('check','sm')} ${t('proposalApplied')}</span>`
      : p.decision === 'declined' ? `<span class="diff-badge no">${t('proposalDeclined')}</span>`
      : p.decision === 'error' ? `<span class="diff-badge err">${t('proposalApplyError')}</span>`
      : p.decision === 'stale' ? `<span class="diff-badge stale">${t('proposalStale')}</span>`
      : p.decision === 'refresh_requested' ? `<span class="diff-badge pending">${t('proposalRefreshQueued')}</span>`
      : `<span class="diff-badge pending">${t('proposalPending')}</span>`)
      + (p.edited ? ` <span class="diff-badge edit">${icon('edit','sm')} ${t('diffEditedBadge')}</span>` : '');
    const tabsLocked = phase !== 'pending' && phase !== 'stale';
    const tabs = edits.length > 1 ? `<div class="diff-tabs">${edits.map((e, i) => `<button class="diff-tab ${i === viewIdx ? 'active' : ''}" data-action="pick-diff-version" data-rid="${tc.rid}" data-idx="${i}" ${tabsLocked ? 'disabled' : ''}>${escapeHtml(e.label || `版本 ${i + 1}`)}</button>`).join('')}</div>` : '';
    const reason = view.reasoning || view.summary || tc.args.explanation || '';
    const cardState = expired ? 'stale' : phase === 'interrupted' ? 'error' : phase === 'applying' ? 'applying' : p.decision === 'applied' ? 'applied' : p.decision === 'declined' ? 'skipped' : p.decision === 'error' ? 'error' : p.decision === 'stale' ? 'stale' : 'pending';
    const locatable = Boolean(proposalCitation(view));
    const actions = decided ? '' : expired ? `<div class="diff-actions"><span class="stale-note">${t('proposalExpiredHint')}</span></div>` : phase === 'interrupted' ? `<div class="diff-actions"><span class="stale-note">${t('proposalApplyInterruptedHint')}</span></div>` : phase === 'applying' ? `<div class="diff-actions"><button class="ghost-btn primary-ghost" disabled>${icon('check','sm')} ${t('proposalApplying')}</button></div>` : phase === 'preparing' ? `<div class="diff-actions"><span class="stale-note">${t('proposalPreparing')}</span></div>` : p.decision === 'stale' ? `<div class="diff-actions stale-footer">
        <button class="ghost-btn primary-ghost" data-action="refresh-diff" data-rid="${tc.rid}">${t('proposalRefresh')}</button>
        <button class="ghost-btn" data-action="skip-stale-diff" data-rid="${tc.rid}">${t('proposalSkip')}</button>
      </div>` : `<div class="diff-actions">
        <button class="ghost-btn primary-ghost" data-action="apply-diff" data-rid="${tc.rid}">${icon('check','sm')} ${t('proposalApply')}</button>
        ${canEditReplacement() && !isLayoutProposal(view) ? `<button class="ghost-btn" data-action="toggle-diff-edit" data-rid="${tc.rid}">${icon('edit','sm')} ${t('diffEditLabel')}</button>` : ''}
        <button class="ghost-btn" data-action="decline-diff" data-rid="${tc.rid}">${t('proposalDecline')}</button>
      </div>`;
    return `<div class="diff-card single">
      <div class="diff-head"><span class="dh-ico">${icon('edit','sm')}</span><span>${t('batchSuggestions').replace('{count}', '1')}</span>${badge}</div>
      ${tc.args.explanation ? `<div class="batch-intro">${escapeHtml(tc.args.explanation)}</div>` : ''}
      ${tabs}
      <div class="batch-item single-item st-${cardState} ${locatable ? 'is-locatable' : ''}" ${locatable ? `data-action="navigate-proposal-item" data-rid="${tc.rid}" data-idx="${viewIdx}" role="button" tabindex="0" title="${escapeHtml(t('proposalLocate'))}"` : ''}>
        <div class="bi-head"><span class="bi-label">${escapeHtml(view.label || t('proposalTitle'))}</span>${locatable ? `<span class="bi-locate" aria-hidden="true">${icon('locate')}</span>` : ''}</div>
        <div class="inline-diff ${isLayoutProposal(view) ? 'layout-diff' : ''}">${inlineDiffHtml(view)}</div>
        ${deletionWarningHtml(view)}
        ${p.editing && phase === 'pending' && canEditReplacement() && !isLayoutProposal(view) ? diffEditHtml(tc.rid, viewIdx, p.pendingReplacement != null ? p.pendingReplacement : String(view.replacement == null ? '' : view.replacement)) : ''}
        ${(p.decision === 'stale' || p.decision === 'error') ? `<div class="stale-note">${escapeHtml((p.error && p.error.message) || t('proposalStaleHint'))}</div>` : ''}
        ${verificationNoteHtml(tc.result && tc.result.verification)}
        ${reason ? `<div class="diff-rationale"><span>${t('proposalReasoning')}</span>${(view.summary || tc.args.explanation) ? `<strong>${escapeHtml(view.summary || tc.args.explanation)}</strong>` : ''}${view.reasoning && view.reasoning !== view.summary ? `<p>${escapeHtml(view.reasoning)}</p>` : ''}</div>` : ''}
      </div>
      ${actions}
      ${phase === 'pending' ? declineReasonHtml(tc, 'single') : ''}
    </div>`;
  }

  // 提案被出卡前拦截的错误卡（proposalError 与 vet/空壳/enrich 拦截共用）：原因全文上屏
  function proposalErrorCardHtml(reason) {
    return `<div class="diff-card"><div class="diff-head"><span class="dh-ico">${icon('edit','sm')}</span><span>${t('proposalTitle')}</span><span class="diff-badge err">${t('statusError')}</span></div><div class="diff-empty">${escapeHtml(reason)}</div></div>`;
  }

  function toolCardHtml(tc) {
    // render 卡片的全部实现在 render-blocks.js，这里只做一次转发。
    // 不要把渲染逻辑挪进本文件：ui.js 历史上被批量重构误删过函数，新增功能一律隔离在独立文件里。
    if (tc.name === 'render' && App.renderBlockHtml) {
      const html = App.renderBlockHtml(tc);
      const flash = tc.blockFlash ? `<div class="blk-hint">${escapeHtml(tc.blockFlash)}</div>` : '';
      if (html) return html + flash;
    }
    if (tc.name === 'propose_edits' && tc.proposalError) return proposalErrorCardHtml(tc.proposalError);
    // 出卡前被拦截（vet 坏公式 / 空壳提案 / enrich 锚点拒绝）：没有 proposal/batch 也没有解析器，
    // 不能再用原始 args 画卡壳（此前兜底渲染成「已过期」假卡，拦截原因只喂给模型用户看不到）——
    // 拦截原因直接上屏
    if (tc.name === 'propose_edits' && tc.status === 'error' && tc.result && !tc.proposal && !tc.batch) {
      return proposalErrorCardHtml(String(tc.result.error || t('proposalInvalid')));
    }
    // 会话中断遗留的挂起卡（running 但解析器丢失，重开/恢复后出现）：同样是死卡，渲染成提示
    // 而不是拿原始 args 画假卡（v148 漏掉的 status=running 形态，2026-09-05 真机复现）
    if (tc.name === 'propose_edits' && tc.status === 'running' && !tc._resolveProposal && !tc.proposal && !tc.batch && !state.isWorking) {
      return proposalErrorCardHtml(t('proposalInterrupted'));
    }
    if (tc.name === 'propose_edits') return (tc.args && Array.isArray(tc.args.changes) && tc.args.changes.length) ? batchCardHtml(tc) : editProposalCardHtml(tc);
    if (!tc.rid) tc.rid = App.id();
    const status = tc.status || 'running';
    const explain = String((tc.args && tc.args.explanation) || '').trim();
    const expanded = !!tc.expanded;
    const ico = status === 'running' ? '<span class="tico spinner"></span>'
      : status === 'complete' ? `<span class="tico ok">${icon('check','sm')}</span>`
      : status === 'error' ? `<span class="tico err">${icon('close','sm')}</span>`
      : `<span class="tico stop">${icon('stop','sm')}</span>`;
    const detail = tc.expanded ? `<div class="tool-detail"><div class="td-label">args</div><pre>${escapeHtml(pretty(tc.args))}</pre>${tc.result ? `<div class="td-label">result</div><pre>${escapeHtml(trimToolText(tc.result))}</pre>` : ''}</div>` : '';
    // 失败原因直接显示在标题行（卡片默认折叠，连续失败时一排灰卡看不出区别）；
    // 连续失败计数来自 api.js 的 consecutiveToolFailures，>1 时追加「×N」
    let errSummary = '';
    let failCount = 0;
    if (status === 'error' && tc.result && typeof tc.result === 'object') {
      const rawError = tc.result.error;
      errSummary = String(typeof rawError === 'string' ? rawError : (rawError && rawError.message) || '').trim();
      if (errSummary.length > 80) errSummary = errSummary.slice(0, 80) + '…';
      failCount = Number(tc.result.consecutiveFailures) || 0;
    }
    return `<div class="tool-card st-${status} ${expanded ? 'expanded' : ''}" data-action="toggle-tool" data-rid="${tc.rid}">
      <div class="tool-row">${ico}<span class="tool-name">${escapeHtml(toolPhrase(tc))}</span>${errSummary ? `<span class="tool-err" title="${escapeHtml(errSummary)}">${escapeHtml(errSummary)}</span>` : ''}<span class="tool-status-txt">${escapeHtml(statusLabel(status))}${failCount > 1 ? ` ×${failCount}` : ''}</span><span class="tool-chev">${icon('chevronRight')}</span></div>
      ${explain ? `<div class="tool-explain">${escapeHtml(explain)}</div>` : ''}
      ${detail}
    </div>`;
  }

  function assistantBubbleHtml(m) {
    let body = '';
    if (m.reasoning) {
      const open = m.reasoningOpen === true;
      const idx = state.messages.indexOf(m);
      body += `<div class="reasoning-block ${open ? 'open' : ''}">
        <button type="button" class="reasoning-head" data-action="toggle-reasoning" data-idx="${idx}" aria-expanded="${open}"><span class="reasoning-chev">${open ? icon('chevronDown') : icon('chevronRight')}</span><span>${t('reasoningLabel')}</span></button>
        ${open ? `<div class="reasoning-body">${escapeHtml(m.reasoning)}</div>` : ''}
      </div>`;
    }
    if (m.content) body += `<div class="markdown">${renderMarkdown(m.content)}</div>`;
    if (m.toolCalls && m.toolCalls.length) {
      body += toolCallsHtml(m.toolCalls);
    }
    return body || `<div class="typing"><span class="phase-dots"><i></i><i></i><i></i></span></div>`;
  }

  // 连续同构工具卡折叠组（v135）：渲染走 App.toolGroups 的分组判定，明细复用 toolCardHtml。
  // 折叠只影响呈现：数据不动，会话级汇总仍在顶部「本次会话的修改记录」。
  function toolCallsHtml(list) {
    if (!(App.toolGroups && typeof App.toolGroups.group === 'function')) return list.map(toolCardHtml).join('');
    return App.toolGroups.group(list).map(part => part.type === 'single'
      ? toolCardHtml(part.tc)
      : toolGroupCardHtml(part.items)).join('');
  }

  function toolGroupCardHtml(items) {
    const first = items[0];
    const st = App.toolGroups.status(items);
    const open = App.toolGroups.isOpen(items);
    const label = toolLabel(first.name);
    const ico = st.key === 'running' ? '<span class="tico spinner"></span>'
      : st.key === 'mixed' ? `<span class="tico err">${icon('close','sm')}</span>`
        : `<span class="tico ok">${icon('check','sm')}</span>`;
    const statusText = st.key === 'done' ? t('statusComplete')
      : st.key === 'running' ? t('statusRunning')
        : t('toolGroupMixed').replace('{done}', st.done).replace('{error}', st.error);
    return `<div class="tool-card group st-${st.key} ${open ? 'expanded' : ''}">
      <div class="tool-row" data-action="toggle-tool-group" data-rid="${escapeHtml(first.rid)}">${ico}<span class="tool-name">${escapeHtml(label)} ×${items.length}</span><span class="tool-status-txt">${escapeHtml(statusText)}</span><span class="tool-chev">${icon(open ? 'chevronDown' : 'chevronRight')}</span></div>
      ${open ? `<div class="group-detail">${items.map(toolCardHtml).join('')}</div>` : ''}
    </div>`;
  }

  function canRegenerate() {
    if (state.isWorking || !state.messages.length) return false;
    const hasUser = state.messages.some(m => m.role === 'user');
    const last = state.messages[state.messages.length - 1];
    return hasUser && last && last.role === 'assistant';
  }

  const PROMPT_COLLAPSE_CHARS = 180;
  const PROMPT_COLLAPSE_LINES = 5;

  function userPromptText(m) {
    return String(m && (m.displayContent != null ? m.displayContent : m.content) || '');
  }

  function isCollapsiblePrompt(text) {
    const value = String(text || '');
    return value.length > PROMPT_COLLAPSE_CHARS || value.split(/\r?\n/).length >= PROMPT_COLLAPSE_LINES;
  }

  function promptSummary(text) {
    const compact = String(text || '').replace(/\s+/g, ' ').trim();
    return compact.length > 58 ? compact.slice(0, 58).trimEnd() + '…' : compact;
  }

  function collapseLatestUserPrompt() {
    for (let i = state.messages.length - 1; i >= 0; i--) {
      const m = state.messages[i];
      if (m.role !== 'user') continue;
      if (isCollapsiblePrompt(userPromptText(m))) m.promptCollapsed = true;
      return;
    }
  }

  function renderMessage(m, idx, arr) {
    const time = new Date(m.timestamp || now()).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const brandShort = t('brand') === 'brand' ? 'AI' : t('brand');
    if (m.role === 'tool') {
      return `<div class="msg tool"><div class="bubble">${escapeHtml(m.name || 'tool')}\n${escapeHtml(trimToolText(m.content))}</div><div class="meta">${time}</div></div>`;
    }
    if (m.role === 'assistant') {
      const isLast = idx === arr.length - 1;
      const inPlanMode = (typeof App.currentAccessMode === 'function' ? App.currentAccessMode() : state.settings.accessMode) === 'plan';
      const planActionable = isLast && inPlanMode && !state.isWorking && String(m.content || '').trim();
      const approveBtn = planActionable
        ? `<button class="ghost-btn" data-action="approve-plan" title="${t('accessConfirm')}">${icon('check','sm')} ${t('approvePlan')}</button>` : '';
      const feedbackBtn = planActionable
        ? `<button class="ghost-btn" data-action="feedback-plan" title="${t('feedbackPlan')}">${icon('edit','sm')} ${t('feedbackPlan')}</button>` : '';
      const regenBtn = isLast && canRegenerate() ? `<button class="ghost-btn" data-action="regenerate" title="${t('regenerate')}">${icon('refresh','sm')} ${t('regenerate')}</button>` : '';
      const actions = (approveBtn || feedbackBtn || regenBtn) ? `<div class="msg-actions">${approveBtn}${feedbackBtn}${regenBtn}</div>` : '';
      return `<div class="msg assistant">
        <div class="msg-identity">${App.modelLogoHtml(state.settings.model) || `<img class="avatar" src="${brandLogoSrc()}" alt=""/>`}<span class="who">${escapeHtml(brandShort)}</span><span class="meta">${time}</span></div>
        <div class="msg-body"><div class="assistant-bubble">${assistantBubbleHtml(m)}</div>${actions}</div>
      </div>`;
    }
    const prompt = userPromptText(m);
    const isLong = isCollapsiblePrompt(prompt);
    // 技能标注：消息确实携带 skillId 生效（api.js 按 latestUserMessage.skillId 注入），
    // 之前只在 chip 上可见、消息流里无痕，用户看不出这条消息带了技能——补上徽标
    const skillEntry = allSkillsList().find(s => s.id === m.skillId);
    const skillTag = skillEntry ? `<span class="msg-skill-tag">${icon('sparkle','sm')}${escapeHtml(skillEntry.name || m.skillId)}</span>` : '';
    const rewindBtn = `<button type="button" class="user-rewind" data-action="rewind-to" data-idx="${idx}" title="${t('rewindTo')}" aria-label="${t('rewindTo')}">${icon('refresh')}</button>`;
    if (!isLong) return `<div class="msg user"><div class="bubble">${escapeHtml(prompt)}</div><div class="meta">${rewindBtn}${skillTag}${time}</div></div>`;
    const hasReplyAfter = arr.slice(idx + 1).some(x => x.role === 'assistant');
    const collapsed = m.promptCollapsed === true || (m.promptCollapsed == null && hasReplyAfter);
    return `<div class="msg user prompt-long ${collapsed ? 'prompt-collapsed' : 'prompt-expanded'}">
      <div class="user-prompt-shell">
        <div class="user-prompt-head">
          <button type="button" class="user-prompt-toggle" data-action="toggle-user-prompt" data-idx="${idx}" data-collapsed="${collapsed}" aria-expanded="${!collapsed}" title="${collapsed ? t('promptExpand') : t('promptCollapse')}">
            <span class="blind-glyph" aria-hidden="true">${collapsed ? icon('chevronRight') : icon('chevronDown')}</span>
            <span class="user-prompt-summary">${escapeHtml(promptSummary(prompt))}</span>
            <span class="user-prompt-count">${escapeHtml(t('promptChars').replace('{count}', prompt.length))}</span>
          </button>
          <button type="button" class="user-prompt-copy" data-action="copy-user-prompt" data-idx="${idx}" title="${t('promptCopy')}">${t('promptCopy')}</button>
          <button type="button" class="user-rewind" data-action="rewind-to" data-idx="${idx}" title="${t('rewindTo')}" aria-label="${t('rewindTo')}">${icon('refresh')}</button>
        </div>
        ${collapsed ? '' : `<div class="bubble user-prompt-content">${escapeHtml(prompt)}</div>`}
      </div>
      <div class="meta">${skillTag}${time}</div>
    </div>`;
  }

  function patchStreamingMessage(assistantUi) {
    const msgEl = document.getElementById('messages');
    if (!msgEl) { render(); return; }
    // 必须是 .assistant-bubble：assistant 消息里没有 .bubble 这个类，
    // 选错会导致每次流式更新都退回全量 render（历史 bug）
    const bubbles = msgEl.querySelectorAll('.msg.assistant .assistant-bubble');
    const bubble = bubbles[bubbles.length - 1];
    if (!bubble) { render(); return; }
    const atBottom = msgEl.scrollHeight - msgEl.scrollTop - msgEl.clientHeight < 40;
    const reasoningBody = bubble.querySelector ? bubble.querySelector('.reasoning-body') : null;
    const reasoningScroll = reasoningBody ? {
      top: reasoningBody.scrollTop,
      atBottom: reasoningBody.scrollHeight - reasoningBody.scrollTop - reasoningBody.clientHeight < 18
    } : null;
    bubble.innerHTML = assistantBubbleHtml(assistantUi);
    const nextReasoningBody = bubble.querySelector ? bubble.querySelector('.reasoning-body') : null;
    if (nextReasoningBody) {
      if (!reasoningScroll || reasoningScroll.atBottom) nextReasoningBody.scrollTop = nextReasoningBody.scrollHeight;
      else nextReasoningBody.scrollTop = Math.min(reasoningScroll.top, Math.max(0, nextReasoningBody.scrollHeight - nextReasoningBody.clientHeight));
    }
    const currentPhase = msgEl.querySelector('.phase-bar');
    const nextPhase = phaseBarHtml();
    // 同一阶段保留原节点，避免每个流式片段都让三点动画重新开始。
    if (currentPhase && nextPhase) {
      if (currentPhase.dataset.phase !== state.workPhase) currentPhase.outerHTML = nextPhase;
    } else if (currentPhase) currentPhase.parentNode.removeChild(currentPhase);
    else if (nextPhase) msgEl.insertAdjacentHTML('beforeend', nextPhase);
    if (atBottom) msgEl.scrollTop = msgEl.scrollHeight;
  }

  function fillPrompt(text) {
    const input = document.getElementById('chat-input');
    if (!input) return;
    input.value = text;
    state.composerDraft = text;
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
  }

  // 临时内联提示：引用跳转失败等轻量错误不再弹原生 alert 打断流程，
  // 在消息区底部插一条几秒内自动消失的提示。直接操作 DOM 而不走 render()，
  // 避免整树重绘；若期间发生 render，提示随之消失，也可接受。
  function showTransientNotice(text) {
    const host = document.getElementById('messages');
    if (!host) { state.error = String(text || ''); render(); return; }
    const el = document.createElement('div');
    el.className = 'inline-notice';
    el.textContent = String(text || '');
    host.appendChild(el);
    host.scrollTop = host.scrollHeight;
    setTimeout(() => { el.remove(); }, 4000);
  }

  function stopActiveRequest() {
    if (!state.isWorking) return;
    state.stopRequested = true;
    // 【24.2】通知在途文档写入：本操作所属任务已被停止，未提交的写命令在下一个
    // 提交前检查点取消（粘性，新运行发送时复位）
    if (App.markDocWritesCancelled) App.markDocWritesCancelled();
    if (state.abortController && typeof state.abortController.abort === 'function') state.abortController.abort();
    // 未决的修改提案一并视为拒绝，避免 promise 悬挂
    for (const m of state.messages) {
      if (m.toolCalls) for (const tc of m.toolCalls) {
        if (tc.name === 'propose_edits' && tc._resolveProposal) {
          if (tc.batch) { for (let i = 0; i < tc.batch.items.length; i++) { if (tc.batch.items[i] === 'pending' || tc.batch.items[i] === 'stale') tc.batch.items[i] = 'skipped'; } settleBatchIfDone(tc); continue; }
          settleProposal(tc, { success: false, declined: true, stopped: true });
        }
        // 未决的 render 交互卡同样结算，否则 agent loop 永远 await 悬挂
        if (tc.name === 'render' && tc._resolveBlock && App.abortRenderBlock) App.abortRenderBlock(tc);
      }
    }
    render();
  }

  function clearComposerDraft() {
    state.composerDraft = '';
    state.slashRange = null;
    const input = document.getElementById('chat-input');
    if (input) input.value = '';
  }

  function takeComposerMessage() {
    let text = String(state.composerDraft || '').trim();
    let skillId = state.activeSkillId;
    const command = extractSkillCommand(text);
    if (command) {
      skillId = command.skill.id;
      text = command.text;
      // 斜杠命令显式点名技能：直接切换当前生效技能（持续生效，直到手动 × 或切换会话）
      state.activeSkillId = command.skill.id;
    }
    if (!text) {
      if (command) {
        clearComposerDraft();
        state.slashItems = null;
        state.slashRange = null;
      }
      return null;
    }
    const message = { text, attachments: (state.attachments || []).slice(), skillId: skillId || null };
    clearComposerDraft();
    state.attachments = [];
    // 技能随发送消费（用户 9-02 拍板，反转旧的「持续生效」设计）：发完 chip 即清，
    // 消息自带 skillId 徽标可追溯；连用就再选一次
    state.activeSkillId = null;
    state.slashItems = null;
    state.slashRange = null;
    return message;
  }

  function scheduleComposerMessage(message) {
    if (!message || !message.text) return;
    setTimeout(() => sendUserMessage(message.text, message.attachments, message.skillId), 0);
  }

  // 任务收尾时纠偏队列里可能还有没来得及注入的补充（任务恰好在下一步注入前结束）。
  // 正常结束：合并成一条新消息自动接续发送；停止/出错收尾：保留在队列里，提示用户手动处理。
  function drainSteerQueue(finished) {
    if (!Array.isArray(state.steerQueue) || !state.steerQueue.length) return;
    if (!finished) {
      showTransientNotice(t('steerLeftNotice').replace('{n}', state.steerQueue.length));
      return;
    }
    const queued = state.steerQueue.splice(0);
    const text = queued.map(q => String(q.text || '').trim()).filter(Boolean).join('\n');
    if (!text) return;
    const attachments = queued.flatMap(q => q.attachments || []);
    const withSkill = [...queued].reverse().find(q => q.skillId);
    scheduleComposerMessage({ text, attachments, skillId: withSkill ? withSkill.skillId : null });
  }

  function settleProposal(tc, result) {
    const p = tc.proposal || {};
    p.decision = result.applied ? 'applied' : result.declined ? 'declined' : result.refreshRequested ? 'refresh_requested' : 'error';
    if (typeof result.appliedIdx === 'number') p.appliedIdx = result.appliedIdx;
    // 结算即清编辑/拒绝的暂存态（卡面随后锁定，不留脏状态）
    delete p.pendingReplacement;
    p.editing = false;
    p.declineOpen = false;
    if (tc._resolveProposal) { const r = tc._resolveProposal; tc._resolveProposal = null; r(result); }
  }

  // ── A4 确认卡三层回应（v146）：①应用 ②卡上编辑替换文字后应用 ③拒绝附结构化理由 ──

  // 卡上编辑过的替换文字（input 事件实时暂存，不经 render 保焦点）。没改过返回 null。
  // 按卡型分流：批量卡只读 batch.pendingReplacements[idx]（卡上同时存在 proposal 对象，
  // 混读会让单条目的编辑文字污染其他条目——review P0-1）
  function pendingReplacementOf(tc, idx, item) {
    let pending = null;
    if (tc.batch) pending = tc.batch.pendingReplacements ? tc.batch.pendingReplacements[idx] : null;
    else if (tc.proposal) pending = tc.proposal.pendingReplacement;
    if (typeof pending !== 'string') return null;
    if (pending === String(item && item.replacement == null ? '' : item.replacement)) return null;
    return pending;
  }

  // 应用前把编辑过的替换文字写回提案项（apply-diff / apply-batch-item / apply-batch 共用）
  function takeReplacementOverride(tc, idx, item) {
    const pending = pendingReplacementOf(tc, idx, item);
    if (pending == null) return false;
    item.replacement = pending;
    if (tc.proposal) tc.proposal.edited = true;
    if (tc.batch) { tc.batch.edited = tc.batch.edited || {}; tc.batch.edited[idx] = true; }
    return true;
  }

  // 拒绝理由回喂给模型的 observation（A4：拒绝不是终点——用户给了反馈，模型必须在本轮
  // 分析并推进下一步，不得只沉淀记忆或只确认收到就停下）
  function declineObservation(reason) {
    const parts = [];
    if (reason.selected) parts.push(`原因：${reason.selected}`);
    if (reason.note) parts.push(`用户补充：${reason.note}`);
    return `用户拒绝了这版提案${parts.length ? `（${parts.join('；')}）` : ''}。你的下一个回复必须推进下一步，不能只把反馈写进记忆或只表示收到：替换文字要改/位置不对 → 立即出一张吸收反馈的修正提案卡；方向不对 → 说出你判断的正确方向，请用户确认后执行；用户自己改 → 明确说明你不会再动文档，以及接下来你能帮什么。反观沉淀记忆可以在推进之后顺带做，但推进本身是本回合必须完成的动作。`;
  }

  const DECLINE_REASONS = ['declineWrongDirection', 'declineNeedTextChange', 'declineWrongPlace', 'declineSelfEdit'];

  function declineReasonHtml(tc, kind) {
    const store = kind === 'single' ? tc.proposal : tc.batch;
    if (!store || !store.declineOpen) return '';
    const note = store.declineNote || '';
    return `<div class="decline-reason">
      <div class="dr-title">${escapeHtml(t('declineReasonTitle'))}</div>
      <div class="dr-chips">${DECLINE_REASONS.map(k => `<button type="button" class="dr-chip ${store.declineReason === k ? 'active' : ''}" data-action="pick-decline-reason" data-kind="${kind}" data-rid="${tc.rid}" data-reason="${k}">${escapeHtml(t(k))}</button>`).join('')}</div>
      <input type="text" class="dr-note" data-kind="${kind}" data-rid="${tc.rid}" value="${escapeHtml(note)}" placeholder="${escapeHtml(t('declineReasonPlaceholder'))}" />
      ${store.declineNeedReason ? `<div class="dr-need">${escapeHtml(t('declineReasonNeed'))}</div>` : ''}
      <div class="diff-actions">
        <button class="ghost-btn primary-ghost" data-action="confirm-decline-${kind === 'single' ? 'diff' : 'batch'}" data-rid="${tc.rid}">${escapeHtml(t('declineSubmit'))}</button>
        <button class="ghost-btn" data-action="cancel-decline" data-kind="${kind}" data-rid="${tc.rid}">${escapeHtml(t('declineBack'))}</button>
      </div>
    </div>`;
  }

  function diffEditHtml(rid, idx, value) {
    return `<div class="diff-edit"><textarea class="diff-edit-input" data-rid="${rid}" data-idx="${idx}" rows="3">${escapeHtml(value)}</textarea><div class="diff-edit-hint">${escapeHtml(t('diffEditHint'))}</div></div>`;
  }

  async function runLoop() {
    const myController = typeof AbortController !== 'undefined' ? new AbortController() : { signal: undefined };
    state.abortController = myController;
    state.isWorking = true;
    collapseLatestUserPrompt();
    render();
    let finished = false;
    try {
      await runAgentLoop();
      persistCurrentSession();
      finished = true;
    } catch (e) {
      if (state.stopRequested || e?.name === 'AbortError') {
        markStoppedMessage();
        persistCurrentSession();
      } else {
        state.error = e.message || String(e);
      }
    } finally {
      // 仅当全局 controller 仍是本次请求时才复位，避免停止后立即重发时覆盖新请求的状态
      if (state.abortController === myController) {
        state.isWorking = false;
        state.abortController = null;
      }
      // 任务落盘时 taskActive 跟随 isWorking 归位（persistCurrentSession 里 taskActive = isWorking）
      persistCurrentSession();
      const next = state.pendingInterruptMessage;
      if (state.pendingInterruptMessage) state.pendingInterruptMessage = null;
      state.workPhase = null;
      render();
      // 回复交付完成（finished）：学习闭环的复盘窗口之一（v137）
      if (finished && App.review) { try { App.review.onTurnEnd(); } catch {} }
      if (next) scheduleComposerMessage(next);
      else drainSteerQueue(finished);
    }
  }

  async function sendUserMessage(text, suppliedAttachments, suppliedSkillId) {
    if (state.isWorking) return;
    // 用户发新消息：在跑的学习复盘立即让位（欠账保留，找机会补跑）
    if (App.review) { try { App.review.abort(); } catch {} }
    // 选区开关开着但没选到文字：拦截发送并把草稿/附件放回输入区，
    // 避免「选区被静默丢弃、消息照样发出、草稿还丢了」的三连击
    if (state.selectionTarget && !(state.selectionSnapshot && (String(state.selectionSnapshot.text || '').trim() || String(state.selectionSnapshot.paragraphText || '').trim()))) {
      state.composerDraft = text;
      if (Array.isArray(suppliedAttachments) && suppliedAttachments.length) state.attachments = suppliedAttachments;
      state.error = t('selEmptyBlock');
      render();
      return;
    }
    state.error = null;
    state.stopRequested = false;
    // 【25.2】新运行开始不再复位取消状态：取消代次单调递增且不可复位——
    // 旧任务的取消事实不因新发送消失；新任务的 applyEdit 捕获新代次不受影响
    ensureSession();
    // 选区注入：开启「针对选区」时，把抓取到的选区原文放在消息最前面
    let content = text;
    if (state.selectionTarget && state.selectionSnapshot && String(state.selectionSnapshot.text || '').trim()) {
      content = `[当前选区]\n${state.selectionSnapshot.text}\n[/当前选区]\n\n请优先处理上面选区中的内容。\n\n` + content;
    } else if (state.selectionTarget && state.selectionSnapshot && String(state.selectionSnapshot.paragraphText || '').trim()) {
      // collapsed：光标只点进某段，注入该段全文，模型可直接拿原文做 find 锚定
      content = `[当前光标所在段]\n${state.selectionSnapshot.paragraphText}\n[/当前光标所在段]\n\n请优先处理这一段，定位时优先用这段原文作 find 锚定。\n\n` + content;
    }
    state.selectionTarget = false;
    state.selectionSnapshot = null;
    // 附件注入：作为消息前缀的参考材料
    const attachments = Array.isArray(suppliedAttachments) ? suppliedAttachments : (state.attachments || []);
    if (attachments.length) {
      const blocks = attachments.map(a => `[Reference file: ${a.name}]\n${a.text}\n[/Reference file]`);
      content = blocks.join('\n\n') + '\n\n' + content;
    }
    // 图片不进消息正文（base64 会撑爆 token）：只注入 attachmentId 引用，
    // 模型用 insert_image(attachmentId) 插入，base64 由插件本地取用。
    // 「单次消费」的正确语义（IMG-6 修正）：数据从 UI 挂起区移交到本条消息对象上——
    // chip 消失、不再注入下一条消息，但本轮 agent loop 执行 insert_image 时仍能按 id 取到
    // （IMG-4 直接清空数据导致模型正确调用却报「未找到」）。
    const pendingImages = state.pendingImages || [];
    if (pendingImages.length) {
      const imgBlocks = pendingImages.map(a => `[Uploaded image: ${a.name} | attachmentId: ${a.id}]`);
      content = imgBlocks.join('\n') + '\n' + 'The attachmentId above is valid for THIS message only. Insert it with insert_image only when the user asked to insert/place an image; when adjusting or removing an already-inserted image use manage_image instead — never insert a second copy.' + '\n\n' + content;
    }
    state.attachments = [];
    state.pendingImages = [];
    clearComposerDraft();
    state.messages.push({ role: 'user', content, displayContent: text, skillId: suppliedSkillId || null, promptCollapsed: false, timestamp: now(), pendingImages: pendingImages.length ? pendingImages : undefined });
    persistCurrentSession();
    await runLoop();
  }

  async function regenerate() {
    if (state.isWorking) return;
    while (state.messages.length && state.messages[state.messages.length - 1].role === 'assistant') state.messages.pop();
    if (!state.messages.some(m => m.role === 'user')) return;
    state.error = null;
    state.stopRequested = false;
    persistCurrentSession();
    await runLoop();
  }

  async function runManualTool() {
    const name = document.getElementById('manual-tool').value;
    const raw = document.getElementById('manual-args').value;
    state.toolOutput = 'running...'; render();
    try {
      const args = raw.trim() ? JSON.parse(raw) : {};
      const out = await executeToolByName(name, args);
      state.toolOutput = pretty(out);
      await maybeFollow(out);
    } catch (e) {
      state.toolOutput = 'ERROR: ' + (e.message || String(e));
    }
    render();
  }

  document.addEventListener('click', async (ev) => {
    const clickedInsideSessionMenu = ev.target.closest('.session-menu-wrap');
    const clickedInsideModelMenu = ev.target.closest('.quick-model-wrap') || ev.target.closest('.quick-access-wrap') || ev.target.closest('.quick-thinking-wrap') || ev.target.closest('.plus-wrap');
    const shouldCloseSessionMenu = state.sessionMenuOpen && !clickedInsideSessionMenu;
    const shouldCloseModelMenu = (state.modelMenuOpen || state.accessMenuOpen || state.thinkingMenuOpen || state.plusMenuOpen) && !clickedInsideModelMenu;
    const cite = ev.target.closest('a.citation');
    if (cite) {
      ev.preventDefault();
      const ref = cite.getAttribute('href').replace('#cite:', '');
      await navigateCitation(ref).catch(e => showTransientNotice(e.message || String(e)));
      return;
    }
    const copyBtn = ev.target.closest('button.md-copy');
    if (copyBtn) {
      const codeEl = copyBtn.closest('.md-codeblock')?.querySelector('pre.md-code code');
      const text = codeEl ? codeEl.textContent : '';
      const done = (ok) => {
        copyBtn.classList.toggle('copied', ok); copyBtn.classList.toggle('copy-fail', !ok);
        copyBtn.innerHTML = App.icon(ok ? 'check' : 'close', 'sm');
        setTimeout(() => { copyBtn.innerHTML = App.icon('copy', 'sm'); copyBtn.classList.remove('copied', 'copy-fail'); }, 1200);
      };
      try {
        if (navigator.clipboard?.writeText) {
          await navigator.clipboard.writeText(text);
        } else {
          const ta = document.createElement('textarea');
          ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
          document.body.appendChild(ta); ta.select();
          document.execCommand('copy'); document.body.removeChild(ta);
        }
        done(true);
      } catch { done(false); }
      return;
    }
    const tab = ev.target.closest('[data-tab]');
    if (tab) { if (shouldCloseSessionMenu) { state.sessionMenuOpen = false; state.pendingDeleteSessionId = null; } if (shouldCloseModelMenu) { state.modelMenuOpen = false; state.accessMenuOpen = false; state.thinkingMenuOpen = false; state.plusMenuOpen = false; } state.tab = tab.dataset.tab; render(); return; }
    const prompt = ev.target.closest('[data-prompt]');
    if (prompt) { fillPrompt(prompt.dataset.prompt || ''); return; }
    const actionEl = ev.target.closest('[data-action]');
    if (!actionEl) { if (shouldCloseSessionMenu || shouldCloseModelMenu) { state.sessionMenuOpen = false; state.pendingDeleteSessionId = null; state.modelMenuOpen = false; state.accessMenuOpen = false; state.thinkingMenuOpen = false; state.plusMenuOpen = false; state.plusSkillsOpen = false; render(); } return; }
    const action = actionEl.dataset.action;
    // render 卡片的交互全部委托给 render-blocks.js，本文件不实现它们
    if (action.indexOf('blk-') === 0 && App.handleBlockAction) {
      if (await App.handleBlockAction(action, actionEl)) return;
    }
    if (action === 'toggle-reasoning') {
      const m = state.messages[Number(actionEl.dataset.idx)];
      if (m) { m.reasoningOpen = !m.reasoningOpen; render(); }
      return;
    }
    if (action === 'navigate-proposal-item') {
      const tc = findProposalByRid(actionEl.dataset.rid);
      if (!tc) return;
      const items = Array.isArray(tc.args && tc.args.changes) && tc.args.changes.length ? tc.args.changes : ((tc.args && tc.args.edits) || []);
      const edit = items[Number(actionEl.dataset.idx) || 0];
      if (edit) await navigateProposalEdit(edit).catch(e => { state.error = e.message || String(e); render(); });
      return;
    }
    if (action === 'toggle-thinking-menu') {
      state.thinkingMenuOpen = !state.thinkingMenuOpen;
      state.modelMenuOpen = false;
      state.accessMenuOpen = false;
      state.plusMenuOpen = false;
      render();
      return;
    }
    if (action === 'set-thinking') {
      state.settings.thinking = actionEl.dataset.level || 'medium';
      state.thinkingMenuOpen = false;
      saveSettings();
      render();
      return;
    }
    if (action === 'toggle-user-prompt') {
      const m = state.messages[Number(actionEl.dataset.idx)];
      if (m && m.role === 'user' && isCollapsiblePrompt(userPromptText(m))) {
        m.promptCollapsed = actionEl.dataset.collapsed !== 'true';
        persistCurrentSession();
        render();
      }
    } else if (action === 'copy-user-prompt') {
      const m = state.messages[Number(actionEl.dataset.idx)];
      if (!m || m.role !== 'user') return;
      const text = userPromptText(m);
      const originalLabel = actionEl.textContent;
      try {
        if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(text);
        else {
          const ta = document.createElement('textarea');
          ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
          document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove();
        }
        actionEl.textContent = t('promptCopied');
      } catch {
        actionEl.textContent = t('statusError');
      }
      setTimeout(() => { if (actionEl.isConnected) actionEl.textContent = originalLabel; }, 1200);
    } else if (action === 'toggle-session-menu') {
      state.sessionMenuOpen = !state.sessionMenuOpen;
      state.pendingDeleteSessionId = null;
      state.modelMenuOpen = false;
      render();
    } else if (action === 'use-provider') {
      const pid = actionEl.dataset.pid;
      const e = providerEntries().find(x => x.id === pid);
      if (e) {
        state.settings.provider = e.id;
        state.settings.customPrefixUrl = e.baseUrl;
        const selected = ((state.settings.providerModels || {})[e.id] || []).filter(id => !App.isRetiredModel || !App.isRetiredModel(e.id, id));
        const remembered = ((state.settings.recentModels || {})[e.id] || []).find(id => selected.includes(id)) || selected[0] || '';
        state.settings.model = remembered;
        state.settings.apiKey = e.key || '';
        if (remembered && App.rememberModel) App.rememberModel(e.id, remembered);
        state.providerEditing = null;
        state.providerModelManaging = null;
        saveSettings({ immediate: true }); render();
      }
    } else if (action === 'new-provider') {
      state.providerEditing = '__new__';
      state.providerModelManaging = null;
      render();
    } else if (action === 'edit-provider') {
      state.providerEditing = actionEl.dataset.pid;
      state.providerModelManaging = null;
      render();
    } else if (action === 'manage-provider-models') {
      const pid = actionEl.dataset.pid;
      state.providerModelManaging = state.providerModelManaging === pid ? null : pid;
      state.providerEditing = null;
      render();
    } else if (action === 'cancel-edit') {
      state.providerEditing = null;
      render();
    } else if (action === 'save-provider') {
      const name = (document.getElementById('pe-name') || {}).value || '';
      const url = ((document.getElementById('pe-url') || {}).value || '').trim();
      const key = ((document.getElementById('pe-key') || {}).value || '').trim();
      const protocol = ((document.getElementById('pe-protocol') || {}).value || 'openai') === 'anthropic' ? 'anthropic' : 'openai';
      if (!name.trim() || !url) { state.error = t('serviceInfoRequired'); render(); return; }
      const pid = actionEl.dataset.pid;
      state.settings.serviceProviders = state.settings.serviceProviders || [];
      if (pid) {
        const provider = state.settings.serviceProviders.find(x => x.id === pid);
        if (provider) {
          Object.assign(provider, { name: name.trim(), baseUrl: url, apiKey: key, protocol });
          if (!Array.isArray(provider.models)) provider.models = [];
          if (state.settings.provider === pid) {
            state.settings.customPrefixUrl = url;
            state.settings.apiKey = key;
            state.settings.providerProtocol = protocol;
          }
        }
      } else {
        const id = App.id().slice(0, 8);
        const newPid = 'custom:' + id;
        state.settings.serviceProviders.push({ id: newPid, name: name.trim(), baseUrl: url, apiKey: key, protocol, models: [] });
        state.settings.provider = newPid;
        state.settings.customPrefixUrl = url;
        state.settings.model = '';
        state.settings.apiKey = key;
        state.settings.providerModels = state.settings.providerModels || {};
        state.settings.knownModels = state.settings.knownModels || {};
        state.settings.providerModels[newPid] = [];
        state.settings.knownModels[newPid] = [];
        state.providerModelManaging = newPid;
      }
      state.providerEditing = null;
      state.error = null;
      if (state.settings.model && App.rememberModel) App.rememberModel(state.settings.provider, state.settings.model);
      saveSettings({ immediate: true }); render();
    } else if (action === 'delete-provider') {
      const pid = actionEl.dataset.pid;
      state.settings.serviceProviders = (state.settings.serviceProviders || []).filter(x => x.id !== pid);
      for (const key of ['providerModels', 'knownModels', 'recentModels', 'hiddenModels']) {
        if (state.settings[key]) delete state.settings[key][pid];
      }
      if (state.settings.provider === pid) {
        const fallback = state.settings.serviceProviders[0];
        state.settings.provider = fallback ? fallback.id : '';
        state.settings.customPrefixUrl = fallback ? fallback.baseUrl : '';
        state.settings.apiKey = fallback ? (fallback.apiKey || '') : '';
        state.settings.providerProtocol = (fallback && fallback.protocol) || 'openai';
        const selected = fallback ? ((state.settings.providerModels || {})[fallback.id] || []) : [];
        state.settings.model = selected[0] || '';
      }
      state.providerEditing = null;
      state.providerModelManaging = null;
      if (state.settings.model && App.rememberModel) App.rememberModel(state.settings.provider, state.settings.model);
      saveSettings({ immediate: true }); render();
    } else if (action === 'add-provider-model') {
      const pid = actionEl.dataset.pid;
      const input = document.getElementById(`provider-model-input-${providerDomId(pid)}`);
      const id = input && input.value.trim();
      if (id) {
        state.settings.knownModels = state.settings.knownModels || {};
        const list = state.settings.knownModels[pid] = state.settings.knownModels[pid] || [];
        if (!list.includes(id)) list.push(id);
        input.value = '';
        saveSettings({ immediate: true });
      }
      render();
    } else if (action === 'fetch-provider-models') {
      const pid = actionEl.dataset.pid;
      const entry = providerEntries().find(x => x.id === pid);
      if (!entry) return;
      state.error = null;
      try {
        const base = encodeURIComponent(entry.baseUrl || '');
        const key = encodeURIComponent(entry.key || '');
        const protocol = entry.protocol === 'anthropic' ? 'anthropic' : 'openai';
        const res = await fetch(App.localApiUrl(`/api/models?base=${base}&key=${key}&protocol=${protocol}`));
        const data = await res.json();
        if (!res.ok || !Array.isArray(data.data)) throw new Error(data.error || `HTTP ${res.status}`);
        const ids = Array.from(new Set(data.data.map(x => x && x.id).filter(Boolean))).filter(id => !App.isRetiredModel || !App.isRetiredModel(pid, id));
        state.settings.knownModels = state.settings.knownModels || {};
        state.settings.knownModels[pid] = Array.from(new Set((state.settings.knownModels[pid] || []).concat(ids)));
        saveSettings({ immediate: true });
        render();
      } catch (e) {
        state.error = String(e.message || e);
        render();
      }
    } else if (action === 'toggle-provider-model') {
      const pid = actionEl.dataset.pid;
      const id = actionEl.dataset.model;
      state.settings.providerModels = state.settings.providerModels || {};
      const list = state.settings.providerModels[pid] = state.settings.providerModels[pid] || [];
      const index = list.indexOf(id);
      if (index >= 0) list.splice(index, 1); else list.push(id);
      if (state.settings.provider === pid) {
        if (!state.settings.model && index < 0) state.settings.model = id;
        else if (state.settings.model === id && index >= 0) state.settings.model = list[0] || '';
        if (state.settings.model && App.rememberModel) App.rememberModel(pid, state.settings.model);
      }
      saveSettings({ immediate: true });
      render();
    } else if (action === 'delete-skill') {
      state.settings.customSkills = (state.settings.customSkills || []).filter(x => x.id !== actionEl.dataset.sid);
      if (state.activeSkillId === actionEl.dataset.sid) state.activeSkillId = null;
      saveSettings(); render();
    } else if (action === 'scan-skills') {
      state.error = null;
      state.skillImportNotice = null;
      try {
        await importSkillFiles(await skillFilesFromServer(), false, 'skillScanResult', true);
      } catch (e) {
        state.error = String(e.message || e);
        render();
      }
    } else if (action === 'slash-pick') {
      const skill = allSkillsList().find(item => item.id === actionEl.dataset.skill);
      selectSlashSkill(skill);
      render();
    } else if (action === 'remove-attachment') {
      state.attachments.splice(Number(actionEl.dataset.idx) || 0, 1);
      render();
    } else if (action === 'remove-pending-image') {
      state.pendingImages.splice(Number(actionEl.dataset.idx) || 0, 1);
      render();
    } else if (action === 'pick-skill') {
      state.activeSkillId = state.activeSkillId === actionEl.dataset.skill ? null : actionEl.dataset.skill;
      state.plusMenuOpen = false;
      state.plusSkillsOpen = false;
      state.skillSearch = '';
      render();
    } else if (action === 'set-access-mode') {
      if (typeof App.setAccessMode === 'function') App.setAccessMode(actionEl.dataset.mode);
      else state.settings.accessMode = actionEl.dataset.mode;
      state.accessMenuOpen = false;
      saveSettings(); render();
    } else if (action === 'toggle-advanced-access') {
      const enabled = typeof App.currentAdvancedAccess === 'function' ? App.currentAdvancedAccess() : false;
      if (typeof App.setAdvancedAccess === 'function') App.setAdvancedAccess(!enabled);
      saveSettings(); render();
    } else if (action === 'toggle-tool-group') {
      // 折叠组开关：按 rid 找到组内第一条，取「从它起连续同名」的段（与 group 切分一致）切 _groupOpen
      const rid = actionEl.dataset.rid;
      for (const msg of state.messages) {
        if (!msg.toolCalls) continue;
        const idx = msg.toolCalls.findIndex(tc => tc.rid === rid);
        if (idx < 0) continue;
        const first = msg.toolCalls[idx];
        let end = idx + 1;
        while (end < msg.toolCalls.length && msg.toolCalls[end].name === first.name && !!msg.toolCalls[end].rid) end++;
        App.toolGroups.toggle(msg.toolCalls.slice(idx, end));
        render();
        break;
      }
    } else if (action === 'toggle-memory') {
      // setEnabled 内部已走 saveSettings，这里只管重渲染
      if (App.memory) { App.memory.setEnabled(!App.memory.enabled()); render(); }
    } else if (action === 'delete-memory') {
      if (App.memory && actionEl.dataset.key) { App.memory.removeTopic(actionEl.dataset.key); render(); }
    } else if (action === 'clear-memory') {
      if (App.memory && window.confirm(t('memoryClearConfirm'))) { App.memory.clearAll(); render(); }
    } else if (action === 'toggle-plus') {
      state.plusMenuOpen = !state.plusMenuOpen;
      state.plusSkillsOpen = false;
      state.modelMenuOpen = false;
      state.accessMenuOpen = false;
      state.thinkingMenuOpen = false;
      render();
    } else if (action === 'plus-skills') {
      state.plusSkillsOpen = !state.plusSkillsOpen;
      render();
    } else if (action === 'clear-skill') {
      state.activeSkillId = null;
      render();
    } else if (action === 'toggle-model-menu') {
      state.modelMenuOpen = !state.modelMenuOpen;
      state.accessMenuOpen = false;
      state.thinkingMenuOpen = false;
      state.sessionMenuOpen = false;
      render();
    } else if (action === 'toggle-access-menu') {
      state.accessMenuOpen = !state.accessMenuOpen;
      state.modelMenuOpen = false;
      state.thinkingMenuOpen = false;
      state.plusMenuOpen = false;
      state.sessionMenuOpen = false;
      render();
    } else if (action === 'pick-quick-model') {
      state.settings.model = actionEl.dataset.model;
      if (App.rememberModel) App.rememberModel(state.settings.provider, state.settings.model);
      saveSettings({ immediate: true });
      state.modelMenuOpen = false;
      render();
    } else if (action === 'toggle-tool') {
      const rid = actionEl.dataset.rid;
      for (const m of state.messages) {
        if (m.toolCalls && m.toolCalls.some(tc => tc.rid === rid)) {
          const tc = m.toolCalls.find(x => x.rid === rid);
          tc.expanded = !tc.expanded;
          break;
        }
      }
      render();
    } else if (action === 'pick-diff-version') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (tc && tc.proposal && !tc.proposal.decision) {
        // 切版本丢弃未应用的编辑暂存：pending 不分版本存，残留会把 v0 的文字带进 v1（review P2-3）
        delete tc.proposal.pendingReplacement;
        tc.proposal.activeIdx = Number(actionEl.dataset.idx) || 0;
        render();
      }
    } else if (action === 'apply-batch-item' || action === 'skip-batch-item' || action === 'retry-batch-item') {
      const t2 = findPendingProposal(actionEl.dataset.rid);
      if (!t2 || !t2.batch || t2.batch.settled) return;
      const i = Number(actionEl.dataset.idx) || 0;
      if (action === 'skip-batch-item') {
        if (t2.batch.items[i] !== 'pending' && t2.batch.items[i] !== 'stale') return;
        // 条目级拒绝同样展开理由区（v151：用户 真机预期「拒绝必附理由」不分粒度）；
        // pendingSkip 记录待跳过条目，提交理由时只跳过这一条
        t2.batch.declineOpen = true;
        t2.batch.pendingSkip = i;
        render();
        return;
      }
      // 失败条目重试：先复位成 pending，再走同一条应用路径
      if (action === 'retry-batch-item') {
        if (t2.batch.items[i] !== 'error') return;
        t2.batch.items[i] = 'pending';
        if (t2.batch.itemErrors) delete t2.batch.itemErrors[i];
      }
      if (t2.batch.items[i] !== 'pending') return;
      t2.batch.items[i] = 'applying';
      render();
      try {
        const c = (t2.args.changes || [])[i];
        if (!App.hasOffice() || !App.host.applyEdit) throw new Error(t('demo'));
        takeReplacementOverride(t2, i, c);
        const out = await App.host.applyEdit(c);
        followInBackground(out); // 【34.4】写入已确认即据实结算，跟随导航后台进行不拖结算
        if (App.queueEditFlash) App.queueEditFlash(c, out);
        t2.batch.items[i] = 'applied';
        t2.batch.appliedCount++;
        // itemVerifs 兜底：正常创建路径总有该字段（拦截时初始化），旧会话/测试数据缺失时
        // 不该把「成功」拖进 catch 变 error
        t2.batch.itemVerifs = t2.batch.itemVerifs || {};
        t2.batch.itemVerifs[i] = (out && out.verification) || null;
      } catch (e) {
        if (e && e.code === 'STALE_EDIT') {
          t2.batch.items[i] = 'stale';
          t2.batch.itemErrors = t2.batch.itemErrors || {};
          t2.batch.itemErrors[i] = { message: e.message || t('proposalStaleHint'), target: e.target || null, currentText: e.currentText || '', currentValues: e.currentValues || null };
        } else {
          t2.batch.items[i] = 'error';
          t2.batch.itemErrors = t2.batch.itemErrors || {};
          // 具体原因与结构化诊断进卡内（不再只闪全局错误条），用户能看出为什么失败、值不值得重试
          t2.batch.itemErrors[i] = { message: e.message || String(e), ...diagnosticsOf(e) };
          state.error = `${(t2.args.changes[i] || {}).label || ''}: ${e.message || String(e)}`;
        }
      }
      settleBatchIfDone(t2);
      persistCurrentSession(); render();
    } else if (action === 'apply-batch') {
      const t2 = findPendingProposal(actionEl.dataset.rid);
      if (!t2 || !t2.batch || t2.batch.settled || t2.batch.applyingAll) return;
      t2.batch.applyingAll = true;
      // 逐项置 applying 并渲染，用户能看到逐条打勾而不是界面冻结到最后
      try {
        for (let i = 0; i < t2.batch.items.length; i++) {
          if (t2.batch.items[i] !== 'pending') continue;
          t2.batch.items[i] = 'applying';
          render();
          try {
            const c = (t2.args.changes || [])[i];
            if (!App.hasOffice() || !App.host.applyEdit) throw new Error(t('demo'));
            takeReplacementOverride(t2, i, c);
            const out = await App.host.applyEdit(c);
            followInBackground(out); // 【34.4】写入已确认即据实结算，跟随导航后台进行不拖结算
            if (App.queueEditFlash) App.queueEditFlash(c, out);
            t2.batch.items[i] = 'applied';
            t2.batch.appliedCount++;
            t2.batch.itemVerifs = t2.batch.itemVerifs || {};
            t2.batch.itemVerifs[i] = (out && out.verification) || null;
          } catch (e) {
            if (e && e.code === 'STALE_EDIT') {
              t2.batch.items[i] = 'stale';
              t2.batch.itemErrors = t2.batch.itemErrors || {};
              t2.batch.itemErrors[i] = { message: e.message || t('proposalStaleHint'), target: e.target || null, currentText: e.currentText || '', currentValues: e.currentValues || null };
            } else {
              t2.batch.items[i] = 'error';
              t2.batch.itemErrors = t2.batch.itemErrors || {};
              // 【49-R4】批量应用同样保留结构化诊断（此前只有 message/writeState/shapeCreated）
              t2.batch.itemErrors[i] = { message: e.message || String(e), ...diagnosticsOf(e) };
            }
          }
          render();
        }
      } finally {
        t2.batch.applyingAll = false;
      }
      settleBatchIfDone(t2);
      persistCurrentSession(); render();
    } else if (action === 'refresh-batch-item') {
      const t2 = findPendingProposal(actionEl.dataset.rid);
      if (!t2 || !t2.batch || t2.batch.settled) return;
      const i = Number(actionEl.dataset.idx) || 0;
      if (t2.batch.items[i] !== 'stale') return;
      t2.batch.items[i] = 'refresh_requested';
      settleBatchIfDone(t2);
      persistCurrentSession(); render();
    } else if (action === 'decline-batch') {
      // A4（v146）：整卡拒绝两步走——先展开结构化理由区，提交才结算
      const t2 = findPendingProposal(actionEl.dataset.rid);
      if (!t2 || !t2.batch || t2.batch.settled) return;
      t2.batch.declineOpen = !t2.batch.declineOpen;
      render();
    } else if (action === 'toggle-batch-reasoning') {
      const t2 = findPendingProposal(actionEl.dataset.rid);
      if (t2) { const i = Number(actionEl.dataset.idx) || 0; t2.batchReasoningOpen = t2.batchReasoningOpen === i ? -1 : i; render(); }
    } else if (action === 'toggle-diff-reasoning') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (tc && tc.proposal) { tc.proposal.reasoningOpen = !tc.proposal.reasoningOpen; render(); }
    } else if (action === 'apply-diff') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc || !tc.proposal || tc.proposal.decision) return;
      const edits = (tc.args && tc.args.edits) || [];
      const idx = Math.min(tc.proposal.activeIdx || 0, edits.length - 1);
      const edit = edits[idx];
      tc.proposal.decision = 'applying';
      render();
      try {
        if (!App.hasOffice() || !App.host.applyEdit) throw new Error(t('demo'));
        takeReplacementOverride(tc, idx, edit);
        const out = await App.host.applyEdit(edit);
        followInBackground(out); // 【34.4】写入已确认即据实结算，跟随导航后台进行不拖结算
        if (App.queueEditFlash) App.queueEditFlash(edit, out);
        tc.result = out;
        settleProposal(tc, { success: true, applied: true, appliedIdx: idx, label: edit.label || `version ${idx + 1}`, result: out });
      } catch (e) {
        if (e && e.code === 'STALE_EDIT') {
          tc.proposal.decision = 'stale';
          tc.proposal.error = { message: e.message || t('proposalStaleHint'), target: e.target || null, currentText: e.currentText || '', currentValues: e.currentValues || null };
          tc.result = { success: false, stale: true, error: e.message || String(e) };
        } else if (e && e.code === 'WRITE_CANCELLED') {
          // 【24.2】停止取消的写入：不是应用错误，按已停止结算（与原停止流程同语义）；
          // 不得把它当普通定位/格式失败再触发任何会写正文的回退
          tc.proposal.decision = 'stopped';
          tc.proposal.error = { message: e.message || String(e) };
          tc.result = { success: false, stopped: true, cancelled: true, error: e.message || String(e) };
          settleProposal(tc, { success: false, declined: true, stopped: true, cancelled: true });
        } else {
          tc.proposal.error = { message: e.message || String(e) };
          tc.result = { success: false, error: e.message || String(e) };
          settleProposal(tc, { success: false, applied: false, error: e.message || String(e) });
        }
      }
      persistCurrentSession();
      render();
    } else if (action === 'refresh-diff') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc || !tc.proposal || tc.proposal.decision !== 'stale') return;
      const edits = (tc.args && tc.args.edits) || [];
      const idx = Math.min(tc.proposal.activeIdx || 0, Math.max(0, edits.length - 1));
      const edit = edits[idx] || {};
      settleProposal(tc, { success: true, refreshRequested: [{ index: idx, label: edit.label, target: edit.target || null, previousFind: edit.find || '', error: tc.proposal.error || null }] });
      persistCurrentSession(); render();
    } else if (action === 'skip-stale-diff') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc || !tc.proposal || tc.proposal.decision !== 'stale') return;
      settleProposal(tc, { success: false, declined: true, stale: true });
      persistCurrentSession(); render();
    } else if (action === 'decline-diff') {
      // A4（v146）：拒绝两步走——先展开结构化理由区，提交才结算
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc || !tc.proposal || tc.proposal.decision) return;
      tc.proposal.declineOpen = !tc.proposal.declineOpen;
      render();
    } else if (action === 'pick-decline-reason') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc) return;
      const store = actionEl.dataset.kind === 'single' ? tc.proposal : tc.batch;
      if (!store) return;
      store.declineReason = store.declineReason === actionEl.dataset.reason ? '' : actionEl.dataset.reason;
      store.declineNeedReason = false;
      render();
    } else if (action === 'cancel-decline') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc) return;
      const store = actionEl.dataset.kind === 'single' ? tc.proposal : tc.batch;
      if (!store) return;
      store.declineOpen = false;
      store.declineNeedReason = false;
      if (store.pendingSkip != null) store.pendingSkip = null;
      render();
    } else if (action === 'confirm-decline-diff') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc || !tc.proposal || tc.proposal.decision) return;
      const p = tc.proposal;
      const note = String(p.declineNote || '').trim();
      if (!p.declineReason && !note) { p.declineNeedReason = true; render(); return; }
      const reason = { selected: p.declineReason ? t(p.declineReason) : '', note };
      settleProposal(tc, { success: false, declined: true, reason, observation: declineObservation(reason) });
      persistCurrentSession();
      render();
    } else if (action === 'confirm-decline-batch') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc || !tc.batch || tc.batch.settled) return;
      const b = tc.batch;
      const note = String(b.declineNote || '').trim();
      if (!b.declineReason && !note) { b.declineNeedReason = true; render(); return; }
      const reason = { selected: b.declineReason ? t(b.declineReason) : '', note };
      const only = b.pendingSkip;
      if (only != null) {
        // 条目级拒绝（v151）：只跳过该条，理由记到该条结果上
        if (b.items[only] === 'pending' || b.items[only] === 'stale') b.items[only] = 'skipped';
        b.skipReasons = b.skipReasons || {};
        b.skipReasons[only] = reason;
        b.pendingSkip = null;
      } else {
        for (let i = 0; i < b.items.length; i++) {
          if (b.items[i] === 'pending' || b.items[i] === 'stale') b.items[i] = 'skipped';
        }
      }
      settleBatchIfDone(tc);
      persistCurrentSession();
      render();
    } else if (action === 'toggle-diff-edit') {
      // A4 第二层：卡上编辑替换文字
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc || !tc.proposal || tc.proposal.decision) return;
      tc.proposal.editing = !tc.proposal.editing;
      render();
    } else if (action === 'toggle-batch-edit') {
      const tc = findPendingProposal(actionEl.dataset.rid);
      if (!tc || !tc.batch || tc.batch.settled) return;
      const i = Number(actionEl.dataset.idx) || 0;
      tc.batch.editing = tc.batch.editing === i ? -1 : i;
      render();
    } else if (action === 'approve-plan') {
      // 只读建议模式下批准方案：切到「审核后修改」继续，改动仍会逐项出 diff 卡确认。
      // 顺手清输入框：用户可能先点过「提修改意见」，残留前缀会变成悬空消息
      if (typeof App.setAccessMode === 'function') App.setAccessMode('confirm');
      else state.settings.accessMode = 'confirm';
      saveSettings();
      state.composerDraft = '';
      const draftInput = document.getElementById('chat-input');
      if (draftInput) draftInput.value = '';
      await sendUserMessage(t('approvePlanGo'));
    } else if (action === 'rewind-to') {
      // undo 式回退（v139，借鉴 Hermes /undo）：截断到选中用户消息之前，原文预填回输入框供编辑重发。
      // 只动会话历史，不碰文档；已写入文档的内容用 ⌘Z / 卡片拒绝各自撤销。
      if (state.isWorking) return;
      const idx = Number(actionEl.dataset.idx);
      const target = state.messages[idx];
      if (!target || target.role !== 'user') return;
      const removed = state.messages.length - idx;
      state.messages = state.messages.slice(0, idx);
      persistCurrentSession();
      fillPrompt(String(target.displayContent || target.content || ''));
      if (App.toast) App.toast(t('rewindDone').replace('{n}', Math.ceil(removed / 2)));
      render();
    } else if (action === 'feedback-plan') {
      // 方案修改意见（feedback-as-deny）：引导前缀填入输入框，用户补完发送，agent 修订方案再等批准
      fillPrompt(t('planFeedbackPrefix'));
    } else if (action === 'toggle-selection-target' || action === 'refresh-selection') {
      if (action === 'toggle-selection-target') state.selectionTarget = !state.selectionTarget;
      if (state.selectionTarget || action === 'refresh-selection') {
        if (App.hasOffice() && App.host && App.host.toolExecutors && App.host.toolExecutors.get_selection) {
          try {
            const sel = await executeToolByName('get_selection', {});
            const text = String((sel && sel.text) || '');
            // collapsed（光标只点进某段）时 host 会带回所在段文本；既无选中文本也无段文本才算真空
            const paragraphText = String((sel && sel.paragraph && sel.paragraph.text) || '');
            const collapsed = !!(sel && sel.collapsed);
            state.selectionSnapshot = { text, paragraphText, collapsed, at: now() };
            // 没抓到任何文字要立即提示（原先写反成清空错误：用户以为抓到了，发送时却被静默丢弃）
            if (!text.trim() && !paragraphText.trim()) state.error = t('selEmptyNotice');
            else if (state.error === t('selEmptyNotice')) state.error = null;
          } catch (e) {
            state.error = e.message || String(e);
            if (action === 'refresh-selection') state.selectionTarget = false;
          }
        } else {
          state.selectionTarget = false;
          state.selectionSnapshot = null;
        }
      } else {
        state.selectionSnapshot = null;
      }
      render();
    } else if (action === 'clear-selection') {
      state.selectionTarget = false;
      state.selectionSnapshot = null;
      render();
    } else if (action === 'toggle-changelog') {
      state.changeLogOpen = !state.changeLogOpen;
      render();
    } else if (action === 'changelog-jump') {
      const idx = Number(actionEl.dataset.idx) || 0;
      const items = App.buildChangeLogItems ? App.buildChangeLogItems(state.messages) : [];
      const item = items[idx];
      const ref = item && App.changeLogRef ? App.changeLogRef(item.target) : '';
      if (ref) await navigateCitation(ref).catch(e => { state.error = e.message || String(e); });
      render();
    } else if (action === 'regenerate') {
      await regenerate();
    } else if (action === 'send') {
      const message = takeComposerMessage();
      if (message && !state.isWorking) await sendUserMessage(message.text, message.attachments, message.skillId);
      else render();
    } else if (action === 'stop') {
      stopActiveRequest();
    } else if (action === 'hold-draft') {
      // 任务进行中发送 = 排入纠偏队列（不打断），agent loop 下一步开始时自动注入
      const message = takeComposerMessage();
      if (message) state.steerQueue.push(message);
      render();
    } else if (action === 'edit-held') {
      const idx = Number(actionEl.dataset.idx);
      const item = state.steerQueue[idx];
      if (item) {
        state.composerDraft = [item.text, state.composerDraft].filter(Boolean).join('\n');
        state.attachments = (item.attachments || []).concat(state.attachments || []);
        state.activeSkillId = item.skillId || state.activeSkillId;
        state.steerQueue.splice(idx, 1);
      }
      render();
    } else if (action === 'steer-held') {
      // 立即打断当前任务并发送这条（排队消息的「我现在就要说」出口）
      const idx = Number(actionEl.dataset.idx);
      const item = state.steerQueue[idx];
      if (!item) return;
      state.steerQueue.splice(idx, 1);
      if (state.isWorking) {
        state.pendingInterruptMessage = item;
        stopActiveRequest();
      } else scheduleComposerMessage(item);
    } else if (action === 'dismiss-recovery') {
      state.recoveredTaskNotice = false;
      render();
    } else if (action === 'adopt-shared-settings') {
      // 【任务 C】冲突恢复：采用共享层当前设置（后果已由警示文案明示，点击即确认）
      if (App.adoptSharedSettings) App.adoptSharedSettings();
    } else if (action === 'new-chat') {
      stopActiveRequest(); // 先停掉进行中的请求并结算未决提案，否则 agent loop 会永久悬挂
      // 【W3W5-R2】草稿按会话归属：离开前把草稿存回离开的会话，新会话从空草稿开始
      const leaving = state.sessions.find(x => x.id === state.currentSessionId);
      if (leaving) leaving.draft = state.composerDraft;
      const s = makeSession(); state.sessions.unshift(s); state.currentSessionId = s.id; state.messages = []; state.composerDraft = ''; state.steerQueue = []; state.activeSkillId = null; state.sessionMenuOpen = false; state.pendingDeleteSessionId = null; saveSessions(); render();
    } else if (action === 'clear') {
      stopActiveRequest();
      state.steerQueue = [];
      state.messages = []; persistCurrentSession(); render();
    } else if (action === 'toggle-theme') {
      state.theme = state.theme === 'dark' ? 'light' : 'dark'; setStoredItem(STORAGE_KEYS.theme, state.theme); render();
    } else if (action === 'toggle-locale') {
      state.locale = state.locale === 'zh' ? 'en' : 'zh'; setStoredItem(STORAGE_KEYS.locale, state.locale); render();
    } else if (action === 'toggle-follow') {
      state.settings.followMode = !state.settings.followMode; saveSettings(); render();
    } else if (action === 'switch-session') {
      stopActiveRequest(); // 同上：先结算当前会话的未决提案再切换
      // 【W3W5-R2】草稿按会话归属：草稿跟随离开的会话暂存、载入目标会话自己的草稿，
      // 保留用户在每个会话里的未发送输入（不再全局共用一个草稿值）
      const leaving2 = state.sessions.find(x => x.id === state.currentSessionId);
      if (leaving2) leaving2.draft = state.composerDraft;
      const s = state.sessions.find(x => x.id === actionEl.dataset.id); if (s) { state.currentSessionId = s.id; state.messages = s.messages || []; state.composerDraft = String(s.draft || ''); state.steerQueue = []; state.activeSkillId = null; state.sessionMenuOpen = false; state.pendingDeleteSessionId = null; App.markSessionRecoveredIfNeeded && App.markSessionRecoveredIfNeeded(s); render(); }
    } else if (action === 'delete-session') {
      state.pendingDeleteSessionId = actionEl.dataset.id;
      render();
    } else if (action === 'cancel-delete-session') {
      state.pendingDeleteSessionId = null;
      render();
    } else if (action === 'confirm-delete-session') {
      const deleteId = actionEl.dataset.id;
      state.sessions = state.sessions.filter(x => x.id !== deleteId);
      state.pendingDeleteSessionId = null;
      if (state.currentSessionId === deleteId) { state.currentSessionId = null; state.messages = []; }
      ensureSession(); saveSessions(); render();
    } else if (action === 'run-tool') {
      await runManualTool();
    }
  });

  document.addEventListener('input', (ev) => {
    // A4（v146）：卡上编辑框与拒绝附言的实时暂存——只写状态不 render，保住输入焦点
    const el = ev.target;
    if (!el || !el.classList) return;
    if (el.classList.contains('diff-edit-input')) {
      const tc = findPendingProposal(el.dataset.rid, { silent: true });
      if (!tc) return;
      const idx = Number(el.dataset.idx) || 0;
      // 按卡型分流写（与 pendingReplacementOf 的读法对称，防跨条目污染）
      if (tc.batch) { tc.batch.pendingReplacements = tc.batch.pendingReplacements || {}; tc.batch.pendingReplacements[idx] = el.value; }
      else if (tc.proposal) tc.proposal.pendingReplacement = el.value;
    } else if (el.classList.contains('dr-note')) {
      const tc = findPendingProposal(el.dataset.rid, { silent: true });
      if (!tc) return;
      const store = el.dataset.kind === 'single' ? tc.proposal : tc.batch;
      if (store) store.declineNote = el.value;
    }
  });

  document.addEventListener('change', async (ev) => {
    // 记忆触发词（v145 A1）：失焦保存，命中词的记忆在 prompt 清单置顶
    if (ev.target.dataset.action === 'edit-memory-triggers') {
      try { App.memory.setTopicTriggers(ev.target.dataset.key, ev.target.value); }
      catch (e) { if (App.toast) App.toast(e.message || String(e), { tone: 'error' }); }
      render();
      return;
    }
    const bind = ev.target.dataset.bind;
    if (bind) {
      if (ev.target.type === 'checkbox') state.settings[bind] = ev.target.checked;
      else state.settings[bind] = ev.target.value;
      if (bind === 'provider') {
        const provider = providerEntries().find(item => item.id === state.settings.provider);
        if (provider) {
          state.settings.customPrefixUrl = provider.baseUrl || state.settings.customPrefixUrl;
          const selected = ((state.settings.providerModels || {})[provider.id] || []);
          state.settings.model = selected[0] || '';
          state.settings.apiKey = provider.key || '';
          state.settings.providerProtocol = provider.protocol || 'openai';
        }
      }
      saveSettings({ immediate: true }); render(); return;
    }
    if (ev.target.dataset.action === 'pick-base-url') {
      if (ev.target.value !== '__custom__') state.settings.customPrefixUrl = ev.target.value;
      saveSettings({ immediate: true }); render(); return;
    }
    if (ev.target.dataset.action === 'pick-model') {
      if (ev.target.value !== '__custom__') state.settings.model = ev.target.value;
      if (App.rememberModel) App.rememberModel(state.settings.provider, state.settings.model);
      saveSettings({ immediate: true }); render(); return;
    }
    if (ev.target.id === 'manual-tool') {
      const args = document.getElementById('manual-args');
      const fn = App.host && App.host.defaultArgsForTool;
      if (args && fn) args.value = fn(ev.target.value);
    }
    if ((ev.target.id === 'skill-file' || ev.target.id === 'skill-folder') && ev.target.files && ev.target.files.length) {
      const files = Array.from(ev.target.files);
      ev.target.value = '';
      state.error = null;
      state.skillImportNotice = null;
      try {
        await importSkillFiles(files, true, 'skillImportResult');
      } catch (e) {
        state.error = String(e.message || e);
        render();
      }
    }
    if (ev.target.id === 'file-input' && ev.target.files && ev.target.files.length) {
      const files = Array.from(ev.target.files);
      ev.target.value = '';
      state.fileLoading = files.map(f => f.name).join(', ');
      state.plusMenuOpen = false;
      render();
      for (const f of files) {
        try {
          const att = await App.readAttachment(f);
          state.attachments.push(att);
        } catch (e) {
          state.error = `${f.name}: ${e.message || String(e)}`;
        }
      }
      state.fileLoading = null;
      render();
    }
    // 【图片上传】图片不进 attachments（那是解析成文本的参考文件），进 pendingImages：
    // base64 留在本地状态，消息里只注入 attachmentId 引用，由 insert_image 按 id 取用
    if (ev.target.id === 'image-input' && ev.target.files && ev.target.files.length) {
      const files = Array.from(ev.target.files);
      ev.target.value = '';
      state.plusMenuOpen = false;
      let seq = (state.pendingImageSeq = (state.pendingImageSeq || 0));
      for (const f of files) {
        if (f.size > 8 * 1024 * 1024) { state.error = `${f.name}: ${t('imageTooLarge')}`; continue; }
        try {
          const dataUrl = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(String(reader.result || ''));
            reader.onerror = () => reject(new Error('read failed'));
            reader.readAsDataURL(f);
          });
          if (!/^data:image\/(png|jpe?g|gif);base64,/.test(dataUrl)) { state.error = `${f.name}: ${t('imageUnsupported')}`; continue; }
          state.pendingImages.push({ id: `img-${++seq}`, name: f.name, dataUrl });
        } catch (e) {
          state.error = `${f.name}: ${e.message || String(e)}`;
        }
      }
      state.pendingImageSeq = seq;
      render();
    }
  });

  // 【图片拖拽】拖图进对话区即挂起为 pendingImages（与 ➕ 上传图片同一条链路）。
  // dragover 必须 preventDefault 否则浏览器直接打开文件；非图片文件拖入不动（避免劫持文本拖放）。
  document.addEventListener('dragover', (ev) => {
    if (ev.dataTransfer && Array.from(ev.dataTransfer.types || []).includes('Files')) ev.preventDefault();
  });
  document.addEventListener('drop', async (ev) => {
    if (!ev.dataTransfer || !Array.from(ev.dataTransfer.types || []).includes('Files')) return;
    const images = Array.from(ev.dataTransfer.files || []).filter(f => /^image\/(png|jpe?g|gif)$/.test(f.type));
    if (!images.length) return;
    ev.preventDefault();
    let seq = (state.pendingImageSeq = state.pendingImageSeq || 0);
    for (const f of images) {
      if (f.size > 8 * 1024 * 1024) { state.error = `${f.name}: ${t('imageTooLarge')}`; continue; }
      try {
        const dataUrl = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(String(reader.result || ''));
          reader.onerror = () => reject(new Error('read failed'));
          reader.readAsDataURL(f);
        });
        state.pendingImages.push({ id: `img-${++seq}`, name: f.name, dataUrl });
      } catch (e) {
        state.error = `${f.name}: ${e.message || String(e)}`;
      }
    }
    state.pendingImageSeq = seq;
    render();
  });

  document.addEventListener('input', (ev) => {
    // render 的 form 原语输入框（不重渲染，避免打断输入焦点）
    if (App.handleBlockInput && App.handleBlockInput(ev.target)) return;
    const bind = ev.target.dataset.bind;
    if (bind) {
      state.settings[bind] = ev.target.value;
      saveSettings();
    }
    if (ev.target.id === 'font-size-range') {
      const v = Number(ev.target.value) || 3;
      state.settings.uiFontSize = v;
      document.documentElement.dataset.fontsize = String(v);
      saveSettings();
      return;
    }
    if (ev.target.id === 'skill-search') {
      state.skillSearch = ev.target.value;
      render();
      return;
    }
    // 斜杠命令：识别光标前最近的 /命令，不限定在消息开头。
    if (ev.target.id === 'chat-input') {
      const v = ev.target.value;
      const cursor = ev.target.selectionStart == null ? v.length : ev.target.selectionStart;
      const hadDraft = Boolean(String(state.composerDraft || '').trim());
      const hadSlashItems = Boolean(state.slashItems && state.slashItems.length);
      let selectedByCommand = false;
      state.composerDraft = v;
      const hasDraft = Boolean(v.trim());
      const completed = completedSlashFragment(v, cursor);
      const completedSkill = completed ? findSkillByCommand(completed.token) : null;
      const slash = completedSkill ? null : slashFragmentAtCursor(v, cursor);
      if (completedSkill) {
        state.slashRange = completed;
        const nextCursor = selectSlashSkill(completedSkill);
        ev.target.value = state.composerDraft;
        try { ev.target.setSelectionRange(nextCursor, nextCursor); } catch {}
        selectedByCommand = true;
      } else if (slash) {
        const items = slashMatches(slash.token);
        state.slashRange = slash;
        state.slashItems = items.length ? items : null;
        state.slashIdx = 0;
      } else if (state.slashItems) {
        state.slashItems = null;
        state.slashRange = null;
      } else {
        state.slashRange = null;
      }
      const hasSlashItems = Boolean(state.slashItems && state.slashItems.length);
      if (hasSlashItems || hadSlashItems !== hasSlashItems || selectedByCommand) render();
      else if (state.isWorking && hadDraft !== hasDraft) render();
    }
  });
  document.addEventListener('keydown', async (ev) => {
    const proposalCard = ev.target.closest && ev.target.closest('[data-action="navigate-proposal-item"]');
    if (proposalCard && (ev.key === 'Enter' || ev.key === ' ')) {
      ev.preventDefault();
      proposalCard.click();
      return;
    }
    if (ev.target.id === 'chat-input' && state.slashItems && state.slashItems.length) {
      if (ev.key === 'ArrowDown') { ev.preventDefault(); state.slashIdx = Math.min(state.slashIdx + 1, state.slashItems.length - 1); render(); return; }
      if (ev.key === 'ArrowUp') { ev.preventDefault(); state.slashIdx = Math.max(state.slashIdx - 1, 0); render(); return; }
      if (ev.key === 'Escape') { state.slashItems = null; state.slashRange = null; render(); return; }
      if (ev.key === 'Enter' && !ev.shiftKey) {
        ev.preventDefault();
        const skill = state.slashItems[state.slashIdx];
        if (skill) {
          state.composerDraft = ev.target.value;
          const nextCursor = selectSlashSkill(skill);
          ev.target.value = state.composerDraft;
          try { ev.target.setSelectionRange(nextCursor, nextCursor); } catch {}
          render();
        }
        return;
      }
    }
    if (ev.target.id === 'chat-input' && ev.key === 'Enter' && !ev.shiftKey) {
      ev.preventDefault();
      state.composerDraft = ev.target.value;
      if (ev.target.value.trim() && !state.isWorking) {
        const message = takeComposerMessage();
        if (message) await sendUserMessage(message.text, message.attachments, message.skillId);
        else render();
      } else if (ev.target.value.trim()) {
        // 任务进行中按回车 = 排入纠偏队列（不打断），下一步自动注入
        const message = takeComposerMessage();
        if (message) state.steerQueue.push(message);
        ev.target.value = '';
        render();
      }
    }
    if (ev.target.id && ev.target.id.startsWith('provider-model-input-') && ev.key === 'Enter') {
      ev.preventDefault();
      const addBtn = ev.target.closest('.provider-model-manager')?.querySelector('[data-action="add-provider-model"]');
      if (addBtn) addBtn.click();
    }
  });

  // ---- 输入区拖拽调高（mousedown 最兼容 Office WebView） ----
  let dragStart = null;
  document.addEventListener('mousedown', (ev) => {
    const bar = ev.target.closest('[data-drag="composer"]');
    if (!bar) return;
    ev.preventDefault();
    const box = document.getElementById('chat-input');
    dragStart = { y: ev.clientY, h: box ? box.offsetHeight : 60, applied: 0 };
    state.dragging = true;
    document.body.style.userSelect = 'none';
  });
  document.addEventListener('mousemove', (ev) => {
    if (!dragStart) return;
    ev.preventDefault();
    const dy = dragStart.y - ev.clientY;
    const h = Math.max(44, Math.min(dragStart.h + dy, Math.round(window.innerHeight * 0.6)));
    const box = document.getElementById('chat-input');
    if (box) box.style.height = h + 'px';
    dragStart.applied = h;
  });
  document.addEventListener('mouseup', () => {
    if (!dragStart) return;
    if (dragStart.applied) {
      state.settings.composerHeight = dragStart.applied;
      saveSettings();
    }
    dragStart = null;
    state.dragging = false;
    document.body.style.userSelect = '';
    render();
  });

  App.render = render;
  App.patchStreamingMessage = patchStreamingMessage;
  App.collapseLatestUserPrompt = collapseLatestUserPrompt;
  App.presentEditProposal = presentEditProposal;
  App.followInBackground = followInBackground;
  App.sendUserMessage = sendUserMessage;
})();
