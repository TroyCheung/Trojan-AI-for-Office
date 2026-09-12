(function () {
  'use strict';
  const App = (window.App = window.App || {});

  function getStoredItem(key) { try { return (typeof window !== 'undefined' && window.localStorage) ? window.localStorage.getItem(key) : null; } catch { return null; } }
  function setStoredItem(key, value) { try { if (typeof window !== 'undefined' && window.localStorage) window.localStorage.setItem(key, value); } catch {} }
  // 与 setStoredItem 的唯一区别：报告写入是否成功（配额满时 localStorage.setItem 抛错）。
  // 会话持久化必须感知失败——曾经静默吞错导致整段对话丢失。
  function trySetStoredItem(key, value) {
    try {
      if (typeof window !== 'undefined' && window.localStorage) { window.localStorage.setItem(key, value); return true; }
    } catch {}
    return false;
  }
  function prefersDarkMode() { return typeof window !== 'undefined' && typeof window.matchMedia === 'function' && window.matchMedia('(prefers-color-scheme: dark)').matches; }
  function getNavigatorLanguage() { return (typeof navigator !== 'undefined' && navigator.language) ? navigator.language : 'en'; }

  function t(key) {
    const hostI18n = App.host && App.host.i18n && App.host.i18n[state.locale];
    if (hostI18n && hostI18n[key] != null) return hostI18n[key];
    return (App.I18N[state.locale] && App.I18N[state.locale][key]) || key;
  }
  function loadLocale() { const v = getStoredItem(App.STORAGE_KEYS.locale); return v === 'en' || v === 'zh' ? v : (getNavigatorLanguage().startsWith('zh') ? 'zh' : 'en'); }
  const storedSettingsText = getStoredItem(App.STORAGE_KEYS.settings);
  const hadStoredSettings = Boolean(storedSettingsText);
  function normalizedSettings(value) {
    const out = Object.assign({}, App.DEFAULT_SETTINGS, value && typeof value === 'object' ? value : {});
    for (const key of ['providerModels', 'hiddenModels', 'recentModels', 'knownModels']) {
      if (!out[key] || typeof out[key] !== 'object' || Array.isArray(out[key])) out[key] = {};
    }
    if (!Array.isArray(out.customProviders)) out.customProviders = [];
    if (!Array.isArray(out.customSkills)) out.customSkills = [];
    if (!out.accessModeByHost || typeof out.accessModeByHost !== 'object' || Array.isArray(out.accessModeByHost)) out.accessModeByHost = {};
    if (!out.advancedAccessByHost || typeof out.advancedAccessByHost !== 'object' || Array.isArray(out.advancedAccessByHost)) out.advancedAccessByHost = {};
    if (typeof App.normalizeCustomSkillEntries === 'function') {
      out.customSkills = App.normalizeCustomSkillEntries(out.customSkills);
    }
    if (!Array.isArray(out.serviceProviders)) {
      const migrated = [];
      for (const id of Object.keys(App.PROVIDERS || {})) {
        const provider = App.PROVIDERS[id];
        migrated.push({
          id,
          name: provider.label,
          // dialect：thinking 等参数的方言。缺省时 dialects.js 会按 baseUrl 自动推断。
          dialect: id === 'zhipu' ? 'zhipu' : (id === 'deepseek' ? 'deepseek' : 'openai'),
          baseUrl: ((provider.baseUrls || [])[0] || {}).value || '',
          apiKey: provider.apiKey || '',
          models: Array.isArray(provider.models) ? provider.models.slice() : []
        });
      }
      for (const cp of out.customProviders) {
        migrated.push({
          id: 'custom:' + cp.id,
          name: cp.name,
          baseUrl: cp.baseUrl,
          apiKey: cp.apiKey || '',
          models: Array.isArray(cp.models) ? cp.models.slice() : []
        });
      }
      out.serviceProviders = migrated;
    }
    // 旧版把自定义服务的模型直接存在 provider.models 中。
    // 新版将它们迁移为该服务已选的常用模型，保持原有使用方式。
    for (const cp of out.customProviders) {
      const pid = 'custom:' + cp.id;
      const legacy = (Array.isArray(cp.models) ? cp.models : []).map(model => model && (model.id || model)).filter(Boolean);
      out.providerModels[pid] = Array.from(new Set((out.providerModels[pid] || []).concat(legacy)));
    }
    for (const key of ['providerModels', 'recentModels', 'knownModels']) {
      if (Array.isArray(out[key].deepseek)) out[key].deepseek = out[key].deepseek.filter(id => !isRetiredModel('deepseek', id));
    }
    if (isRetiredModel(out.provider, out.model)) {
      out.model = (out.recentModels[out.provider] || [])[0]
        || (out.providerModels[out.provider] || [])[0]
        || (out.knownModels[out.provider] || [])[0]
        || '';
    }
    if (out.provider && out.model && !isRetiredModel(out.provider, out.model)) {
      out.recentModels[out.provider] = Array.from(new Set([out.model].concat(out.recentModels[out.provider] || []))).slice(0, 12);
      out.providerModels[out.provider] = Array.from(new Set([out.model].concat(out.providerModels[out.provider] || [])));
    }
    out.settingsUpdatedAt = Number(out.settingsUpdatedAt) || 0;
    return out;
  }
  // 本地 /api/* 服务需要启动时生成的 token（见 server.py 安全模型）。
  // token 由 server.py 注入 taskpane.html；没有 token（如静态打开页面）时原样返回，调用会安全地 403。
  function localApiUrl(url) {
    const token = typeof window !== 'undefined' ? String(window.__LOCAL_API_TOKEN__ || '') : '';
    if (!token || token.startsWith('__')) return url;
    return url + (url.includes('?') ? '&' : '?') + 'token=' + encodeURIComponent(token);
  }
  function loadSettings() { try { return normalizedSettings(JSON.parse(storedSettingsText || '{}')); } catch { return normalizedSettings({}); } }
  let syncTimer = null;
  // 【2026-08-31 事故防线】贫瘠配置禁止覆盖丰富配置：本地一个带 key 的服务商都没有、
  // 而共享层有 → 拒绝上传（当天 Excel 挂载后把残留默认态推上共享层，抹掉了全部配置）。
  // 正常操作（本地本来就有 key 的服务商）不受影响。
  function providersWithKeys(settings) {
    const list = Array.isArray(settings && settings.serviceProviders) ? settings.serviceProviders : [];
    return list.filter(p => p && typeof p.apiKey === 'string' && p.apiKey.trim()).length;
  }
  function sharedUploadBlockedBy(remote) {
    const localKeys = providersWithKeys(state.settings);
    const remoteKeys = providersWithKeys(remote);
    return localKeys === 0 && remoteKeys > 0;
  }
  // 【任务 C】共享设置基线：本地这份设置「基于哪一版」修改的服务端 ETag。
  // 独立本地存储键，不放进 state.settings——那会跟着共享设置复制到其他宿主，基线就串台了。
  function loadBaselineEtag() { return getStoredItem(App.STORAGE_KEYS.settingsBaseline) || null; }
  function saveBaselineEtag(etag) {
    state.settingsBaselineEtag = etag || null;
    try {
      if (typeof window === 'undefined' || !window.localStorage) return;
      if (etag) window.localStorage.setItem(App.STORAGE_KEYS.settingsBaseline, etag);
      else window.localStorage.removeItem(App.STORAGE_KEYS.settingsBaseline);
    } catch {}
  }
  // 【R4】本地待同步状态元数据：saveSettings 标脏、上传成功或明确采用远端才清除。
  // 独立本地键持久化——关闭重开插件后启动同步据此判断「本地有没有没传上去的修改」，
  // 不再只看时间戳决定覆盖。写入失败要如实提示，不能假装跨重开可恢复。
  function readSyncMeta() {
    try {
      const raw = getStoredItem(App.STORAGE_KEYS.settingsSyncMeta);
      return raw ? JSON.parse(raw) : null;
    } catch { return null; }
  }
  function writeSyncMeta(meta) {
    try {
      if (typeof window === 'undefined' || !window.localStorage) return false;
      window.localStorage.setItem(App.STORAGE_KEYS.settingsSyncMeta, JSON.stringify(meta));
      return true;
    } catch { return false; }
  }
  function markLocalDirty() {
    if (!writeSyncMeta({ dirty: true, dirtyAt: Date.now() })) {
      state.settingsSyncWarning = true;
      if (App.render) App.render();
    }
  }
  function markLocalClean() { writeSyncMeta({ dirty: false, cleanAt: Date.now() }); }
  // 上传前的探测 GET 只用来做安全检查，绝不更新基线——拿「刚看到的最新版本」给旧配置
  // 洗白上传，就等于给静默覆盖换了套包装。只有实际采纳远端、或保存成功，才动基线。
  function settingsEquivalent(a, b) {
    try {
      const norm = x => { const c = normalizedSettings(x); delete c.settingsUpdatedAt; return JSON.stringify(c); };
      return norm(a) === norm(b);
    } catch { return false; }
  }
  function settingsSyncWarn(kind) {
    if (kind === 'conflict') { state.settingsSyncConflict = true; state.settingsSyncWarning = false; }
    else if (kind === 'unsupported') { state.settingsSyncUnsupported = true; }
    else state.settingsSyncWarning = true;
    if (App.render) App.render();
  }
  // 上传串行化：请求在途时用户继续编辑 → 成功后基线照常推进（在途编辑就是基于新版本改的），
  // 但要排一次补跑，把在途期间的增量再传上去，不能把新修改标成「已保存」。
  let settingsUploadInFlight = false;
  let settingsUploadQueued = false;
  async function uploadSharedSettings() {
    if (!App.hasOffice || !App.hasOffice()) return;
    if (settingsUploadInFlight) { settingsUploadQueued = true; return; }
    settingsUploadInFlight = true;
    try {
      const snapshot = JSON.stringify(state.settings);
      const probe = await fetch(localApiUrl('https://localhost:18443/api/settings')).catch(() => null);
      const remoteEtag = probe ? (probe.headers && probe.headers.get('ETag')) || null : null;
      if (probe && !remoteEtag) {
        // 旧服务端：没有版本标记就无法安全上传，停止共享上传并提示更新本地服务；本地功能照常
        settingsSyncWarn('unsupported');
        return;
      }
      if (!probe || !probe.ok) { settingsSyncWarn('network'); return; }
      const remote = await probe.json().catch(() => null);
      if (remote && typeof remote === 'object' && sharedUploadBlockedBy(remote)) {
        console.warn('[settings] 上传已拦截：本地无任何带 key 的服务商，共享层有', providersWithKeys(remote), '个');
        return;
      }
      let baseEtag = state.settingsBaselineEtag;
      if (!baseEtag) {
        if (remote && typeof remote === 'object' && !Object.keys(remote).length) {
          baseEtag = remoteEtag;  // 共享层为空：首次写入走空层初始版本的条件检查，不存在可覆盖的数据
        } else if (remote && settingsEquivalent(remote, state.settings)) {
          // 升级后无基线但本地与远端一致（比较最多忽略 settingsUpdatedAt）：建立基线，无需上传
          saveBaselineEtag(remoteEtag);
          markLocalClean();
          state.settingsSyncConflict = false; state.settingsSyncWarning = false;
          if (App.render) App.render();
          return;
        } else {
          // 本地与远端不一致且缺可信基线：保留本地并提示冲突，不因本地时间戳较新就覆盖远端
          console.warn('[settings] 无可信基线且与共享层不一致，已暂停上传，等待采用共享设置或建立基线');
          settingsSyncWarn('conflict');
          return;
        }
      }
      const resp = await fetch(localApiUrl('https://localhost:18443/api/settings'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'If-Match': baseEtag },
        body: snapshot
      }).catch(() => null);
      if (resp && resp.ok) {
        const data = await resp.json().catch(() => null);
        if (data && data.etag) saveBaselineEtag(data.etag);
        state.settingsSyncConflict = false;
        if (JSON.stringify(state.settings) !== snapshot) {
          settingsUploadQueued = true;  // 在途期间的增量修改稍后基于新版本补传（脏标记保留）
        } else {
          markLocalClean();             // 这份快照确认上传成功：跨重开不再是「待同步」
          if (state.settingsSyncWarning) { state.settingsSyncWarning = false; if (App.render) App.render(); }
        }
      } else if (resp && resp.status === 409) {
        const err = await resp.json().catch(() => ({}));
        if (err && err.code === 'settings_conflict') {
          // 另一应用先写入了：保留本地与远端，不自动重试、不刷新时间戳强推，等用户决定
          console.warn('[settings] 共享设置版本冲突：另一应用已更新，本地修改保留');
          settingsSyncWarn('conflict');
        } else {
          settingsSyncWarn('barren');
        }
      } else {
        settingsSyncWarn(resp ? String(resp.status) : 'network');
      }
    } catch {
      settingsSyncWarn('network');
    } finally {
      settingsUploadInFlight = false;
      if (settingsUploadQueued) {
        settingsUploadQueued = false;
        setTimeout(uploadSharedSettings, 600);
      }
    }
  }
  function saveSettings(options = {}) {
    state.settings.settingsUpdatedAt = Date.now();
    setStoredItem(App.STORAGE_KEYS.settings, JSON.stringify(state.settings));
    // 三应用共享：防抖同步到本地服务（Word/PPT/Excel 各自的 localStorage 无法互通）
    if (syncTimer) clearTimeout(syncTimer);
    // 仅真实 Office 环境上传共享：浏览器直接打开页面属于测试场景，
    // 上传会把测试配置污染到 Word/PPT/Excel
    if (!App.hasOffice || !App.hasOffice()) return;
    markLocalDirty();  // 【R4】本地改动尚未确认上传：持久标脏，启动同步据此保护
    if (options.immediate) uploadSharedSettings();
    else syncTimer = setTimeout(uploadSharedSettings, 600);
  }
  // Office 宿主就绪后再同步。【R4】时间戳只辅助显示，不再决定覆盖：
  // 本地有待同步修改（持久 dirty 标记）且远端已前进 → 保留本地提示冲突；
  // 远端仍等于本地基线 → 继续条件上传（不把每次重开都变成人工确认）；
  // 旧数据无元数据且与远端不等 → 保留本地提示选择；干净本地照常采纳新远端。
  async function syncSharedSettings() {
    if (!App.hasOffice || !App.hasOffice()) return;
    try {
      const response = await fetch(localApiUrl('https://localhost:18443/api/settings'));
      const remoteEtag = (response.headers && response.headers.get('ETag')) || null;
      if (response.ok && !remoteEtag) {
        state.settingsSyncUnsupported = true;  // 新客户端 + 旧服务端：只读，停止共享上传
        if (App.render) App.render();
      }
      const rawRemote = response.ok ? await response.json() : null;
      if (!rawRemote || typeof rawRemote !== 'object' || !Object.keys(rawRemote).length) {
        if (hadStoredSettings) uploadSharedSettings();
        return;
      }
      const remote = normalizedSettings(rawRemote);
      const localTime = Number(state.settings.settingsUpdatedAt) || 0;
      const remoteTime = Number(remote.settingsUpdatedAt) || 0;
      const meta = readSyncMeta();
      const dirty = !!(meta && meta.dirty);
      if (remoteTime > localTime) {
        if (dirty) {
          if (state.settingsBaselineEtag && remoteEtag === state.settingsBaselineEtag) {
            uploadSharedSettings();  // 远端相对本地基线没动：条件上传会成功，不必人工确认
            if (App.render) App.render();
            return;
          }
          // 本地有未确认上传的修改，且远端不是本地基线那版：保留本地并提示冲突
          state.settingsSyncConflict = true;
          console.warn('[settings] 启动同步：本地有未同步修改且共享层已前进，保留本地');
          if (App.render) App.render();
          return;
        }
        if (!meta && hadStoredSettings && !settingsEquivalent(remote, state.settings)) {
          // 旧版本升级：无脏标记也无基线可判，内容与远端不等 → 保留本地，让用户选
          state.settingsSyncConflict = true;
          console.warn('[settings] 启动同步：无法确认本地修改是否已同步，保留本地');
          if (App.render) App.render();
          return;
        }
        state.settings = remote;
        setStoredItem(App.STORAGE_KEYS.settings, JSON.stringify(state.settings));
        if (remoteEtag) saveBaselineEtag(remoteEtag);  // 实际采纳远端才建立基线
      } else if (localTime > remoteTime) {
        uploadSharedSettings();
      } else if (!hadStoredSettings) {
        // 旧版本没有时间戳时：已有本地选择优先；全新宿主才采用共享设置。
        state.settings = remote;
        setStoredItem(App.STORAGE_KEYS.settings, JSON.stringify(state.settings));
        if (remoteEtag) saveBaselineEtag(remoteEtag);
      }
      if (App.render) App.render();
    } catch {
      // 【批次 B】同步失败不再静默：本地设置保留（localStorage 是事实源），输入框上方警示
      state.settingsSyncWarning = true;
      if (App.render) App.render();
    }
  }
  // 【任务 C】冲突恢复入口（设置页「采用共享设置」）：放弃本机未同步的修改，采纳共享层当前
  // 版本并重建基线。后果由 UI 文案明示，用户明确点击后才执行；GET 失败则本地原样保留。
  async function adoptSharedSettings() {
    if (!App.hasOffice || !App.hasOffice()) return;
    try {
      const response = await fetch(localApiUrl('https://localhost:18443/api/settings'));
      const remoteEtag = (response.headers && response.headers.get('ETag')) || null;
      if (!response.ok || !remoteEtag) { settingsSyncWarn('network'); return; }
      const rawRemote = await response.json();
      if (!rawRemote || typeof rawRemote !== 'object') { settingsSyncWarn('network'); return; }
      state.settings = normalizedSettings(rawRemote);
      // 不刷新 settingsUpdatedAt：采纳的是远端版本，不是本机修改；下次真实编辑时 saveSettings 自会盖章
      setStoredItem(App.STORAGE_KEYS.settings, JSON.stringify(state.settings));
      saveBaselineEtag(remoteEtag);
      markLocalClean();  // 明确采用远端：本地未同步修改已被放弃，跨重开恢复干净态
      state.settingsSyncConflict = false;
      state.settingsSyncWarning = false;
      if (App.render) App.render();
    } catch {
      settingsSyncWarn('network');  // 采纳失败：本地设置不动，警示保持
    }
  }
  function loadSessions() {
    // 【W3W5-R1】装载即清理：测试运行的临时会话绝不能从存储回流（旧版本曾在生产
    // 保存动作中把它们连带写盘）。识别=显式 isTestRun 标记，或确定性 id 前缀 testrun-
    //（仅用于兼容清理本功能自己产生过的残留，不按名称模糊删任何业务会话）。
    try {
      const list = JSON.parse(getStoredItem(App.STORAGE_KEYS.sessions) || '[]');
      return (Array.isArray(list) ? list : []).filter(s => !isTestRunSession(s)).map(s => { normalizeRestoredProposalCards(s); return s; });
    } catch { return []; }
  }
  // 【W3W5-R1】测试运行的临时会话在任何持久化边界都被排除（保存、降级、淘汰、装载）。
  function isTestRunSession(s) {
    return !!(s && (s.isTestRun || String(s.id || '').startsWith('testrun-')));
  }
  // WKWebView localStorage 配额约 5MB；会话里带着幻灯片快照等大负载，顶满后写入会静默失败
  //（曾经因此丢过整段对话）。保存失败时逐级降级：压缩旧会话重负载 → 全部压缩 →
  // 淘汰最旧会话（永远保留当前会话）→ 仍失败则置 storageQuotaWarning 明示用户。
  const SESSION_SLIM_LIMIT = 300;
  function slimText(value) {
    return typeof value === 'string' && value.length > SESSION_SLIM_LIMIT ? value.slice(0, SESSION_SLIM_LIMIT) + '…[已压缩]' : value;
  }
  function slimSessionForStorage(session) {
    return Object.assign({}, session, {
      messages: (session.messages || []).map(m => {
        const copy = Object.assign({}, m, { content: slimText(m.content), reasoning: slimText(m.reasoning) });
        delete copy.apiParts; // 与 toolCalls 结果重复的 wire 回放结构，是体积大头；压缩后旧会话退回文本摘要回放
        if (Array.isArray(copy.toolCalls)) {
          copy.toolCalls = copy.toolCalls.map(tc => Object.assign({}, tc, { result: slimText(typeof tc.result === 'string' ? tc.result : JSON.stringify(tc.result)) }));
        }
        return copy;
      })
    });
  }
  // 会话落盘时剥掉消息上的 pendingImages（图片 base64 有 MB 级，写进 localStorage 必顶爆配额——
  // 踩坑 8）。内存对象不动；重开侧边栏后历史注入块仍在但数据已失，旧 attachmentId 会按
  // 「未找到」引导用户重新拖图（可接受边界）。
  // 【35-R1】进行态不落盘：_enriching 只在内存证明活跃准备，落盘即剥除；
  // applying（单卡决策/批量条目）的写入结果在重开后不可知，落盘记为 interrupted——
  // 恢复后据实提示「已中断，先核对文档」，不假装仍在应用，也不标成功/失败/可安全重试。
  function stripPendingImages(key, value) {
    if (key === 'pendingImages') return undefined;
    if (key === '_enriching') return undefined;
    if (key === 'proposal' && value && value.decision === 'applying') {
      return Object.assign({}, value, { decision: 'interrupted' });
    }
    if (key === 'batch' && value && Array.isArray(value.items) && value.items.includes('applying')) {
      return Object.assign({}, value, { items: value.items.map(s => (s === 'applying' ? 'interrupted' : s)) });
    }
    return value;
  }

  // 【35-R1】恢复入口归一化进行态卡片：覆盖新保存逻辑生效前已落盘的旧记录
  //（存储里仍是 applying/_enriching）。持有内存解析器的卡是真实在途任务
  //（切换内存会话 ≠ 重载），一律跳过，不动其写入所有权。
  function normalizeRestoredProposalCards(session) {
    for (const m of (session && session.messages) || []) {
      for (const tc of (m && m.toolCalls) || []) {
        if (!tc || tc.name !== 'propose_edits') continue;
        if (tc._resolveProposal) continue;
        delete tc._enriching;
        if (tc.proposal && tc.proposal.decision === 'applying') tc.proposal.decision = 'interrupted';
        if (tc.batch && Array.isArray(tc.batch.items)) {
          tc.batch.items = tc.batch.items.map(s => (s === 'applying' ? 'interrupted' : s));
        }
      }
    }
  }

  function saveSessions() {
    const key = App.STORAGE_KEYS.sessions;
    // 【W3W5-R1】三层持久化（全量/降级/淘汰）都只写业务会话；内存视图始终保留
    // 运行中的测试会话（降级回写不得把它挤出内存）。生产发送/收尾触发的保存
    // 因此不可能把测试会话或测试消息带进 localStorage。
    const bizAll = () => state.sessions.filter(s => !isTestRunSession(s));
    const serialize = list => JSON.stringify(list, stripPendingImages);
    if (trySetStoredItem(key, serialize(bizAll()))) { state.storageQuotaWarning = false; return; }
    // 降级写入成功后把压缩结果同步回内存：否则内存一直比磁盘胖，之后每次保存都重新降级，
    // 警告条永远不消失（业务会话对象保持原样，不损失保真度；测试会话不在落盘视图内）。
    const adopt = (bizList) => {
      state.sessions = state.sessions.filter(isTestRunSession).concat(bizList.map(s => (s.id === businessCurrentSessionId()
        ? bizAll().find(x => x.id === businessCurrentSessionId()) : s)));
    };
    let list = bizAll().map(s => (s.id === businessCurrentSessionId() ? s : slimSessionForStorage(s)));
    if (trySetStoredItem(key, serialize(list))) { adopt(list); state.storageQuotaWarning = true; return; }
    list = bizAll().map(slimSessionForStorage);
    while (!trySetStoredItem(key, serialize(list)) && list.length > 1) {
      // 淘汰最旧的非当前会话
      for (let i = list.length - 1; i >= 0; i--) {
        if (list[i].id !== businessCurrentSessionId()) { list.splice(i, 1); break; }
      }
    }
    adopt(list);
    state.storageQuotaWarning = true;
  }
  // 【W3W5-R1】测试运行期间的「业务当前会话」：测试会话占据 state.currentSessionId 时，
  // 保真与淘汰保护应指向业务会话而非测试会话。
  function businessCurrentSessionId() {
    if (testRunActive) {
      const st = testRunBusinessSessionId;
      if (st) return st;
    }
    return state.currentSessionId;
  }
  function id() { const c = typeof crypto !== 'undefined' ? crypto : null; return (c && typeof c.randomUUID === 'function') ? c.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`; }
  function now() { return Date.now(); }
  function hasOffice() {
    if (typeof Office === 'undefined') return false;
    const runtimeReady = (typeof Excel !== 'undefined' && !!Excel.run) || (typeof Word !== 'undefined' && !!Word.run) || (typeof PowerPoint !== 'undefined' && !!PowerPoint.run);
    // 同时要求已选定一个可用的宿主提供者，避免 provider 选定失败时 UI 仍呈现为可用 Office 会话。
    return runtimeReady && !!(App.host && App.host.available);
  }
  function escapeHtml(s) { return String(s ?? '').replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch])); }
  function pretty(v) { try { return JSON.stringify(v, null, 2); } catch { return String(v); } }
  function normalizeBaseUrl(url) { return String(url || '').replace(/\/+$/, ''); }
  function chatEndpoint() { const base = normalizeBaseUrl(state.settings.customPrefixUrl || App.DEFAULT_BASE_URL); return base.endsWith('/chat/completions') ? base : `${base}/chat/completions`; }

  function uniqueModelIds(values) {
    return Array.from(new Set((values || []).map(x => String(x || '').trim()).filter(Boolean)));
  }

  function isRetiredModel(provider, model) {
    return provider === 'deepseek' && (model === 'deepseek-chat' || model === 'deepseek-reasoner');
  }

  function extraModelsFor(provider) {
    const favorites = (state.settings.providerModels || {})[provider] || [];
    const recent = (state.settings.recentModels || {})[provider] || [];
    return uniqueModelIds(recent.concat(favorites)).filter(id => !isRetiredModel(provider, id)).map(id => ({ id, name: id, custom: true }));
  }

  function rememberModel(provider, model) {
    const pid = String(provider || '');
    const id = String(model || '').trim();
    if (!pid || !id) return;
    state.settings.recentModels = state.settings.recentModels || {};
    const list = uniqueModelIds([id].concat(state.settings.recentModels[pid] || [])).slice(0, 12);
    state.settings.recentModels[pid] = list;
  }

  function getProviderModels(provider) {
    let models;
    const hidden = ((state.settings.hiddenModels || {})[provider] || []);
    const service = (state.settings.serviceProviders || []).find(item => item.id === provider);
    models = service && Array.isArray(service.models)
      ? service.models.map(model => typeof model === 'string' ? { id: model, name: model } : model).filter(model => model && !hidden.includes(model.id))
      : [];
    models = models.concat(extraModelsFor(provider).filter(m => !hidden.includes(m.id)));
    const current = state.settings.model;
    if (current && !isRetiredModel(provider, current) && !models.some(m => m.id === current)) models.unshift({ id: current, name: current });
    return models;
  }

  const state = {
    locale: loadLocale(),
    theme: getStoredItem(App.STORAGE_KEYS.theme) || (prefersDarkMode() ? 'dark' : 'light'),
    tab: 'chat',
    settings: loadSettings(),
    workbookId: null,
    sessions: loadSessions(),
    currentSessionId: null,
    messages: [],
    isWorking: false,
    stopRequested: false,
    abortController: null,
    error: null,
    toolOutput: '',
    workbookLabel: '',
    sessionMenuOpen: false,
    pendingDeleteSessionId: null,
    modelMenuOpen: false,
    accessMenuOpen: false,
    thinkingMenuOpen: false,
    skillMenuOpen: false,
    activeSkillId: null,
    providerEditing: null,
    providerModelManaging: null,
    workPhase: null,
    attachments: [],
    pendingImages: [],   // 【图片上传】挂起的用户图片（id/name/dataUrl），insert_image 按 attachmentId 取用
    composerDraft: '',
    steerQueue: [],

    recoveredTaskNotice: false,
    storageQuotaWarning: false,
    settingsSyncWarning: false,
    settingsSyncConflict: false,     // 【任务 C】版本冲突：另一应用已更新，本地修改保留待用户决定
    settingsSyncUnsupported: false,  // 【任务 C】旧服务端（GET 无 ETag）：暂停共享上传
    settingsBaselineEtag: loadBaselineEtag(),  // 本机共享设置基线（独立存储键，见上）
    changeLogOpen: false,
    selectionTarget: false,
    selectionSnapshot: null,
    pendingInterruptMessage: null,
    slashItems: null,
    slashRange: null,
    slashIdx: 0,
    plusMenuOpen: false,
    plusSkillsOpen: false,
    skillSearch: '',
    dragging: false
  };

  function activeHostKey() {
    const key = App.host && App.host.hostType;
    return ['word', 'powerpoint', 'excel'].includes(key) ? key : 'browser';
  }

  // 【Q1】测试运行的有效权限模式覆盖（纯内存）：固定回复用例需要确认卡流程，但
  // 不能修改 state.settings——否则任何一次普通 saveSettings 都会把临时模式写进
  // localStorage 和共享同步（Astra 隔离复现确认）。判定点优先读本覆盖。
  let accessModeOverride = null;
  function currentAccessMode() {
    if (accessModeOverride && ['confirm', 'auto', 'plan'].includes(accessModeOverride)) return accessModeOverride;
    const mode = (state.settings.accessModeByHost || {})[activeHostKey()];
    return ['confirm', 'auto', 'plan'].includes(mode) ? mode : 'confirm';
  }
  function setAccessModeOverride(mode) {
    accessModeOverride = ['confirm', 'auto', 'plan'].includes(mode) ? mode : null;
  }
  function clearAccessModeOverride() { accessModeOverride = null; }

  function currentAdvancedAccess() {
    return Boolean((state.settings.advancedAccessByHost || {})[activeHostKey()]);
  }

  function applyHostAccessSettings() {
    const key = activeHostKey();
    state.settings.accessModeByHost = state.settings.accessModeByHost || {};
    state.settings.advancedAccessByHost = state.settings.advancedAccessByHost || {};
    if (!['confirm', 'auto', 'plan'].includes(state.settings.accessModeByHost[key])) state.settings.accessModeByHost[key] = 'confirm';
    if (typeof state.settings.advancedAccessByHost[key] !== 'boolean') state.settings.advancedAccessByHost[key] = false;
    // 保留旧字段只为向后兼容，实际规则始终读取按宿主隔离的新字段。
    state.settings.accessMode = state.settings.accessModeByHost[key];
    return state.settings.accessMode;
  }

  function setAccessMode(mode) {
    const safe = ['confirm', 'auto', 'plan'].includes(mode) ? mode : 'confirm';
    const key = activeHostKey();
    state.settings.accessModeByHost = state.settings.accessModeByHost || {};
    state.settings.accessModeByHost[key] = safe;
    state.settings.accessMode = safe;
  }

  function setAdvancedAccess(enabled) {
    const key = activeHostKey();
    state.settings.advancedAccessByHost = state.settings.advancedAccessByHost || {};
    state.settings.advancedAccessByHost[key] = Boolean(enabled);
  }

  // 【W3W5-隔离】测试运行内存标志（纯内存，随桥注入安装/卸载翻转）：测试回合期间
  // 学习复盘等后台入口据此暂停，防止测试消息进入学习产物或触发无关后台模型请求。
  // 与 accessModeOverride 同款约束：不落盘、不进共享同步、清理按运行作用域。
  // beginTestRun 记录业务当前会话 id：持久化/淘汰的「当前会话保真」在测试期间仍指向
  // 业务会话，不借测试运行改写业务消息或任务状态，也不让测试会话顶替它受保护。
  let testRunActive = false;
  let testRunBusinessSessionId = null;
  function beginTestRun(businessSessionId) {
    testRunActive = true;
    testRunBusinessSessionId = (businessSessionId && String(businessSessionId)) || null;
  }
  function endTestRun() { testRunActive = false; testRunBusinessSessionId = null; }
  function isTestRunActive() { return testRunActive; }

  function activeWorkbookKey() { return state.workbookId || 'browser'; }
  function sessionsForCurrentWorkbook() {
    const key = activeWorkbookKey();
    return state.sessions.filter(s => (s.workbookId || 'browser') === key);
  }
  function makeSession(name) {
    return { id: id(), workbookId: activeWorkbookKey(), name: name || (state.locale === 'zh' ? '新对话' : 'NEW CHAT'), messages: [], createdAt: now(), updatedAt: now() };
  }
  function ensureSession() {
    const available = sessionsForCurrentWorkbook();
    if (state.currentSessionId && available.some(s => s.id === state.currentSessionId)) return;
    let session = available[0];
    if (!session) { session = makeSession(); state.sessions.unshift(session); saveSessions(); }
    state.currentSessionId = session.id;
    state.messages = session.messages || [];
    state.pptTaskScope = session.pptTaskScope || null;
    App.markSessionRecoveredIfNeeded(session);
  }
  // 任务中断恢复：会话带着 taskActive 标记说明上次任务死在半路（侧边栏被关/崩溃）。
  // 清掉标记并在界面上提示一次「进度已保留，说继续即可接力」。
  function markSessionRecoveredIfNeeded(session) {
    if (!session) return;
    // UI 切换会话时也走这里，范围账本必须随会话切换，不能留在上一个 PPT 会话里。
    state.pptTaskScope = session.pptTaskScope || null;
    // 【35-R1】恢复入口同时归一化进行态卡片（applying→interrupted、剥除 _enriching），
    // 覆盖未经 loadSessions 的恢复路径与存量旧记录；在途（有解析器）的卡不受影响
    normalizeRestoredProposalCards(session);
    if (session.taskActive) {
      session.taskActive = false;
      state.recoveredTaskNotice = true;
      saveSessions();
    }
  }
  function persistCurrentSession() {
    const s = state.sessions.find(x => x.id === state.currentSessionId);
    if (!s) return;
    s.messages = state.messages;
    s.taskActive = !!state.isWorking;
    // 任务范围账本不依赖 apiParts：工具历史可能为节省空间被折叠，但整稿覆盖事实必须可恢复。
    s.pptTaskScope = state.pptTaskScope || null;
    s.updatedAt = now();
    const firstUser = state.messages.find(m => m.role === 'user');
    if (firstUser && (!s.name || s.name === '新对话' || s.name === 'NEW CHAT')) {
      // 会话名用用户实际输入（displayContent），不带 [当前选区]/附件等注入前缀
      const text = String(firstUser.displayContent || firstUser.content || '').trim();
      s.name = text.length > 40 ? text.slice(0, 37) + '...' : text;
    }
    state.sessions.sort((a, b) => b.updatedAt - a.updatedAt);
    saveSessions();
  }

  App.allSkills = function () {
    return (App.SKILLS || []).concat(state.settings.customSkills || []);
  };
  App.state = state;
  // 旧配置自愈：当前服务失效时改用现有服务的第一个。
  (function normalizeProvider() {
    const p = state.settings.provider;
    const services = state.settings.serviceProviders || [];
    const valid = services.some(service => service.id === p);
    if (!valid) {
      const fallback = services[0];
      state.settings.provider = fallback ? fallback.id : '';
      state.settings.customPrefixUrl = fallback ? fallback.baseUrl : '';
      state.settings.apiKey = fallback ? (fallback.apiKey || '') : '';
      const selected = fallback ? ((state.settings.providerModels || {})[fallback.id] || []) : [];
      state.settings.model = selected[0] || '';
    }
    if (state.settings.customPrefixUrl === 'https://open.bigmodel.cn/api/paas/v4') {
      state.settings.customPrefixUrl = 'https://open.bigmodel.cn/api/coding/paas/v4';
      const zhipu = services.find(service => service.id === 'zhipu');
      if (zhipu) zhipu.baseUrl = state.settings.customPrefixUrl;
    }
  })();
  App.getStoredItem = getStoredItem;
  App.setStoredItem = setStoredItem;
  App.prefersDarkMode = prefersDarkMode;
  App.getNavigatorLanguage = getNavigatorLanguage;
  App.t = t;
  App.loadLocale = loadLocale;
  App.loadSettings = loadSettings;
  App.saveSettings = saveSettings;
  App.syncSharedSettings = syncSharedSettings;
  App.uploadSharedSettings = uploadSharedSettings; // 供 settings-guard 测试直调（内含贫瘠覆盖拦截闸）
  App.adoptSharedSettings = adoptSharedSettings;   // 【任务 C】冲突恢复：采用共享设置（设置页入口）
  App.rememberModel = rememberModel;
  App.isRetiredModel = isRetiredModel;
  App.loadSessions = loadSessions;
  App.saveSessions = saveSessions;
  App.id = id;
  App.now = now;
  App.hasOffice = hasOffice;
  App.escapeHtml = escapeHtml;
  App.pretty = pretty;
  App.normalizeBaseUrl = normalizeBaseUrl;
  App.chatEndpoint = chatEndpoint;
  App.getProviderModels = getProviderModels;
  App.activeWorkbookKey = activeWorkbookKey;
  App.sessionsForCurrentWorkbook = sessionsForCurrentWorkbook;
  App.makeSession = makeSession;
  App.ensureSession = ensureSession;
  App.persistCurrentSession = persistCurrentSession;
  App.markSessionRecoveredIfNeeded = markSessionRecoveredIfNeeded;
  App.localApiUrl = localApiUrl;
  App.activeHostKey = activeHostKey;
  App.currentAccessMode = currentAccessMode;
  App.setAccessModeOverride = setAccessModeOverride;     // 【Q1】测试桥专用：内存覆盖，不触持久设置
  App.clearAccessModeOverride = clearAccessModeOverride;
  App.currentAdvancedAccess = currentAdvancedAccess;
  App.applyHostAccessSettings = applyHostAccessSettings;
  App.setAccessMode = setAccessMode;
  App.setAdvancedAccess = setAdvancedAccess;
  App.beginTestRun = beginTestRun;                     // 【W3W5-隔离】测试桥专用：内存标志
  App.endTestRun = endTestRun;
  App.isTestRunActive = isTestRunActive;
})();
