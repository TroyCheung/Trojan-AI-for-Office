(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ── 学习闭环执行器（v137）：计数器触发 + 三窗口 + 欠账补跑 ──
  // 借鉴 NousResearch/hermes-agent（MIT）的 nudge+后台复盘理念，改造为纯前端形态：
  // 回复交付完成/确认卡等待/启动补账三个窗口里，静默发一次廉价复盘调用，从对话中
  // 提取值得长期记住的偏好与经验，经 App.memory.writeMemoryTool 落库。
  // 铁律：复盘只读写记忆库，永不触碰文档；用户发新消息立即让位（abort），欠账保留。
  // 触发是周期性的，被打断的复盘无需弥补性紧迫——欠账只是「找机会再跑一次」的提示。

  const TURN_INTERVAL = 3;          // 每 N 个成功回合尝试一次复盘
  const MAX_CONTEXT_MESSAGES = 24;  // 喂给复盘的对话条数上限（短平快，给窗口留命中机会）
  const MAX_OPS = 5;                // 单次复盘最多应用的记忆操作数（防爆炸）
  const MAX_DEBTS = 3;              // 欠账队列上限（FIFO）

  let turnCount = 0;
  let running = false;
  let abortController = null;

  function debtKey() { return (App.STORAGE_KEYS && App.STORAGE_KEYS.reviewDebt) || 'office-ai-trial-review-debt-v1'; }
  // state.js 的封装通常已就位；直接回退 window.localStorage，测试与极端加载顺序下也不炸
  function debtGet() {
    if (typeof App.getStoredItem === 'function') return App.getStoredItem(debtKey());
    return (typeof window !== 'undefined' && window.localStorage) ? window.localStorage.getItem(debtKey()) : null;
  }
  function debtSet(value) {
    if (typeof App.setStoredItem === 'function') { App.setStoredItem(debtKey(), value); return; }
    if (typeof window !== 'undefined' && window.localStorage) window.localStorage.setItem(debtKey(), value);
  }
  function loadDebts() {
    try { return JSON.parse(debtGet() || '[]') || []; } catch { return []; }
  }
  function saveDebts(list) { debtSet(JSON.stringify(list.slice(0, MAX_DEBTS))); }

  function isRunning() { return running; }

  function hasDebt() { return loadDebts().length > 0; }

  // ── 三个窗口的入口 ──

  // 【W3W5-隔离】测试运行期间复盘整体静默：测试消息/拒绝理由/预设回复不得进入学习
  // 产物，也不得产生测试之外的后台模型请求。入口级拦截（三个窗口全覆盖）。
  function testRunSuspended() {
    try { return !!(App.isTestRunActive && App.isTestRunActive()); } catch (e) { return false; }
  }

  // 回合成功结束（回复已交付给用户）：计数 + 找机会跑。返回 promise 供测试 await；生产调用方可不 await
  function onTurnEnd() {
    if (testRunSuspended()) return Promise.resolve();
    turnCount += 1;
    if (turnCount % TURN_INTERVAL !== 0 && !hasDebt()) return Promise.resolve();
    const debt = hasDebt() ? loadDebts()[0] : null;
    return run(debt || { current: true });
  }

  // 确认卡/拍板卡等待用户决策：结构性空隙，AI 反正在等
  function onCardWait() {
    if (testRunSuspended()) return Promise.resolve();
    if (running) return Promise.resolve();
    const debt = hasDebt() ? loadDebts()[0] : null;
    if ((turnCount > 0 && turnCount % TURN_INTERVAL === 0) || debt) return run(debt || { current: true }, { waitingCard: true });
    return Promise.resolve();
  }

  // 插件启动：消化欠账
  function onStartup() {
    if (testRunSuspended()) return Promise.resolve();
    if (!hasDebt()) return Promise.resolve();
    return run(loadDebts()[0]);
  }

  // 用户发新消息：立即让位。abort 一个 fetch 无副作用（复盘产物是一次性写入，没写成即没做）
  function abort() {
    if (abortController && typeof abortController.abort === 'function') abortController.abort();
  }

  function guard(opts) {
    const state = App.state || {};
    if (running) return false;
    if (!App.memory || !App.memory.enabled()) return false;
    // AI 正在产出/执行工具时不跑；唯一例外是确认卡等待期（waitingCard 由 api.js 在 await 卡片前置位）
    if (state.isWorking && !(opts && opts.waitingCard)) return false;
    return true;
  }

  // ── 对话摘录组装 ──

  function collectMessages(debt) {
    if (debt && debt.sessionId && !(debt.current)) {
      const sessions = App.loadSessions ? App.loadSessions() : [];
      const session = sessions.find(s => s.id === debt.sessionId);
      if (!session) return null; // 会话已被删，放弃这笔欠账
      return session.messages || [];
    }
    return (App.state && App.state.messages) || [];
  }

  function transcript(messages) {
    const lines = [];
    for (const m of messages.slice(-MAX_CONTEXT_MESSAGES)) {
      if (m.role === 'user') {
        const text = String(m.displayContent || m.content || '').replace(/\s+/g, ' ').trim();
        if (text) lines.push(`[User] ${text.slice(0, 300)}`);
      } else if (m.role === 'assistant') {
        const text = String(m.content || '').replace(/\s+/g, ' ').trim();
        const tools = (m.toolCalls || []).map(tc => tc.name);
        if (text || tools.length) {
          lines.push(`[AI] ${text.slice(0, 240)}${tools.length ? `（调用工具：${tools.join('、')}）` : ''}`);
        }
      }
    }
    return lines.join('\n') || '（本轮对话极短）';
  }

  // ── 复盘提示词（负面清单与行动优先级借鉴 Hermes，MIT）──

  function buildPrompt(messages) {
    const stats = App.memory.stats();
    const list = stats.count
      ? stats.topics.map(t => `- ${t.key} · ${t.title}${t.summary ? ' · ' + t.summary : ''}（${t.chars} 字符）`).join('\n')
      : '（记忆库为空）';
    return [
      '你在为一位用户的 Office AI 助手做学习复盘。以下是刚结束的一段对话摘录和现有记忆清单。',
      '',
      '## 对话摘录',
      transcript(messages),
      '',
      '## 现有记忆',
      list,
      '',
      '## 任务',
      '从这段对话中提取值得跨会话记住的用户偏好、工作规则与经验，输出 JSON 数组（可为空）。每项格式：',
      '{"action":"create|append|update","topic":"snake_case_key","title":"简短标题","summary":"一句话","content":"规则正文"}',
      '',
      '## 判断标准（按优先级）',
      '1. 先看现有记忆是否需要修订：用户纠正了已有做法、某条记忆过时或缺步骤 → 用 append/update 修订对应主题（优先于新建）。',
      '2. 用户明确表达的偏好、反复出现的工作模式 → 沉淀。',
      '3. 绝不入库：宿主或环境问题（如某 Office 版本缺 API）；对工具/功能的否定断言（会固化成日后的拒绝理由）；未解决的失败不得写成可靠做法；一次性任务细节；文档具体内容；隐私与密钥。',
      '4. 条目写成陈述事实而非对助手的命令（「用户偏好简洁回复」正确，「总是要简洁回复」错误）。时间敏感内容标注日期。',
      '5. action=create 时 topic 必须是新的 snake_case 键；append/update 必须用清单里已有的 key。最多输出 ' + MAX_OPS + ' 项。',
      '6. 只输出 JSON 数组，不要输出任何其他文字。确实没有值得记的就输出 []。'
    ].join('\n');
  }

  function parseOps(text) {
    const raw = String(text || '');
    const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/);
    const candidate = fenced ? fenced[1] : (raw.match(/\[[\s\S]*\]/) || [])[0] || '';
    let arr;
    try { arr = JSON.parse(candidate.trim()); } catch { return null; }
    if (!Array.isArray(arr)) return null;
    return arr.filter(op => op && ['create', 'append', 'update', 'delete'].includes(op.action) && op.topic).slice(0, MAX_OPS);
  }

  // ── 模型调用（非流式，协议分派）──

  async function callModel(userPrompt, signal) {
    const s = App.state.settings;
    // A7（v146）：复盘可指定降级模型（settings.reviewModel，设置卡选填），空则用主模型
    const model = String(s.reviewModel || '').trim() || s.model;
    const isA = typeof App.currentProtocol === 'function' && App.currentProtocol() === 'anthropic' && App.anthropicProtocol;
    const url = isA ? App.anthropicProtocol.endpoint(s.customPrefixUrl) : App.chatEndpoint();
    const headers = isA
      ? App.anthropicProtocol.headers(s.apiKey)
      : { 'Content-Type': 'application/json', 'Authorization': `Bearer ${s.apiKey}` };
    const messages = [{ role: 'user', content: userPrompt }];
    const body = isA
      ? App.anthropicProtocol.outbound({ model, messages, stream: false })
      : { model, messages, stream: false, temperature: 0.2 };
    const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body), signal });
    if (!res.ok) throw new Error(`API ${res.status}: ${(await res.text().catch(() => '')).slice(0, 200)}`);
    const json = await res.json();
    if (isA) {
      return (Array.isArray(json.content) ? json.content : [])
        .filter(b => b.type === 'text').map(b => b.text).join('');
    }
    return String((json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || '');
  }

  // ── 主流程 ──

  async function run(debt, opts) {
    if (!guard(opts)) return;
    running = true;
    abortController = (typeof AbortController !== 'undefined') ? new AbortController() : null;
    const state = App.state;
    if (state) state.reviewRunning = true;
    if (App.render) App.render();
    try {
      const messages = collectMessages(debt);
      if (!messages || !messages.length) {
        if (debt && debt.sessionId) saveDebts(loadDebts().filter(d => d.sessionId !== debt.sessionId));
        return;
      }
      const text = await callModel(buildPrompt(messages), abortController && abortController.signal);
      const ops = parseOps(text);
      if (ops == null) {
        // 解析失败：静默放弃；欠账补跑场景同样弃账，防止每次窗口都对着同一笔账永久重试
        if (debt && debt.sessionId) saveDebts(loadDebts().filter(d => d.sessionId !== debt.sessionId));
        return;
      }
      let applied = 0;
      for (const op of ops) {
        try {
          const r = await App.memory.writeMemoryTool(op);
          if (r && r.success !== false) applied++; // 语义失败（撞 key/主题不存在）是返回值不是异常
        } catch { /* 单条失败跳过 */ }
      }
      // 成功落库：清掉对应的欠账
      if (debt && debt.sessionId) saveDebts(loadDebts().filter(d => d.sessionId !== debt.sessionId));
      if (applied > 0) {
        const msg = (typeof App.t === 'function' ? App.t('reviewToast') : '') || `复盘完成，沉淀了 {n} 条记忆（可在设置中查看）`;
        if (App.toast) App.toast(msg.replace('{n}', applied));
      }
    } catch (e) {
      if (e && (e.name === 'AbortError' || App.state && App.state.stopRequested)) {
        // 用户发新消息让位：当前会话的复盘记欠账，找机会再来
        const sid = App.state && App.state.currentSessionId;
        if (sid) {
          const debts = loadDebts().filter(d => d.sessionId !== sid);
          debts.unshift({ sessionId: sid, workbookId: (App.activeWorkbookKey && App.activeWorkbookKey()) || 'browser', since: Date.now() });
          saveDebts(debts);
        }
      }
      // 其他失败（网络/接口）：静默放弃，不欠账防死循环；留一行警告供排查
      if (!(e && e.name === 'AbortError')) console.warn('[review] 复盘失败：', e && e.message || e);
    } finally {
      running = false;
      abortController = null;
      if (state) state.reviewRunning = false;
      if (App.render) App.render();
    }
  }

  App.review = { onTurnEnd, onCardWait, onStartup, abort, isRunning, hasDebt, buildPrompt, parseOps, transcript, TURN_INTERVAL, MAX_OPS };
})();
