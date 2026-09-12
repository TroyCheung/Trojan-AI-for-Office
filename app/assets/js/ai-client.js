(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ── AI 网络客户端（v140，Pi 式 R1 拆分）：请求构造 + fetch/SSE + 协议适配 + 重试韧性 ──
  // 边界：loop 只通过 callChatCompletions(messages, options) 使用本文件；
  // 策略（权限/意图分类/技能注入）经 options.hooks 注入，本文件不感知业务。
  // 错误语义：失败 throw（带 retryableRequest 标记的网络错误可安全原地重试）。

  // 被剥离的历史 reasoning_content 的备份字段（非枚举，不会被 JSON.stringify 发出）
  App.REASONING_BACKUP = '_reasoningEchoBackup';

  function makeAbortError() {
    const e = new Error('Request aborted');
    e.name = 'AbortError';
    return e;
  }

  function messagesContainImages(messages) {
    return (Array.isArray(messages) ? messages : []).some(message => Array.isArray(message.content)
      && message.content.some(part => part && (part.type === 'image_url' || part.type === 'input_image')));
  }

  function textOnlyMessages(messages) {
    return (Array.isArray(messages) ? messages : []).map(message => {
      if (!Array.isArray(message.content)) return message;
      const text = message.content
        .filter(part => part && part.type === 'text')
        .map(part => String(part.text || ''))
        .filter(Boolean)
        .join('\n');
      return Object.assign({}, message, {
        content: `${text}\n当前模型接口不接受图片输入，本次请依据 get_slide 的 layoutAudit、形状坐标和格式快照继续。`
      });
    });
  }

  async function apiErrorMessage(response) {
    let message = await response.text().catch(() => '');
    try {
      const parsed = JSON.parse(message);
      message = parsed.error?.message || parsed.message || message;
    } catch {}
    return message || response.statusText;
  }

  const visionUnsupportedModels = new Set();
  // 记录「跨轮历史也必须回传 reasoning_content」的接口：命中过一次 400 后不再剥离，
  // 避免每轮都先失败一次再重试。key 与 currentVisionKey 相同（接口地址 + 模型）。
  const reasoningEchoRequiredModels = new Set();

  function currentVisionKey(state) {
    return `${String(state.settings.customPrefixUrl || '')}|${String(state.settings.model || '')}`;
  }

  // 当前服务商的协议（v136）：'openai'（缺省）| 'anthropic'。镜像值 providerProtocol 由设置卡
  // 切换/保存服务时同步；旧配置无该字段一律 openai，既有行为零变化。
  function currentProtocol() {
    const s = (App.state && App.state.settings) || {};
    const svc = (s.serviceProviders || []).find(p => p.id === s.provider);
    if (svc && svc.protocol) return svc.protocol === 'anthropic' ? 'anthropic' : 'openai';
    return s.providerProtocol === 'anthropic' ? 'anthropic' : 'openai';
  }
  async function callChatCompletions(messages, options = {}) {
    // 策略 hooks（Pi 式边界，v140）：由调用方（agent loop）注入；网络层不感知权限/意图/技能注入
    const { requestAccessMode, toolsForAccessMode, toolsForOfficeTask, filterCapabilityGatedTools, advancedAccessEnabled, effectiveOfficeThinkingLevel } = options.hooks || {};
    const state = App.state;
    const tmp = Number(state.settings.temperature);
    const temperature = Number.isFinite(tmp) ? tmp : 0.2;
    const mode = requestAccessMode();
    const taskTools = toolsForOfficeTask(App.host.toolDefinitions, options.taskProfile);
    // 历史账本不是当前任务权限；仅向本轮活跃的 PPT 整稿任务提供只读设计决定入口。
    if (App.host.hostType === 'powerpoint' && state.activePptTaskScope && App.pptTaskScope?.toolDefinition) {
      taskTools.push(App.pptTaskScope.toolDefinition);
    }
    const gatedTools = filterCapabilityGatedTools(taskTools);
    const tools = toolsForAccessMode(gatedTools, mode, advancedAccessEnabled());
    const body = { model: state.settings.model, messages, temperature, stream: true };
    const visionKey = currentVisionKey(state);
    if (visionUnsupportedModels.has(visionKey) && messagesContainImages(body.messages)) {
      body.messages = textOnlyMessages(body.messages);
    }
    // 仅在有工具时才声明 tools/tool_choice；空数组配 tool_choice:'auto' 会被部分 OpenAI 兼容后端拒绝。
    const toolList = Array.isArray(tools) ? tools.slice() : [];
    if (state.settings.searchEnabled && state.settings.tavilyKey) {
      toolList.push({ type: 'function', function: { name: 'web_search', description: 'Search the web for up-to-date information. Use when the question needs facts beyond the document (news, prices, latest data, unfamiliar entities). When the user asks to FIND an image online and insert it, search with a query describing the picture — the response includes direct image links you can pass straight to insert_image(url).', parameters: { type: 'object', properties: { query: { type: 'string', description: 'search query' }, max_results: { type: 'number', description: 'max results, default 5' } }, required: ['query'] } } });
    }
    // 读用户电脑上的文件（本地服务器提供，绕过浏览器沙箱）。始终可用，与宿主和任务分类无关。
    // read_local_file 之后是 load_skill / get_current_time（宿主无关内置工具）
    toolList.push({ type: 'function', function: { name: 'read_local_file', description: 'READ. Read a file from the user\'s computer by path when the user references a local file instead of pasting content (e.g. ~/Desktop/方案.pptx, 微信收到的文件路径). Supports PDF/DOCX/PPTX/TXT/MD/CSV/JSON. Any folder in the user\'s home directory works (Desktop, Documents, Downloads, work folders, WeChat files); hidden paths and ~/Library are off-limits.', parameters: { type: 'object', properties: { path: { type: 'string', description: 'Absolute path or ~-prefixed path' }, explanation: { type: 'string' } }, required: ['path'] } } });
    // load_skill：Excel 方法论技能懒加载（批 5）——任务路由只注入一句路由提示，
    // 模型匹配场景时拉正文，避免六篇方法论全量占每轮 token（取舍学官方 GLM in Excel）
    // 模型匹配场景时拉正文，避免六篇方法论全量占每轮 token（取舍学官方 GLM in Excel）
    if (App.allSkills && App.allSkills().length) {
      toolList.push({ type: 'function', function: { name: 'load_skill', description: 'READ. Load a skill imported by this user. No skills are preinstalled. Use only a skill actually present in the user skill list; never invent a skill name.', parameters: { type: 'object', properties: { name: { type: 'string', description: 'Exact name or id of an imported skill' } }, required: ['name'] } } });
    }
    if (App.pptLayoutSkills && App.host && App.host.hostType === 'powerpoint') {
      toolList.push({ type: 'function', function: { name: 'load_skill', description: 'READ. Load a skill imported by this user. No skills are preinstalled. Use only a skill actually present in the user skill list; never invent a skill name.', parameters: { type: 'object', properties: { name: { type: 'string', description: 'Exact name or id of an imported skill' } }, required: ['name'] } } });
    }
    toolList.push({ type: 'function', function: { name: 'extract_images', description: 'Extract direct image URLs from a web page. Use after web_search when the results are article pages rather than image links: pass the article URL here, get its embedded image links, then insert with insert_image(url). Only usable in a real Office host.', parameters: { type: 'object', properties: { url: { type: 'string', description: 'Absolute http(s) page URL.' }, explanation: { type: 'string' } }, required: ['url'] } } });
    toolList.push({ type: 'function', function: { name: 'get_current_time', description: 'Get the current local date and time from the user\'s machine. Use for any time/date question (current time, today\'s date, day of week) — do NOT use web_search for this. When writing time into cells prefer =NOW()/=TODAY() formulas.', parameters: { type: 'object', properties: { locale: { enum: ['zh', 'en'], description: 'Output locale. Default zh' } } } } });
    // render：结构化呈现工具（宿主无关，按宿主裁剪可用的导出类型）。所有访问模式下都可用，
    // 因为它只渲染界面卡片，不修改文档；真正落回文档的动作由用户点击按钮触发。
    if (App.renderBlockTool) toolList.push(App.renderBlockTool(App.host && App.host.hostType));
    if (toolList.length && !options.noTools) { body.tools = toolList; body.tool_choice = 'auto'; }
    // thinking 参数各家方言不同（智谱要 thinking 对象，DeepSeek 的 effort 无 medium），统一由 dialects.js 翻译
    const dialect = App.currentDialect();
    const thinkingLevel = effectiveOfficeThinkingLevel(state.settings.thinking, options.taskProfile, options.agentStep);
    dialect.applyThinking(body, thinkingLevel, state.settings.model);
    // ── 网络韧性（借鉴 GenOffice agent-core 的抗造机制）──
    // 看门狗：fetch 连接 60 秒、流读取空闲 120 秒。内部 AbortController 与用户停止（options.signal）
    // 共用中止通道，用 timeoutError 区分触发来源：用户停止保持 AbortError 语义，超时报明确中文错误。
    const CONNECT_TIMEOUT_MS = 60000;
    const READ_IDLE_TIMEOUT_MS = 120000;
    // 空流/首字节前网络失败的原地退避重试（最多 2 次）；已收到任何内容后断流不重试，避免内容重复。
    const NETWORK_RETRY_DELAYS = [1000, 3000];
    const userAborted = () => Boolean(state.stopRequested) || Boolean(options.signal && options.signal.aborted);
    // 协议分派（v136）：anthropic 时请求出口+SSE 入口双向转换；openai 全部原样
    const useAnthropic = currentProtocol() === 'anthropic' && !!App.anthropicProtocol;

    async function attemptRequest() {
      const internal = typeof AbortController !== 'undefined' ? new AbortController() : null;
      const timersAvailable = typeof setTimeout === 'function' && typeof clearTimeout === 'function';
      let timeoutError = null;
      let timer = null;
      const arm = (ms, message) => {
        if (!internal || !timersAvailable) return;
        clearTimeout(timer);
        timer = setTimeout(() => {
          timeoutError = new Error(message);
          timeoutError.name = 'TimeoutError';
          internal.abort();
        }, ms);
      };
      const disarm = () => { if (timer) { clearTimeout(timer); timer = null; } };
      const userSignal = options.signal;
      const onUserAbort = () => { if (internal) internal.abort(); };
      if (userSignal && typeof userSignal.addEventListener === 'function') {
        if (userSignal.aborted) onUserAbort();
        else userSignal.addEventListener('abort', onUserAbort, { once: true });
      }
      const translateAbort = (e) => {
        if (timeoutError) return timeoutError;
        if (e && e.name === 'AbortError') return makeAbortError();
        return e;
      };
      const request = payload => {
        arm(CONNECT_TIMEOUT_MS, '连接超时（60 秒没有响应），请检查网络或接口地址后重试。');
        // 协议分派（v136）：anthropic 走 /messages + x-api-key + 请求体转换；openai 原样。
        // payload 始终是 OpenAI 形状，转换在适配器内完成（工具回传也走同一出口，loop 不感知）。
        const url = useAnthropic
          ? App.anthropicProtocol.endpoint(state.settings.customPrefixUrl)
          : App.chatEndpoint();
        const hdrs = useAnthropic
          ? App.anthropicProtocol.headers(state.settings.apiKey)
          : { 'Content-Type': 'application/json', 'Authorization': `Bearer ${state.settings.apiKey}` };
        const wire = useAnthropic ? App.anthropicProtocol.outbound(payload) : payload;
        // 代理分派（v142）：taskpane 是浏览器环境，上游不带 CORS 头时直连必然
        // "Load failed"（Anthropic 官方端点、多数中转）。本地 server.py 提供
        // /api/chat 流式代理（与获取模型同机制），token 鉴权与 settings 同级。
        // anthropic 协议直接走代理；openai 先直连，网络层失败再自动落代理。
        const proxyAvailable = typeof App.localApiUrl === 'function'
          && !String(window && window.__LOCAL_API_TOKEN__ || '').startsWith('__');
        const viaProxy = (useAnthropic && proxyAvailable) || Boolean(forceProxy && proxyAvailable);
        const target = viaProxy
          ? App.localApiUrl('https://localhost:18443/api/chat')
          : url;
        const wireHeaders = viaProxy
          ? { 'Content-Type': 'application/json' }
          : hdrs;
        const wireBody = viaProxy
          ? JSON.stringify({ url, headers: hdrs, body: wire })
          : JSON.stringify(wire);
        return fetch(target, {
          method: 'POST',
          headers: wireHeaders,
          body: wireBody,
          signal: internal ? internal.signal : options.signal
        }).then(response => { disarm(); return response; }, e => {
          disarm();
          const err = translateAbort(e);
          if (err === e) {
            err.retryableRequest = true;
            // 直连网络失败且本地代理可用：标记 fallbackProxy，重试循环改走代理
            if (!viaProxy && proxyAvailable && (e instanceof TypeError)) err.fallbackProxy = true;
          }
          throw err;
        });
      };
      try {
        let res = await request(body);
        if (!res.ok) {
          // 响应体只能读一次，先统一取出错误信息
          let errorStatus = res.status;
          let errorMessage = await apiErrorMessage(res);
          // 部分接口校验历史中带 tool_calls 的 assistant 消息必须回传 reasoning_content。
          // 本轮回放时已剥离历史推理内容，若因此被拒，恢复原样重试一次并记住该接口。
          if (errorStatus === 400 && !useAnthropic && /reasoning/i.test(errorMessage) && body.messages.some(m => m && m[App.REASONING_BACKUP])) {
            const restored = body.messages.map(m => (m && m[App.REASONING_BACKUP])
              ? Object.assign({}, m, { reasoning_content: m[App.REASONING_BACKUP] }) : m);
            res = await request(Object.assign({}, body, { messages: restored }));
            if (res.ok) {
              reasoningEchoRequiredModels.add(currentVisionKey(state));
            } else {
              errorStatus = res.status;
              errorMessage = await apiErrorMessage(res);
            }
          }
          if (!res.ok) {
            const canRetryWithoutImages = messagesContainImages(body.messages) && [400, 413, 415, 422].includes(errorStatus);
            if (!canRetryWithoutImages) throw new Error(`API ${errorStatus}: ${errorMessage}`);

            const fallbackBody = Object.assign({}, body, { messages: textOnlyMessages(body.messages) });
            res = await request(fallbackBody);
            if (!res.ok) {
              const fallbackMessage = await apiErrorMessage(res);
              throw new Error(`API ${res.status}: ${fallbackMessage}`);
            }
            visionUnsupportedModels.add(visionKey);
          }
        }

        const contentType = res.headers && res.headers.get ? (res.headers.get('content-type') || '') : '';
        if (!res.body || typeof res.body.getReader !== 'function' || contentType.includes('application/json')) {
          const json = await res.json();
          const jsonMsg = json.choices?.[0]?.message || {};
          const content = jsonMsg.content || '';
          if (content && options.onContent) options.onContent(content);
          // 空响应（零内容零工具调用零推理）视为可重试的中转/上游故障
          const empty = !String(content).trim() && !String(jsonMsg.reasoning_content || '').trim()
            && !(Array.isArray(jsonMsg.tool_calls) && jsonMsg.tool_calls.length);
          return { result: json, empty };
        }

        const reader = res.body.getReader();
        const decoder = new TextDecoder();
        const toolCalls = [];
        let content = '';
        let reasoning = '';
        let usage = null;
        let finishReason = '';
        let streamError = '';
        let buffer = '';
        // 收到过任何内容（正文/推理/工具调用片段）后断流不再重试，避免重复拼接
        let receivedAny = false;
        // Anthropic SSE 翻译器（v136）：openai 协议为 null，事件原样走 OpenAI 解析
        const sseTranslator = useAnthropic ? App.anthropicProtocol.createSseTranslator() : null;

        const processLine = (line) => {
          const trimmed = line.trim();
          if (!trimmed || trimmed.startsWith(':')) return;
          if (!trimmed.startsWith('data:')) return;
          const data = trimmed.slice(5).trim();
          if (!data || data === '[DONE]') return;
          let chunk;
          try { chunk = JSON.parse(data); } catch { return; }
          if (sseTranslator) {
            const converted = sseTranslator(data);
            if (!converted) return;
            chunk = converted;
          }
          // 中转网关常见故障：扣费后上游失败，把 {"error":...} 推进 SSE 再关流。拦截并报出来，不当空回复
          if (chunk && chunk.error) { streamError = String(chunk.error.message || JSON.stringify(chunk.error)); return; }
          if (chunk.usage) usage = chunk.usage;
          const choice = chunk.choices && chunk.choices[0];
          if (choice && choice.finish_reason) finishReason = String(choice.finish_reason);
          const delta = choice && choice.delta ? choice.delta : {};
          const rDelta = dialect.readReasoningDelta(delta);
          if (rDelta) {
            receivedAny = true;
            reasoning += rDelta;
            if (options.onReasoning) options.onReasoning(rDelta, reasoning);
          }
          if (typeof delta.content === 'string') {
            receivedAny = true;
            content += delta.content;
            if (options.onContent) options.onContent(delta.content, content);
          }
          if (Array.isArray(delta.tool_calls)) { receivedAny = true; mergeToolCallDeltas(toolCalls, delta.tool_calls); }
        };

        while (true) {
          arm(READ_IDLE_TIMEOUT_MS, '响应超时（120 秒没有新数据），请检查网络后重试。');
          let readChunk;
          try { readChunk = await reader.read(); } catch (e) {
            disarm();
            const err = translateAbort(e);
            // 还没收到任何内容的断流可以重试；已有部分内容的断流直接抛错
            if (err === e && !receivedAny) err.retryableRequest = true;
            throw err;
          }
          disarm();
          const { value, done } = readChunk;
          buffer += decoder.decode(value || new Uint8Array(), { stream: !done });
          const lines = buffer.split(/\r?\n/);
          buffer = lines.pop() || '';
          for (const line of lines) processLine(line);
          if (done) break;
        }
        if (buffer.trim()) processLine(buffer);
        if (streamError) throw new Error(`上游接口在流式传输中报错：${streamError}`);

        const message = { role: 'assistant', content };
        if (reasoning) message.reasoning_content = reasoning;
        const normalizedToolCalls = toolCalls.filter(Boolean).map(tc => ({
          id: tc.id || `call_${App.id().replace(/[^a-zA-Z0-9_]/g, '')}`,
          type: tc.type || 'function',
          function: { name: tc.function.name || '', arguments: tc.function.arguments || '{}' }
        })).filter(tc => tc.function.name);
        if (normalizedToolCalls.length) message.tool_calls = normalizedToolCalls;
        const empty = !String(content).trim() && !String(reasoning).trim() && !normalizedToolCalls.length;
        return { result: { choices: [{ message, finish_reason: finishReason }], usage }, empty };
      } finally {
        disarm();
        if (userSignal && typeof userSignal.removeEventListener === 'function') userSignal.removeEventListener('abort', onUserAbort);
      }
    }

    // 空流重试：首个字节前的网络异常、或流正常结束但零内容零工具调用时，原地退避重试。
    // 与上方的 400 自愈（reasoning 回传、图片降级）正交：自愈在一次尝试内部完成，这里是整次尝试重试。
    let forceProxy = false;
    for (let attempt = 0; ; attempt++) {
      try {
        const outcome = await attemptRequest(forceProxy);
        if (!outcome.empty || attempt >= NETWORK_RETRY_DELAYS.length || userAborted()) return outcome.result;
      } catch (e) {
        // 直连被 CORS/网络掐死（Load failed 类）且本地代理可用：切代理重试，不消耗 attempt
        if (e && e.fallbackProxy && !forceProxy) { forceProxy = true; continue; }
        if (!e || !e.retryableRequest || attempt >= NETWORK_RETRY_DELAYS.length || userAborted()) throw e;
      }
      await new Promise(resolve => setTimeout(resolve, NETWORK_RETRY_DELAYS[attempt]));
    }
  }

  function mergeToolCallDeltas(toolCalls, deltas) {
    for (const part of deltas) {
      const index = Number.isInteger(part.index) ? part.index : 0;
      const target = toolCalls[index] || { id: '', type: 'function', function: { name: '', arguments: '' } };
      if (part.id) target.id = part.id;
      if (part.type) target.type = part.type;
      if (part.function) {
        if (part.function.name) target.function.name += part.function.name;
        if (part.function.arguments) target.function.arguments += part.function.arguments;
      }
      toolCalls[index] = target;
    }
  }

  App.currentVisionKey = currentVisionKey;
  App.currentProtocol = currentProtocol;
  App.aiClient = {
    callChatCompletions,
    // 跨请求状态查询：toApiMessages 判断历史推理是否必须回传（api.js 使用）
    reasoningEchoRequired: function (visionKey) { return reasoningEchoRequiredModels.has(visionKey); }
  };
})();
