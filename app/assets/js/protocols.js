(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ── Anthropic Messages 协议适配器（v136）──
  // 依据 Anthropic 官方 Messages API 公开文档实现（协议是公开标准；不复制任何项目代码）。
  // 定位：纯函数转换层。loop 与历史存储保持 OpenAI 形状不变，只在请求出口（outbound）
  // 与 SSE 入口（createSseTranslator）做双向翻译。网络韧性/重试/降级全部复用现有层。

  const DEFAULT_MAX_TOKENS = 8192;

  function safeParseJson(text) {
    try { return JSON.parse(String(text || '{}')); } catch { return {}; }
  }

  // OpenAI user content（string 或 blocks 数组）→ Anthropic content blocks。
  // 统一产出 block 数组：相邻 user 合并时 concat 始终类型安全（裸字符串混进 block 数组必 400）。
  function toUserBlocks(content) {
    if (typeof content === 'string') return [{ type: 'text', text: content }];
    if (!Array.isArray(content)) return [{ type: 'text', text: String(content == null ? '' : content) }];
    const blocks = [];
    for (const b of content) {
      if (typeof b === 'string') { blocks.push({ type: 'text', text: b }); continue; }
      if (!b || !b.type) continue;
      if (b.type === 'text') { blocks.push({ type: 'text', text: String(b.text || '') }); continue; }
      if (b.type === 'image_url') {
        const src = String((b.image_url && b.image_url.url) || '');
        const m = src.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
        if (m) blocks.push({ type: 'image', source: { type: 'base64', media_type: m[1], data: m[2] } });
        else blocks.push({ type: 'text', text: '[图片链接无法直接发给 Anthropic 接口，请重新上传图片文件]' });
        continue;
      }
      // 未知 block 类型降级为文本占位，不静默丢
      blocks.push({ type: 'text', text: `[不支持的输入块: ${b.type}]` });
    }
    return blocks.length ? blocks : [{ type: 'text', text: ' ' }];
  }

  // ── 请求出口：OpenAI 形状 body → Anthropic /messages body ──
  // 白名单重建（不从入身上改字段）：dialect 注入的 thinking 等专有参数一律不透传。
  function outbound(body) {
    const srcMessages = Array.isArray(body.messages) ? body.messages : [];
    const systems = [];
    const rest = [];
    for (const m of srcMessages) {
      if (m && m.role === 'system') systems.push(String(m.content || ''));
      else rest.push(m);
    }

    // OpenAI wire → Anthropic messages。role:'tool' 累积成一条 user 消息的 tool_result blocks
    //（Anthropic 规定 tool_result 必须住在 user 消息里，且紧跟所属 tool_use 的 assistant 消息之后）
    const out = [];
    let pendingToolResults = null;
    const flushToolResults = () => {
      if (!pendingToolResults) return;
      out.push({ role: 'user', content: pendingToolResults });
      pendingToolResults = null;
    };
    for (const m of rest) {
      if (!m || !m.role) continue;
      if (m.role === 'tool') {
        if (!pendingToolResults) pendingToolResults = [];
        pendingToolResults.push({
          type: 'tool_result',
          tool_use_id: String(m.tool_call_id || ''),
          content: String(m.content == null ? '' : m.content)
        });
        continue;
      }
      flushToolResults();
      if (m.role === 'user') {
        out.push({ role: 'user', content: toUserBlocks(m.content) });
      } else if (m.role === 'assistant') {
        const blocks = [];
        const text = String(m.content || '');
        if (text.trim()) blocks.push({ type: 'text', text });
        for (const tc of (Array.isArray(m.tool_calls) ? m.tool_calls : [])) {
          blocks.push({
            type: 'tool_use',
            id: String(tc.id || ''),
            name: String(tc.function && tc.function.name || ''),
            input: safeParseJson(tc.function && tc.function.arguments)
          });
        }
        if (!blocks.length) blocks.push({ type: 'text', text: ' ' }); // 空 assistant 不合法，占位
        out.push({ role: 'assistant', content: blocks });
      }
    }
    flushToolResults();

    // 相邻同 role 合并（Anthropic 要求 user/assistant 交替；保守自并，防网关挑剔）
    const merged = [];
    for (const m of out) {
      const prev = merged[merged.length - 1];
      if (prev && prev.role === m.role) {
        prev.content = prev.content.concat(m.content);
      } else merged.push(m);
    }

    const converted = {
      model: String(body.model || ''),
      max_tokens: DEFAULT_MAX_TOKENS,
      messages: merged
    };
    if (body.stream) converted.stream = true; // 不带 stream Anthropic 返回非流式 JSON，整条流式链路就断了
    if (systems.length) converted.system = systems.join('\n\n');
    if (typeof body.temperature === 'number') converted.temperature = Math.max(0, Math.min(1, body.temperature));
    if (Array.isArray(body.tools) && body.tools.length) {
      converted.tools = body.tools.map(t => ({
        name: String(t.function && t.function.name || ''),
        description: String(t.function && t.function.description || ''),
        input_schema: (t.function && t.function.parameters) || { type: 'object', properties: {} }
      }));
      converted.tool_choice = { type: 'auto' };
    }
    return converted;
  }

  // ── SSE 入口：每次请求一个翻译器，把 Anthropic 事件合成 OpenAI 形状 chunk ──
  // 返回的函数吃 data: 行的 JSON 字符串，产出喂给现有 processLine 的 chunk（或 null 跳过）。
  // processLine 只认 data: 行，event: 行天然被忽略，无需特判。
  function createSseTranslator() {
    let toolBlockCount = 0;              // 已出现的 tool_use 块数 → OpenAI tool_calls 的 index
    const blockIndexToToolIndex = {};    // Anthropic content_block index → tool_calls index
    const usageAcc = {};                 // usage 在 message_start/message_delta 两段分发，消费方是覆盖语义，产出时给全量
    return function translate(data) {
      const ev = safeParseJson(data);
      if (!ev || !ev.type) return null;
      if (ev.type === 'error') {
        return { error: { message: (ev.error && ev.error.message) || 'Anthropic 流中报错' } };
      }
      if (ev.type === 'message_start') {
        const u = ev.message && ev.message.usage;
        if (u && typeof u.input_tokens === 'number') usageAcc.prompt_tokens = u.input_tokens;
        return Object.keys(usageAcc).length ? { choices: [{ delta: {} }], usage: Object.assign({}, usageAcc) } : null;
      }
      if (ev.type === 'content_block_start' && ev.content_block && ev.content_block.type === 'tool_use') {
        const toolIndex = toolBlockCount++;
        blockIndexToToolIndex[ev.index] = toolIndex;
        return { choices: [{ delta: { tool_calls: [{ index: toolIndex, id: ev.content_block.id || '', type: 'function', function: { name: ev.content_block.name || '', arguments: '' } }] } }] };
      }
      if (ev.type === 'content_block_delta') {
        const d = ev.delta || {};
        if (d.type === 'text_delta' && typeof d.text === 'string') {
          return { choices: [{ delta: { content: d.text } }] };
        }
        if (d.type === 'input_json_delta' && typeof d.partial_json === 'string') {
          const toolIndex = blockIndexToToolIndex[ev.index];
          if (toolIndex == null) return null;
          return { choices: [{ delta: { tool_calls: [{ index: toolIndex, function: { arguments: d.partial_json } }] } }] };
        }
        // thinking_delta 等不透传（Anthropic 思考流不做回放）
        return null;
      }
      if (ev.type === 'message_delta') {
        const map = { tool_use: 'tool_calls', end_turn: 'stop', stop_sequence: 'stop', max_tokens: 'length' };
        const finish = ev.delta && ev.delta.stop_reason ? (map[ev.delta.stop_reason] || 'stop') : undefined;
        const chunk = { choices: [{ delta: {}, finish_reason: finish }] };
        if (ev.usage && typeof ev.usage.output_tokens === 'number') {
          usageAcc.completion_tokens = ev.usage.output_tokens;
          chunk.usage = Object.assign({}, usageAcc);
        }
        return chunk;
      }
      // content_block_stop / message_stop / ping → 无需产出
      return null;
    };
  }

  // ── 传输参数 ──
  function endpoint(baseUrl) {
    const base = String(baseUrl || '').replace(/\/+$/, '');
    if (/\/messages$/.test(base)) return base;
    if (/\/v1$/.test(base)) return base + '/messages';
    return base + '/v1/messages';
  }

  function headers(apiKey) {
    return {
      'Content-Type': 'application/json',
      'x-api-key': String(apiKey || ''),
      // 标准是 x-api-key；部分中转只认 Bearer，两个都带不冲突
      'Authorization': `Bearer ${String(apiKey || '')}`,
      'anthropic-version': '2023-06-01',
      // Anthropic 官方 API 默认拒绝浏览器直连（不返回 CORS 头），此头是官方提供的
      // 浏览器直连开关——缺失时 WKWebView 的 fetch 直接报 "Load failed"
      'anthropic-dangerous-direct-browser-access': 'true'
    };
  }

  App.anthropicProtocol = { outbound, createSseTranslator, endpoint, headers, DEFAULT_MAX_TOKENS };
})();
