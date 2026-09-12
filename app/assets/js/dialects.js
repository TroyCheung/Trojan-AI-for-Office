(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 各厂商 OpenAI 兼容接口的「方言」差异。
  // UI 只表达意图（none/low/medium/high），这里负责翻译成各家实际认的参数。
  // 参考：
  //   智谱   https://docs.bigmodel.cn/cn/guide/develop/openai/introduction
  //   DeepSeek https://api-docs.deepseek.com/guides/thinking_mode/
  const DIALECTS = {
    // 通用 OpenAI 兼容：兼容只代表接口形状相似，不代表接受推理参数。
    // 默认不发送；明确识别为智谱或 DeepSeek 后再由对应方言处理。
    openai: {
      applyThinking() {},
      readReasoningDelta(delta) { return (delta && (delta.reasoning_content || delta.reasoning)) || ''; },
      needsReasoningEcho: false
    },

    // 智谱 GLM：thinking 对象是开关；GLM-5.2 及以上支持 reasoning_effort。
    // 官方档位（docs.bigmodel.cn/cn/guide/capabilities/thinking）：
    //   GLM-5.3 起思考强制开启（disabled 会报错），effort 仅 low / high / max，缺省是 max；
    //   GLM-5.2 支持 disabled，effort 里 low/medium 会被服务端映射成 high（没有真正的轻量档）。
    // UI 档位 → 原生：不思考 → disabled（5.3+ 退化为 low）；快速 → low；标准 → high；深度 → max。
    zhipu: {
      applyThinking(body, level, model) {
        const version = glmVersion(model);
        const effortSupported = version && (version.major > 5 || (version.major === 5 && version.minor >= 2));
        const thinkingForced = version && (version.major > 5 || (version.major === 5 && version.minor >= 3));
        if (!level || level === 'none') {
          if (thinkingForced) {
            // GLM-5.3 不能关闭思考，用最低档代替「不思考」
            body.thinking = { type: 'enabled', clear_thinking: false };
            body.reasoning_effort = 'low';
          } else {
            body.thinking = { type: 'disabled' };
          }
          return;
        }
        // clear_thinking:false 让多轮工具调用之间保留推理上下文
        body.thinking = { type: 'enabled', clear_thinking: false };
        if (effortSupported) {
          body.reasoning_effort = level === 'low' ? 'low' : (level === 'medium' ? 'high' : 'max');
        }
      },
      readReasoningDelta(delta) { return (delta && delta.reasoning_content) || ''; },
      needsReasoningEcho: true
    },

    // DeepSeek V4：effort 取值 low/high/max，无 medium；
    // 带工具调用时 reasoning_content 必须原样回传，否则接口返回 400
    deepseek: {
      applyThinking(body, level) {
        if (!level || level === 'none') { body.thinking = { type: 'disabled' }; return; }
        body.thinking = { type: 'enabled' };
        body.reasoning_effort = level === 'low' ? 'low' : (level === 'medium' ? 'high' : 'max');
      },
      readReasoningDelta(delta) { return (delta && delta.reasoning_content) || ''; },
      needsReasoningEcho: true
    }
  };

  function glmVersion(model) {
    const match = String(model || '').match(/glm-(\d+)\.(\d+)/i);
    return match ? { major: Number(match[1]), minor: Number(match[2]) } : null;
  }

  // 优先用服务配置里显式声明的 dialect；没有就从 baseUrl 推断（旧配置自愈，无需迁移）
  function dialectIdFor(providerId, baseUrl) {
    const services = (App.state && App.state.settings && App.state.settings.serviceProviders) || [];
    const svc = services.find(s => s && s.id === providerId);
    if (svc && svc.dialect && DIALECTS[svc.dialect]) return svc.dialect;
    const url = String(baseUrl || (svc && svc.baseUrl) || '').toLowerCase();
    if (url.includes('bigmodel.cn') || url.includes('z.ai') || url.includes('zhipu')) return 'zhipu';
    if (url.includes('deepseek')) return 'deepseek';
    return 'openai';
  }

  function currentDialect() {
    const s = (App.state && App.state.settings) || {};
    return DIALECTS[dialectIdFor(s.provider, s.customPrefixUrl)] || DIALECTS.openai;
  }

  App.DIALECTS = DIALECTS;
  App.dialectIdFor = dialectIdFor;
  App.currentDialect = currentDialect;
})();
