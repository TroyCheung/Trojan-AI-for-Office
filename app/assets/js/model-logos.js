(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 模型/服务商名 → 品牌 logo（Proma 式「模型即品牌」：切换模型，logo 跟着走）。
  // 资产：Simple Icons（CC0）+ 智谱官网 Z mark（自裁剪），落 assets/brand/logos/。
  // 匹配按小写子串、先命中先用；没命中返回空串，调用方退化为纯文字，不硬造。
  const LOGOS = [
    { file: 'zhipu-mark.svg', name: 'GLM', keys: ['glm', 'zhipu', 'bigmodel', 'chatglm'] },
    { file: 'claude.svg', name: 'Claude', keys: ['claude', 'anthropic'] },
    { file: 'deepseek.svg', name: 'DeepSeek', keys: ['deepseek'] },
    { file: 'qwen.svg', name: 'Qwen', keys: ['qwen', 'tongyi'] },
    { file: 'kimi.svg', name: 'Kimi', keys: ['kimi', 'moonshot'] }
  ];

  // App.modelLogoHtml('GLM-4.6') → '<img class="mlogo" ...>'；未知模型 → ''
  function modelLogoHtml(modelId) {
    const s = String(modelId || '').toLowerCase();
    if (!s) return '';
    const hit = LOGOS.find(l => l.keys.some(k => s.includes(k)));
    if (!hit) return '';
    return `<img class="mlogo" src="assets/brand/logos/${hit.file}" alt="" aria-hidden="true" title="${hit.name}" />`;
  }

  App.modelLogoHtml = modelLogoHtml;
})();
