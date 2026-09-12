(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 轻量 toast：底部浮层，自动消失。样式在 blocks.css 尾部 .toast-* 区。
  // 设计约定（2026-09 UI 批次）：只用于「无内联回执的异步结果」场景；
  // 已有可见反馈的操作（复制按钮变字、删除后列表变化等）不要重复接，会噪音化。
  let host = null;

  function ensureHost() {
    if (host && document.body && document.body.contains(host)) return host;
    host = document.createElement('div');
    host.className = 'toast-host';
    host.setAttribute('aria-live', 'polite');
    document.body.appendChild(host);
    return host;
  }

  // App.toast(message, { tone: 'info' | 'success' | 'error', ms: 2600 })
  function toast(message, opts) {
    if (!document.body) return;
    const o = opts || {};
    const el = document.createElement('div');
    el.className = 'toast tone-' + (o.tone || 'info');
    el.textContent = String(message == null ? '' : message);
    ensureHost().appendChild(el);
    // 下一帧再加 in 类，保证入场过渡必然播放
    requestAnimationFrame(() => { requestAnimationFrame(() => el.classList.add('in')); });
    setTimeout(() => {
      el.classList.remove('in');
      el.addEventListener('transitionend', () => el.remove(), { once: true });
      setTimeout(() => { if (el.parentNode) el.remove(); }, 500); // transitionend 兜底（reduced-motion 下不触发）
    }, Math.max(800, Number(o.ms) || 2600));
  }

  App.toast = toast;
})();
