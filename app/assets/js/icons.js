(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 统一图标集。
  // 规则（新增图标必须遵守，否则又会退回到 emoji + 几何字符混用的状态）：
  //   1. 全部 16×16 viewBox，1.5 描边，round cap / round join
  //   2. 只用 currentColor，不写死颜色，由外层 color 决定
  //   3. 不用 fill（除非是实心点/实心圆这种语义就是实心的）
  //   4. 不用 emoji。emoji 的字重、基线、光学尺寸不受控，还会跟随系统字体变化
  const PATHS = {
    // 状态
    check:      '<path d="M3 8.5l3.2 3.2L13 5"/>',
    close:      '<path d="M4 4l8 8M12 4l-8 8"/>',
    alert:      '<path d="M8 5.5v3.2M8 11.2v.3"/><path d="M6.8 2.6L1.7 11.4a1.3 1.3 0 001.2 2h10.2a1.3 1.3 0 001.2-2L9.2 2.6a1.35 1.35 0 00-2.4 0z"/>',
    info:       '<circle cx="8" cy="8" r="6.2"/><path d="M8 7.4v3.4M8 5.3v.2"/>',
    stop:       '<rect x="4.2" y="4.2" width="7.6" height="7.6" rx="1.4"/>',
    dot:        '<circle cx="8" cy="8" r="3.2" fill="currentColor" stroke="none"/>',

    // 方向
    chevronRight: '<path d="M6.2 3.6L10.6 8l-4.4 4.4"/>',
    chevronDown:  '<path d="M3.6 6.2L8 10.6l4.4-4.4"/>',
    arrowRight:   '<path d="M2.8 8h10.4M9.2 4l4 4-4 4"/>',
    arrowUp:      '<path d="M8 13.2V2.8M4 6.8l4-4 4 4"/>',
    swap:         '<path d="M2.6 5.6h10.8M10.6 2.8l2.8 2.8-2.8 2.8"/><path d="M13.4 10.4H2.6M5.4 7.6l-2.8 2.8 2.8 2.8"/>',
    drag:         '<path d="M8 2.6v10.8M5.4 5.2L8 2.6l2.6 2.6M5.4 10.8L8 13.4l2.6-2.6"/>',

    // 动作
    locate:     '<circle cx="8" cy="8" r="2.1"/><path d="M8 1.4v2.2M8 12.4v2.2M14.6 8h-2.2M3.6 8H1.4"/>',
    edit:       '<path d="M11.1 2.6l2.3 2.3L5.6 12.7 2.5 13.5l.8-3.1z"/>',
    refresh:    '<path d="M13.3 7.1A5.4 5.4 0 003.3 5.4"/><path d="M2.7 8.9a5.4 5.4 0 0010 1.7"/><path d="M2.7 2.9v2.6h2.6M13.3 13.1v-2.6h-2.6"/>',
    plus:       '<path d="M8 3.4v9.2M3.4 8h9.2"/>',
    minus:      '<path d="M3.4 8h9.2"/>',
    send:       '<path d="M13.6 2.4L7.2 8.8M13.6 2.4l-4.1 11.2-2.3-4.8-4.8-2.3z"/>',
    download:   '<path d="M8 2.4v7.6M4.8 7l3.2 3.2L11.2 7"/><path d="M2.6 12.2v.4a1.2 1.2 0 001.2 1.2h8.4a1.2 1.2 0 001.2-1.2v-.4"/>',
    copy:       '<rect x="5.4" y="5.4" width="8" height="8" rx="1.4"/><path d="M10.6 5.4V4a1.4 1.4 0 00-1.4-1.4H4a1.4 1.4 0 00-1.4 1.4v5.2A1.4 1.4 0 004 10.6h1.4"/>',

    upload:     '<path d="M8 13.6V6M4.8 9.2L8 6l3.2 3.2"/><path d="M2.6 3.4h10.8"/>',
    folder:     '<path d="M14.1 12.4a1.3 1.3 0 01-1.3 1.3H3.2a1.3 1.3 0 01-1.3-1.3V3.6a1.3 1.3 0 011.3-1.3h3.2l1.3 2h5.1a1.3 1.3 0 011.3 1.3z"/>',
    clock:      '<circle cx="8" cy="8" r="6.2"/><path d="M8 4.4V8l2.4 1.4"/>',
    enter:      '<path d="M13.4 3.2v4.2a1.4 1.4 0 01-1.4 1.4H2.9"/><path d="M5.9 5.7L2.9 8.8l3 3.1"/>',
    // 对象
    image:     '<rect x="2" y="3" width="12" height="10" rx="1.5"/><circle cx="5.6" cy="6.2" r="1"/><path d="M2.6 11.6l3.2-3 2.5 2.3 2.5-2.7 2.6 3"/>',
    file:       '<path d="M9.2 1.9H4.6a1.3 1.3 0 00-1.3 1.3v9.6a1.3 1.3 0 001.3 1.3h6.8a1.3 1.3 0 001.3-1.3V5.2z"/><path d="M9.2 1.9v3.3h3.5"/>',
    comment:    '<path d="M13.6 9.6a1.3 1.3 0 01-1.3 1.3H5.1L2.4 13.6V3.7a1.3 1.3 0 011.3-1.3h8.6a1.3 1.3 0 011.3 1.3z"/>',
    table:      '<rect x="2.2" y="2.8" width="11.6" height="10.4" rx="1.3"/><path d="M2.2 6.4h11.6M6.5 6.4v6.8"/>',
    layers:     '<path d="M8 1.9L1.9 5 8 8.1 14.1 5z"/><path d="M1.9 11L8 14.1 14.1 11M1.9 8L8 11.1 14.1 8"/>',
    slide:      '<rect x="1.9" y="3" width="12.2" height="8.2" rx="1.2"/><path d="M8 11.2v2.1M5.9 13.3h4.2"/>',
    sparkle:    '<path d="M8 2.1l1.5 3.9 3.9 1.5-3.9 1.5L8 12.9 6.5 9 2.6 7.5 6.5 6z"/>',
    search:     '<circle cx="7.2" cy="7.2" r="4.6"/><path d="M10.6 10.6l3 3"/>',

    // 模式
    shield:     '<path d="M8 1.9l5.1 2.1v4c0 3-2.1 5.2-5.1 6.1-3-0.9-5.1-3.1-5.1-6.1V4z"/>',
    zap:        '<path d="M8.9 1.9L3.4 9h4l-.3 5.1L13 7h-4z"/>',
    listView:   '<path d="M5.6 4.2h8.2M5.6 8h8.2M5.6 11.8h8.2M2.4 4.2v.1M2.4 8v.1M2.4 11.8v.1"/>',
    bolt:       '<path d="M8 1.9l4.4 4.4-4.4 4.4-4.4-4.4z"/><path d="M8 10.7v3.4"/>',

    // 主题
    sun:        '<circle cx="8" cy="8" r="3.1"/><path d="M8 1.4v1.5M8 13.1v1.5M14.6 8h-1.5M2.9 8H1.4M12.7 3.3l-1.1 1.1M4.4 11.6l-1.1 1.1M12.7 12.7l-1.1-1.1M4.4 4.4L3.3 3.3"/>',
    moon:       '<path d="M13.6 9.4A5.9 5.9 0 016.6 2.4a5.9 5.9 0 107 7z"/>'
  };

  const SOLID = { dot: true };

  // 尺寸档位。与 blocks.css 的 --ui-icon-* 对齐，不要在调用处写死 px。
  function icon(name, opts) {
    const d = PATHS[name];
    if (!d) return '';
    const o = opts || {};
    const cls = ['ui-ico', o.size ? `ui-ico-${o.size}` : '', o.className || ''].filter(Boolean).join(' ');
    const label = o.label ? ` role="img" aria-label="${String(o.label).replace(/"/g, '&quot;')}"` : ' aria-hidden="true"';
    const stroke = SOLID[name] ? '' : ' fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"';
    return `<svg class="${cls}" viewBox="0 0 16 16"${stroke}${label}>${d}</svg>`;
  }

  App.icon = icon;
  App.ICON_NAMES = Object.keys(PATHS);
})();
