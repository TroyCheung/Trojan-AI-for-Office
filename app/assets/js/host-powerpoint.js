(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const requireOffice = () => App.requireOffice();
  function clampText(s, n) { s = String(s == null ? '' : s); return s.length > n ? s.slice(0, n) + '…' : s; }

  // 本宿主已验证的 Office.js 行为约束——唯一来源。
  // 以前这些经验散在三处：插件自己的防御性代码注释里、抛出的错误字符串里、以及模型每轮对话的重新试错里，
  // 于是同一个坑要被发现很多次。现在集中在这里，由 eval_officejs 的说明、recolor_slide 的实现和读页提示共同消费。
  const HOST_QUIRKS = {
    batchSync: '批量 load 多个形状后再一次 sync 会整体失败；必须逐个形状独立 load + sync。',
    collections: '集合用 items[i] 索引，没有 getItem(index)。',
    fillApi: 'fill 没有 setType()；改填充色用 fill.setSolidColor("#RRGGBB")。',
    groups: '组合（Group）的子形状读不到（shapes.items 恒为空），也没有 ungroup()——组合内部的元素无法程序化修改。需要改动组合内部时，直接请用户在 PowerPoint 里选中该组合手动取消组合（Cmd+Shift+G）后再跑一次，不要反复尝试各种读取方式。',
    duplicate: 'slide.duplicate() 不存在时，duplicate_slide 自动走文件级复制（getFileAsync + insertSlidesFromBase64 + sourceSlideIds）；仍失败才走 list_layouts + add_slide(layoutId) 借版。',
    pageSetup: 'pageSetup.slideWidth / slideHeight 可能读回 undefined；页面尺寸以 get_slide 返回的 page 为准。',
    imageReplace: '图片没有原位替换 API：replace_image 在同位置删除旧图再插入新图（shapes.addImage），新图的叠放次序是顶层，需要压到底层时用 apply_layout 的 zOrder 修正。',
    tableCellFill: '表格单元格底色（TableCell.fill）是否可写因宿主而异：set_table_style 会逐格探测，写不了时在结果里明说并给替代路径，不会谎报成功。'
  };
  const HOST_QUIRKS_TEXT = Object.values(HOST_QUIRKS).map((line, i) => `(${i + 1}) ${line}`).join(' ');

  function powerPointApiSupported(version) {
    try { return Boolean(Office.context.requirements.isSetSupported('PowerPointApi', version)); }
    catch { return false; }
  }

  // 运行时实测能力缓存。API 版本号声明支持 != 真能用：Mac PowerPoint 16.112 声明 PowerPointApi 1.1-1.10
  // 全量支持，但 slide.duplicate 不是函数、pageSetup.slideWidth 读回 undefined。
  // 这里只记录「实际调用时观察到的结果」，不写死任何一台机器的探测报告——换机器或微软补齐 API 后会自动纠正。
  const runtimeProbe = {
    duplicateSlide: null, exactPageSize: null, imageReplaceInPlace: null, tableCellFill: null, chartSeriesValues: null,
    slideBackgroundApi: null, tableStructureApi: null, pictureCrop: null, pictureTransparency: null, duplicateViaBase64: null,
    hyperlinkApi: null, shapeAdjustments: null
  };
  // insertSlidesFromBase64 的 sourceSlideIds 实际接受的 id 格式（实测成功后记录）：
  // 'full' = 完整 slide.id（如 2147481232#588170797），'suffix' = # 后面的部分。null = 还没探测。
  let duplicateBase64IdFormat = null;

  function noteRuntimeProbe(key, value) {
    runtimeProbe[key] = Boolean(value);
    return runtimeProbe[key];
  }

  function presentationCapabilities() {
    const modern = powerPointApiSupported('1.8');
    return {
      shapeLayout: powerPointApiSupported('1.4'),
      zOrder: modern,
      zOrderMethod: 'shape.setZOrder',
      slidePreview: modern,
      slidePreviewMethod: 'slide.getImageAsBase64',
      // 声明值只作初值：真读过一次页面尺寸后以实测结果为准
      exactPageSize: runtimeProbe.exactPageSize === null ? powerPointApiSupported('1.10') : runtimeProbe.exactPageSize,
      slideLayouts: powerPointApiSupported('1.3'),
      duplicateSlide: runtimeProbe.duplicateSlide,
      // 文件级复制路径（slide.duplicate 缺失时的自动降级）：null = 还没探测，true/false = 实测结果
      duplicateViaBase64: runtimeProbe.duplicateViaBase64,
      duplicateBase64IdFormat,
      // 以下四项只有实测过才有值：null = 还没探测（调一次对应工具即实测写回），不写死
      slideBackgroundApi: runtimeProbe.slideBackgroundApi,
      tableStructureApi: runtimeProbe.tableStructureApi,
      pictureCrop: runtimeProbe.pictureCrop,
      pictureTransparency: runtimeProbe.pictureTransparency,
      hyperlinkApi: runtimeProbe.hyperlinkApi,
      // 【43】绝对圆角半径（卡头色条贴合圆角外框用）：null=未实测，先用声明值；实测过以实测为准
      shapeAdjustments: runtimeProbe.shapeAdjustments === null ? powerPointApiSupported('1.10') : runtimeProbe.shapeAdjustments,
      note: (modern
        ? 'This host supports shape.setZOrder and slide.getImageAsBase64.'
        : 'This host can add and format shapes when shapeLayout is true, but cannot change z-order or render a slide preview through Office.js.')
        + ' duplicateSlide: null = not probed yet; false = no native slide.duplicate — duplicate_slide then falls back to file-level copy via insertSlidesFromBase64 (duplicateViaBase64 reports whether that worked); only if both fail, borrow a design with list_layouts + add_slide(layoutId) instead.'
        + ' slideBackgroundApi/tableStructureApi/pictureCrop/pictureTransparency: null = not probed yet — call set_slide_background / edit_table_structure / crop_image / set_picture_opacity once to probe; false = confirmed unsupported, use the fallback the tool error message names.'
        + ' hyperlinkApi: null = not probed yet — call set_hyperlink once to probe; false = confirmed unsupported, write the URL as plain text instead and tell the user to add the link manually (Cmd+K).'
    };
  }

  function auditColor(value) {
    const text = String(value || '').trim().toUpperCase();
    return /^#[0-9A-F]{6}$/.test(text) ? text : '';
  }

  function uniqueSortedNumbers(values) {
    return [...new Set(values.filter(Number.isFinite).map(value => Math.round(value * 10) / 10))].sort((a, b) => b - a);
  }

  function overlapArea(a, b) {
    const width = Math.max(0, Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left));
    const height = Math.max(0, Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top));
    return width * height;
  }

  function textWidthUnits(text) {
    let units = 0;
    for (const char of String(text || '')) {
      if (/\s/.test(char)) units += 0.3;
      else if (/[\u2E80-\u9FFF\uF900-\uFAFF]/.test(char)) units += 1;
      else if (/[A-Z]/.test(char)) units += 0.66;
      else if (/[a-z]/.test(char)) units += 0.52;
      else if (/[0-9]/.test(char)) units += 0.56;
      else units += 0.38;
    }
    return units;
  }

  function estimatedWrappedLines(text, availableWidth, fontSize, wordWrap) {
    const capacity = Math.max(1, availableWidth / Math.max(1, fontSize));
    let lines = 0;
    for (const paragraph of String(text || '').split(/\r?\n/)) {
      if (!paragraph) { lines += 1; continue; }
      if (wordWrap === false) { lines += 1; continue; }
      let used = 0;
      const tokens = paragraph.split(/(\s+)/).filter(Boolean);
      for (const token of tokens) {
        let width = textWidthUnits(token);
        if (width <= capacity) {
          if (used > 0 && used + width > capacity) { lines += 1; used = 0; }
          used += width;
          continue;
        }
        if (used > 0) { lines += 1; used = 0; }
        lines += Math.max(0, Math.ceil(width / capacity) - 1);
        width %= capacity;
        used = width;
      }
      lines += 1;
    }
    return Math.max(1, lines);
  }

  function textFitEstimate(shape) {
    const fontSize = Number(shape.fontSize);
    if (!Number.isFinite(fontSize) || fontSize <= 0) return null;
    const horizontalMargins = Math.max(0, Number(shape.marginLeft) || 0) + Math.max(0, Number(shape.marginRight) || 0);
    const verticalMargins = Math.max(0, Number(shape.marginTop) || 0) + Math.max(0, Number(shape.marginBottom) || 0);
    const availableWidth = Math.max(1, Number(shape.width) - horizontalMargins);
    const availableHeight = Math.max(1, Number(shape.height) - verticalMargins);
    const lineHeight = fontSize * 1.18;
    const estimatedLines = estimatedWrappedLines(shape.text, availableWidth, fontSize, shape.wordWrap);
    const availableLines = Math.max(1, Math.floor(availableHeight / lineHeight));
    return {
      shapeId: shape.id,
      estimatedLines,
      availableLines,
      fontSize,
      autoSizeSetting: String(shape.autoSizeSetting || ''),
      availableWidth: Math.round(availableWidth * 10) / 10,
      availableHeight: Math.round(availableHeight * 10) / 10
    };
  }

  function buildLayoutAudit(shapes, page = {}) {
    const pageWidth = Number.isFinite(page.width) ? page.width : 960;
    const pageHeight = Number.isFinite(page.height) ? page.height : 540;
    const positioned = (Array.isArray(shapes) ? shapes : []).filter(shape =>
      Number.isFinite(shape.left) && Number.isFinite(shape.top) && Number.isFinite(shape.width) && Number.isFinite(shape.height)
      && Math.max(shape.width, shape.height) > 1 && shape.width >= 0 && shape.height >= 0
    );
    const visible = positioned.filter(shape => shape.width > 1 && shape.height > 1);
    const textShapes = visible.filter(shape => String(shape.text || '').trim());
    const issues = [];
    // 降级快照（getSlide 的 fallback 分支产出 styleReadable:false）没有字号、填充、边距数据。
    // 此时样式类检查必然误报/漏报，只保留几何类检查，并在结果里标明审查不完整。
    const stylesReadable = !positioned.some(shape => shape.styleReadable === false);

    const outside = positioned.filter(shape => shape.left < -1 || shape.top < -1
      || shape.left + shape.width > pageWidth + 1 || shape.top + shape.height > pageHeight + 1);
    if (outside.length) {
      // 图片/无文本形状出界多为刻意的出血设计，降为 low；文字出界仍然 high
      const hasTextContent = shape => Boolean(String(shape.text || '').trim());
      const outsideText = outside.filter(hasTextContent);
      const outsideVisual = outside.filter(shape => !hasTextContent(shape));
      if (outsideText.length) {
        issues.push({
          code: 'out_of_bounds', severity: 'high',
          message: `${outsideText.length} 个文字形状超出幻灯片边界。`,
          shapeIds: outsideText.map(shape => shape.id)
        });
      }
      if (outsideVisual.length) {
        issues.push({
          code: 'out_of_bounds', severity: 'low',
          message: `${outsideVisual.length} 个无文本形状（图片/装饰）超出幻灯片边界，可能是刻意的出血设计。`,
          shapeIds: outsideVisual.map(shape => shape.id)
        });
      }
    }

    // 装饰性重叠豁免：一方是字号 2 倍以上的装饰大字，或透明度 >0.5 的水印式文字，不判 high
    const isDecorativeOverlap = (a, b) => {
      const sizeA = Number(a.fontSize), sizeB = Number(b.fontSize);
      if (Number.isFinite(sizeA) && Number.isFinite(sizeB) && sizeA > 0 && sizeB > 0
        && Math.max(sizeA, sizeB) >= 2 * Math.min(sizeA, sizeB)) return true;
      return Number(a.fillTransparency) > 0.5 || Number(b.fillTransparency) > 0.5;
    };
    const collisions = [];
    const decorativeCollisions = [];
    for (let i = 0; i < textShapes.length; i++) {
      for (let j = i + 1; j < textShapes.length; j++) {
        const a = textShapes[i], b = textShapes[j];
        const area = overlapArea(a, b);
        const smaller = Math.min(a.width * a.height, b.width * b.height);
        if (smaller > 0 && area / smaller >= 0.12) {
          (isDecorativeOverlap(a, b) ? decorativeCollisions : collisions).push([a.id, b.id]);
        }
      }
    }
    if (collisions.length) {
      issues.push({
        code: 'text_overlap', severity: 'high',
        message: `${collisions.length} 组文字形状存在明显重叠。`,
        pairs: collisions
      });
    }
    if (decorativeCollisions.length) {
      issues.push({
        code: 'text_overlap', severity: 'low',
        message: `${decorativeCollisions.length} 组文字重叠涉及装饰性大字/水印式文字，通常为刻意设计。`,
        pairs: decorativeCollisions
      });
    }

    const fontSizes = stylesReadable ? uniqueSortedNumbers(textShapes.map(shape => Number(shape.fontSize))) : [];
    if (stylesReadable && textShapes.length >= 3 && fontSizes.length < 2) {
      issues.push({ code: 'flat_type_hierarchy', severity: 'medium', message: '文字较多，但字号没有形成清楚的层级。' });
    }
    if (stylesReadable && fontSizes.length > 5) {
      issues.push({ code: 'too_many_font_sizes', severity: 'medium', message: `本页使用了 ${fontSizes.length} 种字号，视觉层级可能过碎。` });
    }

    const tinyText = stylesReadable ? textShapes.filter(shape => Number.isFinite(Number(shape.fontSize)) && Number(shape.fontSize) < 12 && String(shape.text || '').trim().length > 8) : [];
    if (tinyText.length) {
      issues.push({
        code: 'tiny_text', severity: 'medium',
        message: `${tinyText.length} 个文本框字号小于 12pt。`,
        shapeIds: tinyText.map(shape => shape.id)
      });
    }

    const fitEstimates = stylesReadable ? textShapes.map(textFitEstimate).filter(Boolean) : [];
    const overflowing = fitEstimates.filter(item => item.estimatedLines > item.availableLines
      && item.autoSizeSetting !== 'AutoSizeTextToFitShape');
    if (overflowing.length) {
      issues.push({
        code: 'text_overflow_risk', severity: 'high',
        message: `${overflowing.length} 个文本框按当前字号和空间无法完整容纳文字。`,
        shapeIds: overflowing.map(item => item.shapeId),
        details: overflowing
      });
    }

    const autoShrinking = fitEstimates.filter(item => item.estimatedLines > item.availableLines
      && item.autoSizeSetting === 'AutoSizeTextToFitShape');
    if (autoShrinking.length) {
      issues.push({
        code: 'text_autoshrink_risk', severity: 'medium',
        message: `${autoShrinking.length} 个文本框需要自动缩小字号才能装下，可能破坏参考页字号层级。`,
        shapeIds: autoShrinking.map(item => item.shapeId),
        details: autoShrinking
      });
    }

    const unexpectedWrap = fitEstimates.filter(item => item.fontSize >= 18 && item.availableLines <= 1 && item.estimatedLines > 1);
    if (unexpectedWrap.length) {
      issues.push({
        code: 'unexpected_title_wrap', severity: 'high',
        message: `${unexpectedWrap.length} 个单行标题或标签发生了意外换行。`,
        shapeIds: unexpectedWrap.map(item => item.shapeId),
        details: unexpectedWrap
      });
    }

    const thinRules = positioned.filter(shape => Math.min(Number(shape.width), Number(shape.height)) <= 3
      && Math.max(Number(shape.width), Number(shape.height)) >= 24);
    const ruleCollisions = [];
    for (const textShape of textShapes) {
      for (const rule of thinRules) {
        if (textShape.id === rule.id) continue;
        const horizontal = rule.width >= rule.height;
        const coordinate = horizontal ? rule.top + rule.height / 2 : rule.left + rule.width / 2;
        const crossesText = horizontal
          ? coordinate > textShape.top + 2 && coordinate < textShape.top + textShape.height - 2
            && Math.min(rule.left + rule.width, textShape.left + textShape.width) - Math.max(rule.left, textShape.left) > 12
          : coordinate > textShape.left + 2 && coordinate < textShape.left + textShape.width - 2
            && Math.min(rule.top + rule.height, textShape.top + textShape.height) - Math.max(rule.top, textShape.top) > 12;
        if (crossesText) ruleCollisions.push([textShape.id, rule.id]);
      }
    }
    if (ruleCollisions.length) {
      issues.push({
        code: 'text_rule_collision', severity: 'high',
        message: `${ruleCollisions.length} 处文字框与分隔线穿插。`,
        pairs: ruleCollisions
      });
    }

    // 【37/39】空占位框：有占位身份、文本为空、且 containedType 可靠读回 null
    //（confirmed-empty）才算确认空。containedType 非空（Image/Chart…）是有非文字内容，
    // 读不到（unknown）不冒充空——两者都不得进清理候选，只在 metrics 里诚实记录。
    const emptyPlaceholders = visible.filter(shape => shape.placeholderType && shape.placeholderContent === 'confirmed-empty' && !String(shape.text || '').trim());
    const unreadablePlaceholders = visible.filter(shape => shape.placeholderType && shape.placeholderContent === 'unknown' && !String(shape.text || '').trim());
    const placeholderOverlaps = [];
    for (const textShape of textShapes) {
      for (const ph of emptyPlaceholders) {
        if (textShape.id === ph.id) continue;
        // 【39-R2】相交占比较小者面积：大占位框完全包含小文字时占比=1（按占位框面积
        // 算只有几个百分点，曾漏检）；边缘轻触（占文字面积 <12%）不报
        const area = overlapArea(textShape, ph);
        const smaller = Math.min(textShape.width * textShape.height, ph.width * ph.height);
        if (smaller > 0 && area / smaller >= 0.12) placeholderOverlaps.push([textShape.id, ph.id]);
      }
    }
    if (placeholderOverlaps.length) {
      issues.push({
        code: 'empty_placeholder_overlap', severity: 'high',
        message: `${placeholderOverlaps.length} 处文字压在确认空的占位框上：被新布局替代的空占位框应在提案中明确复用或删除（apply_layout deleteShape），不能把新内容叠在上面。`,
        pairs: placeholderOverlaps
      });
    }
    if (unreadablePlaceholders.length) {
      issues.push({
        code: 'placeholder_content_unreadable', severity: 'low',
        message: `${unreadablePlaceholders.length} 个疑似空占位框的内容状态不可读（旧宿主或读取失败）：未纳入清理候选，不要按空占位框处理。`,
        shapeIds: unreadablePlaceholders.map(shape => shape.id)
      });
    }

    const colors = new Set();
    if (stylesReadable) {
      for (const shape of visible) {
        for (const value of [shape.fillColor, shape.lineColor, shape.fontColor]) {
          const color = auditColor(value);
          if (color) colors.add(color);
        }
      }
    }
    if (stylesReadable && colors.size > 9) {
      issues.push({ code: 'too_many_colors', severity: 'low', message: `本页检测到 ${colors.size} 种颜色，可能削弱统一感。` });
    }

    const alignmentAnchors = (key) => {
      const buckets = [];
      for (const shape of textShapes) {
        const value = Math.round(Number(shape[key]));
        if (!Number.isFinite(value)) continue;
        const bucket = buckets.find(item => Math.abs(item.value - value) <= 3);
        if (bucket) { bucket.values.push(value); bucket.value = Math.round(bucket.values.reduce((sum, item) => sum + item, 0) / bucket.values.length); }
        else buckets.push({ value, values: [value] });
      }
      return buckets.filter(item => item.values.length >= 2).map(item => ({ position: item.value, count: item.values.length }));
    };

    return {
      page: { width: pageWidth, height: pageHeight, exact: Boolean(page.exact) },
      stylesReadable,
      metrics: {
        shapeCount: visible.length,
        textShapeCount: textShapes.length,
        fontSizes,
        colorCount: colors.size,
        emptyPlaceholders: emptyPlaceholders.map(shape => ({ id: shape.id, placeholderType: shape.placeholderType })),
        placeholdersUnreadable: unreadablePlaceholders.map(shape => ({ id: shape.id, placeholderType: shape.placeholderType })),
        repeatedLeftEdges: alignmentAnchors('left'),
        repeatedTopEdges: alignmentAnchors('top')
      },
      issues,
      note: stylesReadable
        ? '这是几何与格式检查，不代替对整页预览图的视觉判断。没有 high 级问题不等于整页视觉验收通过——视觉结论必须依据 get_slide_preview 的实际预览图。'
        : '本次为降级读取，样式数据不可读：只完成了几何类检查（越界/重叠/穿插），字号、颜色、文字装载未检查。没有 high 级问题不等于整页视觉验收通过。'
    };
  }

  async function pageSizeForContext(context) {
    const fallback = { width: 960, height: 540, exact: false };
    if (!powerPointApiSupported('1.10') || !context.presentation.pageSetup) return fallback;
    try {
      const setup = context.presentation.pageSetup;
      setup.load('slideWidth,slideHeight');
      await context.sync();
      const width = Number(setup.slideWidth);
      const height = Number(setup.slideHeight);
      // Mac PowerPoint 16.112 声明支持 1.10，实测这两个值读回 undefined。不校验就会把 NaN 当成页面尺寸，
      // 越界/重叠/穿插审计全部静默失效，还对模型谎称 exact:true。
      if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
        noteRuntimeProbe('exactPageSize', false);
        return fallback;
      }
      noteRuntimeProbe('exactPageSize', true);
      return { width, height, exact: true };
    } catch (e) {
      noteRuntimeProbe('exactPageSize', false);
      return fallback;
    }
  }

  function shapeLabel(shape) {
    try { return String(shape.name || shape.id || ''); } catch { return ''; }
  }

  // 逐形状独立处理的骨架（HOST_QUIRKS.batchSync 的直接后果）：每个形状自己 load、自己 sync、自己 try，
  // 单个形状失败只记下它自己，不拖垮整页。getSlideCore 的样式读取和 recolorSlide 的改色共用这一个骨架，
  // 避免同一个模式在文件里写第三遍。
  async function eachShapeIsolated(context, slide, visit) {
    const failures = [];
    for (const shape of (slide.shapes.items || [])) {
      try { await visit(shape); }
      catch (e) { failures.push({ shape: shapeLabel(shape), error: e && e.message ? e.message : String(e) }); }
    }
    return { failures };
  }

  // 先整批试一次，失败了再逐项退避。正常形状零额外开销；异形形状不会因为一个属性不存在就整条数据全丢。
  // 这是 HOST_QUIRKS.batchSync 的正确应对：一批 load 里任何一个请求失败，整批 sync 都会被取消。
  async function loadIsolated(context, loaders) {
    const ok = new Set();
    try {
      loaders.forEach(loader => loader.apply());
      await context.sync();
      loaders.forEach(loader => ok.add(loader.key));
      return ok;
    } catch { /* 整批失败，下面逐项重试 */ }
    for (const loader of loaders) {
      try { loader.apply(); await context.sync(); ok.add(loader.key); } catch {}
    }
    return ok;
  }

  // 读一个形状的样式，能读到多少读多少。
  // Line / Connector 没有 textFrame，Group 既没有 textFrame 也没有 fill，Picture 没有 textFrame——
  // 把它们和普通形状放进同一批 load，整批会被取消，连本来读得到的线条颜色也一起丢
  // （「红箭头是 #FF0000 却没出现在调色板里」就是这么来的）。所以逐子对象退避。
  async function readShapeStyle(context, shape) {
    const ok = await loadIsolated(context, [
      { key: 'fill', apply: () => shape.fill.load('type,foregroundColor,transparency') },
      { key: 'line', apply: () => shape.lineFormat.load('visible,color,transparency,weight') },
      { key: 'frame', apply: () => shape.textFrame.load('hasText,leftMargin,rightMargin,topMargin,bottomMargin,verticalAlignment,wordWrap,autoSizeSetting') }
    ]);
    let hasText = false;
    if (ok.has('frame')) {
      try { hasText = Boolean(shape.textFrame.hasText); } catch { ok.delete('frame'); }
    }
    if (hasText) {
      // 文字这一层同样要分开退避：paragraphFormat 读不到时，字体名和字色本来是读得到的，
      // 放在同一批里会被一起取消——换肤时就会看不见这个形状用了什么字体。
      const textOk = await loadIsolated(context, [
        { key: 'text', apply: () => shape.textFrame.textRange.load('text') },
        { key: 'font', apply: () => shape.textFrame.textRange.font.load('name,size,color,bold') },
        { key: 'paragraph', apply: () => shape.textFrame.textRange.paragraphFormat.load('horizontalAlignment') }
      ]);
      textOk.forEach(key => ok.add(key));
      if (!textOk.size) hasText = false;
    }
    return { ok, hasText };
  }

  // 读取形状文本：不同平台 shape.textFrame 兼容性不一，统一防御性处理。
  // 单个形状独立 sync 提取文本：无文本框（图片等）的形状 sync 会失败，跳过即可，不再拖垮整页
  async function shapeTextSafe(context, shape) {
    try {
      shape.textFrame.textRange.load('text');
      await context.sync();
      const t = shape.textFrame.textRange.text;
      return typeof t === 'string' ? t : '';
    } catch (e) { return ''; }
  }

  // 【harden-ppt-excel-recovery 2.1】带可读性的文本读取：读取失败必须与「真空文本」分开。
  // 锚点核验拿不到证据时按不可核验处理（拒绝/打回），不把读取失败冒充空文本。
  async function shapeTextDetailed(context, shape) {
    try {
      const range = shape.textFrame.textRange;
      range.load('text');
      await context.sync();
      const t = range.text;
      return { text: typeof t === 'string' ? t : '', unreadable: false };
    } catch (e) {
      return { text: '', unreadable: true, error: String((e && e.message) || e).slice(0, 120) };
    }
  }

  // 表格文本：逐单元格独立 sync 读取，任何一格失败不影响其余（Table API 仅在部分宿主可用，整体失败返回空）
  async function tableTextSafe(context, shape) {
    try {
      const table = shape.table;
      if (!table) return '';
      table.load('rowCount,columnCount');
      await context.sync();
      const rows = Number(table.rowCount) || 0;
      const cols = Number(table.columnCount) || 0;
      const getCell = r => c => (typeof table.getCell === 'function' ? table.getCell(r, c) : table.cell(r, c));
      const lines = [];
      for (let r = 0; r < rows; r++) {
        const cells = [];
        for (let c = 0; c < cols; c++) {
          try {
            const cell = getCell(r)(c);
            cell.textFrame.textRange.load('text');
            await context.sync();
            cells.push(String(cell.textFrame.textRange.text || '').trim());
          } catch { cells.push(''); }
        }
        const rowText = cells.filter(Boolean).join(' | ');
        if (rowText) lines.push(rowText);
      }
      return lines.join('\n');
    } catch { return ''; }
  }

  // 表格单元格字体：模型要「统一缩小表格字号」之前，必须先知道表格现在用多大字号，
  // 否则只能瞎猜。逐单元格独立 try + sync（HOST_QUIRKS.batchSync：批中一个失败整批取消）。
  // 返回 null 表示整个 Table API 读不到字体——调用方必须用 unreadableStyle 明说
  // （「读不到」和「没有」是两回事，不能让模型把盲区当成空表格）。
  async function tableFontsSafe(context, shape) {
    try {
      const table = shape.table;
      if (!table) return null;
      table.load('rowCount,columnCount');
      await context.sync();
      const rows = Number(table.rowCount) || 0;
      const cols = Number(table.columnCount) || 0;
      const getCell = r => c => (typeof table.getCell === 'function' ? table.getCell(r, c) : table.cell(r, c));
      const cells = [];
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          try {
            const font = getCell(r)(c).textFrame.textRange.font;
            font.load('name,size,color');
            await context.sync();
            cells.push({ row: r, col: c, fontName: String(font.name || ''), fontSize: font.size, fontColor: String(font.color || '') });
          } catch { cells.push({ row: r, col: c, unreadable: true }); }
        }
      }
      return { rowCount: rows, columnCount: cols, cells };
    } catch { return null; }
  }

  // 递归读取形状文本：组合（Group）的子形状和表格（Table）的单元格在顶层 textFrame 上读不到，
  // 需要深入一层。逐形状/逐单元格独立 try，某个失败不影响其他，也不会把整页 sync 拖崩。
  async function shapeTextDeep(context, shape, depth = 0) {
    const parts = [];
    const direct = await shapeTextSafe(context, shape);
    if (direct) parts.push(direct);
    if (depth >= 3) return parts.join('\n'); // 防御异常嵌套，组合最多递归 3 层
    let type = '';
    try { shape.load('type'); await context.sync(); type = String(shape.type || ''); } catch {}
    if (type === 'Group' && shape.shapes) {
      try {
        shape.shapes.load('items');
        await context.sync();
        for (const child of (shape.shapes.items || [])) {
          const text = await shapeTextDeep(context, child, depth + 1).catch(() => '');
          if (text) parts.push(text);
        }
      } catch {}
    } else if (type === 'Table') {
      const text = await tableTextSafe(context, shape);
      if (text) parts.push(text);
    }
    return parts.join('\n');
  }

  function shapeText(shape) {
    try {
      const tf = shape.textFrame;
      if (tf && tf.textRange && typeof tf.textRange.text === 'string') return tf.textRange.text;
    } catch {}
    return '';
  }

  // 只组装真正读到的部分。unreadableStyle 明确区分「这个形状没有填充色」和「填充色我读不到」——
  // 前者可以放心不管，后者模型必须知道自己是盲的。
  function loadedShapeDetails(shape, ok, hasText) {
    const details = { text: '', styleReadable: ok.size > 0 };
    const grab = (key, fn) => { if (ok.has(key)) { try { fn(); } catch {} } };
    grab('fill', () => {
      details.fillType = String(shape.fill.type || '');
      details.fillColor = String(shape.fill.foregroundColor || '');
      details.fillTransparency = shape.fill.transparency;
    });
    grab('line', () => {
      details.lineVisible = shape.lineFormat.visible;
      details.lineColor = String(shape.lineFormat.color || '');
      details.lineTransparency = shape.lineFormat.transparency;
      details.lineWidth = shape.lineFormat.weight;
    });
    grab('frame', () => {
      details.marginLeft = shape.textFrame.leftMargin;
      details.marginRight = shape.textFrame.rightMargin;
      details.marginTop = shape.textFrame.topMargin;
      details.marginBottom = shape.textFrame.bottomMargin;
      details.verticalAlignment = String(shape.textFrame.verticalAlignment || '');
      details.wordWrap = shape.textFrame.wordWrap;
      details.autoSizeSetting = String(shape.textFrame.autoSizeSetting || '');
    });
    if (hasText) {
      grab('text', () => { details.text = String(shape.textFrame.textRange.text || ''); });
      grab('font', () => {
        const font = shape.textFrame.textRange.font;
        details.fontName = String(font.name || '');
        details.fontSize = font.size;
        details.fontColor = String(font.color || '');
        details.bold = font.bold;
      });
      grab('paragraph', () => {
        details.horizontalAlignment = String(shape.textFrame.textRange.paragraphFormat.horizontalAlignment || '');
      });
    }
    const missing = ['fill', 'line', 'frame'].concat(hasText ? ['text', 'font'] : []).filter(key => !ok.has(key));
    if (missing.length) details.unreadableStyle = missing;
    return details;
  }

  async function getPresentationOutline(args = {}) {
    requireOffice();
    const maxSlides = Math.max(1, Number(args.maxSlides || 100));
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const items = slides.items.slice(0, maxSlides);
      items.forEach(s => { s.load('id'); s.shapes.load('items'); });
      await context.sync();
      const out = [];
      for (let idx = 0; idx < items.length; idx++) {
        const s = items[idx];
        const shapes = (s.shapes.items || []).slice(0, 15); // 每页最多读 15 个形状，控制耗时
        const texts = [];
        for (const sh of shapes) {
          const t = await shapeTextDeep(context, sh); // 组合/表格内的文字也读出，否则对 AI 隐形
          if (t.trim()) texts.push(t.trim());
        }
        out.push({ index: idx, id: s.id, title: clampText(texts[0] || '', 120), texts: texts.map(x => clampText(x, 600)) });
      }
      return { success: true, slideCount: slides.items.length, returned: out.length, hasMore: slides.items.length > out.length, slides: out, note: 'slides[].texts 可能被截断（尾部有 … 表示截断）。不要把截断文本用作 propose_edits 的 find，必须先用 get_slide 读取完整原文。' };
    });
  }

  async function getSlideCore(index, includeStyles) {
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      slide.shapes.load('items');
      await context.sync();
      const page = await pageSizeForContext(context);
      const shapes = [];
      for (const sh of (slide.shapes.items || [])) {
        const fields = includeStyles ? ['id', 'name', 'type', 'left', 'top', 'width', 'height'] : ['id', 'name', 'left', 'top', 'width', 'height'];
        if (includeStyles && powerPointApiSupported('1.8')) fields.push('zOrderPosition');
        try { sh.load(fields.join(',')); } catch {}
      }
      await context.sync();
      // 【37/39-R1/40】占位框身份：placeholderFormat.type（Title/Body/…）+ containedType
      //（占位框包含的 ShapeType，空=null，PowerPointApi 1.8）让模型分辨母版空占位框
      // 与普通形状。无文字 ≠ 空：containedType 非 null（Image/Chart/Table…）说明有非文字
      // 内容，不得标空；containedType 读回 undefined（缺失）或读取失败记 unknown，不冒充空。
      // 前置过滤（40）：只对「shape.type 可读且为 Placeholder」的形状发占位请求——非占位框
      // 按官方契约访问 placeholderFormat 抛 GeneralException，不该制造这种请求；type 读不到
      //（旧包宿主整批取消后的降级路径）就保持未知，不猜。typeLoaded=样式路径已批量加载 type。
      const placeholderInfoOf = async (sh, typeLoaded) => {
        let shapeType = '';
        if (typeLoaded) {
          try { shapeType = String(sh.type || ''); } catch { return null; }
        } else {
          try { sh.load('type'); await context.sync(); shapeType = String(sh.type || ''); } catch { return null; }
        }
        if (shapeType !== 'Placeholder') return null;
        let pf;
        try { pf = sh.placeholderFormat; } catch { return null; }   // 防御：契约外宿主
        if (!pf || typeof pf.load !== 'function') return null;
        try {
          pf.load('type');
          await context.sync();
        } catch { return null; }   // sync 阶段失败（批量取消）受控降级为未知，不猜
        const type = String(pf.type || '');
        if (!type || type === 'None') return null;
        let content = 'unknown';
        if (powerPointApiSupported('1.8')) {
          try {
            pf.load('containedType');
            await context.sync();
            // 严格 null 才确认空：undefined（属性缺失/未写回）与读取失败都是 unknown
            content = pf.containedType === null ? 'confirmed-empty' : (pf.containedType === undefined ? 'unknown' : 'has-content');
          } catch { content = 'unknown'; }   // 读不到如实记 unknown，不反复制造 sync 错误
        }
        return { type, content };
      };
      await eachShapeIsolated(context, slide, async sh => {
        const placeholderInfo = await placeholderInfoOf(sh, includeStyles);
        if (!includeStyles) {
          const text = await shapeTextDeep(context, sh);
          const snapshot = { id: sh.id, name: sh.name, text, left: sh.left, top: sh.top, width: sh.width, height: sh.height, styleReadable: false, isEmpty: !String(text || '').trim() };
          if (placeholderInfo) { snapshot.placeholderType = placeholderInfo.type; snapshot.placeholderContent = placeholderInfo.content; }
          shapes.push(snapshot);
          return;
        }
        try {
          const style = await readShapeStyle(context, sh);
          const snapshot = { id: sh.id, name: sh.name, type: sh.type, left: sh.left, top: sh.top, width: sh.width, height: sh.height, ...loadedShapeDetails(sh, style.ok, style.hasText) };
          // 组合/表格的顶层 textFrame 读不到文字，递归读取子形状/单元格，避免内容对 AI 隐形
          if ((sh.type === 'Group' || sh.type === 'Table') && !String(snapshot.text || '').trim()) {
            snapshot.text = await shapeTextDeep(context, sh);
          }
          // 表格再补单元格字体：模型改表格字号前必须先知道当前字号，读不到就明说
          if (sh.type === 'Table') {
            const tableFonts = await tableFontsSafe(context, sh);
            if (tableFonts) snapshot.tableCellFonts = tableFonts;
            else snapshot.unreadableStyle = (snapshot.unreadableStyle || []).concat('tableCellFonts');
          }
          if (powerPointApiSupported('1.8')) snapshot.zOrderPosition = sh.zOrderPosition;
          if (placeholderInfo) { snapshot.placeholderType = placeholderInfo.type; snapshot.placeholderContent = placeholderInfo.content; }
          snapshot.isEmpty = !String(snapshot.text || '').trim();
          shapes.push(snapshot);
        } catch {
          // 该形状样式不可读：保留几何数据，文本走防御性深读
          const text = await shapeTextDeep(context, sh);
          const snapshot = { id: sh.id, name: sh.name, type: sh.type, text, left: sh.left, top: sh.top, width: sh.width, height: sh.height, styleReadable: false, isEmpty: !String(text || '').trim() };
          if (placeholderInfo) { snapshot.placeholderType = placeholderInfo.type; snapshot.placeholderContent = placeholderInfo.content; }
          shapes.push(snapshot);
        }
      });
      const result = { success: true, index, id: slide.id, capabilities: presentationCapabilities(), page, shapes, layoutAudit: buildLayoutAudit(shapes, page) };
      // 页面上有组合就直说本宿主够不到它的内部，省得模型为此试探好几轮才发现
      const groupCount = shapes.filter(s => String(s.type || '') === 'Group').length;
      if (groupCount) {
        result.groupCount = groupCount;
        result.groupLimitation = HOST_QUIRKS.groups;
      }
      return result;
    });
  }

  async function getSlide(args) {
    requireOffice();
    const index = Number(args.index);
    if (powerPointApiSupported('1.4')) {
      try { return await getSlideCore(index, true); }
      catch (e) {
        // 带样式读取整批失败（如旧包宿主 load type 字段整批取消）：降级到几何+文字读取，
        // 但必须把「样式读不到」显式写进结果——模型看到 stylesDegraded 才知道样式是被宿主
        // 拒绝的，不会把盲区当成「形状本来就没样式」
        const fallback = await getSlideCore(index, false);
        if (fallback && typeof fallback === 'object' && fallback.success) {
          fallback.stylesDegraded = true;
          fallback.stylesDegradedReason = String((e && e.message) || e).slice(0, 300);
        }
        return fallback;
      }
    }
    const result = await getSlideCore(index, false);
    if (result && typeof result === 'object' && result.success) {
      result.stylesDegraded = true;
      result.stylesDegradedReason = 'PowerPointApi 1.4 unavailable on this host: shape type/style fields were not requested';
    }
    return result;
  }

  async function getSlidePreview(args = {}) {
    requireOffice();
    if (!powerPointApiSupported('1.8')) {
      throw new Error('当前 PowerPoint 不支持页面预览 API 1.8。仍可使用 get_slide 返回的版式检查结果。');
    }
    const index = Number(args.index);
    const height = Math.max(240, Math.min(900, Number(args.height || 540)));
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      const image = slide.getImageAsBase64({ height });
      await context.sync();
      const imageBase64 = String(image.value || '').replace(/^data:image\/png;base64,/i, '');
      if (!imageBase64) throw new Error('PowerPoint 没有返回页面预览图。');
      return {
        success: true,
        index,
        id: slide.id,
        mimeType: 'image/png',
        requestedHeight: height,
        imageBase64,
        note: '这是当前幻灯片的整页预览图，用于判断视觉层级、平衡、留白和风格一致性。'
      };
    });
  }

  async function getSelectedSlides() {
    requireOffice();
    return new Promise(resolve => {
      try {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, asyncResult => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const slides = (asyncResult.value && asyncResult.value.slides) || [];
            // Office SlideRange 的 slides[].index 是 1-based，本插件所有工具统一按 0-based slides.items[index]，这里减 1 对齐
            resolve({ success: true, slides: slides.map(s => ({ id: s.id, title: s.title, index: Number(s.index) - 1 })) });
          } else {
            resolve({ success: false, error: asyncResult.error ? asyncResult.error.message : 'getSelectedDataAsync failed' });
          }
        });
      } catch (e) { resolve({ success: false, error: e.message || String(e) }); }
    });
  }

  const HEX_RE = /^#[0-9A-Fa-f]{6}$/;
  function normColor(value) {
    const text = String(value == null ? '' : value).trim().toUpperCase();
    return HEX_RE.test(text) ? text : '';
  }

  // 换肤/风格迁移的专用工具。以前这件事只能让模型手写 eval 遍历，每写一次都要重踩一遍 HOST_QUIRKS，
  // 还容易漏掉同色系的另一个色值（改了 #F00018、漏了 #FF0000）或误删自己刚加的形状。
  // dryRun 先返回本页实际用到的完整调色板和字体清单，看准了再写入。
  async function recolorSlide(args = {}) {
    requireOffice();
    const index = Number(args.index);
    const dryRun = Boolean(args.dryRun);
    const colorMap = {};
    for (const key of Object.keys(args.colorMap || {})) {
      const from = normColor(key);
      const to = normColor(args.colorMap[key]);
      if (from && to) colorMap[from] = to;
    }
    const fontMap = {};
    for (const key of Object.keys(args.fontMap || {})) {
      if (key && args.fontMap[key]) fontMap[String(key)] = String(args.fontMap[key]);
    }
    const fontName = args.fontName ? String(args.fontName) : '';
    if (!dryRun && !Object.keys(colorMap).length && !Object.keys(fontMap).length && !fontName) {
      throw new Error('recolor_slide 至少需要 colorMap / fontMap / fontName 之一；只想看这一页用了哪些颜色和字体，传 dryRun:true。');
    }
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      slide.shapes.load('items');
      await context.sync();

      const palette = new Map();
      const fonts = new Map();
      const fontSizes = new Map();
      const changes = [];
      const unreadable = [];
      let groupCount = 0;
      const bump = (map, key) => { if (key) map.set(key, (map.get(key) || 0) + 1); };

      // 表格（Table 形状）顶层 textFrame 读不到，单元格字体要逐格迁移，映射规则与普通形状一致。
      // 逐单元格独立 try + sync（HOST_QUIRKS.batchSync）：一格失败不拖垮整表；
      // 表格整体读不到时静默跳过并交给 unreadableShapes 计数，不谎报成「表格里没有颜色」。
      const recolorTableCells = async (shape, label) => {
        let table;
        try {
          table = shape.table;
          if (!table) return;
          table.load('rowCount,columnCount');
          await context.sync();
        } catch { return; }
        const getCell = r => c => (typeof table.getCell === 'function' ? table.getCell(r, c) : table.cell(r, c));
        for (let r = 0; r < (Number(table.rowCount) || 0); r++) {
          for (let c = 0; c < (Number(table.columnCount) || 0); c++) {
            try {
              const font = getCell(r)(c).textFrame.textRange.font;
              font.load('name,size,color');
              await context.sync();
              const cellLabel = `${label}[${r},${c}]`;
              const currentFont = String(font.name || '');
              const fontColor = normColor(font.color);
              bump(fonts, currentFont);
              bump(fontSizes, font.size);
              bump(palette, fontColor);
              if (fontColor && colorMap[fontColor]) {
                if (!dryRun) font.color = colorMap[fontColor];
                changes.push({ shape: cellLabel, field: 'fontColor', from: fontColor, to: colorMap[fontColor] });
              }
              const nextFont = fontMap[currentFont] || (fontName && currentFont !== fontName ? fontName : '');
              if (nextFont) {
                if (!dryRun) font.name = nextFont;
                changes.push({ shape: cellLabel, field: 'font', from: currentFont, to: nextFont });
              }
              if (!dryRun) await context.sync();
            } catch { /* 单个单元格读/写失败不影响其余 */ }
          }
        }
      };

      const walk = await eachShapeIsolated(context, slide, async shape => {
        try { shape.load('id,name,type'); await context.sync(); } catch {}
        const label = shapeLabel(shape);
        if (String(shape.type || '') === 'Group') groupCount += 1;
        if (String(shape.type || '') === 'Table') await recolorTableCells(shape, label);
        const style = await readShapeStyle(context, shape);
        if (style.ok.size === 0) { unreadable.push(label); return; }

        const fillColor = style.ok.has('fill') ? normColor(shape.fill.foregroundColor) : '';
        bump(palette, fillColor);
        if (fillColor && colorMap[fillColor]) {
          if (!dryRun) {
            try { shape.fill.setSolidColor(colorMap[fillColor]); }
            catch { shape.fill.foregroundColor = colorMap[fillColor]; }
          }
          changes.push({ shape: label, field: 'fill', from: fillColor, to: colorMap[fillColor] });
        }

        const lineColor = style.ok.has('line') ? normColor(shape.lineFormat.color) : '';
        bump(palette, lineColor);
        if (lineColor && colorMap[lineColor] && shape.lineFormat.visible !== false) {
          if (!dryRun) shape.lineFormat.color = colorMap[lineColor];
          changes.push({ shape: label, field: 'line', from: lineColor, to: colorMap[lineColor] });
        }

        if (style.hasText && style.ok.has('font')) {
          const font = shape.textFrame.textRange.font;
          const currentFont = String(font.name || '');
          const fontColor = normColor(font.color);
          bump(fonts, currentFont);
          bump(fontSizes, font.size);
          bump(palette, fontColor);
          if (fontColor && colorMap[fontColor]) {
            if (!dryRun) font.color = colorMap[fontColor];
            changes.push({ shape: label, field: 'fontColor', from: fontColor, to: colorMap[fontColor] });
          }
          const nextFont = fontMap[currentFont] || (fontName && currentFont !== fontName ? fontName : '');
          if (nextFont) {
            if (!dryRun) font.name = nextFont;
            changes.push({ shape: label, field: 'font', from: currentFont, to: nextFont });
          }
        }
        if (!dryRun) await context.sync();
      });

      const asList = map => [...map.entries()].sort((a, b) => b[1] - a[1]).map(entry => ({ value: entry[0], count: entry[1] }));
      const result = {
        success: true, index, dryRun,
        palette: asList(palette), fonts: asList(fonts), fontSizes: asList(fontSizes),
        changed: changes.length, changes: changes.slice(0, 60),
        unreadableShapes: unreadable.length + walk.failures.length,
        unreadableShapeNames: unreadable.slice(0, 20),
        _navTarget: { slideId: slide.id }
      };
      if (groupCount) {
        result.groupCount = groupCount;
        result.groupLimitation = HOST_QUIRKS.groups;
      }
      if (dryRun) result.hint = '这是预演，没有写入任何修改。palette 是本页实际用到的全部颜色——同一个色系可能有好几个色值，补全 colorMap 后去掉 dryRun 再调用一次。';
      return result;
    });
  }

  // ---- T1 工具：表格样式 / 多形状对齐 / 图片替换 / 图表读取 ----

  // 四个工具共用的定位骨架：按 index 找页、按 shapeId 找形状。
  // 批量只 load 身份/几何字段（所有形状类型都支持），样式类字段一律逐形状独立退避（HOST_QUIRKS.batchSync）。
  async function locateShapeOnSlide(context, index, shapeId, fields) {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    const slide = slides.items[index];
    if (!slide) throw new Error(`Slide ${index} not found`);
    slide.load('id');
    slide.shapes.load('items');
    await context.sync();
    slide.shapes.items.forEach(sh => sh.load(fields || 'id,name,type'));
    await context.sync();
    const shape = slide.shapes.items.find(sh => sh.id === shapeId);
    return { slide, shape };
  }

  // 表格单元格样式：字体/字号/字色/加粗/底色，支持整表或按行/列/单元格限定范围。
  // scaleFontSize 解决「统一缩小表格字号」这个最高频需求：先读各单元格现有字号再等比缩放，
  // 不写死一个值（写死会把原本的字号层级——表头大、数据小——直接抹平）。
  async function setTableStyle(args = {}) {
    requireOffice();
    const target = args.target || {};
    const index = Number(target.index);
    const shapeId = target.shapeId ? String(target.shapeId) : '';
    if (!Number.isInteger(index) || index < 0) throw new Error('set_table_style 需要 target.index（0 起始页码）和 target.shapeId——先用 get_slide 读该页，找到 type 为 Table 的形状 id。');
    if (!shapeId) throw new Error('set_table_style 缺少 target.shapeId——先用 get_slide 读该页，找到 type 为 Table 的形状 id。');
    const fontName = args.fontName ? String(args.fontName).slice(0, 120) : '';
    const fontSize = args.fontSize != null ? finiteNumber(args.fontSize, 'fontSize', { positive: true }) : null;
    const fontColor = args.fontColor ? normalizedColor(args.fontColor, 'fontColor') : '';
    const hasBold = hasOwn(args, 'bold');
    const cellFill = args.cellFill ? normalizedColor(args.cellFill, 'cellFill') : '';
    const scale = args.scaleFontSize != null ? Number(args.scaleFontSize) : null;
    if (scale != null && (!Number.isFinite(scale) || scale <= 0 || scale > 2)) {
      throw new Error('scaleFontSize 必须是 0~2 之间的正数（0.85 = 各单元格现有字号 ×0.85）。它和 fontSize 的区别：fontSize 是写死值，scaleFontSize 是等比缩放、保留原字号层级。');
    }
    if (!fontName && fontSize == null && !fontColor && !hasBold && !cellFill && scale == null) {
      throw new Error('set_table_style 没有任何样式参数。支持：fontName / fontSize / fontColor / bold / cellFill（单元格底色）/ scaleFontSize（等比缩放现有字号，如 0.85）。范围限定可选 rows / cols / cells，缺省为整表。');
    }
    const rowSet = new Set((Array.isArray(args.rows) ? args.rows : []).map(Number).filter(Number.isInteger));
    const colSet = new Set((Array.isArray(args.cols) ? args.cols : []).map(Number).filter(Number.isInteger));
    const cellSet = new Set((Array.isArray(args.cells) ? args.cells : []).filter(pair => Array.isArray(pair) && pair.length >= 2).map(pair => `${Number(pair[0])},${Number(pair[1])}`));
    const scoped = rowSet.size > 0 || colSet.size > 0 || cellSet.size > 0;
    const inScope = (r, c) => !scoped || rowSet.has(r) || colSet.has(c) || cellSet.has(`${r},${c}`);

    return PowerPoint.run(async context => {
      const { slide, shape } = await locateShapeOnSlide(context, index, shapeId, 'id,name,type');
      if (!shape) throw new Error(`第 ${index} 页找不到形状 ${shapeId}。先用 get_slide 重读该页拿到当前形状列表（形状可能已被移动或删除）。`);
      if (String(shape.type || '') !== 'Table') {
        throw new Error(`形状 ${shapeId} 的类型是 ${shape.type}，不是表格，set_table_style 改不了它。普通形状的字体/填充用 apply_layout(updateShape)；整页配色/字体迁移用 recolor_slide；表格整体位置尺寸用 apply_layout 改 left/top/width/height。`);
      }
      let table;
      try {
        table = shape.table;
        if (!table) throw new Error('shape.table missing');
        table.load('rowCount,columnCount');
        await context.sync();
      } catch {
        throw new Error('本机读不到这个表格的 Table API（shape.table 不可用）。仍然能改的：整页字体/配色迁移用 recolor_slide，表格整体位置尺寸用 apply_layout；单元格级修改在这台机器上走不通，不要反复试。');
      }
      const rows = Number(table.rowCount) || 0;
      const cols = Number(table.columnCount) || 0;
      // load 成功不等于数据可信：部分宿主 Table API 半身不遂，rowCount/columnCount 读回 0 或 undefined。
      // 这时一格都改不到，绝不允许 success:true changed:0 静默通过——返回失败并给替代路径。
      if (rows <= 0 || cols <= 0) {
        return {
          success: false,
          error: '表格行列数读回为空（Table API 部分不可用），本次没有改动任何单元格。替代路径：recolor_slide 整页迁移或 apply_layout 调整整体几何。',
          dimensionsUnreadable: true
        };
      }
      const getCell = r => c => (typeof table.getCell === 'function' ? table.getCell(r, c) : table.cell(r, c));
      let changed = 0;
      let fillWritten = 0;
      let fillUnsupported = 0;
      let scaleSkipped = 0;
      const unreadableCells = [];
      // 逐单元格独立 try + sync（HOST_QUIRKS.batchSync）：一格失败只记它自己，不拖垮整表
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          if (!inScope(r, c)) continue;
          try {
            const cell = getCell(r)(c);
            const font = cell.textFrame.textRange.font;
            font.load('name,size,color,bold');
            await context.sync();
            if (fontName) font.name = fontName;
            if (fontSize != null) font.size = fontSize;
            if (fontColor) font.color = fontColor;
            if (hasBold) font.bold = Boolean(args.bold);
            if (scale != null) {
              const current = Number(font.size);
              // 混排字号的单元格 font.size 可能读回 null：跳过缩放并计数，不谎报
              if (Number.isFinite(current) && current > 0) font.size = Math.max(1, Math.round(current * scale * 10) / 10);
              else scaleSkipped += 1;
            }
            if (cellFill) {
              // TableCell.fill 是否可写按宿主实测探测，不写死（HOST_QUIRKS.tableCellFill）
              if (cell.fill && typeof cell.fill.setSolidColor === 'function') {
                cell.fill.setSolidColor(cellFill);
                noteRuntimeProbe('tableCellFill', true);
                fillWritten += 1;
              } else {
                noteRuntimeProbe('tableCellFill', false);
                fillUnsupported += 1;
              }
            }
            await context.sync();
            changed += 1;
          } catch {
            unreadableCells.push([r, c]);
          }
        }
      }
      const result = {
        success: true, index, shapeId,
        changed, unreadable: unreadableCells.length, unreadableCells: unreadableCells.slice(0, 20),
        _navTarget: { slideId: slide.id }
      };
      if (scaleSkipped) result.scaleSkipped = scaleSkipped;
      if (cellFill) {
        result.cellFill = { written: fillWritten, unsupported: fillUnsupported };
        if (fillUnsupported) {
          result.cellFillHint = '本机写不了单元格底色（TableCell.fill 不可用），底色没有生效，其余样式已正常应用。替代路径：整页配色迁移用 recolor_slide；或在表格下方垫一个同色矩形（apply_layout addShape + zOrder sendToBack）。';
        }
      }
      if (unreadableCells.length) result.note = `${unreadableCells.length} 个单元格读不到或写入失败，已跳过，其余单元格正常。`;
      return result;
    });
  }

  // 新建表格：PPT Office.js 没有原生插入表格 API（TOOL-COVERAGE 死区），
  // 用「每格一个矩形（边框+底色）+ 一个文本框」拼出视觉表格。产出是普通形状，不是 Table——
  // 后续改单元格文字用 set_text，调样式用 apply_layout / recolor_slide；set_table_style 不适用。
  // 逐形状独立 try（HOST_QUIRKS.batchSync）：一格失败只记它自己，不拖垮整表。12×12 上限防失控。
  const TABLE_GRID_MAX_DIM = 12;
  async function addTableGrid(args = {}) {
    requireOffice();
    const index = Number(args.index);
    if (!Number.isInteger(index) || index < 0) throw new Error('add_table_grid 需要 index（0 起始页码）。');
    const data = Array.isArray(args.data)
      ? args.data.map(row => Array.isArray(row) ? row.map(cell => String(cell == null ? '' : cell)) : [String(row == null ? '' : row)])
      : [];
    if (!data.length) throw new Error('add_table_grid 需要 data（二维数组文本，第一行是表头）。例：[["项目","数值"],["收入","1.2亿"]]');
    const dataCols = Math.max(...data.map(row => row.length));
    const rows = args.rows != null ? Number(args.rows) : data.length;
    const columns = args.columns != null ? Number(args.columns) : dataCols;
    if (!Number.isInteger(rows) || rows < 1 || rows > TABLE_GRID_MAX_DIM || !Number.isInteger(columns) || columns < 1 || columns > TABLE_GRID_MAX_DIM) {
      throw new Error(`add_table_grid 最多 ${TABLE_GRID_MAX_DIM}×${TABLE_GRID_MAX_DIM}（当前 rows=${rows}, columns=${columns}）——形状拼的表，格子再多页面和性能都撑不住。更大的表请拆到多页，或请用户手动插入原生表格后用 set_table_style 调样式。`);
    }
    if (data.length > rows || dataCols > columns) {
      throw new Error(`data 是 ${data.length} 行 ${dataCols} 列，超出 rows×columns（${rows}×${columns}）。把 rows/columns 调大，或删掉多余数据。`);
    }
    const left = finiteNumber(args.left, 'left');
    const top = finiteNumber(args.top, 'top');
    const width = finiteNumber(args.width, 'width', { positive: true });
    const height = finiteNumber(args.height, 'height', { positive: true });
    const headerFill = args.headerFill ? normalizedColor(args.headerFill, 'headerFill') : '';
    const cellFill = args.cellFill ? normalizedColor(args.cellFill, 'cellFill') : ''; // 非表头格底色（可空）
    const borderColor = args.borderColor ? normalizedColor(args.borderColor, 'borderColor') : '#BFBFBF';
    const fontName = args.fontName ? String(args.fontName).slice(0, 120) : '';
    const fontSize = args.fontSize != null ? finiteNumber(args.fontSize, 'fontSize', { positive: true }) : null;

    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      if (!slide.shapes || typeof slide.shapes.addGeometricShape !== 'function' || typeof slide.shapes.addTextBox !== 'function') {
        throw new Error('本机连 addGeometricShape/addTextBox 都没有，拼表也拼不出来。替代路径：请用户在 PowerPoint 里「插入 → 表格」手动建表，之后用 set_table_style 调样式、edit_table_structure 探测结构操作。');
      }
      const cellW = width / columns;
      const cellH = height / rows;
      const round2 = v => Math.round(v * 100) / 100;
      let cellsCreated = 0;
      const failed = [];
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < columns; c++) {
          const geometry = { left: round2(left + c * cellW), top: round2(top + r * cellH), width: round2(cellW), height: round2(cellH) };
          try {
            const rect = slide.shapes.addGeometricShape('Rectangle', geometry);
            // 底色：表头行优先 headerFill，其余格用 cellFill；都没给就显式 fill.clear()——
            // 不写 fill 不是「透明」，会露出宿主的主题默认填充色，拼出来的表直接花掉
            const cellBackground = (r === 0 && headerFill) ? headerFill : cellFill;
            if (cellBackground) rect.fill.setSolidColor(cellBackground);
            else if (rect.fill && typeof rect.fill.clear === 'function') rect.fill.clear();
            rect.lineFormat.color = borderColor;
            const box = slide.shapes.addTextBox((data[r] && data[r][c]) || '', geometry);
            if (fontName) box.textFrame.textRange.font.name = fontName;
            if (fontSize != null) box.textFrame.textRange.font.size = fontSize;
            if (r === 0) box.textFrame.textRange.font.bold = true; // 表头行加粗
            await context.sync();
            cellsCreated += 1;
          } catch (e) {
            failed.push({ cell: [r, c], error: e && e.message ? e.message : String(e) });
          }
        }
      }
      const result = {
        success: cellsCreated > 0, index, rows, columns,
        cellsPlanned: rows * columns, cellsCreated, shapesCreated: cellsCreated * 2,
        failed: failed.length, failedCells: failed.slice(0, 10),
        note: '这是用矩形+文本框拼的视觉表格，不是原生 Table（Office.js 没有插入表格 API）。后续：单元格文字用 set_text 逐格改，样式用 apply_layout / recolor_slide；set_table_style 只适用原生表格。',
        _navTarget: { slideId: slide.id }
      };
      if (failed.length) result.note = `${failed.length} 个单元格创建失败，已跳过，其余正常。` + result.note;
      return result;
    });
  }

  const ALIGN_ACTIONS = new Set(['left', 'center', 'right', 'top', 'middle', 'bottom',
    'alignLeft', 'alignCenter', 'alignRight', 'alignTop', 'alignMiddle', 'alignBottom',
    'distributeH', 'distributeV']);
  const ALIGN_ACTION_ALIASES = {
    alignLeft: 'left', alignCenter: 'center', alignRight: 'right',
    alignTop: 'top', alignMiddle: 'middle', alignBottom: 'bottom'
  };

  // 多形状对齐/等距分布：Office.js 没有现成对齐 API，按形状几何自己算，纯几何移动不动样式。
  // 基准：reference 'first'（默认，shapeIds 里的第一个形状）或 'slide'（页面边界）。
  async function alignShapes(args = {}) {
    requireOffice();
    const target = args.target || {};
    const index = Number(target.index);
    if (!Number.isInteger(index) || index < 0) throw new Error('align_shapes 需要 target.index（0 起始页码）。先用 get_slide 读该页拿到形状 id 列表。');
    const rawAction = String(args.action || '');
    const action = ALIGN_ACTION_ALIASES[rawAction] || rawAction;
    if (!ALIGN_ACTIONS.has(action)) {
      throw new Error(`align_shapes 不支持 action "${rawAction}"。支持：left/center/right/top/middle/bottom（对齐，也接受 alignLeft 等别名）和 distributeH/distributeV（等距分布，至少 3 个形状）。`);
    }
    const shapeIds = (Array.isArray(args.shapeIds) ? args.shapeIds : []).map(String).filter(Boolean);
    const distribute = action === 'distributeH' || action === 'distributeV';
    const need = distribute ? 3 : 2;
    if (shapeIds.length < need) {
      throw new Error(`align_shapes 的 ${action} 至少需要 ${need} 个 shapeId（当前 ${shapeIds.length} 个）。先用 get_slide 读该页形状列表拿 id。`);
    }
    const reference = String(args.reference || 'first') === 'slide' ? 'slide' : 'first';

    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      slide.shapes.load('items');
      await context.sync();
      slide.shapes.items.forEach(sh => sh.load('id,left,top,width,height'));
      await context.sync();
      const byId = new Map(slide.shapes.items.map(sh => [sh.id, sh]));
      const missing = shapeIds.filter(id => !byId.has(id));
      if (missing.length) {
        const available = slide.shapes.items.map(sh => sh.id).slice(0, 20).join(', ');
        throw new Error(`第 ${index} 页找不到这些形状：${missing.join(', ')}。本页现有形状 id：${available}。先用 get_slide 重读再试。`);
      }
      // 按用户给的顺序取形状：reference 'first' 时 shapeIds 的第一个就是基准
      const targets = shapeIds.map(id => byId.get(id)).map(sh => ({
        shape: sh, left: Number(sh.left), top: Number(sh.top), width: Number(sh.width), height: Number(sh.height)
      }));
      const page = reference === 'slide' ? await pageSizeForContext(context) : null;
      const anchor = reference === 'slide'
        ? { left: 0, top: 0, width: page.width, height: page.height }
        : targets[0];
      const updates = targets.map(t => ({ shape: t.shape, left: t.left, top: t.top }));
      const horizontal = action === 'left' || action === 'center' || action === 'right' || action === 'distributeH';
      if (!distribute) {
        updates.forEach((u, i) => {
          const t = targets[i];
          if (action === 'left') u.left = anchor.left;
          else if (action === 'center') u.left = anchor.left + (anchor.width - t.width) / 2;
          else if (action === 'right') u.left = anchor.left + anchor.width - t.width;
          else if (action === 'top') u.top = anchor.top;
          else if (action === 'middle') u.top = anchor.top + (anchor.height - t.height) / 2;
          else if (action === 'bottom') u.top = anchor.top + anchor.height - t.height;
        });
      } else {
        // 分布：按主轴位置排序。reference first = 保持首尾不动、中间均分间距；
        // slide = 在页面范围内均分（含两端留白，经典「页面内等距排列」）
        const axis = horizontal ? 'left' : 'top';
        const sizeKey = horizontal ? 'width' : 'height';
        const order = targets.map((t, i) => i).sort((a, b) => targets[a][axis] - targets[b][axis]);
        const totalSize = targets.reduce((sum, t) => sum + t[sizeKey], 0);
        let cursor;
        let gap;
        if (reference === 'slide') {
          const span = horizontal ? page.width : page.height;
          gap = (span - totalSize) / (targets.length + 1);
          cursor = gap;
        } else {
          const first = targets[order[0]];
          const last = targets[order[order.length - 1]];
          gap = (last[axis] + last[sizeKey] - first[axis] - totalSize) / (targets.length - 1);
          cursor = first[axis];
        }
        for (const i of order) {
          updates[i][axis] = cursor;
          cursor += targets[i][sizeKey] + gap;
        }
      }
      // 逐形状独立 try + sync：一个形状写不动不拖垮其余
      let moved = 0;
      const failed = [];
      for (const u of updates) {
        try {
          u.shape.left = Math.round(u.left * 100) / 100;
          u.shape.top = Math.round(u.top * 100) / 100;
          await context.sync();
          moved += 1;
        } catch (e) {
          failed.push({ shape: shapeLabel(u.shape), error: e && e.message ? e.message : String(e) });
        }
      }
      return {
        success: true, index, action, reference, moved, failed: failed.length, failedShapes: failed.slice(0, 10),
        note: `纯几何移动：${action}，基准=${reference === 'slide' ? '页面边界' : 'shapeIds 里的第一个形状'}，样式未动。完成后建议 get_slide_preview 复检视觉结果。`,
        _navTarget: { slideId: slide.id }
      };
    });
  }

  // 图片替换：Office.js 没有标准的图片原位替换 API，先探测（不写死），
  // 没有就降级为「同位置删除旧图 + addImage 插入新图」，位置尺寸保持原样（HOST_QUIRKS.imageReplace）。
  async function replaceImage(args = {}) {
    requireOffice();
    const target = args.target || {};
    const index = Number(target.index);
    const shapeId = target.shapeId ? String(target.shapeId) : '';
    if (!Number.isInteger(index) || index < 0 || !shapeId) throw new Error('replace_image 需要 target.index 和 target.shapeId——先用 get_slide 读该页，找到图片形状的 id。');
    const base64 = String(args.base64 || '').replace(/^data:image\/[a-zA-Z+]+;base64,/i, '').trim();
    if (!base64) throw new Error('replace_image 缺少 base64 图片数据（不带 data: 前缀的 base64 字符串）。');
    return PowerPoint.run(async context => {
      const { slide, shape } = await locateShapeOnSlide(context, index, shapeId, 'id,name,type,left,top,width,height');
      if (!shape) throw new Error(`第 ${index} 页找不到形状 ${shapeId}。先用 get_slide 重读该页拿到当前形状列表。`);
      if (String(shape.type || '') !== 'Picture') {
        throw new Error(`形状 ${shapeId} 的类型是 ${shape.type}，不是图片。replace_image 只替换图片形状：换文字用 set_text，换表格样式用 set_table_style，换图表数据本机不支持（get_chart 只能读）。`);
      }
      const geometry = { left: Number(shape.left), top: Number(shape.top), width: Number(shape.width), height: Number(shape.height) };
      // 路径 A（探测）：宿主若有原位替换 API 就用它，能保留叠放次序和裁剪
      if (shape.image && typeof shape.image.replace === 'function') {
        shape.image.replace(base64);
        await context.sync();
        noteRuntimeProbe('imageReplaceInPlace', true);
        return { success: true, index, shapeId, path: 'inPlace', _navTarget: { slideId: slide.id } };
      }
      noteRuntimeProbe('imageReplaceInPlace', false);
      // 路径 B（降级）：同位置删除旧图 + 插入新图。新图在顶层，需要原层级时用 apply_layout 的 zOrder 修正。
      if (!slide.shapes || typeof slide.shapes.addImage !== 'function') {
        throw new Error('本机既没有图片原位替换 API，也没有 shapes.addImage，程序化替换图片走不通。替代路径：insert_image 在当前选区插入新图 → 用 apply_layout(updateShape) 按 get_slide 读到的原图几何摆位 → apply_layout(deleteShape) 删旧图。');
      }
      shape.delete();
      const picture = slide.shapes.addImage(base64, geometry);
      picture.load('id');
      await context.sync();
      return {
        success: true, index, shapeId, newShapeId: picture.id, path: 'deleteInsert', geometry,
        note: '本机没有图片原位替换 API，已在原位置删除旧图并插入新图（位置和尺寸不变）。新图现在位于顶层：如果它压住了不该压的内容，用 apply_layout 的 zOrder（如 sendBackward）调整。',
        _navTarget: { slideId: slide.id }
      };
    });
  }

  // 图表数据读取：Chart API 薄，chartType 和系列名一般能读，数值/类别很多宿主读不到。
  // 逐字段独立退避，读不到的部分用 unreadableStyle / valuesUnreadable 明说——「读不到」和「没有」是两回事。
  async function getChart(args = {}) {
    requireOffice();
    const target = args.target || {};
    const index = Number(target.index);
    const shapeId = target.shapeId ? String(target.shapeId) : '';
    if (!Number.isInteger(index) || index < 0 || !shapeId) throw new Error('get_chart 需要 target.index 和 target.shapeId——先用 get_slide 读该页，找到 type 为 Chart 的形状 id。');
    return PowerPoint.run(async context => {
      const { slide, shape } = await locateShapeOnSlide(context, index, shapeId, 'id,name,type');
      if (!shape) throw new Error(`第 ${index} 页找不到形状 ${shapeId}。先用 get_slide 重读该页拿到当前形状列表。`);
      const type = String(shape.type || '');
      if (type !== 'Chart') {
        const extra = type === 'Picture'
          ? '这个「图表」其实是一张贴图，没有可读数据，只能 get_slide_preview 目测。'
          : '如果页面上有图表，先用 get_slide 找 type 为 Chart 的形状。';
        throw new Error(`形状 ${shapeId} 的类型是 ${type}，不是图表。${extra}`);
      }
      let chart;
      try {
        chart = shape.chart;
        if (!chart) throw new Error('shape.chart missing');
      } catch {
        throw new Error('本机 Chart API 不可用（shape.chart 不存在）。替代路径：get_slide_preview 看图目测数据，或请用户在 PowerPoint 里右键图表 →「编辑数据」，从打开的 Excel 表里读取。');
      }
      const unreadable = [];
      let chartType = '';
      try {
        chart.load('chartType');
        await context.sync();
        chartType = String(chart.chartType || '');
      } catch { unreadable.push('chartType'); }
      const series = [];
      let seriesOk = true;
      try {
        chart.series.load('items');
        await context.sync();
      } catch { seriesOk = false; unreadable.push('series'); }
      if (seriesOk) {
        for (const s of (chart.series.items || [])) {
          const entry = {};
          try { s.load('name'); await context.sync(); entry.name = String(s.name || ''); }
          catch { entry.nameUnreadable = true; }
          // 系列数值：逐系列独立试，读不到就明说（实测写回 runtimeProbe，不写死）
          try {
            s.load('values');
            await context.sync();
            if (Array.isArray(s.values)) { entry.values = s.values.slice(); noteRuntimeProbe('chartSeriesValues', true); }
            else { entry.valuesUnreadable = true; noteRuntimeProbe('chartSeriesValues', false); }
          } catch {
            entry.valuesUnreadable = true;
            noteRuntimeProbe('chartSeriesValues', false);
          }
          series.push(entry);
        }
      }
      const result = { success: true, index, shapeId, chartType, seriesCount: series.length, series };
      if (unreadable.length) result.unreadableStyle = unreadable;
      if (series.some(s => s.valuesUnreadable)) {
        result.note = '本机读不到系列的数值/类别（Chart API 只有图表类型和系列名可读）。这是「读不到」，不是「图表没有数据」。需要数值时：get_slide_preview 目测，或请用户右键图表 →「编辑数据」打开 Excel 表。';
      }
      return result;
    });
  }

  // ---- T2 工具：样式模板 / 页面背景 / 表格结构 / 图片裁剪与透明度 / 动手前问卷 ----

  // 样式模板的持久化：主存储是插件共享设置（App.state.settings.pptStyleTemplates + App.saveSettings()），
  // 跨文档可用——换肤场景里客户每次给的都是新文件，模板「跟着旧文件走」等于每次重新调研，
  // 放共享设置才能「存一次、次次套用」。
  // document.settings 只做向后兼容读取：旧版本存在文档里的模板仍然读得到，
  // 并在下一次 save_style_template 写入时随新数据一并迁进共享设置（一次性迁移，新写入不再进文档）。
  function styleTemplatesKey() { return (App.STORAGE_KEYS && App.STORAGE_KEYS.styleTemplates) || 'office-ai-trial-ppt-style-templates-v1'; }

  function loadStyleTemplates() {
    const settings = (App.state && App.state.settings) || {};
    const fromSettings = (settings.pptStyleTemplates && typeof settings.pptStyleTemplates === 'object' && !Array.isArray(settings.pptStyleTemplates)) ? settings.pptStyleTemplates : {};
    let legacy = {};
    try { legacy = JSON.parse(App.loadDocSetting(styleTemplatesKey(), '{}') || '{}') || {}; } catch { legacy = {}; }
    // 共享设置优先；文档里的旧数据并入（同名以共享设置为准）
    return Object.assign({}, legacy, fromSettings);
  }

  // 把当前 deck 的设计标记（配色/字体/字号层级）存成命名模板。
  // 提取直接复用 recolor_slide 的 dryRun——逐形状退避、表格单元格字体这些读取坑它都解决了，
  // 再写一遍扫描逻辑只会重新踩坑。背景色 Office.js 读不到（没有 slide.background 读 API），存 null 明说。
  async function saveStyleTemplate(args = {}) {
    requireOffice();
    const name = String(args.name || '').trim().slice(0, 60);
    if (!name) throw new Error('save_style_template 需要 name（模板名，如「品牌模板」）。已存模板用 list_style_templates 查看，同名会覆盖。');
    const rawIndexes = Array.isArray(args.indexes) ? args.indexes : [];
    const picked = [...new Set(rawIndexes.map(Number).filter(i => Number.isInteger(i) && i >= 0))].slice(0, 5);
    const targets = picked.length ? picked : [0, 1, 2];
    const scanned = [];
    const failed = [];
    for (const index of targets) {
      try {
        const scan = await recolorSlide({ index, dryRun: true });
        scanned.push({ index, palette: scan.palette || [], fonts: scan.fonts || [], fontSizes: scan.fontSizes || [] });
      } catch (e) {
        failed.push({ index, error: e && e.message ? e.message : String(e) });
      }
    }
    if (!scanned.length) {
      throw new Error(`代表页（${targets.join('/')}）都读不到：${failed[0] ? failed[0].error : '未知原因'}。先用 get_presentation_outline 确认页码范围，再用 indexes 指定能读的页。`);
    }
    // 合并多页清单：按出现次数加权排序，头部就是这套设计的代表色/代表字体
    const merge = key => {
      const totals = new Map();
      for (const page of scanned) for (const entry of page[key]) totals.set(entry.value, (totals.get(entry.value) || 0) + entry.count);
      return [...totals.entries()].sort((a, b) => b[1] - a[1]).map(entry => entry[0]);
    };
    const colors = merge('palette').slice(0, 12);
    const fonts = merge('fonts').slice(0, 8);
    // 字号层级按值从大到小排（merge 按频次排会丢掉「标题大正文小」的层级语义）
    const fontSizes = [...new Set(scanned.flatMap(p => p.fontSizes.map(f => Number(f.value))).filter(Number.isFinite))].sort((a, b) => b - a).slice(0, 8);
    const template = {
      name,
      savedAt: new Date().toISOString(),
      sourceIndexes: scanned.map(p => p.index),
      background: null, // Office.js 读不到页面背景填充；套用时用 set_slide_background 显式指定
      primaryColor: colors[0] || null,
      accentColors: colors.slice(1, 4),
      colors, fonts, fontSizes,
      note: String(args.note || '').slice(0, 200) || undefined
    };
    const all = loadStyleTemplates(); // 含文档里的旧版模板：本次写入会把它们一并迁进共享设置
    const replaced = Boolean(all[name]);
    all[name] = template;
    App.state.settings.pptStyleTemplates = all;
    App.saveSettings();
    return {
      success: true, name, replaced, templateCount: Object.keys(all).length, template,
      hint: '模板已保存到插件共享设置，跨文档可用（不再随单个文件走）。套用时：colors 做 recolor_slide 的 colorMap 目标值，fonts[0] 做 fontName，fontSizes 是字号层级参考。背景色没有存（Office.js 读不到），需要时用 set_slide_background 显式指定。'
    };
  }

  async function listStyleTemplates() {
    requireOffice();
    const all = loadStyleTemplates();
    const templates = Object.keys(all).sort().map(key => {
      const tpl = all[key];
      return {
        name: tpl.name, savedAt: tpl.savedAt, sourceIndexes: tpl.sourceIndexes,
        primaryColor: tpl.primaryColor, accentColors: tpl.accentColors, colors: tpl.colors,
        fonts: tpl.fonts, fontSizes: tpl.fontSizes, note: tpl.note
      };
    });
    const result = { success: true, count: templates.length, templates };
    if (!templates.length) result.hint = '还没有已存模板。换肤完成后用 save_style_template({ name }) 把当前 deck 的配色/字体/字号层级存下来，下次服务同一客户直接套用。';
    return result;
  }

  // set_slide_background 降级路径（全页矩形置底）的固定形状名：按它识别并复用旧矩形
  const PAGE_BACKGROUND_SHAPE_NAME = '__AI_PageBackground';

  // 页面背景：Office.js 大概率没有 slide.background 写 API，探测（不写死）；没有就降级为
  // 「全页矩形置底」并在结果里明说路径。降级需要层级 API（1.8）把矩形压到内容底下，
  // 连层级都没有时宁可报错也不留一页蒙版盖在内容上。
  async function setSlideBackground(args = {}) {
    requireOffice();
    const index = Number(args.index);
    if (!Number.isInteger(index) || index < 0) throw new Error('set_slide_background 需要 index（0 起始页码）和 color（#RRGGBB）。');
    const color = normalizedColor(args.color, 'color');
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      await context.sync();
      // 路径 A（探测）：宿主真有背景填充 API 就直接写
      try {
        const bgFill = slide.background && slide.background.fill;
        if (bgFill && typeof bgFill.setSolidColor === 'function') {
          bgFill.setSolidColor(color);
          await context.sync();
          noteRuntimeProbe('slideBackgroundApi', true);
          return { success: true, index, color, path: 'backgroundApi', _navTarget: { slideId: slide.id } };
        }
      } catch { /* 落入降级路径 */ }
      noteRuntimeProbe('slideBackgroundApi', false);
      // 路径 B（降级）：全页矩形置底
      if (!powerPointApiSupported('1.8') || !slide.shapes || typeof slide.shapes.addGeometricShape !== 'function') {
        throw new Error('本机既没有页面背景写 API（slide.background），也没有层级 API（1.8）把垫底矩形压到内容下面，程序化设背景走不通。替代路径：(1) 请用户在 PowerPoint 里「设计 → 设置背景格式」手动填纯色；(2) 用 list_layouts + add_slide({ layoutId }) 借一个自带背景的版式。');
      }
      const page = await pageSizeForContext(context);
      // 垫底矩形固定命名：第二次调用同一页时按名字找到旧矩形改色，而不是再垫一个——
      // 矩形在最底层，肉眼看不见叠加，但形状列表会越积越脏。
      slide.shapes.load('items');
      await context.sync();
      slide.shapes.items.forEach(sh => sh.load('id,name'));
      await context.sync();
      const existing = slide.shapes.items.find(sh => sh.name === PAGE_BACKGROUND_SHAPE_NAME);
      if (existing) {
        existing.fill.setSolidColor(color);
        await context.sync();
        return {
          success: true, index, color, path: 'fullPageRect', updated: true, shapeId: existing.id,
          note: `本机没有页面背景写 API，已把上次创建的全页垫底矩形（${PAGE_BACKGROUND_SHAPE_NAME}）改为新颜色，没有新建叠加。注意：它会盖住母版/版式自带的背景元素（banner/logo 在渲染层位于 slide 形状之下）——若这页的 banner 来自母版，不要用此工具，走借版（list_layouts + add_slide）。`,
          _navTarget: { slideId: slide.id }
        };
      }
      const rect = slide.shapes.addGeometricShape('Rectangle', { left: 0, top: 0, width: page.width, height: page.height });
      rect.fill.setSolidColor(color);
      try { rect.lineFormat.visible = false; } catch {}
      rect.setZOrder('SendToBack');
      try { rect.name = PAGE_BACKGROUND_SHAPE_NAME; } catch {}
      rect.load('id');
      await context.sync();
      return {
        success: true, index, color, path: 'fullPageRect', shapeId: rect.id,
        note: `本机没有页面背景写 API，已降级为全页矩形置底（sendToBack），命名为 ${PAGE_BACKGROUND_SHAPE_NAME}，再次调用本工具会改它而不是叠加新矩形。它不是真背景：全选/导出时会被当成普通形状，删掉这个 shapeId 即可撤销背景。它会盖住母版/版式自带的背景元素（banner/logo 在渲染层位于 slide 形状之下）——若这页的 banner 来自母版，不要用此工具，走借版（list_layouts + add_slide）。` + (page.exact ? '' : '页面尺寸是回退值 960x540，矩形可能没盖满或超出，建议 get_slide_preview 复检。'),
        _navTarget: { slideId: slide.id }
      };
    });
  }

  // 表格行列增删：Office.js 的 Table API 大概率没有结构操作，探测候选方法名（不写死）；
  // 全都没有就报错给替代路径，不假支持。
  const TABLE_STRUCTURE_ACTIONS = {
    addRow: { methods: ['addRow', 'insertRow'], dimension: 'row' },
    removeRow: { methods: ['deleteRow', 'removeRow'], dimension: 'row' },
    addColumn: { methods: ['addColumn', 'insertColumn'], dimension: 'column' },
    removeColumn: { methods: ['deleteColumn', 'removeColumn'], dimension: 'column' }
  };

  async function editTableStructure(args = {}) {
    requireOffice();
    const target = args.target || {};
    const index = Number(target.index);
    const shapeId = target.shapeId ? String(target.shapeId) : '';
    if (!Number.isInteger(index) || index < 0 || !shapeId) throw new Error('edit_table_structure 需要 target.index 和 target.shapeId——先用 get_slide 读该页，找到 type 为 Table 的形状 id。');
    const action = String(args.action || '');
    const spec = TABLE_STRUCTURE_ACTIONS[action];
    if (!spec) throw new Error(`edit_table_structure 不支持 action "${action}"。支持：${Object.keys(TABLE_STRUCTURE_ACTIONS).join('/')}。position 是 0 起始的行/列号：add 表示插入到该位置（缺省追加到末尾），remove 表示删该位置（缺省删最后一行/列）。`);
    return PowerPoint.run(async context => {
      const { slide, shape } = await locateShapeOnSlide(context, index, shapeId, 'id,name,type');
      if (!shape) throw new Error(`第 ${index} 页找不到形状 ${shapeId}。先用 get_slide 重读该页拿到当前形状列表。`);
      if (String(shape.type || '') !== 'Table') throw new Error(`形状 ${shapeId} 的类型是 ${shape.type}，不是表格。edit_table_structure 只改表格结构；普通形状用 apply_layout，单元格样式用 set_table_style。`);
      let table;
      try {
        table = shape.table;
        if (!table) throw new Error('shape.table missing');
        table.load('rowCount,columnCount');
        await context.sync();
      } catch {
        throw new Error('本机读不到这个表格的 Table API（shape.table 不可用），结构修改更无从谈起。替代路径：请用户在 PowerPoint 里手动增删行列，不要反复试。');
      }
      const rows = Number(table.rowCount) || 0;
      const cols = Number(table.columnCount) || 0;
      const size = spec.dimension === 'row' ? rows : cols;
      const adding = action.indexOf('add') === 0;
      const method = spec.methods.find(name => typeof table[name] === 'function');
      if (!method) {
        noteRuntimeProbe('tableStructureApi', false);
        throw new Error(`本机 Table API 没有行列增删方法（${spec.methods.join('/')} 都不存在），表格结构改不了。替代路径：(1) 请用户在 PowerPoint 里右键表格手动增删行列（最稳）；(2) 用 eval_officejs 读出全部单元格文本后删旧表重建——先确认 slide.shapes.addTable 存在，多数宿主没有；(3) 只是要「视觉上少一行」可以改用 set_table_style 清空该行文字。不要重试其他方法名，已经探测过了。`);
      }
      const maxPosition = adding ? size : size - 1;
      const position = args.position != null ? Number(args.position) : maxPosition;
      if (!Number.isInteger(position) || position < 0 || position > maxPosition) {
        throw new Error(`position ${args.position} 越界：当前表格 ${rows} 行 ${cols} 列，${action} 的合法 position 是 0~${maxPosition}。`);
      }
      table[method](position);
      await context.sync();
      noteRuntimeProbe('tableStructureApi', true);
      // 结构变化后重读行列数，返回给模型核对
      table.load('rowCount,columnCount');
      await context.sync();
      return { success: true, index, shapeId, action, position, method, rowCount: Number(table.rowCount), columnCount: Number(table.columnCount), _navTarget: { slideId: slide.id } };
    });
  }

  // 图片裁剪：Office.js 图片形状大概率没有裁剪 API，探测 shape.cropFormat（不写死）；
  // 没有就报错给替代路径，不假支持。裁剪值是 0~1 的比例（该边往里裁的比例）。
  async function cropImage(args = {}) {
    requireOffice();
    const target = args.target || {};
    const index = Number(target.index);
    const shapeId = target.shapeId ? String(target.shapeId) : '';
    if (!Number.isInteger(index) || index < 0 || !shapeId) throw new Error('crop_image 需要 target.index 和 target.shapeId——先用 get_slide 读该页，找到 type 为 Picture 的形状 id。');
    const crops = {};
    for (const key of ['cropTop', 'cropBottom', 'cropLeft', 'cropRight']) {
      if (!hasOwn(args, key)) continue;
      const value = Number(args[key]);
      if (!Number.isFinite(value) || value < 0 || value >= 1) throw new Error(`${key} 必须是 0~1 之间的小数（0.1 = 该边往里裁 10%）。`);
      crops[key] = value;
    }
    if (!Object.keys(crops).length) throw new Error('crop_image 至少要给一个裁剪值：cropTop / cropBottom / cropLeft / cropRight（0~1 的比例）。');
    return PowerPoint.run(async context => {
      const { slide, shape } = await locateShapeOnSlide(context, index, shapeId, 'id,name,type');
      if (!shape) throw new Error(`第 ${index} 页找不到形状 ${shapeId}。先用 get_slide 重读该页拿到当前形状列表。`);
      if (String(shape.type || '') !== 'Picture') throw new Error(`形状 ${shapeId} 的类型是 ${shape.type}，不是图片。crop_image 只裁剪图片形状。`);
      const cropFormat = shape.cropFormat;
      if (!cropFormat || typeof cropFormat.load !== 'function') {
        noteRuntimeProbe('pictureCrop', false);
        throw new Error('本机图片形状没有裁剪 API（shape.cropFormat 不存在）。替代路径：(1) replace_image 直接换成预先裁好的图（最干净）；(2) 请用户在 PowerPoint 里「图片格式 → 裁剪」手动裁；(3) 只是要挡住某条边，用 apply_layout 加一个与背景同色的矩形遮罩。不要重试其他属性名，已经探测过了。');
      }
      try {
        cropFormat.load('cropTop,cropBottom,cropLeft,cropRight');
        await context.sync();
        for (const key of Object.keys(crops)) cropFormat[key] = crops[key];
        await context.sync();
      } catch (e) {
        noteRuntimeProbe('pictureCrop', false);
        throw new Error(`本机 cropFormat 存在但写入失败（${e && e.message ? e.message : e}）。替代路径：replace_image 换预裁图，或请用户手动裁剪。`);
      }
      noteRuntimeProbe('pictureCrop', true);
      return { success: true, index, shapeId, crops, _navTarget: { slideId: slide.id } };
    });
  }

  // 图片透明度：Office.js 大概率不可写，探测 shape.pictureFormat.transparency（不写死）；
  // 不支持就报错给「垫半透明矩形」替代路径，不假支持。
  async function setPictureOpacity(args = {}) {
    requireOffice();
    const target = args.target || {};
    const index = Number(target.index);
    const shapeId = target.shapeId ? String(target.shapeId) : '';
    if (!Number.isInteger(index) || index < 0 || !shapeId) throw new Error('set_picture_opacity 需要 target.index 和 target.shapeId——先用 get_slide 读该页，找到 type 为 Picture 的形状 id。');
    const opacity = Number(args.opacity);
    if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1) throw new Error('opacity 必须是 0~1 之间的数（1 = 不透明，0.4 = 四成不透明）。');
    const veilTransparency = (1 - opacity).toFixed(2);
    return PowerPoint.run(async context => {
      const { slide, shape } = await locateShapeOnSlide(context, index, shapeId, 'id,name,type');
      if (!shape) throw new Error(`第 ${index} 页找不到形状 ${shapeId}。先用 get_slide 重读该页拿到当前形状列表。`);
      if (String(shape.type || '') !== 'Picture') throw new Error(`形状 ${shapeId} 的类型是 ${shape.type}，不是图片。set_picture_opacity 只改图片形状；普通形状的填充透明度用 apply_layout 的 fillTransparency。`);
      const pictureFormat = shape.pictureFormat;
      if (!pictureFormat || typeof pictureFormat.load !== 'function') {
        noteRuntimeProbe('pictureTransparency', false);
        throw new Error(`本机写不了图片透明度（shape.pictureFormat 不存在）。替代路径：用 apply_layout 在图片上盖一个与背景同色的半透明矩形——geometry 从 get_slide 抄图片的 left/top/width/height，fillTransparency 填 ${veilTransparency}——视觉等同于图片变淡；或请用户手动「图片格式 → 透明度」。`);
      }
      try {
        pictureFormat.load('transparency');
        await context.sync();
        pictureFormat.transparency = 1 - opacity;
        await context.sync();
      } catch (e) {
        noteRuntimeProbe('pictureTransparency', false);
        throw new Error(`本机 pictureFormat 存在但透明度写入失败（${e && e.message ? e.message : e}）。替代路径：apply_layout 盖半透明矩形（fillTransparency ${veilTransparency}），或请用户手动设置。`);
      }
      noteRuntimeProbe('pictureTransparency', true);
      return { success: true, index, shapeId, opacity, _navTarget: { slideId: slide.id } };
    });
  }

  // ask_clarification 的真正通道在 api.js：出卡前被拦截，转给 render-blocks.js 的
  // App.presentClarification（render form 的语义化包装）。这个执行器只是兜底——
  // 走到这说明界面通道没加载（比如无界面的测试环境），让模型改用文字提问。
  async function askClarificationTool() {
    return { success: false, error: 'ask_clarification 的问卷卡通道（App.presentClarification）没有加载，卡片出不来。降级：把问题直接用正文写出来问用户，等用户回复后再继续。' };
  }

  // 列出母版与版式：宿主没有 slide.duplicate 时，「按指定版式新建页」是唯一能借到整套设计
  // （背景 / banner / logo / 占位符几何）的官方途径，所以这是借版流程的第一步。
  async function listLayouts() {
    requireOffice();
    if (!powerPointApiSupported('1.3')) throw new Error('当前 PowerPoint 不支持母版/版式 API（需要 PowerPointApi 1.3）。');
    return PowerPoint.run(async context => {
      const masters = context.presentation.slideMasters;
      masters.load('items');
      await context.sync();
      masters.items.forEach(m => { m.load('id,name'); m.layouts.load('items'); });
      await context.sync();
      masters.items.forEach(m => m.layouts.items.forEach(l => l.load('id,name')));
      await context.sync();
      const out = masters.items.map(m => ({
        masterId: m.id,
        masterName: m.name,
        layouts: m.layouts.items.map(l => ({ layoutId: l.id, name: l.name }))
      }));
      const layoutCount = out.reduce((n, m) => n + m.layouts.length, 0);
      return { success: true, masterCount: out.length, layoutCount, masters: out, hint: 'Pass one of these layoutId values to add_slide to build a new slide on that design.' };
    });
  }

  async function addSlide(args = {}) {
    requireOffice();
    const layoutId = args.layoutId ? String(args.layoutId) : '';
    const slideMasterId = args.slideMasterId ? String(args.slideMasterId) : '';
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      if (layoutId) {
        const options = { layoutId };
        if (slideMasterId) options.slideMasterId = slideMasterId;
        slides.add(options);
      } else {
        slides.add(); // 默认版式
      }
      slides.load('items');
      try { await context.sync(); }
      catch (e) {
        const hint = layoutId ? `（layoutId="${layoutId}" 可能无效，先用 list_layouts 取有效 id）` : '';
        throw new Error(`新建幻灯片失败${hint}：${e && e.message ? e.message : e}`);
      }
      const newIndex = slides.items.length - 1;
      const slide = slides.items[newIndex];
      slide.load('id');
      await context.sync();
      return {
        success: true, index: newIndex, id: slide.id, _navTarget: { slideId: slide.id },
        // C1 母版感知：新页的占位区几何只有读了才知道——不读就填内容必然 off-brand
        hint: 'New slide created' + (layoutId ? ' on the requested layout' : '') + '. Call get_slide on this index next: the layout placeholders (their positions and sizes) decide where text may land — put content inside them, never on top of the banner or logo.'
      };
    });
  }

  async function deleteSlide(args) {
    requireOffice();
    const index = Number(args.index);
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.delete();
      await context.sync();
      return { success: true, deletedIndex: index };
    });
  }

  // 把当前演示文稿整个读成 base64。
  // 原理：Office.context.document.getFileAsync(FileType.Compressed) 返回当前 pptx 文件的切片序列，
  // 逐片 getSliceAsync 取 slice.data 拼接即得完整文件的 base64，读完必须 closeAsync。
  // 用途：宿主没有 slide.duplicate() 时，把「当前文件自己」当作 insertSlidesFromBase64 的源文件——
  // 源和目标同一份文件、只插指定的一页，效果等同于复制该页，且版式/占位符/备注全部保留。
  // 两道护栏：
  // 1. 大小守卫——拼装中的 base64 超过约 32MB（对应原文件 ~24MB）就中止：再大 insertSlidesFromBase64
  //    会长时间卡死甚至拖垮宿主，而且复制结果用户也等不到。中止时报错指向借版路径。
  // 2. 超时——getFileAsync/getSliceAsync 是回调式 API，宿主偶发不回电时 Promise 永远悬置，
  //    用 Promise.race 套 60 秒上限，超时报同样的借版提示。
  const FILE_COPY_MAX_BASE64 = 32 * 1024 * 1024;
  const FILE_COPY_TIMEOUT_MS = 60000;
  function promiseWithTimeout(promise, ms, message) {
    // 测试沙箱（Node vm）里没有 setTimeout，守卫自动失效，不影响主流程
    if (typeof setTimeout !== 'function') return promise;
    return Promise.race([promise, new Promise((_, reject) => setTimeout(() => {
      const err = new Error(message);
      err.isTimeout = true;
      reject(err);
    }, ms))]);
  }
  function fileCopyUnavailable(reason) {
    noteRuntimeProbe('duplicateViaBase64', false);
    return new Error(`${reason}，文件级复制不可用，请走 list_layouts + add_slide 借版路径`);
  }
  function readPresentationBase64() {
    const reading = new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error('getFileAsync 读取当前文件失败：' + (result.error ? result.error.message : result.status)));
          return;
        }
        const file = result.value;
        const parts = [];
        let totalBytes = 0;
        const readSlice = i => {
          if (i >= file.sliceCount) {
            file.closeAsync();
            resolve(parts.join(''));
            return;
          }
          file.getSliceAsync(i, sliceResult => {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(new Error('getSliceAsync 读取文件切片失败：' + (sliceResult.error ? sliceResult.error.message : sliceResult.status)));
              return;
            }
            const data = String((sliceResult.value && sliceResult.value.data) || '');
            totalBytes += data.length;
            if (totalBytes > FILE_COPY_MAX_BASE64) {
              file.closeAsync();
              // base64 长度 ×3/4 ≈ 原文件字节数
              reject(fileCopyUnavailable(`文件过大（约 ${Math.round(totalBytes * 3 / 4 / 1024 / 1024)} MB）`));
              return;
            }
            parts.push(data);
            readSlice(i + 1);
          });
        };
        readSlice(0);
      });
    });
    return promiseWithTimeout(reading, FILE_COPY_TIMEOUT_MS, '读取当前文件超时（60 秒无响应）')
      .catch(e => {
        // 超时报错与其他失败统一收口成「借版路径」提示，并探测写回 capabilities
        if (e && e.isTimeout) throw fileCopyUnavailable(e.message);
        throw e;
      });
  }

  // 路径 2 的核心：用当前文件的 base64 把指定页插回它自己后面。
  // slide.id 形如 "2147481232#588170797"，sourceSlideIds 接受的 id 格式因宿主而异（完整 id 或 # 后缀），
  // 先试完整 id，抛错再试 # 后缀；成功的格式记入 duplicateBase64IdFormat 并探测写回 capabilities。
  async function insertSlideCopyFromBase64(context, slideId) {
    if (typeof context.presentation.insertSlidesFromBase64 !== 'function') {
      noteRuntimeProbe('duplicateViaBase64', false);
      throw new Error('presentation.insertSlidesFromBase64 不存在');
    }
    const base64 = await readPresentationBase64();
    const candidates = [slideId];
    const hashIndex = slideId.indexOf('#');
    if (hashIndex >= 0 && hashIndex < slideId.length - 1) candidates.push(slideId.slice(hashIndex + 1));
    let lastError = null;
    for (const sourceId of candidates) {
      try {
        // targetSlideId = 源页自己：新页插在源页之后
        context.presentation.insertSlidesFromBase64(base64, { sourceSlideIds: [sourceId], targetSlideId: slideId });
        await context.sync();
        duplicateBase64IdFormat = sourceId === slideId ? 'full' : 'suffix';
        noteRuntimeProbe('duplicateViaBase64', true);
        return;
      } catch (e) {
        lastError = e;
      }
    }
    noteRuntimeProbe('duplicateViaBase64', false);
    throw lastError || new Error('insertSlidesFromBase64 两种 slide id 格式都被拒绝');
  }

  async function duplicateSlide(args) {
    requireOffice();
    const index = Number(args.index);
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      // 复制前记录全部页 id：复制完成后，id 不在该集合里的那一页就是新页（通常紧随源页之后）
      slides.items.forEach(s => s.load('id'));
      await context.sync();
      const beforeIds = new Set(slides.items.map(s => s.id));

      if (typeof slide.duplicate === 'function') {
        // 路径 1：原生 API（Mac 16.112 上没有；走到这里说明宿主较新）
        noteRuntimeProbe('duplicateSlide', true);
        slide.duplicate();
        await context.sync();
      } else {
        // 路径 2：文件级复制——Mac 16.112 实测 insertSlidesFromBase64 存在，借它落地「借版」
        noteRuntimeProbe('duplicateSlide', false);
        try {
          await insertSlideCopyFromBase64(context, String(slide.id || ''));
        } catch (e) {
          throw new Error(`本机没有 slide.duplicate()，文件级复制（getFileAsync + insertSlidesFromBase64）也失败了（${e && e.message ? e.message : e}）。借版降级路径：list_layouts 列出母版下的版式 → add_slide({ layoutId }) 按目标版式新建页（自带背景/banner/logo/占位符）→ 在新页上填内容 → delete_slide 删原页。新页在末尾，页序需要用户自己在 PowerPoint 里拖。`);
        }
      }

      slides.load('items');
      await context.sync();
      slides.items.forEach(s => s.load('id'));
      await context.sync();
      let newIndex = slides.items.findIndex(s => !beforeIds.has(s.id));
      if (newIndex < 0) newIndex = Math.min(index + 1, slides.items.length - 1); // 兜底：按 PowerPoint 插入位置约定
      const newSlide = slides.items[newIndex];
      const newSlideId = newSlide ? newSlide.id : null;
      // 返回新页身份，AI 不用猜新页在哪；_navTarget 供跟随模式直接跳转到新页
      return { success: true, sourceIndex: index, index: newIndex, slideId: newSlideId, _navTarget: { slideId: newSlideId } };
    });
  }

  async function getSlideNotes(args = {}) {
    requireOffice();
    const hasIndex = args.index !== undefined && args.index !== null && args.index !== '';
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const index = Number(args.index);
      if (hasIndex && !slides.items[index]) throw new Error(`Slide ${index} not found`);
      const items = hasIndex ? [{ slide: slides.items[index], index }] : slides.items.map((slide, i) => ({ slide, index: i }));
      const out = [];
      for (const item of items) {
        let notesShapes;
        try {
          item.slide.notesSlide.shapes.load('items');
          await context.sync();
          notesShapes = item.slide.notesSlide.shapes.items;
        } catch (e) {
          // 读取失败最常见的原因是这一页还没有备注页实例（从未写过备注），而不是宿主缺少备注 API。
          // 整体抛错会让模型误判「本机不支持备注」从此绕道，所以按「该页无备注」处理，继续读下一页。
          out.push({ index: item.index, notes: '', empty: true, reason: '该页暂无备注（尚未创建备注页实例）' });
          continue;
        }
        const texts = [];
        for (const shape of notesShapes) {
          const t = await shapeTextSafe(context, shape);
          const trimmed = t.trim();
          // 过滤页码占位符这类纯数字碎片，剩下的才是讲稿正文
          if (trimmed && !/^\d{1,3}$/.test(trimmed)) texts.push(trimmed);
        }
        out.push({ index: item.index, notes: texts.join('\n\n') });
      }
      if (hasIndex) return Object.assign({ success: true }, out[0]);
      return { success: true, slideCount: slides.items.length, slides: out };
    });
  }

  async function setSlideNotes(args) {
    requireOffice();
    const index = Number(args.index);
    const text = String(args.text || '');
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      let notes;
      try { notes = slide.notesSlide; notes.shapes.load('items'); await context.sync(); }
      catch (e) { throw new Error('无法访问该页的备注页，通常是这一页还没有备注页实例。让用户在 PowerPoint 里给该页先手动输入任意一个字的备注再重试，或改用 eval_officejs。'); }
      // 优先按占位符类型定位备注正文（body）：直接取「第一个有 textFrame 的形状」
      // 可能误中页眉/页码/幻灯片图像占位符；找不到 body 占位符再回退原逻辑
      let target = null;
      try {
        notes.shapes.items.forEach(sh => { try { sh.placeholderFormat.load('type'); } catch {} });
        await context.sync();
        target = notes.shapes.items.find(sh => {
          try { return String(sh.placeholderFormat.type || '') === 'Body'; } catch { return false; }
        }) || null;
      } catch { /* placeholderFormat 不可用：走下方回退 */ }
      if (!target) {
        for (const sh of notes.shapes.items) { try { sh.textFrame.textRange.load('text'); target = target || sh; } catch {} }
        await context.sync();
      }
      if (!target) throw new Error('No notes placeholder found on this slide');
      target.textFrame.textRange.text = text;
      await context.sync();
      return { success: true, index, _navTarget: { slideId: slide.id } };
    });
  }

  async function insertTextbox(args) {
    requireOffice();
    const { index, text = '', left = 50, top = 50, width = 400, height = 100 } = args;
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[Number(index)];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      if (!slide.shapes || typeof slide.shapes.addTextBox !== 'function') throw new Error('shapes.addTextBox not available in this PowerPoint host; use eval_officejs');
      slide.shapes.addTextBox(text, { left, top, width, height });
      await context.sync();
      return { success: true, index: Number(index), _navTarget: { slideId: slide.id } };
    });
  }

  async function setText(args) {
    requireOffice();
    const { index, shapeId, text = '' } = args;
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[Number(index)];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      slide.shapes.load('items');
      await context.sync();
      slide.shapes.items.forEach(sh => sh.load('id'));
      await context.sync();
      const shape = shapeId ? slide.shapes.items.find(sh => sh.id === shapeId) : slide.shapes.items[0];
      if (!shape) throw new Error('Target shape not found');
      // 保留基础字符格式：优先对整段文本的 range 做 insertText(Replace)（继承原格式），
      // 而不是直接赋 textRange.text（会按宿主默认格式重置）；getSubstring 不可用时降级为直接赋值。
      let replaced = false;
      try {
        const range = shape.textFrame.textRange;
        range.load('text');
        await context.sync();
        range.getSubstring(0, String(range.text || '').length).insertText(String(text), 'Replace');
        replaced = true;
      } catch { /* 落入下方直接赋值降级路径 */ }
      if (!replaced) shape.textFrame.textRange.text = text;
      await context.sync();
      return { success: true, index: Number(index), _navTarget: { slideId: slide.id } };
    });
  }

  // PPT-BLUEPRINT P1：超链接写入（vendor 16.0.20416 解锁，TextRange/Shape.setHyperlink，门控 PowerPointApi 1.10）。
  // 三条路径：命中文本子串（target+find）/ 整个形状（target 无 find）/ 当前选中文字（无 target，getSelectedTextRange 1.5）。
  // 写后读回（Word W2 三态）：setHyperlink 返回 Hyperlink → load('address,screenTip') → 比对。
  // API 依据（vendored 取证）：TextRange.setHyperlink 与 Shape.setHyperlink 均存在且返回 e.Hyperlink；
  // Hyperlink 标量属性 address/screenTip/type（_scalarPropertyNames 实录）。
  async function setHyperlinkTool(args = {}) {
    requireOffice();
    const address = String(args.address || '').trim();
    if (!address) throw new Error('address is required (https://…, mailto:…, or a file name).');
    const screenTip = args.screenTip ? String(args.screenTip) : '';
    const target = args.target || {};
    const find = args.find != null ? String(args.find) : '';
    if (!powerPointApiSupported('1.10')) {
      noteRuntimeProbe('hyperlinkApi', false);
      throw new Error('本机 PowerPoint 不支持超链接 API（PowerPointApi 1.10）。替代路径：用 set_text 把链接地址写成纯文本，并告诉用户在 PowerPoint 里选中后按 Cmd+K 手动添加链接。');
    }
    return PowerPoint.run(async context => {
      let range = null;        // TextRange 级链接目标
      let shape = null;        // Shape 级链接目标
      let mode = '';
      let nav = {};
      if (target.shapeId || target.slideId || target.index != null) {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        const slide = target.slideId
          ? slides.items.find(item => item.id === target.slideId)
          : slides.items[Number(target.index)];
        if (!slide) throw new Error(`Slide ${target.slideId || target.index} not found. Get the real id/index from get_presentation_outline first.`);
        slide.load('id');
        slide.shapes.load('items');
        await context.sync();
        slide.shapes.items.forEach(sh => sh.load('id,name'));
        await context.sync();
        shape = target.shapeId ? slide.shapes.items.find(sh => sh.id === target.shapeId) : null;
        if (!shape) throw new Error(`Shape ${target.shapeId || ''} not found on slide ${target.slideId || target.index}. Get shapeId from get_slide first.`);
        nav = { index: slides.items.indexOf(slide), slideId: slide.id, shapeId: shape.id };
        if (find) {
          const textRange = shape.textFrame.textRange;
          textRange.load('text');
          await context.sync();
          const full = String(textRange.text || '');
          const at = full.indexOf(find);
          if (at < 0) {
            throw new Error(`find "${find}" does not appear in the shape text. Read the exact text with get_slide first; "find" must be an exact substring.`);
          }
          range = textRange.getSubstring(at, find.length);
          mode = 'substring';
        } else {
          mode = 'shape';
        }
      } else {
        // 无 target：作用于当前选中的文字（用户在 PowerPoint 里选好一段文本再让 AI 加链接）
        if (!powerPointApiSupported('1.5')) {
          throw new Error('本机不支持读取选中文本（PowerPointApi 1.5）。请用 target + find 指定要加链接的文字。');
        }
        range = context.presentation.getSelectedTextRange();
        range.load('text');
        await context.sync();
        const selected = String(range.text || '');
        if (!selected.trim()) throw new Error('当前没有选中文字。请在 PowerPoint 里选中要加链接的文本，或用 target + find 指定。');
        mode = 'selection';
      }
      let hyperlink;
      try {
        hyperlink = (mode === 'shape' ? shape : range).setHyperlink({ address, screenTip: screenTip || undefined });
        hyperlink.load('address,screenTip');
        await context.sync();
      } catch (e) {
        noteRuntimeProbe('hyperlinkApi', false);
        throw new Error(`设置超链接失败（${String((e && e.message) || e).slice(0, 200)}）。本机 office.js 可能不含该能力。替代路径：用 set_text 把地址写成纯文本，请用户 Cmd+K 手动加链接。`);
      }
      noteRuntimeProbe('hyperlinkApi', true);
      // 写后读回三态（Word W2 模式）
      const mismatches = [];
      if (String(hyperlink.address || '') !== address) mismatches.push(`address 读回 ${hyperlink.address}`);
      if (screenTip && String(hyperlink.screenTip || '') !== screenTip) mismatches.push(`screenTip 读回 ${hyperlink.screenTip}`);
      const verification = mismatches.length
        ? { verified: false, verifyWarning: `疑似未生效：${mismatches.join('；')}` }
        : { verified: true };
      return Object.assign({ success: true, mode, address, _navTarget: nav.slideId ? { slideId: nav.slideId } : undefined }, verification, nav);
    });
  }

  const LAYOUT_OPERATIONS = new Set(['addShape', 'updateShape', 'deleteShape']);
  const LAYOUT_SHAPE_TYPES = new Set(['rectangle', 'roundRectangle', 'ellipse', 'textBox', 'line']);
  const LAYOUT_Z_ORDERS = new Set(['sendToBack', 'sendBackward', 'bringForward', 'bringToFront']);

  function hasOwn(obj, key) { return Object.prototype.hasOwnProperty.call(obj || {}, key); }
  function finiteNumber(value, label, options = {}) {
    const number = Number(value);
    if (!Number.isFinite(number)) throw new Error(`${label} must be a finite number`);
    if (options.positive && number <= 0) throw new Error(`${label} must be greater than 0`);
    if (Math.abs(number) > 10000) throw new Error(`${label} is outside the safe range`);
    return number;
  }
  function normalizedColor(value, label) {
    const text = String(value || '').trim();
    if (!/^#[0-9a-f]{6}$/i.test(text)) throw new Error(`${label} must use #RRGGBB format`);
    return text.toUpperCase();
  }
  function normalizedTransparency(value, label) {
    const number = Number(value);
    if (!Number.isFinite(number) || number < 0 || number > 1) throw new Error(`${label} must be between 0 and 1`);
    return number;
  }

  // updateShape 改了哪个字段，就需要当前快照里对应的字段组（enrich 与校验共用这张表）
  const EXPECTED_BY_DESIRED = {
    name: ['name'], text: ['text'], left: ['left'], top: ['top'], width: ['width'], height: ['height'],
    fillColor: ['fillType', 'fillColor', 'fillTransparency'], fillTransparency: ['fillType', 'fillColor', 'fillTransparency'],
    lineVisible: ['lineVisible', 'lineColor', 'lineTransparency', 'lineWidth'], lineColor: ['lineVisible', 'lineColor', 'lineTransparency', 'lineWidth'], lineTransparency: ['lineVisible', 'lineColor', 'lineTransparency', 'lineWidth'], lineWidth: ['lineVisible', 'lineColor', 'lineTransparency', 'lineWidth'],
    fontName: ['fontName', 'fontSize', 'fontColor', 'bold'], fontSize: ['fontName', 'fontSize', 'fontColor', 'bold'], fontColor: ['fontName', 'fontSize', 'fontColor', 'bold'], bold: ['fontName', 'fontSize', 'fontColor', 'bold'],
    horizontalAlignment: ['horizontalAlignment'], verticalAlignment: ['verticalAlignment'],
    marginLeft: ['marginLeft'], marginRight: ['marginRight'], marginTop: ['marginTop'], marginBottom: ['marginBottom'], wordWrap: ['wordWrap'], autoSizeSetting: ['autoSizeSetting'],
    zOrder: ['zOrderPosition']
  };

  // 【47.2】按对象能力读字段：这些大类在真机（16.112 实测契约，与 fake UNSUPPORTED 同源）
  // 没有 textFrame——对它们 load 文字字段会让整批 sync 取消；Line/Connector 连 fill 都没有。
  // 快照按能力捕获，不伪造空文本；组合（Group）内部成员在本宿主不可读，如实标注。
  const SHAPE_CAPABILITY_LIMITS = {
    Group: ['textFrame'],
    Picture: ['textFrame'],
    Line: ['textFrame', 'fill'],
    Connector: ['textFrame', 'fill'],
    Chart: ['textFrame'],
    Table: ['textFrame']
  };
  const TEXTFRAME_FIELDS = ['text', 'fontName', 'fontSize', 'fontColor', 'bold', 'horizontalAlignment', 'verticalAlignment', 'marginLeft', 'marginRight', 'marginTop', 'marginBottom', 'wordWrap', 'autoSizeSetting'];
  const FILL_FIELDS = ['fillType', 'fillColor', 'fillTransparency'];
  function describeEditTarget(target) {
    if (!target) return '未知目标';
    const slide = target.slideId != null ? `slide ${target.slideId}` : (target.index != null ? `第 ${target.index} 页` : '未知页');
    return `${slide} ${target.shapeId != null ? `shape ${target.shapeId}` : ''}`.trim();
  }

  function validateLayoutEdit(edit) {
    const operation = edit && edit.operation;
    // 报错必须带上模型尝试的 type 和支持列表，并指明表格的正确改法——
    // 真实事故：模型连试 14 次（resizeTable / shrinkTable / setCellFont…），每次只看到
    // 一句「unsupported operation」，既不知道自己错在哪，也不知道表格该走哪条路，只能瞎试。
    if (!operation || !LAYOUT_OPERATIONS.has(operation.type)) {
      const attempted = operation && operation.type != null ? `"${operation.type}"` : '（operation 或 operation.type 缺失）';
      throw new Error(`Layout proposal has an unsupported operation ${attempted}. Supported: ${[...LAYOUT_OPERATIONS].join('/')}. 表格（Table 形状）的单元格字体/填充修改用 set_table_style（apply_layout 不支持单元格级修改）；表格整体位置和尺寸可用 updateShape 改 left/top/width/height。`);
    }
    const target = edit.target || {};
    const index = Number(target.index);
    if (!target.slideId && (target.index == null || !Number.isInteger(index) || index < 0)) throw new Error('Layout proposal is missing a valid slide target');
    const desired = operation.shape || {};
    if (operation.type === 'addShape') {
      if (!LAYOUT_SHAPE_TYPES.has(desired.shapeType)) throw new Error('Layout proposal has an unsupported shape type');
      for (const key of ['left', 'top', 'width', 'height']) {
        if (!hasOwn(desired, key)) throw new Error(`New shape is missing ${key}`);
        finiteNumber(desired[key], key, { positive: key === 'width' || key === 'height' });
      }
      // 【43】圆角衔接修饰（仅 addShape）：cornerStyle:'top' = 上圆下直（卡头色条贴圆角外框），
      // cornerRadius = 绝对磅值半径；配套使用时色条与外框必须同 left/top/width、同 cornerRadius
      if (hasOwn(desired, 'cornerStyle')) {
        if (desired.shapeType !== 'roundRectangle') throw new Error('cornerStyle 仅适用于 roundRectangle');
        if (desired.cornerStyle !== 'top') throw new Error('cornerStyle 只支持 "top"（上圆下直）');
      }
      if (hasOwn(desired, 'cornerRadius')) {
        finiteNumber(desired.cornerRadius, 'cornerRadius');
        const maxRadius = Math.min(Number(desired.width), Number(desired.height)) / 2;
        if (Number(desired.cornerRadius) > maxRadius) {
          throw new Error(`cornerRadius ${desired.cornerRadius} 超过该形状允许的最大半径 ${maxRadius}（min(width,height)/2）。调小 cornerRadius 或增大形状高度。`);
        }
      }
    } else {
      if (!target.shapeId) throw new Error('Layout proposal is missing shapeId');
      // 【47.2】出卡时的快照读取失败（或能力限制）已逐项记录：按结构化信息打回，
      // 不再只给笼统的 missing snapshot，也不把确定失败标成可原样重试
      if (operation._snapshotError) {
        const err = operation._snapshotError;
        throw new Error(`无法核验目标现状，此项没有应用（对象：${err.object || '未知'}，阶段：${err.phase}）：${err.reason}${err.retryable ? '请先重新 get_slide 确认现状，再重出提案；不要按原参数原样重试。' : '此项不可原样重试，请按上面的替代路径处理。'}`);
      }
      const expected = operation.expected;
      if (!expected || !['name', 'type', 'text', 'left', 'top', 'width', 'height'].some(key => hasOwn(expected, key))) {
        throw new Error('Layout proposal is missing the current shape snapshot');
      }
      if (operation.type === 'updateShape' && !Object.keys(desired).length) throw new Error('Layout update does not contain any changes');
      if (operation.type === 'updateShape') {
        const expected = operation.expected;
        for (const key of Object.keys(desired)) {
          if (key === 'shapeType') continue;
          const required = EXPECTED_BY_DESIRED[key] || [];
          if (required.some(field => !hasOwn(expected, field))) throw new Error(`Layout proposal is missing the current ${required.join('/')} value`);
        }
      }
    }
    for (const key of ['left', 'top', 'width', 'height']) {
      if (hasOwn(desired, key)) finiteNumber(desired[key], key, { positive: key === 'width' || key === 'height' });
    }
    if (hasOwn(desired, 'fillColor')) normalizedColor(desired.fillColor, 'fillColor');
    if (hasOwn(desired, 'fontColor')) normalizedColor(desired.fontColor, 'fontColor');
    if (hasOwn(desired, 'lineColor')) normalizedColor(desired.lineColor, 'lineColor');
    if (hasOwn(desired, 'fillTransparency')) normalizedTransparency(desired.fillTransparency, 'fillTransparency');
    if (hasOwn(desired, 'lineTransparency')) normalizedTransparency(desired.lineTransparency, 'lineTransparency');
    if (hasOwn(desired, 'lineWidth')) finiteNumber(desired.lineWidth, 'lineWidth', { positive: true });
    if (hasOwn(desired, 'fontSize')) finiteNumber(desired.fontSize, 'fontSize', { positive: true });
    if (hasOwn(desired, 'autoSizeSetting') && !['AutoSizeNone', 'AutoSizeTextToFitShape', 'AutoSizeShapeToFitText'].includes(desired.autoSizeSetting)) {
      throw new Error('autoSizeSetting is unsupported');
    }
    if (desired.zOrder && !LAYOUT_Z_ORDERS.has(desired.zOrder)) throw new Error('Layout proposal has an unsupported z-order operation');
    return { operation, target, desired };
  }

  async function currentShapeSnapshot(context, shape, expected) {
    const fillFields = ['fillType', 'fillColor', 'fillTransparency'];
    const lineFields = ['lineVisible', 'lineColor', 'lineTransparency', 'lineWidth'];
    const frameFields = ['marginLeft', 'marginRight', 'marginTop', 'marginBottom', 'verticalAlignment', 'wordWrap', 'autoSizeSetting'];
    const fontFields = ['fontName', 'fontSize', 'fontColor', 'bold'];
    const needsTextRange = ['text', 'horizontalAlignment', ...fontFields].some(key => hasOwn(expected, key));
    if (fillFields.some(key => hasOwn(expected, key))) shape.fill.load('type,foregroundColor,transparency');
    if (lineFields.some(key => hasOwn(expected, key))) shape.lineFormat.load('visible,color,transparency,weight');
    if (frameFields.some(key => hasOwn(expected, key))) shape.textFrame.load('leftMargin,rightMargin,topMargin,bottomMargin,verticalAlignment,wordWrap,autoSizeSetting');
    if (needsTextRange) {
      shape.textFrame.textRange.load('text');
      if (fontFields.some(key => hasOwn(expected, key))) shape.textFrame.textRange.font.load('name,size,color,bold');
      if (hasOwn(expected, 'horizontalAlignment')) shape.textFrame.textRange.paragraphFormat.load('horizontalAlignment');
    }
    if (hasOwn(expected, 'zOrderPosition')) shape.load('zOrderPosition');
    await context.sync();
    const snapshot = {
      name: shape.name,
      type: shape.type,
      left: Number(shape.left),
      top: Number(shape.top),
      width: Number(shape.width),
      height: Number(shape.height)
    };
    if (fillFields.some(key => hasOwn(expected, key))) Object.assign(snapshot, { fillType: String(shape.fill.type || ''), fillColor: String(shape.fill.foregroundColor || ''), fillTransparency: shape.fill.transparency });
    if (lineFields.some(key => hasOwn(expected, key))) Object.assign(snapshot, { lineVisible: shape.lineFormat.visible, lineColor: String(shape.lineFormat.color || ''), lineTransparency: shape.lineFormat.transparency, lineWidth: shape.lineFormat.weight });
    if (frameFields.some(key => hasOwn(expected, key))) Object.assign(snapshot, { marginLeft: shape.textFrame.leftMargin, marginRight: shape.textFrame.rightMargin, marginTop: shape.textFrame.topMargin, marginBottom: shape.textFrame.bottomMargin, verticalAlignment: String(shape.textFrame.verticalAlignment || ''), wordWrap: shape.textFrame.wordWrap, autoSizeSetting: String(shape.textFrame.autoSizeSetting || '') });
    if (needsTextRange) snapshot.text = String(shape.textFrame.textRange.text || '');
    if (fontFields.some(key => hasOwn(expected, key))) Object.assign(snapshot, { fontName: String(shape.textFrame.textRange.font.name || ''), fontSize: shape.textFrame.textRange.font.size, fontColor: String(shape.textFrame.textRange.font.color || ''), bold: shape.textFrame.textRange.font.bold });
    if (hasOwn(expected, 'horizontalAlignment')) snapshot.horizontalAlignment = String(shape.textFrame.textRange.paragraphFormat.horizontalAlignment || '');
    if (hasOwn(expected, 'zOrderPosition')) snapshot.zOrderPosition = shape.zOrderPosition;
    return snapshot;
  }

  function shapeSnapshotMismatchedFields(expected, current) {
    const mismatched = [];
    for (const key of ['left', 'top', 'width', 'height', 'fillTransparency', 'lineTransparency', 'lineWidth', 'fontSize', 'marginLeft', 'marginRight', 'marginTop', 'marginBottom', 'zOrderPosition']) {
      if (hasOwn(expected, key) && Math.abs(Number(expected[key]) - Number(current[key])) > 0.25) mismatched.push(key);
    }
    for (const key of ['name', 'type', 'text', 'fillType', 'fillColor', 'lineColor', 'fontName', 'fontColor', 'horizontalAlignment', 'verticalAlignment', 'autoSizeSetting']) {
      if (hasOwn(expected, key) && String(expected[key] == null ? '' : expected[key]) !== String(current[key] == null ? '' : current[key])) mismatched.push(key);
    }
    for (const key of ['lineVisible', 'bold', 'wordWrap']) {
      if (hasOwn(expected, key) && Boolean(expected[key]) !== Boolean(current[key])) mismatched.push(key);
    }
    return mismatched;
  }

  function shapeSnapshotMatches(expected, current) {
    return shapeSnapshotMismatchedFields(expected, current).length === 0;
  }

  function shapeOptions(desired) {
    const options = {};
    for (const key of ['left', 'top', 'width', 'height']) {
      if (hasOwn(desired, key)) options[key] = finiteNumber(desired[key], key, { positive: key === 'width' || key === 'height' });
    }
    return options;
  }

  function applyShapeFormatting(shape, desired) {
    for (const key of ['left', 'top', 'width', 'height']) {
      if (hasOwn(desired, key)) shape[key] = finiteNumber(desired[key], key, { positive: key === 'width' || key === 'height' });
    }
    if (hasOwn(desired, 'name')) shape.name = String(desired.name || '').slice(0, 250);
    if (hasOwn(desired, 'fillColor')) shape.fill.setSolidColor(normalizedColor(desired.fillColor, 'fillColor'));
    if (hasOwn(desired, 'fillTransparency')) shape.fill.transparency = normalizedTransparency(desired.fillTransparency, 'fillTransparency');
    if (hasOwn(desired, 'lineVisible')) shape.lineFormat.visible = Boolean(desired.lineVisible);
    if (hasOwn(desired, 'lineColor')) shape.lineFormat.color = normalizedColor(desired.lineColor, 'lineColor');
    if (hasOwn(desired, 'lineTransparency')) shape.lineFormat.transparency = normalizedTransparency(desired.lineTransparency, 'lineTransparency');
    if (hasOwn(desired, 'lineWidth')) shape.lineFormat.weight = finiteNumber(desired.lineWidth, 'lineWidth', { positive: true });

    const hasTextSetting = ['text', 'fontName', 'fontSize', 'fontColor', 'bold', 'horizontalAlignment', 'verticalAlignment', 'marginLeft', 'marginRight', 'marginTop', 'marginBottom', 'wordWrap', 'autoSizeSetting'].some(key => hasOwn(desired, key));
    if (hasTextSetting) {
      const textRange = shape.textFrame.textRange;
      if (hasOwn(desired, 'text')) textRange.text = String(desired.text == null ? '' : desired.text).slice(0, 20000);
      if (hasOwn(desired, 'fontName')) textRange.font.name = String(desired.fontName || '').slice(0, 120);
      if (hasOwn(desired, 'fontSize')) textRange.font.size = finiteNumber(desired.fontSize, 'fontSize', { positive: true });
      if (hasOwn(desired, 'fontColor')) textRange.font.color = normalizedColor(desired.fontColor, 'fontColor');
      if (hasOwn(desired, 'bold')) textRange.font.bold = Boolean(desired.bold);
      if (hasOwn(desired, 'horizontalAlignment')) {
        const value = String(desired.horizontalAlignment || '');
        if (!['Left', 'Center', 'Right', 'Justify', 'Distributed'].includes(value)) throw new Error('horizontalAlignment is unsupported');
        textRange.paragraphFormat.horizontalAlignment = value;
      }
      if (hasOwn(desired, 'verticalAlignment')) {
        const value = String(desired.verticalAlignment || '');
        if (!['Top', 'Middle', 'Bottom', 'TopCentered', 'MiddleCentered', 'BottomCentered'].includes(value)) throw new Error('verticalAlignment is unsupported');
        shape.textFrame.verticalAlignment = value;
      }
      const marginMap = { marginLeft: 'leftMargin', marginRight: 'rightMargin', marginTop: 'topMargin', marginBottom: 'bottomMargin' };
      for (const [source, destination] of Object.entries(marginMap)) {
        if (hasOwn(desired, source)) shape.textFrame[destination] = finiteNumber(desired[source], source);
      }
      if (hasOwn(desired, 'wordWrap')) shape.textFrame.wordWrap = Boolean(desired.wordWrap);
      if (hasOwn(desired, 'autoSizeSetting')) shape.textFrame.autoSizeSetting = desired.autoSizeSetting;
    }
    if (desired.zOrder) shape.setZOrder(desired.zOrder[0].toUpperCase() + desired.zOrder.slice(1));
  }

  // 按 shapeId 查找形状：先查顶层，找不到再递归查各组合（Group）的子形状——
  // 组合内的形状对 AI 可见（读取已递归），写路径找不到会造成「能读不能改」的死路
  async function findShapeById(context, slide, shapeId, loadFields) {
    const top = (slide.shapes.items || []).find(item => item.id === shapeId);
    if (top) return top;
    for (const candidate of (slide.shapes.items || [])) {
      if (String(candidate.type || '') !== 'Group' || !candidate.shapes) continue;
      try {
        candidate.shapes.load('items');
        await context.sync();
        candidate.shapes.items.forEach(child => child.load(loadFields));
        await context.sync();
        const found = (candidate.shapes.items || []).find(item => item.id === shapeId);
        if (found) return found;
      } catch {}
    }
    return null;
  }

  // 【48-R1/49-R2】组合成员证据：可读且非空才作为防误删证据（members 参与防覆盖比对）。
  // 子 Group 递归核验（限深，防异常结构）；成员文字逐个隔离读取——混合组合里个别成员
  // （如图片）没有 textFrame，不能让一个成员的文字请求把整批读数取消。
  // 证据不足（incomplete）的情形：任一成员文字读取异常（本应可读，不得吞成空串）、
  // 嵌套超限、成员为空或集合不可读（无法区分「真空组」与「本机不可见」）——
  // 由调用方限制整组删除，不伪造空成员。
  const GROUP_MEMBER_MAX_DEPTH = 4;
  async function currentGroupMembers(context, shape, depth = 0) {
    if (!shape.shapes || typeof shape.shapes.load !== 'function') throw new Error('no shapes collection');
    if (depth > GROUP_MEMBER_MAX_DEPTH) return { members: [], incomplete: true };
    shape.shapes.load('items');
    await context.sync();
    const kids = shape.shapes.items || [];
    const members = [];
    let incomplete = false;
    for (const k of kids) {
      k.load('id,type,left,top,width,height');
      await context.sync();
      const memberType = String(k.type || '');
      if (memberType === 'Group') {
        const nested = await currentGroupMembers(context, k, depth + 1);
        if (nested.incomplete) incomplete = true;
        members.push(...nested.members);
        continue;
      }
      const record = { id: String(k.id), type: memberType, left: Number(k.left), top: Number(k.top), width: Number(k.width), height: Number(k.height) };
      const textless = (SHAPE_CAPABILITY_LIMITS[memberType] || []).includes('textFrame');
      if (textless) {
        record.text = '';   // 该类型本无文字框架：文字缺席是事实，不是读取失败（49-R2 与异常区分）
      } else {
        try {
          k.textFrame.textRange.load('text');
          await context.sync();
          record.text = String(k.textFrame.textRange.text || '');
        } catch {
          record.textUnreadable = true;   // 本应可读文字的成员读取异常：不吞成空串，整体证据按不足处理
          incomplete = true;
        }
      }
      members.push(record);
    }
    return { members, incomplete };
  }

  function groupMembersMismatched(expectedMembers, currentMembers) {
    const byId = new Map(currentMembers.map(m => [m.id, m]));
    const mismatched = [];
    for (const m of expectedMembers) {
      const cur = byId.get(m.id);
      if (!cur) { mismatched.push(String(m.id)); continue; }
      if (String(m.text || '') !== String(cur.text || '') || String(m.type || '') !== String(cur.type || '')
        || ['left', 'top', 'width', 'height'].some(key => Math.abs(Number(m[key]) - Number(cur[key])) > 0.25)) {
        mismatched.push(String(m.id));
      }
    }
    for (const cur of currentMembers) {
      if (!expectedMembers.some(m => m.id === cur.id)) mismatched.push(cur.id);
    }
    return mismatched;
  }

  // 【43】把圆角半径钉成绝对磅值：PPT 圆角默认是 adj×min(w,h) 的相对值，不同高度的
  // 形状（高外框 vs 矮色条）默认半径必然不同——真机 B 的「尖角硬拼」根因之一。
  // 写后读回验证实际生效（声明支持 1.10 ≠ 真生效）；读回不符/中途失败删除残形再报错。
  async function setAbsoluteCornerRadius(context, shape, desired) {
    const radius = Number(desired.cornerRadius);
    const minDim = Math.min(Number(desired.width), Number(desired.height));
    const adj = Math.min(0.5, radius / minDim);
    // 【45-R2】delete 只是排队，PowerPoint.run 异常退出不保证隐式 sync——失败清理必须显式
    // 等待删除提交后才报错；清理失败/未确认时报错带 shapeId 与清理状态，不得宣称「已清理」
    // 或让人按「可安全重试」继续。
    let shapeId = null;
    const readShapeId = async () => {
      try { shape.load('id'); await context.sync(); shapeId = String(shape.id); } catch { /* id 读不到：清理状态里如实说未知 */ }
    };
    // true=删除已提交；false=删除失败或未确认（残形可能仍在页面上）
    const cleanupLeftover = async () => {
      try {
        shape.delete();
        await context.sync();
        return true;
      } catch { return false; }
    };
    const leftoverSuffix = cleaned => cleaned
      ? '残形已删除并提交。'
      : `删除残形未确认提交（shapeId=${shapeId || '未能读取'}），残形可能仍在页面上：不要按「已清理」继续重试，先核对该形状并删除，或调整设计后重试。`;
    const attachWriteState = (error, shapeCreated, cleaned) => {
      // 【48-R3/R4】写入状态显式化，供工具结果与失败汇总区分「确认未写入/已回滚/清理未确认」
      error.shapeCreated = shapeCreated;
      error.writeState = cleaned ? 'rolled_back' : 'cleanup_unconfirmed';
      if (shapeId) error.shapeId = shapeId;
      return error;
    };
    try {
      if (!shape.adjustments || typeof shape.adjustments.set !== 'function') throw new Error('adjustments unavailable');
      shape.adjustments.set(0, adj);
      await context.sync();
      const back = shape.adjustments.get(0);
      await context.sync();
      const actualAdj = Number(back && back.value);
      const actualRadius = actualAdj * minDim;
      const ok = Number.isFinite(actualAdj) && Math.abs(actualRadius - radius) < 0.5;
      noteRuntimeProbe('shapeAdjustments', ok);
      if (!ok) {
        await readShapeId();
        const cleaned = await cleanupLeftover();
        throw attachWriteState(new Error(`圆角半径写后读回不符（请求 ${radius}pt，实际 ${actualRadius}pt）——本机 adjustments 未真实生效。${leftoverSuffix(cleaned)}不要降级硬拼，改直角设计。`), true, cleaned);
      }
      return { requested: radius, actual: Math.round(actualRadius * 100) / 100, adjustment: actualAdj };
    } catch (e) {
      if (e && /圆角半径写后读回不符/.test(e.message)) throw e;   // 读回不符错误已带清理状态，就地终结
      noteRuntimeProbe('shapeAdjustments', false);
      await readShapeId();
      const cleaned = await cleanupLeftover();
      throw attachWriteState(new Error('本机圆角半径设置失败（' + ((e && e.message) || e) + '）。' + leftoverSuffix(cleaned) + '不要改用默认圆角硬拼：把卡片与色条都改成直角矩形，或调整设计后重试。'), 'unknown', cleaned);
    }
  }

  // 【48-R3】addShape 分阶段诊断：创建/属性设置（提交前同步异常）与提交（sync）失败分开；
  // 完整去敏参数与目标、Office code/message/debugInfo 全量随错误对象字段回传（api.js 拷贝进
  // 工具结果）；写入状态显式化——not_committed=异常在任何提交之前（零写入）/ unknown=提交
  // 阶段失败需 get_slide 核对，不用「通常没创建」这类含糊措辞代替状态。
  function enrichAddShapeError(e, desired, target, phase, committed) {
    const office = {};
    if (e && e.code) office.code = e.code;
    if (e && e.message) office.message = String(e.message);
    if (e && e.debugInfo !== undefined) {
      try { office.debugInfo = JSON.parse(JSON.stringify(e.debugInfo)); } catch { office.debugInfo = String(e.debugInfo); }
    }
    const detail = [];
    if (office.code) detail.push(`code=${office.code}`);
    if (office.message) detail.push(office.message);
    const stateText = committed
      ? '未知（提交阶段失败：先 get_slide 该页核对此位置有无残留形状，避免重复新增）'
      : '未写入（异常发生在任何提交之前，零写入）';
    const err = new Error(`新增形状失败（阶段 ${phase}；${desired.shapeType}${desired.name ? `「${desired.name}」` : ''}，目标 slide ${target.slideId != null ? target.slideId : `index ${target.index}`}）：${detail.join('；') || 'Office 未返回详细原因'}。形状是否落盘：${stateText}。不要在未核对/未解决前按原参数原样重试，也不要把这个失败当成已成功继续汇报。`);
    err.phase = phase;
    err.target = { slideId: target.slideId != null ? target.slideId : null, index: target.index != null ? target.index : null };
    err.shape = desired;            // 完整去敏参数（原始字符串/数值/布尔）
    err.office = office;
    err.shapeCreated = committed ? 'unknown' : false;
    err.writeState = committed ? 'unknown' : 'not_committed';
    return err;
  }

  // existingOp【31-A】：上层入口（applyLayoutDirect/提案卡）已建的操作上下文——
  // 本函数复用同一令牌与计数，不重复登记、不重新捕获取消代次。
  async function applyLayout(edit, existingOp) {
    requireOffice();
    const { operation, target, desired } = validateLayoutEdit(edit);
    const capabilities = presentationCapabilities();
    if (!capabilities.shapeLayout) throw new Error('当前 PowerPoint 不支持形状布局 API 1.4，未执行任何修改。');
    if (desired.zOrder && !capabilities.zOrder) throw new Error('当前 PowerPoint 不支持层级 API 1.8，未执行任何修改。');

    // 【R1/30.2】布局写入与文本写入同生命周期：在途计数+取消代次；检查点在各提交点前
    const op = existingOp || App.docWriteBegin('ppt.applyLayout');
    try {
    return await PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      slides.items.forEach(slide => slide.load('id'));
      await context.sync();
      const slide = target.slideId ? slides.items.find(item => item.id === target.slideId) : slides.items[Number(target.index)];
      if (!slide) throw App.makeStaleEditError('目标幻灯片已移动或删除，请重新读取此项。', { target });
      // target 只给 slideId 时 Number(target.index) 是 NaN，复检会被跳过；用定位到的真实下标代替
      const slideIndex = slides.items.findIndex(item => item === slide);

      if (operation.type === 'addShape') {
        // 【31-A】取消检查先于任何宿主变更 API：add 形状本身就是变更，检查必须在 add 之前
        op.token.throwIfCancelled();
        // 【43】cornerRadius 需要 adjustments（PowerPointApi 1.10）：声明不支持就不建形，
        // 避免留下「默认圆角」残形（声明支持但实测失败的情形在建形后删除残形再报错）
        if (hasOwn(desired, 'cornerRadius') && !powerPointApiSupported('1.10')) {
          noteRuntimeProbe('shapeAdjustments', false);
          throw new Error('本机不支持设置绝对圆角半径（Shape.adjustments 需要 PowerPointApi 1.10）。不要改用默认圆角硬拼：把卡片与色条都改成直角矩形，或调整设计后重试。');
        }
        // 【48-R3】创建与属性设置是提交前的同步代理调用，异常与提交阶段分开诊断
        const options = shapeOptions(desired);
        let shape;
        try {
          if (desired.shapeType === 'textBox') shape = slide.shapes.addTextBox(String(desired.text || ''), options);
          else if (desired.shapeType === 'line') shape = slide.shapes.addLine('Straight', options);
          else {
            // 【43】cornerStyle:'top' → 原生上圆下直形状（Round2SameRectangle，PowerPointApi 1.4，
            // 文档枚举可编辑形状）；同色填充+无描边叠在外框顶部即无缝卡头色条
            const typeMap = { rectangle: 'Rectangle', roundRectangle: 'RoundRectangle', ellipse: 'Ellipse' };
            const geoType = desired.shapeType === 'roundRectangle' && desired.cornerStyle === 'top' ? 'Round2SameRectangle' : typeMap[desired.shapeType];
            shape = slide.shapes.addGeometricShape(geoType, options);
          }
          applyShapeFormatting(shape, desired);
        } catch (e) {
          throw enrichAddShapeError(e, desired, target, 'create-format', false);
        }
        let cornerRadiusResult = null;
        try {
          if (hasOwn(desired, 'cornerRadius')) {
            cornerRadiusResult = await setAbsoluteCornerRadius(context, shape, desired);
          }
          shape.load('id');
          await context.sync();
        } catch (e) {
          // 【47.4/48-R3】圆角链路的错误已带结构化写入状态（shapeCreated/writeState/shapeId），
          // 原样上抛不重复包装；提交阶段失败分阶段诊断并显式标注写入状态
          if (e && /圆角半径|写后读回不符/.test(String(e && e.message))) throw e;
          throw enrichAddShapeError(e, desired, target, 'commit-sync', true);
        }
        return { success: true, operation: operation.type, shapeId: shape.id, index: slideIndex, _navTarget: { slideId: slide.id }, ...(cornerRadiusResult ? { cornerRadius: cornerRadiusResult } : {}) };
      }

      slide.shapes.load('items');
      await context.sync();
      slide.shapes.items.forEach(shape => shape.load('id,name,type,left,top,width,height'));
      await context.sync();
      const shape = await findShapeById(context, slide, target.shapeId, 'id,name,type,left,top,width,height');
      // 【47.3】「找不到」如实说找不到：区分已删除与「在组合内部本宿主读不到」两种可能，
      // 带上 slideId/shapeId，不让模型把它和属性变化混为一谈、也不让它原样重试
      if (!shape) {
        throw App.makeStaleEditError(`目标形状未找到（slideId=${target.slideId || ''}，shapeId=${target.shapeId}）：可能已被删除，也可能位于组合内部——本宿主读不到组合内的形状。请 get_slide 该页核对；若形状仍在，按其现状重出提案；不要按原参数原样重试。`, { target });
      }
      const current = await currentShapeSnapshot(context, shape, operation.expected);
      const mismatched = shapeSnapshotMismatchedFields(operation.expected, current);
      if (mismatched.length) {
        throw App.makeStaleEditError(`目标形状在生成提案后发生变化（${mismatched.join('/')}），为避免覆盖你的修改，此项没有应用。`, { target, currentShape: current, mismatchedFields: mismatched });
      }
      // 【R1/30.2】定位与 expected 校验之后、格式/删除提交之前检查取消
      op.token.throwIfCancelled();
      if (operation.type === 'deleteShape') {
        // 【48-R1】组合删除的成员级防覆盖：外框相同不代表组内未变
        if (hasOwn(operation.expected, 'members')) {
          // 【50-R1】保留完整读取结果：currentGroupMembers 的成员文字读取失败不抛异常，
          // 而是返回 incomplete=true——外层 catch 抓不到它。原文字为空时丢弃 incomplete
          // 会把「现在读不到」误当成「仍然为空」后执行删除。只有证据完整 + 成员比对通过
          // + 同一操作 token 未被取消，才允许删除（与出卡阶段同一「证据完整」判定）。
          let evidence;
          try {
            evidence = await currentGroupMembers(context, shape);
          } catch (e) {
            // 【49-R2】应用期成员读取失败 ≠ 成员未变：按证据不足受限，不盲删
            throw App.makeStaleEditError(`组合成员在应用时读取失败（${String((e && e.message) || e).slice(0, 160)}），无法核验组内现状，此项没有应用。请重新读取后重出提案。`, { target });
          }
          if (evidence.incomplete) {
            throw App.makeStaleEditError(`组合成员在应用时读取不完整（存在读不到的文字/结构），无法核验组内现状，此项没有应用、保留该组合。请重新读取后重出提案；若仍不可核验，请让用户手动处理该组合。`, { target });
          }
          // 【49-R3】成员核验经历多次 sync：读取完成后、变更提交前重新检查取消代次——
          // 停止落在核验窗口内时不得继续删除（同一 op.token，不重建代次）
          op.token.throwIfCancelled();
          const currentMembers = evidence.members;
          const changedMembers = groupMembersMismatched(operation.expected.members, currentMembers);
          if (changedMembers.length) {
            throw App.makeStaleEditError(`组合成员在生成提案后发生变化（${changedMembers.join('、')}），为避免覆盖组合内的修改，此项没有应用。`, { target, mismatchedMembers: changedMembers, currentShape: { members: currentMembers } });
          }
        }
        // 【48-R1】成员不可核验（读不到或为空）：明确受限处理——保留组合、列为未完成，
        // 不把常规卡片确认当成绕过防覆盖核验的手段
        if (hasOwn(operation.expected, 'groupMembersUnreadable')) {
          throw new Error('组合内容在本宿主不可核验（成员读不到），为防止覆盖组合内的修改，此项没有应用、保留该组合。请让用户在 PowerPoint 中手动删除该组合，或先解组后对顶层形状重出提案。此项列为未完成项，向用户如实说明；不要按原参数重试。');
        }
        shape.delete();
      } else {
        applyShapeFormatting(shape, desired);
      }
      await context.sync();
      return { success: true, operation: operation.type, shapeId: target.shapeId, index: slideIndex, _navTarget: { slideId: slide.id } };
    });
    } finally {
      if (!existingOp) op.end();
    }
  }

  function validateEditProposal(edit) {
    if (!(edit && (edit.kind === 'layout' || edit.operation))) return true;
    // 返回具体失败原因而不是笼统的「不完整」——错误文本会作为工具结果回喂模型，
    // 它看到「缺 shapeId」就知道该先去 get_slide，能自我修正，不需要人介入。
    try { validateLayoutEdit(edit); return true; }
    catch (e) { return String((e && e.message) || e); }
  }

  // ================= 文字提案锚点核验（出卡前 + 应用时共用） =================
  // 【harden-ppt-excel-recovery 2.1】此前文字卡（省略 kind 与 target.kind:"notes"）只在
  // applyEdit 里做匹配：错字 find、过期 shapeId、重复命中要等用户点卡后才失败，报错又缺
  // 当前文本对照。现在出卡前先核验（任一项不合法即拒绝整卡，只读零写入），应用时再核验保留。
  // 定位骨架只有这一份：locateTextEditTarget 同时服务 enrich 与 applyEdit，避免两套语义。

  // 错误里的引文：超长文本显式截断并标注原文长度，不让模型把截断尾巴当原文
  function quoteAnchorText(value, rereadTool) {
    const text = String(value == null ? '' : value);
    if (text.length <= 500) return text;
    return text.slice(0, 500) + `…（已截断：原文 ${text.length} 字，请用 ${rereadTool || 'get_slide'} 读取完整原文）`;
  }

  // 文字卡目标定位：slideId/index → 页；notes 走备注形状扫描，普通文字按 shapeId 限定
  // （顶层 + 组合递归）或扫描顶层 + 组合子形状；findSafeTextMatch 唯一命中才算定位成功。
  // 返回定位证据（候选/唯一命中/同框多中/不可读），失败原因由 describeTextAnchorProblem 统一描述。
  async function locateTextEditTarget(context, edit) {
    const target = edit.target || {};
    const find = String(edit.find == null ? '' : edit.find);
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    slides.items.forEach(s => s.load('id'));
    await context.sync();
    const slide = target.slideId ? slides.items.find(s => s.id === target.slideId) : slides.items[Number(target.index)];
    const safeId = shape => { try { return shape && shape.id != null ? String(shape.id) : ''; } catch { return ''; } };
    const safeType = shape => { try { return shape && shape.type != null ? String(shape.type) : ''; } catch { return ''; } };
    const result = {
      target, find, slide,
      slideIndex: slide ? slides.items.indexOf(slide) : Number(target.index),
      slidesCount: slides.items.length,
      kind: target.kind === 'notes' ? 'notes' : 'shape',
      candidates: [], located: [], ambiguous: [], slideShapeIds: []
    };
    if (!slide) return result;
    const classify = entry => {
      result.candidates.push(entry);
      if (entry.unreadable) return;
      const match = App.findSafeTextMatch(entry.text, find);
      if (match.status === 'matched') result.located.push(entry);
      else if (match.status === 'ambiguous') result.ambiguous.push(Object.assign({ count: match.count }, entry));
    };
    if (result.kind === 'notes') {
      let notes;
      try { notes = slide.notesSlide; notes.shapes.load('items'); await context.sync(); }
      catch { result.notesUnavailable = true; return result; }
      for (const shape of notes.shapes.items || []) {
        const read = await shapeTextDetailed(context, shape);
        classify({ shape, shapeId: safeId(shape), type: safeType(shape), text: read.text, unreadable: read.unreadable });
      }
      return result;
    }
    slide.shapes.load('items');
    await context.sync();
    slide.shapes.items.forEach(s => s.load('id,type'));
    await context.sync();
    result.slideShapeIds = (slide.shapes.items || []).map(safeId).filter(Boolean);
    let shapes;
    if (target.shapeId) {
      const found = await findShapeById(context, slide, target.shapeId, 'id');
      shapes = found ? [found] : [];
    } else {
      // 候选形状 = 顶层形状 + 各组合的子形状：读取侧对组合已递归可见，写侧必须能定位回去
      shapes = [...slide.shapes.items];
      for (const candidate of slide.shapes.items) {
        if (String(candidate.type || '') !== 'Group' || !candidate.shapes) continue;
        try {
          candidate.shapes.load('items');
          await context.sync();
          candidate.shapes.items.forEach(child => { try { child.load('id'); } catch {} });
          await context.sync();
          shapes.push(...(candidate.shapes.items || []));
        } catch {}
      }
    }
    for (const shape of shapes) {
      const read = await shapeTextDetailed(context, shape);
      classify({ shape, shapeId: safeId(shape), type: safeType(shape), text: read.text, unreadable: read.unreadable });
    }
    return result;
  }

  // 统一失败原因（出卡前拒绝与应用时拦截同语义）：必含提交 find 全文、目标当前文本证据
  // （>500 字显式截断）与下一步读取动作。重复命中给正确消歧：形状跨框用 target.shapeId，
  // 同框与「一框唯一 + 另框多次」的混合歧义都要求更长唯一片段，绝不静默取第一处、不猜唯一框；
  // 备注锚点只能靠文字定位，不得推荐无效的 shapeId 消歧。返回 null 表示锚点合法（全页唯一）。
  function describeTextAnchorProblem(edit, loc) {
    const target = loc.target || {};
    const find = String(edit.find == null ? '' : edit.find);
    const isNotes = loc.kind === 'notes';
    const reread = isNotes ? 'get_slide_notes' : 'get_slide';
    const label = edit.label ? `「${String(edit.label).slice(0, 30)}」` : '';
    const findLine = `你提交的 find 全文：「${quoteAnchorText(find, reread)}」`;
    if (!find.trim()) {
      return `${label}缺 find：文字卡必须锚定页面上真实存在的文字，程序不会猜目标（${findLine}）。纯新增文字请改用 insert_textbox 直通写入，不走提案卡。`;
    }
    if (!loc.slide) {
      const at = target.slideId ? `slideId=${target.slideId}` : `target.index=${target.index == null ? '未提供' : target.index}`;
      return `${label}目标页不存在（${at}，当前演示文稿共 ${loc.slidesCount} 页）。${findLine}请先用 get_presentation_outline 或 get_slide 核对页号，按现状重出提案。`;
    }
    if (loc.notesUnavailable) {
      return `${label}无法访问第 ${loc.slideIndex} 页的备注页：这页可能还没有备注实例（不是宿主缺少备注 API，也不等于备注为空）。${findLine}请先用 get_slide_notes 确认该页确有备注文字；没有备注的页不要出 target.kind:"notes" 的卡。`;
    }
    if (!isNotes && target.shapeId && !loc.candidates.length) {
      const ids = (loc.slideShapeIds || []).join('、') || '（无）';
      return `${label}第 ${loc.slideIndex} 页找不到形状 shapeId「${target.shapeId}」——形状可能已被移动或删除。本页现有形状 id：${ids}。${findLine}请先用 get_slide 重读该页拿到当前形状列表，按现状重出提案。`;
    }
    if (!isNotes && target.shapeId && loc.candidates.length === 1 && loc.candidates[0].unreadable && String(loc.candidates[0].type || '') === 'Table') {
      return `${label}目标形状「${target.shapeId}」是表格（Table）：文字卡定位不了表格单元格。表格文字在各单元格里，形状没有整体 textFrame，get_slide 深读到的表格文字不能用作 propose_edits 的 find，按原样重出文字卡必然再次失败。替代路径：单元格字体/字号/字色/加粗/底色用 set_table_style；加删行列用 edit_table_structure（多数宿主没有该 API，本机能力探测通过才可用）；表格整体位置尺寸用 apply_layout 改 left/top/width/height；要改单元格里的文字内容，本插件没有专用工具，请如实告诉用户需要手动处理。${findLine}`;
    }
    if (!isNotes && target.shapeId && loc.candidates.length === 1 && loc.candidates[0].unreadable) {
      return `${label}无法读取形状「${target.shapeId}」的文字（读取失败，不是空文本：组合、图片等对象可能没有本宿主可读的文字框架）。${findLine}请先用 get_slide 读取该形状的当前状态；读不到文字的形状不要出文字卡。`;
    }
    if (isNotes && (loc.located.length > 1 || loc.ambiguous.length)) {
      const parts = [];
      for (const c of loc.located) parts.push(`一个备注文字形状唯一命中（当前文本：「${quoteAnchorText(c.text, reread)}」）`);
      for (const a of loc.ambiguous) parts.push(`一个备注文字形状内出现 ${a.count} 次（当前文本：「${quoteAnchorText(a.text, reread)}」）`);
      return `${label}find 在第 ${loc.slideIndex} 页备注内定位不唯一：${parts.join('；')}。程序不会静默取第一处；备注锚点只能靠文字本身定位。${findLine}请把 find 扩长为包含足够上下文、在整个备注内唯一的片段。`;
    }
    if (loc.located.length > 1) {
      const list = loc.located.map(c => `shapeId「${c.shapeId}」当前文本：「${quoteAnchorText(c.text, reread)}」`).join('；');
      return `${label}find 在 ${loc.located.length} 个文本框中各自唯一命中，无法确定要改哪一个：${list}。${findLine}请用 target.shapeId 限定目标，或改用只在目标文本框中出现的更长原文。`;
    }
    if (loc.ambiguous.length) {
      // 同框多次命中；或一框唯一命中 + 另一框多次命中（混合歧义）——唯一的一框未必是目标，不得猜
      const parts = [];
      if (loc.located.length === 1) parts.push(`shapeId「${loc.located[0].shapeId}」唯一命中（当前文本：「${quoteAnchorText(loc.located[0].text, reread)}」）`);
      for (const a of loc.ambiguous) parts.push(`shapeId「${a.shapeId}」内出现 ${a.count} 次（当前文本：「${quoteAnchorText(a.text, reread)}」）`);
      const hint = target.shapeId
        ? '请把 find 扩长为包含足够上下文、在该框内唯一的原文片段。'
        : '请用 target.shapeId 指定目标文本框，且 find 在该框内必须只出现 1 次（必要时扩长为带上下文的唯一片段）。';
      return `${label}find 在第 ${loc.slideIndex} 页定位不唯一：${parts.join('；')}。程序不会静默取第一处，也不会替你猜哪个框才是目标。${findLine}${hint}`;
    }
    if (!loc.located.length) {
      const readable = loc.candidates.filter(c => !c.unreadable);
      const shown = readable.slice(0, 6).map(c => `shapeId「${c.shapeId}」当前文本：「${quoteAnchorText(c.text, reread)}」`).join('\n');
      const more = readable.length > 6 ? `\n（其余 ${readable.length - 6} 个文本框文字未列出，请用 ${reread} 读取）` : '';
      const unreadableCount = loc.candidates.length - readable.length;
      const unreadNote = unreadableCount > 0 ? `\n另有 ${unreadableCount} 个形状文字读取失败（读取失败不是空文本，无法核验）。` : '';
      const tableNote = loc.candidates.some(c => String(c.type || '') === 'Table')
        ? '\n该页有表格：表格文字在单元格里（get_slide 能深读到），文字卡不能定位表格；表格样式用 set_table_style，加删行列用 edit_table_structure，整体位置尺寸用 apply_layout，改单元格文字内容需手动处理。'
        : '';
      const scope = isNotes
        ? `该页备注页共扫描 ${loc.candidates.length} 个形状`
        : target.shapeId ? '目标形状的文字如下' : `第 ${loc.slideIndex} 页共扫描 ${loc.candidates.length} 个文本框`;
      const evidence = readable.length ? `${scope}：\n${shown}${more}${unreadNote}${tableNote}` : `${scope}，其中没有可读出文字的形状。${unreadNote}${tableNote}`;
      return `${label}find 未在目标范围命中。${evidence}\n${findLine}请先用 ${reread} 读取目标范围完整原文，把 find 改成与上述当前文本之一逐字一致的片段。`;
    }
    return null;
  }

  // 出卡前核验：任一文字项定位不合法即抛错拒绝整卡（api.js 把错误转成工具结果回喂模型，不出卡）。
  // 只读零写入；不猜测、不自动替换模型提交的 find/target。
  async function verifyTextAnchorsBeforeCard(textItems) {
    await PowerPoint.run(async context => {
      const problems = [];
      for (let i = 0; i < textItems.length; i++) {
        try {
          const loc = await locateTextEditTarget(context, textItems[i]);
          const reason = describeTextAnchorProblem(textItems[i], loc);
          if (reason) problems.push(`第 ${i + 1} 项文字项${reason}`);
        } catch (e) {
          const rereadHint = textItems[i] && textItems[i].target && textItems[i].target.kind === 'notes' ? 'get_slide_notes' : 'get_slide';
          problems.push(`第 ${i + 1} 项文字项现状读取失败（${String((e && e.message) || e).slice(0, 160)}）：请先用 ${rereadHint} 核对目标现状后重出提案。`);
        }
      }
      if (problems.length) {
        throw new Error(`PPT 文字提案在出卡前核验未通过，已整卡拒绝（未写入任何内容）：\n${problems.join('\n')}\n请按各问题的当前证据修正后，把整批修改一次性重出提案；不要按原参数原样重试，也不要凭记忆微调锚点。`);
      }
    });
  }

  // PPT 版上下文核实（对应 Word 的 enrichEditProposal）：layout 卡片的 operation.expected
  // 由程序在出卡时从活文档实时读取，模型自己编的快照一律覆盖。两个目的：
  // 1) 模型不必背诵精确快照（它根本拿不准），消灭「提案内容不完整」的反复打回；
  // 2) 快照只保留与本次改动相关的字段组（+身份字段），批量卡逐项应用时，
  //    不会被无关变化（例如先应用了加背景矩形导致层级编号偏移）级联误报 stale。
  // 【47.2】逐项隔离 + 按能力读字段：一项读取失败只记到该项（_snapshotError：对象/阶段/
  // 原因/可重试性），不再被外层 catch 一票吞掉、连累同批其余项全部丢快照；Group 等无
  // textFrame 的类型不再强行读文字（真机整批取消的根因），快照保留可核验的身份+几何，
  // 不伪造空文本、不让模型猜 expected。
  async function enrichEditProposal(args) {
    const items = [...((args && args.edits) || []), ...((args && args.changes) || [])];
    const targets = items.filter(item => item && (item.kind === 'layout' || item.operation)
      && item.operation && item.operation.type !== 'addShape'
      && item.target && (item.target.shapeId || item.target.slideId || item.target.index != null));
    const textItems = items.filter(item => item && !(item.kind === 'layout' || item.operation));
    if (typeof PowerPoint === 'undefined') return args;
    // 【harden-ppt-excel-recovery 2.1】文字项（含省略 kind 与 target.kind:"notes"）出卡前先核验：
    // 任一项无法唯一定位即抛错拒绝整卡，错误带 find 全文、目标当前文本与重读指引回喂模型。
    if (textItems.length) await verifyTextAnchorsBeforeCard(textItems);
    if (!targets.length) return args;
    const snapshotErrorOf = (target, reason, retryable) => ({ object: describeEditTarget(target), phase: 'enrich', reason, retryable });
    try {
      await PowerPoint.run(async context => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        slides.items.forEach(s => s.load('id'));
        await context.sync();
        for (const item of targets) {
          const target = item.target;
          try {
            const slide = target.slideId ? slides.items.find(s => s.id === target.slideId) : slides.items[Number(target.index)];
            if (!slide) {
              item.operation._snapshotError = snapshotErrorOf(target, `目标页不存在（共 ${slides.items.length} 页）——提案生成后页面可能被删除或重排。请重新 get_slide 后针对现状重出提案。`, true);
              continue; // 记录到该项，由校验按结构化信息打回，不再静默留空
            }
            slide.shapes.load('items');
            await context.sync();
            slide.shapes.items.forEach(s => s.load('id,name,type,left,top,width,height'));
            await context.sync();
            const shape = target.shapeId ? await findShapeById(context, slide, target.shapeId, 'id,name,type,left,top,width,height') : null;
            if (!shape) {
              item.operation._snapshotError = snapshotErrorOf(target, '目标形状未找到：可能已被删除，也可能位于组合内部（本宿主读不到组合内的形状）。请 get_slide 该页核对；若形状仍在顶层，按其现状重出提案。', true);
              continue;
            }
            const wanted = { name: true, type: true };
            const desired = item.operation.shape || {};
            if (item.operation.type === 'deleteShape' || !Object.keys(desired).length) {
              Object.assign(wanted, { text: true, left: true, top: true, width: true, height: true });
            } else {
              for (const key of Object.keys(desired)) {
                for (const field of (EXPECTED_BY_DESIRED[key] || [])) wanted[field] = true;
              }
            }
            // 【47.2】按能力处理：desired（真正要写什么）撞上能力限制 → 明确不可重试的原因；
            // wanted（要读什么身份证据）撞上限制 → 只剥离字段，不拒绝（组合删除仍然合法，
            // 保留身份+几何防误删，不伪造空文本）
            const limits = SHAPE_CAPABILITY_LIMITS[String(shape.type || '')] || [];
            const desiredTextDemand = item.operation.type === 'updateShape' && TEXTFRAME_FIELDS.some(key => hasOwn(desired, key));
            const desiredFillDemand = item.operation.type === 'updateShape' && FILL_FIELDS.some(key => hasOwn(desired, key));
            if (limits.includes('textFrame') && desiredTextDemand) {
              const label = String(shape.type || '') === 'Group' ? '组合（Group）没有 textFrame：本宿主读不到组合内文字，也无法读写组内内容' : `${String(shape.type || '')} 类型没有 textFrame：不能读取或设置文字字段`;
              item.operation._snapshotError = snapshotErrorOf(target, `${label}。请改为对组合外/有文字框架的形状操作，或在提案中明确说明该项需要用户手动处理。`, false);
              continue;
            }
            if (limits.includes('fill') && desiredFillDemand) {
              item.operation._snapshotError = snapshotErrorOf(target, `${String(shape.type || '')} 类型没有可读填充：不能核验 fillColor/fillType。`, false);
              continue;
            }
            for (const field of Object.keys(wanted)) {
              if (limits.includes('textFrame') && TEXTFRAME_FIELDS.includes(field)) delete wanted[field];
              if (limits.includes('fill') && FILL_FIELDS.includes(field)) delete wanted[field];
            }
            item.operation.expected = await currentShapeSnapshot(context, shape, wanted);
            // 【49-R1】成员证据只对真实 Group 生效：Picture/Line/Chart/Table 同样没有
            // textFrame，但不是组合——按自身能力保留快照与删除路径，不得标
            // groupMembersUnreadable 误伤普通形状删除
            if (String(shape.type || '') === 'Group') {
              // 【48-R1/49-R2】成员证据：可读且非空才捕获为 members 参与防覆盖比对；
              // 子 Group 递归核验（限深），任一证据不足（读取失败/嵌套过深/为空，无法区分
              // 「真空组」与「本机不可见」）→ groupMembersUnreadable，由应用侧限制整组删除——
              // 不伪造空成员，不把外框相同解释成组内未变
              try {
                const evidence = await currentGroupMembers(context, shape);
                if (evidence.incomplete || !evidence.members.length) item.operation.expected.groupMembersUnreadable = true;
                else item.operation.expected.members = evidence.members;
              } catch { item.operation.expected.groupMembersUnreadable = true; }
            }
          } catch (e) {
            item.operation._snapshotError = snapshotErrorOf(target, `现状读取失败（${String((e && e.message) || e).slice(0, 200)}）。`, false);
          }
        }
      });
    } catch (e) {
      // 【47.2】run 级失败（页面级读取取消等）：逐项记录，不再静默吞掉让校验只能给笼统报错
      const reason = `现状批量读取失败（${String((e && e.message) || e).slice(0, 200)}）。`;
      for (const item of targets) {
        if (!item.operation.expected && !item.operation._snapshotError) {
          item.operation._snapshotError = snapshotErrorOf(item.target, reason, true);
        }
      }
    }
    return args;
  }

  // 【图片链路】insert_image 的图片来源解析（Word/PPT 同款，两宿主各自带一份小实现——
  // 夹具只加载宿主文件，不能依赖 host.js 之外的共享层）：
  //   1. base64（兼容旧用法）
  //   2. attachmentId —— 用户经 ➕ 上传图片/拖拽进对话区的图片（App.state.pendingImages），
  //      消息里只注入了 id 引用，base64 留在本地（防 token 爆炸）
  //   3. url —— 图片直链，走本地服务 /api/fetch-image 下载（WKWebView 直连外网图片会被 CORS 拦）
  async function resolveImageSource(args) {
    const { base64, attachmentId, url } = args || {};
    if (base64) return String(base64).replace(/^data:image\/\w+;base64,/, '');
    if (attachmentId) {
      // 查找范围：UI 挂起区（还没发送的）+ 本会话各条消息携带的图片（IMG-6：发送时数据从
      // state 移交到消息对象上，本轮 agent loop 执行时从这里取）
      const fromState = (App.state && Array.isArray(App.state.pendingImages)) ? App.state.pendingImages : [];
      const msgs = (App.state && Array.isArray(App.state.messages)) ? App.state.messages : [];
      const fromMessages = [];
      for (const m of msgs) if (m && Array.isArray(m.pendingImages)) for (const img of m.pendingImages) fromMessages.push(img);
      const list = fromMessages.concat(fromState);
      const found = list.find(p => p && p.id === String(attachmentId));
      if (!found) {
        throw new Error(`attachmentId "${attachmentId}" 未找到（已上传的图片只在当次插件会话内有效，重开侧边栏或更换会话后失效）。当前可用：${list.length ? list.map(p => p.id + '(' + p.name + ')').join('、') : '无'}。请让用户重新拖一张图进对话框。`);
      }
      return String(found.dataUrl || '').replace(/^data:image\/\w+;base64,/, '');
    }
    if (url) {
      const target = 'https://localhost:18443/api/fetch-image?url=' + encodeURIComponent(url);
      const endpoint = typeof App.localApiUrl === 'function' ? App.localApiUrl(target) : target;
      const resp = await fetch(endpoint).catch(() => null);
      const data = resp ? await resp.json().catch(() => null) : null;
      if (!resp || !resp.ok || !data || !data.base64) {
        throw new Error(`图片下载失败${resp ? '（HTTP ' + resp.status + '）' : '（本地服务未响应）'}：${(data && data.error) || url}。替代路径：让用户用 ➕ 菜单上传图片或直接拖进对话区，然后用 attachmentId 插入。`);
      }
      return data.base64;
    }
    throw new Error('需要图片来源之一：base64 / attachmentId（用户上传的图片，推荐）/ url（图片直链）。用户上传过图片时优先用 attachmentId，不要自己生成 base64。');
  }

  async function insertImage(args) {
    requireOffice();
    const base64 = await resolveImageSource(args);
    return new Promise(resolve => {
      try {
        Office.context.document.setSelectedDataAsync(base64, { coercionType: Office.CoercionType.Image }, r => {
          if (r.status === Office.AsyncResultStatus.Succeeded) resolve({ success: true });
          else resolve({ success: false, error: r.error ? r.error.message : 'setSelectedDataAsync failed' });
        });
      } catch (e) { resolve({ success: false, error: e.message || String(e) }); }
    });
  }

  async function gotoSlide(args) {
    requireOffice();
    const index = Number(args.index);
    const slideId = await PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      await context.sync();
      return slide.id;
    });
    const navigated = await selectSlideById(slideId);
    return { success: navigated, index, id: slideId, navigated };
  }

  function selectSlideById(slideId) {
    // slide.id 在不同宿主/版本下可能是纯数字也可能是 "id#creationId" 复合格式，
    // goToByIdAsync 只认其中一种——逐候选尝试，全失败才算导航失败。
    const raw = String(slideId || '');
    const candidates = [raw];
    if (raw.includes('#')) {
      const parts = raw.split('#');
      candidates.push(parts[0], parts[1]);
    }
    return new Promise(resolve => {
      const tryNext = () => {
        const id = candidates.shift();
        if (id == null) return resolve(false);
        try {
          const doc = Office.context.document;
          if (typeof doc.goToByIdAsync !== 'function') return resolve(false);
          doc.goToByIdAsync(id, Office.GoToType.Slide, r => {
            if (r.status === Office.AsyncResultStatus.Succeeded) resolve(true);
            else tryNext();
          });
        } catch { resolve(false); }
      };
      tryNext();
    });
  }

  async function evalOfficeJs(args) {
    requireOffice();
    const code = args.code || '';
    const modifiedSlideIndexes = [...new Set((Array.isArray(args.modifiedSlideIndexes) ? args.modifiedSlideIndexes : [])
      .map(Number).filter(index => Number.isInteger(index) && index >= 0))];
    return PowerPoint.run(async context => {
      const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
      const fn = new AsyncFunction('context', 'PowerPoint', code);
      const result = await fn(context, PowerPoint);
      return { success: true, result: result ?? null, modifiedSlideIndexes };
    });
  }

  async function navigateCitation(ref) {
    requireOffice();
    // 格式: "s:<index>" 幻灯片序号 | "id:<slideId>" 幻灯片 id | 纯数字回退为序号
    if (/^id:/.test(ref)) { const navigated = await selectSlideById(ref.slice(3)); return { success: navigated, navigated }; }
    const index = Number(/^s:/.test(ref) ? ref.slice(2) : ref);
    return gotoSlide({ index });
  }

  async function maybeFollow(result) {
    if (!App.state.settings.followMode || !result) return;
    const nav = result._navTarget;
    if (nav && nav.slideId) { await selectSlideById(nav.slideId).catch(console.warn); }
  }

  // 直调 apply_layout：expected 快照由插件从活文档实时读取，不要求模型提供。
  // expected 的作用是防「提案卡生成之后文档又被改动」，而直调的读与写在同一次调用里，没有这个时间窗。
  // 以前把这份负担压给模型，两条路都是死的：prompt 说「不要自己填 expected」→ 报 missing shape snapshot；
  // 模型改成自己填 → 又要求填齐 fillType/fillColor/fillTransparency 这类字段组，照样打回。
  // 于是模型放弃这个工具，转去手写 eval_officejs——今天下午那十几次失败就是从这里开始的。
  async function applyLayoutDirect(args) {
    const operation = args && args.operation;
    // confirm 模式只把「纯格式」放行。删形状和换文字不是格式，必须走提案卡让用户过目。
    if (currentAccessMode() === 'confirm' && operation) {
      if (operation.type === 'deleteShape') {
        throw new Error('确认模式下不能直接删除形状：请改用 propose_edits 出一张提案卡，让用户确认后再删。重复原样调用仍会被拒绝，第一次被拒就换提案卡路径。');
      }
      if (operation.shape && hasOwn(operation.shape, 'text')) {
        throw new Error('确认模式下不能直接改写文字：几何和配色可以直接应用，文字内容请走 propose_edits 出卡确认。');
      }
    }
    // 【31-A】操作上下文在实际工具入口、第一次异步补读之前创建：补读（enrich）期间
    // 发生的停止必须让本操作失效。若等补读完成后才 docWriteBegin，捕获的是停止之后的
    // 新代次，被停止的旧操作会被当成新操作放行（Astra 独立探针实证）。令牌贯穿补读与
    // applyLayout，避免重复登记和重新捕获代次；带 expected 直入与提案卡路径同一上下文。
    const op = App.docWriteBegin('ppt.apply_layout');
    try {
      // 只在模型没提供 expected 时才自动补。模型主动给了快照就仍然走严格校验——
      // 那说明它是基于某个时点的读数做的决定，过期检测依然有意义。
      if (operation && operation.type !== 'addShape' && !operation.expected) {
        op.token.throwIfCancelled();
        const enriched = await enrichEditProposal({ edits: [args] });
        const edit = (enriched && enriched.edits && enriched.edits[0]) || args;
        return await applyLayout(edit, op);
      }
      return await applyLayout(args, op);
    } finally {
      op.end();
    }
  }

  // 返回「本轮有效访问模式」而不仅是持久化全局设置：全局 auto 但用户本轮要求先审时，
  // api.js 的 requestAccessMode() 会临时降级为 confirm，apply_layout 的删形状/改文字分支
  // 必须读同一个值，否则 confirm 直通工具调进来时闸门仍看到 auto 直接放行（host-word.js gateMode 同款）。
  function currentAccessMode() {
    try {
      if (typeof App.effectiveAccessMode === 'function') return App.effectiveAccessMode();
      const raw = typeof App.currentAccessMode === 'function' ? App.currentAccessMode() : ((App.state && App.state.settings && App.state.settings.accessMode) || 'confirm');
      return raw === 'auto' && typeof App.requestRequiresReview === 'function' && App.requestRequiresReview() ? 'confirm' : raw;
    } catch { return 'confirm'; }
  }

  const TOOL_EXECUTORS = {
    get_presentation_outline: getPresentationOutline,
    get_slide: getSlide,
    get_slide_preview: getSlidePreview,
    get_selected_slides: getSelectedSlides,
    get_slide_notes: getSlideNotes,
    list_layouts: listLayouts,
    list_style_templates: listStyleTemplates,
    save_style_template: saveStyleTemplate,
    set_slide_background: setSlideBackground,
    edit_table_structure: editTableStructure,
    crop_image: cropImage,
    set_picture_opacity: setPictureOpacity,
    ask_clarification: askClarificationTool,
    read_memory: (args) => App.memory.readMemoryTool(args),
    write_memory: (args) => App.memory.writeMemoryTool(args),
    recolor_slide: recolorSlide,
    set_table_style: setTableStyle,
    add_table_grid: addTableGrid,
    align_shapes: alignShapes,
    replace_image: replaceImage,
    get_chart: getChart,
    add_slide: addSlide,
    delete_slide: deleteSlide,
    duplicate_slide: duplicateSlide,
    set_slide_notes: setSlideNotes,
    insert_textbox: insertTextbox,
    set_text: setText,
    set_hyperlink: setHyperlinkTool,
    apply_layout: applyLayoutDirect,
    insert_image: insertImage,
    goto_slide: gotoSlide,
    eval_officejs: evalOfficeJs
  };

  async function getPresentationMetadata() {
    requireOffice();
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const items = slides.items.slice(0, 60);
      items.forEach(s => { s.load('id'); s.shapes.load('items'); });
      await context.sync();
      items.forEach(s => s.shapes.items.forEach(sh => { try { sh.textFrame.textRange.load('text'); } catch {} }));
      // 批量 sync 可能因组合等无 textFrame 的形状整体失败：失败后退回逐形状独立深读，不再拖垮整个 metadata
      let batchOk = true;
      try { await context.sync(); } catch { batchOk = false; }
      const out = [];
      for (let idx = 0; idx < items.length; idx++) {
        const s = items[idx];
        let texts = s.shapes.items.map(shapeText).map(x => x.trim()).filter(Boolean);
        if (!batchOk) {
          texts = [];
          for (const sh of s.shapes.items) {
            const t = (await shapeTextDeep(context, sh)).trim();
            if (t) texts.push(t);
          }
        }
        out.push({ index: idx, id: s.id, title: clampText(texts[0] || '', 100), refId: `s:${idx}` });
      }
      return { success: true, presentationId: App.state.workbookId || 'presentation', slideCount: slides.items.length, capabilities: presentationCapabilities(), slides: out };
    });
  }

  // 全文注入：低于阈值时把整份演示文稿的文字一次性交给模型，省掉多次读取往返。
  const FULL_CONTEXT_CHAR_LIMIT = 15000;
  const FULL_CONTEXT_SLIDE_BATCH = 20;

  async function getFullContextBatched() {
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const items = slides.items;
      items.forEach(s => { s.load('id'); s.shapes.load('items'); });
      await context.sync();

      const blocks = [];
      let charCount = 0;
      for (let batchStart = 0; batchStart < items.length; batchStart += FULL_CONTEXT_SLIDE_BATCH) {
        const batch = items.slice(batchStart, batchStart + FULL_CONTEXT_SLIDE_BATCH);
        const shapes = [];
        for (const slide of batch) {
          for (const shape of (slide.shapes.items || [])) {
            shape.load('id,type');
            shape.textFrame.load('hasText');
            shapes.push(shape);
          }
        }
        await context.sync();
        const textShapes = shapes.filter(shape => shape.textFrame.hasText);
        textShapes.forEach(shape => shape.textFrame.textRange.load('text'));
        if (textShapes.length) await context.sync();

        for (let offset = 0; offset < batch.length; offset++) {
          const idx = batchStart + offset;
          const slide = batch[offset];
          const lines = [];
          for (const shape of (slide.shapes.items || [])) {
            let hasText = false;
            try { hasText = Boolean(shape.textFrame.hasText); } catch {}
            let text = hasText ? shapeText(shape).trim() : '';
            // 组合/表格：顶层 textFrame 读不到文字，递归读取子形状/单元格（批量 hasText 因组合失败时会整体退回 sequential 路径，同样覆盖）
            if (!text && (shape.type === 'Group' || shape.type === 'Table')) text = (await shapeTextDeep(context, shape)).trim();
            if (!text) continue;
            charCount += text.length;
            lines.push(`  [shape ${shape.id}] ${text}`);
          }
          blocks.push(`[slide ${idx}] (id=${slide.id})\n${lines.join('\n') || '  (无文本)'}`);
          if (charCount > FULL_CONTEXT_CHAR_LIMIT) {
            return { truncated: true, charCount, slideCount: items.length };
          }
        }
      }
      return { truncated: false, charCount, slideCount: items.length, text: blocks.join('\n\n') };
    });
  }

  // 兼容旧版 PowerPoint：若批量读取 textFrame.hasText 不可用，退回逐 shape 独立同步。
  async function getFullContextSequential() {
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const items = slides.items;
      items.forEach(slide => { slide.load('id'); slide.shapes.load('items'); });
      await context.sync();

      const blocks = [];
      let charCount = 0;
      for (let idx = 0; idx < items.length; idx++) {
        const slide = items[idx];
        slide.shapes.items.forEach(shape => { try { shape.load('id'); } catch {} });
        await context.sync();
        const lines = [];
        for (const shape of (slide.shapes.items || [])) {
          const text = (await shapeTextDeep(context, shape)).trim(); // 深读：组合子形状与表格单元格也要覆盖
          if (!text) continue;
          charCount += text.length;
          lines.push(`  [shape ${shape.id}] ${text}`);
        }
        blocks.push(`[slide ${idx}] (id=${slide.id})\n${lines.join('\n') || '  (无文本)'}`);
        if (charCount > FULL_CONTEXT_CHAR_LIMIT) return { truncated: true, charCount, slideCount: items.length };
      }
      return { truncated: false, charCount, slideCount: items.length, text: blocks.join('\n\n') };
    });
  }

  async function getFullContext() {
    requireOffice();
    let result;
    try { result = await getFullContextBatched(); }
    catch { result = await getFullContextSequential(); }
    result.capabilities = presentationCapabilities();
    return result;
  }

  const SYSTEM_PROMPT = `你是嵌在 Microsoft PowerPoint 里的中文写作助手。用户是中文母语的专业文字工作者，你的全部回复和所有写入演示文稿的内容都用中文。

工作方式：
1. 先判断用户到底要什么。要求含糊时先问，不要猜着改。
2. 动手前先说明你的整体判断：这份演示文稿或这页的问题在哪、你打算怎么处理。这段话用自然语言写在正文里，不要塞进工具参数。
3. 然后才提出具体修改。
4. 每一处改动都要能追溯到用户的要求，不要顺手改用户没提的地方。

呈现规则（决定用户看到的是卡片还是文字，很重要）：
- 当你的回答是「N 个同构的项」「几个方案并排比较」「行列矩阵」「层级大纲」「需要用户填的表单」时，调用 render 工具，界面会负责编号、对齐、跳转和操作按钮。
- 用了 render 之后，正文里不要再把同样的内容复述一遍，写一句话引出即可。
- 反过来，下面这些情况一律用正文回复，不要调 render：普通问答、解释概念、只有一两句话的回答、单个连续段落的叙述或改写、闲聊与确认。给一段话套上卡片只会让界面变吵，不会让它变清楚。
- 判断标准：需要对着比较、需要逐项操作、或者项与项之间字段相同，就用 render；只是要读一遍，就用正文。


You are an AI assistant integrated into Microsoft PowerPoint with access to read and modify the presentation.

Available tools:
READ:
- get_presentation_outline: List slides with their titles and text
- get_slide: Read all shapes and text on a specific slide (by index)
- get_slide_preview: Render a slide as a PNG so you can judge the actual visual result
- get_selected_slides: Get the currently selected slides
- get_slide_notes: Read speaker notes (one slide by index, or all slides without index)
- list_layouts: List slide masters and their layouts (id + name) so a new slide can be built on an existing design
- get_chart: Read a chart shape's data (chart type, series names, values where the host allows)
- list_style_templates: List saved design templates (named palette + fonts + font-size hierarchy in the add-in shared settings, available across documents)

WRITE:
- add_slide: Add a new slide, optionally on a specific layout (layoutId from list_layouts) so it inherits that design
- delete_slide: Delete a slide by index
- duplicate_slide: Duplicate a slide by index
- set_slide_notes: Set the speaker notes of a slide
- insert_textbox: Add a text box to a slide (position/size in points)
- set_text: Replace the text of a shape on a slide
- recolor_slide: Remap colors and fonts across a whole slide in one call. Run it with dryRun first to get the slide actual palette and font list, then run it again with a complete colorMap. Use this for restyling instead of hand-written eval loops.
- set_table_style: Restyle cells of a table shape (font name/size/color/bold, cell background, or scaleFontSize to shrink every cell size proportionally). Whole table or limited rows/cols/cells.
- add_table_grid: Create a NEW table on a slide as a grid of shapes (rectangles + text boxes). PowerPoint Office.js has no native insert-table API, so the result is NOT a native Table — it only looks like one. For an existing native table, use set_table_style.
- align_shapes: Align or evenly distribute multiple shapes on a slide (left/center/right/top/middle/bottom, distributeH/distributeV).
- replace_image: Replace a picture shape's content with a new base64 image, keeping its position and size.
- save_style_template: Save the current deck's design tokens (colors, fonts, size hierarchy) as a named template in the add-in shared settings, available across documents
- set_slide_background: Set a slide's background to a solid color (falls back to a full-page rectangle sent to back when the host has no background API)
- edit_table_structure: Add/remove a row or column of a table shape (most hosts lack this API; the error names the fallback)
- crop_image: Crop a picture shape by edge fractions (most hosts lack this API; the error names alternatives)
- set_picture_opacity: Set a picture's opacity 0-1 (most hosts lack this API; the error names alternatives)
- apply_layout: Add, update, or delete a shape with explicit geometry and formatting. Follow the effective access mode and this user's review request; direct formatting permission is not permission to bypass a requested proposal.
- insert_image: 在当前选区插入图片。用户上传/拖拽过图片时必须用 attachmentId（消息里的 [Uploaded image: …] 块），图片直链用 url，不要自己编 base64
- goto_slide: Navigate to a slide by index
- eval_officejs: Execute PowerPoint.run code when the listed tools are not enough
- ask_clarification: Show the user a question card (select/text fields) and return their answers keyed by question id. Use it BEFORE acting when the direction is unclear — ask, do not guess.

Layout rules:
- Read the target slide with get_slide before proposing or applying layout work. Reuse its actual slide id and shape id; the app supplies the live shape snapshot, never invent operation.expected.
- For whole-deck design, understand the requested pages and their narrative before editing. Choose the layout from the content and any references provided by the user. No reference library is preinstalled. Diagnose each page's reading problem, then choose hierarchy and composition. Preserve client branding, facts and the necessary relationships conveyed by existing arrows. Re-express these relationships through a new composition; deleting arrows alone does not fulfill a relayout request. A prototype is only part of the requested scope.
- 占位框清理依据（三态契约）：get_slide 的 isEmpty 只表示「没有文字」，绝不能单独作为删除依据——图片/图表占位框同样没有文字。只有 placeholderContent 为 confirmed-empty（containedType 可靠读回 null）的空占位框、且被本次新布局替代，才可在提案中复用其几何或用 apply_layout(deleteShape) 删除；has-content（含图片/图表/表格等非文字内容）与 unknown（旧宿主或读取失败）一律保护，不得按空占位框清理。新内容叠在确认空的占位框上会被 layoutAudit 报为 high 级问题。有真实文字的占位框、Logo、页码、母版/版式自带对象一律保留，不属于清理范围。
- For visual layout, also call get_slide_preview when capabilities.slidePreview=true. Coordinates alone are not enough to judge balance, whitespace, or visual style.
- Treat "微调/优化/顺一下/好看一点" as an in-style polish: preserve the existing palette, fonts, and design language; improve hierarchy, alignment, spacing, proportions, and composition. Do not introduce a new visual style unless the user explicitly asks for a redesign.
- 换肤/风格迁移前必须先调研现有 deck：用 get_slide 读 2-3 页代表页，提取实际在用的字体（font.name 是英文内部名）、配色、字号层级和版心安全区（底部 banner/页脚条等固定占位内的区域不能再放内容）。中文字体名与内部英文名映射：微软雅黑=Microsoft YaHei，微软正黑（体）=Microsoft JhengHei，苹方=PingFang SC，宋体=SimSun，黑体=SimHei，仿宋=FangSong，楷体=KaiTi。用户报中文字体名时，先查现有页面字体的英文对应名再设置，不要按字面硬设；同一页面的中西文字体风格必须协调（衬线配衬线，黑体配无衬线西文）。
- 换肤/风格迁移改配色和字体时用 recolor_slide，不要手写 eval 遍历：先 dryRun 拿到本页完整调色板（同一个色系常有多个色值，例如 #F00018 和 #FF0000 都是红），补全 colorMap 后再正式跑一次。
- 换肤/风格迁移前先调 list_style_templates：已存了这个客户的模板（如「品牌模板」）就直接用它的 colors/fonts 组装 recolor_slide 的 colorMap/fontMap，不要每次重新调研；没有就在换肤完成、复检通过后用 save_style_template 存一份，下次直接套用。
- 用户提供的版式参考只借结构（分组/顺序/层级/阅读顺序）：其颜色、字体、圆角、装饰与绝对坐标不是客户页规范，不得照搬。视觉属性按「用户本次明确要求 > 当前客户模板 > 同模板邻近页证据 > 简单一致默认」取值：用 get_slide 的字体/fillColor/lineColor/lineWidth 读模板证据。shape.type 只有大类（几何形状一律 GeometricShape），圆角/直角细类型读不到：不得从 GeometricShape 推断直角或圆角，不得声称已识别客户圆角；圆角/直角凭视觉证据（用户指定示范页或明确说法，capabilities.slidePreview=true 时用 get_slide_preview 目视核对），无证据记未知、用简单一致默认（推荐直角）并在提案中声明，不把蓝图样式冒充客户风格。同一卡片的色条与外框不得轮廓冲突、凸角、留缝或双边：直角模板用普通 rectangle；圆角模板外框与色条同 left/top/width 且同一个显式 cornerRadius（绝对磅值 ≤min(width,height)/2），色条加 cornerStyle:"top"（宿主映射原生 Round2SameRectangle 并钉绝对半径）；capabilities.shapeAdjustments=false 或写入报错时改直角设计，不要默认圆角硬拼。
- 仅在缺少的信息会改变核心内容、客户风格或改动范围时用 ask_clarification；普通排版选择由已读内容和用户要求判断，不把字号、坐标、参考编号变成用户必须回答的问题。
- 改表格单元格样式（字体/字号/字色/加粗/底色，或 scaleFontSize 统一缩放字号）用 set_table_style；多个形状对齐或等距分布用 align_shapes；替换图片内容用 replace_image。不要为这些手写 eval。
- 表格两条路别混：Office.js 没有原生插入表格 API，新建表格用 add_table_grid（矩形+文本框拼的视觉表格，不是原生 Table，之后改样式走 apply_layout/recolor_slide）；页面上已有的原生表格用 set_table_style 调样式、edit_table_structure 探测结构操作。不要在 eval 里找 addTable，它不存在。
- 换字体/换肤后必须重新跑版面审计（get_slide 的 layoutAudit，必要时加 get_slide_preview 复检）：字体不同度量不同，同样的字号在新字体下可能溢出、异常换行或重叠。这类由字体度量变化引起的问题，通过微调字号（fontSize）解决，不要靠移动元素位置硬躲；用户说「不用调整版式」是指不移动元素位置（left/top/width/height），不等于字号不可调。本条与下面的 high 级问题 blocking 规则衔接：换肤后审计仍报 high 级问题时，同样不许宣布完成。
- 换肤/迁移风格的首选策略是「借版」而不是原地重建。借版有两条路，按宿主能力选：(1) duplicate_slide 复制目标风格里版式最接近的一页；(2) 宿主不支持复制时（capabilities.duplicateSlide=false，或 duplicate_slide 报错说没有该 API），用 list_layouts 列出母版下的版式，挑最接近的一个用 add_slide({ layoutId }) 新建页——新页自带该版式的背景、banner、logo 和占位符——把内容填进去，再 delete_slide 删原页。新页在末尾，页序要提醒用户自己拖。已有页面的版式不可切换，不要尝试给 slide.layout 赋值。只有确实没有可借的版式时才原地重建；原地重建若连续两次达不到效果，立即改用借版，不要硬撑到改不下去才停手。
- 「借版」身份铁律（违反会毁掉无关页面）：新建页必须用 add_slide 工具并以其返回的 index/id 定位，禁止用 eval 裸调 slides.add()（它不返回新页身份），禁止凭猜测用索引定位任何"刚建的页"。红旗：如果你认为刚创建的页面里出现非占位符的旧内容（成段的正文文字），说明定位错了页——立即停手重新核对，绝不要清空它。
- 删除原页前必须先验证：用 get_slide 读取替代页，确认应迁移的文字确实都在，才能 delete_slide 原页；验证不过就先修内容，绝不能在替代页为空或不完整时删原页。
- 预览图可能滞后于刚写入的修改。相同字节数不能证明缓存；图像与实际读回内容不符时说明预览可能未更新，数据快照只能核验内容与几何，不能替代视觉验收。
- A layout request must result in real geometry or formatting work when needed; do not reduce it to changing a few words.
- After direct layout edits, inspect the refreshed get_slide audit and preview. Make at most one focused correction pass; do not keep restyling indefinitely.
- High-severity text_overflow_risk, unexpected_title_wrap, text_overlap, text_rule_collision, or out_of_bounds issues are blocking. Do not report completion while one remains. Low-severity findings (e.g. bleeding images, decorative large-text overlaps) are informational.
- Reconcile completion claims with the requested pages and actual tool results. Failed, stale, or skipped items are not completed; a failed write may have partial effects, so use writeState and reread before describing the document. A read page is not automatically redesigned, and a geometry audit is not visual approval. Pages may remain unchanged for a concrete design reason. List remaining or unverifiable work; never claim a page or the whole deck is rebuilt/done while requested work is incomplete, even if the pages you touched passed an audit.
- In confirm mode, additive and formatting layout work (adding shapes, colors, fonts, positions) applies directly via apply_layout — no proposal card needed. DESTRUCTIVE operations must go through propose_edits first: deleting shapes and deleting slides are both destructive — apply_layout refuses deleteShape in confirm mode, and repeating the same refused call will not pass; switch to a proposal card on the first refusal. Text content changes also require propose_edits. Never describe a visual change as a text replacement.
- Never fill operation.expected yourself for updateShape/deleteShape; the add-in reads the live shape snapshot for you on both the direct and the proposal path, so omitting it is correct and will not be rejected — the app reads the live shape snapshot automatically at review/apply time; you only provide target (slideId/shapeId) and the desired changes.
- A background card that must sit behind text requires capabilities.zOrder=true. The correct API is shape.setZOrder; there is no shape.zOrder property.
- Keep each layout card atomic. One card adds, updates, or deletes one shape.

Note: Basic shape layout requires PowerPointApi 1.4. Z-order requires 1.8. The structured layout executor checks both before writing and will make no change when a required capability is unavailable. Use eval_officejs only for operations that cannot be expressed by apply_layout. When eval_officejs changes slides, always include every zero-based slide index in modifiedSlideIndexes so the add-in can run its mandatory post-write review.

Citations: Use markdown links with #cite: hash to reference slides. Clicking navigates there.
- Slide by index: [agenda](#cite:s:1)
Example: [see slide 2](#cite:s:1)
MUST: whenever your answer refers to specific slides — which page contains what, differences between pages, where a problem is — write those references as clickable citations [short label](#cite:s:N), never as bare numbers like "slide 2" or 「第2页」. This is a hard formatting requirement; answers about slide locations without clickable citations are considered non-compliant.


修改演示文稿前先确认你读到的是完整原文，不是摘要或截断文本。`;

  const POWERPOINT_LAYOUT_SHAPE = {
    type: 'object',
    properties: {
      shapeType: { type: 'string', enum: ['rectangle', 'roundRectangle', 'ellipse', 'textBox', 'line'] },
      name: { type: 'string' }, text: { type: 'string' },
      left: { type: 'number' }, top: { type: 'number' }, width: { type: 'number' }, height: { type: 'number' },
      fillColor: { type: 'string', description: '#RRGGBB only.' }, fillTransparency: { type: 'number', description: '0 to 1.' },
      lineVisible: { type: 'boolean' }, lineColor: { type: 'string', description: '#RRGGBB only.' }, lineTransparency: { type: 'number', description: '0 to 1.' }, lineWidth: { type: 'number' },
      fontName: { type: 'string' }, fontSize: { type: 'number' }, fontColor: { type: 'string', description: '#RRGGBB only.' }, bold: { type: 'boolean' },
      horizontalAlignment: { type: 'string', enum: ['Left', 'Center', 'Right', 'Justify', 'Distributed'] },
      verticalAlignment: { type: 'string', enum: ['Top', 'Middle', 'Bottom', 'TopCentered', 'MiddleCentered', 'BottomCentered'] },
      marginLeft: { type: 'number' }, marginRight: { type: 'number' }, marginTop: { type: 'number' }, marginBottom: { type: 'number' }, wordWrap: { type: 'boolean' },
      autoSizeSetting: { type: 'string', enum: ['AutoSizeNone', 'AutoSizeTextToFitShape', 'AutoSizeShapeToFitText'] },
      zOrder: { type: 'string', enum: ['sendToBack', 'sendBackward', 'bringForward', 'bringToFront'] }
    }
  };

  const POWERPOINT_LAYOUT_EXPECTED = {
    type: 'object',
    description: 'Optional; do NOT fill this yourself. The add-in captures the live shape snapshot automatically at review/apply time (47: also for Groups and unreadable shapes it records what is verifiable). You only supply target and desired changes.',
    properties: {
      name: { type: 'string' }, type: { type: 'string' }, text: { type: 'string' },
      left: { type: 'number' }, top: { type: 'number' }, width: { type: 'number' }, height: { type: 'number' }, zOrderPosition: { type: 'number' },
      fillType: { type: 'string' }, fillColor: { type: 'string' }, fillTransparency: { type: 'number' },
      lineVisible: { type: 'boolean' }, lineColor: { type: 'string' }, lineTransparency: { type: 'number' }, lineWidth: { type: 'number' },
      fontName: { type: 'string' }, fontSize: { type: 'number' }, fontColor: { type: 'string' }, bold: { type: 'boolean' },
      horizontalAlignment: { type: 'string' }, verticalAlignment: { type: 'string' },
      marginLeft: { type: 'number' }, marginRight: { type: 'number' }, marginTop: { type: 'number' }, marginBottom: { type: 'number' }, wordWrap: { type: 'boolean' }, autoSizeSetting: { type: 'string' }
    }
  };

  const POWERPOINT_LAYOUT_OPERATION = {
    type: 'object',
    properties: {
      type: { type: 'string', enum: ['addShape', 'updateShape', 'deleteShape'] },
      shape: POWERPOINT_LAYOUT_SHAPE,
      expected: POWERPOINT_LAYOUT_EXPECTED
    },
    required: ['type']
  };

  const POWERPOINT_EDIT_ITEM = {
    type: 'object',
    properties: {
      kind: { type: 'string', enum: ['text', 'layout'], description: 'Use layout for shape geometry/format/layer changes; otherwise text.' },
      label: { type: 'string', description: 'Short name for this suggestion.' },
      placement: { type: 'string', enum: ['replace', 'after', 'before'], description: 'replace swaps find; after/before inserts replacement beside the anchor.' },
      find: { type: 'string', description: 'Original shape or notes text, or a focused exact excerpt.' },
      replacement: { type: 'string', description: 'Exact new text.' },
      contextBefore: { type: 'string', description: 'Optional short unchanged context before find, for display only.' },
      contextAfter: { type: 'string', description: 'Optional short unchanged context after find, for display only.' },
      target: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' }, slideId: { type: 'string', description: 'Stable slide id returned by get_slide when available.' }, shapeId: { type: 'string' }, kind: { type: 'string', description: '"notes" for speaker notes; omit for shape text.' } }, required: ['index'] },
      operation: POWERPOINT_LAYOUT_OPERATION,
      reasoning: { type: 'string' },
      summary: { type: 'string' }
    },
    required: ['label', 'target', 'reasoning', 'summary']
  };

  const TOOL_DEFINITIONS = [
    { type: 'function', function: { name: 'get_presentation_outline', description: 'List slides with titles and text content.', parameters: { type: 'object', properties: { maxSlides: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'get_slide', description: 'Read all shapes and text on a slide by index.', parameters: { type: 'object', properties: { index: { type: 'number' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'get_slide_preview', description: 'Render one slide as a PNG for visual review. Use for layout or beautification work when slidePreview capability is available.', parameters: { type: 'object', properties: { index: { type: 'number' }, height: { type: 'number', description: 'Preview height in pixels, 240 to 900. Default 540.' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'get_selected_slides', description: 'Get the currently selected slide(s).', parameters: { type: 'object', properties: { explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'list_layouts', description: 'READ. List every slide master and its layouts (id + name). Call this before add_slide whenever the new slide should carry an existing design: background, banner, logo, and placeholder geometry all come from the layout.', parameters: { type: 'object', properties: { explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'add_slide', description: 'WRITE. Add a new slide at the end. Pass layoutId (from list_layouts) to build it on a specific layout, which brings that layout background, banner, logo and placeholders — on hosts without slide duplication this is how you borrow a design. Omit layoutId for the default layout. The layout of an EXISTING slide cannot be changed via API, and slides cannot be reordered — if page order matters, tell the user to drag it in PowerPoint.', parameters: { type: 'object', properties: { layoutId: { type: 'string', description: 'Layout id from list_layouts. Omit for the default layout.' }, slideMasterId: { type: 'string', description: 'Master id from list_layouts. Only needed when the same layout id appears under more than one master.' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'delete_slide', description: 'WRITE. Delete a slide by index. Irreversible in this add-in: before deleting, verify with get_slide that any content meant to be moved already exists on the replacement slide — never delete the source when the replacement is empty or incomplete.', parameters: { type: 'object', properties: { index: { type: 'number' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'duplicate_slide', description: 'WRITE. Duplicate a slide by index.', parameters: { type: 'object', properties: { index: { type: 'number' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'get_slide_notes', description: 'READ. Get the speaker notes of one slide (pass index) or of every slide (omit index).', parameters: { type: 'object', properties: { index: { type: 'number', description: 'Slide index. Omit to read notes of all slides.' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'set_slide_notes', description: 'WRITE. Set the speaker notes text of a slide.', parameters: { type: 'object', properties: { index: { type: 'number' }, text: { type: 'string' }, explanation: { type: 'string' } }, required: ['index', 'text'] } } },
    { type: 'function', function: { name: 'insert_textbox', description: 'WRITE. Add a text box to a slide. Position/size in points.', parameters: { type: 'object', properties: { index: { type: 'number' }, text: { type: 'string' }, left: { type: 'number' }, top: { type: 'number' }, width: { type: 'number' }, height: { type: 'number' }, explanation: { type: 'string' } }, required: ['index', 'text'] } } },
    { type: 'function', function: { name: 'set_text', description: 'WRITE. Replace the text of a shape on a slide (first shape if shapeId omitted).', parameters: { type: 'object', properties: { index: { type: 'number' }, shapeId: { type: 'string' }, text: { type: 'string' }, explanation: { type: 'string' } }, required: ['index', 'text'] } } },
    { type: 'function', function: { name: 'set_hyperlink', description: 'WRITE. Set a hyperlink without changing any text. Three targets: a substring inside one shape (target + find), a whole shape (target, no find), or the currently selected text in PowerPoint (omit target — the user selects text first). Requires hyperlinkApi capability (PowerPointApi 1.10); the result is verified by reading the link back. Writing the URL as plain text is the fallback when this API is unavailable.', parameters: { type: 'object', properties: { address: { type: 'string', description: 'Link target: https://…, mailto:…, or a file name.' }, target: { type: 'object', properties: { slideId: { type: 'string' }, shapeId: { type: 'string' } }, description: 'Slide + shape that owns the text. Omit to use the current text selection.' }, find: { type: 'string', description: 'Exact substring inside the shape text to hyperlink. Omit to link the whole shape. Must match get_slide text exactly.' }, screenTip: { type: 'string', description: 'Optional tooltip shown on hover.' }, explanation: { type: 'string' } }, required: ['address'] } } },
    { type: 'function', function: { name: 'recolor_slide', description: 'WRITE. Remap colors and fonts across one whole slide: every shape fill, line color, font color and font name that matches is replaced, including fonts inside table cells. Call it with dryRun:true first — that writes nothing and returns the slide actual palette, font list and font sizes, so the colorMap can cover every value of a color family instead of just one. Geometry and text are never touched.', parameters: { type: 'object', properties: { index: { type: 'number' }, colorMap: { type: 'object', description: 'Exact hex to hex map, e.g. { "#F00018": "#8D866D", "#FF0000": "#8D866D" }. Applies to fills, line colors and font colors alike.' }, fontMap: { type: 'object', description: 'Font name to font name map, e.g. { "AliHYAiHei": "Microsoft JhengHei UI" }.' }, fontName: { type: 'string', description: 'Set every text shape on the slide to this font. fontMap wins where both match.' }, dryRun: { type: 'boolean', description: 'Preview only: write nothing, return the palette, the fonts and what would change.' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'apply_layout', description: 'WRITE. Apply one structured PowerPoint shape operation under the current access mode. In confirm mode, deletions and text content changes require propose_edits; eligible formatting may apply directly. When the user asks to review a proposal first, use propose_edits for the requested changes. Plan mode does not allow writes.', parameters: { type: 'object', properties: { kind: { type: 'string', enum: ['layout'] }, label: { type: 'string' }, target: POWERPOINT_EDIT_ITEM.properties.target, operation: POWERPOINT_LAYOUT_OPERATION, explanation: { type: 'string' } }, required: ['target', 'operation'] } } },
    { type: 'function', function: { name: 'insert_image', description: 'WRITE. Insert an image into the current selection. Preferred source: attachmentId — the id shown in the [Uploaded image: … | attachmentId: img-N] block when the user uploaded or dragged an image into the chat. Also accepts url (direct link, downloaded via the local service) or raw base64 — never generate base64 yourself.', parameters: { type: 'object', properties: { attachmentId: { type: 'string' }, url: { type: 'string' }, base64: { type: 'string' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'set_table_style', description: 'WRITE. Restyle the cells of ONE table on a slide: fontName/fontSize/fontColor/bold and cellFill (cell background), applied to the whole table or limited by rows/cols/cells. scaleFontSize (e.g. 0.85) multiplies every cell current font size in one call — the right tool for "统一缩小表格字号", since it preserves the existing size hierarchy instead of flattening it to one value. Returns how many cells changed and how many were unreadable. Do NOT hand-write eval loops for table styling.', parameters: { type: 'object', properties: { target: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' }, shapeId: { type: 'string', description: 'Id of the Table shape, from get_slide.' } }, required: ['index', 'shapeId'] }, fontName: { type: 'string' }, fontSize: { type: 'number' }, fontColor: { type: 'string', description: '#RRGGBB only.' }, bold: { type: 'boolean' }, cellFill: { type: 'string', description: 'Cell background color, #RRGGBB only. May be unsupported on some hosts — check the result cellFill field.' }, scaleFontSize: { type: 'number', description: 'Multiply each cell current font size by this factor (0-2), e.g. 0.85. Alternative to fontSize.' }, rows: { type: 'array', items: { type: 'number' }, description: 'Limit to these 0-based rows.' }, cols: { type: 'array', items: { type: 'number' }, description: 'Limit to these 0-based columns.' }, cells: { type: 'array', items: { type: 'array', items: { type: 'number' } }, description: 'Limit to these [row, col] cells.' }, explanation: { type: 'string' } }, required: ['target'] } } },
    { type: 'function', function: { name: 'add_table_grid', description: 'WRITE. Create a NEW table on a slide. PowerPoint Office.js has NO native insert-table API, so this builds a visual table from one rectangle (border + fill) plus one text box per cell — the result is ordinary shapes, NOT a native Table. First data row is the header (bold, headerFill background). Max 12x12. To restyle an EXISTING native table use set_table_style instead.', parameters: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' }, data: { type: 'array', description: '2D array of cell text; first row is the header.', items: { type: 'array', items: { type: 'string' } } }, rows: { type: 'number', description: 'Row count, 1-12. Default: data row count.' }, columns: { type: 'number', description: 'Column count, 1-12. Default: widest data row.' }, left: { type: 'number' }, top: { type: 'number' }, width: { type: 'number' }, height: { type: 'number' }, headerFill: { type: 'string', description: 'Header row background, #RRGGBB only. Optional.' }, cellFill: { type: 'string', description: 'Non-header cell background, #RRGGBB only. Optional; omitted = no fill (only borders).' }, borderColor: { type: 'string', description: 'Cell border color, #RRGGBB only. Default #BFBFBF.' }, fontName: { type: 'string' }, fontSize: { type: 'number' }, explanation: { type: 'string' } }, required: ['index', 'data', 'left', 'top', 'width', 'height'] } } },
    { type: 'function', function: { name: 'align_shapes', description: 'WRITE. Align or evenly distribute several shapes on one slide. Pure geometry move, styles untouched. action: left/center/right/top/middle/bottom (align, aliases alignLeft etc. accepted) or distributeH/distributeV (even spacing, needs 3+ shapes). reference: "first" (default, anchor on the first shapeId in the list) or "slide" (page bounds). Do NOT hand-write eval loops for alignment.', parameters: { type: 'object', properties: { target: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' } }, required: ['index'] }, shapeIds: { type: 'array', items: { type: 'string' }, description: 'Shape ids from get_slide, in any order; with reference "first" the first id is the anchor.' }, action: { type: 'string', enum: ['left', 'center', 'right', 'top', 'middle', 'bottom', 'alignLeft', 'alignCenter', 'alignRight', 'alignTop', 'alignMiddle', 'alignBottom', 'distributeH', 'distributeV'] }, reference: { type: 'string', enum: ['first', 'slide'], description: 'Anchor for alignment/distribution. Default "first".' }, explanation: { type: 'string' } }, required: ['target', 'shapeIds', 'action'] } } },
    { type: 'function', function: { name: 'replace_image', description: 'WRITE. Replace the content of ONE picture shape with a new base64 image, keeping its position and size. On hosts without an in-place image replace API this deletes the old picture and inserts the new one at the same geometry — the new picture then sits on top; fix stacking with apply_layout zOrder if needed. The result path field says which way was used.', parameters: { type: 'object', properties: { target: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' }, shapeId: { type: 'string', description: 'Id of the Picture shape, from get_slide.' } }, required: ['index', 'shapeId'] }, base64: { type: 'string', description: 'New image, base64 without data: prefix.' }, explanation: { type: 'string' } }, required: ['target', 'base64'] } } },
    { type: 'function', function: { name: 'get_chart', description: 'READ. Read the data of ONE chart shape on a slide: chart type, series names, and values/categories where the host Chart API exposes them. Parts the host cannot read are marked valuesUnreadable/unreadableStyle — "unreadable" never means "no data". For numbers on thin-API hosts, fall back to get_slide_preview or ask the user to open the chart data in Excel.', parameters: { type: 'object', properties: { target: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' }, shapeId: { type: 'string', description: 'Id of the Chart shape, from get_slide.' } }, required: ['index', 'shapeId'] }, explanation: { type: 'string' } }, required: ['target'] } } },
    { type: 'function', function: { name: 'list_style_templates', description: 'READ. List saved design templates (name, color palette, fonts, font-size hierarchy) kept in the add-in shared settings — captured from this or earlier decks and available across documents. Check this BEFORE restyling: if a template for this client already exists, reuse its colors/fonts instead of re-surveying the deck.', parameters: { type: 'object', properties: { explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'save_style_template', description: 'WRITE. Capture the current deck design tokens — color palette, fonts, font-size hierarchy — from representative slides (default: first 3) and save them as a named template in the add-in shared settings, available across documents. Save after a restyle so the same client look can be reused next time. Slide background color cannot be read via Office.js and is NOT stored.', parameters: { type: 'object', properties: { name: { type: 'string', description: 'Template name, e.g. "品牌模板". Same name overwrites.' }, indexes: { type: 'array', items: { type: 'number' }, description: 'Representative 0-based slide indexes to sample (max 5). Default: first 3 slides.' }, note: { type: 'string', description: 'Optional short note, e.g. client or usage.' }, explanation: { type: 'string' } }, required: ['name'] } } },
    { type: 'function', function: { name: 'set_slide_background', description: 'WRITE. Set one slide background to a solid color. Hosts without a background API get a full-page rectangle sent to back instead (named __AI_PageBackground; calling again on the same slide recolors it instead of stacking a new one) — the result path field says which way was used ("backgroundApi" or "fullPageRect"). The rectangle covers background elements coming from the master/layout (banner/logo render below slide shapes) — if this slide banner comes from the master, do NOT use this tool; borrow the design via list_layouts + add_slide instead.', parameters: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' }, color: { type: 'string', description: '#RRGGBB only.' }, explanation: { type: 'string' } }, required: ['index', 'color'] } } },
    { type: 'function', function: { name: 'edit_table_structure', description: 'WRITE. Add or remove one row/column of ONE table shape. action: addRow/removeRow/addColumn/removeColumn; position is 0-based (add = insert at, remove = delete; default: append/last). Most hosts have NO table-structure API — then it fails with an explicit fallback instead of pretending. Cell styling is set_table_style; whole-table geometry is apply_layout.', parameters: { type: 'object', properties: { target: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' }, shapeId: { type: 'string', description: 'Id of the Table shape, from get_slide.' } }, required: ['index', 'shapeId'] }, action: { type: 'string', enum: ['addRow', 'removeRow', 'addColumn', 'removeColumn'] }, position: { type: 'number', description: '0-based row/column position. Default: append for add, last for remove.' }, explanation: { type: 'string' } }, required: ['target', 'action'] } } },
    { type: 'function', function: { name: 'crop_image', description: 'WRITE. Crop ONE picture shape by fractions of each edge (0-1, e.g. cropTop 0.1 cuts 10% off the top). At least one of cropTop/cropBottom/cropLeft/cropRight is required. Most hosts have NO crop API — then it fails with explicit alternatives (replace_image with a pre-cropped image, or manual crop), it never pretends.', parameters: { type: 'object', properties: { target: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' }, shapeId: { type: 'string', description: 'Id of the Picture shape, from get_slide.' } }, required: ['index', 'shapeId'] }, cropTop: { type: 'number' }, cropBottom: { type: 'number' }, cropLeft: { type: 'number' }, cropRight: { type: 'number' }, explanation: { type: 'string' } }, required: ['target'] } } },
    { type: 'function', function: { name: 'set_picture_opacity', description: 'WRITE. Set the opacity of ONE picture shape (0-1, 1 = fully opaque). Most hosts cannot write picture transparency — then it fails with the fallback (overlay a semi-transparent rectangle via apply_layout, or set it manually), it never pretends.', parameters: { type: 'object', properties: { target: { type: 'object', properties: { index: { type: 'number', description: 'Slide index.' }, shapeId: { type: 'string', description: 'Id of the Picture shape, from get_slide.' } }, required: ['index', 'shapeId'] }, opacity: { type: 'number', description: '0 to 1. 1 = fully opaque.' }, explanation: { type: 'string' } }, required: ['target', 'opacity'] } } },
    { type: 'function', function: { name: 'read_memory', description: 'SYSTEM. Read your cross-session memory (habits and preferences learned from past work). Call with no topic to list what is stored; pass topic (key) to read its full content. Check the list when the task touches a stored topic (client style, layout habits, review rules) — apply it instead of re-deriving from scratch.', parameters: { type: 'object', properties: { topic: { type: 'string', description: 'Memory topic key. Omit to list available topics first.' } } } } },
    { type: 'function', function: { name: 'write_memory', description: 'SYSTEM. Save a stable work habit/preference to cross-session memory so future sessions inherit it (client style preferences, formatting rules the user corrected you on). action: create (new topic) / append (small addition, preferred) / update (reorganize one topic) / delete. Only save preferences the user explicitly expressed or repeatedly showed; never promote a single occurrence into a rule; timestamp time-sensitive facts; never store secrets or private data. The user can review and delete memory in settings. Optionally set 1-3 triggers (keywords) when the topic binds to a specific document/task type; when the latest user message matches a trigger, that topic is promoted to the top of the memory list.', parameters: { type: 'object', properties: { action: { type: 'string', enum: ['create', 'append', 'update', 'delete'], description: 'create=new topic; append=small addition (preferred); update=rewrite one topic; delete=remove.' }, topic: { type: 'string', description: 'Topic key, snake_case, e.g. ppt_style_habits.' }, title: { type: 'string', description: 'Short title shown in settings (create only).' }, summary: { type: 'string', description: 'One-line summary for the memory list (create/update).' }, content: { type: 'string', description: 'The rule/preference text (create/append/update).' }, triggers: { type: 'array', items: { type: 'string' }, description: 'Keywords binding this topic to a document/task type (e.g. ["路演", "品牌"]). Latest-message match pins the topic to the top of the list. Set on create/update; merged on append.' } }, required: ['action', 'topic'] } } },
    { type: 'function', function: { name: 'ask_clarification', description: 'Ask the user BEFORE acting when the direction is unclear (what to keep, which style, how far to go). Shows a question card with select/text fields and returns the answers keyed by question id. Do NOT guess when you can ask. Not for simple yes/no confirmations that the edit mode already handles.', parameters: { type: 'object', properties: { questions: { type: 'array', description: '1-8 questions.', items: { type: 'object', properties: { id: { type: 'string', description: 'Short key the answer is returned under, e.g. "keep_elements".' }, label: { type: 'string', description: 'The question shown to the user.' }, type: { type: 'string', enum: ['select', 'text'], description: 'select = pick from options (default); text = free input.' }, options: { type: 'array', items: { type: 'string' }, description: 'Required for select, 2-12 options.' } }, required: ['id', 'label'] } }, summary: { type: 'string', description: 'One or two sentences shown above the card: why you are asking.' }, submitLabel: { type: 'string', description: 'Submit button text.' }, explanation: { type: 'string' } }, required: ['questions'] } } },
    { type: 'function', function: { name: 'goto_slide', description: 'Navigate to a slide by index.', parameters: { type: 'object', properties: { index: { type: 'number' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'eval_officejs', description: 'Execute arbitrary Office.js code. Escape hatch for unsupported operations. Your code runs INSIDE an existing PowerPoint.run: use the provided `context` and `PowerPoint` parameters directly, NEVER wrap code in another PowerPoint.run (nested runs silently drop your return value). load() takes comma-separated names or an array (e.g. shape.load("name,left,top") or shape.load(["name"])), not slash paths. Always `return JSON.stringify(yourData)` so the result is readable. Known behaviour of THIS host, do not rediscover it by trial and error: ' + HOST_QUIRKS_TEXT + ' When the code writes to slides, list every affected zero-based slide index in modifiedSlideIndexes so mandatory layout review can run.', parameters: { type: 'object', properties: { code: { type: 'string' }, modifiedSlideIndexes: { type: 'array', items: { type: 'number' }, description: 'Every zero-based slide index modified by this code. Required for slide-writing code.' }, explanation: { type: 'string' } }, required: ['code'] } } },
    { type: 'function', function: { name: 'propose_edits', description: 'PROPOSAL (do not write directly). Use top-level edits for alternative versions of one location, or top-level changes for several independently reviewable locations. Put the recommendation and rationale in the card before review. When the user declines with a reason, your next reply MUST act on it in the same turn — a revised proposal card that absorbs the feedback, a proposed direction for the user to confirm, or what you will do instead; never just save the feedback to memory or acknowledge it and stop.', parameters: { type: 'object', properties: { edits: { type: 'array', description: 'ONE location, optionally with alternative versions.', items: POWERPOINT_EDIT_ITEM }, changes: { type: 'array', description: 'MULTIPLE locations in one review batch. Never nest this inside edits.', items: POWERPOINT_EDIT_ITEM }, explanation: { type: 'string', description: 'Short overall recommendation shown before review.' } }, required: ['explanation'] } } }
  ];

  const SAMPLE_ARGS = {
    get_presentation_outline: { maxSlides: 100 },
    get_slide: { index: 0 },
    get_slide_preview: { index: 0, height: 540 },
    get_selected_slides: {},
    add_slide: {},
    delete_slide: { index: 1 },
    duplicate_slide: { index: 0 },
    get_slide_notes: {},
    set_slide_notes: { index: 0, text: '这一页的讲稿备注。' },
    insert_textbox: { index: 0, text: '标题文字', left: 50, top: 50, width: 400, height: 80 },
    set_text: { index: 0, text: '替换后的文字' },
    apply_layout: { kind: 'layout', label: '新增背景卡片', target: { index: 0, slideId: 'slide-id' }, operation: { type: 'addShape', shape: { shapeType: 'roundRectangle', name: 'AI_Background_Card', left: 48, top: 110, width: 120, height: 260, fillColor: '#F2F2F2', lineVisible: false, zOrder: 'sendToBack' } } },
    insert_image: { base64: '<BASE64_IMAGE>' },
    set_table_style: { target: { index: 0, shapeId: 'table-shape-id' }, scaleFontSize: 0.85 },
    add_table_grid: { index: 0, left: 60, top: 120, width: 480, height: 150, headerFill: '#3F4042', borderColor: '#BFBFBF', fontName: 'Microsoft JhengHei UI', fontSize: 14, data: [['项目', '数值'], ['收入', '1.2亿'], ['利润', '0.3亿']] },
    align_shapes: { target: { index: 0 }, shapeIds: ['shape-id-1', 'shape-id-2', 'shape-id-3'], action: 'left' },
    replace_image: { target: { index: 0, shapeId: 'picture-shape-id' }, base64: '<BASE64_IMAGE>' },
    get_chart: { target: { index: 0, shapeId: 'chart-shape-id' } },
    list_style_templates: {},
    save_style_template: { name: '品牌模板', indexes: [0, 1, 2] },
    set_slide_background: { index: 0, color: '#F5F1E8' },
    edit_table_structure: { target: { index: 0, shapeId: 'table-shape-id' }, action: 'addRow', position: 2 },
    crop_image: { target: { index: 0, shapeId: 'picture-shape-id' }, cropTop: 0.1, cropBottom: 0.1 },
    set_picture_opacity: { target: { index: 0, shapeId: 'picture-shape-id' }, opacity: 0.5 },
    ask_clarification: { questions: [{ id: 'keep_elements', label: '换肤时要保留哪些现有元素？', type: 'select', options: ['全部保留', '只保留 logo', '全部替换'] }] },
    goto_slide: { index: 2 },
    eval_officejs: { code: "const slides = context.presentation.slides;\nslides.load('items');\nawait context.sync();\nreturn slides.items.length;", modifiedSlideIndexes: [] },
    propose_edits: { changes: [{ kind: 'layout', label: '新增背景卡片', target: { index: 1, slideId: 'slide-id' }, operation: { type: 'addShape', shape: { shapeType: 'roundRectangle', left: 48, top: 110, width: 120, height: 260, fillColor: '#F2F2F2', lineVisible: false, zOrder: 'sendToBack' } }, reasoning: '用浅灰色卡片区分信息列', summary: '为第一列增加底色卡片' }], explanation: '保留文字内容，用底色建立列之间的视觉区分。' }
  };
  function defaultArgsForTool(name) { return App.pretty(SAMPLE_ARGS[name] || {}); }

  function revisedPowerPointText(currentText, edit) {
    const current = String(currentText || '');
    const match = App.findSafeTextMatch(current, String(edit.find || ''));
    if (match.status !== 'matched') {
      throw App.makeStaleEditError(match.status === 'ambiguous' ? '目标文字出现多次，无法安全判断要修改哪一处。' : '幻灯片中的原文已变化，请重新读取此项。', { target: edit.target || null, currentText: current });
    }
    const replacement = String(edit.replacement == null ? '' : edit.replacement);
    const placement = edit.placement || 'replace';
    const inserted = placement === 'after' ? match.text + '\n' + replacement
      : placement === 'before' ? replacement + '\n' + match.text
      : replacement;
    return current.slice(0, match.start) + inserted + current.slice(match.end);
  }

  // 应用用户在 Diff 卡上选中的提案版本：只改目标片段，不覆盖同一文本框的其他内容。
  async function applyEdit(edit) {
    if (edit && (edit.kind === 'layout' || edit.operation)) return applyLayout(edit);
    // 【R1/30.2】写入生命周期：进入即计入 pendingDocWrites 并捕获取消代次；检查点在
    // 异步定位完成之后、文本提交之前——停止后不再写入；已提交的写入按实际结果完成。
    const op = App.docWriteBegin('ppt.applyEdit');
    try {
      const target = edit.target || {};
      return await PowerPoint.run(async context => {
      // 【harden-ppt-excel-recovery 2.1】定位骨架与出卡前核验共用 locateTextEditTarget（只有一份语义）；
      // 应用时再核验：出卡后用户可能又改了文档，定位失败即 STALE_EDIT 零写入，
      // 错误带提交 find 与目标当前文本的完整对照（describeTextAnchorProblem）。
      const loc = await locateTextEditTarget(context, edit);
      // 定位唯一 = 恰好一框唯一命中，且没有任何别框同文多次命中的混合歧义（不许猜）
      const entry = loc.located.length === 1 && !loc.ambiguous.length ? loc.located[0] : null;
      const match = entry ? App.findSafeTextMatch(entry.text, String(edit.find || '')) : null;
      if (!entry || !match || match.status !== 'matched') {
        const reason = describeTextAnchorProblem(edit, loc) || '目标文本框或备注中的原文已变化，请重新读取此项。';
        throw App.makeStaleEditError(`生成提案后目标已变化，为避免覆盖你的修改，此项未应用。${reason}`, { target, currentText: entry ? entry.text : undefined });
      }
      const slide = loc.slide;
      const textShape = entry.shape;
      const currentText = entry.text;
      // 保留字符级格式：只对命中片段的 range 做 insertText(Replace)，文本框内其余 run 的
      // 加粗/混排字体/局部颜色保持不动。getSubstring 不可用（旧宿主）时降级为整框替换。
      // 【R1/30.2】异步定位（多段 sync）已完成、文本尚未提交：此处检查取消——
      // 停止发生即 WRITE_CANCELLED，形状保持原文本。
      op.token.throwIfCancelled();
      const replacement = String(edit.replacement == null ? '' : edit.replacement);
      const placement = edit.placement || 'replace';
      const inserted = placement === 'after' ? match.text + '\n' + replacement
        : placement === 'before' ? replacement + '\n' + match.text
        : replacement;
      let partialApplied = false;
      try {
        const subRange = textShape.textFrame.textRange.getSubstring(match.start, match.end - match.start);
        subRange.insertText(inserted, 'Replace');
        partialApplied = true;
      } catch { /* getSubstring 不可用：落入整框替换降级路径 */ }
      if (!partialApplied) textShape.textFrame.textRange.text = revisedPowerPointText(currentText, edit);
      await context.sync();
      return { success: true, index: Number(target.index), partialReplace: true, preservedFormatting: partialApplied, _navTarget: { slideId: slide.id } };
      });
    } finally {
      op.end();
    }
  }

  App.HOSTS.powerpoint = {
    hostType: 'powerpoint',
    available: true,
    metadataLabel: 'Presentation outline',
    systemPrompt: SYSTEM_PROMPT,
    toolDefinitions: TOOL_DEFINITIONS,
    toolExecutors: TOOL_EXECUTORS,
    // 探测摘牌（键 = runtimeProbe 的键）：对应探针被实测为 false 时，api.js 装配工具列表会把该工具摘掉，
    // 不再提供给模型（null = 未探测，保留；只摘实测 false 的）
    capabilityToolGates: { crop_image: 'pictureCrop', set_picture_opacity: 'pictureTransparency', edit_table_structure: 'tableStructureApi', set_hyperlink: 'hyperlinkApi' },
    presentationCapabilities,
    defaultArgsForTool,
    evalToolName: 'eval_officejs',
    getMetadata: getPresentationMetadata,
    getFullContext,
    navigateCitation,
    follow: maybeFollow,
    applyEdit,
    validateEditProposal,
    enrichEditProposal,
    // confirm 模式下可直接使用的纯格式工具（颜色/字体/位置/装饰形状；文字与删页仍走提案卡）。
    // set_hyperlink 只设链接不动文字（textToDisplay 未暴露，不会换显示文本），属格式类可直通（踩坑 11 已核无破坏性分支）
    directFormattingTools: ['apply_layout', 'recolor_slide', 'set_table_style', 'align_shapes', 'set_slide_background', 'crop_image', 'set_picture_opacity', 'set_hyperlink'],
    i18n: {
      zh: {
        brand: 'Trojan AI', brandFooter: 'Trojan AI for Office',
        title: '准备好处理你的演示文稿', subtitle: '你可以让我生成幻灯片、撰写内容或整理结构',
        input: '告诉我你想如何处理这份演示文稿…',
        chart: '幻灯片生成', chartDesc: '按主题快速生成幻灯片与文本',
        fix: '内容润色优化', fixDesc: '精炼要点、统一措辞与结构',
        analyze: '演示结构解析', analyzeDesc: '提炼大纲、生成讲稿备注',
        chartPrompt: '请根据当前演示文稿主题，新增一页合适的幻灯片',
        fixPrompt: '帮我精炼现有幻灯片的要点文字，使其更简洁有力',
        analyzePrompt: '帮我读取演示文稿大纲并总结整体结构',
        demo: '当前不在 PowerPoint/Office 环境中，PPT 工具只能在插件侧边栏里运行。'
      },
      en: {
        brand: 'Trojan AI', brandFooter: 'Trojan AI for Office',
        title: 'Ready to work with your presentation', subtitle: 'Ask me to create slides, write content, or organize structure',
        input: 'Tell me what to do with this presentation…',
        chart: 'Slide Generation', chartDesc: 'Generate slides and text by topic',
        fix: 'Content Polish', fixDesc: 'Refine bullets, unify wording & structure',
        analyze: 'Deck Analysis', analyzeDesc: 'Extract outline and write speaker notes',
        chartPrompt: 'Add a fitting new slide based on the presentation topic',
        fixPrompt: 'Refine the bullet text on the existing slides to be more concise',
        analyzePrompt: 'Read the presentation outline and summarize the overall structure',
        demo: 'Not currently running inside PowerPoint/Office. PowerPoint tools only work in the add-in task pane.'
      }
    }
  };
})();
