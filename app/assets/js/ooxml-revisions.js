(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // ---- OOXML 修订（tracked changes）只读解析 ----
  // 供 Word 的 get_tracked_changes 在原生 TrackedChange API 不可用时兜底（Range.getOoxml 读回
  // 正文 OOXML，解析其中尚存的 w:ins / w:del 修订节点）。纯函数、零 Office 依赖：
  //   App.parseTrackedChangesOoxml(xmlText, { parseXml, contextRadius, maxTextLength })
  // parseXml(xmlText) 必须返回 namespace-aware 的 Document（浏览器传 DOMParser；
  // 测试传 jsdom 的 DOMParser）。解析失败抛 Error，由调用方归入 failed 状态。
  //
  // 规范依据（WORD-TRACKED-CHANGES-IMPLEMENTATION-2026-09-06 4.2 + OpenXML WordprocessingML）：
  //   - 文字插入 = w:ins 内的 w:t；文字删除 = w:del 内的 w:delText。删除文字不在当前正文
  //     文本流里，只有读修订节点才能拿回（这正是原生正文读取会漏掉的部分）。
  //   - 一律按 namespaceURI + localName 判定元素，不依赖固定 w: 前缀（文档换前缀照样解析）。
  //   - 同一修订元素内的多个 run 按文档顺序合并；w:tab→\t、w:br/w:cr→\n；textContent
  //     原样保留空格（application/xml 不做空白折叠）。不把修订元素的整个 textContent
  //     拼进正文——只取 w:t / w:delText 的文本，格式元数据（rPr 等）不混入。
  //   - Flat OPC（pkg:package 包多部件）只解析 contentType 为 wordprocessingml.document.main
  //     的主文档部件，绝不扫描包中其他部件，防止重复计数。
  //   - 段落标记删除（w:pPr/w:rPr/w:del）、格式修订（*PrChange）、移动（moveFrom/moveTo）、
  //     嵌套修订（w:del 内 w:ins 或反之）等无法可靠解释的：记入 unsupportedKinds 并让上层把
  //     结果标为 partial，绝不拼成似是而非的删改文本。
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';

  function isW(node, localName) {
    return node && node.nodeType === 1 && node.namespaceURI === W_NS && node.localName === localName;
  }
  function elementChildren(el) {
    const out = [];
    for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n);
    return out;
  }
  // 属性按 localName 读（前缀无关；xml:space 等带命名空间的属性也照此处理）
  function attrByLocal(el, name) {
    if (!el.attributes) return null;
    for (let i = 0; i < el.attributes.length; i++) {
      const a = el.attributes[i];
      if (a.localName === name) return a.value;
    }
    return null;
  }
  function findDescendant(root, predicate) {
    for (const child of elementChildren(root)) {
      if (predicate(child)) return child;
      const deep = findDescendant(child, predicate);
      if (deep) return deep;
    }
    return null;
  }

  // Flat OPC：pkg:package 下按 contentType 定位主文档部件里的 w:document；
  // 普通 OOXML（根就是 w:document）直接返回根。
  function locateMainDocumentPart(doc) {
    const root = doc.documentElement;
    if (!root) throw new Error('OOXML 为空：没有根元素');
    if (root.namespaceURI === PKG_NS && root.localName === 'package') {
      for (const part of elementChildren(root)) {
        if (!(part.namespaceURI === PKG_NS && part.localName === 'part')) continue;
        const contentType = attrByLocal(part, 'contentType') || '';
        if (!/wordprocessingml\.document\.main\+xml/i.test(contentType)) continue;
        const xmlData = findDescendant(part, n => n.namespaceURI === PKG_NS && n.localName === 'xmlData');
        if (xmlData) {
          const docEl = elementChildren(xmlData).find(n => isW(n, 'document'));
          if (docEl) return docEl;
        }
      }
      throw new Error('Flat OPC 包中未找到 wordprocessingml 主文档部件');
    }
    return root;
  }

  // 修订文字收集：mode 'ins' 取 w:t，'del' 取 w:delText；tab/br/cr 转义。
  // 递归中检测反向修订标记（del 内 ins / ins 内 del）——嵌套修订无法可靠解释，返回 nested 标记，
  // 上层记 unsupportedKind 并放弃该条目的文字（不拼错文本）。非 W 命名空间的子树跳过（其中
  // 可能有 fallback 内容如 mc:AlternateContent，绝不能把两份文字都算一遍）。
  function collectRevisionText(el, mode, state) {
    let out = '';
    for (const child of elementChildren(el)) {
      if (child.namespaceURI !== W_NS) continue;
      const ln = child.localName;
      if (mode === 'ins' && ln === 'del') { state.nested = true; return out; }
      if (mode === 'del' && ln === 'ins') { state.nested = true; return out; }
      if (mode === 'ins' && ln === 't') out += child.textContent || '';
      else if (mode === 'del' && ln === 'delText') out += child.textContent || '';
      else if (mode === 'del' && ln === 't') { state.strayTextInDel = true; }
      else if (ln === 'tab') out += '\t';
      else if (ln === 'br' || ln === 'cr') out += '\n';
      else if (ln === 'noBreakHyphen') out += '-';
      else out += collectRevisionText(child, mode, state);
    }
    return out;
  }

  // 正文「接受后文本」追加：非修订区域的 w:t / tab / br。delText 不进正文流。
  // run 内 rPr 的 *PrChange（格式修订）在此检测——run 文本照收，但缺口要记全。
  function appendAcceptedRun(el, collector, out) {
    for (const child of elementChildren(el)) {
      if (child.namespaceURI !== W_NS) continue;
      const ln = child.localName;
      if (ln === 't') collector.accepted += child.textContent || '';
      else if (ln === 'delText') { /* 已删除文字不在当前正文里 */ }
      else if (ln === 'tab') collector.accepted += '\t';
      else if (ln === 'br' || ln === 'cr') collector.accepted += '\n';
      else if (ln === 'noBreakHyphen') collector.accepted += '-';
      else if (UNSUPPORTED_FORMAT_RE.test(ln)) recordUnsupported(out, 'formatChange');
      else appendAcceptedRun(child, collector, out);
    }
  }

  const UNSUPPORTED_FORMAT_RE = /^(rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange)$/;
  const UNSUPPORTED_MOVE_RE = /^move(From|To)(RangeStart|RangeEnd)?$|^move(From|To)$/;
  const UNSUPPORTED_TABLE_RE = /^(cellIns|cellDel|tblPrEx)$/;

  function recordUnsupported(out, kind) {
    const existing = out.unsupportedKinds.find(k => k.kind === kind);
    if (existing) existing.count++;
    else out.unsupportedKinds.push({ kind, count: 1 });
  }

  // 单个修订元素 → 条目（author/date/id 按 localName 读属性）
  function makeChange(out, kind, el, text, collector, tableCtx) {
    const pos = collector.accepted.length;
    const entry = {
      kind, // 'insert' | 'delete'
      author: attrByLocal(el, 'author'),
      date: attrByLocal(el, 'date'),
      revisionId: attrByLocal(el, 'id'),
      text,
      posStart: kind === 'insert' ? pos : pos,
      posEnd: kind === 'insert' ? pos + text.length : pos
    };
    if (tableCtx) entry.location = Object.assign({}, tableCtx);
    collector.entries.push(entry);
  }

  function finalizeParagraph(collector, out, radius) {
    for (const entry of collector.entries) {
      const before = collector.accepted.slice(Math.max(0, entry.posStart - radius), entry.posStart);
      const after = collector.accepted.slice(entry.posEnd, entry.posEnd + radius);
      out.changes.push({
        kind: entry.kind,
        author: entry.author,
        date: entry.date,
        revisionId: entry.revisionId,
        text: entry.text,
        contextBefore: (entry.posStart - radius > 0 ? '…' : '') + before,
        contextAfter: after + (entry.posEnd + radius < collector.accepted.length ? '…' : ''),
        location: Object.assign({ paragraphOrdinal: out.paragraphOrdinal }, entry.location || {})
      });
    }
  }

  function parseTrackedChangesOoxml(xmlText, options = {}) {
    const radius = Math.max(10, Number(options.contextRadius || 42));
    const maxTextLength = Math.max(50, Number(options.maxTextLength || 2000));
    const parseXml = options.parseXml
      || (typeof DOMParser === 'function' ? (t) => new DOMParser().parseFromString(t, 'application/xml') : null);
    if (!parseXml) throw new Error('当前环境没有可用的 XML 解析器（DOMParser）');

    const doc = parseXml(String(xmlText));
    const parserErrors = doc.getElementsByTagName('parsererror');
    if (parserErrors.length) {
      throw new Error('OOXML 解析失败：' + String(parserErrors[0].textContent || '').slice(0, 200));
    }

    const partRoot = locateMainDocumentPart(doc);
    const bodyEl = findDescendant(partRoot, n => isW(n, 'body')) || partRoot;

    const out = { changes: [], unsupportedKinds: [], warnings: [], paragraphOrdinal: -1, tableCount: 0 };

    // walk：tableCtx = 最近表格上下文（inTable + tableIndex，嵌套表格继承最近外层序号）
    function walk(el, collector, tableCtx, inRPr) {
      for (const child of elementChildren(el)) {
        if (child.namespaceURI !== W_NS) continue;
        const ln = child.localName;
        if (ln === 'p') { walkParagraph(child, tableCtx); continue; }
        if (ln === 'tbl') {
          const tableIndex = out.tableCount++;
          walk(child, collector, { inTable: true, tableIndex }, false);
          continue;
        }
        if (ln === 'ins' && !inRPr) {
          const state = {};
          const text = collectRevisionText(child, 'ins', state);
          if (state.nested) { recordUnsupported(out, 'nestedRevision'); continue; }
          if (state.strayTextInDel) out.warnings.push('w:del 内出现 w:t（应为 w:delText），该部分文字未计入删除文本');
          makeChange(out, 'insert', child, text, collector, tableCtx);
          collector.accepted += text; // 插入文字是接受后正文的一部分
          continue;
        }
        if (ln === 'del') {
          if (inRPr) { recordUnsupported(out, 'paragraphMarkDelete'); continue; }
          const state = {};
          const text = collectRevisionText(child, 'del', state);
          if (state.nested) { recordUnsupported(out, 'nestedRevision'); continue; }
          if (state.strayTextInDel) out.warnings.push('w:del 内出现 w:t（应为 w:delText），该部分文字未计入删除文本');
          makeChange(out, 'delete', child, text, collector, tableCtx);
          continue;
        }
        if (UNSUPPORTED_FORMAT_RE.test(ln)) { recordUnsupported(out, 'formatChange'); continue; }
        if (UNSUPPORTED_MOVE_RE.test(ln)) { recordUnsupported(out, 'move'); continue; }
        if (UNSUPPORTED_TABLE_RE.test(ln)) { recordUnsupported(out, 'tableStructureChange'); continue; }
        if (ln === 'r') { appendAcceptedRun(child, collector, out); continue; }
        // 其余容器（body/document/sdt/hyperlink/pPr/rPr/tc/tr 等）继续深入。
        // rPr 只为探测段落标记删除（w:pPr/w:rPr/w:del），其中的 w:del 不是文字删除。
        walk(child, collector, tableCtx, inRPr || ln === 'rPr');
      }
    }

    function walkParagraph(pEl, tableCtx) {
      out.paragraphOrdinal++;
      const collector = { accepted: '', entries: [] };
      walk(pEl, collector, tableCtx, false);
      finalizeParagraph(collector, out, radius);
    }

    walk(bodyEl, { accepted: '', entries: [] }, null, false);

    // 全局条目编号 + 截断标注（截断不影响文字本身的正确性，expand 入口在工具层）
    out.changes.forEach((c, index) => {
      c.index = index;
      c.refId = `tc:${index}`;
      if (c.text.length > maxTextLength) {
        c.text = c.text.slice(0, maxTextLength);
        c.truncated = true;
      }
    });
    out.paragraphCount = out.paragraphOrdinal + 1;
    delete out.paragraphOrdinal;
    return out;
  }

  App.parseTrackedChangesOoxml = parseTrackedChangesOoxml;
})();
