(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const requireOffice = () => App.requireOffice();

  // ---- 段落锚点 id：用 Word 内容控件 tag 作为稳定引用，回退到段落序号 ----
  function clampText(s, n) { s = String(s == null ? '' : s); return s.length > n ? s.slice(0, n) + '…' : s; }
  // 真机 Word 的单元格 value 尾部可能带 \x07(单元格结束符)——复核 P3-5:不剥掉的话
  // 「空单元格」判定(trim 滤不掉 \x07)和写后读回比对都会被它骗。所有消费 cell.value 的
  // 地方统一经 cellText。夹具按带 \x07 建模,防实现与夹具共谋。
  function cellText(v) { return String(v ?? '').replace(/\x07/g, ''); }

  // 最长公共前缀/后缀长度（长文本差分替换用，下标均为 JS 字符串的 UTF-16 单元，与 Word range 偏移一致）
  function commonPrefixLength(a, b) {
    const n = Math.min(a.length, b.length);
    let i = 0;
    while (i < n && a[i] === b[i]) i++;
    return i;
  }
  function commonSuffixLength(a, b, prefixLen) {
    const n = Math.min(a.length, b.length) - prefixLen;
    let i = 0;
    while (i < n && a[a.length - 1 - i] === b[b.length - 1 - i]) i++;
    return i;
  }

  // insertText(Replace) 的新文字会继承插入点前后 run 的格式——比如紧跟在加粗引导语
  // （"Text:"）后面插入正文，整段会被「传染」成粗体。替换前先抓被替换区间首字符的字体
  // 属性，插入后强制套回，保证「只改文字、不动格式」。
  async function insertTextPreservingFormat(context, targetRange, text, token) {
    // 【24.2】写命令提交前的取消检查：函数入口（进入本函数前的等待不计入）。
    if (token) token.throwIfCancelled();
    const font = targetRange.font;
    // 【25.3】无 font 快速路径同样是正文写命令：提交前检查取消
    if (!font || typeof font.load !== 'function') {
      if (token) token.throwIfCancelled();
      return targetRange.insertText(text, Word.InsertLocation.replace);
    }
    const props = ['name', 'size', 'bold', 'italic', 'underline', 'color', 'highlightColor'];
    let snapshot = null;
    try {
      font.load(props);
      await context.sync();
      snapshot = {};
      for (const k of props) { const v = font[k]; if (v !== null && v !== undefined) snapshot[k] = v; }
    } catch (e) { snapshot = null; } // 快照读不出来就裸替换，不能让格式问题搞砸文字修改
    // 【25.3】字体快照读取（含失败降级）结束之后、insertText 紧前的最终取消检查——
    // 停止若发生在字体读取等待中，在此拦截；该检查不在任何 fallback try/catch 之内
    if (token) token.throwIfCancelled();
    const insertedRange = targetRange.insertText(text, Word.InsertLocation.replace);
    await context.sync(); // 文字写入独立成批：这步失败才是真的应用失败
    if (snapshot) {
      // 格式套回单独成批、best-effort：宿主不认某个属性时（真机出现过 Invalid argument）
      // 只损失格式保真度，绝不把已经写成功的文字标成「应用失败」（卡片说谎比格式瑕疵更糟）
      try {
        const newFont = insertedRange.font;
        for (const k of Object.keys(snapshot)) newFont[k] = snapshot[k];
        await context.sync();
      } catch (e) {}
    }
    return insertedRange;
  }

  // ---- 写后读回验证（W2）----
  // 学 Excel 批 21 教训：typeof 存在性检查会被 Office.js 代理对象骗，写后读回才算数。
  // 三态：verified:true（读回相符）/ verified:false + verifyWarning（读回了但不符，疑似未生效）/
  // verified:null + verifyError（读回动作本身抛错——写入可能已生效但无法确认，与不匹配分开报）。
  // 验证绝不抛错打断已成功的写入，只把结论写进结果对象，不静默假成功。
  async function verifyWriteBack(readBack) {
    let mismatches;
    try {
      mismatches = await readBack();
    } catch (e) {
      return { verified: null, verifyError: `读回验证本身失败（写入可能已生效，无法确认）：${String(e && e.message || e)}` };
    }
    if (mismatches && mismatches.length) {
      return { verified: false, verifyWarning: `疑似未生效：${mismatches.join('；')}` };
    }
    return { verified: true };
  }
  // 读回必须换新代理重新取回：在原代理上读到的可能只是本地赋值缓存，不是宿主真实状态。
  function freshRange(range) {
    return (range && typeof range.getRange === 'function') ? range.getRange() : range;
  }

  // 锚定解析三态（真机反馈「给第 N 段加批注/居中某标题」做不到，模型无法把「第二段」变成选区）：
  // 都不传 → 当前选区；paragraphIndex → 该段整段 range（越界报错）；find → 全文唯一命中区间
  // （0 命中/多命中报错）。两者都传时 paragraphIndex 为准，find 仅作段内校验。
  // manageComment add / applyStyle / setParagraphFormat 三处共用，不许复制粘贴三份。
  async function resolveAnchorRange(context, args) {
    const hasParagraphIndex = args.paragraphIndex !== undefined && args.paragraphIndex !== null && args.paragraphIndex !== '';
    const find = String(args.find || '').trim();
    if (hasParagraphIndex) {
      const idx = Number(args.paragraphIndex);
      const paras = context.document.body.paragraphs;
      paras.load('items');
      await context.sync();
      if (!Number.isInteger(idx) || idx < 0 || !paras.items[idx]) {
        throw new Error(`Paragraph ${args.paragraphIndex} not found (document has ${paras.items.length} paragraphs; valid index 0-${Math.max(0, paras.items.length - 1)})`);
      }
      const anchorRange = paras.items[idx].getRange();
      if (find) {
        const hits = anchorRange.search(find, { matchCase: false, matchWholeWord: false, matchWildcards: false });
        hits.load('items');
        await context.sync();
        if (!hits.items.length) throw new Error(`find '${clampText(find, 60)}' not found inside paragraph ${idx} (paragraphIndex wins; find is only a verification)`);
      }
      return { anchorRange, anchorMode: 'paragraph' };
    }
    if (find) {
      const hits = context.document.body.search(find, { matchCase: false, matchWholeWord: false, matchWildcards: false });
      hits.load('items');
      await context.sync();
      if (!hits.items.length) throw new Error(`find '${clampText(find, 60)}' not found anywhere in the document`);
      if (hits.items.length > 1) throw new Error(`find '${clampText(find, 60)}' matches ${hits.items.length} places; provide a longer, unique anchor text (or pass paragraphIndex)`);
      return { anchorRange: hits.items[0], anchorMode: 'find' };
    }
    return { anchorRange: context.document.getSelection(), anchorMode: 'selection' };
  }

  // collapsed 防御（真机四连修①②）：无锚定且选区为空（只有光标）时，格式写进光标处而非用户
  // 说的目标段；读回验证验的是同一个错误目标，verified:true 也是假成功。显式 note，不再静默。
  function collapsedAnchorNote(extra) {
    return `未指定目标且选区为空（只有光标），格式仅作用于光标处。${extra ? `${extra}。` : ''}要作用于指定段落，请传 paragraphIndex 或 find 锚定。`;
  }

  async function getDocumentOutline(args = {}) {
    requireOffice();
    const maxParagraphs = Math.max(1, Number(args.maxParagraphs || 400));
    return Word.run(async context => {
      const body = context.document.body;
      const paras = body.paragraphs;
      paras.load('items');
      await context.sync();
      const items = paras.items.slice(0, maxParagraphs);
      items.forEach(p => p.load('text,styleBuiltIn,style'));
      await context.sync();
      const outline = [];
      const paragraphs = [];
      items.forEach((p, idx) => {
        const text = (p.text || '').trim();
        const style = p.styleBuiltIn || p.style || '';
        const isHeading = /heading/i.test(String(style));
        const level = isHeading ? (parseInt(String(style).replace(/\D/g, ''), 10) || 1) : 0;
        if (text || isHeading) paragraphs.push({ index: idx, text: clampText(text, 600), style: String(style) });
        if (isHeading && text) outline.push({ index: idx, level, text: clampText(text, 200) });
      });
      return { success: true, paragraphCount: paras.items.length, returned: items.length, hasMore: paras.items.length > items.length, outline, paragraphs, note: 'paragraphs[].text 可能被截断（尾部有 … 表示截断）。不要把截断文本用作 propose_edits 的 find，必须先用 get_paragraphs 读取完整原文。' };
    });
  }

  async function getSelection() {
    requireOffice();
    return Word.run(async context => {
      const sel = context.document.getSelection();
      sel.load('text,styleBuiltIn,style,font/bold,font/italic,font/size,font/name,font/color');
      await context.sync();
      const result = { success: true, text: sel.text || '', style: String(sel.styleBuiltIn || sel.style || ''), font: { bold: sel.font.bold, italic: sel.font.italic, size: sel.font.size, name: sel.font.name, color: sel.font.color } };
      // 光标只是点进某段（collapsed，无选中文本）时，额外返回所在段文本，供选区 chip 注入给模型
      if (!result.text) {
        const para = sel.paragraphs.getFirst();
        para.load('text');
        await context.sync();
        const paragraphText = clampText(para.text || '', 600);
        if (paragraphText.trim()) {
          result.collapsed = true;
          result.paragraph = { text: paragraphText };
        }
      }
      return result;
    });
  }

  async function getParagraphs(args = {}) {
    requireOffice();
    const start = Math.max(0, Number(args.start || 0));
    const count = Math.max(1, Number(args.count || 20));
    return Word.run(async context => {
      const paras = context.document.body.paragraphs;
      paras.load('items');
      await context.sync();
      const slice = paras.items.slice(start, start + count);
      slice.forEach(p => p.load('text,styleBuiltIn'));
      await context.sync();
      return { success: true, total: paras.items.length, paragraphs: slice.map((p, i) => ({ index: start + i, text: p.text || '', style: String(p.styleBuiltIn || '') })) };
    });
  }

  function contextWindow(paragraphText, match) {
    const text = String(paragraphText || '');
    const radius = 42;
    const beforeStart = Math.max(0, match.start - radius);
    const afterEnd = Math.min(text.length, match.end + radius);
    return {
      contextBefore: (beforeStart > 0 ? '…' : '') + text.slice(beforeStart, match.start),
      contextAfter: text.slice(match.end, afterEnd) + (afterEnd < text.length ? '…' : ''),
      contextIsWholeParagraph: match.start === 0 && match.end === text.length
    };
  }

  // ---- 表格段落锚定（2026-09-01 W3 提前批）：表格单元格的段落不在 body.paragraphs 里，
  // 提案卡 find 锚到表格文字时永远报「缺少可辨认的前后文」。这里把正文段落与
  // body.tables → rows → cells → cell.body.paragraphs 逐层合并成统一候选集。
  // API 面已 grep vendored word-mac-16.00.js（16.0.20416.15170）取证：
  //   Body.tables → TableCollection（"Body.tables"，WordApi 1.3）；
  //   Table.rows → TableRowCollection；TableRow.cells → TableCellCollection；
  //   TableCell.body → Body → .paragraphs（ParagraphCollection）；
  //   TableCell.value 可读写、rowIndex/cellIndex 可读；Table.rowCount 可读。
  // 返回的 body 候选带 { paragraph, index }（index 对 paragraphIndex 语义负责）；
  // 表格候选带 { paragraph, tableIndex, rowIndex, columnIndex }、无 index——
  // 表格没有段落序号，paragraphIndex 锚定不适用（精确 find + 上下文消歧照旧）。
  // 性能守卫（参照 clampText/maxParagraphs 风格）：表格多于 maxTables 或单元格总数
  // 超过 maxCells 时截断并在 scanNote 里明说；不带 scanNote 即全量扫描。
  async function loadBodyParagraphs(context, options = {}) {
    const maxTables = Math.max(0, Number(options.maxTables || 20));
    const maxCells = Math.max(1, Number(options.maxCells || 400));
    const body = context.document.body;
    const bodyParas = body.paragraphs;
    bodyParas.load('items');
    // 防御：宿主/测试面没有 body.tables 时只扫正文段落（不报错）
    const tables = body.tables;
    const hasTables = !!(tables && typeof tables.load === 'function');
    if (hasTables) tables.load('items');
    await context.sync();
    const tableItems = hasTables ? (tables.items || []) : [];
    const candidates = bodyParas.items.map((p, index) => ({ paragraph: p, index }));
    const scan = { tableCount: tableItems.length, tablesScanned: 0, cellsScanned: 0 };
    let cellsTruncated = false;
    for (const t of tableItems.slice(0, maxTables)) {
      if (scan.cellsScanned >= maxCells) { cellsTruncated = true; break; }
      const rows = t.rows;
      rows.load('items');
      await context.sync();
      for (const row of rows.items) {
        if (scan.cellsScanned >= maxCells) { cellsTruncated = true; break; }
        const cells = row.cells;
        cells.load('items');
        await context.sync();
        for (const cell of cells.items) {
          if (scan.cellsScanned >= maxCells) { cellsTruncated = true; break; }
          const cellParas = cell.body.paragraphs;
          cellParas.load('items');
          cell.load('rowIndex,cellIndex');
          await context.sync();
          cellParas.items.forEach(p => candidates.push({ paragraph: p, tableIndex: scan.tablesScanned, rowIndex: Number(cell.rowIndex), columnIndex: Number(cell.cellIndex) }));
          scan.cellsScanned++;
        }
        if (cellsTruncated) break;
      }
      if (!cellsTruncated) scan.tablesScanned++;
    }
    const truncated = tableItems.length > maxTables || cellsTruncated;
    if (truncated) {
      scan.scanNote = `文档含 ${scan.tableCount} 张表格，本次仅扫描前 ${scan.tablesScanned} 张（${scan.cellsScanned} 个单元格）内的段落；如需定位未扫描区域，请缩小范围或用 search_text/get_tables 确认。`;
    }
    return { bodyItems: bodyParas.items, candidates, ...scan };
  }

  // Word 卡片的上下文由当前段落自动补齐，不依赖模型主动填写。
  async function enrichEditProposal(args) {
    const normalized = Object.assign({}, args || {});
    const key = Array.isArray(normalized.changes) && normalized.changes.length ? 'changes' : 'edits';
    const items = Array.isArray(normalized[key]) ? normalized[key].map(item => Object.assign({}, item)) : [];
    normalized[key] = items;
    const textItems = items.filter(item => item && item.kind !== 'layout' && String(item.find || '').trim());
    if (!textItems.length) return normalized;

    return Word.run(async context => {
      // 候选集 = 正文段落 + 表格单元格段落（W3 提前批：find 锚到表格文字时也要能核实填充）。
      // paragraphIndex 语义维持只对正文段落（bodyItems）——表格段落没有序号，find 命中即可核实。
      const { bodyItems, candidates } = await loadBodyParagraphs(context);
      const requested = Array.from(new Set(textItems.map(item => Number(item.target && item.target.paragraphIndex)).filter(index => Number.isInteger(index) && index >= 0 && bodyItems[index])));
      const needsGlobalSearch = textItems.some(item => {
        const index = Number(item.target && item.target.paragraphIndex);
        return !Number.isInteger(index) || index < 0 || !bodyItems[index];
      });
      const loaded = needsGlobalSearch ? candidates.map(c => c.paragraph) : requested.map(index => bodyItems[index]);
      loaded.forEach(paragraph => paragraph.load('text'));
      await context.sync();

      const unmatched = [];
      for (const item of textItems) {
        const find = String(item.find || '');
        const index = Number(item.target && item.target.paragraphIndex);
        let paragraph = Number.isInteger(index) && index >= 0 ? bodyItems[index] : null;
        let match = paragraph ? App.findSafeTextMatch(String(paragraph.text || ''), find) : { status: 'missing' };
        if (match.status !== 'matched' && needsGlobalSearch) {
          const locations = [];
          for (const candidate of candidates) {
            const candidateMatch = App.findSafeTextMatch(String(candidate.paragraph.text || ''), find);
            if (candidateMatch.status === 'matched') locations.push({ paragraph: candidate.paragraph, match: candidateMatch });
          }
          if (locations.length === 1) { paragraph = locations[0].paragraph; match = locations[0].match; }
        }
        if (paragraph && match.status === 'matched') Object.assign(item, contextWindow(paragraph.text, match));
        else unmatched.push(`「${find.slice(0, 30)}」`);
      }
      // 锚点全部落空 = 提案在当前文档上必然 STALE（v141 真机：空文档插入时模型连发
      // 占位提案全部过期告终）。在出卡前拒绝并指明正确路径，比出卡后再 STALE 循环强得多。
      if (textItems.length && unmatched.length === textItems.length) {
        const bodyEmpty = bodyItems.every(p => !String(p.text || '').trim());
        throw new Error(bodyEmpty
          ? '文档当前为空，没有任何文字可作提案锚点。纯新增内容（标题/段落/备注）请直接调用 insert_heading 或 insert_text——确认模式下纯插入直通生效、可 ⌘Z 撤销，不需要提案卡。'
          : `提案的 find 锚点在文档中全部未命中（${unmatched.join('、')}）。请先用 get_paragraphs 读取真实原文，引用确实存在的文字作为锚点后重新提案。`);
      }
      return normalized;
    });
  }

  async function getTables(args = {}) {
    requireOffice();
    const maxTables = Math.max(1, Number(args.maxTables || 20));
    return Word.run(async context => {
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      const slice = tables.items.slice(0, maxTables);
      slice.forEach(t => t.load('values,rowCount'));
      await context.sync();
      return { success: true, tableCount: tables.items.length, tables: slice.map((t, i) => ({ index: i, rowCount: t.rowCount, values: t.values })) };
    });
  }

  async function searchText(args) {
    requireOffice();
    const { query, matchCase = false, matchWholeWord = false, useWildcards = false, maxResults = 100 } = args;
    if (!query) throw new Error('query is required');
    return Word.run(async context => {
      const results = context.document.body.search(query, { matchCase, matchWholeWord, matchWildcards: useWildcards });
      results.load('items');
      await context.sync();
      const slice = results.items.slice(0, maxResults);
      slice.forEach(r => r.load('text'));
      await context.sync();
      return { success: true, totalFound: results.items.length, returned: slice.length, matches: slice.map((r, i) => ({ index: i, text: clampText(r.text, 160) })) };
    });
  }

  // 仅在跟随模式开启时选中目标 range（select 会滚动视图）。统一受 followMode 控制，
  // 避免在工具内部硬编码 select 而绕过开关。在 range 仍有效的 Word.run 内调用。
  function followSelect(range) {
    if (App.state.settings.followMode) { try { range.select(); } catch {} }
  }

  const WHERE_MAP = { Replace: 'Replace', Start: 'Start', End: 'End', Before: 'Before', After: 'After' };

  // 闸口模式读法与下发层对齐（复核 P2-1）：auto 下用户说「审核后再改」时，下发层
  // requestAccessMode(api.js)会降级为 confirm 语义，执行器闸口必须同口径——否则 Replace 这类
  // 覆盖操作会在「用户以为在审核」时直通。Excel 侧六个闸口是同一个既有洞，留给 Excel 线统一修。
  function gateMode() {
    const raw = typeof App.currentAccessMode === 'function' ? App.currentAccessMode() : (App.state.settings.accessMode || 'confirm');
    if (raw === 'auto' && typeof App.requestRequiresReview === 'function' && App.requestRequiresReview()) return 'confirm';
    return raw;
  }

  async function insertText(args) {
    requireOffice();
    const { text = '', location = 'End', style } = args;
    const where = WHERE_MAP[location] || 'End';
    // confirm 闸门（踩坑 11：按操作分不按工具分）：工具本身在 confirm 下直通（2026-09-01 空文档
    // + 审核模式死锁修复，纯增量插入可撤销、无审批价值），但 Replace 是覆盖选区的破坏性操作，
    // 拦下引导走提案卡（同 manageComment delete 分支先例）。
    const accessMode = gateMode();
    if (accessMode === 'confirm' && where === 'Replace') {
      return { success: false, blocked: true, error: '替换选区文字在确认模式下不可用。请改用 propose_edits 生成修改卡片，由用户确认后应用；Start/End/Before/After 等纯插入不受限。' };
    }
    return Word.run(async context => {
      const body = context.document.body;
      let range;
      if (where === 'Replace' || where === 'Before' || where === 'After') {
        const sel = context.document.getSelection();
        range = sel.insertText(text, where === 'Replace' ? Word.InsertLocation.replace : (where === 'Before' ? Word.InsertLocation.before : Word.InsertLocation.after));
      } else {
        range = body.insertText(text, where === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end);
      }
      if (style) range.styleBuiltIn = style;
      followSelect(range);
      await context.sync();
      // 写后读回（W2）：换代理读回文本（及 style），与刚写入的内容比对
      const verification = await verifyWriteBack(async () => {
        const fresh = freshRange(range);
        fresh.load(style ? 'text,styleBuiltIn' : 'text');
        await context.sync();
        const bad = [];
        if (String(fresh.text || '').replace(/\r\n?/g, '\n') !== String(text).replace(/\r\n?/g, '\n')) bad.push(`文本读回为 ${clampText(JSON.stringify(String(fresh.text)), 60)}，与写入的 ${clampText(JSON.stringify(String(text)), 60)} 不符`);
        if (style && String(fresh.styleBuiltIn || '') !== String(style)) bad.push(`样式读回为「${fresh.styleBuiltIn}」，预期「${style}」`);
        return bad;
      });
      return { success: true, inserted: clampText(text, 120), location: where, ...verification, _navTarget: { kind: 'selection' } };
    });
  }

  async function replaceText(args) {
    requireOffice();
    const { query, replacement = '', matchCase = false, matchWholeWord = false, useWildcards = false, replaceAll = true } = args;
    if (!query) throw new Error('query is required');
    return Word.run(async context => {
      const results = context.document.body.search(query, { matchCase, matchWholeWord, matchWildcards: useWildcards });
      results.load('items');
      await context.sync();
      const targets = replaceAll ? results.items : results.items.slice(0, 1);
      const insertedRanges = targets.map(r => r.insertText(replacement, Word.InsertLocation.replace));
      await context.sync();
      // 写后读回（W2）：insertText 返回的是新文本的 range（非本地赋过值的代理），逐一读回比对
      let verification = { verified: true };
      if (targets.length) {
        verification = await verifyWriteBack(async () => {
          insertedRanges.forEach(r => { if (r && typeof r.load === 'function') r.load('text'); });
          await context.sync();
          const bad = [];
          insertedRanges.forEach((r, i) => {
            const got = r ? String(r.text || '') : '';
            if (got !== String(replacement)) bad.push(`第 ${i + 1} 处替换读回为 ${clampText(JSON.stringify(got), 60)}，预期 ${clampText(JSON.stringify(String(replacement)), 60)}`);
          });
          return bad;
        });
      }
      return { success: true, replaced: targets.length, totalFound: results.items.length, ...verification, _navTarget: { kind: 'selection' } };
    });
  }

  async function applyStyle(args) {
    requireOffice();
    const { target = 'selection', style, font = {} } = args;
    return Word.run(async context => {
      // 锚定三态（与 manageComment add 同一套 resolveAnchorRange）：传了 paragraphIndex/find 时
      // 锚定优先于 target；否则退回 target（selection 默认 / document）。
      const hasAnchor = (args.paragraphIndex !== undefined && args.paragraphIndex !== null && args.paragraphIndex !== '') || !!String(args.find || '').trim();
      let range, anchorMode;
      if (hasAnchor) {
        const resolved = await resolveAnchorRange(context, args);
        range = resolved.anchorRange;
        anchorMode = resolved.anchorMode;
      } else if (target === 'document') {
        range = context.document.body.getRange();
        anchorMode = 'document';
      } else {
        range = context.document.getSelection();
        anchorMode = 'selection';
      }
      // collapsed 防御：isEmpty 随下面第一次 sync 一并取回（Range.isEmpty 已在新包，grep 取证）
      if (anchorMode === 'selection') range.load('isEmpty');
      if (style) range.styleBuiltIn = style;
      if (font.bold != null) range.font.bold = !!font.bold;
      if (font.italic != null) range.font.italic = !!font.italic;
      if (font.underline != null) range.font.underline = font.underline ? 'Single' : 'None';
      if (font.size) range.font.size = font.size;
      if (font.name) range.font.name = font.name;
      if (font.color) range.font.color = font.color;
      await context.sync();
      // 写后读回（W2）：只比对本次设置过的属性，未指定的属性不比（避免误报）
      const verification = await verifyWriteBack(async () => {
        const loads = [];
        if (style) loads.push('styleBuiltIn');
        if (font.bold != null) loads.push('font/bold');
        if (font.italic != null) loads.push('font/italic');
        if (font.underline != null) loads.push('font/underline');
        if (font.size) loads.push('font/size');
        if (font.name) loads.push('font/name');
        if (font.color) loads.push('font/color');
        if (!loads.length) return [];
        const fresh = freshRange(range);
        fresh.load(loads.join(','));
        await context.sync();
        const bad = [];
        if (style && String(fresh.styleBuiltIn || '') !== String(style)) bad.push(`样式读回为「${fresh.styleBuiltIn}」，预期「${style}」`);
        const f = fresh.font || {};
        if (font.bold != null && !!f.bold !== !!font.bold) bad.push(`粗体读回为 ${f.bold}，预期 ${!!font.bold}`);
        if (font.italic != null && !!f.italic !== !!font.italic) bad.push(`斜体读回为 ${f.italic}，预期 ${!!font.italic}`);
        if (font.underline != null && String(f.underline) !== (font.underline ? 'Single' : 'None')) bad.push(`下划线读回为 ${f.underline}，预期 ${font.underline ? 'Single' : 'None'}`);
        if (font.size && Number(f.size) !== Number(font.size)) bad.push(`字号读回为 ${f.size}，预期 ${font.size}`);
        if (font.name && String(f.name) !== String(font.name)) bad.push(`字体读回为「${f.name}」，预期「${font.name}」`);
        if (font.color && String(f.color || '').toLowerCase() !== String(font.color).toLowerCase()) bad.push(`字色读回为 ${f.color}，预期 ${font.color}`);
        return bad;
      });
      // collapsed 防御（真机四连修①）：无锚定且选区为空时格式只落在光标处，读回验的是同一个
      // 错误目标——note 里说破；字体属性在光标处只是键入属性（无可见效果）也要说清。
      const hasFontProps = Object.values(font).some(v => v !== undefined && v !== null);
      const note = (anchorMode === 'selection' && range.isEmpty)
        ? collapsedAnchorNote(hasFontProps ? '字体属性在光标处只是键入属性，对已有文字无可见效果' : '')
        : undefined;
      return { success: true, target, anchorMode, style: style || null, ...(note ? { note } : {}), ...verification, _navTarget: { kind: 'selection' } };
    });
  }

  async function setParagraphFormat(args) {
    requireOffice();
    const { target = 'selection', alignment, lineSpacing, lineSpacingMultiple, leftIndent, spaceBefore, spaceAfter } = args;
    const props = {};
    if (alignment) props.alignment = alignment; // Left/Centered/Right/Justified
    // 行距单位（真机反馈：模型按「倍」传 1.5，Word 当磅值写入，对话框显示 0.3 倍）。
    // Word OM 多倍行距约定为 12 磅/倍：1.5 倍 = 18 磅。优先收 lineSpacingMultiple(倍数),换算成磅;
    // lineSpacing 保留为磅值直传通道。本机 runtime 无 lineSpacingRule(vendored 零命中),只能靠该约定换算。
    if (lineSpacingMultiple) props.lineSpacing = Number(lineSpacingMultiple) * 12;
    else if (lineSpacing) props.lineSpacing = lineSpacing;
    if (leftIndent != null) props.leftIndent = leftIndent;
    if (spaceBefore != null) props.spaceBefore = spaceBefore;
    if (spaceAfter != null) props.spaceAfter = spaceAfter;
    const keys = Object.keys(props);
    return Word.run(async context => {
      // 锚定三态（与 manageComment add 同一套 resolveAnchorRange）：传了 paragraphIndex/find 时
      // 锚定优先于 target；否则退回 target（selection 默认 / document）。
      const hasAnchor = (args.paragraphIndex !== undefined && args.paragraphIndex !== null && args.paragraphIndex !== '') || !!String(args.find || '').trim();
      let range, anchorMode;
      if (hasAnchor) {
        const resolved = await resolveAnchorRange(context, args);
        range = resolved.anchorRange;
        anchorMode = resolved.anchorMode;
      } else if (target === 'document') {
        range = context.document.body.getRange();
        anchorMode = 'document';
      } else {
        range = context.document.getSelection();
        anchorMode = 'selection';
      }
      // collapsed 防御：isEmpty 随本批第一次 sync 一并取回（Range.isEmpty 已在新包，grep 取证）
      if (anchorMode === 'selection') range.load('isEmpty');
      // 2026-08-31 复核 vendored word-mac-16.00.js（16.0.20416.15170）：ParagraphFormat 类已进包，
      // 但只挂在 Style（门控 WordApi 1.5）与 ConditionalStyle 上，Range 类体内 paragraphFormat
      // 仍零命中——本机 runtime 下 range.paragraphFormat 仍是 undefined，段落格式属性直接挂在
      // Paragraph 上且标记可写（_scalarPropertyUpdateable）。官方新面有 paragraphFormat 时优先
      // 走它；没有则逐段写入。这是对客户端 getter 的存在性判断，不是对宿主能力的猜测。
      const pf = range.paragraphFormat;
      if (pf) {
        for (const k of keys) pf[k] = props[k];
      } else if (keys.length) {
        const paras = range.paragraphs;
        paras.load('items');
        await context.sync();
        paras.items.forEach(p => { for (const k of keys) p[k] = props[k]; });
      }
      await context.sync();
      // 写后读回（W2）：换代理按同一路径读回刚设置的每个属性，逐段比对
      const verification = await verifyWriteBack(async () => {
        if (!keys.length) return [];
        const fresh = freshRange(range);
        const readFrom = [];
        const pf2 = fresh.paragraphFormat;
        if (pf2) {
          pf2.load(keys.join(','));
          await context.sync();
          readFrom.push(pf2);
        } else {
          const paras2 = fresh.paragraphs;
          paras2.load('items');
          await context.sync();
          paras2.items.forEach(p => p.load(keys.join(',')));
          await context.sync();
          readFrom.push(...paras2.items);
        }
        const bad = [];
        readFrom.forEach((obj, i) => {
          keys.forEach(k => {
            const ok = Number.isFinite(props[k]) ? Number(obj[k]) === Number(props[k]) : String(obj[k]) === String(props[k]);
            if (!ok) bad.push(`${k} 第 ${i + 1} 段读回为 ${obj[k]}，预期 ${props[k]}`);
          });
        });
        return bad;
      });
      // collapsed 防御（真机四连修②）：无锚定且选区为空时格式只落在光标所在段，note 里说破
      const note = (anchorMode === 'selection' && range.isEmpty) ? collapsedAnchorNote('') : undefined;
      return { success: true, target, anchorMode, ...(note ? { note } : {}), ...verification, _navTarget: { kind: 'selection' } };
    });
  }

  async function insertHeading(args) {
    requireOffice();
    const { text = '', level = 1, location = 'End' } = args;
    return Word.run(async context => {
      const body = context.document.body;
      const expectedStyle = `Heading${Math.min(Math.max(Number(level) || 1, 1), 6)}`;
      const range = body.insertParagraph(text, location === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end);
      range.styleBuiltIn = expectedStyle;
      followSelect(range);
      await context.sync();
      // 写后读回（W2）：换代理读回段文本与标题样式
      const verification = await verifyWriteBack(async () => {
        const fresh = freshRange(range);
        fresh.load('text,styleBuiltIn');
        await context.sync();
        const bad = [];
        if (String(fresh.text || '').replace(/\r\n?/g, '\n') !== String(text).replace(/\r\n?/g, '\n')) bad.push(`文本读回为 ${clampText(JSON.stringify(String(fresh.text)), 60)}，与写入的 ${clampText(JSON.stringify(String(text)), 60)} 不符`);
        if (String(fresh.styleBuiltIn || '') !== expectedStyle) bad.push(`样式读回为「${fresh.styleBuiltIn}」，预期「${expectedStyle}」`);
        return bad;
      });
      return { success: true, text: clampText(text, 120), level, ...verification, _navTarget: { kind: 'selection' } };
    });
  }

  async function insertTable(args) {
    requireOffice();
    // 参数别名（真机反馈：模型把 rows 传成 table 撞校验）：rows 缺失而 table 是 2D 数组时自动采用，
    // 结果里带 note 说明；两者都缺时报错文案列出实际收到的键名，方便模型自我修正。
    let rows = args.rows;
    let aliasNote;
    if (rows === undefined && args.table !== undefined) {
      if (Array.isArray(args.table) && args.table.length && args.table.every(Array.isArray)) {
        rows = args.table;
        aliasNote = 'used table as rows';
      }
    }
    if (!Array.isArray(rows) || !rows.length || !rows.every(Array.isArray)) {
      const receivedKeys = Object.keys(args || {}).filter(k => k !== 'explanation');
      throw new Error(`rows must be a non-empty 2D array (received keys: ${receivedKeys.length ? receivedKeys.join(', ') : '(none)'}). Tip: a 2D array passed as "table" is accepted as an alias.`);
    }
    const { location = 'End', headerRow = true } = args;
    const rowCount = rows.length;
    // P2-1(CODE-REVIEW-ACTION-PLAN):列宽取所有行的最大列数,短行补 ''——不截断,不丢数据。
    const colCount = Math.max(...rows.map(r => r.length));
    if (!colCount) throw new Error('every row is empty: at least one row must contain at least one cell (column width is the maximum row length)');
    // 归一(真机反馈③:参差不齐的 rows 真机 insertTable 直接抛错,模型退化成 ASCII 假表格):
    // 短行补 '' 到最大列宽;单元格一律 String 强转(null/undefined → '')。
    const values = rows.map(r => Array.from({ length: colCount }, (_, i) => (r[i] == null ? '' : String(r[i]))));
    // 归一发生了补行要上报,模型才能告知用户
    const paddedRows = rows.filter(r => r.length < colCount).length;
    const normalized = paddedRows ? { paddedRows } : {};
    return Word.run(async context => {
      const body = context.document.body;
      const table = body.insertTable(rowCount, colCount, location === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end, values);
      if (headerRow && rowCount > 0) {
        // 首行作为表头：加粗。不用 styleFirstRow，因其依赖表样式是否定义了表头格式，跨版本不可靠。
        try { table.getCell(0, 0).parentRow.font.bold = true; } catch {}
      }
      followSelect(table);
      await context.sync();
      // 写后读回（W2）：比对行数与首格文本（values[0][0]，rowCount/getCell 读回均非本地赋值代理）
      const verification = await verifyWriteBack(async () => {
        table.load('rowCount');
        const firstCell = table.getCell(0, 0);
        firstCell.load('value');
        await context.sync();
        const bad = [];
        if (Number(table.rowCount) !== rowCount) bad.push(`行数读回为 ${table.rowCount}，预期 ${rowCount}`);
        const expectedFirst = values[0][0];
        if (cellText(firstCell.value) !== expectedFirst) bad.push(`首格文本读回为 ${clampText(JSON.stringify(cellText(firstCell.value)), 60)}，预期 ${clampText(JSON.stringify(expectedFirst), 60)}`);
        return bad;
      });
      return { success: true, rowCount, columnCount: colCount, ...(aliasNote ? { note: aliasNote } : {}), ...normalized, ...verification, _navTarget: { kind: 'selection' } };
    });
  }

  // fill_table_cells（W3 提前批，WORD-BLUEPRINT 4.4）：按 (tableIndex, row, column) 定位单元格
  // 写入 text，只动 cell 文本不动表格结构/列宽；一次可写多格。写后读回验证（W2 模式）。
  // confirm 闸门（Excel 批次 22 哲学）：空单元格直通，覆盖已有内容在 confirm 下拦，
  // 引导走 propose_edits 提案卡（W3 提前批起提案路径已支持表格锚定，卡能正常出）。
  async function fillTableCells(args) {
    requireOffice();
    const { tableIndex, updates } = args;
    if (!Number.isInteger(Number(tableIndex)) || Number(tableIndex) < 0) throw new Error('tableIndex must be a non-negative integer (index from get_tables)');
    const list = Array.isArray(updates) ? updates : [];
    if (!list.length) throw new Error('updates must be a non-empty array of { row, column, text }');
    for (const u of list) {
      if (!Number.isInteger(Number(u && u.row)) || !Number.isInteger(Number(u && u.column))) {
        throw new Error('each update must have integer row and column (0-based, counting from the first row/column of the table)');
      }
    }
    // 复核 P3-3：同一坐标重复出现时后者覆盖前者并上报——不去重的话读回比对会对先写条目报「疑似未生效」假警
    const byCoord = new Map();
    for (const u of list) byCoord.set(`${Number(u.row)},${Number(u.column)}`, u);
    const deduped = Array.from(byCoord.values());
    const dedupeNote = deduped.length < list.length ? `同坐标重复 ${list.length - deduped.length} 处，已按后者为准` : null;
    const accessMode = gateMode();
    return Word.run(async context => {
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      const table = tables.items[Number(tableIndex)];
      if (!table) throw new Error(`Table ${tableIndex} not found (document has ${tables.items.length} tables; valid index 0-${Math.max(0, tables.items.length - 1)})`);
      let cells;
      try {
        cells = deduped.map(u => {
          const cell = table.getCell(Number(u.row), Number(u.column));
          cell.load('value');
          return { u, cell };
        });
        await context.sync();
      } catch (e) {
        // 复核 P3-4：任何 sync 抛错都会被包到这里(不限于越界),措辞放宽避免误归因
        throw new Error(`读取表格 ${tableIndex} 的单元格失败（可能是坐标越界，先用 get_tables 确认表格尺寸）: ${String(e && e.message || e)}`);
      }
      const blockedTargets = cells.filter(({ cell }) => cellText(cell.value).trim());
      if (accessMode === 'confirm' && blockedTargets.length) {
        const where = blockedTargets.map(({ u }) => `(${u.row},${u.column})`).join(' ');
        return { success: false, blocked: true, error: `以下单元格已有内容，确认模式下不允许直接覆盖：${where}。请改用 propose_edits 生成修改卡片（find 填该单元格的现有原文，提案路径已支持表格锚定），由用户确认后应用。空单元格的写入不受限。` };
      }
      cells.forEach(({ u, cell }) => { cell.value = String(u.text == null ? '' : u.text); });
      await context.sync();
      // 写后读回（W2）：换代理逐格重新取回 value 比对（原代理可能只回本地赋值缓存）
      const verification = await verifyWriteBack(async () => {
        const fresh = cells.map(({ u }) => {
          const c = table.getCell(Number(u.row), Number(u.column));
          c.load('value');
          return { u, c };
        });
        await context.sync();
        const bad = [];
        fresh.forEach(({ u, c }, i) => {
          const expected = String(u.text == null ? '' : u.text);
          const got = cellText(c.value);
          if (got !== expected) bad.push(`第 ${i + 1} 格 (${u.row},${u.column}) 读回为 ${clampText(JSON.stringify(got), 60)}，预期 ${clampText(JSON.stringify(expected), 60)}`);
        });
        return bad;
      });
      return { success: true, tableIndex: Number(tableIndex), updated: cells.length, ...(dedupeNote ? { note: dedupeNote } : {}), ...verification, _navTarget: { kind: 'selection' } };
    });
  }

  async function insertPageBreak(args = {}) {
    requireOffice();
    // 默认 selection：用户说「插到这里/光标处」时模型常不传 location，默认落光标处才符合直觉；
    // 明确要文末/文首时模型传 End/Start
    const { location = 'selection' } = args;
    return Word.run(async context => {
      const body = context.document.body;
      const range = location === 'selection' ? context.document.getSelection() : body.getRange(location === 'Start' ? 'Start' : 'End');
      range.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
      followSelect(range);
      await context.sync();
      return { success: true, _navTarget: { kind: 'selection' } };
    });
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
    const { location = 'End' } = args;
    const base64 = await resolveImageSource(args);
    return Word.run(async context => {
      const body = context.document.body;
      const range = location === 'selection' ? context.document.getSelection() : body;
      const pic = range.insertInlinePictureFromBase64(base64, location === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end);
      followSelect(pic);
      await context.sync();
      return { success: true, _navTarget: { kind: 'selection' } };
    });
  }

  // ---- 图片管理（IMG-4，真机三轮反馈：插入重复后无法删除/调整）----
  // vendored word-mac-16.00.js 取证：InlinePicture 标量 width/height 可读写（单位磅），
  // 方法含 delete() / getRange() / getNext()；Body.inlinePictures → InlinePictureCollection（items）。
  // delete 按操作级闸门拦（对齐 manage_comment delete / Excel 删行列先例：confirm 拦引导切 auto），
  // resize 是格式类直通（工具在 directFormattingTools，delete 分支单独拦）。
  async function manageImage(args = {}) {
    requireOffice();
    const action = String(args.action || 'list');
    const index = Number(args.index);
    if (!['list', 'resize', 'delete'].includes(action)) throw new Error(`action must be list/resize/delete (got: ${action})`);
    return Word.run(async context => {
      const pics = context.document.body.inlinePictures;
      pics.load('items');
      await context.sync();
      if (action === 'list') {
        const items = pics.items || [];
        items.forEach(p => p.load('width,height,imageFormat,altTextDescription'));
        await context.sync();
        return {
          success: true, count: items.length,
          images: items.map((p, i) => ({ index: i, width: p.width, height: p.height, format: p.imageFormat, altText: p.altTextDescription || '' })),
          note: 'index 是文档内图片序号（按读取顺序，0 起）。resize 用 index+width/height（磅）；delete 用 index（确认模式下被拦，需切「直接修改」模式）。'
        };
      }
      if (!Number.isInteger(index) || index < 0 || !(pics.items && pics.items[index])) {
        throw new Error(`index ${args.index} 无效。先用 action:"list" 取当前图片序号（共 ${pics.items ? pics.items.length : 0} 张）。注意插入/删除后序号会变。`);
      }
      const pic = pics.items[index];
      if (action === 'resize') {
        const width = args.width != null ? Number(args.width) : null;
        const height = args.height != null ? Number(args.height) : null;
        if (width == null && height == null) throw new Error('resize 需要 width 和/或 height（单位磅，1 磅 ≈ 1/72 英寸；只给一项时另一项按当前比例推算后同时传入最稳）。');
        if (width != null && (!(width > 0) || width > 1584)) throw new Error(`width ${args.width} 超出合理范围（0-1584 磅，A4 纸宽）`);
        if (height != null && (!(height > 0) || height > 1584)) throw new Error(`height ${args.height} 超出合理范围（0-1584 磅）`);
        pic.load('width,height');
        await context.sync();
        const before = { width: pic.width, height: pic.height };
        // 等比推算：只给一边时按原比例补另一边（宿主 lockAspectRatio 不总是锁定，程序自己算最稳）
        const ratio = before.width && before.height ? before.height / before.width : 0.75;
        if (width != null && height == null) { pic.width = width; pic.height = Math.round(width * ratio); }
        else if (height != null && width == null) { pic.height = height; pic.width = Math.round(height / ratio); }
        else { pic.width = width; pic.height = height; }
        await context.sync();
        // 写后读回（W2 模式：不符则如实报，不静默）
        pic.load('width,height');
        await context.sync();
        const verified = Math.abs(pic.width - (width != null ? width : pic.width)) <= 0.5
          && Math.abs(pic.height - (height != null ? height : pic.height)) <= 0.5;
        return Object.assign({
          success: true, index, before, after: { width: pic.width, height: pic.height },
          verified,
          verifyWarning: verified ? undefined : '疑似未生效：读回尺寸与请求不符，请用 list 复核。'
        }, { note: '尺寸单位磅。若比例不符预期，再次 resize 微调即可。' });
      }
      // delete
      const accessMode = gateMode();
      if (accessMode === 'confirm') {
        return { success: false, blocked: true, error: '删除图片在确认模式下不可用。请告诉用户：需要删除图片时切换到「直接修改」模式后重试，或手动选中图片按 Delete。' };
      }
      const before = pics.items.length;
      pic.delete();
      await context.sync();
      pics.load('items');
      await context.sync();
      const after = pics.items ? pics.items.length : 0;
      return { success: true, index, before, after, verified: after === before - 1, note: after === before - 1 ? '已删除，后续图片序号前移。' : '已执行删除，请用 list 复核。' };
    });
  }

  // 批注全套（W1）：API 面以 vendored word-mac-16.00.js 为准——Comment{ id, content(可写),
  // authorName, creationDate, resolved(可写), replies(CommentReplyCollection), reply(text), delete(),
  // getRange() }；枚举走 body.getComments()（WordApiOnline 1.1，本机 office.js 已含）。
  async function manageComment(args) {
    requireOffice();
    const { operation = 'add', text = '', commentId } = args;
    // confirm 闸门（踩坑 11：按操作分不按工具分）：删除批注有破坏性（可能删掉用户手写的批注），
    // 与 Excel 删批注同待遇——工具本身在 confirm 下直通，delete 分支操作级拦截。
    const accessMode = gateMode();
    if (accessMode === 'confirm' && operation === 'delete') {
      return { success: false, blocked: true, error: '删除批注在确认模式下不可用。请告诉用户：需要删除批注时切换到「直接修改」模式后重试。' };
    }
    return Word.run(async context => {
      if (operation === 'add') {
        if (!String(text).trim()) throw new Error('text is required for add');
        // 锚定三态走共用的 resolveAnchorRange（与 applyStyle/setParagraphFormat 同一套）。
        // insertComment 直接挂在 range 上，无需 select()。
        const { anchorRange, anchorMode } = await resolveAnchorRange(context, args);
        if (typeof anchorRange.insertComment !== 'function') throw new Error('Comment API not available in this Word host');
        const comment = anchorRange.insertComment(String(text));
        comment.load('id');
        await context.sync();
        return { success: true, operation, commentId: comment.id, anchorMode, _navTarget: { kind: 'selection' } };
      }
      if (!['reply', 'resolve', 'delete'].includes(operation)) throw new Error(`Unsupported comment operation: ${operation}`);
      if (!commentId) throw new Error(`commentId is required for ${operation}`);
      if (operation === 'reply' && !String(text).trim()) throw new Error('text is required for reply');
      const comments = context.document.body.getComments();
      comments.load('items');
      await context.sync();
      comments.items.forEach(c => c.load('id'));
      await context.sync();
      const target = comments.items.find(c => String(c.id) === String(commentId));
      if (!target) throw new Error(`Comment '${commentId}' not found`);
      if (operation === 'reply') target.reply(String(text));
      if (operation === 'resolve') target.resolved = true;
      if (operation === 'delete') target.delete();
      await context.sync();
      return { success: true, operation, commentId: String(commentId), _navTarget: { kind: 'selection' } };
    });
  }

  // 只读：列出全部批注（作者、日期、锚定文本、回复线程、resolved 状态），review 类任务的基础设施。
  // 注意每个代理对象的属性只在自己 load 过之后才可读——getRange() 每次调用都是新代理，
  // 必须先把引用存下来再 load 再读（真机教训，夹具按此行为复刻）。
  async function getComments() {
    requireOffice();
    return Word.run(async context => {
      const comments = context.document.body.getComments();
      comments.load('items');
      await context.sync();
      const rows = comments.items.map(c => {
        c.load('id,authorName,content,creationDate,resolved');
        let anchor = null;
        try { anchor = c.getRange(); anchor.load('text'); } catch { anchor = null; }
        let replies = null;
        try { replies = c.replies; replies.load('items'); } catch { replies = null; }
        return { c, anchor, replies };
      });
      await context.sync();
      for (const row of rows) {
        if (!row.replies) continue;
        try { (row.replies.items || []).forEach(r => r.load('id,authorName,content,creationDate')); } catch {}
      }
      await context.sync();
      const list = rows.map(({ c, anchor, replies }) => {
        const replyItems = ((replies && replies.items) || []).map(r => ({ id: r.id, author: r.authorName || '', date: r.creationDate || null, content: r.content || '' }));
        return {
          id: c.id,
          refId: `c:${c.id}`,
          author: c.authorName || '',
          date: c.creationDate || null,
          content: c.content || '',
          resolved: !!c.resolved,
          anchor: anchor ? clampText(anchor.text || '', 120) : '',
          replyCount: replyItems.length,
          replies: replyItems
        };
      });
      // 计数语义显式化（真机反馈：模型曾凭对话记忆直接答批注数而答错）：
      // totalComments 只计父批注；回复嵌套在线程内（replyCount），绝不作为顶层条目。
      return { success: true, totalComments: list.length, count: list.length, comments: list, note: `共 ${list.length} 条父批注（totalComments 只计父批注）；回复嵌套在各批注的 replies 内、由 replyCount 计数，不作为独立批注计入总数。` };
    });
  }

  // ---- get_tracked_changes（修订只读读取，WORD-TRACKED-CHANGES-IMPLEMENTATION 4.1/4.2/5）----
  // 只读：绝不调用 accept/reject/acceptAll/rejectAll，绝不切换修订模式（夹具宿主根本不提供这些
  // 方法，实现也不得经 eval 绕路）。两条路径选一为主来源，不做双份相加：
  //   1) 原生 Body/Range.getTrackedChanges()（WordApi 1.6，vendored 16.0.20416 grep 取证）——
  //      官方面，author/date/text/type；缺位置信息，插入条目 best-effort 用正文唯一命中补上下文。
  //   2) 原生不可用（typeof 缺失或 sync 抛门控错）→ Range.getOoxml() 读回正文 OOXML，
  //      交 App.parseTrackedChangesOoxml 解析 w:ins/w:del（含被删文字、上下文、表格位置）。
  // typeof 只作初筛；真机 sync 时才验证 1.6 门控，失败属当下这一次的失败，探测缓存不写死。
  function shortHash(s) {
    let h = 5381;
    for (let i = 0; i < s.length; i++) h = (((h << 5) + h) + s.charCodeAt(i)) >>> 0;
    return h.toString(36);
  }
  function changesFingerprint(changes) {
    return shortHash(changes.map(c => [c.kind, c.author || '', c.date || '', String(c.text || '').slice(0, 60)].join('|')).join('§'));
  }

  function parseCursor(raw) {
    if (raw == null || raw === '') return { offset: 0 };
    let payload;
    try { payload = JSON.parse(String(raw)); } catch { throw new Error('cursor 无效（不是本工具签发的分页令牌）。请不带 cursor 重新读取。'); }
    if (!payload || typeof payload.o !== 'number' || payload.o < 0 || typeof payload.f !== 'string') {
      throw new Error('cursor 无效（缺少 fingerprint/offset）。请不带 cursor 重新读取。');
    }
    return payload;
  }

  // 原生路径读取。单条 load/读字段失败记入 failures（partial），不拖垮整批。
  async function readNativeTrackedChanges(context, range) {
    const report = { changes: [], failures: [] };
    const collection = range.getTrackedChanges();
    collection.load('items');
    await context.sync();
    const items = (collection.items || []).slice();
    const rows = [];
    for (const item of items) {
      try {
        item.load('author,date,text,type');
        rows.push({ item });
      } catch (e) {
        rows.push({ item: null, failure: item });
        report.failures.push(`第 ${rows.length} 条修订字段读取失败：${String((e && e.message) || e)}`);
      }
    }
    await context.sync();
    rows.forEach((row, index) => {
      if (!row.item) return;
      try {
        const type = String(row.item.type || '');
        report.changes.push({
          index,
          refId: `tc:${index}`,
          kind: type === 'Insert' ? 'insert' : type === 'Delete' ? 'delete' : type.toLowerCase() || 'unknown',
          author: row.item.author || null,
          date: row.item.date || null,
          revisionId: null, // 原生 TrackedChange 无持久 id（官方标量只有 author/date/text/type），不编造
          text: String(row.item.text || ''),
          contextBefore: null,
          contextAfter: null,
          location: null
        });
      } catch (e) {
        report.failures.push(`第 ${index + 1} 条修订字段读取失败：${String((e && e.message) || e)}`);
      }
    });
    return report;
  }

  // 原生路径没有位置信息：插入条目用正文唯一命中补上下文（best-effort，只做前 20 条，
  // 删除文字不在正文里、search 不到，如实留空）。失败静默跳过，由 limitations 说明。
  async function attachInsertContexts(context, changes) {
    const targets = changes.filter(c => c.kind === 'insert').slice(0, 20);
    for (const c of targets) {
      const needle = String(c.text || '').replace(/\s+/g, ' ').trim().slice(0, 80);
      if (!needle) continue;
      try {
        const hits = context.document.body.search(needle, { matchCase: false, matchWholeWord: false, matchWildcards: false });
        hits.load('items');
        await context.sync();
        if (!hits.items || hits.items.length !== 1) continue;
        const para = hits.items[0].paragraphs.getFirst();
        para.load('text');
        await context.sync();
        const paraText = String(para.text || '');
        const at = paraText.indexOf(needle);
        if (at < 0) continue;
        const radius = 42;
        const beforeStart = Math.max(0, at - radius);
        const afterEnd = Math.min(paraText.length, at + needle.length + radius);
        c.contextBefore = (beforeStart > 0 ? '…' : '') + paraText.slice(beforeStart, at);
        c.contextAfter = paraText.slice(at + needle.length, afterEnd) + (afterEnd < paraText.length ? '…' : '');
      } catch (e) { /* 上下文是锦上添花，失败不升格为工具错误 */ }
    }
  }

  async function readOoxmlTrackedChanges(context, range) {
    const result = range.getOoxml();
    await context.sync();
    const xml = result && result.value;
    if (typeof xml !== 'string' || !xml.trim()) throw new Error('getOoxml 读回为空，无法解析修订');
    const parsed = App.parseTrackedChangesOoxml(xml, { parseXml: t => new DOMParser().parseFromString(t, 'application/xml') });
    return parsed;
  }

  // 相邻删除+插入、作者相同时给候选替换配对：只作提示，原始两条记录永远分开保留。
  // 相邻与同作者都不足以证明一一对应（MD 第 5 节），置信度只给 low/medium。
  function suggestReplacementPairs(changes) {
    const pairs = [];
    for (let i = 1; i < changes.length; i++) {
      const prev = changes[i - 1], cur = changes[i];
      if (prev.kind !== 'delete' || cur.kind !== 'insert') continue;
      const sameAuthor = prev.author && prev.author === cur.author;
      const sameDate = prev.date && prev.date === cur.date;
      pairs.push({
        deleteRefId: prev.refId,
        insertRefId: cur.refId,
        confidence: sameAuthor && sameDate ? 'medium' : 'low',
        note: '候选替换配对（相邻删除+插入）。作者/日期相同仅提高可疑度，不构成一一对应证明；无法配对时两条分别展示。'
      });
    }
    return pairs;
  }

  function trackedChangesNote(status, source, opts) {
    const parts = [];
    if (status === 'ok' && opts.total === 0) {
      parts.push('当前范围没有尚存的修订记录。这只说明现在读不到修订，不能证明历史上没人改过——已接受/拒绝的旧修订通常无法从当前记录恢复。');
    }
    if (source === 'ooxml') parts.push('数据来自正文 OOXML 解析（原生 TrackedChange API 不可用）；refId tc:N 是本次快照的局部引用，重读文档后可能指向不同条目。');
    if (source === 'officejs') parts.push('数据来自原生 TrackedChange API；该源不提供段落/表格定位，插入条目的上下文为 best-effort（唯一命中才填）；refId tc:N 是本次快照的局部引用。');
    if (opts.hasMore) parts.push(`尚有未返回条目：带 nextCursor 继续读取。用户要「所有修改」时必须翻页到 hasMore=false，不得把部分结果当全部。`);
    if (opts.cursorInvalid) parts.push('文档在两次读取之间发生了变化，旧游标已失效——请丢弃之前的列表从头重读，不要拼接两个版本的修订列表。');
    if (opts.anyTruncated) parts.push('有超长文本被截断（truncated:true）；需要该条完整文字时，把其 refId 传入 expand 参数重新读取。');
    if (opts.pairs) parts.push('suggestedPairs 只是候选替换提示：相邻删除+插入，配对可信度有限，展示时保留原始两条记录。');
    return parts.join(' ');
  }

  async function getTrackedChanges(args = {}) {
    requireOffice();
    const scope = args.scope === 'selection' ? 'selection' : 'body';
    const limit = Math.min(Math.max(1, Number(args.limit || 50) || 50), 200);
    const maxTextLength = 2000;

    return Word.run(async context => {
      const range = scope === 'body' ? context.document.body : context.document.getSelection();
      let report = null;
      let source = null;
      let nativeError = null;

      if (range && typeof range.getTrackedChanges === 'function') {
        try {
          report = await readNativeTrackedChanges(context, range);
          source = 'officejs';
        } catch (e) {
          nativeError = String((e && e.message) || e); // 初筛通过但 sync 失败（如宿主不支持 1.6 门控）——降级 OOXML
          report = null;
        }
      }
      if (!report && range && typeof range.getOoxml === 'function') {
        try {
          const parsed = await readOoxmlTrackedChanges(context, range);
          report = {
            changes: parsed.changes,
            failures: [],
            unsupportedKinds: parsed.unsupportedKinds,
            warnings: parsed.warnings,
            paragraphCount: parsed.paragraphCount,
            tableCount: parsed.tableCount
          };
          source = 'ooxml';
        } catch (e) {
          if (nativeError) {
            return {
              status: 'failed', source: null, scope, complete: false, total: null, returned: 0, hasMore: false,
              changes: [], limitations: ['页眉、页脚和脚注未读取'],
              failures: [`原生 TrackedChange API 失败：${nativeError}`, `OOXML 路径失败：${String((e && e.message) || e)}`],
              note: '两条只读路径都失败。不要推测修订内容，也不要用零条记录冒充「没有修订」。可请用户提供保留了修订记录的文档副本。'
            };
          }
          throw e;
        }
      }
      if (!report) {
        // 原生真的试过且失败（门控错）而 OOXML 面不存在 → failed（试过但失败）；
        // 两个面都不存在（typeof 就没有）→ unsupported（根本没试）。两者都不得冒充零条。
        if (nativeError) {
          return {
            status: 'failed', source: null, scope, complete: false, total: null, returned: 0, hasMore: false,
            changes: [], limitations: ['页眉、页脚和脚注未读取'],
            failures: [`原生 TrackedChange API 失败：${nativeError}`, '本宿主没有可用的 OOXML 读取路径（getOoxml 不可用），无法兜底'],
            note: '修订读取失败。不要推测修订内容，也不要用零条记录冒充「没有修订」。可请用户提供保留了修订记录的文档副本。'
          };
        }
        return {
          status: 'unsupported', source: null, scope, complete: false, total: null, returned: 0, hasMore: false,
          changes: [], limitations: ['页眉、页脚和脚注未读取'],
          failures: [],
          note: '当前 Word 宿主既没有原生 TrackedChange API（WordApi 1.6），也不支持读取 OOXML。不要推测修订内容；明确告知用户该范围不可读，可请用户提供保留了修订记录的副本。'
        };
      }

      const expandRef = args.expand != null ? String(args.expand) : null;
      let entries = report.changes;
      // 统一截断口径；expand 指定的条目给完整文字（续读入口）
      for (const c of entries) {
        if (expandRef === c.refId) continue;
        if (String(c.text || '').length > maxTextLength) {
          c.text = String(c.text).slice(0, maxTextLength);
          c.truncated = true;
        }
      }
      if (expandRef) {
        const target = entries.find(c => c.refId === expandRef);
        if (!target) {
          return { status: 'failed', source, scope, complete: false, total: null, returned: 0, hasMore: false, changes: [], limitations: ['页眉、页脚和脚注未读取'], failures: [`expand 引用 ${expandRef} 不在本次读取结果中（refId 是快照局部引用，请先重新读取获取最新列表）`], note: '请先不带 expand 重新读取，再用返回列表里的 refId。' };
        }
        const fingerprint = changesFingerprint(entries);
        return {
          status: 'ok', source, scope, complete: true, total: entries.length, returned: 1, hasMore: false,
          snapshotId: `${fingerprint}-${Date.now().toString(36)}`, fingerprint,
          changes: [target], limitations: ['页眉、页脚和脚注未读取'],
          note: 'expand 返回的是该条修订在当前文档状态下的完整文字；若文档已改动，请整表重读。'
        };
      }

      // 原生路径：插入条目 best-effort 补上下文（全量算完再分页，保证跨页上下文一致）
      if (source === 'officejs') await attachInsertContexts(context, entries);

      const fingerprint = changesFingerprint(entries);
      const total = entries.length;
      const cursor = parseCursor(args.cursor);
      let cursorInvalid = false;
      let offset = Number(cursor.o) || 0;
      if (cursor.f != null && cursor.f !== fingerprint) {
        cursorInvalid = true;
        offset = 0;
      }
      if (offset > total) offset = total;
      const slice = entries.slice(offset, offset + limit);
      const hasMore = offset + limit < total;
      const anyTruncated = entries.some(c => c.truncated);
      const hasUnsupported = (report.unsupportedKinds || []).length > 0;
      const hasFailures = (report.failures || []).length > 0;
      const status = (hasUnsupported || hasFailures) ? 'partial' : 'ok';
      const complete = !hasMore && !hasUnsupported && !hasFailures && !anyTruncated;

      const result = {
        status, source, scope, complete, total, returned: slice.length, hasMore,
        snapshotId: `${fingerprint}-${Date.now().toString(36)}`, fingerprint,
        changes: slice,
        limitations: ['页眉、页脚和脚注未读取']
      };
      if (hasMore) result.nextCursor = JSON.stringify({ f: fingerprint, o: offset + limit });
      if (cursorInvalid) result.cursorInvalid = true;
      // 原生初筛通过但真调用失败、靠 OOXML 兜底成功时，把原生失败原因如实带出（不冒充一切正常）
      if (nativeError) result.nativeFallbackError = nativeError;
      if (hasUnsupported) result.unsupportedKinds = report.unsupportedKinds;
      if (hasFailures) result.failures = report.failures;
      if (report.warnings && report.warnings.length) result.warnings = report.warnings;
      const pairs = suggestReplacementPairs(entries);
      if (pairs.length) result.suggestedPairs = pairs;
      result.note = trackedChangesNote(status, source, { total, hasMore, cursorInvalid, anyTruncated, pairs: pairs.length > 0 });
      return result;
    });
  }

  // 内容控件作为可导航锚点：add 时返回 tag，navigateCitation 时按 tag/title 选中
  async function manageContentControl(args) {
    requireOffice();
    const { operation = 'add', tag, title, text } = args;
    return Word.run(async context => {
      if (operation === 'add') {
        const sel = context.document.getSelection();
        const cc = sel.insertContentControl();
        if (tag) cc.tag = tag;
        if (title) cc.title = title;
        if (text) cc.insertText(text, Word.InsertLocation.replace);
        cc.load('id,tag,title');
        await context.sync();
        return { success: true, operation, id: cc.id, tag: cc.tag, title: cc.title, _navTarget: { kind: 'contentControl', tag: cc.tag || tag } };
      }
      if (operation === 'list') {
        const ccs = context.document.contentControls;
        ccs.load('items');
        await context.sync();
        ccs.items.forEach(c => c.load('id,tag,title,text'));
        await context.sync();
        return { success: true, controls: ccs.items.map(c => ({ id: c.id, tag: c.tag, title: c.title, text: clampText(c.text, 120) })) };
      }
      throw new Error(`Unsupported content control operation: ${operation}`);
    });
  }

  async function evalOfficeJs(args) {
    requireOffice();
    const code = args.code || '';
    return Word.run(async context => {
      const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
      const fn = new AsyncFunction('context', 'Word', code);
      const result = await fn(context, Word);
      return { success: true, result: result ?? null };
    });
  }

  async function selectByParagraphIndex(index) {
    // 引用加固（W2）：p: 空串会被 Number('') 变成 0 静默跳首段；非整数先拦下，不许乱猜
    const raw = String(index == null ? '' : index).trim();
    if (!/^\d+$/.test(raw)) throw new Error(`Invalid paragraph reference: '${index}' (expected a non-negative integer)`);
    return Word.run(async context => {
      const paras = context.document.body.paragraphs;
      paras.load('items');
      await context.sync();
      const p = paras.items[Number(raw)];
      if (!p) throw new Error(`Paragraph ${raw} not found (document has ${paras.items.length} paragraphs)`);
      p.select();
      await context.sync();
      return { success: true };
    });
  }
  async function selectByContentControlTag(tag) {
    // 引用加固（W2）：空 tag 会命中第一个无 tag 的控件，必须先拦
    const raw = String(tag == null ? '' : tag).trim();
    if (!raw) throw new Error('Invalid content control reference: empty tag');
    return Word.run(async context => {
      const ccs = context.document.contentControls;
      ccs.load('items');
      await context.sync();
      ccs.items.forEach(c => c.load('tag'));
      await context.sync();
      const cc = ccs.items.find(c => c.tag === raw);
      if (!cc) throw new Error(`Content control '${raw}' not found`);
      cc.select();
      await context.sync();
      return { success: true };
    });
  }

  async function selectByCommentId(commentId) {
    const raw = String(commentId == null ? '' : commentId).trim();
    if (!raw) throw new Error('Invalid comment reference: empty id');
    return Word.run(async context => {
      const comments = context.document.body.getComments();
      comments.load('items');
      await context.sync();
      comments.items.forEach(c => c.load('id'));
      await context.sync();
      // 与 manageComment 同一匹配方式：String 化后按 id 精确比对
      const target = comments.items.find(c => String(c.id) === raw);
      if (!target) throw new Error(`Comment '${raw}' not found`);
      // Comment.getRange 无参（vendored 16.0.20416 grep 取证：经典 Comment 类 delete 后紧跟
      // getRange=function()）；命中后选中锚定区间即完成跳转
      target.getRange().select();
      await context.sync();
      return { success: true };
    });
  }

  async function navigateCitation(ref) {
    requireOffice();
    // 格式: "p:<index>" 段落序号 | "cc:<tag>" 内容控件 | "c:<commentId>" 批注 | 纯数字回退为段落序号
    if (/^cc:/.test(ref)) return selectByContentControlTag(ref.slice(3));
    if (/^c:/.test(ref)) return selectByCommentId(ref.slice(2));
    if (/^p:/.test(ref)) return selectByParagraphIndex(ref.slice(2));
    return selectByParagraphIndex(ref);
  }

  async function maybeFollow(result) {
    if (!App.state.settings.followMode || !result) return;
    const nav = result._navTarget;
    if (!nav) return;
    try {
      if (nav.kind === 'contentControl' && nav.tag) await selectByContentControlTag(nav.tag);
      // kind === 'selection'：插入类工具已对新内容调用 select()，选区即落在新内容上，无需再次定位。
    } catch (e) { console.warn(e); }
  }

  const TOOL_EXECUTORS = {
    read_memory: (args) => App.memory.readMemoryTool(args),
    write_memory: (args) => App.memory.writeMemoryTool(args),
    get_document_outline: getDocumentOutline,
    get_selection: getSelection,
    get_paragraphs: getParagraphs,
    get_tables: getTables,
    search_text: searchText,
    insert_text: insertText,
    replace_text: replaceText,
    apply_style: applyStyle,
    set_paragraph_format: setParagraphFormat,
    insert_heading: insertHeading,
    insert_table: insertTable,
    fill_table_cells: fillTableCells,
    insert_page_break: insertPageBreak,
    insert_image: insertImage,
    manage_image: manageImage,
    manage_comment: manageComment,
    get_comments: getComments,
    get_tracked_changes: getTrackedChanges,
    check_consistency: checkConsistency,
    insert_hyperlink: insertHyperlink,
    manage_content_control: manageContentControl,
    eval_officejs: evalOfficeJs
  };

  async function getDocumentMetadata() {
    requireOffice();
    return Word.run(async context => {
      const body = context.document.body;
      const paras = body.paragraphs;
      paras.load('items');
      const sel = context.document.getSelection();
      sel.load('text,styleBuiltIn');
      await context.sync();
      const headingItems = paras.items.slice(0, 300);
      headingItems.forEach(p => p.load('text,styleBuiltIn'));
      await context.sync();
      const outline = [];
      headingItems.forEach((p, idx) => {
        const style = String(p.styleBuiltIn || '');
        if (/heading/i.test(style) && (p.text || '').trim()) {
          outline.push({ index: idx, level: parseInt(style.replace(/\D/g, ''), 10) || 1, text: clampText(p.text.trim(), 120), refId: `p:${idx}` });
        }
      });
      return { success: true, documentId: App.state.workbookId || 'document', paragraphCount: paras.items.length, selection: { text: clampText(sel.text || '', 200), style: String(sel.styleBuiltIn || '') }, outline };
    });
  }

  // 全文注入：低于阈值时把整篇正文一次性交给模型，省掉 2 次读取往返。
  // 行首 [n] 是段落序号，可直接用于 propose_edits 的 target.paragraphIndex。
  const FULL_CONTEXT_CHAR_LIMIT = 15000;

  async function getFullContext() {
    requireOffice();
    return Word.run(async context => {
      const paras = context.document.body.paragraphs;
      paras.load('items');
      const sel = context.document.getSelection();
      sel.load('text');
      await context.sync();
      paras.items.forEach(p => p.load('text,styleBuiltIn'));
      await context.sync();

      const texts = paras.items.map(p => String(p.text || ''));
      const charCount = texts.reduce((sum, x) => sum + x.length, 0);
      if (charCount > FULL_CONTEXT_CHAR_LIMIT) {
        return { truncated: true, charCount, paragraphCount: paras.items.length };
      }
      const lines = texts.map((text, i) => {
        const style = String(paras.items[i].styleBuiltIn || '');
        const tag = /heading/i.test(style) ? `(${style})` : '';
        return `[${i}]${tag} ${text}`;
      });
      return {
        truncated: false,
        charCount,
        paragraphCount: paras.items.length,
        selection: String(sel.text || '').trim(),
        text: lines.join('\n')
      };
    });
  }

  const SYSTEM_PROMPT = `你是嵌在 Microsoft Word 里的中文写作助手。用户是中文母语的专业文字工作者，你的全部回复和所有写入文档的内容都用中文。

工作方式：
1. 先判断用户到底要什么。要求含糊时先问，不要猜着改。
2. 动手前先说明你的整体判断：这份文档或这段文字的问题在哪、你打算怎么处理。这段话用自然语言写在正文里，不要塞进工具参数。
3. 然后才提出具体修改。
4. 每一处改动都要能追溯到用户的要求，不要顺手改用户没提的地方。

呈现规则（决定用户看到的是卡片还是文字，很重要）：
- 当你的回答是「N 个同构的项」「几个方案并排比较」「行列矩阵」「层级大纲」「需要用户填的表单」时，调用 render 工具，界面会负责编号、对齐、跳转和操作按钮。
- 用了 render 之后，正文里不要再把同样的内容复述一遍，写一句话引出即可。
- 反过来，下面这些情况一律用正文回复，不要调 render：普通问答、解释概念、只有一两句话的回答、单个连续段落的叙述或改写、闲聊与确认。给一段话套上卡片只会让界面变吵，不会让它变清楚。
- 判断标准：需要对着比较、需要逐项操作、或者项与项之间字段相同，就用 render；只是要读一遍，就用正文。


You are an AI assistant integrated into Microsoft Word with full access to read and modify the document.

Available tools:
READ:
- get_document_outline: Read heading structure and paragraph summaries
- get_selection: Read the currently selected text and its formatting（光标未选中文字时返回所在段文本）
- get_paragraphs: Read a range of paragraphs by index
- get_tables: Read table contents
- search_text: Find text (supports match case / whole word / wildcards)
- get_comments: List all comments (author, date, anchor text, reply thread, resolved status). Each comment carries refId（c:<id>）；回答批注相关问题时用 [文字](#cite:c:<commentId>) 生成可点击跳转。totalComments 只计父批注，回复嵌套在线程内不计入总数。列出、统计或处理批注前必须调用本工具，禁止凭对话记忆作答
- get_tracked_changes: List surviving tracked changes (revisions): author, date, kind (insert/delete), exact inserted/deleted text, surrounding context. refId（tc:<index>）是本次快照的局部引用，不支持 #cite 跳转。只读：不能接受/拒绝修订。用户要「所有修改」时翻页到 hasMore=false；正文零修订只说明当前范围没有尚存修订，不能断言历史上没人改过

WRITE:
- insert_text: Insert text at Start/End of document, or Before/After the selection. 确认模式下 Replace 被拦截（覆盖选区有破坏性），替换内容请走 propose_edits 提案卡；Start/End/Before/After 纯插入直通
- replace_text: Find and replace text (replaceAll optional)
- apply_style: Apply a built-in style (e.g. Heading1, Normal, Quote) and font formatting. 用户点名位置（「第一小节标题」「某句话」）时必须用 paragraphIndex 或 find 锚定——无锚定且无选区时格式只落在光标处，字体属性只是键入属性、没有可见效果
- set_paragraph_format: Alignment / line spacing / indent. 锚定方式与 apply_style 相同（paragraphIndex / find，否则选区）。行距优先用 lineSpacingMultiple 传倍数（1.5 = 1.5 倍行距，程序自动换算成磅值）；lineSpacing 是磅值直传，别按倍数传
- insert_heading: Insert a heading paragraph at a given level
- insert_table: Insert a table from a 2D array of rows
- fill_table_cells: Write text into existing table cells by (tableIndex, row, column). 只动单元格文本，不动表格结构；空单元格直通，覆盖已有内容在确认模式下被拦、走 propose_edits（提案卡已支持锚定表格内文字）
- check_consistency: 术语/数字一致性检索底稿（只读）。传入你抽取的候选术语（terms，最多 20 个），返回每个术语的全部出现位置与 ±30 字上下文；数字自动聚合（出现 ≥2 次，千分位写法归一）。工具只检索计数，一致性判断由你做：发现问题列成带引用位置的清单，确认后用 propose_edits 改
- insert_hyperlink: 给文字设超链接（不动文字）。find 必须全文唯一命中（多命中拒绝执行，防止加错位置）；省略 find 用当前选区。需要 WordApiDesktop 1.3；不可用时退回纯文本 + 用户 Cmd+K
- insert_page_break: Insert a page break
- insert_image: 插入内联图片——任何编辑模式下都直接调用，绝不走 propose_edits（提案卡只能携带文字，包图片的卡会退化成文字重复，图丢失）。用户上传/拖拽过图片时必须用 attachmentId（消息里的 [Uploaded image: …] 块），不要自己编 base64；图片直链用 url（web_search 返回的 direct image links 可直接用）；「插到这里/光标处」必须 location:'selection'——光标位置在执行瞬间由插件读取，不需要预先读取或反复确认；「文末」用 'End'
- manage_image: 管理文档里已插入的图片：list 列序号尺寸 / resize 调大小（磅，只给一边自动等比）/ delete 删除（确认模式拦）。图片插重复或尺寸不对时用它修正，不要靠再插一张来修
- manage_comment: Add a comment (anchor: current selection by default, or a whole paragraph via paragraphIndex, or unique exact text via find), or reply to / resolve / delete a comment thread by commentId (find ids with get_comments; delete is blocked in confirm mode). 用 paragraphIndex 锚定前必须先调 get_paragraphs/get_document_outline 确认该段实际内容——段号是机械计数，标题、署名行、空行都占段号，用户口语的「第 N 段」通常只数正文段，禁止凭猜段号；能引用原文时优先用 find 锚定。
- manage_content_control: Add/list content controls (use as named, navigable anchors)
- eval_officejs: Execute Word.run code when the listed tools are not enough
- propose_edits: Present an edit proposal (diff card) for the user to review. 相互依赖的修改（对调、交换、联动改动）必须合并为一张卡——用覆盖整个受影响区域的 find/replacement 一次完成；不得拆成多张有先后顺序依赖的卡，因为应用前一张会改变后一张的原文

Citations: Use markdown links with #cite: hash to reference document locations. Clicking navigates there.
- Paragraph by index: [intro](#cite:p:0)
- Content control by tag: [summary](#cite:cc:summary)
- Comment by id: [comment](#cite:c:123)（id 见 get_comments 返回的 refId；回答批注问题时必须用它生成跳转链接）
Example: [see the introduction](#cite:p:0)


批注与修订（用户问「所有修改意见」「删了什么」「为什么改」「修订痕迹」时）：
1. 这是批注与修订的联合检查：分别调用 get_comments 和 get_tracked_changes，再结合相关正文。只读到批注时先说明「修订还未核对」，不得据此得出「没有文字修改」。用户只问批注数量时不强制扩大到修订。
2. 结果分三类写清楚：记录显示的修改事实 / 审稿人明确写出的理由 / 你根据上下文推测的原因。审稿人没写理由的条目标「未说明，以下为推测」，推测不得写成结论。
3. 来源链接只证明有人提供了参考，不等于已核实属实；引用是否支持具体事实要打开原始来源核对，未核对就标待核实。不凭几条批注给整份稿件背书。
4. 「删得对不对」按信息必要性、重复、准确性证据、语气与逻辑评估，给出保留/采纳/待核实及依据；不默认所有删除都有合理理由。
5. 修订读取失败或范围受限（unsupported/failed/partial、页眉页脚未覆盖）要如实说明，不重复调用同一批工具期待不同结果；只有 get_tracked_changes 返回 source:"ooxml" 时才能说查过底层 XML，否则不得声称读过 XML。


修改文档前先确认你读到的是完整原文，不是摘要或截断文本。`;

  const WORD_EDIT_ITEM = {
    type: 'object',
    properties: {
      label: { type: 'string', description: 'Short name for this suggestion.' },
      placement: { type: 'string', enum: ['replace', 'after', 'before'], description: 'replace swaps find; after/before inserts replacement beside the anchor.' },
      find: { type: 'string', description: 'Exact, focused original text or insertion anchor. Keep it under 180 characters when possible. 相互依赖的修改（对调、交换、联动改动）不得拆成多张卡：find/replacement 必须覆盖整个受影响区域一次完成，因为应用前一张卡会改变后一张卡的原文。' },
      replacement: { type: 'string', description: 'Exact replacement text; may be empty only for a pure deletion.' },
      contextBefore: { type: 'string', description: 'Short unchanged context immediately before find. The app verifies and fills this from the target paragraph.' },
      contextAfter: { type: 'string', description: 'Short unchanged context immediately after find. The app verifies and fills this from the target paragraph.' },
      target: { type: 'object', properties: { paragraphIndex: { type: 'number', description: 'Paragraph index returned by get_paragraphs/get_document_outline.' } } },
      reasoning: { type: 'string', description: 'Concrete reason this edit is needed.' },
      summary: { type: 'string', description: 'One-line summary shown on the card.' }
    },
    required: ['label', 'find', 'replacement', 'reasoning', 'summary']
  };

  const TOOL_DEFINITIONS = [
    { type: 'function', function: { name: 'get_document_outline', description: 'Read the document heading structure and a paragraph summary. Returns outline (headings with index/level) and paragraphs (index/text/style).', parameters: { type: 'object', properties: { maxParagraphs: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'get_selection', description: 'Read the currently selected text and its style/font. If no text is selected (cursor only), also returns the paragraph at the cursor as paragraph.text with collapsed=true.', parameters: { type: 'object', properties: { explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'get_paragraphs', description: 'Read a range of paragraphs by index.', parameters: { type: 'object', properties: { start: { type: 'number' }, count: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'get_tables', description: 'Read the contents of tables in the document.', parameters: { type: 'object', properties: { maxTables: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'search_text', description: 'Find text in the document. Supports matchCase, matchWholeWord, useWildcards.', parameters: { type: 'object', properties: { query: { type: 'string' }, matchCase: { type: 'boolean' }, matchWholeWord: { type: 'boolean' }, useWildcards: { type: 'boolean' }, maxResults: { type: 'number' }, explanation: { type: 'string' } }, required: ['query'] } } },
    { type: 'function', function: { name: 'insert_text', description: "WRITE. Insert text. location: Start/End (document) or Replace/Before/After (current selection). Optional built-in style. In confirm mode Replace is blocked (overwriting the selection is destructive) — use propose_edits for replacements; Start/End/Before/After go through directly.", parameters: { type: 'object', properties: { text: { type: 'string' }, location: { enum: ['Start', 'End', 'Replace', 'Before', 'After'] }, style: { type: 'string' }, explanation: { type: 'string' } }, required: ['text'] } } },
    { type: 'function', function: { name: 'replace_text', description: 'WRITE. Find and replace text. replaceAll defaults true.', parameters: { type: 'object', properties: { query: { type: 'string' }, replacement: { type: 'string' }, matchCase: { type: 'boolean' }, matchWholeWord: { type: 'boolean' }, useWildcards: { type: 'boolean' }, replaceAll: { type: 'boolean' }, explanation: { type: 'string' } }, required: ['query', 'replacement'] } } },
    { type: 'function', function: { name: 'apply_style', description: 'WRITE. Apply a built-in style and/or font formatting. Anchor three ways: (a) no anchor args → target (selection default, or whole document); (b) paragraphIndex → that whole paragraph (MECHANICAL index from get_paragraphs/get_document_outline — read first, never guess); (c) find → the unique document-wide match of that exact text (0 or multiple matches → error). Use (b)/(c) whenever the user names a location. Without an anchor and with an empty selection (cursor only), formatting lands on the cursor position only and font props become typing attributes with no visible effect on existing text.', parameters: { type: 'object', properties: { target: { enum: ['selection', 'document'] }, paragraphIndex: { type: 'number', description: 'Paragraph index from get_paragraphs/get_document_outline to format; wins over target.' }, find: { type: 'string', description: 'Exact anchor text; must match exactly once document-wide. Wins over target.' }, style: { type: 'string' }, font: { type: 'object', properties: { bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' }, size: { type: 'number' }, name: { type: 'string' }, color: { type: 'string' } } }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'set_paragraph_format', description: 'WRITE. Set paragraph alignment / line spacing / indent / spacing. Same anchoring as apply_style: no anchor args → target (selection default, or whole document); paragraphIndex → that whole paragraph (mechanical index, read first); find → unique document-wide match. Without an anchor and with an empty selection, formatting lands on the cursor paragraph only — pass paragraphIndex or find when the user names a location.', parameters: { type: 'object', properties: { target: { enum: ['selection', 'document'] }, paragraphIndex: { type: 'number', description: 'Paragraph index from get_paragraphs/get_document_outline to format; wins over target.' }, find: { type: 'string', description: 'Exact anchor text; must match exactly once document-wide. Wins over target.' }, alignment: { enum: ['Left', 'Centered', 'Right', 'Justified'] }, lineSpacing: { type: 'number', description: 'Line spacing in POINTS (12 = single). Do NOT pass multiples here — use lineSpacingMultiple instead.' }, lineSpacingMultiple: { type: 'number', description: 'Line spacing as a multiple, e.g. 1.5 for 1.5x. Preferred over lineSpacing; converted to points internally.' }, leftIndent: { type: 'number' }, spaceBefore: { type: 'number' }, spaceAfter: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'insert_heading', description: 'WRITE. Insert a heading paragraph at a level (1-6).', parameters: { type: 'object', properties: { text: { type: 'string' }, level: { type: 'number' }, location: { enum: ['Start', 'End'] }, explanation: { type: 'string' } }, required: ['text'] } } },
    { type: 'function', function: { name: 'insert_table', description: 'WRITE. Insert a table from a 2D array of rows. Column width is the maximum row length: shorter rows are padded with empty cells (nothing is truncated or lost); cell values are stringified (null/undefined → empty).', parameters: { type: 'object', properties: { rows: { type: 'array', items: { type: 'array', items: {} }, description: '2D array of cell values. Alias: if rows is missing, a 2D array passed as "table" is accepted.' }, location: { enum: ['Start', 'End'] }, headerRow: { type: 'boolean' }, explanation: { type: 'string' } }, required: ['rows'] } } },
    { type: 'function', function: { name: 'fill_table_cells', description: 'WRITE. Write text into existing table cells by position: tableIndex (from get_tables) + updates array of { row, column, text } (0-based). Only cell text is changed — table structure, column widths and formatting are untouched. Writes to EMPTY cells go through directly; overwriting a cell that already has content is blocked in confirm mode — use propose_edits instead (proposals support anchoring on table cell text). Values are verified by reading each cell back after the write.', parameters: { type: 'object', properties: { tableIndex: { type: 'number', description: 'Table index from get_tables.' }, updates: { type: 'array', items: { type: 'object', properties: { row: { type: 'number', description: '0-based row index.' }, column: { type: 'number', description: '0-based column index.' }, text: { type: 'string', description: 'New cell text.' } }, required: ['row', 'column', 'text'] } }, explanation: { type: 'string' } }, required: ['tableIndex', 'updates'] } } },
    { type: 'function', function: { name: 'insert_page_break', description: 'WRITE. Insert a page break.', parameters: { type: 'object', properties: { location: { enum: ['Start', 'End', 'selection'] }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'insert_image', description: 'WRITE. Insert an inline image. Preferred source: attachmentId — the id shown in the [Uploaded image: … | attachmentId: img-N] block when the user uploaded or dragged an image into the chat. Also accepts url (direct image link; downloaded through the local service) or raw base64 (no data: prefix) — never generate base64 yourself.', parameters: { type: 'object', properties: { attachmentId: { type: 'string', description: 'Id from the user uploaded image block, e.g. img-1.' }, url: { type: 'string', description: 'Direct image URL (https://…/x.png).' }, base64: { type: 'string' }, location: { enum: ['Start', 'End', 'selection'], description: "Where to insert: 'selection' = at the user's cursor / selected position (DEFAULT — use whenever the user says 在这里/光标处/this spot), 'End' = end of document (文末), 'Start' = beginning." }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'manage_image', description: 'WRITE. Manage the inline images already in the document: list (index + size + format for every picture), resize (index + width and/or height in points, aspect ratio auto-kept when only one side is given, read back verified), delete (index; blocked in confirm mode — switch to direct mode or delete manually). Call list first to get valid indexes; indexes shift after every insert/delete. Use this to fix a wrong-size or duplicated image instead of inserting another copy.', parameters: { type: 'object', properties: { action: { enum: ['list', 'resize', 'delete'] }, index: { type: 'number', description: 'Image index from action=list. 0-based.' }, width: { type: 'number', description: 'New width in points (1pt = 1/72 inch), 0-1584.' }, height: { type: 'number', description: 'New height in points, 0-1584.' }, explanation: { type: 'string' } }, required: ['action'] } } },
    { type: 'function', function: { name: 'manage_comment', description: 'WRITE. Manage comments. add (text required) anchors three ways: (a) no anchor args → current selection; (b) paragraphIndex → the whole paragraph (MECHANICAL index from get_paragraphs/get_document_outline counting EVERY paragraph incl. title/byline/blank lines — a colloquial「第N段」from the user usually means body paragraphs only, so MUST read first and verify the paragraph text matches what the user means before anchoring; out-of-range → error); (c) find → the unique document-wide match of that exact anchor text (0 matches or multiple matches → error; on multiple, give a longer unique anchor). Both paragraphIndex and find → paragraphIndex wins, find is verified inside that paragraph. reply/resolve/delete: act on an existing thread by commentId (find ids with get_comments; reply requires text; delete is blocked in confirm mode).', parameters: { type: 'object', properties: { operation: { enum: ['add', 'reply', 'resolve', 'delete'] }, commentId: { type: 'string' }, text: { type: 'string' }, paragraphIndex: { type: 'number', description: 'add only: paragraph index from get_paragraphs/get_document_outline to anchor the comment on.' }, find: { type: 'string', description: 'add only: exact anchor text; must match exactly once document-wide (or inside the paragraphIndex paragraph, as verification).' }, explanation: { type: 'string' } }, required: ['operation'] } } },
    { type: 'function', function: { name: 'get_comments', description: 'READ. List all comments in the document: author, date, anchor text, content, reply thread and resolved status. Each comment carries refId "c:<id>" — when answering comment-related questions, link to it as [text](#cite:c:<commentId>) so the user can click to jump to the comment anchor. Output includes totalComments (top-level parent comments only; replies are nested inside each thread, never counted as separate comments). MUST be called before listing, counting or processing comments — never answer comment questions from conversation memory.', parameters: { type: 'object', properties: { explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'get_tracked_changes', description: 'READ. List tracked changes (revisions) still recorded in the document: author, date, kind (insert/delete), the exact inserted/deleted text, and surrounding context. ZERO writes — cannot accept/reject revisions or toggle track-changes mode. scope: "body" (default) or "selection". Paging: pass limit + nextCursor; when the user asks for ALL changes keep paging until hasMore=false — never present a partial list as complete. status ok|partial|unsupported|failed; complete=false or unsupportedKinds/failures present means the result has gaps — say so explicitly, never report an empty list as "no edits were ever made" (accepted/rejected old revisions are unrecoverable). suggestedPairs marks an adjacent delete+insert as a CANDIDATE replacement only — always show both original records. Pass a changes[].refId via expand to fetch one entry\'s full untruncated text.', parameters: { type: 'object', properties: { scope: { enum: ['body', 'selection'], description: 'Read revisions in the whole body (default) or only the current selection.' }, limit: { type: 'number', description: 'Max entries per page (default 50, max 200).' }, cursor: { type: 'string', description: 'nextCursor from the previous page.' }, expand: { type: 'string', description: 'A refId from changes (e.g. "tc:3") to fetch that entry\'s full untruncated text.' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'check_consistency', description: 'READ. Consistency scan support for terms and numbers (document-wide, including table cell text). Pass terms: the candidate wordings you extracted (e.g. brand names, product names, abbreviations) — each is counted with every occurrence\'s location and ±30 chars of context. Numbers are scanned automatically: every value occurring 2+ times is aggregated (thousand-separator variants like "1,200" vs "1200" are grouped under one key, raw forms kept). The tool only searches and counts — YOU judge whether occurrences are inconsistent (same metric with different values, same concept with different spellings), report findings as a list with citation locations, and fix via propose_edits only after the user confirms. Zero writes.', parameters: { type: 'object', properties: { terms: { type: 'array', items: { type: 'string' }, description: 'Candidate terms to count document-wide (max 20). Omit to get the number aggregation only.' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'insert_hyperlink', description: 'WRITE. Set a hyperlink without changing any text. Two targets: a unique document-wide text match (find — must match exactly once; multiple matches are rejected to avoid linking the wrong spot) or the current selection in Word (omit find). Requires the Word Hyperlink API (WordApiDesktop 1.3); the result is verified by reading the address back. Fallback when unavailable: insert the URL as plain text (insert_text) and tell the user to select it and press Cmd+K.', parameters: { type: 'object', properties: { address: { type: 'string', description: 'Link target: https://…, mailto:…, or a file name.' }, find: { type: 'string', description: 'Exact text to hyperlink; must match exactly once document-wide (verify with search_text first).' }, screenTip: { type: 'string', description: 'Optional tooltip shown on hover.' }, explanation: { type: 'string' } }, required: ['address'] } } },
    { type: 'function', function: { name: 'manage_content_control', description: 'WRITE. Add/list content controls. Added controls act as named, navigable anchors (use tag in #cite:cc:tag).', parameters: { type: 'object', properties: { operation: { enum: ['add', 'list'] }, tag: { type: 'string' }, title: { type: 'string' }, text: { type: 'string' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'read_memory', description: 'SYSTEM. Read your cross-session memory (habits and preferences learned from past work). Call with no topic to list what is stored; pass topic (key) to read its full content. Check the list when the task touches a stored topic (review rules, formatting habits, client preferences) — apply it instead of re-deriving from scratch.', parameters: { type: 'object', properties: { topic: { type: 'string', description: 'Memory topic key. Omit to list available topics first.' } } } } },
    { type: 'function', function: { name: 'write_memory', description: 'SYSTEM. Save a stable work habit/preference to cross-session memory so future sessions inherit it (document conventions the user corrected you on, style preferences, recurring task patterns). action: create (new topic) / append (small addition, preferred) / update (reorganize one topic) / delete. Only save preferences the user explicitly expressed or repeatedly showed; never promote a single occurrence into a rule; timestamp time-sensitive facts; never store secrets or private data. The user can review and delete memory in settings. Optionally set 1-3 triggers (keywords) when the topic binds to a specific document/task type; when the latest user message matches a trigger, that topic is promoted to the top of the memory list.', parameters: { type: 'object', properties: { action: { type: 'string', enum: ['create', 'append', 'update', 'delete'], description: 'create=new topic; append=small addition (preferred); update=rewrite one topic; delete=remove.' }, topic: { type: 'string', description: 'Topic key, snake_case, e.g. doc_review_rules.' }, title: { type: 'string', description: 'Short title shown in settings (create only).' }, summary: { type: 'string', description: 'One-line summary for the memory list (create/update).' }, content: { type: 'string', description: 'The rule/preference text (create/append/update).' }, triggers: { type: 'array', items: { type: 'string' }, description: 'Keywords binding this topic to a document/task type (e.g. ["标书", "周报"]). Latest-message match pins the topic to the top of the list. Set on create/update; merged on append.' } }, required: ['action', 'topic'] } } },
    { type: 'function', function: { name: 'eval_officejs', description: 'Execute arbitrary Office.js code in Word.run. Escape hatch. Code receives context and Word.', parameters: { type: 'object', properties: { code: { type: 'string' }, explanation: { type: 'string' } }, required: ['code'] } } },
    { type: 'function', function: { name: 'propose_edits', description: 'PROPOSAL (do not write directly). For overwriting or modifying EXISTING content. Pure insertions (new headings, paragraphs, notes in empty space) must use insert_heading/insert_text instead — they apply directly and are undoable. Appending to the document start or end is ALSO a pure insertion: use insert_text with location Start/End, or placement after/before anchored on the exact first/last text — NEVER implement an append by replacing the last paragraph or line with the new sentence, because replace deletes that original text. Use top-level edits for alternative versions of one location, or top-level changes for several independently reviewable locations. Every edit item must carry a real find anchor (existing document text) and a substantive replacement — placeholder/empty items are rejected. Put the recommendation and rationale in the card before review. When the user declines with a reason, your next reply MUST act on it in the same turn — a revised proposal card that absorbs the feedback, a proposed direction for the user to confirm, or what you will do instead; never just save the feedback to memory or acknowledge it and stop.', parameters: { type: 'object', properties: { edits: { type: 'array', description: 'ONE location, optionally with alternative versions.', items: WORD_EDIT_ITEM }, changes: { type: 'array', description: 'MULTIPLE locations in one review batch. Never nest this inside edits. Items must be independent of each other — mutually dependent edits (swap, exchange, linked changes) MUST be merged into ONE card whose find/replacement covers the whole affected region, because applying an earlier card changes the original text a later card anchors on.', items: WORD_EDIT_ITEM }, explanation: { type: 'string', description: 'Short overall recommendation shown before the user reviews the cards.' } }, required: ['explanation'] } } }
  ];

  const SAMPLE_ARGS = {
    get_document_outline: { maxParagraphs: 400 },
    get_selection: {},
    get_paragraphs: { start: 0, count: 20 },
    get_tables: { maxTables: 20 },
    search_text: { query: 'keyword', matchCase: false },
    insert_text: { text: '这是新插入的一段文字。', location: 'End' },
    replace_text: { query: '旧文字', replacement: '新文字', replaceAll: true },
    apply_style: { paragraphIndex: 2, style: 'Heading1', font: { bold: true } },
    set_paragraph_format: { find: '第一小节标题原文', alignment: 'Centered', lineSpacingMultiple: 1.5 },
    insert_heading: { text: '小结', level: 1, location: 'End' },
    insert_table: { rows: [['姓名', '分数'], ['张三', 90], ['李四', 85]], location: 'End', headerRow: true },
    fill_table_cells: { tableIndex: 0, updates: [{ row: 1, column: 0, text: '王五' }, { row: 2, column: 0, text: '张三' }] },
    insert_page_break: { location: 'End' },
    insert_image: { attachmentId: 'img-1', location: 'selection' },
    manage_comment: { operation: 'add', paragraphIndex: 2, text: '这段需要补充数据来源。' },
    get_comments: {},
    get_tracked_changes: { scope: 'body', limit: 50 },
    check_consistency: { terms: ['智谱', '智谱AI'] },
    insert_hyperlink: { address: 'https://example.com', find: '点击这里' },
    manage_image: { action: 'list' },
    manage_content_control: { operation: 'add', tag: 'summary', title: '总结', text: '总结内容' },
    eval_officejs: { code: "const sel = context.document.getSelection();\nsel.load('text');\nawait context.sync();\nreturn sel.text;" },
    propose_edits: { edits: [{ label: '精简版', find: '（原文上下文…待修改句…上下文）', replacement: '（修改后的完整段落）', reasoning: '删掉了冗余的修饰，把结论提前', summary: '删冗余、结论前置' }] }
  };
  function defaultArgsForTool(name) { return App.pretty(SAMPLE_ARGS[name] || {}); }

  async function replaceTargetedWordEdit(edit, placement, token) {
    const find = String(edit.find || '');
    const replacement = String(edit.replacement == null ? '' : edit.replacement);
    const requestedIndex = Number(edit.target && edit.target.paragraphIndex);
    return Word.run(async context => {
      // 候选集 = 正文段落 + 表格单元格段落（W3 提前批）。paragraphIndex 语义只对正文段落：
      // ±3 邻近窗口只取 body 段落；表格段落不参与邻近，仅在唯一命中/上下文消歧路径上可命中。
      const { bodyItems, candidates, scanNote } = await loadBodyParagraphs(context);
      let candidateParas;
      const indexValid = Number.isInteger(requestedIndex) && requestedIndex >= 0 && Boolean(bodyItems[requestedIndex]);
      if (indexValid) {
        const indexes = [requestedIndex];
        for (let d = 1; d <= 3; d++) { indexes.push(requestedIndex - d, requestedIndex + d); }
        candidateParas = indexes.filter((idx, pos) => idx >= 0 && bodyItems[idx] && indexes.indexOf(idx) === pos).map(idx => bodyItems[idx]);
      } else candidateParas = candidates.map(c => c.paragraph);
      candidateParas.forEach(p => p.load('text'));
      await context.sync();
      const locateIn = list => {
        const found = [];
        for (const p of list) {
          const match = App.findSafeTextMatch(String(p.text || ''), find);
          if (match.status === 'ambiguous') throw App.makeStaleEditError('原文在目标段落中出现多次，无法安全判断要修改哪一处。', { currentText: String(p.text || '') });
          if (match.status === 'matched') found.push({ paragraph: p, match });
        }
        return found;
      };
      // find 多处命中时用卡片上的 contextBefore/contextAfter 消歧（enrichEditProposal 生成时已核实填充）。
      // 比对口径与填充口径（contextWindow）一致：同一 paragraph.text、按 match.start/end 切片取前后文。
      // 填充截断时会带「…」前/后缀，规范化后剥掉再比；空的一侧不参与过滤（段首/段尾本就为空）。
      // 恰好一个命中通过才算消歧成功；0 个或 2+ 个通过维持 stale（报错语义带「上下文也无法区分」）。
      const rawBefore = String(edit.contextBefore || '');
      const rawAfter = String(edit.contextAfter || '');
      const ctxBefore = rawBefore ? App.normalizeComparableText(rawBefore).replace(/^…+/, '') : '';
      const ctxAfter = rawAfter ? App.normalizeComparableText(rawAfter).replace(/…+$/, '') : '';
      let contextFailed = false;
      const disambiguateByContext = list => {
        if (list.length <= 1 || (!ctxBefore && !ctxAfter)) return null;
        const passing = list.filter(loc => {
          const text = String(loc.paragraph.text || '');
          if (ctxBefore && !App.normalizeComparableText(text.slice(0, loc.match.start)).endsWith(ctxBefore)) return false;
          if (ctxAfter && !App.normalizeComparableText(text.slice(loc.match.end)).startsWith(ctxAfter)) return false;
          return true;
        });
        if (passing.length !== 1) { contextFailed = true; return null; }
        return passing[0];
      };
      let located = locateIn(candidateParas);
      let location = located.find(x => x.paragraph === bodyItems[requestedIndex]) || (located.length === 1 ? located[0] : null) || disambiguateByContext(located);
      // 全文兜底：批量卡片里，前面卡片的插入/删除会让后续段落编号整体偏移，±3 段可能找不到。
      // 兜底扫全部候选（含表格单元格段落）：锚文本唯一或能被上下文消歧时应用仍然安全；
      // 找不到或不唯一才报「原文已变化」。
      if (!location && indexValid) {
        const globalParas = candidates.map(c => c.paragraph);
        globalParas.forEach(p => p.load('text'));
        await context.sync();
        const globalLocated = locateIn(globalParas);
        location = globalLocated.find(x => x.paragraph === bodyItems[requestedIndex]) || (globalLocated.length === 1 ? globalLocated[0] : null) || disambiguateByContext(globalLocated);
        if (!location) located = globalLocated;
      }
      if (!location) {
        const current = indexValid ? String(bodyItems[requestedIndex].text || '') : '';
        const multiMsg = contextFailed ? '附近有多个相似位置，上下文也无法区分，无法安全应用此项。' : '附近有多个相似位置，无法安全应用此项。';
        const baseMsg = located.length > 1 ? multiMsg : '原文已变化，请重新读取后生成此项。';
        // 表格扫描被性能守卫截断时要明说：找不到的可能在未扫描区域，不是真的已变化
        throw App.makeStaleEditError(scanNote ? `${baseMsg}（${scanNote}）` : baseMsg, { target: edit.target || null, currentText: current });
      }
      const paragraph = location.paragraph;
      const actualFind = location.match.text;
      const original = String(paragraph.text || '');
      const joiner = '\n';
      const inserted = placement === 'after' ? actualFind + joiner + replacement
        : placement === 'before' ? replacement + joiner + actualFind
        : replacement;
      if (actualFind.length <= 180) {
        const matches = paragraph.getRange().search(actualFind, { matchCase: false, matchWholeWord: false, matchWildcards: false });
        matches.load('items');
        await context.sync();
        const match = matches.items[0];
        if (!match) throw App.makeStaleEditError('定位过程中原文再次发生变化，请重新读取此项。', { target: edit.target || null, currentText: original });
        // 【24.2】读取/定位 sync 之后、写命令提交之前的取消检查
        if (token) token.throwIfCancelled();
        const range = await insertTextPreservingFormat(context, match, inserted, token);
        followSelect(range);
        try { await context.sync(); } catch (e) { /* 跟随选择是装饰性动作，失败不影响修改 */ }
        return { success: true, replaced: 1, totalFound: matches.items.length, usedParagraphTarget: true, ...(scanNote ? { scanNote } : {}), _navTarget: { kind: 'selection' } };
      }
      // 长文本（>180 字符）替换保格式：整段 insertText(Replace) 会把段内加粗、超链接等 run 格式全部抹平。
      // 改为公共前后缀差分——用 range.split 在段内偏移处切出中间差异区间，只重写中间段，
      // 前后缀对应的 run 保持不动，格式得以保留。偏移量来自 findSafeTextMatch 返回的真实区间（match.start/end），
      // 不是简单 indexOf（锚文本可能在段内出现多次，且匹配可能是规范化后的非精确命中）。
      // split API 在旧版 Word 不可用或切分失败时，降级回原有的整段重写路径。
      const prefixLen = commonPrefixLength(actualFind, inserted);
      const suffixLen = commonSuffixLength(actualFind, inserted, prefixLen);
      if (prefixLen + suffixLen < actualFind.length && typeof paragraph.getRange === 'function') {
        // split 降级只覆盖「切分」本身；写入成功后的收尾动作（选中跳转的 sync）失败
        // 绝不能落入整段重写——否则文字会被重复插入（真机出现过跟随 sync 报错触发重写）
        let pieces = null;
        try {
          const midStart = location.match.start + prefixLen;
          const midEnd = location.match.end - suffixLen;
          pieces = paragraph.getRange().split([midStart, midEnd]);
          pieces.load('items');
          await context.sync();
        } catch (e) { pieces = null; }
        if (pieces && pieces.items.length === 3) {
          if (token) token.throwIfCancelled();   // 【24.2】切分 sync 之后、写命令提交之前
          const range = await insertTextPreservingFormat(context, pieces.items[1], inserted.slice(prefixLen, inserted.length - suffixLen), token);
          followSelect(range);
          try { await context.sync(); } catch (e) { /* 跟随选择是装饰性动作，失败不影响修改 */ }
          return { success: true, replaced: 1, totalFound: 1, usedParagraphFallback: true, preservedFormatting: true, ...(scanNote ? { scanNote } : {}), _navTarget: { kind: 'selection' } };
        }
      }
      // 【24.2】整段回退路径同样是正文写命令：提交前检查取消，取消异常不再落入其他写入
      if (token) token.throwIfCancelled();
      const updated = original.slice(0, location.match.start) + inserted + original.slice(location.match.end);
      const range = paragraph.insertText(updated, Word.InsertLocation.replace);
      followSelect(range);
      await context.sync();
      return { success: true, replaced: 1, totalFound: 1, usedParagraphFallback: true, ...(scanNote ? { scanNote } : {}), _navTarget: { kind: 'selection' } };
    });
  }

  // 应用用户在 Diff 卡上选中的提案版本：短文本优先精确搜索，长文本按段落定位。
  // 【25.2/R1/30.2】文档写入的取消代次与在途计数已上移到 host.js 的共享生命周期
  //（App.markDocWritesCancelled / App.docWriteBegin）——Excel/PPT 写入路径同语义接入，
  // 三宿主一套代次。本文件的 applyEdit 只负责进入登记与 token 传递。
  async function applyEdit(edit) {
    // 【23.3 R4-B】在途写入计数：applyEdit 进入=写入已提交宿主、Promise 尚未返回。
    // 停止协议据此区分「回合已终态但写入仍在途」——停止确认必须等该计数归零，
    // 迟到的写入收敛不允许被可恢复的停止确认掩盖。
    // 【24.2/24.3/25.2】操作级取消令牌：进入时捕获取消代次快照；之后任何停止发生
    // 即粘性取消——新运行的 stopRequested 重置不能让已取消的旧写入复活。已提交且
    // 不可取消的写入允许按实际结果完成（finally 归零，不提前确认、不宣称回滚）。
    const op = App.docWriteBegin('word.applyEdit');
    try {
      return await applyEditInner(edit, op.token);
    } finally {
      op.end();
    }
  }

  async function applyEditInner(edit, token) {
    const find = String(edit && edit.find || '');
    if (!find.trim()) throw new Error('提案缺少原文定位文本。');
    const placement = edit.placement || 'replace';
    return replaceTargetedWordEdit(edit, placement, token);
  }

  // ---- W3：check_consistency（蓝本 4.3，只读聚合）----
  // 术语/数字一致性的检索底稿：工具层做全文检索与计数（含表格单元格段落，复用 W3 提前批
  // 的统一候选集），一致性判断由模型负责；发现问题经用户确认后走 propose_edits，本工具零写入。
  // 数字扫描取舍：只聚合出现 ≥2 次的数字（出现一次的没有一致性问题）；归一 key 去千分位
  // 逗号与空白（'1,200' 与 '1200' 同 key），但 occurrences 保留各自原样，写法统一由模型判断。
  function normalizeNumberToken(raw) {
    return String(raw).replace(/[，,\s]/g, '');
  }
  async function checkConsistency(args = {}) {
    requireOffice();
    const terms = (Array.isArray(args.terms) ? args.terms : []).map(t => String(t || '').trim()).filter(Boolean).slice(0, 20);
    return Word.run(async context => {
      const { candidates, scanNote } = await loadBodyParagraphs(context);
      candidates.forEach(c => c.paragraph.load('text'));
      await context.sync();
      const paras = candidates.map(c => {
        const where = c.index != null
          ? { paragraphIndex: c.index }
          : { tableIndex: c.tableIndex, row: c.rowIndex, column: c.columnIndex };
        return Object.assign({ text: String(c.paragraph.text || '') }, where);
      });
      const lower = s => String(s || '').toLowerCase();
      const termResults = terms.map(term => {
        const needle = lower(term);
        const occurrences = [];
        for (const p of paras) {
          const hay = lower(p.text);
          let at = hay.indexOf(needle);
          while (at >= 0 && occurrences.length < 50) {
            const { text, ...where } = p;
            occurrences.push(Object.assign({
              contextBefore: text.slice(Math.max(0, at - 30), at),
              contextAfter: text.slice(at + needle.length, at + needle.length + 30)
            }, where));
            at = hay.indexOf(needle, at + needle.length);
          }
        }
        return { term, count: occurrences.length, occurrences };
      });
      const numberMap = new Map();
      for (const p of paras) {
        for (const m of p.text.matchAll(/\d[\d,，.．]*\s*%?/g)) {
          const raw = m[0].trim();
          const key = normalizeNumberToken(raw);
          if (!key || !/\d/.test(key)) continue;
          if (!numberMap.has(key)) numberMap.set(key, []);
          const { text, ...where } = p;
          numberMap.get(key).push(Object.assign({ raw, context: text.slice(Math.max(0, m.index - 24), m.index) }, where));
        }
      }
      const numbers = [...numberMap.entries()]
        .filter(([, list]) => list.length >= 2)
        .map(([value, list]) => ({ value, count: list.length, occurrences: list.slice(0, 30) }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 60);
      return {
        success: true, paragraphCount: paras.length, terms: termResults, numbers, scanNote,
        note: '本工具只做检索与计数，零写入。术语一致性（同一概念的不同写法）与数字一致性（同一指标的数值冲突）由你比对 occurrences 判断；发现的问题先列成带引用位置的清单向用户报告，确认后用 propose_edits 修改。'
      };
    });
  }

  // ---- W3：insert_hyperlink（蓝本 4.5）----
  // vendored word-mac-16.00.js（16.0.20416.15170）取证：Range 与 Document 均导航 hyperlinks →
  // HyperlinkCollection.add(anchor, options)（官方 WordApiDesktop 1.3，desktop-only，Mac 桌面适用）；
  // Hyperlink 标量含 address/screenTip/textToDisplay。textToDisplay 不暴露——不换显示文字，
  // 与 PPT set_hyperlink 同口径，保持纯格式类。写后读回三态（W2 模式）。
  // find 多命中直接拒绝：Word 是文字文档，链接加错位置比拒一次更糟。
  async function insertHyperlink(args = {}) {
    requireOffice();
    const address = String(args.address || '').trim();
    if (!address) throw new Error('address is required (https://…, mailto:…, or a file name).');
    const screenTip = args.screenTip ? String(args.screenTip) : '';
    const find = args.find != null ? String(args.find) : '';
    return Word.run(async context => {
      let anchor;
      if (find) {
        const results = context.document.body.search(find, { matchCase: true });
        results.load('items');
        await context.sync();
        if (!(results.items && results.items.length)) {
          throw new Error(`find "${find}" 在正文中没有命中。先用 search_text 确认原文；find 必须与文档文字完全一致。`);
        }
        if (results.items.length > 1) {
          throw new Error(`find "${find}" 命中 ${results.items.length} 处，为避免链接加错位置已拒绝执行。补充前后文让 find 唯一，或让用户在 Word 里选中目标文字后不带 find 重试。`);
        }
        anchor = results.items[0];
      } else {
        anchor = context.document.getSelection();
        anchor.load('text');
        await context.sync();
        if (!String(anchor.text || '').trim()) throw new Error('当前没有选中文本。请在 Word 里选中要加链接的文字，或用 find 指定目标文本。');
      }
      let hyperlink;
      try {
        hyperlink = anchor.hyperlinks.add(anchor, { address, screenTip: screenTip || undefined });
        hyperlink.load('address');
        await context.sync();
      } catch (e) {
        throw new Error(`设置超链接失败（${String((e && e.message) || e).slice(0, 200)}）。本机 Word 可能不含 Hyperlink API（WordApiDesktop 1.3）。替代路径：把地址作为纯文本插入（insert_text），请用户选中后 Cmd+K 手动加链接。`);
      }
      const mismatches = [];
      if (String(hyperlink.address || '') !== address) mismatches.push(`address 读回 ${hyperlink.address}`);
      const verification = mismatches.length
        ? { verified: false, verifyWarning: `疑似未生效：${mismatches.join('；')}` }
        : { verified: true };
      return Object.assign({ success: true, mode: find ? 'find' : 'selection', address }, verification);
    });
  }

  // 真机能力探测（照 probeExcelCapabilities 模式）：宿主挂载后只读探测一次，零写入，
  // 全部 try/catch，结果由 app.js 挂到 host.capabilities。探测为 false 时 capabilityToolGates 摘牌。
  async function probeWordCapabilities() {
    const caps = { probedAt: new Date().toISOString() };
    try {
      await Word.run(async context => {
        const body = context.document.body;
        caps.getCommentsApi = !!(body && typeof body.getComments === 'function');
        const sel = context.document.getSelection();
        caps.insertCommentApi = !!(sel && typeof sel.insertComment === 'function');
        // Hyperlink 初筛（真值由 insert_hyperlink 首调的写后读回兜底——typeof 会被代理对象骗，W2 教训）
        caps.hyperlinksApi = !!(sel && sel.hyperlinks && typeof sel.hyperlinks.add === 'function');
        // 修订只读双路径初筛（MD 4.1/4.2）：原生 WordApi 1.6 或 OOXML 任一可读即暴露工具；
        // typeof 只是初筛，真机 sync 才验证门控，运行时失败由 get_tracked_changes 如实报、不写死探测缓存
        caps.getTrackedChangesApi = !!(body && typeof body.getTrackedChanges === 'function');
        caps.getOoxmlApi = !!(body && typeof body.getOoxml === 'function');
        caps.trackedChangesRead = !!(caps.getTrackedChangesApi || caps.getOoxmlApi);
      });
      caps.comments = !!(caps.getCommentsApi && caps.insertCommentApi);
    } catch (e) { caps.comments = false; caps.error = String(e && e.message || e); }
    return caps;
  }

  // 提案层校验（v141，对齐 Excel 的 vet 精神）：残缺/无变化的提案在出卡前拦截，
  // 回喂模型重新生成——真机实况：空文档插入时模型连发占位垃圾卡（label「跳过」、
  // replacement「占位」），渲染成卡只会让用户困惑。
  function vetEditProposalArgs(args) {
    const items = []
      .concat(Array.isArray(args && args.edits) ? args.edits : [])
      .concat(Array.isArray(args && args.changes) ? args.changes : []);
    if (!items.length) return '提案缺少 edits 或 changes 数组。纯插入（标题/新段落/备注）请直接用 insert_heading/insert_text，不要用 propose_edits。';
    for (let i = 0; i < items.length; i++) {
      const item = items[i] || {};
      const find = String(item.find == null ? '' : item.find).trim();
      const replacement = String(item.replacement == null ? '' : item.replacement);
      const placement = String(item.placement || 'replace');
      if (!replacement.trim() && placement !== 'replace') {
        return `edits/changes[${i}] 的 replacement 为空。插入类提案（placement: after/before）必须给出要插入的完整文本。`;
      }
      if (placement === 'replace' && !find) {
        return `edits/changes[${i}]（${String(item.label || '').slice(0, 30)}）是 replace 但 find 为空：replace 必须引用文档中的现有原文作为锚点。纯新增内容请改用 placement: "after"/"before" 并给出锚点，或直接用 insert_text/insert_heading。`;
      }
      if (placement === 'replace' && find === replacement.trim()) {
        return `edits/changes[${i}] 的 find 与 replacement 完全相同，这是无变化的提案。请给出真正修改后的文本。`;
      }
    }
    return null;
  }

  App.HOSTS.word = {
    hostType: 'word',
    available: true,
    metadataLabel: 'Document outline',
    systemPrompt: SYSTEM_PROMPT,
    toolDefinitions: TOOL_DEFINITIONS,
    toolExecutors: TOOL_EXECUTORS,
    defaultArgsForTool,
    evalToolName: 'eval_officejs',
    getMetadata: getDocumentMetadata,
    getFullContext,
    navigateCitation,
    follow: maybeFollow,
    applyEdit,
    enrichEditProposal,
    probeWordCapabilities,
    // confirm 模式下可直通：批注是低风险说明性写入（同 Excel comments 先例）；
    // delete 分支由 manageComment 内闸门单独拦截（操作级闸门先例见踩坑 11）。
    // apply_style / set_paragraph_format 是纯格式修改（字体/对齐/行距等），立即可见且可撤销,
    // 两个编辑模式下都直通（2026-08-31 用户 拍板:小格式修改不该被 confirm 拦,只读建议模式仍只给建议）。
    // 纯插入工具同为纯增量、可 ⌘Z、无审批价值，confirm 直通（2026-09-01 空文档 + 审核模式死锁修复，
    // 哲学对齐 Excel 批次 22「空目标直通、有占用才走卡」）；insert_text 的 Replace 覆盖选区有破坏性，
    // 由执行器内操作级闸门拦截并引导走 propose_edits。fill_table_cells 同理：空单元格直通，
    // 覆盖已有内容的单元格由执行器内闸门拦截（W3 提前批）。
    vetEditProposalArgs,
    directFormattingTools: ['manage_comment', 'apply_style', 'set_paragraph_format', 'insert_text', 'insert_table', 'fill_table_cells', 'insert_heading', 'insert_page_break', 'insert_image', 'insert_hyperlink', 'manage_image'],
    // 能力闸：探测 comments 为 false 时批注两工具自动摘牌；未探测（undefined）时保留。
    // insert_hyperlink 的 hyperlinksApi 是 typeof 初筛，执行时写后读回兜底。
    // get_tracked_changes 双路径只要一条可用就保留（trackedChangesRead = 原生 || OOXML）；
    // 两条全无才摘牌。运行时失败不算「不支持」，不进探测缓存。
    capabilityToolGates: { manage_comment: 'comments', get_comments: 'comments', insert_hyperlink: 'hyperlinksApi', get_tracked_changes: 'trackedChangesRead' },
    runtimeCapabilities() { return this.capabilities || {}; },
    i18n: {
      zh: {
        brand: 'Trojan AI', brandFooter: 'Trojan AI for Office',
        title: '准备好处理你的 Word 文档', subtitle: '你可以让我撰写、润色、排版或检索文档内容',
        input: '告诉我你想如何处理这份文档…',
        chart: '智能写作生成', chartDesc: '根据要求生成或续写段落',
        fix: '文档润色校对', fixDesc: '检查并修正措辞、语法与格式',
        analyze: '文档结构解析', analyzeDesc: '提炼大纲、生成摘要与结论',
        chartPrompt: '请根据当前文档主题，在文末续写一段合适的内容',
        fixPrompt: '帮我校对全文，修正语法和措辞问题并给出修改建议',
        analyzePrompt: '帮我读取文档大纲并生成一段整体摘要',
        demo: '当前不在 Word/Office 环境中，Word 工具只能在插件侧边栏里运行。'
      },
      en: {
        brand: 'Trojan AI', brandFooter: 'Trojan AI for Office',
        title: 'Ready to work with your Word document', subtitle: 'Ask me to write, edit, format, or search the document',
        input: 'Tell me what to do with this document…',
        chart: 'Writing Generation', chartDesc: 'Draft or continue paragraphs',
        fix: 'Proofread & Polish', fixDesc: 'Fix grammar, wording and formatting',
        analyze: 'Document Analysis', analyzeDesc: 'Extract outline and summarize',
        chartPrompt: 'Continue the document with a fitting new paragraph at the end',
        fixPrompt: 'Proofread the whole document and fix grammar and wording issues',
        analyzePrompt: 'Read the document outline and write an overall summary',
        demo: 'Not currently running inside Word/Office. Word tools only work in the add-in task pane.'
      }
    }
  };
})();
