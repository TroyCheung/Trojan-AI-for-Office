(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const requireOffice = () => App.requireOffice();

  // 闸口读「本轮有效访问模式」，与下发层同口径：全局 auto 但用户本轮要求先审时，
  // api.js 的 requestAccessMode() 会临时降级为 confirm，本文件的覆盖/删除/隐藏闸门必须读
  // 同一个值——否则模型经由 confirm 直通工具（set_cell_range 等）调进危险分支时，
  // 闸门仍看到持久化的 auto 而直接放行，用户以为在审核实际已被直写（host-word.js gateMode 同款）。
  function effectiveAccessMode() {
    try {
      if (typeof App.effectiveAccessMode === 'function') return App.effectiveAccessMode();
      const raw = typeof App.currentAccessMode === 'function' ? App.currentAccessMode() : (App.state.settings.accessMode || 'confirm');
      return raw === 'auto' && typeof App.requestRequiresReview === 'function' && App.requestRequiresReview() ? 'confirm' : raw;
    } catch { return 'confirm'; }
  }

  // ---- Excel 稳定 sheetId 映射 ----
  async function getSheetMap(context, sheets) {
    let map = {};
    try { map = JSON.parse(App.loadDocSetting(App.STORAGE_KEYS.sheetMap, '{}') || '{}'); } catch { map = {}; }
    let max = Object.values(map).reduce((a, b) => Math.max(a, Number(b) || 0), 0);
    let dirty = false;
    for (const s of sheets) {
      if (!map[s.id]) { map[s.id] = ++max; dirty = true; }
    }
    if (dirty) await App.saveDocSetting(App.STORAGE_KEYS.sheetMap, JSON.stringify(map)).catch(() => {});
    return new Map(Object.entries(map));
  }
  async function worksheetById(context, stableId) {
    const sheets = context.workbook.worksheets;
    sheets.load('items');
    await context.sync();
    for (const sheet of sheets.items) sheet.load('id,name');
    await context.sync();
    const map = await getSheetMap(context, sheets.items);
    for (const sheet of sheets.items) if (Number(map.get(sheet.id)) === Number(stableId)) return sheet;
    return null;
  }
  function colName(index) { let s = '', n = index; while (n >= 0) { s = String.fromCharCode(n % 26 + 65) + s; n = Math.floor(n / 26) - 1; } return s; }
  function a1(row, col) { return `${colName(col)}${row + 1}`; }
  // Excel 序列数转 ISO 日期：25569 是 1970-01-01 在 Excel 1900 日期系统中的序列数
  function excelSerialToIsoDate(serial) {
    if (typeof serial !== 'number' || !Number.isFinite(serial)) return null;
    const ms = Math.round((serial - 25569) * 86400000);
    const d = new Date(ms);
    return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
  }
  // 判断 numberFormat 是否为日期/时间格式：去掉引号字面量、[...] 区域/颜色段和反斜杠转义后，仍含 y/m/d/h/s 字母即视为日期时间格式
  function isDateNumberFormat(format) {
    if (typeof format !== 'string') return false;
    const cleaned = format.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
    return /[ymdhs]/i.test(cleaned);
  }
  // 占用检查（set_cell_range 原有思路的复用）：目标区域存在非空单元格时抛错，让模型带 allow_overwrite=true 与用户确认后重试
  function assertRangeEmpty(range, label) {
    const start = parseStart(range.address);
    const occupied = [];
    for (let i = 0; i < range.rowCount; i++) for (let j = 0; j < range.columnCount; j++) {
      if ((range.values[i][j] !== null && range.values[i][j] !== '') || (typeof range.formulas[i][j] === 'string' && range.formulas[i][j].startsWith('='))) occupied.push(a1(start.startRow + i, start.startCol + j));
    }
    if (occupied.length) throw new Error(`Would overwrite ${occupied.length} non-empty cell(s) in ${label}: ${occupied.slice(0, 10).join(', ')}${occupied.length > 10 ? '...' : ''}. Retry with allow_overwrite=true if confirmed.`);
  }
  function quoteSheetName(name) { return `'${String(name || '').replace(/'/g, "''")}'`; }
  function parseStart(address) {
    const part = String(address || 'A1').split('!').pop().split(':')[0].replace(/'/g, '');
    const m = part.match(/([A-Z]+)(\d+)/i);
    if (!m) return { startCol: 0, startRow: 0 };
    const col = m[1].toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    return { startCol: col, startRow: Number(m[2]) - 1 };
  }
  function rangeForDimension(ref, count, dim) {
    count = Number(count || 1);
    if (dim === 'rows') {
      const start = Number(ref || 1);
      return `${start}:${start + count - 1}`;
    }
    const start = String(ref || 'A').toUpperCase();
    const n = start.split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    return `${start}:${colName(n + count - 1)}`;
  }

  // ---- 公式安全双段验证（EXCEL-BLUEPRINT 批 1）----
  // 写前静态检查：只查结构性语法错误（空公式/引号不配对/括号不平衡/结尾是运算符），
  // 不做语义判断。返回 null 表示通过。
  function validateFormulaSyntax(formula) {
    if (typeof formula !== 'string' || !formula.startsWith('=')) return null;
    const body = formula.slice(1);
    if (!body.trim()) return 'Empty formula';
    if (((body.match(/"/g) || []).length) % 2 !== 0) return 'Unbalanced quotes';
    let depth = 0, inQuote = false;
    for (const ch of body) {
      if (ch === '"') { inQuote = !inQuote; continue; }
      if (inQuote) continue;
      if (ch === '(') depth++;
      else if (ch === ')' && --depth < 0) return 'Unbalanced parentheses';
    }
    if (depth !== 0) return 'Unbalanced parentheses';
    if (/[+\-*/^&,]$/.test(body.trim())) return 'Formula ends with an operator';
    return null;
  }
  // 写后回扫：写入 sync 后回读 values，收集 Excel 错误码闭集（#REF! #DIV/0! #N/A …）。
  // 闭集判断，不把「#1 热门话题」这类普通文本误报成公式错误（review P2#5）。
  const EXCEL_ERROR_RE = /^#(REF!|DIV\/0!|VALUE!|NAME\?|NUM!|N\/A|NULL!|SPILL!|CALC!|GETTING_DATA|BLOCKED!|CONNECT!|FIELD!)/;
  // formulaMatrix 是本次写入的矩阵，用于在错误清单里附上原公式方便模型修复。
  function scanFormulaErrors(values, startRow, startCol, formulaMatrix) {
    const errors = [];
    for (let i = 0; i < values.length; i++) for (let j = 0; j < values[i].length; j++) {
      const v = values[i][j];
      if (typeof v === 'string' && EXCEL_ERROR_RE.test(v.trim())) {
        const f = formulaMatrix && formulaMatrix[i] && formulaMatrix[i][j];
        errors.push({ address: a1(startRow + i, startCol + j), error: v.trim(), ...(typeof f === 'string' && f.startsWith('=') ? { formula: f } : {}) });
      }
    }
    return errors;
  }
  // 回读验证的体积上限：超过则跳过并明示（与读取 cellLimit 同档，避免大写入拖慢回合）
  const WRITE_VERIFY_CELL_LIMIT = 2000;
  // 提交前的校验/拒绝一律带 writeState:'not_committed'：这些错误发生在任何写入之前，
  // 「确定未写入」让结局合同判成 not_written，同参数修正重试不被重放保护误拦。
  function notCommittedError(message) {
    const e = new Error(message);
    e.writeState = 'not_committed';
    return e;
  }
  // expectedCells 单格语义的唯一出处：setCellRange 应用期核验与 enrichEditProposal 出卡前
  // 核验共用同一解析，两处口径不得分叉。
  // {formula:"=.."} 按公式文本比（读 formulas）；{value:..}/裸值按值比（读 values，
  // 经 App.valuesEquivalent，null/空串互认、数字按 Object.is、文本走归一化比对）。
  function expectedCellSpec(spec) {
    const hasFormula = spec && typeof spec === 'object' && typeof spec.formula === 'string' && spec.formula.startsWith('=');
    const expected = spec && typeof spec === 'object' && Object.prototype.hasOwnProperty.call(spec, 'value') ? spec.value : (hasFormula ? spec.formula : spec);
    return { isFormula: hasFormula, expected };
  }
  function expectedCellMatches(spec, actualValue, actualFormula) {
    const { isFormula, expected } = expectedCellSpec(spec);
    return App.valuesEquivalent(isFormula ? actualFormula : actualValue, expected);
  }
  // 快照失配的证据展示：封顶行数与超长截断（60 字符），不倾倒整块矩阵。
  const SNAPSHOT_EVIDENCE_CELL_CAP = 8;
  // set_cell_range 单次写入与提案出卡前核验共用的格数上限（同一常量，不另立阈值）：
  // verifyCellSnapshots 借同一上限拦下「快照核验就要物化超大区域」的提案。
  const MAX_WRITE_CELLS = 10000;
  function clampCellEvidence(s) {
    const t = typeof s === 'string' ? s : String(s);
    return t.length > 60 ? t.slice(0, 60) + `…（已截断，全长 ${t.length} 字符）` : t;
  }
  // 展示单格证据：先截内容再加引号——引号包着的是截断后的文本，超长原文不整段透出。
  // 对象走 JSON（封顶截断）：{val:5} 这类畸形规格不得显示成 [object Object]。
  function cellPieceText(v) {
    if (v === null || v === undefined || v === '') return '(空)';
    if (typeof v === 'string') return `"${clampCellEvidence(v)}"`;
    if (typeof v === 'object') return clampCellEvidence(JSON.stringify(v));
    return clampCellEvidence(String(v));
  }
  // expectedCells 单格规格的形态判定：合法形态 = null / 裸原始值 / {value: …} /
  // {formula: "=…"}；同带两者或字段不可识别 → 返回可读问题串（vet 静态拒绝
  // 与 enrich 运行时兜底共用）。只做形态判定，不改两侧比较语义。
  function expectedSpecProblem(spec) {
    if (spec == null || typeof spec !== 'object') return null;
    const hasValue = Object.prototype.hasOwnProperty.call(spec, 'value');
    const hasFormula = Object.prototype.hasOwnProperty.call(spec, 'formula');
    const validFormula = typeof spec.formula === 'string' && spec.formula.startsWith('=');
    if (hasValue && hasFormula) return '同时带 value 与 formula，请二选一（值填 {value: …}，公式填 {formula: "=…"}）';
    if (hasValue) return null;   // {value:null} / {value:""} 是合法空值形态，保留
    if (validFormula) return null;
    if (hasFormula) return `formula 必须以 "=" 开头的公式文本（收到 ${clampCellEvidence(JSON.stringify(spec.formula))}）`;
    const keys = Object.keys(spec).map(k => JSON.stringify(k)).join(', ');
    return `没有可识别的 value / formula 字段（收到字段：${keys || '（空对象）'}）`;
  }
  // 矩阵最大行宽的安全求法：Math.max(...rows) 对模型可能提交的超大 expectedCells 会
  // 在上限检查前就栈溢出（spread 展开几十万元素），一律用循环。
  function matrixRowWidth(matrix) {
    let w = 0;
    for (const row of matrix || []) if (Array.isArray(row) && row.length > w) w = row.length;
    return w;
  }
  function expectedCellEvidence(address, spec, actualValue, actualFormula) {
    const { isFormula, expected } = expectedCellSpec(spec);
    const expectedText = isFormula ? `公式 "${clampCellEvidence(expected)}"` : `值 ${cellPieceText(expected)}`;
    const formulaText = typeof actualFormula === 'string' && actualFormula.startsWith('=') ? `，公式 "${clampCellEvidence(actualFormula)}"` : '';
    return `${address}：提交期望 ${expectedText}，实际 值 ${cellPieceText(actualValue)}${formulaText}`;
  }
  // 失配报错（出卡前与应用期共用骨架）：总数明示 + 证据行封顶 + read_range 指引 + 零写入声明。
  function staleSnapshotMessage(label, sheetId, range, lines, total, action) {
    const head = `目标单元格已变化：${label ? `提案「${String(label).slice(0, 40)}」的` : ''}expectedCells 快照与当前工作表不符（sheetId ${sheetId} 区域 ${range}，失配 ${total} 格）`;
    const capNote = total > lines.length ? `\n（失配共 ${total} 格，以上仅展示前 ${lines.length} 格，其余以 read_range 读取结果为准。）` : '';
    return `${head}\n${lines.join('\n')}${capNote}\n${action} 未写入任何单元格。`;
  }
  // 会话级撤销栈（Office.js 写入不进 Excel 撤销栈——平台限制；快照来自每次写入前的真实值）
  const UNDO_STACK = [];
  async function undoLastWrite() {
    requireOffice();
    const entry = UNDO_STACK.pop();
    if (!entry) return { success: false, error: '没有可撤销的写入（本会话的撤销记录为空；格式调整与结构操作不在撤销范围内）。' };
    return TOOL_EXECUTORS.set_cell_range({ sheetId: entry.sheetId, range: entry.range, cells: entry.before, allow_overwrite: true, _skipUndo: true });
  }

  async function getWorkbookMetadata() {
    requireOffice();
    return Excel.run(async context => {
      const wb = context.workbook;
      wb.load('name');
      const sheets = wb.worksheets;
      sheets.load('items');
      const active = sheets.getActiveWorksheet();
      active.load('id,name');
      const selected = wb.getSelectedRange();
      selected.load('address');
      await context.sync();
      for (const sheet of sheets.items) sheet.load('id,name,position,visibility');
      await context.sync();
      const map = await getSheetMap(context, sheets.items);
      const info = [];
      for (const sheet of sheets.items) {
        const used = sheet.getUsedRangeOrNullObject();
        used.load('address,rowCount,columnCount');
        await context.sync();
        info.push({ sheetId: Number(map.get(sheet.id)), name: sheet.name, nativeId: sheet.id, position: sheet.position, visibility: sheet.visibility, maxRows: used.isNullObject ? 0 : used.rowCount, maxColumns: used.isNullObject ? 0 : used.columnCount, usedRange: used.isNullObject ? null : used.address.split('!').pop() });
      }
      return { success: true, workbookId: App.state.workbookId || 'workbook', workbookName: wb.name || '', activeSheet: { sheetId: Number(map.get(active.id)), name: active.name }, selectedRange: selected.address.includes('!') ? selected.address.split('!').pop() : selected.address, worksheets: info };
    });
  }

  // 目录注入（返回格式对齐 host-word.js 的 getFullContext）：
  // 每个工作表读前 8 行 × 前 10 列的值（CSV 形式，行尾空单元格与整块空行折叠），
  // 让模型每轮直接看到各表开头的真实数据，不再只靠 sheet 名单和尺寸猜范围盲读。
  // 全部表加起来超限时按 Word 同款协议返回 { truncated: true }，由 api.js 提示改用读取工具。
  const FULL_CONTEXT_CHAR_LIMIT = 15000;
  const FULL_CONTEXT_ROWS = 8;
  const FULL_CONTEXT_COLS = 10;

  function csvCell(value) {
    const text = (value === null || value === undefined) ? '' : String(value);
    return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  }

  async function getFullContext() {
    requireOffice();
    return Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items');
      await context.sync();
      for (const sheet of sheets.items) sheet.load('id,name');
      await context.sync();
      const map = await getSheetMap(context, sheets.items);
      const sections = [];
      for (const sheet of sheets.items) {
        const header = `[sheet ${Number(map.get(sheet.id))}] ${sheet.name}`;
        // 逐表独立 try：单表读取失败（保护表/异常区域）不拖垮整次注入
        try {
          const range = sheet.getRange(`A1:${colName(FULL_CONTEXT_COLS - 1)}${FULL_CONTEXT_ROWS}`);
          range.load('values');
          await context.sync();
          const rows = range.values.map(row => {
            const cells = row.map(csvCell);
            while (cells.length && cells[cells.length - 1] === '') cells.pop();
            return cells.join(',');
          });
          while (rows.length && rows[rows.length - 1] === '') rows.pop();
          sections.push(rows.length ? `${header}\n${rows.join('\n')}` : `${header}\n(空)`);
        } catch (e) {
          sections.push(`${header}\n(读取失败：${e.message})`);
        }
      }
      const text = sections.join('\n\n');
      if (text.length > FULL_CONTEXT_CHAR_LIMIT) {
        return { truncated: true, charCount: text.length, sheetCount: sheets.items.length };
      }
      return { truncated: false, charCount: text.length, sheetCount: sheets.items.length, text };
    });
  }

  async function getCellRanges(args) {
    requireOffice();
    const { sheetId, ranges, includeStyles = true, cellLimit = 2000 } = args;
    const requestedRanges = Array.isArray(ranges) ? ranges : [];
    const limit = Math.max(1, Number(cellLimit || 2000));
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.load('name');
      const used = sheet.getUsedRangeOrNullObject();
      used.load('address');
      await context.sync();
      const dimension = used.isNullObject ? 'A1' : used.address.split('!').pop();
      const cells = {}, formulas = {}, styles = {};
      let count = 0, hasMore = false;
      for (let rangeIndex = 0; rangeIndex < requestedRanges.length; rangeIndex++) {
        if (count >= limit) { hasMore = true; break; }
        const rangeText = requestedRanges[rangeIndex];
        const range = sheet.getRange(rangeText);
        range.load('values,formulas,address,rowCount,columnCount');
        await context.sync();
        const start = parseStart(range.address);
        const styleCells = [];
        let stoppedInsideRange = false;
        outer: for (let r = 0; r < range.rowCount; r++) {
          for (let c = 0; c < range.columnCount; c++) {
            if (count >= limit) { stoppedInsideRange = true; break outer; }
            const key = a1(start.startRow + r, start.startCol + c);
            const val = range.values[r][c];
            const formula = range.formulas[r][c];
            let returnedSomething = false;
            if (val !== null && val !== '' && typeof val !== 'undefined') { cells[key] = val; returnedSomething = true; }
            if (typeof formula === 'string' && formula.startsWith('=')) { formulas[key] = formula; returnedSomething = true; }
            if (returnedSomething) {
              count++;
              if (includeStyles) styleCells.push([key, range.getCell(r, c)]);
            }
          }
        }
        if (stoppedInsideRange || (count >= limit && rangeIndex < requestedRanges.length - 1)) hasMore = true;
        if (includeStyles && styleCells.length) {
          styleCells.forEach(([, cell]) => { cell.format.font.load('name,size,color,bold,italic,underline,strikethrough'); cell.format.fill.load('color'); cell.format.load('horizontalAlignment'); cell.load('numberFormat'); });
          await context.sync();
          for (const [key, cell] of styleCells) {
            const style = {};
            if (cell.format.font.name) style.fontFamily = cell.format.font.name;
            if (cell.format.font.size) style.fontSize = cell.format.font.size;
            if (cell.format.font.bold !== null) style.fontWeight = cell.format.font.bold ? 'bold' : 'normal';
            if (cell.format.font.italic !== null) style.fontStyle = cell.format.font.italic ? 'italic' : 'normal';
            if (cell.format.font.color) style.fontColor = cell.format.font.color;
            if (cell.format.fill.color) style.backgroundColor = cell.format.fill.color;
            if (cell.format.horizontalAlignment) style.horizontalAlignment = String(cell.format.horizontalAlignment).toLowerCase();
            if (cell.numberFormat) {
              style.numberFormat = cell.numberFormat;
              // 日期格式下读到的值是 Excel 序列数（如 45292），附带 ISO 日期避免把「2024/1/1」误读成数字
              if (isDateNumberFormat(cell.numberFormat)) {
                const iso = excelSerialToIsoDate(cells[key]);
                if (iso) style.isoDate = iso;
              }
            }
            if (Object.keys(style).length) styles[key] = style;
          }
        }
        if (hasMore) break;
      }
      return { success: true, hasMore, worksheet: { name: sheet.name, sheetId, dimension, cells, formulas, styles, borders: {} } };
    });
  }

  async function getRangeAsCsv(args) {
    requireOffice();
    const { sheetId, range, includeHeaders = true, maxRows = 500 } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.load('name');
      const r = sheet.getRange(range);
      r.load('values,rowCount,columnCount');
      await context.sync();
      const start = includeHeaders ? 0 : 1;
      const rows = [];
      for (let i = start; i < Math.min(r.rowCount, start + maxRows); i++) rows.push(r.values[i].map(csvEscape).join(','));
      return { success: true, csv: rows.join('\n'), rowCount: rows.length, columnCount: r.columnCount, hasMore: r.rowCount - start > maxRows, sheetName: sheet.name };
    });
  }
  function csvEscape(v) { if (v == null) return ''; const s = String(v); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; }

  // read_range：三模式统一读取（EXCEL-BLUEPRINT 批 1，对齐官方形态）。
  // compact（默认）＝markdown 表＋公式旁注＋全空明示；csv＝裸 CSV；detailed＝markdown 表＋
  // 公式/数字格式稀疏映射（改写场景用，日期序列数附 ISO）。旧 get_cell_ranges / get_range_as_csv
  // 从工具定义摘牌（模型只见 read_range），执行器保留原实现作兼容别名。
  // 读取上限：maxRows 截行 + 总单元格 2000 截断（整行/整列引用的 token 炸弹防护，review P2#6）。
  const READ_CELL_LIMIT = 2000;
  async function readRange(args) {
    requireOffice();
    const { sheetId, range, mode = 'compact', maxRows = 500 } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.load('name');
      const r = sheet.getRange(range);
      r.load('values,formulas,address,rowCount,columnCount');
      if (mode === 'detailed') r.load('numberFormat');
      await context.sync();
      const start = parseStart(r.address);
      if (r.columnCount > READ_CELL_LIMIT) {
        throw new Error(`Range too wide: ${r.columnCount} columns exceeds the ${READ_CELL_LIMIT}-cell read budget. Read specific columns instead (e.g. 'A1:Z100'), never whole rows like '1:1'.`);
      }
      const rowBudget = Math.max(1, Math.floor(READ_CELL_LIMIT / r.columnCount));
      const limit = Math.min(Math.max(1, Number(maxRows || 500)), rowBudget);
      const hasMore = r.rowCount > limit;
      const values = r.values.slice(0, Math.min(r.rowCount, limit));
      const formulas = r.formulas.slice(0, Math.min(r.rowCount, limit));
      const cellText = v => v == null ? '' : String(v).replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');
      const table = [ ['', ...Array.from({ length: r.columnCount }, (_, j) => colName(start.startCol + j))] ];
      for (let i = 0; i < values.length; i++) table.push([String(start.startRow + i + 1), ...values[i].map(cellText)]);
      const result = { success: true, mode, sheetName: sheet.name, range: r.address.split('!').pop(), rowCount: r.rowCount, columnCount: r.columnCount, hasMore };
      const isEmpty = values.every(row => row.every(v => v == null || v === ''));
      if (mode === 'csv') {
        result.csv = values.map(row => row.map(csvEscape).join(',')).join('\n');
        return result;
      }
      // markdown 表插入表头分隔行（| --- |）
      const lines = table.map(row => `| ${row.join(' | ')} |`);
      lines.splice(1, 0, `| ${table[0].map(() => '---').join(' | ')} |`);
      let text = lines.join('\n');
      if (isEmpty) text = `**${sheet.name}!${result.range}** — _All cells are empty._`;
      else {
        text = `**${sheet.name}!${result.range}** (${r.rowCount}×${r.columnCount}${hasMore ? `, first ${values.length} rows` : ''})\n\n${text}`;
        if (mode === 'compact') {
          const sidebar = [];
          for (let i = 0; i < formulas.length; i++) for (let j = 0; j < formulas[i].length; j++) {
            const f = formulas[i][j];
            if (typeof f === 'string' && f.startsWith('=')) sidebar.push(`${a1(start.startRow + i, start.startCol + j)}: ${f}`);
          }
          if (sidebar.length) text += `\n\nFormulas:\n${sidebar.join('\n')}`;
        }
      }
      result.markdown = text;
      if (mode === 'detailed') {
        const fMap = {}, nfMap = {}, isoMap = {};
        for (let i = 0; i < formulas.length; i++) for (let j = 0; j < formulas[i].length; j++) {
          const key = a1(start.startRow + i, start.startCol + j);
          const f = formulas[i][j];
          if (typeof f === 'string' && f.startsWith('=')) fMap[key] = f;
          const nf = r.numberFormat && r.numberFormat[i] && r.numberFormat[i][j];
          if (typeof nf === 'string' && nf && nf !== 'General') {
            nfMap[key] = nf;
            if (isDateNumberFormat(nf)) {
              const iso = excelSerialToIsoDate(values[i][j]);
              if (iso) isoMap[key] = iso;
            }
          }
        }
        result.formulas = fMap;
        result.numberFormats = nfMap;
        if (Object.keys(isoMap).length) result.isoDates = isoMap;
      }
      return result;
    });
  }

  async function searchData(args) {
    requireOffice();
    const { searchTerm, sheetId, range, offset = 0, options = {} } = args;
    const { matchCase = false, matchEntireCell = false, matchFormulas = false, useRegex = false, maxResults = 500 } = options;
    return Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items'); await context.sync();
      for (const s of sheets.items) s.load('id,name'); await context.sync();
      const map = await getSheetMap(context, sheets.items);
      const targetSheets = sheetId ? [await worksheetById(context, sheetId)].filter(Boolean) : sheets.items;
      const matches = [];
      const regex = useRegex ? new RegExp(searchTerm, matchCase ? 'g' : 'ig') : null;
      for (const sheet of targetSheets) {
        const r = range ? sheet.getRange(range) : sheet.getUsedRangeOrNullObject();
        r.load('values,formulas,address,rowCount,columnCount'); await context.sync();
        if (r.isNullObject) continue;
        const start = parseStart(r.address);
        for (let row = 0; row < r.rowCount; row++) for (let col = 0; col < r.columnCount; col++) {
          const val = r.values[row][col];
          const formula = r.formulas[row][col];
          const text = String(matchFormulas && formula ? formula : (val ?? ''));
          let ok;
          if (regex) { regex.lastIndex = 0; ok = regex.test(text); }
          else { const a = matchCase ? text : text.toLowerCase(); const b = matchCase ? searchTerm : searchTerm.toLowerCase(); ok = matchEntireCell ? a === b : a.includes(b); }
          if (ok) matches.push({ sheetName: sheet.name, sheetId: Number(map.get(sheet.id)), a1: a1(start.startRow + row, start.startCol + col), value: val, formula: typeof formula === 'string' && formula.startsWith('=') ? formula : null, row: start.startRow + row + 1, column: start.startCol + col + 1 });
        }
      }
      const slice = matches.slice(offset, offset + maxResults);
      return { success: true, matches: slice, totalFound: matches.length, returned: slice.length, offset, hasMore: offset + maxResults < matches.length, searchTerm, nextOffset: offset + maxResults < matches.length ? offset + maxResults : null };
    });
  }

  // writeToken【31-B】：提案应用路径（applyEdit）经第二参数传入的本操作取消令牌；
  // 直通调用（auto 写空单元格等）不传、无检查点，行为不变。
  async function setCellRange(args, writeToken) {
    requireOffice();
    const { sheetId, range, expectedCells, copyToRange, resizeWidth, resizeHeight, allow_overwrite = false, _skipUndo } = args;
    let cells = args.cells; // let：Excel.run 内按目标区域做行宽补齐后重赋
    if (!Array.isArray(cells) || !cells.length || !cells.every(Array.isArray)) throw new Error('cells 必须是二维数组（每个内层数组对应目标区域的一行）。例如 A1:A2 写入 500：cells=[[{value:500}],[{value:500}]]，而不是 [{value:500},{value:500}] 这样的扁平数组。请修正 cells 结构后重试。');
    const writeCellCount = cells.reduce((sum, row) => sum + row.length, 0);
    if (writeCellCount > MAX_WRITE_CELLS) throw notCommittedError(`Refusing to write ${writeCellCount} cells in one call. Split the write into chunks of ${MAX_WRITE_CELLS} cells or fewer.`);
    // 【用户 2026-08-31 拍板：空目标直通】confirm 模式下往空单元格写入是低风险操作（对齐官方
    // write_cells 语义），直接放行不出卡；覆盖已有内容仍必须走 propose_edits 卡。
    // 提案应用路径（带 expectedCells 快照）不受此闸——那是用户已在卡上点过「应用」的写入。
    const accessMode = effectiveAccessMode();
    const isProposalPath = Array.isArray(expectedCells);
    // 放行两类：提案应用（带 expectedCells=用户已在卡上批准）与 allow_overwrite 重试
    // （模型按占用报错指引带 true 重试的既定自愈路径；undo_last_write 恢复快照也走此路）
    if (accessMode === 'confirm' && !isProposalPath && !allow_overwrite) {
      const occupied = await Excel.run(async context => {
        const sheet = await worksheetById(context, sheetId);
        if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
        const targets = [sheet.getRange(range)];
        if (copyToRange) targets.push(sheet.getRange(copyToRange));
        for (const t of targets) t.load('values,formulas');
        await context.sync();
        return targets.some(t => (t.values || []).some(row => row.some(v => v !== null && v !== '' && v !== undefined))
          || (t.formulas || []).some(row => row.some(f => typeof f === 'string' && f.startsWith('='))));
      });
      if (occupied) {
        return { success: false, blocked: true, error: 'Review mode blocks overwriting existing cells. Writing into EMPTY cells is allowed directly; to modify existing content you MUST use propose_edits so the user reviews a diff card.' };
      }
    }
    // 【预检段】读当前值/校验/快照（只读不改——写入段见下方独立 run，保证 Excel 撤销单元干净）
    let undoBefore, undoBeforeFormulas;
    await Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.load('name');
      const r = sheet.getRange(range);
      r.load('values,formulas,address,rowCount,columnCount');
      // copyToRange 的目标区域也要先读出占用情况，在写入前一并校验，避免静默覆盖已有数据
      const copyTarget = copyToRange ? sheet.getRange(copyToRange) : null;
      if (copyTarget) copyTarget.load('values,formulas,address,rowCount,columnCount');
      await context.sync();
      // 行宽宽进（终验三轮：模型的 cells 常有空行/短行——语义是「这格不写」，补 null 对齐而非拒绝；
      // 超宽仍拒绝，防止写错位）。行数不足同样补全 null 行。
      const padded = [];
      for (let i = 0; i < r.rowCount; i++) {
        const row = Array.isArray(cells[i]) ? cells[i].slice() : [];
        if (row.length > r.columnCount) throw new Error(`cells[${i}] column count (${row.length}) exceeds target range column count (${r.columnCount}) — wrong target range?`);
        while (row.length < r.columnCount) row.push(null);
        padded.push(row);
      }
      cells = padded;
      const start = parseStart(r.address);
      if (expectedCells) {
        if (!Array.isArray(expectedCells) || expectedCells.length !== r.rowCount || expectedCells.some(row => !Array.isArray(row) || row.length !== r.columnCount)) {
          throw notCommittedError('expectedCells dimensions must match the target range');
        }
        // 比较口径抽到 expectedCellMatches（与出卡前核验同源）；错误从「地址清单 + 整块矩阵」
        // 改为逐格证据（提交期望 vs 真实值/公式，封顶 8 行、超长截断），并给 read_range 指引。
        // currentValues/currentFormulas 载荷保留（卡面「刷新重提」路径消费），消息本体不倾倒矩阵。
        const evidence = [];
        let mismatchTotal = 0;
        for (let i = 0; i < r.rowCount; i++) for (let j = 0; j < r.columnCount; j++) {
          if (expectedCellMatches(expectedCells[i][j], r.values[i][j], r.formulas[i][j])) continue;
          mismatchTotal++;
          if (evidence.length < SNAPSHOT_EVIDENCE_CELL_CAP) {
            evidence.push(expectedCellEvidence(a1(start.startRow + i, start.startCol + j), expectedCells[i][j], r.values[i][j], r.formulas[i][j]));
          }
        }
        if (mismatchTotal) {
          throw App.makeStaleEditError(
            staleSnapshotMessage(null, sheetId, range, evidence, mismatchTotal,
              `内容可能在出卡后被修改。请用 read_range(sheetId=${sheetId}, range="${range}") 重新读取该区域，按真实内容重填 expectedCells 后重新提交。`),
            { target: { sheetId, range }, mismatchedCells: mismatchTotal, evidence, currentValues: r.values, currentFormulas: r.formulas }
          );
        }
      }
      if (!allow_overwrite) assertRangeEmpty(r, range);
      if (copyTarget && !allow_overwrite) assertRangeEmpty(copyTarget, copyToRange);
      const matrix = cells.map(row => row.map(cell => cell && typeof cell === 'object' ? (cell.formula || (cell.value ?? null)) : (cell ?? null)));
      // 写前静态检查：结构性语法错误直接拒绝写入（模型会收到错误并重写公式）
      for (let i = 0; i < matrix.length; i++) for (let j = 0; j < matrix[i].length; j++) {
        const bad = validateFormulaSyntax(matrix[i][j]);
        if (bad) throw new Error(`Formula syntax error at ${a1(start.startRow + i, start.startCol + j)}: ${bad}. Rewrite the formula and retry.`);
      }
      // 撤销快照（插件层 undo_last_write 用）：写入前值已在本 run 读到
      undoBefore = r.values.map(row => row.slice());
      undoBeforeFormulas = r.formulas.map(row => row.slice());
    });
    // 【写入段】独立的干净 Excel.run：只有赋值 + 一次 sync，不 load 任何东西——
    // 官方 ChatGPT/Claude/GLM 插件的写入都可被 ⌘Z 撤销（用户 2026-08-31 实测），
    // 此前读写验证混在一个 run 多段 sync，破坏了 Excel 的撤销单元。写入必须自成一段。
    await Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const r = sheet.getRange(range);
      const matrix = cells.map(row => row.map(cell => cell && typeof cell === 'object' ? (cell.formula || (cell.value ?? null)) : (cell ?? null)));
      // 【R1/30.2 + 31-B】预检（第一个 run 的 expectedCells/占用校验）之后、真正赋值
      // 之前检查取消：读取的是本操作经参数传入的令牌（writeToken），与其他并发操作的
      // 生命周期互不干扰；停止发生即 WRITE_CANCELLED，本次写入零提交。
      if (writeToken) writeToken.throwIfCancelled();
      r.formulas = matrix;
      const sheetPrefix = `${quoteSheetName(sheet.name)}!`;
      const start2 = parseStart(range);
      for (let i = 0; i < cells.length; i++) {
        for (let j = 0; j < cells[i].length; j++) {
          const cellAddress = `${sheetPrefix}${a1(start2.startRow + i, start2.startCol + j)}`;
          applyCellOptions(context, r.getCell(i, j), cells[i][j], cellAddress);
        }
      }
      const copyTarget2 = copyToRange ? sheet.getRange(copyToRange) : null;
      if (copyTarget2) copyTarget2.copyFrom(r, Excel.RangeCopyType.all, false, false);
      applyResize(r, resizeWidth, resizeHeight);
      await context.sync();
    });
    // 【验证段】独立的读 run：写后回读扫公式错误（与官方机制不同，覆盖写入也验证——
    // Diff 卡应用恰恰是最需要「公式算没算出错」回执的场景）
    let verification;
    const start = parseStart(range);
    const matrixForScan = cells.map(row => row.map(cell => cell && typeof cell === 'object' ? (cell.formula || (cell.value ?? null)) : (cell ?? null)));
    if (writeCellCount <= WRITE_VERIFY_CELL_LIMIT) {
      await Excel.run(async context => {
        const sheet = await worksheetById(context, sheetId);
        const r = sheet.getRange(range);
        r.load('values');
        await context.sync();
        const formulaErrors = scanFormulaErrors(r.values, start.startRow, start.startCol, matrixForScan);
        verification = { checked: true, formulaErrorCount: formulaErrors.length, ...(formulaErrors.length ? { formulaErrors } : {}) };
      });
    } else {
      verification = { checked: false, reason: 'too-large', cellCount: writeCellCount };
    }
    // 写入成功后入撤销栈：公式格还原公式（{formula}），纯值格还原值（{value}）；
    // 上限 20 条防内存膨胀
    if (!_skipUndo) UNDO_STACK.push({
      sheetId, range,
      before: undoBeforeFormulas.map((row, i) => row.map((f, j) =>
        typeof f === 'string' && f.startsWith('=') ? { formula: f } : { value: undoBefore[i][j] ?? null }))
    });
    if (UNDO_STACK.length > 20) UNDO_STACK.shift();
    const dirty = [{ sheetId, range }];
    if (copyToRange) dirty.push({ sheetId, range: copyToRange });
    return { success: true, writtenRange: range, copiedTo: copyToRange || null, verification, _dirtyRanges: dirty };
  }
  function applyCellOptions(context, cell, spec, cellAddress) {
    if (!spec || typeof spec !== 'object') return;
    if (spec.note) {
      if (!context.workbook.notes || typeof context.workbook.notes.add !== 'function') throw new Error('Excel Notes API is not available in this Office host');
      context.workbook.notes.add(cellAddress, String(spec.note));
    }
    const st = spec.cellStyles || {};
    if (st.fontWeight) cell.format.font.bold = st.fontWeight === 'bold';
    if (st.fontStyle) cell.format.font.italic = st.fontStyle === 'italic';
    if (st.fontSize) cell.format.font.size = st.fontSize;
    if (st.fontFamily) cell.format.font.name = st.fontFamily;
    if (st.fontColor) cell.format.font.color = st.fontColor;
    if (st.backgroundColor) cell.format.fill.color = st.backgroundColor;
    if (st.horizontalAlignment) cell.format.horizontalAlignment = st.horizontalAlignment;
    if (st.numberFormat) cell.numberFormat = [[st.numberFormat]];
    const borderMap = { top: 'EdgeTop', bottom: 'EdgeBottom', left: 'EdgeLeft', right: 'EdgeRight' };
    for (const [side, val] of Object.entries(spec.borderStyles || {})) {
      if (!val || !borderMap[side]) continue;
      const b = cell.format.borders.getItem(borderMap[side]);
      if (val.color) b.color = val.color;
      if (val.weight) b.weight = val.weight;
      if (val.style) b.style = val.style === 'solid' ? 'Continuous' : val.style;
    }
  }
  function applyResize(r, width, height) { if (width) r.getEntireColumn().format.columnWidth = width.value; if (height) r.getEntireRow().format.rowHeight = height.value; }

  async function clearCellRange(args) {
    requireOffice();
    const { sheetId, range, clearType = 'contents' } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const map = { contents: Excel.ClearApplyTo.contents, formats: Excel.ClearApplyTo.formats, all: Excel.ClearApplyTo.all };
      sheet.getRange(range).clear(map[clearType] || Excel.ClearApplyTo.contents);
      await context.sync();
      return { success: true, cleared: range, clearType, _dirtyRanges: [{ sheetId, range }] };
    });
  }

  async function copyTo(args) {
    requireOffice();
    const { sheetId, sourceRange, destinationRange, allow_overwrite = false } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const dest = sheet.getRange(destinationRange);
      // 覆盖保护：目标区域有非空单元格且未显式允许覆盖时拒绝执行，避免 copyFrom 静默抹掉已有数据
      if (!allow_overwrite) {
        dest.load('values,formulas,address,rowCount,columnCount');
        await context.sync();
        assertRangeEmpty(dest, destinationRange);
      }
      dest.copyFrom(sheet.getRange(sourceRange), Excel.RangeCopyType.all, false, false);
      await context.sync();
      return { success: true, sourceRange, destinationRange, _dirtyRanges: [{ sheetId, range: destinationRange }] };
    });
  }

  // ---- 结构操作（EXCEL-BLUEPRINT 批 2：modify_sheet_structure 按官方语义三拆）----
  // 旧名保留为薄路由别名（TOOL_DEFINITIONS 已摘牌），insert/delete/hide/unhide/freeze 分发到新工具。
  // confirm 闸门（工具放行、破坏性分支拦截，沿用 v76 方案）：delete 拦、其余直放。
  async function insertDeleteRowsColumns(args) {
    requireOffice();
    const { sheetId, operation, dimension = 'rows', reference, count = 1, position = 'before' } = args;
    const accessMode = effectiveAccessMode();
    if (accessMode === 'confirm' && operation === 'delete') {
      return { success: false, blocked: true, error: '删除行列在确认模式下不可用（此模式下 eval_officejs 同样不可用）。请告诉用户：删除行列需要切换到「直接修改」模式后重试。' };
    }
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      let ref = reference;
      if (operation === 'insert' && position === 'after') {
        if (dimension === 'rows') ref = String(Number(reference || 1) + 1);
        else ref = colName(String(reference || 'A').toUpperCase().split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0));
      }
      const range = sheet.getRange(rangeForDimension(ref, count, dimension));
      if (operation === 'insert') range.insert(dimension === 'rows' ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.right);
      else range.delete(dimension === 'rows' ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left);
      await context.sync();
      return { success: true, operation, dimension, sheetId, _dirtyRanges: [{ sheetId, range: '*' }] };
    });
  }

  async function hideUnhideRowsColumns(args) {
    requireOffice();
    const { sheetId, operation, dimension = 'rows', reference, count = 1 } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const range = sheet.getRange(rangeForDimension(reference, count, dimension));
      if (dimension === 'rows') range.rowHidden = operation === 'hide';
      else range.columnHidden = operation === 'hide';
      await context.sync();
      return { success: true, operation, dimension, sheetId, _dirtyRanges: [{ sheetId, range: '*' }] };
    });
  }

  async function freezePanesTool(args) {
    requireOffice();
    const { sheetId, action, count, reference } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const fp = sheet.freezePanes;
      if (action === 'unfreeze') fp.unfreeze();
      else if (action === 'freeze_rows' || action === 'freeze_columns') {
        const n = Number(count);
        if (!Number.isFinite(n) || n < 1) throw new Error(`${action} requires a positive numeric count (got: ${count})`);
        if (action === 'freeze_rows') fp.freezeRows(n); else fp.freezeColumns(n);
      }
      else if (action === 'freeze_at') {
        // 'B3' = 冻结其上方行与左侧列；优先 freezeAt，旧宿主没有则降级为行列数冻结
        if (typeof fp.freezeAt === 'function') fp.freezeAt(String(reference || 'B2'));
        else {
          const m = String(reference || 'B2').match(/^([A-Z]+)(\d+)$/i);
          if (!m) throw new Error(`freeze_at expects a single cell like 'B3', got: ${reference}`);
          const cols = m[1].toUpperCase().split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
          const rows = Number(m[2]) - 1;
          if (cols > 0) fp.freezeColumns(cols);
          if (rows > 0) fp.freezeRows(rows);
        }
      } else throw new Error(`Unknown freeze_panes action: ${action}`);
      await context.sync();
      return { success: true, action, sheetId, _dirtyRanges: [{ sheetId, range: '*' }] };
    });
  }

  async function viewSettings(args) {
    requireOffice();
    const { sheetId, action, color } = args;
    // confirm 闸门：隐藏工作表把用户正在看的表藏起来，破坏性不亚于删表，拦下要求走 auto
    const accessMode = effectiveAccessMode();
    if (accessMode === 'confirm' && (action === 'hide_sheet' || action === 'very_hide_sheet')) {
      return { success: false, blocked: true, error: '隐藏工作表在确认模式下不可用。请告诉用户：需要隐藏工作表时切换到「直接修改」模式后重试。' };
    }
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      if (action === 'get') {
        sheet.load('name,showGridlines,showHeadings,tabColor,visibility');
        await context.sync();
        return { success: true, action, sheetId, sheetName: sheet.name, gridlines: sheet.showGridlines, headings: sheet.showHeadings, tabColor: sheet.tabColor, visibility: sheet.visibility };
      }
      if (action === 'show_gridlines') sheet.showGridlines = true;
      else if (action === 'hide_gridlines') sheet.showGridlines = false;
      else if (action === 'show_headings') sheet.showHeadings = true;
      else if (action === 'hide_headings') sheet.showHeadings = false;
      else if (action === 'set_tab_color') sheet.tabColor = color || '';
      else if (action === 'hide_sheet') sheet.visibility = 'Hidden';
      else if (action === 'show_sheet') sheet.visibility = 'Visible';
      else if (action === 'very_hide_sheet') sheet.visibility = 'VeryHidden';
      else throw new Error(`Unknown view_settings action: ${action}`);
      await context.sync();
      return { success: true, action, sheetId, _dirtyRanges: [] };
    });
  }

  async function modifySheetStructure(args) {
    // 旧工具名兼容路由：按 operation 分发到拆分后的三个工具
    const { operation } = args;
    if (operation === 'insert' || operation === 'delete') return insertDeleteRowsColumns(args);
    if (operation === 'hide' || operation === 'unhide') return hideUnhideRowsColumns(args);
    if (operation === 'freeze') {
      // 旧语义：reference 可以给边界（'C'=冻结前 3 列，'5'=冻结前 5 行），换算成 count 再走新工具
      const { dimension, reference, count } = args;
      let n;
      if (count != null && count !== '') n = Number(count);
      else if (reference != null && reference !== '') {
        n = dimension === 'columns'
          ? String(reference).toUpperCase().split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0)
          : Number(reference);
      }
      if (!Number.isFinite(n) || n < 1) n = 1;
      return freezePanesTool({ sheetId: args.sheetId, action: dimension === 'columns' ? 'freeze_columns' : 'freeze_rows', count: n });
    }
    if (operation === 'unfreeze') return freezePanesTool({ sheetId: args.sheetId, action: 'unfreeze' });
    throw new Error(`Unknown modify_sheet_structure operation: ${operation}`);
  }

  async function modifyWorkbookStructure(args) {
    requireOffice();
    // 宽进容错（终验：模型首调常传 name 而非 sheetName）：name 归一为 sheetName
    if (args.sheetName == null && args.name != null) args = Object.assign({}, args, { sheetName: args.name });
    const { operation, sheetId, sheetName, newName, tabColor, confirmed = false } = args;
    // 删除工作表是 Excel 中唯一不可撤销的操作（⌘Z 无效），auto 模式下也必须先拿到用户确认：
    // 未带 confirmed: true 时不执行，返回 needsConfirmation 让模型先向用户说明并征得同意，再以 confirmed: true 重调
    if (operation === 'delete' && confirmed !== true) {
      return { success: false, needsConfirmation: true, error: '删除工作表不可撤销。请先用一句话向用户确认（说明将删除哪个工作表、其中有什么内容），用户明确同意后再以 confirmed: true 重新调用。' };
    }
    return Excel.run(async context => {
      let result = { success: true, operation };
      if (operation === 'create') {
        const s = context.workbook.worksheets.add(sheetName || 'Sheet');
        if (tabColor) s.tabColor = tabColor;
        s.load('id,name'); await context.sync();
        const map = await getSheetMap(context, [s]);
        result.sheetId = Number(map.get(s.id)); result.name = s.name;
      } else {
        const sheet = await worksheetById(context, sheetId);
        if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
        if (operation === 'delete') sheet.delete();
        if (operation === 'rename') { sheet.name = newName || sheetName || sheet.name; if (tabColor) sheet.tabColor = tabColor; }
        if (operation === 'duplicate') {
          const copy = sheet.copy(Excel.WorksheetPositionType.after, sheet);
          if (newName) copy.name = newName;
          if (tabColor) copy.tabColor = tabColor;
          copy.load('id,name'); await context.sync();
          const map = await getSheetMap(context, [copy]);
          result.sheetId = Number(map.get(copy.id)); result.name = copy.name;
        }
      }
      await context.sync();
      result._dirtyRanges = result.sheetId ? [{ sheetId: result.sheetId, range: '*' }] : (sheetId ? [{ sheetId, range: '*' }] : []);
      return result;
    });
  }

  async function resizeRange(args) {
    requireOffice();
    const { sheetId, range, width, height } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const r = range ? sheet.getRange(range) : sheet.getUsedRangeOrNullObject();
      applyResize(r, width, height);
      await context.sync();
      return { success: true, sheetId, range: range || '*', _dirtyRanges: [{ sheetId, range: range || '*' }] };
    });
  }

  // ---- format_cells（EXCEL-BLUEPRINT 批 3）----
  // 命名样式预设（对齐官方 11 种）：先铺底再逐属性覆盖；range 支持逗号分隔多区域。
  const NAMED_STYLES = {
    header: { fill_color: '#134CFF', font_color: '#FFFFFF', bold: true },
    'total-row': { bold: true },
    subtotal: { bold: true },
    input: { fill_color: '#FFF2CC' },
    'blank-section': { fill_color: '#F2F2F2' },
    number: { number_format: '0.00' },
    integer: { number_format: '0' },
    currency: { number_format: '¥#,##0.00' },
    percent: { number_format: '0.0%' },
    ratio: { number_format: '0.00' },
    text: { number_format: '@' }
  };
  const BORDER_SIDES = { top: 'EdgeTop', bottom: 'EdgeBottom', left: 'EdgeLeft', right: 'EdgeRight' };

  async function formatCells(args) {
    requireOffice();
    // 宽进（终验收官：模型把 style 写成 namedStyle 被静默忽略还报成功）：
    // 常见别名归一 + 未识别参数给 warning 回执让模型自纠，不再无声吞掉
    const STYLE_ALIASES = ['namedStyle', 'namedStyles', 'preset', 'presetStyle', 'styles'];
    if (args.style == null) {
      const hit = STYLE_ALIASES.find(k => args[k] != null);
      if (hit) args = Object.assign({}, args, { style: args[hit] });
    }
    const KNOWN = ['sheetId', 'range', 'style', 'bold', 'italic', 'underline', 'font_color', 'font_size', 'font_name', 'fill_color', 'horizontal_alignment', 'wrap_text', 'number_format', 'borders', 'explanation'].concat(STYLE_ALIASES);
    const ignored = Object.keys(args).filter(k => !KNOWN.includes(k) && args[k] !== undefined && args[k] !== null);
    const warnings = ignored.map(k => `Ignored parameter "${k}" — not a format_cells parameter. Did you mean "style" or a per-property override (bold/font_color/number_format/…)?`);
    const { sheetId, range, style, borders, ...rest } = args;
    const overrides = {};
    for (const [k, v] of Object.entries(rest)) if (v !== undefined && v !== null) overrides[k] = v;
    const merged = {};
    for (const s of [].concat(style || [])) {
      if (NAMED_STYLES[s]) Object.assign(merged, NAMED_STYLES[s]);
      else throw new Error(`Unknown named style "${s}". Available: ${Object.keys(NAMED_STYLES).join(', ')}`);
    }
    Object.assign(merged, overrides);
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const parts = String(range || '').split(',').map(s => s.trim()).filter(Boolean);
      if (!parts.length) throw new Error('range is required (supports comma-separated multi-area, e.g. "A1:B2,D1:E2")');
      const applied = [];
      for (const part of parts) {
        const r = sheet.getRange(part);
        r.load('rowCount,columnCount,address');
        await context.sync();
        // 体积闸：整行/整列 × 数字格式矩阵会构造百万级数组（review P2#6），拒绝并要求具体区域
        if (merged.number_format !== undefined && r.rowCount * r.columnCount > 2000) {
          throw new Error(`Refusing to set number_format on ${part} (${r.rowCount}×${r.columnCount} cells, budget 2000). Apply it to specific rows/columns instead of whole rows/columns.`);
        }
        const f = r.format;
        if (merged.bold !== undefined) f.font.bold = merged.bold;
        if (merged.italic !== undefined) f.font.italic = merged.italic;
        if (merged.underline !== undefined) f.font.underline = merged.underline ? 'Single' : 'None';
        if (merged.font_color !== undefined) f.font.color = merged.font_color;
        if (merged.font_size !== undefined) f.font.size = merged.font_size;
        if (merged.font_name !== undefined) f.font.name = merged.font_name;
        if (merged.fill_color !== undefined) f.fill.color = merged.fill_color;
        if (merged.horizontal_alignment !== undefined) f.horizontalAlignment = { left: 'Left', center: 'Center', right: 'Right' }[merged.horizontal_alignment] || merged.horizontal_alignment;
        if (merged.wrap_text !== undefined) f.wrapText = merged.wrap_text;
        if (merged.number_format !== undefined) {
          // Range.numberFormat 需要二维矩阵；构造与区域同形的矩阵广播
          r.numberFormat = Array.from({ length: r.rowCount }, () => Array.from({ length: r.columnCount }, () => merged.number_format));
        }
        for (const [side, spec] of Object.entries(borders || {})) {
          if (!spec || !BORDER_SIDES[side]) continue;
          const b = f.borders.getItem(BORDER_SIDES[side]);
          if (spec.color) b.color = spec.color;
          if (spec.weight) b.weight = spec.weight;
          if (spec.style) b.style = spec.style === 'solid' ? 'Continuous' : spec.style;
        }
        applied.push(part);
      }
      await context.sync();
      // _dirtyRanges 按单区域拆条：maybeFollow 的 selectRange 不收逗号多区域地址（review P2#7）
      return { success: true, appliedRanges: applied, applied: Object.keys(merged).concat(Object.keys(borders || {})), ...(warnings.length ? { warnings } : {}), _dirtyRanges: parts.map(p => ({ sheetId, range: p })) };
    });
  }

  // ---- comments（批 3；2026-08-31 真机探测 Excel.Comment 存在，notesApi 不可用走此路）----
  async function commentsTool(args) {
    requireOffice();
    // 参数容错（终验第 5 步：模型会猜 list 动词和 text 参数名）——action 别名归一、content/text 互认
    let { sheetId, action, range, commentId, content, text } = args;
    if (action === 'list') action = 'read';
    if (content == null && text != null) content = text;
    // 再兜一层常见别名（终验三轮：模型把批注文本塞进 comment/note/body）
    if (content == null) {
      const alias = ['comment', 'note', 'body', 'message'].find(k => typeof args[k] === 'string' && args[k].trim());
      if (alias) content = args[alias];
    }
    // range 带表名前缀（'Sheet1!B2'）时剥掉——批注地址是「所在表 + 单格」
    if (typeof range === 'string' && range.includes('!')) range = range.split('!').pop().replace(/'/g, '');
    args = Object.assign({}, args, { action, content, range });
    const validActions = ['read', 'add', 'update', 'reply', 'delete', 'resolve', 'reopen'];
    if (!validActions.includes(action)) throw new Error(`action must be one of ${validActions.join('/')} (got: ${action}). Use "read" to list comments — it needs NO commentId.`);
    // confirm 闸门：删除批注有破坏性（可能删掉用户手写的批注），与删行列同待遇
    const accessMode = effectiveAccessMode();
    if (accessMode === 'confirm' && action === 'delete') {
      return { success: false, blocked: true, error: '删除批注在确认模式下不可用。请告诉用户：需要删除批注时切换到「直接修改」模式后重试。' };
    }
    return Excel.run(async context => {
      const coll = context.workbook.comments;
      if (!coll || typeof coll.add !== 'function') throw new Error('Excel Comments API is not available in this Office host');
      if (action === 'read') {
        coll.load('items');
        await context.sync();
        // 属性名按官方文档：正文是 content（不是 text），位置用 getRange()
        for (const c of coll.items) {
          try {
            c.load('id,authorName,content,resolved');
            c.getRange().load('address');
          } catch { continue; }
        }
        await context.sync();
        const list = [];
        for (const c of coll.items) {
          let cell = null;
          try { cell = c.getRange().address; } catch {}
          list.push({ id: c.id, author: c.authorName, content: c.content, resolved: c.resolved, ...(cell ? { cell } : {}) });
        }
        return { success: true, action, comments: list, count: list.length };
      }
      const needsContent = action === 'add' || action === 'update' || action === 'reply';
      if (needsContent && !String(content || '').trim()) {
        // 报错带非空参数清单与示例（终验三轮：模型不带 content 空转三次被截停，它需要一眼看懂差什么）
        const sent = Object.keys(args).filter(k => !['explanation'].includes(k) && args[k] !== undefined && args[k] !== null && args[k] !== '').join(', ');
        throw new Error(`content is required for ${action}. You sent non-empty parameters: [${sent || 'none'}]. Correct example: { "action": "add", "range": "B2", "content": "批注文字" }. The comment TEXT must go in the "content" parameter.`);
      }
      if (action === 'add') {
        const sheet = await worksheetById(context, sheetId);
        if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
        sheet.load('name');
        await context.sync();
        if (String(range || '').includes(':')) throw new Error('comments add requires a single cell, not a range');
        const c = coll.add(`${quoteSheetName(sheet.name)}!${range}`, String(content));
        c.load('id');
        await context.sync();
        return { success: true, action, commentId: c.id, _dirtyRanges: [{ sheetId, range }] };
      }
      if (!commentId) throw new Error(`commentId is required for ${action}`);
      const c = coll.getItem(String(commentId));
      if (action === 'update') { c.content = String(content); }
      if (action === 'reply') { c.replies.add(String(content)); }
      if (action === 'delete') { c.delete(); }
      if (action === 'resolve') { c.resolved = true; }
      if (action === 'reopen') { c.resolved = false; }
      await context.sync();
      return { success: true, action, commentId: String(commentId), _dirtyRanges: action === 'delete' ? [] : [{ sheetId, range: range || '*' }] };
    });
  }

  // ---- trace_dependencies（EXCEL-BLUEPRINT 批 4，公式诊断的地基）----
  // Office.js 无原生公式追踪 API，自行实现：解析公式串中的单元格引用（precedents 上游），
  // 反向扫描全表公式建索引（dependents 下游），BFS 递归到指定深度，防环。
  // 已知不解析：命名区域、结构化引用（Table[列]）——遇到按原样展示，不标错。
  function stripStringLiterals(body) { return body.replace(/"(?:[^"]|"")*"/g, '""'); }
  function extractRefs(formula) {
    if (typeof formula !== 'string' || !formula.startsWith('=')) return [];
    const body = stripStringLiterals(formula.slice(1));
    const refs = [];
    // 左边界断言 (?<![A-Za-z0-9_.])：挡住 Table1[ 的 ble1、Sheet1:Sheet2!A1 的 eet1、1E5 的 E5 这类幽灵引用；
    // (?![(\w]) 右边界排除 LOG10( 函数名。整列/整行（A:A、1:1）与命名区域、结构化引用不解析（已知限制）。
    const cell = '(?<![A-Za-z0-9_.])(\\$?[A-Z]{1,3}\\$?\\d+)(?![(\\w])';
    // 跨表：'带空格 名'!A1:B2 或 Sheet1!A1
    const crossRe = new RegExp("(?:'((?:[^']|'')+)'|([A-Za-z_][\\w.]*))!(" + cell + '(?::' + cell + ')?)', 'gi');
    let m;
    const crossSpans = [];
    while ((m = crossRe.exec(body))) {
      refs.push({ sheet: (m[1] || m[2]).replace(/''/g, "'"), ref: m[3].replace(/\$/g, '') });
      crossSpans.push([m.index, m.index + m[0].length]);
    }
    // 同表：先抹掉跨表片段再匹配，避免重复计入
    let same = '';
    let pos = 0;
    for (const [s, e] of crossSpans) { same += body.slice(pos, s) + ' '; pos = e; }
    same += body.slice(pos);
    const sameRe = new RegExp('(' + cell + ')(?::(' + cell + '))?', 'gi');
    while ((m = sameRe.exec(same))) {
      const a = m[1].replace(/\$/g, ''), b = m[3] ? m[3].replace(/\$/g, '') : null;
      refs.push({ sheet: null, ref: b ? `${a}:${b}` : a });
    }
    return refs;
  }
  function refToRect(ref) {
    const [a, b] = ref.split(':');
    const pa = parseStart(a), pb = b ? parseStart(b) : pa;
    return { r1: Math.min(pa.startRow, pb.startRow), c1: Math.min(pa.startCol, pb.startCol), r2: Math.max(pa.startRow, pb.startRow), c2: Math.max(pa.startCol, pb.startCol) };
  }
  function rectContains(rect, row, col) { return row >= rect.r1 && row <= rect.r2 && col >= rect.c1 && col <= rect.c2; }
  function normalizeSheetName(name) { return String(name || '').toLowerCase().replace(/^'|'$/g, '').replace(/''/g, "'"); }

  async function traceDependencies(args) {
    requireOffice();
    const { cell, mode = 'precedents', depth } = args;
    if (String(cell || '').includes(':')) throw new Error('trace_dependencies expects a single cell, not a range.');
    const maxDepth = Math.min(Math.max(1, Number(depth) || 2), 5);
    const m = String(cell || '').match(/^(?:'([^']+)'|([^!]+))!(\$?[A-Z]{1,3}\$?\d+)$/i);
    if (!m) throw new Error(`cell must include the sheet name, e.g. "Sheet1!D10" (got: ${cell})`);
    const rootSheet = m[1] || m[2];
    const rootAddr = m[3].replace(/\$/g, '');
    return Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items');
      await context.sync();
      for (const s of sheets.items) s.load('id,name');
      await context.sync();
      const rootSheetObj = sheets.items.find(s => normalizeSheetName(s.name) === normalizeSheetName(rootSheet));
      if (!rootSheetObj) throw new Error(`Worksheet "${rootSheet}" not found`);
      // 读起点格
      const rootRange = rootSheetObj.getRange(rootAddr);
      rootRange.load('values,formulas');
      await context.sync();
      const rootFormula = typeof rootRange.formulas[0][0] === 'string' && rootRange.formulas[0][0].startsWith('=') ? rootRange.formulas[0][0] : null;
      const rootValue = rootRange.values[0][0];

      // 依赖索引：所有工作表的公式格 → 引用列表（dependents 方向用；precedents 不需要全表）
      let formulaIndex = null;
      if (mode === 'dependents') {
        formulaIndex = [];
        for (const s of sheets.items) {
          const used = s.getUsedRangeOrNullObject();
          // 两次 sync：先拿尺寸过体积闸，再加载 formulas——否则整表公式先过桥传输完才抛错（review 批 4 P1#3）
          used.load('address,rowCount,columnCount');
          await context.sync();
          if (used.isNullObject) continue;
          if (used.rowCount * used.columnCount > 50000) throw new Error(`Sheet "${s.name}" is too large (${used.rowCount}×${used.columnCount}) for dependents tracing. Remove unused rows/columns or trace precedents instead.`);
          used.load('formulas');
          await context.sync();
          const start = parseStart(used.address);
          for (let r = 0; r < used.rowCount; r++) for (let c = 0; c < used.columnCount; c++) {
            const f = used.formulas[r] && used.formulas[r][c];
            if (typeof f === 'string' && f.startsWith('=') && extractRefs(f).length) {
              formulaIndex.push({ sheet: s.name, addr: a1(start.startRow + r, start.startCol + c), formula: f, refs: extractRefs(f) });
            }
          }
        }
      }

      const visited = new Set([`${normalizeSheetName(rootSheet)}!${rootAddr}`]);
      const lines = [`**${mode === 'dependents' ? 'Dependents' : 'Precedents'} tree for ${rootSheet}!${rootAddr}**${rootFormula ? ` (${rootFormula})` : ` = ${rootValue ?? ''}`}:`];
      let found = 0;
      const RANGE_EXPAND_LIMIT = 16; // 区域 ≤16 格展开为单格继续追公式链；更大保持区域节点（防节点爆炸）

      // 引用列表 → 子节点列表（precedents 用）：小区域展开单格，大区域折叠为 range 节点
      function refsToChildren(refs, contextSheet) {
        const children = [];
        for (const ref of refs) {
          const childSheet = ref.sheet || contextSheet;
          const key = `${normalizeSheetName(childSheet)}!${ref.ref}`;
          if (visited.has(key)) continue;
          visited.add(key);
          if (ref.ref.includes(':')) {
            const rect = refToRect(ref.ref);
            const size = (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
            if (size <= RANGE_EXPAND_LIMIT) {
              // 展开单格也逐格标 visited：否则 =A1+SUM(A1:B1) 会渲染两份 A1 且子树形态不一致（review P2#1）
              for (let r = rect.r1; r <= rect.r2; r++) for (let c = rect.c1; c <= rect.c2; c++) {
                const cellKey = `${normalizeSheetName(childSheet)}!${a1(r, c)}`;
                if (visited.has(cellKey)) continue;
                visited.add(cellKey);
                children.push({ sheet: childSheet, addr: a1(r, c) });
              }
            } else {
              children.push({ sheet: childSheet, addr: ref.ref, isRange: true });
            }
          } else {
            children.push({ sheet: childSheet, addr: ref.ref });
          }
        }
        return children;
      }

      // 起点：root 的直接引用（precedents）或引用 root 的公式格（dependents）——树节点对象
      let frontier;
      if (mode === 'precedents') {
        frontier = rootFormula ? refsToChildren(extractRefs(rootFormula), rootSheet) : [];
      } else {
        const pr = parseStart(rootAddr);
        frontier = [];
        for (const entry of formulaIndex) {
          const hit = entry.refs.some(ref => {
            const refSheet = ref.sheet ? normalizeSheetName(ref.sheet) : normalizeSheetName(entry.sheet);
            return refSheet === normalizeSheetName(rootSheet) && rectContains(refToRect(ref.ref), pr.startRow, pr.startCol);
          });
          if (hit) {
            const key = `${normalizeSheetName(entry.sheet)}!${entry.addr}`;
            if (!visited.has(key)) { visited.add(key); frontier.push({ sheet: entry.sheet, addr: entry.addr }); }
          }
        }
      }
      for (const n of frontier) n.children = [];
      const roots = frontier.slice(); // BFS 会滚动覆盖 frontier，渲染入口必须是根层

      // 分层批读公式（sync 效率），同时把子节点挂到父节点上构建真树。
      // 展开条件 level < maxDepth：最深一层只读值不展开 children——
      // 展开了也读不到值，渲染成「= 」空值假象（review 批 4 P1#2 实测）
      let level = 1;
      while (frontier.length && level <= maxDepth) {
        const nextFrontier = [];
        if (mode === 'precedents') {
          const groups = new Map();
          for (const n of frontier) {
            if (n.isRange) continue;
            const k = normalizeSheetName(n.sheet);
            if (!groups.has(k)) groups.set(k, { sheetName: n.sheet, nodes: [] });
            groups.get(k).nodes.push(n);
          }
          const readNodes = [];
          for (const g of groups.values()) {
            const sheetObj = sheets.items.find(s => normalizeSheetName(s.name) === normalizeSheetName(g.sheetName));
            if (!sheetObj) continue;
            for (const n of g.nodes) {
              const r = sheetObj.getRange(n.addr);
              r.load('values,formulas');
              readNodes.push({ node: n, r });
            }
          }
          if (readNodes.length) await context.sync();
          for (const item of readNodes) {
            const f = item.r.formulas && item.r.formulas[0] && item.r.formulas[0][0];
            const v = item.r.values && item.r.values[0] && item.r.values[0][0];
            item.node.formula = typeof f === 'string' && f.startsWith('=') ? f : null;
            item.node.value = v;
            if (item.node.formula && level < maxDepth) {
              item.node.children = refsToChildren(extractRefs(item.node.formula), item.node.sheet);
              for (const c of item.node.children) c.children = [];
              nextFrontier.push(...item.node.children);
            } else {
              item.node.children = [];
            }
          }
        } else {
          for (const node of frontier) {
            const entry = formulaIndex.find(e => normalizeSheetName(e.sheet) === normalizeSheetName(node.sheet) && e.addr === node.addr);
            if (!entry) continue;
            node.formula = entry.formula;
            const pr = parseStart(node.addr);
            node.children = [];
            if (level >= maxDepth) continue;
            for (const candidate of formulaIndex) {
              const hit = candidate.refs.some(ref => {
                const refSheet = ref.sheet ? normalizeSheetName(ref.sheet) : normalizeSheetName(candidate.sheet);
                return refSheet === normalizeSheetName(node.sheet) && rectContains(refToRect(ref.ref), pr.startRow, pr.startCol);
              });
              if (!hit) continue;
              const key = `${normalizeSheetName(candidate.sheet)}!${candidate.addr}`;
              if (!visited.has(key)) { visited.add(key); const child = { sheet: candidate.sheet, addr: candidate.addr, children: [] }; node.children.push(child); nextFrontier.push(child); }
            }
          }
        }
        frontier = nextFrontier;
        level++;
      }

      // 递归渲染真树（父子关系由缩进正确表达）
      const renderNode = function (node, depth) {
        const label = `${node.sheet}!${node.addr}${node.isRange ? ' (range)' : node.formula ? ` ${node.formula}` : ` = ${node.value ?? ''}`}`;
        lines.push(`${'  '.repeat(depth)}- ${label}`);
        found++;
        for (const c of node.children || []) renderNode(c, depth + 1);
      };
      if (roots.length) roots.forEach(n => renderNode(n, 1));
      // roots 为空（常量格/无依赖）：不渲染空壳根节点，直接给 No direct found（review P2#4）

      if (!found) lines.push('', `_No direct ${mode} found._`);
      return { success: true, cell: `${rootSheet}!${rootAddr}`, mode, depth: maxDepth, markdown: lines.join('\n'), count: found, _dirtyRanges: [] };
    });
  }

  // ---- get_workbook_overview（批 4）----
  // 无参：工作簿结构总览（sheet 名/位置/尺寸/对象计数 + 命名区域 + 表格）。
  // 传 sheet（表名）：该表详情（尺寸、表格、对象计数、8×10 数据预览）。
  // 与宿主契约的 getMetadata 分工：getMetadata 供界面/会话标签，本工具供模型按需深看。
  async function getWorkbookOverview(args = {}) {
    requireOffice();
    const { sheet: sheetName } = args;
    return Excel.run(async context => {
      const wb = context.workbook;
      wb.load('name');
      const sheets = wb.worksheets;
      sheets.load('items');
      await context.sync();
      for (const s of sheets.items) s.load('id,name,position,visibility');
      await context.sync();
      const map = await getSheetMap(context, sheets.items);

      // 命名区域（工作簿级，API 缺失则跳过）
      const names = [];
      try {
        const nameColl = wb.names;
        if (nameColl && typeof nameColl.load === 'function') {
          nameColl.load('items');
          await context.sync();
          for (const n of nameColl.items) { try { n.load('name,type,value'); } catch { continue; } }
          await context.sync();
          for (const n of nameColl.items) names.push({ name: n.name, type: String(n.type || ''), value: String(n.value || '') });
        }
      } catch {}

      const info = [];
      for (const s of sheets.items) {
        const used = s.getUsedRangeOrNullObject();
        used.load('address,rowCount,columnCount');
        const tables = s.tables;
        if (tables && typeof tables.load === 'function') tables.load('items');
        // 真实 API 是 getCount()（返回 ClientResult，无需 load、sync 后读 .value）；
        // getCountOrNullObject 不存在（review 批 4 实锤：vendored excel-mac 源码 0 命中）
        const chartCount = typeof s.charts.getCount === 'function' ? s.charts.getCount() : null;
        const pivotCount = typeof s.pivotTables.getCount === 'function' ? s.pivotTables.getCount() : null;
        await context.sync();
        if (tables && Array.isArray(tables.items)) for (const t of tables.items) { try { t.load('name,rowCount,columnCount'); } catch { continue; } }
        await context.sync();
        info.push({
          sheetId: Number(map.get(s.id)), name: s.name, position: s.position, visibility: s.visibility,
          usedRange: used.isNullObject ? null : used.address.split('!').pop(),
          rows: used.isNullObject ? 0 : used.rowCount, columns: used.isNullObject ? 0 : used.columnCount,
          tables: tables && Array.isArray(tables.items) ? tables.items.map(t => ({ name: t.name, rows: t.rowCount, columns: t.columnCount })) : [],
          charts: chartCount ? chartCount.value : 0,
          pivotTables: pivotCount ? pivotCount.value : 0
        });
      }

      const result = { success: true, workbookName: wb.name || '', namedRanges: names, worksheets: info };
      if (!sheetName) {
        result.markdown = [
          `**${wb.name || 'Workbook'}** — ${info.length} sheet(s), ${names.length} named range(s)`,
          ...info.map(i => `- [sheet ${i.sheetId}] ${i.name}${i.visibility !== 'Visible' ? ` (${i.visibility})` : ''}${i.usedRange ? ` ${i.usedRange} (${i.rows}×${i.columns})` : ' (空)'}${i.tables.length ? ` · tables: ${i.tables.map(t => t.name).join(',')}` : ''}${i.charts || i.pivotTables ? ` · charts ${i.charts}/pivots ${i.pivotTables}` : ''}`)
        ].join('\n');
        return result;
      }
      const target = info.find(i => normalizeSheetName(i.name) === normalizeSheetName(sheetName));
      if (!target) throw new Error(`Worksheet "${sheetName}" not found. Available: ${info.map(i => i.name).join(', ')}`);
      const sheetObj = sheets.items.find(s => normalizeSheetName(s.name) === normalizeSheetName(sheetName));
      let dataPreview = [];
      try {
        const r = sheetObj.getRange(`A1:${colName(FULL_CONTEXT_COLS - 1)}${FULL_CONTEXT_ROWS}`);
        r.load('values');
        await context.sync();
        dataPreview = r.values.map(row => row.map(csvCell).join(','));
      } catch { dataPreview = ['(预览读取失败)']; }
      return Object.assign(result, { detail: target, dataPreview, sheetId: target.sheetId });
    });
  }

  // ---- conditional_format（批 3；2026-08-31 真机探测 conditionalFormats.add 不可用，
  // 由 capabilityToolGates 在本机自动摘牌，代码保留原生实现供支持环境使用）。
  // API 面按微软官方文档（review 批 3 修正）：公式规则类型是 Custom 而非 Formula，
  // 样式挂在规则子对象的 format 上（cf.custom.format / cf.cellValue.format）。
  function applyCfStyle(format, st) {
    if (!st) return;
    if (st.fill_color) format.fill.color = st.fill_color;
    if (st.font_color) format.font.color = st.font_color;
    if (st.bold !== undefined) format.font.bold = st.bold;
    if (st.italic !== undefined) format.font.italic = st.italic;
  }
  async function conditionalFormat(args) {
    requireOffice();
    const { sheetId, action, range, type, formula, operator, value, style } = args;
    // confirm 闸门：clearAll 抹掉区域内用户已有全部条件格式规则，破坏性与删批注同级
    const accessMode = effectiveAccessMode();
    if (accessMode === 'confirm' && action === 'clear') {
      return { success: false, blocked: true, error: '清除条件格式在确认模式下不可用。请告诉用户：需要清除时切换到「直接修改」模式后重试。' };
    }
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const r = sheet.getRange(range);
      const cfColl = r.conditionalFormats;
      if (!cfColl || (action === 'clear' ? typeof cfColl.clearAll !== 'function' : typeof cfColl.add !== 'function')) {
        throw new Error('conditionalFormats API is not available on this host');
      }
      if (action === 'clear') {
        cfColl.clearAll();
      } else if (action === 'add') {
        let cf;
        if (type === 'formula') {
          if (!formula) throw new Error('formula rules require formula (e.g. "=$A1>0")');
          cf = cfColl.add('Custom');
          cf.custom.rule = { formula: String(formula) };
          applyCfStyle(cf.custom.format, style);
        } else if (type === 'cell_value') {
          if (!operator) throw new Error('cell_value rules require operator (e.g. GreaterThan)');
          cf = cfColl.add('CellValue');
          cf.cellValue.rule = { operator: String(operator), formula1: String(value) };
          applyCfStyle(cf.cellValue.format, style);
        } else {
          throw new Error(`type must be "formula" or "cell_value" (got: ${type})`);
        }
        // 执行层兜底（2026-08-31 真机：代理对象让 typeof 检查假阳性，add 排队后 sync 不报错
        // 但规则并未生效——静默假成功）。sync 后读回规则数验证真的加上了，没加上就明说。
        if (typeof cfColl.getCount === 'function') {
          const count = cfColl.getCount();
          await context.sync();
          if (!count.value) throw new Error('conditionalFormats API accepted the call but no rule was created — this host does not actually support conditional formatting. Tell the user this capability is unavailable and offer manual formatting (format_cells) instead.');
        } else {
          await context.sync();
        }
      } else throw new Error(`Unknown conditional_format action: ${action}`);
      return { success: true, action, sheetId, range, _dirtyRanges: [{ sheetId, range }] };
    });
  }

  async function getAllObjects(args = {}) {
    requireOffice();
    const { sheetId, id: objectId } = args;
    return Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items'); await context.sync();
      for (const s of sheets.items) s.load('id,name'); await context.sync();
      const map = await getSheetMap(context, sheets.items);
      const targetSheets = sheetId ? [await worksheetById(context, sheetId)].filter(Boolean) : sheets.items;
      const objects = [];
      for (const sheet of targetSheets) {
        const charts = sheet.charts; charts.load('items');
        const pivots = sheet.pivotTables; pivots.load('items');
        await context.sync();
        for (const c of charts.items) { c.load('id,name'); await context.sync(); if (!objectId || c.id === objectId) objects.push({ id: c.id, type: 'chart', name: c.name, sheetId: Number(map.get(sheet.id)), sheetName: sheet.name }); }
        for (const p of pivots.items) { p.load('id,name'); await context.sync(); if (!objectId || p.id === objectId) objects.push({ id: p.id, type: 'pivotTable', name: p.name, sheetId: Number(map.get(sheet.id)), sheetName: sheet.name }); }
      }
      return { success: true, objects };
    });
  }

  async function modifyObject(args) {
    requireOffice();
    // 宽进容错（终验：模型首调常把 source/chartType 扁平放顶层而非 properties 内）：
    // 顶层的图表/透视参数并入 properties
    const flatKeys = ['source', 'chartType', 'range', 'anchor', 'title', 'name', 'rows', 'columns', 'values'];
    const merged = Object.assign({}, args.properties || {});
    for (const k of flatKeys) if (args[k] !== undefined && merged[k] === undefined) merged[k] = args[k];
    args = Object.assign({}, args, { properties: merged });
    const { operation, sheetId, objectType, id: objectId, properties = {} } = args;
    // confirm 闸门：删除图表/透视表拦（破坏性），create/update 直通——
    // 创建是低风险可见操作（2026-08-31 终验第 7 步：create 被 confirm 摘牌曾让图表成为死路）
    const accessMode = effectiveAccessMode();
    if (accessMode === 'confirm' && operation === 'delete') {
      return { success: false, blocked: true, error: '删除图表/透视表在确认模式下不可用。请告诉用户：需要删除时切换到「直接修改」模式后重试。' };
    }
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      let result = { success: true, operation, objectType, sheetId };
      if (objectType === 'chart') {
        if (operation === 'create') {
          if (!properties.source || !properties.chartType) throw new Error('Chart creation requires source and chartType');
          const chart = sheet.charts.add(properties.chartType, sheet.getRange(properties.source), Excel.ChartSeriesBy.auto);
          if (properties.name) chart.name = properties.name;
          if (properties.title) { chart.title.text = properties.title; chart.title.visible = true; }
          if (properties.anchor) chart.setPosition(properties.anchor);
          chart.load('id,name'); await context.sync();
          result.id = chart.id; result.name = chart.name;
        } else {
          if (!objectId) throw new Error('Chart update/delete requires id');
          const chart = sheet.charts.getItem(objectId);
          if (operation === 'delete') chart.delete();
          if (operation === 'update') {
            if (properties.title) { chart.title.text = properties.title; chart.title.visible = true; }
            if (properties.name) chart.name = properties.name;
            if (properties.anchor) chart.setPosition(properties.anchor);
          }
        }
      } else if (objectType === 'pivotTable') {
        if (operation === 'create') {
          if (!properties.source || !properties.range) throw new Error('PivotTable creation requires source and range');
          const pivot = sheet.pivotTables.add(properties.name || `Pivot_${Date.now()}`, properties.source, properties.range);
          await context.sync();
          const fieldErrors = await addPivotFields(context, pivot, properties);
          pivot.load('id,name'); await context.sync();
          result.id = pivot.id; result.name = pivot.name;
          if (fieldErrors.length) { result.success = false; result.errors = fieldErrors; }
        } else {
          if (!objectId) throw new Error('PivotTable update/delete requires id');
          const pivot = sheet.pivotTables.getItem(objectId);
          if (operation === 'delete') pivot.delete();
          if (operation === 'update') {
            const fieldErrors = await addPivotFields(context, pivot, properties);
            if (fieldErrors.length) { result.success = false; result.errors = fieldErrors; }
          }
        }
      }
      await context.sync();
      result._dirtyRanges = [{ sheetId, range: properties.range || properties.anchor || '*' }];
      return result;
    });
  }
  async function addPivotFields(context, pivot, p) {
    const errors = [];
    for (const x of (p.rows || [])) {
      try {
        pivot.rowHierarchies.add(pivot.hierarchies.getItem(x.field));
        await context.sync();
      } catch (e) {
        errors.push({ area: 'rows', field: x.field, error: e.message || String(e) });
      }
    }
    for (const x of (p.columns || [])) {
      try {
        pivot.columnHierarchies.add(pivot.hierarchies.getItem(x.field));
        await context.sync();
      } catch (e) {
        errors.push({ area: 'columns', field: x.field, error: e.message || String(e) });
      }
    }
    for (const x of (p.values || [])) {
      try {
        const h = pivot.hierarchies.getItem(x.field);
        const dh = pivot.dataHierarchies.add(h);
        if (x.summarizeBy) dh.summarizeBy = pivotSummarizeBy(x.summarizeBy);
        await context.sync();
      } catch (e) {
        errors.push({ area: 'values', field: x.field, summarizeBy: x.summarizeBy || null, error: e.message || String(e) });
      }
    }
    return errors;
  }
  function pivotSummarizeBy(value) {
    const map = {
      sum: Excel.AggregationFunction && Excel.AggregationFunction.sum || 'Sum',
      count: Excel.AggregationFunction && Excel.AggregationFunction.count || 'Count',
      average: Excel.AggregationFunction && Excel.AggregationFunction.average || 'Average',
      max: Excel.AggregationFunction && Excel.AggregationFunction.max || 'Max',
      min: Excel.AggregationFunction && Excel.AggregationFunction.min || 'Min'
    };
    return map[value] || value;
  }

  async function evalOfficeJs(args) {
    requireOffice();
    const code = args.code || '';
    return Excel.run(async context => {
      const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
      const fn = new AsyncFunction('context', 'Excel', code);
      const result = await fn(context, Excel);
      return { success: true, result: result ?? null, _dirtyRanges: [{ sheetId: -1, range: '*' }] };
    });
  }

  async function maybeFollow(result) {
    if (!App.state.settings.followMode || !result) return;
    const dirty = result._dirtyRanges;
    if (!Array.isArray(dirty) || !dirty.length) return;
    const first = dirty.find(x => x.sheetId && x.sheetId > 0);
    if (first) await selectRange(first.sheetId, first.range === '*' ? undefined : first.range).catch(console.warn);
  }
  async function selectRange(sheetId, range) {
    requireOffice();
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.activate();
      if (range) sheet.getRange(range).select();
      await context.sync();
      return { success: true };
    });
  }
  async function navigateCitation(ref) {
    const [sid, range] = ref.split('!');
    return selectRange(Number(sid), range);
  }

  const TOOL_EXECUTORS = {
    read_memory: (args) => App.memory.readMemoryTool(args),
    write_memory: (args) => App.memory.writeMemoryTool(args),
    read_range: readRange,
    get_cell_ranges: getCellRanges,
    get_range_as_csv: getRangeAsCsv,
    search_data: searchData,
    get_workbook_overview: getWorkbookOverview,
    trace_dependencies: traceDependencies,
    get_all_objects: getAllObjects,
    set_cell_range: setCellRange,
    clear_cell_range: clearCellRange,
    copy_to: copyTo,
    insert_delete_rows_columns: insertDeleteRowsColumns,
    hide_unhide_rows_columns: hideUnhideRowsColumns,
    freeze_panes: freezePanesTool,
    view_settings: viewSettings,
    modify_sheet_structure: modifySheetStructure,
    modify_workbook_structure: modifyWorkbookStructure,
    resize_range: resizeRange,
    format_cells: formatCells,
    comments: commentsTool,
    conditional_format: conditionalFormat,
    instructions: instructionsTool,
    undo_last_write: undoLastWrite,
    modify_object: modifyObject,
    eval_officejs: evalOfficeJs
  };

  const SYSTEM_PROMPT = `你是嵌在 Microsoft Excel 里的中文数据助手。用户是中文母语的专业工作者，你的全部回复都用中文。

工作方式：
1. 先判断用户到底要什么。要求含糊时先问，不要猜着做。
2. 动手前先说明你的整体判断：数据的问题在哪、你打算怎么处理。这段话用自然语言写在正文里，不要塞进工具参数。
3. 然后才提出具体修改。写入目标为空单元格时（建表、填空、追加）可直接调用 set_cell_range 立即生效，不必出提案卡；修改已有内容则必须先用 read_range 读取目标区域的当前值/公式，走修改提案，expectedCells 原样采用读到的内容（公式格填 {formula: "=..."}）——凭对话记忆填快照会因原文已变化被拒。空目标的提案（如需批量审阅时）find 填 "(empty)"，expectedCells 用与 cells 同形状的全 null 矩阵。不要因为目标是空单元格就改用 render 卡片让用户自己导出。
4. 每一处改动都要能追溯到用户的要求，不要顺手改用户没提的地方。
5. Excel 无法用 ⌘Z 撤销插件写入（平台限制）：用户说「撤销/退回」时调用 undo_last_write（每次撤销最近一次写入，可连续调用逐层回退）。

呈现规则（决定用户看到的是卡片还是文字，很重要）：
- 当你的回答是「N 个同构的项」「几个方案并排比较」「行列矩阵」「层级大纲」「需要用户填的表单」时，调用 render 工具，界面会负责编号、对齐、跳转和操作按钮。
- 用了 render 之后，正文里不要再把同样的内容复述一遍，写一句话引出即可。
- 反过来，下面这些情况一律用正文回复，不要调 render：普通问答、解释概念、只有一两句话的回答、单个连续段落的叙述或改写、闲聊与确认。给一段话套上卡片只会让界面变吵，不会让它变清楚。
- 判断标准：需要对着比较、需要逐项操作、或者项与项之间字段相同，就用 render；只是要读一遍，就用正文。


You are an AI assistant integrated into Microsoft Excel with full access to read and modify spreadsheet data.

Available tools:
READ:
- read_range: Read cell values from a range. mode: compact (markdown table + formula sidebar, default) / csv (raw values) / detailed (adds per-cell formulas and number formats). Always read before modifying — never guess what's in the sheet.
- search_data: Find text across the spreadsheet
- get_workbook_overview: Structural overview (sheets, named ranges, tables, object counts); pass a sheet name for detail + preview
- trace_dependencies: Trace formula lineage (precedents/dependents) for a cell like "Sheet1!D10" — the tool for formula debugging
- get_all_objects: List charts, pivot tables, and other objects

WRITE:
- set_cell_range: Write values, formulas, notes, and formatting. After writing, formula results are verified and any Excel errors (#REF!, #DIV/0!, …) are reported back for repair.
- clear_cell_range: Clear contents or formatting
- copy_to: Copy ranges with formula translation
- insert_delete_rows_columns: Insert or delete rows/columns (reference '5'/'C', count, position)
- hide_unhide_rows_columns: Hide or unhide rows/columns
- freeze_panes: freeze_rows/freeze_columns/freeze_at('B3')/unfreeze
- view_settings: gridlines, headings, tab color, sheet visibility
- format_cells: Named styles (header/currency/…) + per-property overrides; multi-area ranges
- comments: Read/add/update/reply/resolve cell comments (notes API is unavailable on this host — always use comments)
- conditional_format: Formula or cell-value rules with fill/font styling
- modify_workbook_structure: Create/delete/rename/duplicate sheets
- resize_range: Adjust column widths and row heights
- modify_object: Create/update/delete charts and pivot tables
- eval_officejs: Execute Office.js code when the listed tools are not enough

Citations: Use markdown links with #cite: hash to reference sheets/cells. Clicking navigates there.
- Sheet only: [Sheet Name](#cite:sheetId)
- Cell/range: [A1:B10](#cite:sheetId!A1:B10)
Example: [Exchange Ratio](#cite:3) or [see cell B5](#cite:3!B5)


When the user asks about their workbook data, read it first. Use A1 notation for cell references.`;

  const EXCEL_EDIT_ITEM = {
    type: 'object',
    properties: {
      label: { type: 'string', description: 'Short name for this suggestion.' },
      find: { type: 'string', description: 'Short description of the original value(s), e.g. "B3: 120". When writing NEW data into empty cells, write "(empty)".' },
      replacement: { type: 'string', description: 'Short description of the new value(s), e.g. "B3: 135", or of the data being written.' },
      target: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, expectedCells: { type: 'array', description: 'Snapshot of current values/formulas before the edit, same 2D shape as cells — MUST come from a fresh read_range, never from memory. For formula cells use {formula: "=..."}; for plain values use {value: ...}. For an empty target use null in every position.', items: { type: 'array', items: { type: 'object' } } }, cells: { type: 'array', description: 'Desired new values/formulas.', items: { type: 'array', items: { type: 'object' } } } }, required: ['sheetId', 'range', 'expectedCells', 'cells'] },
      reasoning: { type: 'string' },
      summary: { type: 'string' }
    },
    required: ['label', 'find', 'replacement', 'target', 'reasoning', 'summary']
  };

  // format_cells 的单边框 schema 片段
  const BORDER_SPEC = { type: 'object', properties: { color: { type: 'string', description: 'Hex, e.g. "#999999"' }, weight: { type: 'string', description: 'Hairline/Thin/Medium/Thick' }, style: { type: 'string', description: "'solid' maps to Continuous" } } };
  const TOOL_DEFINITIONS = [
    { type: 'function', function: { name: 'read_range', description: 'Read cell values from a range. mode "compact" (default) returns a markdown table with a formula sidebar; "csv" returns raw CSV values; "detailed" additionally includes per-cell formulas and number formats. Always read before modifying — never guess what is in the sheet.', parameters: { type: 'object', properties: { sheetId: { type: 'number', description: 'The worksheet ID (1-based index)' }, range: { type: 'string', description: 'Range in A1 notation, e.g. "A1:D10"' }, mode: { enum: ['compact', 'csv', 'detailed'], description: 'Output mode. Default compact' }, maxRows: { type: 'number', description: 'Row limit. Default 500' } }, required: ['sheetId', 'range'] } } },
    { type: 'function', function: { name: 'get_workbook_overview', description: 'Get a structural overview of the workbook: sheet names, dimensions, named ranges, tables, and object counts. Use at the start of a conversation or to locate data before reading ranges. Pass a sheet name for that sheet\'s detail (dimensions, tables, objects, data preview).', parameters: { type: 'object', properties: { sheet: { type: 'string', description: 'Sheet name for detailed info. Omit for the workbook-level overview' } } } } },
    { type: 'function', function: { name: 'trace_dependencies', description: 'READ. Trace formula lineage for a cell: precedents (upstream inputs) or dependents (downstream impact), recursively up to depth (default 2, max 5). Cell must include the sheet name, e.g. "Sheet1!D10". Use for formula debugging and impact analysis before editing.', parameters: { type: 'object', properties: { cell: { type: 'string', description: 'Cell to trace, e.g. "Sheet1!D10". Single cell with sheet name' }, mode: { enum: ['precedents', 'dependents'], description: 'Trace direction. Default precedents' }, depth: { type: 'number', description: 'Recursion levels. Default 2, max 5' } }, required: ['cell'] } } },
    { type: 'function', function: { name: 'search_data', description: 'Find text or values across the spreadsheet. Supports regex and case-sensitive search.', parameters: { type: 'object', properties: { searchTerm: { type: 'string' }, sheetId: { type: 'number' }, range: { type: 'string' }, offset: { type: 'number' }, options: { type: 'object', properties: { matchCase: { type: 'boolean' }, matchEntireCell: { type: 'boolean' }, matchFormulas: { type: 'boolean' }, useRegex: { type: 'boolean' }, maxResults: { type: 'number' } } }, explanation: { type: 'string' } }, required: ['searchTerm'] } } },
    { type: 'function', function: { name: 'get_all_objects', description: 'List all charts, pivot tables, and other objects in the workbook.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, id: { type: 'string' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'set_cell_range', description: "WRITE. Write values, formulas, notes, and formatting to cells. By default fails if target cells (or copyToRange cells) contain data. Retry with allow_overwrite=true after confirmation.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, cells: { type: 'array', items: { type: 'array', items: { type: 'object', properties: { value: {}, formula: { type: 'string' }, note: { type: 'string' }, cellStyles: { type: 'object' }, borderStyles: { type: 'object' } } } } }, copyToRange: { type: 'string' }, resizeWidth: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, resizeHeight: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, allow_overwrite: { type: 'boolean' }, explanation: { type: 'string' } }, required: ['sheetId', 'range', 'cells'] } } },
    { type: 'function', function: { name: 'clear_cell_range', description: "WRITE. Clear contents, formatting, or both from a range. clearType: contents/formats/all.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, clearType: { enum: ['contents', 'formats', 'all'] }, explanation: { type: 'string' } }, required: ['sheetId', 'range'] } } },
    { type: 'function', function: { name: 'copy_to', description: 'WRITE. Copy a range to another location with formula translation. If destination is larger, source pattern repeats. Fails if the destination contains data; confirm with the user and retry with allow_overwrite=true.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, sourceRange: { type: 'string' }, destinationRange: { type: 'string' }, allow_overwrite: { type: 'boolean', description: 'Allow overwriting non-empty cells in the destination. Default false' }, explanation: { type: 'string' } }, required: ['sheetId', 'sourceRange', 'destinationRange'] } } },
    { type: 'function', function: { name: 'insert_delete_rows_columns', description: "WRITE. Insert or delete rows and columns. Use reference like '5' for row 5 or 'C' for column C; count inserts/deletes multiple at once. Deleting is blocked in confirm mode.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, operation: { enum: ['insert', 'delete'] }, dimension: { enum: ['rows', 'columns'] }, reference: { type: 'string' }, count: { type: 'number' }, position: { enum: ['before', 'after'] }, explanation: { type: 'string' } }, required: ['sheetId', 'operation', 'dimension'] } } },
    { type: 'function', function: { name: 'hide_unhide_rows_columns', description: "WRITE. Hide or unhide rows and columns. Use reference like '5' for row 5 or 'C' for column C; count hides/unhides multiple at once.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, operation: { enum: ['hide', 'unhide'] }, dimension: { enum: ['rows', 'columns'] }, reference: { type: 'string' }, count: { type: 'number' }, explanation: { type: 'string' } }, required: ['sheetId', 'operation', 'dimension', 'reference'] } } },
    { type: 'function', function: { name: 'freeze_panes', description: "WRITE. Freeze or unfreeze panes. freeze_rows/freeze_columns take count; freeze_at takes a cell reference like 'B3' (freezes rows above + columns left of it).", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, action: { enum: ['freeze_rows', 'freeze_columns', 'freeze_at', 'unfreeze'] }, count: { type: 'number', description: 'Required for freeze_rows / freeze_columns' }, reference: { type: 'string', description: "Cell reference for freeze_at, e.g. 'B3'" }, explanation: { type: 'string' } }, required: ['sheetId', 'action'] } } },
    { type: 'function', function: { name: 'view_settings', description: 'WRITE. Read or change worksheet display settings: gridlines, row/column headings, tab color, and sheet visibility. Use action "get" to inspect the current state. Hiding sheets is blocked in confirm mode.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, action: { enum: ['get', 'show_gridlines', 'hide_gridlines', 'show_headings', 'hide_headings', 'set_tab_color', 'hide_sheet', 'show_sheet', 'very_hide_sheet'] }, color: { type: 'string', description: 'Tab color in #RRGGBB ("" to clear). Only for set_tab_color' }, explanation: { type: 'string' } }, required: ['sheetId', 'action'] } } },
    { type: 'function', function: { name: 'modify_workbook_structure', description: "WRITE. Create, delete, rename, or duplicate worksheets. Deleting a sheet is irreversible (undo does NOT work): without confirmed=true the call is refused; first ask the user in one sentence which sheet will be deleted and what it contains, then call again with confirmed=true only after explicit approval.", parameters: { type: 'object', properties: { operation: { enum: ['create', 'delete', 'rename', 'duplicate'] }, sheetId: { type: 'number' }, sheetName: { type: 'string' }, newName: { type: 'string' }, tabColor: { type: 'string' }, confirmed: { type: 'boolean', description: 'Required for operation=delete: set true only after the user explicitly approved deleting the sheet' }, explanation: { type: 'string' } }, required: ['operation'] } } },
    { type: 'function', function: { name: 'resize_range', description: "WRITE. Adjust column widths or row heights. Use 'A:D' for columns, '1:5' for rows, or omit range for entire sheet.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, width: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, height: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, explanation: { type: 'string' } }, required: ['sheetId'] } } },
    { type: 'function', function: { name: 'format_cells', description: "WRITE. Apply formatting to a cell range: named styles and individual property overrides. Named styles: header, total-row, subtotal, input, blank-section, number, integer, currency, percent, ratio, text. Range supports comma-separated multi-area: 'A1:B2,D1:E2'. Visible immediately and undo-friendly.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, style: { description: "Named style(s): a single string or an array", anyOf: [{ type: 'string', enum: ['header', 'total-row', 'subtotal', 'input', 'blank-section', 'number', 'integer', 'currency', 'percent', 'ratio', 'text'] }, { type: 'array', items: { type: 'string' } }] }, bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' }, font_color: { type: 'string', description: 'Hex, e.g. "#FF0000"' }, font_size: { type: 'number' }, font_name: { type: 'string' }, fill_color: { type: 'string', description: 'Hex fill color' }, horizontal_alignment: { enum: ['left', 'center', 'right'] }, wrap_text: { type: 'boolean' }, number_format: { type: 'string' }, borders: { type: 'object', properties: { top: BORDER_SPEC, bottom: BORDER_SPEC, left: BORDER_SPEC, right: BORDER_SPEC } }, explanation: { type: 'string' } }, required: ['sheetId', 'range'] } } },
    { type: 'function', function: { name: 'comments', description: "WRITE. Manage cell comments. action=read lists all comments (NO commentId needed); add puts a comment on a single cell (range + content); update/reply/delete/resolve/reopen take commentId (+ content for update/reply). The text parameter is ALWAYS named \"content\". Deleting is blocked in confirm mode.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, action: { enum: ['read', 'add', 'update', 'reply', 'delete', 'resolve', 'reopen'] }, range: { type: 'string', description: "Target cell in A1 notation. Required for add; single cell only" }, commentId: { type: 'string', description: 'Required for update/reply/delete/resolve/reopen (from a previous read)' }, content: { type: 'string', description: 'Comment text. Required for add/update/reply' }, explanation: { type: 'string' } }, required: ['action'] } } },
    { type: 'function', function: { name: 'conditional_format', description: 'WRITE. Add or clear conditional formatting rules. Supports custom formula rules and cell value rules (GreaterThan, LessThan, Between, EqualTo, …) with fill/font styling.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, action: { enum: ['add', 'clear'] }, range: { type: 'string' }, type: { enum: ['formula', 'cell_value'] }, formula: { type: 'string', description: 'Custom formula for formula rules, e.g. "=$A1>0"' }, operator: { type: 'string', description: 'CellValue operator: Between/NotBetween/EqualTo/NotEqualTo/GreaterThan/LessThan/GreaterThanOrEqual/LessThanOrEqual' }, value: { description: 'Comparison value(s) for cell_value rules' }, style: { type: 'object', properties: { fill_color: { type: 'string' }, font_color: { type: 'string' }, bold: { type: 'boolean' }, italic: { type: 'boolean' } } }, explanation: { type: 'string' } }, required: ['action', 'sheetId', 'range'] } } },
    { type: 'function', function: { name: 'modify_object', description: 'WRITE. Create, update, or delete charts and pivot tables.', parameters: { type: 'object', properties: { operation: { enum: ['create', 'update', 'delete'] }, sheetId: { type: 'number' }, objectType: { enum: ['pivotTable', 'chart'] }, id: { type: 'string' }, properties: { type: 'object', properties: { name: { type: 'string' }, source: { type: 'string' }, range: { type: 'string' }, anchor: { type: 'string' }, rows: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' } } } }, columns: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' } } } }, values: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' }, summarizeBy: { enum: ['sum', 'count', 'average', 'max', 'min'] } } } }, title: { type: 'string' }, chartType: { enum: ['columnClustered', 'barClustered', 'line', 'pie', 'scatter', 'area', 'doughnut'] } } }, explanation: { type: 'string' } }, required: ['operation', 'sheetId', 'objectType'] } } },
    { type: 'function', function: { name: 'undo_last_write', description: 'SYSTEM. Undo the most recent cell write from this session (restores the exact previous values/formulas). Excel cannot undo add-in writes with Cmd+Z — this is the replacement. Each call undoes one write, most recent first (up to 20 levels). Formatting and structural operations are NOT covered. Use when the user says 撤销/退回/undo.', parameters: { type: 'object', properties: {} } } },
    { type: 'function', function: { name: 'instructions', description: 'SYSTEM. Update persistent rules for yourself: level=user applies to all files (personal preferences), level=workbook applies to this file only. Rules persist across sessions and are injected into your context automatically. Use append to add, replace to rewrite the full text. Only set rules the user actually asked for.', parameters: { type: 'object', properties: { action: { enum: ['append', 'replace'] }, level: { enum: ['user', 'workbook'] }, content: { type: 'string', description: 'Rule text to save' } }, required: ['action', 'level', 'content'] } } },
    { type: 'function', function: { name: 'eval_officejs', description: 'Execute arbitrary Office.js code in Excel.run. Escape hatch for unsupported operations. Code receives context and Excel.', parameters: { type: 'object', properties: { code: { type: 'string' }, explanation: { type: 'string' } }, required: ['code'] } } },
    { type: 'function', function: { name: 'propose_edits', description: 'PROPOSAL (do not write directly). Use top-level edits for alternative versions of one location, or top-level changes for several independently reviewable locations. Creating NEW data (building a table, filling empty cells) also goes through this tool: set find to "(empty)" and expectedCells to a null matrix of the same shape as cells. target.cells MUST be a two-dimensional array matching the target range shape (one inner array per row, e.g. [[{value: 500}], [{value: 500}]] for A1:A2) — a flat one-dimensional array is rejected. Put the recommendation and rationale in the card before review. When the user declines with a reason, your next reply MUST act on it in the same turn — a revised proposal card that absorbs the feedback, a proposed direction for the user to confirm, or what you will do instead; never just save the feedback to memory or acknowledge it and stop.', parameters: { type: 'object', properties: { edits: { type: 'array', description: 'ONE location, optionally with alternative versions.', items: EXCEL_EDIT_ITEM }, changes: { type: 'array', description: 'MULTIPLE locations in one review batch. Never nest this inside edits.', items: EXCEL_EDIT_ITEM }, explanation: { type: 'string', description: 'Short overall recommendation shown before review.' } }, required: ['explanation'] } } },
    { type: 'function', function: { name: 'read_memory', description: 'SYSTEM. Read your cross-session memory (work habits and preferences learned from past work). Call with no topic to list what is stored; pass topic (key) to read its full content. Different from instructions rules (user-set, auto-injected): memory is what YOU learned and saved. Check the list when the task touches a stored topic (formatting habits, client preferences) — apply it instead of re-deriving from scratch.', parameters: { type: 'object', properties: { topic: { type: 'string', description: 'Memory topic key. Omit to list available topics first.' } } } } },
    { type: 'function', function: { name: 'write_memory', description: 'SYSTEM. Save a stable work habit/preference to cross-session memory so future sessions inherit it (spreadsheet conventions the user corrected you on, style preferences, recurring task patterns). action: create (new topic) / append (small addition, preferred) / update (reorganize one topic) / delete. Use instructions for rules the user explicitly asks to persist; use write_memory for habits YOU noticed. Only save preferences the user explicitly expressed or repeatedly showed; never promote a single occurrence into a rule; timestamp time-sensitive facts; never store secrets or private data. The user can review and delete memory in settings. Optionally set 1-3 triggers (keywords) when the topic binds to a specific document/task type; when the latest user message matches a trigger, that topic is promoted to the top of the memory list.', parameters: { type: 'object', properties: { action: { type: 'string', enum: ['create', 'append', 'update', 'delete'], description: 'create=new topic; append=small addition (preferred); update=rewrite one topic; delete=remove.' }, topic: { type: 'string', description: 'Topic key, snake_case, e.g. sheet_format_habits.' }, title: { type: 'string', description: 'Short title shown in settings (create only).' }, summary: { type: 'string', description: 'One-line summary for the memory list (create/update).' }, content: { type: 'string', description: 'The rule/preference text (create/append/update).' }, triggers: { type: 'array', items: { type: 'string' }, description: 'Keywords binding this topic to a document/task type (e.g. ["报表", "对账"]). Latest-message match pins the topic to the top of the list. Set on create/update; merged on append.' } }, required: ['action', 'topic'] } } }
  ];

  const SAMPLE_ARGS = {
    read_range: { sheetId: 1, range: 'A1:D10', mode: 'compact', maxRows: 500 },
    get_workbook_overview: {},
    trace_dependencies: { cell: '数据表!D10', mode: 'precedents', depth: 2 },
    search_data: { searchTerm: 'keyword', options: { matchCase: false, useRegex: false, maxResults: 100 } },
    get_all_objects: {},
    set_cell_range: { sheetId: 1, range: 'A1:B2', cells: [[{ value: '标题1', cellStyles: { fontWeight: 'bold' } }, { value: '标题2', cellStyles: { fontWeight: 'bold' } }], [{ value: 1 }, { formula: '=A2*2' }]], allow_overwrite: false },
    clear_cell_range: { sheetId: 1, range: 'A1:B2', clearType: 'contents' },
    copy_to: { sheetId: 1, sourceRange: 'A1:B2', destinationRange: 'D1:E2' },
    insert_delete_rows_columns: { sheetId: 1, operation: 'insert', dimension: 'rows', reference: '5', count: 1, position: 'before' },
    hide_unhide_rows_columns: { sheetId: 1, operation: 'hide', dimension: 'rows', reference: '5', count: 1 },
    freeze_panes: { sheetId: 1, action: 'freeze_rows', count: 1 },
    view_settings: { sheetId: 1, action: 'set_tab_color', color: '#134cff' },
    modify_workbook_structure: { operation: 'create', sheetName: 'AI分析结果', tabColor: '#134cff' },
    resize_range: { sheetId: 1, range: 'A:D', width: { type: 'points', value: 90 } },
    format_cells: { sheetId: 1, range: 'A1:D1', style: 'header' },
    comments: { sheetId: 1, action: 'add', range: 'B2', content: '此口径已与财务核对（2026-08）' },
    conditional_format: { sheetId: 1, action: 'add', range: 'B2:B10', type: 'cell_value', operator: 'GreaterThan', value: 100, style: { fill_color: '#FFC7CE', font_color: '#9C0006' } },
    modify_object: { operation: 'create', sheetId: 1, objectType: 'chart', properties: { source: 'A1:B10', chartType: 'columnClustered', anchor: 'E2', title: 'Chart' } },
    eval_officejs: { code: "const range = context.workbook.worksheets.getActiveWorksheet().getRange('A1');\nrange.load('values');\nawait context.sync();\nreturn range.values;" },
    propose_edits: { edits: [{ label: '更新值', find: 'B3: 120', replacement: 'B3: 135', target: { sheetId: 1, range: 'B3', expectedCells: [[{ value: 120 }]], cells: [[{ value: 135 }]] }, reasoning: '按最新报价更新', summary: '价格 120→135' }] }
  };
  function defaultArgsForTool(name) { return App.pretty(SAMPLE_ARGS[name] || {}); }

  // 单元格提案项的取值口径（vet / enrich / applyEdit 三处同源）：target 内嵌或提案项顶层任填其一。
  function proposalCellFields(edit) {
    const target = (edit && edit.target) || {};
    return {
      cells: target.cells ?? (edit ? edit.cells : undefined),
      expectedCells: target.expectedCells ?? (edit ? edit.expectedCells : undefined)
    };
  }

  // 应用用户在 Diff 卡上选中的提案版本（覆盖写入）
  async function applyEdit(edit) {
    const target = edit.target || {};
    const fields = proposalCellFields(edit);
    if (!Array.isArray(fields.expectedCells)) throw App.makeStaleEditError('提案缺少单元格原值快照（target.expectedCells 或提案项顶层 expectedCells 任填其一，必须来自 read_range 的真实读取），请重新读取此项。', { target: { sheetId: target.sheetId, range: target.range } });
    // 【R1/30.2 + 31-B】写入生命周期：进入即计入 pendingDocWrites 并捕获取消代次；
    // 令牌作为本操作的内部参数传入 setCellRange，在实际提交点读取——不用模块级
    // 可覆盖变量（批量逐项应用的异步点击可让两次 applyEdit 重叠，全局令牌会被另一
    // 操作的 finally 清空，丢失本操作的取消身份）。模型经 auto/直通路径直接调
    // set_cell_range 时不传令牌、行为不变。
    const op = App.docWriteBegin('excel.applyEdit');
    try {
      return await TOOL_EXECUTORS.set_cell_range({ sheetId: target.sheetId, range: target.range, cells: fields.cells, expectedCells: fields.expectedCells, allow_overwrite: true }, op.token);
    } finally {
      op.end();
    }
  }

  // 提案层校验（api.js 拦截 propose_edits 时调用）：静态可判定的缺陷在出卡前打回模型重写，
  // 返回 null 表示通过；返回错误串则作为 tool result 回喂。查公式语法与 cells/expectedCells
  // 的二维形状、单格规格形态、与目标区域行列数的一致性。目标是否真实存在属运行时信息，
  // 由 enrichEditProposal 异步读取后核验。
  function parseA1Shape(range) {
    const part = String(range || '').split('!').pop().replace(/\$/g, '').replace(/'/g, '').trim();
    if (!/^[A-Za-z]{1,3}\d+(:[A-Za-z]{1,3}\d+)?$/.test(part)) return null;
    const [a, b] = part.split(':');
    const pa = parseStart(a), pb = parseStart(b || a);
    return { rows: Math.abs(pb.startRow - pa.startRow) + 1, cols: Math.abs(pb.startCol - pa.startCol) + 1 };
  }
  function vetEditProposalArgs(args) {
    const items = []
      .concat(Array.isArray(args && args.edits) ? args.edits : [])
      .concat(Array.isArray(args && args.changes) ? args.changes : []);
    for (let index = 0; index < items.length; index++) {
      const item = items[index];
      const at = `edits/changes[${index}]${item && item.label ? `（${String(item.label).slice(0, 40)}）` : ''}`;
      // cells 形状在出卡前点名（v151 真机：一维 cells 出卡后应用必败，模型盲试）。
      // 取值与 applyEdit 同源（proposalCellFields）：target 内嵌或提案项顶层任填其一。
      const rawCells = proposalCellFields(item).cells;
      if (rawCells != null && (!Array.isArray(rawCells) || !rawCells.length || !rawCells.every(Array.isArray))) {
        return `${at} 的 cells（target.cells 或提案项顶层 cells 任填其一）必须是二维数组（每个内层数组对应目标区域的一行），例如 [[{value:500}],[{value:500}]]。收到的是${Array.isArray(rawCells) ? '扁平数组' : '非数组'}，请修正后重新提交。`;
      }
      const cells = rawCells;
      const range = String((item.target && item.target.range) || '');
      // expectedCells 形状前置校验：扁平数组、畸形规格与行列数不符在这里打回，原先要到
      // 应用期 set_cell_range 才抛「dimensions must match」，必败卡已先摆到用户面前。
      const rawExpected = proposalCellFields(item).expectedCells;
      if (rawExpected != null) {
        if (!Array.isArray(rawExpected) || !rawExpected.length || !rawExpected.every(Array.isArray)) {
          return `${at} 的 expectedCells（target.expectedCells 或提案项顶层 expectedCells 任填其一）必须是二维数组（每个内层数组对应目标区域的一行）。例如 A1:A2 的原值 120、130：expectedCells=[[{value:120}],[{value:130}]]，而不是 [{value:120},{value:130}] 这样的扁平数组；空格填 null，公式格填 {formula:"=..."}。收到的是${Array.isArray(rawExpected) ? '扁平数组' : '非数组'}，请先用 read_range 读取该区域，按真实行列原样填写后重新提交。`;
        }
        // 单格规格形态：{formula:""}、{val:5}、value+formula 同带等畸形规格，此前要到
        // 应用期比较才以失配暴露且证据不可读；这里静态点名并展示实际字段。
        for (let i = 0; i < rawExpected.length; i++) {
          for (let j = 0; j < rawExpected[i].length; j++) {
            const problem = expectedSpecProblem(rawExpected[i][j]);
            if (problem) {
              return `${at} 的 expectedCells 第 ${i + 1} 行第 ${j + 1} 列是无法识别的规格：${problem}。合法形态：{value: 值}、{formula: "=公式"} 或 null（空格填 null，裸数字/文字也兼容）。请先用 read_range 读取该区域，按真实内容原样重填 expectedCells 后重新提交。收到的是 ${clampCellEvidence(JSON.stringify(rawExpected[i][j]))}`;
            }
          }
        }
        const shape = parseA1Shape(range) || (Array.isArray(cells) ? { rows: cells.length, cols: matrixRowWidth(cells) } : null);
        if (shape) {
          const rows = rawExpected.length;
          const cols = matrixRowWidth(rawExpected);
          if (rows !== shape.rows || cols !== shape.cols) {
            return `${at} 的 expectedCells 形状（${rows} 行 × ${cols} 列）与目标区域 ${range || '(未给 range)'}（${shape.rows} 行 × ${shape.cols} 列）不一致。expectedCells 必须与 range 同形状（行数、列数完全相等，空格填 null，公式格填 {formula:"=..."}），例如 B2:C3 要写 [[{value:1},{value:2}],[{value:3},{value:4}]]。请先用 read_range 读取该区域，按返回的真实行列原样填写后重新提交。`;
          }
        }
      }
      if (!Array.isArray(cells)) continue;
      for (let i = 0; i < cells.length; i++) {
        if (!Array.isArray(cells[i])) continue;
        for (let j = 0; j < cells[i].length; j++) {
          const cell = cells[i][j];
          const formula = cell && typeof cell === 'object' && typeof cell.formula === 'string'
            ? cell.formula
            : (typeof cell === 'string' && cell.startsWith('=') ? cell : null);
          const bad = validateFormulaSyntax(formula);
          if (bad) return `提案中的公式语法错误（区域 ${range} 第 ${i + 1} 行第 ${j + 1} 列）：${bad}。请修正公式后重新提交提案，不要原样重发。`;
        }
      }
    }
    return null;
  }

  // ---- 提案出卡前核验（对应 Word 宿主的 enrichEditProposal）----
  // ui.js 出卡前（vet 之后）异步调用。单元格类提案项对 expectedCells 做只读实值核验：
  // 此前实值校验只在用户点「应用」时的 setCellRange 做（vet 只查形状），快照失配的必败卡
  // 会先摆到用户面前。规则：
  //   - 任一格失配即抛错回喂模型（api.js 把 enrich 抛错转 success:false 工具结果，卡不出、零写入）
  //   - 比较口径与 setCellRange 同源（expectedCellMatches / expectedCellSpec），公式比文本、
  //     值比 valuesEquivalent（null/空串互认），绝不因「读不到」视为相符
  //   - 不自动改写模型提交的快照：模型必须按 read_range 证据重新提案
  //   - 缺失快照的项不在此拦（遵从既有 vet 约束；应用时由 applyEdit 报「缺少快照」）
  //   - 形状先按真机行列数核对（vet 的 parseA1Shape 查不出整列/整行等写法），再读值；超过
  //     10000 格的区域直接引导拆分（与 set_cell_range 写入上限同档），避免出卡前物化百万行矩阵
  async function enrichEditProposal(args) {
    const items = []
      .concat(Array.isArray(args && args.changes) ? args.changes : [])
      .concat(Array.isArray(args && args.edits) ? args.edits : []);
    await verifyCellSnapshots(items);
    return args;
  }
  async function verifyCellSnapshots(items) {
    if (typeof Excel === 'undefined') return;
    const jobs = [];
    for (const item of items) {
      if (!item) continue;
      const fields = proposalCellFields(item);
      if (!Array.isArray(fields.expectedCells)) continue;
      const sheetId = item.target && item.target.sheetId;
      const range = String((item.target && item.target.range) || '');
      const label = String(item.label || '');
      // 带快照但目标不完整（sheetId 缺失/''/非数字或 range 空）此前静默跳过，vet 也不拦，
      // 必败卡照出。这里拒绝并指明 get_workbook_overview / read_range。
      if (sheetId == null || sheetId === '' || !Number.isFinite(Number(sheetId)) || !range.trim()) {
        throw notCommittedError(`提案「${label.slice(0, 40)}」带 expectedCells 快照但目标不完整（sheetId=${JSON.stringify(sheetId == null ? null : sheetId)}，range=${JSON.stringify(range)}）。请先用 get_workbook_overview 获取真实 sheetId、read_range 确认目标区域地址，再重新提案。未写入。`);
      }
      // 规格形态运行时兜底（正常链路 vet 已静态拒绝）：无法识别的规格不进比较，直接点名。
      for (let i = 0; i < fields.expectedCells.length; i++) {
        const row = fields.expectedCells[i];
        if (!Array.isArray(row)) continue;   // 非法行形状交给下方真机形状核对统一报错
        for (let j = 0; j < row.length; j++) {
          const problem = expectedSpecProblem(row[j]);
          if (problem) {
            throw notCommittedError(`提案「${label.slice(0, 40)}」的 expectedCells 第 ${i + 1} 行第 ${j + 1} 列是无法识别的规格：${problem}。合法形态：{value: 值}、{formula: "=公式"} 或 null（空格填 null，裸数字/文字也兼容）。请先用 read_range(sheetId=${Number(sheetId)}, range="${range}") 读取真实内容，按原样重填 expectedCells 后重新提案。未写入。`);
          }
        }
      }
      jobs.push({ item, sheetId: Number(sheetId), range, expectedCells: fields.expectedCells });
    }
    if (!jobs.length) return;
    await Excel.run(async context => {
      for (const job of jobs) {
        const label = String(job.item.label || '');
        const sheet = await worksheetById(context, job.sheetId);
        if (!sheet) {
          throw notCommittedError(`提案「${label.slice(0, 40)}」的目标工作表 sheetId ${job.sheetId}（区域 ${job.range}）不存在。先用 get_workbook_overview 核对真实 sheetId，再重新提案。未写入。`);
        }
        // 先只取行列元数据核对形状：整列/整行等 vet 查不出的大范围在这里就地拦下，
        // 不为注定失配的快照物化百万行 values/formulas。
        const meta = sheet.getRange(job.range);
        meta.load('address,rowCount,columnCount');
        await context.sync();
        const rows = job.expectedCells.length;
        if (rows !== meta.rowCount || job.expectedCells.some(row => !Array.isArray(row) || row.length !== meta.columnCount)) {
          const cols = matrixRowWidth(job.expectedCells);
          throw notCommittedError(`提案「${label.slice(0, 40)}」的 expectedCells 形状（${rows} 行 × ${cols} 列）与目标区域 ${job.range} 的真实形状（${meta.rowCount} 行 × ${meta.columnCount} 列）不一致。请先用 read_range(sheetId=${job.sheetId}, range="${job.range}") 按真实行列原样重填 expectedCells 后重新提案。未写入。`);
        }
        const cellCount = meta.rowCount * meta.columnCount;
        if (cellCount > MAX_WRITE_CELLS) {
          throw notCommittedError(`提案「${label.slice(0, 40)}」的目标区域 ${job.range} 有 ${cellCount} 格，超过单次写入/核验上限（${MAX_WRITE_CELLS} 格）。请把修改拆成多个 ≤${MAX_WRITE_CELLS} 格的提案分批提交。未写入。`);
        }
        const r = sheet.getRange(job.range);
        r.load('values,formulas');
        await context.sync();
        const start = parseStart(meta.address);
        const evidence = [];
        let mismatchTotal = 0;
        for (let i = 0; i < meta.rowCount; i++) for (let j = 0; j < meta.columnCount; j++) {
          if (expectedCellMatches(job.expectedCells[i][j], r.values[i][j], r.formulas[i][j])) continue;
          mismatchTotal++;
          if (evidence.length < SNAPSHOT_EVIDENCE_CELL_CAP) {
            evidence.push(expectedCellEvidence(a1(start.startRow + i, start.startCol + j), job.expectedCells[i][j], r.values[i][j], r.formulas[i][j]));
          }
        }
        if (mismatchTotal) {
          throw App.makeStaleEditError(
            staleSnapshotMessage(label, job.sheetId, job.range, evidence, mismatchTotal,
              `请先用 read_range(sheetId=${job.sheetId}, range="${job.range}") 读取真实值/公式，按读到的内容原样重填 expectedCells 后重新提案；不要凭记忆微调快照。`),
            { target: { sheetId: job.sheetId, range: job.range }, mismatchedCells: mismatchTotal, evidence }
          );
        }
      }
    });
  }

  // 真机能力探测（学 PPT 的 HOST_QUIRKS 方法论）：宿主挂载后只读探测一次，
  // 探明本机 Excel 对各 API 的真实支持度。零写入，全部 try/catch，结果挂 host.capabilities。
  // 用途：①「工具」页显示真机能力 ②为假宿主夹具提供"按真机行为复刻"的依据
  async function probeExcelCapabilities() {
    const caps = { probedAt: new Date().toISOString() };
    try {
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load('name,showGridlines,showHeadings,tabColor,visibility');
        const r = sheet.getRange('A1');
        r.load('numberFormat,formulas');
        await context.sync();
        caps.sheetViewProps = true;
        caps.rangeNumberFormat = Array.isArray(r.numberFormat);
      });
    } catch (e) { caps.sheetViewProps = false; caps.error = e.message; }
    try {
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        caps.notesApi = !!(context.workbook.notes && typeof context.workbook.notes.add === 'function');
        caps.freezeAt = !!(sheet.freezePanes && typeof sheet.freezePanes.freezeAt === 'function');
        caps.conditionalFormats = !!(sheet.conditionalFormats && typeof sheet.conditionalFormats.add === 'function');
      });
    } catch (e) { caps.existenceProbeError = e.message; }
    try { caps.comments = await (async () => { const v = Excel.Comment; return typeof v !== 'undefined'; })(); } catch { caps.comments = false; }
    // 插件写入能否被 ⌘Z 撤销。ExcelApi 1.20（2025-09）起宿主才把插件写入计入撤销栈，
    // 在此之前是「任何插件写入清空整个撤销栈」。两个条件缺一不可：
    //   ①宿主够新（Mac 需 Office 16.100+）②加载的 office.js 本身是 1.20 之后的构建。
    // 2026-08-31 的疑难杂症正是栽在②：vendor 冻结在 2022 年的 16.0.15407，
    // 宿主明明支持（同机 ChatGPT/Claude 可撤销），我们的写入却全程不进撤销栈。
    // 这条探针就是为了让「撤销不可用」当场可见，不必再靠真机试 ⌘Z 反推。
    try {
      const req = typeof Office !== 'undefined' && Office.context && Office.context.requirements;
      caps.undoSupported = Boolean(req && req.isSetSupported('ExcelApi', '1.20'));
    } catch (e) { caps.undoSupported = false; caps.undoProbeError = e.message; }
    try { caps.officeJsSource = (typeof window !== 'undefined' && window.__officeJsSource__) || ''; } catch { caps.officeJsSource = ''; }
    return caps;
  }

  // ---- instructions：两级持久规则（批 6；user 级 localStorage 全局、workbook 级按文档存）----
  const USER_RULES_KEY = App.STORAGE_KEYS.userRules;
  async function instructionsTool(args) {
    const { action, level, content } = args;
    if (!['user', 'workbook'].includes(level)) throw new Error(`level must be "user" or "workbook" (got: ${level})`);
    if (!['append', 'replace'].includes(action)) throw new Error(`action must be "append" or "replace" (got: ${action})`);
    const text = String(content || '').trim();
    if (action === 'append' && !text) throw new Error('content is required for append');
    if (action === 'replace' && !text) throw new Error('content is required for replace (to clear, say so to the user instead)');
    if (level === 'user') {
      let current = '';
      try { current = (typeof window !== 'undefined' && window.localStorage ? window.localStorage.getItem(USER_RULES_KEY) : '') || ''; } catch { current = ''; }
      const next = action === 'replace' ? text : (current ? current + '\n' + text : text);
      try { if (typeof window === 'undefined' || !window.localStorage) throw new Error('storage unavailable'); window.localStorage.setItem(USER_RULES_KEY, next); } catch (e) { throw new Error(`保存用户级规则失败：${e.message || e}`); }
      return { success: true, level, action, ruleCount: next.split('\n').filter(Boolean).length };
    }
    let current = '';
    try { current = App.loadDocSetting('workbookRules', '') || ''; } catch { current = ''; }
    const next = action === 'replace' ? text : (current ? current + '\n' + text : text);
    await App.saveDocSetting('workbookRules', next).catch(e => { throw new Error(`保存本文件规则失败：${e.message}`); });
    return { success: true, level, action, ruleCount: next.split('\n').filter(Boolean).length };
  }
  async function getPersistentRules() {
    const out = {};
    try { out.user = (typeof window !== 'undefined' && window.localStorage ? window.localStorage.getItem(USER_RULES_KEY) : '') || ''; } catch { out.user = ''; }
    try { out.workbook = App.loadDocSetting('workbookRules', '') || ''; } catch { out.workbook = ''; }
    return out;
  }

  App.HOSTS.excel = {
    hostType: 'excel',
    available: true,
    metadataLabel: 'Workbook metadata',
    systemPrompt: SYSTEM_PROMPT,
    toolDefinitions: TOOL_DEFINITIONS,
    toolExecutors: TOOL_EXECUTORS,
    defaultArgsForTool,
    evalToolName: 'eval_officejs',
    getPersistentRules,
    getMetadata: getWorkbookMetadata,
    getFullContext,
    navigateCitation,
    follow: maybeFollow,
    applyEdit,
    vetEditProposalArgs,
    enrichEditProposal,
    probeExcelCapabilities,
    // confirm 模式下可直接生效的工具：resize_range 调列宽/行高、insert_delete（delete 分支由执行器
    // 内部闸门单独拦截）、hide_unhide、freeze_panes、view_settings（hide_sheet 分支同拦）、
    // format_cells（纯格式、可 ⌘Z）、comments（低风险说明性写入，delete 分支同拦）、
    // conditional_format（样式规则、可清可撤）、set_cell_range（**空目标直通**——覆盖已有内容
    // 由执行器内闸门拦截强制走提案）、modify_workbook_structure（create/rename/duplicate 直通，
    // delete 由 confirmed 二段闸拦截——终验第 6 步：建分析表曾被摘成死路）、
    // modify_object（create/update 直通，delete 分支闸拦——终验第 7 步图表同理）。
    directFormattingTools: ['resize_range', 'insert_delete_rows_columns', 'hide_unhide_rows_columns', 'freeze_panes', 'view_settings', 'format_cells', 'comments', 'conditional_format', 'set_cell_range', 'modify_workbook_structure', 'modify_object'],
    // 能力闸（对齐 PPT 机制）：2026-08-31 真机探测 conditionalFormats.add 不可用 → 本机自动摘牌，
    // 模型看不到该工具；探测为 null（未测）时保留。runtimeCapabilities 由 app.js 启动探测后填充。
    capabilityToolGates: { conditional_format: 'conditionalFormats' },
    runtimeCapabilities() { return this.capabilities || {}; },
    i18n: {
      zh: {
        brand: 'Trojan AI', brandFooter: 'Trojan AI for Office',
        title: '准备好处理你的 Excel 数据', subtitle: '你可以让我分析、可视化或转换你的数据',
        input: '告诉我你想如何处理这份表格…',
        chart: '智能图表生成', chartDesc: '自动推荐图表类型并一键生成',
        fix: '公式错误诊断', fixDesc: '定位报错原因并生成修复公式',
        analyze: '跨表智能解析', analyzeDesc: '自动关联多表数据并输出结论',
        chartPrompt: '根据当前表格数据生成合适的图表，先核对数据范围和要表达的关系。',
        fixPrompt: '检查当前表格中的公式错误，定位原因并按当前修改模式提出或应用修复。',
        analyzePrompt: '分析相关工作表内容，核对数据范围后给出结论和依据。',
        demo: '当前不在 Excel/Office 环境中，Excel 工具只能在插件侧边栏里运行。'
      },
      en: {
        brand: 'Trojan AI', brandFooter: 'Trojan AI for Office',
        title: 'Ready to work with your Excel data', subtitle: 'Ask me to analyze, visualize, or transform your data',
        input: 'Tell me what to do with this workbook…',
        chart: 'Chart Generation', chartDesc: 'One-click visualization & styling',
        fix: 'Error Fix', fixDesc: 'Auto-detect & fix formula errors',
        analyze: 'Multi-Sheet Analysis', analyzeDesc: 'Cross-sheet automation and conclusions',
        chartPrompt: 'Create a suitable chart after checking the data range and intended comparison.',
        fixPrompt: 'Locate formula errors and propose or apply fixes according to the selected editing mode.',
        analyzePrompt: 'Analyze the relevant sheets and give conclusions supported by the data.',
        demo: 'Not currently running inside Excel/Office. Excel tools only work in the add-in task pane.'
      }
    }
  };
})();
