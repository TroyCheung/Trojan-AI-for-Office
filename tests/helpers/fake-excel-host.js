'use strict';
// Excel 仿真机夹具（公开仓精简版）：从内部仓 tests/fake-excel-host.js 裁剪出
// excel-anchor-recovery.test.js 走到的 API 面——worksheetById 的 sheetMap 稳定映射、
// range 的 values/formulas 读取与 formulas 写入落库、提案核验的元数据读取。
// 写入语义同源：formulas 赋值落库并做简单求值模拟——公式含 /0 产生 '#DIV/0!'、
// 含 REF 产生 '#REF!'、其余公式的计算值按 null 处理（不模拟真实计算）。
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ROOT = path.resolve(__dirname, '..', '..');

function load(relativePath, sandbox) {
  const source = fs.readFileSync(path.join(ROOT, relativePath), 'utf8');
  vm.runInNewContext(source, sandbox, { filename: relativePath });
}

function a1ToPos(ref) {
  const s = String(ref).trim();
  const m = s.match(/^([A-Z]+)(\d+)$/i);
  if (m) {
    const col = m[1].toUpperCase().split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    return { row: Number(m[2]) - 1, col };
  }
  if (/^\d+$/.test(s)) return { row: Number(s) - 1, col: 0 }; // 整行引用 '5'
  const col = s.toUpperCase().split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return { row: 0, col }; // 整列引用 'C'
}

function parseRangeAddress(address) {
  const part = String(address).split('!').pop().replace(/'/g, '');
  const [a, b] = part.split(':');
  const start = a1ToPos(a);
  const end = b ? a1ToPos(b) : start;
  return { startRow: start.row, startCol: start.col, rowCount: end.row - start.row + 1, columnCount: end.col - start.col + 1 };
}

// 写入后的求值模拟：只模拟错误类结果，正常公式一律 null（见文件头注释）
function simulateEvaluate(entry) {
  if (typeof entry !== 'string' || !entry.startsWith('=')) return entry;
  if (/\/\s*0(?![.\d])/.test(entry) || /DIV\/0/.test(entry)) return '#DIV/0!';
  if (/REF!/.test(entry)) return '#REF!';
  return null;
}

// 两种形状：
//   excelFixture(grid)                —— 单表（表名「数据表」）
//   excelFixture({ sheets: [{ name, grid }] }) —— 多表
function excelFixture(input = {}) {
  const fixtureDocSettings = {}; // 每次 fixture 独立，防跨测试泄漏
  const multi = input && Array.isArray(input.sheets);
  const sheetSpecs = multi
    ? input.sheets.map((s, i) => ({ name: s.name || ('Sheet' + (i + 1)), grid: s.grid || {} }))
    : [{ name: '数据表', grid: input || {} }];
  const state = { settings: { followMode: false }, workbookId: 'wb' };
  const sheets = [];
  const stores = [];
  for (const spec of sheetSpecs) {
    const store = new Map();
    for (const [ref, data] of Object.entries(spec.grid)) {
      const { row, col } = a1ToPos(ref);
      store.set(`${row},${col}`, Object.assign({}, data));
    }
    stores.push(store);
    const sheetName = spec.name;
    const sheet = {
      id: `native-${sheets.length + 1}-${sheetName}`,
      name: sheetName,
      position: sheets.length,
      visibility: 'Visible',
      showGridlines: true,
      showHeadings: true,
      tabColor: null,
      load() {},
      getRange(address) {
        const { startRow, startCol, rowCount, columnCount } = parseRangeAddress(address);
        const readGrid = prop => {
          const fallback = prop === 'formula' ? '' : null;
          const out = [];
          for (let r = 0; r < rowCount; r++) {
            const row = [];
            for (let c = 0; c < columnCount; c++) {
              const cell = store.get(`${startRow + r},${startCol + c}`) || {};
              row.push(cell[prop] ?? fallback);
            }
            out.push(row);
          }
          return out;
        };
        const range = {
          address: `'${sheetName}'!${address}`,
          rowCount,
          columnCount,
          format: {
            font: { bold: null, italic: null, underline: null, color: null, size: null, name: null, load() {} },
            fill: { color: null, load() {} },
            horizontalAlignment: null, wrapText: null,
            load() {},
            borders: { getItem: () => ({}) }
          },
          load() {},
          getCell(r, c) {
            const data = store.get(`${startRow + r},${startCol + c}`) || {};
            return {
              numberFormat: data.numberFormat || 'General',
              load() {},
              format: {
                font: Object.assign({ name: null, size: null, color: null, bold: null, italic: null, underline: null, strikethrough: null }, data.font || {}, { load() {} }),
                fill: Object.assign({ color: null }, data.fill || {}, { load() {} }),
                horizontalAlignment: data.horizontalAlignment || null,
                load() {},
                borders: { getItem: () => ({}) }
              }
            };
          },
          getEntireColumn() { return { format: {} }; },
          getEntireRow() { return { format: {} }; }
        };
        Object.defineProperty(range, 'values', { get: () => readGrid('value') });
        Object.defineProperty(range, 'formulas', {
          get: () => readGrid('formula'),
          set: matrix => {
            for (let r = 0; r < matrix.length; r++) {
              for (let c = 0; c < matrix[r].length; c++) {
                const entry = matrix[r][c];
                const isFormula = typeof entry === 'string' && entry.startsWith('=');
                store.set(`${startRow + r},${startCol + c}`, { value: simulateEvaluate(entry), formula: isFormula ? entry : '' });
              }
            }
          }
        });
        Object.defineProperty(range, 'text', { get: () => readGrid('value').map(row => row.map(v => v == null ? '' : String(v))) });
        Object.defineProperty(range, 'numberFormat', { get: () => readGrid('numberFormat') });
        return range;
      },
      getUsedRangeOrNullObject() {
        if (!store.size) return { isNullObject: true, address: '', rowCount: 0, columnCount: 0, load() {} };
        let maxR = 0, maxC = 0;
        for (const key of store.keys()) { const [r, c] = key.split(',').map(Number); if (r > maxR) maxR = r; if (c > maxC) maxC = c; }
        const colNameOf = index => { let s = '', n = index; while (n >= 0) { s = String.fromCharCode(n % 26 + 65) + s; n = Math.floor(n / 26) - 1; } return s; };
        const readUsed = prop => {
          const out = [];
          for (let r = 0; r <= maxR; r++) {
            const row = [];
            for (let c = 0; c <= maxC; c++) { const cell = store.get(`${r},${c}`) || {}; row.push(cell[prop] ?? (prop === 'formula' ? '' : null)); }
            out.push(row);
          }
          return out;
        };
        return { isNullObject: false, address: `'${sheetName}'!A1:${colNameOf(maxC)}${maxR + 1}`, rowCount: maxR + 1, columnCount: maxC + 1, get values() { return readUsed('value'); }, get formulas() { return readUsed('formula'); }, load() {} };
      }
    };
    sheets.push(sheet);
  }

  const context = {
    workbook: {
      load() {},
      getSelectedRange: () => ({ address: 'A1', load() {} }),
      worksheets: { items: sheets, load() {}, getActiveWorksheet: () => sheets[0] }
    },
    async sync() {}
  };

  const App = {
    HOSTS: {},
    state,
    STORAGE_KEYS: { sheetMap: 'sheetMap' },
    requireOffice() {},
    hasOffice: () => true,
    loadDocSetting: (k, d) => (k in fixtureDocSettings ? fixtureDocSettings[k] : (d ?? '{}')),
    saveDocSetting: async (k, v) => { fixtureDocSettings[k] = v; },
    pretty: value => JSON.stringify(value),
    valuesEquivalent: (a, b) => a === b,
    makeStaleEditError(message) { const e = new Error(message); e.code = 'STALE_EDIT'; return e; },
    currentAccessMode() { return this.state.settings.accessMode || 'confirm'; }
  };
  const Excel = {
    run: async callback => callback(context),
    RangeCopyType: { all: 'all' },
    ClearApplyTo: { contents: 'contents', formats: 'formats', all: 'all' },
    InsertShiftDirection: { down: 'down', right: 'right' },
    DeleteShiftDirection: { up: 'up', left: 'left' },
    WorksheetPositionType: { after: 'after' }
  };
  // 用户级持久规则（instructions 工具）用的 storage 桩
  const ruleStore = {};
  const localStorageStub = {
    getItem: k => ruleStore[k] ?? null,
    setItem: (k, v) => { ruleStore[k] = String(v); },
    removeItem: k => { delete ruleStore[k]; }
  };
  load('app/assets/js/host.js', { window: { App, localStorage: localStorageStub }, Excel, console });
  // host.js 会把 saveDocSetting/loadDocSetting 覆盖为需要 Office 环境的真实现：
  // 夹具恢复自己的桩（文档设置走 fixtureDocSettings），requireOffice/hasOffice 保留共享实现
  App.saveDocSetting = async (k, v) => { fixtureDocSettings[k] = v; };
  App.loadDocSetting = (k, d) => (k in fixtureDocSettings ? fixtureDocSettings[k] : (d ?? '{}'));
  load('app/assets/js/host-excel.js', { window: { App, localStorage: localStorageStub }, Excel, console });
  // sheet 指第一张表；sheets/stores 供多表用例按表序取仓库
  return { App, state, sheet: sheets[0], store: stores[0] || new Map(), sheets, stores };
}

module.exports = { excelFixture };
