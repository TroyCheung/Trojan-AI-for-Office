'use strict';
// Excel 出卡前锚点修复（自内部仓 harden-ppt-excel-recovery 的 Excel 部分移植）：
// 单元格提案出卡前对 expectedCells 做只读实值核验。
// 旧缺口：vetEditProposalArgs 只查形状，快照失配要等用户点「应用」才在 setCellRange 爆
// 「目标单元格已变化」——必败卡先摆到用户面前。
// 本文件钉住移植后的合同：出卡前任一格失配即拒（坐标 + 提交期望 + 真实值/公式 + read_range 指引，
// 超长截断、不倾倒矩阵）、零写入、不改写模型快照、apply 期防陈旧核验保留且错误改善。
// 与 setCellRange 的比较口径必须同源（expectedCellSpec/valuesEquivalent），此处同时钉口径。
// 注：内部仓同名列测试中的「结构卡项 enrich」与「propose_edits 工具描述对齐」两条用例
// 依赖公开版尚未携带的特性（structure 卡、描述文案修正），不在本次移植范围，未搬运。
const assert = require('node:assert/strict');
const test = require('node:test');
const { excelFixture } = require('./helpers/fake-excel-host.js');

function storeSnapshot(store) {
  return JSON.stringify([...store.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)));
}

test('出卡前核验：快照相符（值/公式/空值/裸值/数字与字符串数字互认）→ 通过且不改写模型快照', async () => {
  const { App, store } = excelFixture({
    A1: { value: 120 },
    B1: { value: 240, formula: '=A1*2' },
    C1: { value: null },
    D1: { value: 130 },
    E1: { value: 135 }
  });
  const host = App.HOSTS.excel;
  const expectedCells = [[{ value: 120 }, { formula: '=A1*2' }, null, 130, '135']];
  const args = {
    explanation: '核对一致',
    changes: [{
      label: '改 A1', find: 'A1: 120', replacement: 'A1: 121',
      target: { sheetId: 1, range: 'A1:E1', expectedCells, cells: [[{ value: 121 }, { formula: '=A1*2' }, null, 131, 136]] }
    }]
  };
  const before = storeSnapshot(store);
  const out = await host.enrichEditProposal(args);
  assert.equal(out, args, 'ui.js 读返回值：原对象原样返回');
  assert.deepEqual(args.changes[0].target.expectedCells, expectedCells, '不自动改写模型提交的快照');
  assert.equal(args.changes[0].find, 'A1: 120', '单元格项的卡面字段原样保留');
  assert.equal(storeSnapshot(store), before, '核验只读，零写入');
});

test('出卡前核验：任一格值失配 → 拒绝出卡，错误含坐标/提交期望/真实值/read_range 指引，零写入', async () => {
  const { App, store } = excelFixture({ B3: { value: 135 } });
  const host = App.HOSTS.excel;
  const before = storeSnapshot(store);
  const args = {
    explanation: '改价格',
    changes: [{
      label: '改 B3', find: 'B3: 120', replacement: 'B3: 140',
      target: { sheetId: 1, range: 'B3', expectedCells: [[{ value: 120 }]], cells: [[{ value: 140 }]] }
    }]
  };
  let err = null;
  await assert.rejects(
    host.enrichEditProposal(args),
    e => { err = e; return true; }
  );
  assert.equal(err.code, 'STALE_EDIT');
  assert.match(err.message, /目标单元格已变化/, '失配按陈旧快照拒绝');
  assert.match(err.message, /B3/, '错误含坐标');
  assert.match(err.message, /提交期望 值 120/, '错误含提交期望');
  assert.match(err.message, /实际 值 135/, '错误含真实值');
  assert.match(err.message, /read_range\(sheetId=1, range="B3"\)/, '错误含 read_range 指引');
  assert.match(err.message, /未写入/, '错误明示零写入');
  assert.match(err.message, /「改 B3」/, '错误点名提案项 label');
  assert.equal(storeSnapshot(store), before, '拒绝路径零写入');
});

test('出卡前核验：公式格按公式文本比对，失配给出两侧公式', async () => {
  const { App, store } = excelFixture({ C2: { value: 6, formula: '=A1+B1' } });
  const host = App.HOSTS.excel;
  const before = storeSnapshot(store);
  const args = {
    explanation: '改公式',
    changes: [{
      label: '修 C2 公式', find: 'C2 =A1*2', replacement: 'C2 =A1*3',
      target: { sheetId: 1, range: 'C2', expectedCells: [[{ formula: '=A1*2' }]], cells: [[{ formula: '=A1*3' }]] }
    }]
  };
  let err = null;
  await assert.rejects(host.enrichEditProposal(args), e => { err = e; return true; });
  assert.match(err.message, /C2/);
  assert.match(err.message, /提交期望 公式 "=A1\*2"/, '提交期望给出公式原文');
  assert.match(err.message, /公式 "=A1\+B1"/, '真实公式原文一并给出');
  assert.equal(storeSnapshot(store), before, '拒绝路径零写入');
});

test('出卡前核验：空值与数字类型口径 —— null/空串互认，0 与空互不相等，同值公式可提交计算值', async () => {
  const { App } = excelFixture({ A1: { value: 0 }, B1: { value: '' }, C1: { value: 6, formula: '=A1+6' } });
  const host = App.HOSTS.excel;
  // 相符面：B1 空串格提交 null（互认）；C1 公式格提交计算值 6（值口径，同源语义）
  await host.enrichEditProposal({
    changes: [{ label: 'A', target: { sheetId: 1, range: 'B1:C1', expectedCells: [[null, 6]], cells: [[{ value: 1 }, 7]] } }]
  });
  // 失配面：A1 是数字 0，不是空；提交「空」必须拒（0 是真实内容，不得当空格覆盖）
  let err = null;
  await assert.rejects(
    host.enrichEditProposal({
      changes: [{ label: 'B', target: { sheetId: 1, range: 'A1', expectedCells: [[null]], cells: [[{ value: 9 }]] } }]
    }),
    e => { err = e; return true; }
  );
  assert.match(err.message, /A1：提交期望 值 \(空\)，实际 值 0/, '0 与空互不相等且证据可读');
});

test('出卡前核验：cells/expectedCells 放提案项顶层同样核验（相符与不符）', async () => {
  const { App, store } = excelFixture({ A1: { value: 15 }, A2: { value: 16 } });
  const host = App.HOSTS.excel;
  const before = storeSnapshot(store);
  // 相符：顶层 expectedCells 与真实一致 → 通过
  const ok = { changes: [{ label: '顶层数字', target: { sheetId: 1, range: 'A1' }, expectedCells: [[15]], cells: [[16]] }] };
  await host.enrichEditProposal(ok);
  assert.deepEqual(ok.changes[0].expectedCells, [[15]], '顶层快照同样不被改写');
  // 失配：顶层 expectedCells 指错 → 与 target 内嵌同一条拒绝路径
  let err = null;
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ label: '顶层失配', target: { sheetId: 1, range: 'A2' }, expectedCells: [[99]], cells: [[17]] }] }),
    e => { err = e; return true; }
  );
  assert.match(err.message, /A2：提交期望 值 99，实际 值 16/);
  assert.equal(storeSnapshot(store), before, '两条路径均零写入');
});

test('出卡后用户改动：applyEdit 防陈旧核验保留，错误含对照与指引，零写入', async () => {
  const { App, store } = excelFixture({ B3: { value: 120 } });
  const host = App.HOSTS.excel;
  const edit = {
    label: '改 B3', find: 'B3: 120', replacement: 'B3: 140',
    target: { sheetId: 1, range: 'B3', expectedCells: [[{ value: 120 }]], cells: [[{ value: 140 }]] }
  };
  await host.enrichEditProposal({ changes: [edit] });
  store.set('2,1', { value: 135, formula: '' });   // 出卡后、应用前用户手动改了 B3（B3 = 0 基行列 2,1）
  const before = storeSnapshot(store);
  let err = null;
  await assert.rejects(host.applyEdit(edit), e => { err = e; return true; });
  assert.equal(err.code, 'STALE_EDIT', '应用期仍是 STALE_EDIT 合同');
  assert.match(err.message, /B3：提交期望 值 120，实际 值 135/, '应用期错误同样给逐格对照');
  assert.match(err.message, /read_range\(sheetId=1, range="B3"\)/, '应用期错误指引重读');
  assert.match(err.message, /未写入/);
  assert.equal(JSON.stringify(err.currentValues), '[[135]]', 'currentValues 载荷保留（卡面刷新重提路径消费），消息本体不倾倒矩阵');
  assert.equal(err.mismatchedCells, 1, '失配计数进细节');
  assert.equal(storeSnapshot(store), before, '应用期拒绝零写入');
});

test('超长内容显式截断；失配格超过 8 个只列前 8 并给总数，不吐超大 matrix', async () => {
  const long = 'L'.repeat(200) + 'TAIL_END';
  const grid = { B1: { value: long } };
  for (let i = 1; i <= 12; i++) grid['A' + i] = { value: i };
  const { App } = excelFixture(grid);
  const host = App.HOSTS.excel;
  // 超长值：错误含前 60 字符 + 截断标注，不含尾部
  let err = null;
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ label: '长文本', target: { sheetId: 1, range: 'B1', expectedCells: [[{ value: '短' }]], cells: [[{ value: '新' }]] } }] }),
    e => { err = e; return true; }
  );
  assert.ok(err.message.includes('L'.repeat(60)), '截断保留前 60 字符');
  assert.match(err.message, /已截断/, '截断必须显式标注');
  assert.ok(!err.message.includes('TAIL_END'), '不透出超长尾部');
  // 失配 12 格：总数明示，证据行 ≤ 8（A9 起不列）
  let err2 = null;
  await assert.rejects(
    host.enrichEditProposal({
      changes: [{
        label: '整列失配',
        target: { sheetId: 1, range: 'A1:A12', expectedCells: Array.from({ length: 12 }, () => [{ value: 0 }]), cells: Array.from({ length: 12 }, () => [{ value: 1 }]) }
      }]
    }),
    e => { err2 = e; return true; }
  );
  assert.match(err2.message, /失配 12 格/, '总数明示');
  assert.match(err2.message, /仅展示前 8 格/, '封顶必须显式说明只展示前 8 项');
  assert.ok(!err2.message.includes('A9：'), '证据行封顶 8 条');
  assert.match(err2.message, /A8：/, '前 8 条在场');
});

test('缺失快照遵从既有 vet 约束：enrich 不拦、不补写快照', async () => {
  const { App } = excelFixture({ A1: { value: 10 } });
  const host = App.HOSTS.excel;
  const args = { changes: [{ label: '无快照', target: { sheetId: 1, range: 'A1', cells: [[11]] } }] };
  assert.equal(host.vetEditProposalArgs(args), null, 'vet 对缺失快照本就不拦（既有约束）');
  const out = await host.enrichEditProposal(args);
  assert.equal(out, args);
  assert.equal(args.changes[0].target.expectedCells, undefined, '不自动补写快照');
});

test('多表：sheetId 2 的失配按对应表拒绝', async () => {
  const { App } = excelFixture({ sheets: [{ name: '正本', grid: { A1: { value: 1 } } }, { name: '草稿', grid: { A1: { value: 2 } } }] });
  const host = App.HOSTS.excel;
  let err = null;
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ label: '改草稿', target: { sheetId: 2, range: 'A1', expectedCells: [[99]], cells: [[3]] } }] }),
    e => { err = e; return true; }
  );
  assert.match(err.message, /sheetId 2/);
  assert.match(err.message, /A1：提交期望 值 99，实际 值 2/);
});

test('目标表不存在：出卡前拒绝并指向 get_workbook_overview，不静默出必败卡', async () => {
  const { App } = excelFixture({ A1: { value: 1 } });
  const host = App.HOSTS.excel;
  let err = null;
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ label: '找不到表', target: { sheetId: 99, range: 'A1', expectedCells: [[1]], cells: [[2]] } }] }),
    e => { err = e; return true; }
  );
  assert.match(err.message, /sheetId 99/);
  assert.match(err.message, /get_workbook_overview/);
});

test('真实形状与快照形状不符：出卡前按真机读数拒绝并给 read_range 指引', async () => {
  const { App } = excelFixture({ A1: { value: 1 }, A2: { value: 2 } });
  const host = App.HOSTS.excel;
  let err = null;
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ label: '形状错', target: { sheetId: 1, range: 'A1:A2', expectedCells: [[1]], cells: [[9], [9]] } }] }),
    e => { err = e; return true; }
  );
  assert.match(err.message, /1 行 × 1 列/, '提交形状');
  assert.match(err.message, /2 行 × 1 列/, '真实形状');
  assert.match(err.message, /read_range/);
});

test('独立审查1：带快照但目标不完整（sheetId 缺失/空/非数字、range 空）→ 出卡前拒绝并指路，零写入', async () => {
  const badTargets = [
    { range: 'A1' },          // 缺 sheetId
    { sheetId: '', range: 'A1' },
    { sheetId: 'abc', range: 'A1' },
    { sheetId: 1, range: '' },
    { sheetId: 1 }            // 缺 range
  ];
  for (const target of badTargets) {
    const { App, store } = excelFixture({ A1: { value: 1 } });
    const before = storeSnapshot(store);
    let err = null;
    await assert.rejects(
      App.HOSTS.excel.enrichEditProposal({
        changes: [{ label: '目标不完整', target: Object.assign({}, target, { cells: [[2]], expectedCells: [[1]] }) }]
      }),
      e => { err = e; return true; }
    );
    assert.equal(err.writeState, 'not_committed');
    assert.match(err.message, /目标不完整/);
    assert.match(err.message, /get_workbook_overview/, '指明表身份来源');
    assert.match(err.message, /read_range/, '指明区域读取动作');
    assert.match(err.message, /未写入/);
    assert.equal(storeSnapshot(store), before, '拒绝路径零写入');
  }
});

test('独立审查2：畸形 expectedCells 规格（{formula:""}/{formula:"SUM(A1)"}/{val:5}/value+formula 同带）vet 静态拒绝且字段可读', () => {
  const { App } = excelFixture({ A1: { value: 1 } });
  const host = App.HOSTS.excel;
  const cases = [
    [{ formula: '' }, /没有可识别|formula 必须/],
    [{ formula: 'SUM(A1)' }, /formula 必须以 "=" 开头.*"SUM\(A1\)"/s],
    [{ val: 5 }, /没有可识别的 value \/ formula 字段.*"val"/s],
    [{ value: 1, formula: '=A1' }, /同时带 value 与 formula.*二选一/s]
  ];
  for (const [bad, pattern] of cases) {
    const verdict = host.vetEditProposalArgs({
      changes: [{ label: '畸形规格', target: { sheetId: 1, range: 'A1', cells: [[2]], expectedCells: [[bad]] } }]
    });
    assert.ok(typeof verdict === 'string' && pattern.test(verdict), `应静态拒绝 ${JSON.stringify(bad)}：${verdict}`);
    assert.match(verdict, /合法形态：\{value: 值\}、\{formula: "=公式"\} 或 null/);
    assert.match(verdict, /read_range/);
    assert.ok(verdict.includes(JSON.stringify(bad).slice(0, 20)), '展示实际收到的 JSON 字段');
  }
  // 合法形态不误伤：{value:null}/{value:''}/{formula:'=A1'}/null/裸值
  for (const good of [[{ value: null }], [{ value: '' }], [{ formula: '=A1' }], [null], [120]]) {
    assert.equal(host.vetEditProposalArgs({
      changes: [{ label: '合法', target: { sheetId: 1, range: 'A1', cells: [[2]], expectedCells: [good] } }]
    }), null, `合法规格不应被拦：${JSON.stringify(good)}`);
  }
});

test('独立审查2：enrich 对畸形规格运行时兜底拒绝（零写入）；失配证据展示 JSON 字段而非 [object Object]', async () => {
  const { App, store } = excelFixture({ A1: { value: 1 } });
  const host = App.HOSTS.excel;
  const before = storeSnapshot(store);
  // enrich 兜底（绕过 vet 的直调路径）
  let err = null;
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ label: '兜底', target: { sheetId: 1, range: 'A1', cells: [[2]], expectedCells: [[{ val: 5 }]] } }] }),
    e => { err = e; return true; }
  );
  assert.equal(err.writeState, 'not_committed');
  assert.match(err.message, /无法识别的规格.*"val"/s);
  assert.match(err.message, /合法形态：\{value: 值\}、\{formula: "=公式"\} 或 null/);
  assert.equal(storeSnapshot(store), before, '兜底拒绝零写入');
  // 直调 set_cell_range（不经 vet/enrich）：畸形规格按失配拒绝，证据展示 JSON 字段
  let err2 = null;
  await assert.rejects(
    host.toolExecutors.set_cell_range({ sheetId: 1, range: 'A1', cells: [[2]], expectedCells: [[{ val: 5 }]], allow_overwrite: true }),
    e => { err2 = e; return true; }
  );
  assert.equal(err2.code, 'STALE_EDIT');
  assert.ok(err2.message.includes('"val":5'), `证据含实际 JSON 字段：${err2.message}`);
  assert.ok(!err2.message.includes('[object Object]'), '不得显示 [object Object]');
  assert.equal(storeSnapshot(store), before, '两条拒绝路径均零写入');
});
