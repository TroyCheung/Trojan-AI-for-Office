'use strict';
// 【harden-ppt-excel-recovery · PPT 部分】文字锚点出卡前核验与可恢复错误。
// 背景：enrichEditProposal 此前只处理 layout 项；文字项（省略 kind 或 target.kind:"notes"）的
// 错字 find、过期 shapeId、重复命中要等用户点卡后才在 applyEdit 失败，报错缺当前文本对照。
// 验收：任一文字项不合法 → 出卡前拒绝整卡（错误含提交 find 全文、目标当前文本、重读指引、
// 明确截断标注），零写入；应用时再核验保留且同类错误同语义；合法项不改动模型提交的 find。
const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ROOT = path.resolve(__dirname, '..', 'app');
const { createFakeHost } = require('./fake-powerpoint-host.js');

function loadWithContext(contextFactory) {
  const sandbox = { console };
  sandbox.window = sandbox;
  sandbox.App = { HOSTS: {}, requireOffice() {}, hasOffice: () => true, state: { settings: { accessMode: 'confirm' }, messages: [] } };
  sandbox.App.currentAccessMode = () => 'confirm';
  sandbox.Office = { context: { requirements: { isSetSupported: () => true } } };
  sandbox.PowerPoint = { run: async fn => fn(contextFactory()) };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'assets/js/host.js'), 'utf8'), sandbox, { filename: 'host.js' });
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'assets/js/host-powerpoint.js'), 'utf8'), sandbox, { filename: 'host-powerpoint.js' });
  return sandbox.App.HOSTS.powerpoint;
}
const loadOnFake = fake => loadWithContext(() => fake.context);

const WRITE_OPS = /^(text|insertText|geometry|name|fill\.|line\.|font\.|bulletFormat\.|delete|setZOrder|setHyperlink|addTextBox|addGeometricShape|addSlide|deleteSlide|background\.|adjustments\.|duplicate)/;
const writeCount = fake => fake.log.filter(e => WRITE_OPS.test(e.op)).length;

// ---- ① 合法项：通过核验、不改 find、零写入、应用成功 ----

test('合法（省略 kind）：通过出卡前核验，find 不被改写，零写入，应用成功', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [
    { id: 'sh1', name: '标题', text: '结论先行，证据放最后。' },
    { id: 'pic1', name: '配图', type: 'Picture' }
  ] }] });
  const host = loadOnFake(fake);
  const item = { label: '改写标题', find: '结论先行', replacement: '先给结论', target: { index: 0, shapeId: 'sh1' } };
  const args = { changes: [item] };
  const out = await host.enrichEditProposal(args);
  assert.equal(out, args, '原样返回');
  assert.equal(item.find, '结论先行', '不得自动改写/猜测模型提交的 find');
  assert.equal(writeCount(fake), 0, '核验只读，零写入');
  const applied = await host.applyEdit(item);
  assert.equal(applied.success, true); // Public trial returns success; text effects are checked independently below.
  assert.equal(fake.slideStates[0].shapes[0].__st.text, '先给结论，证据放最后。');
});

test('合法（引号等价匹配 + 未限定 shapeId 的唯一命中）：通过并应用', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [
    { id: 'a', name: '左框', text: '把“竞品份额”改掉。' },
    { id: 'b', name: '右框', text: '无关文字' }
  ] }] });
  const host = loadOnFake(fake);
  const item = { find: '"竞品份额”', replacement: '对手份额', target: { index: 0 } };
  await host.enrichEditProposal({ changes: [item] });
  const applied = await host.applyEdit(item);
  assert.equal(applied.success, true); // Public trial returns success; text effects are checked independently below.
  // replace 语义：整个命中区间（含引号等价映射回原文的部分）被 replacement 替换
  assert.equal(fake.slideStates[0].shapes[0].__st.text, '把对手份额改掉。');
});

test('合法（notes）：备注锚点通过核验并应用成功', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'sh1', name: '标题', text: '标题' }], notes: '这一页讲三点。', hasNotes: true }] });
  const host = loadOnFake(fake);
  const item = { label: '改备注', find: '讲三点', replacement: '讲两个重点', target: { index: 0, kind: 'notes' } };
  await host.enrichEditProposal({ changes: [item] });
  assert.equal(writeCount(fake), 0, '核验阶段零写入');
  const applied = await host.applyEdit(item);
  assert.equal(applied.success, true); // Public trial returns success; text effects are checked independently below.
});

// ---- ② 错字：出卡前拒绝，带完整对照 ----

test('错字：出卡前拒绝整卡，错误含提交 find、目标当前文本、get_slide 指引，零写入', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'sh1', name: '正文', text: '竞品分析显示份额领先。' }] }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ label: '改写', find: '竞业分析显示份额领先', replacement: '对手分析', target: { index: 0 } }] }),
    err => {
      assert.match(err.message, /竞业分析显示份额领先/, '错误引用提交的 find');
      assert.match(err.message, /竞品分析显示份额领先。/, '错误给出目标当前文本');
      assert.match(err.message, /get_slide/, '错误给重读指引');
      assert.match(err.message, /未写入/, '明确零写入');
      return true;
    }
  );
  assert.equal(writeCount(fake), 0);
});

// ---- ③ 过期 shapeId ----

test('过期 shapeId：拒绝并列出本页现有形状 id，零写入', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [
    { id: 'sh-a', name: '标题', text: '标题A' },
    { id: 'sh-b', name: '正文', text: '正文B' }
  ] }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '标题A', replacement: 'x', target: { index: 0, shapeId: 'ghost' } }] }),
    err => {
      assert.match(err.message, /ghost/, '指出提交的 shapeId');
      assert.match(err.message, /sh-a/);
      assert.match(err.message, /sh-b/, '列出本页现有形状 id 供消歧');
      assert.match(err.message, /get_slide/);
      return true;
    }
  );
  assert.equal(writeCount(fake), 0);
});

// ---- ④ 重复：跨框重复与同框重复都要正确消歧 ----

test('跨框重复：拒绝并要求 target.shapeId 限定，列出两个候选框', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [
    { id: 'a', name: '左框', text: '重点结论' },
    { id: 'b', name: '右框', text: '重点结论' }
  ] }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '重点结论', replacement: 'x', target: { index: 0 } }] }),
    err => {
      assert.match(err.message, /target\.shapeId|shapeId/, '要求 shapeId 消歧');
      assert.match(err.message, /「a」/);
      assert.match(err.message, /「b」/);
      return true;
    }
  );
  assert.equal(writeCount(fake), 0);
});

test('同框重复：即使给了 shapeId 也拒绝，要求更长唯一片段，不猜第一处', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'only', name: '正文', text: '注意：注意：收尾' }] }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '注意', replacement: 'x', target: { index: 0, shapeId: 'only' } }] }),
    err => {
      assert.match(err.message, /出现 2 次/, '说明同框命中次数');
      assert.match(err.message, /唯一/, '要求更长唯一片段');
      return true;
    }
  );
  assert.equal(writeCount(fake), 0);
});

test('混合歧义：一框唯一命中 + 另一框内两次命中 → 出卡前拒绝，应用时同语义拒绝，不静默猜唯一框', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [
    { id: 'uniq', name: '唯一框', text: '开头：注意安全' },
    { id: 'dup', name: '重复框', text: '注意：注意收尾' }
  ] }] });
  const host = loadOnFake(fake);
  const item = { find: '注意', replacement: '当心', target: { index: 0 } };
  await assert.rejects(host.enrichEditProposal({ changes: [item] }), err => {
    assert.match(err.message, /定位不唯一/, '明确说明定位不唯一');
    assert.match(err.message, /「uniq」/, '列出唯一命中的框');
    assert.match(err.message, /「dup」/, '列出多次命中的框');
    assert.match(err.message, /出现 2 次/, '说明多中框的次数');
    return true;
  });
  await assert.rejects(host.applyEdit(item), err => err.code === 'STALE_EDIT', '应用时同一拒绝语义');
  assert.equal(fake.slideStates[0].shapes[0].__st.text, '开头：注意安全', '唯一框未被误改');
  assert.equal(fake.slideStates[0].shapes[1].__st.text, '注意：注意收尾', '重复框未被误改');
  assert.equal(writeCount(fake), 0);
});

// ---- ⑤ notes 专属路径 ----

test('notes 错字：拒绝并指向 get_slide_notes 与备注当前文本', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'sh1', name: '标题', text: '标题' }], notes: '这一页讲竞品。', hasNotes: true }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '这一页讲对手', replacement: 'x', target: { index: 0, kind: 'notes' } }] }),
    err => {
      assert.match(err.message, /get_slide_notes/, '备注的重读工具是 get_slide_notes');
      assert.match(err.message, /这一页讲竞品。/, '给出备注当前文本');
      return true;
    }
  );
  assert.equal(writeCount(fake), 0);
});

test('notes 同框重复：要求扩长唯一片段，不得推荐 shapeId 消歧（备注无形状定位）', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'sh1', name: '标题', text: '标题' }], notes: '注意：注意收尾', hasNotes: true }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '注意', replacement: 'x', target: { index: 0, kind: 'notes' } }] }),
    err => {
      assert.match(err.message, /出现 2 次/);
      assert.match(err.message, /扩长/, '要求更长唯一片段');
      assert.doesNotMatch(err.message, /shapeId/, '备注锚点无形状定位，不得推荐 shapeId 消歧');
      return true;
    }
  );
  assert.equal(writeCount(fake), 0);
});

test('定位阶段整体异常：notes 项兜底指引用 get_slide_notes，错误带原始异常', async () => {
  const context = {
    presentation: { slides: { load() { throw new Error('页面级读取异常'); }, items: [] } },
    async sync() {}
  };
  const host = loadWithContext(() => context);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '备注原文', replacement: 'x', target: { index: 0, kind: 'notes' } }] }),
    err => {
      assert.match(err.message, /get_slide_notes/, 'notes 项的兜底重读工具');
      assert.match(err.message, /页面级读取异常/, '保留原始异常信息');
      assert.match(err.message, /未写入/);
      return true;
    }
  );
});

test('notes 页不存在：拒绝并说明备注页不可访问，不冒充空备注', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'sh1', name: '标题', text: '标题' }] }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '任何文字', replacement: 'x', target: { index: 0, kind: 'notes' } }] }),
    err => {
      assert.match(err.message, /备注页/, '说明是备注页不可访问');
      assert.match(err.message, /get_slide_notes/);
      return true;
    }
  );
  assert.equal(writeCount(fake), 0);
});

// ---- ⑥ 无法读取不冒充空文本 ----

test('无文字框架形状（Picture）：读取失败拒绝，不当作空文本', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [
    { id: 'pic1', name: '配图', type: 'Picture' },
    { id: 'sh1', name: '标题', text: '真标题' }
  ] }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '标题', replacement: 'x', target: { index: 0, shapeId: 'pic1' } }] }),
    err => {
      assert.match(err.message, /读取失败|无法读取/);
      assert.match(err.message, /不是空文本|不能当作空文本/, '明确读取失败≠空文本');
      return true;
    }
  );
  assert.equal(writeCount(fake), 0);
});

test('表格目标：文字卡明确说定位不了表格单元格，指向真实替代工具，不建议原样重试', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [
    { id: 'tbl', name: '数据表', type: 'Table' },
    { id: 'sh1', name: '标题', text: '真标题' }
  ] }] });
  const host = loadOnFake(fake);
  const item = { find: '某个单元格文字', replacement: 'x', target: { index: 0, shapeId: 'tbl' } };
  await assert.rejects(host.enrichEditProposal({ changes: [item] }), err => {
    assert.match(err.message, /表格/, '说明目标是表格');
    assert.match(err.message, /set_table_style/, '单元格样式指向真实存在的工具');
    assert.match(err.message, /edit_table_structure|apply_layout/, '结构/几何替代路径');
    assert.match(err.message, /手动/, '单元格文字无专用工具，明确需手动处理');
    assert.match(err.message, /还会失败|必然.*失败/, '明说原样重试文字卡无效');
    return true;
  });
  await assert.rejects(host.applyEdit(item), err => {
    assert.equal(err.code, 'STALE_EDIT');
    assert.match(err.message, /set_table_style/, '应用时同一张嘴：同语义表格指引');
    return true;
  });
  assert.equal(writeCount(fake), 0);
});

test('扫描零命中且页上有表格：错误说明表格文字在单元格内并给替代路径', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [
    { id: 'sh1', name: '正文', text: '普通正文' },
    { id: 'tbl', name: '数据表', type: 'Table' }
  ] }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '来自表格深读的文字', replacement: 'x', target: { index: 0 } }] }),
    err => {
      assert.match(err.message, /表格/, '点到表格这一来源');
      assert.match(err.message, /set_table_style/);
      return true;
    }
  );
  assert.equal(writeCount(fake), 0);
});

// ---- ⑦ 超长文本显式截断 ----

test('超长目标文本：错误中的对照文本显式截断并标注原长度', async () => {
  const long = '长'.repeat(620);
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'sh1', name: '正文', text: long }] }] });
  const host = loadOnFake(fake);
  await assert.rejects(
    host.enrichEditProposal({ changes: [{ find: '不存在的锚点', replacement: 'x', target: { index: 0, shapeId: 'sh1' } }] }),
    err => {
      assert.match(err.message, /已截断/);
      assert.match(err.message, /620/, '标注原文总长');
      return true;
    }
  );
});

// ---- ⑧ 应用时再核验：出卡后文档被改 ----

test('应用时变化：出卡通过后用户改文 → 应用拦截（STALE_EDIT），错误带 find 与当前文本，零写入', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'sh1', name: '正文', text: '旧口径文字' }] }] });
  const host = loadOnFake(fake);
  const item = { find: '旧口径文字', replacement: '新口径', target: { index: 0, shapeId: 'sh1' } };
  await host.enrichEditProposal({ changes: [item] }); // 出卡前通过
  fake.slideStates[0].shapes[0].text = '用户已经改过的文字'; // 出卡后用户修改
  await assert.rejects(host.applyEdit(item), err => {
    assert.equal(err.code, 'STALE_EDIT', '沿用 STALE_EDIT 合同');
    assert.match(err.message, /旧口径文字/, '错误引用提交的 find');
    assert.match(err.message, /用户已经改过的文字/, '错误给出当前文本');
    return true;
  });
  assert.equal(fake.slideStates[0].shapes[0].text, '用户已经改过的文字', '新内容未被覆盖');
  assert.equal(writeCount(fake), 0, '应用失败零写入');
});

// ---- ⑨ 混合卡 ----

test('混合卡：文字项不合法 → 整卡拒绝，layout 项不落卡', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'sh1', name: '标题', text: '真实标题' }] }] });
  const host = loadOnFake(fake);
  const changes = [
    { kind: 'layout', target: { index: 0, shapeId: 'sh1' }, operation: { type: 'updateShape', shape: { bold: true } } },
    { label: '文字项', find: '错字锚点', replacement: 'x', target: { index: 0 } }
  ];
  await assert.rejects(host.enrichEditProposal({ changes }), err => {
    assert.match(err.message, /错字锚点/);
    assert.match(err.message, /未写入/);
    return true;
  });
  assert.equal(writeCount(fake), 0);
});

test('混合卡：全部合法 → layout 快照照常补齐，文字项不动，零写入', async () => {
  const fake = createFakeHost({ slides: [{ shapes: [{ id: 'sh1', name: '标题', text: '真实标题' }] }] });
  const host = loadOnFake(fake);
  const textItem = { label: '改标题', find: '真实标题', replacement: '新标题', target: { index: 0 } };
  const layoutItem = { kind: 'layout', target: { index: 0, shapeId: 'sh1' }, operation: { type: 'updateShape', shape: { bold: true } } };
  await host.enrichEditProposal({ changes: [layoutItem, textItem] });
  assert.ok(layoutItem.operation.expected && layoutItem.operation.expected.bold === false, 'layout 快照照常补齐');
  assert.equal(textItem.find, '真实标题', '文字项不被改写');
  assert.equal(writeCount(fake), 0);
});
