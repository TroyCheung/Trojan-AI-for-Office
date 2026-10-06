'use strict';
// 【recovery-loop】progress-guard.js 单元测试（移植自 Office-AI-addin 的
// tests/progress-guard.test.js，按公开版无 host.js 合同的场景裁剪）：
// 指纹忽略装饰字段、重试元数据不掩盖相同错误、成功写入开新纪元、失败记忆结转、
// 分页与参数/结果变化都继续。阈值本身（2 提示 3 停止）由 recovery-loop.test.js
// 在真实 Agent Loop 层验收，这里钉住守卫自身的判定语义。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const ROOT = path.resolve(__dirname, '..');
function loadGuard() {
  const App = {};
  vm.runInNewContext(fs.readFileSync(path.join(ROOT, 'app/assets/js/progress-guard.js'), 'utf8'), { window: { App } }, { filename: 'progress-guard.js' });
  return App.ProgressGuard;
}

test('同参同结果失败：continue → correct（2 提示）→ stop（3 停止）', () => {
  const PG = loadGuard();
  const g = PG.create();
  const fail = { success: false, error: 'nope' };
  assert.equal(g.observe('insert_textbox', { target: { index: 1 } }, fail).action, 'continue');
  const correct = g.observe('insert_textbox', { target: { index: 1 } }, fail);
  assert.equal(correct.action, 'correct');
  assert.equal(correct.reason, 'repeat-fail');
  assert.match(correct.guidance, /相同参数/);
  const stop = g.observe('insert_textbox', { target: { index: 1 } }, fail);
  assert.equal(stop.action, 'stop');
  assert.match(stop.guidance, /insert_textbox/);
});

test('装饰字段（explanation/label/summary）变化不改变指纹，任意深度生效', () => {
  const PG = loadGuard();
  assert.equal(
    PG.fingerprint('propose_edits', { changes: [{ find: '旧标题', replacement: '新标题', label: 'A', explanation: 'x' }] }),
    PG.fingerprint('propose_edits', { changes: [{ summary: 'y', find: '旧标题', replacement: '新标题' }] }),
    'label/explanation/summary 是装饰字段；find/replacement 是实质字段'
  );
  assert.notEqual(
    PG.fingerprint('propose_edits', { changes: [{ find: '旧标题', replacement: '新标题' }] }),
    PG.fingerprint('propose_edits', { changes: [{ find: '另一个标题', replacement: '新标题' }] })
  );
});

test('重试元数据（consecutiveFailures/guidance/stopRetrying）不掩盖相同错误', () => {
  const g = loadGuard().create();
  const actions = [];
  for (let i = 1; i <= 3; i++) {
    actions.push(g.observe('insert_text', { find: 'x' }, { success: false, error: 'same', consecutiveFailures: i, guidance: 'retry' + i }).action);
  }
  assert.deepEqual(actions, ['continue', 'correct', 'stop']);
});

test('参数改变或结果改变都是新证据，不触发停止', () => {
  const g = loadGuard().create();
  for (let i = 0; i < 4; i++) {
    assert.equal(g.observe('insert_textbox', { target: { index: i } }, { success: false, error: 'same' }).action, 'continue', '参数变化 ' + i);
  }
  for (let i = 0; i < 4; i++) {
    assert.equal(g.observe('insert_textbox', { target: { index: 9 } }, { success: false, error: 'err' + i }).action, 'continue', '结果变化 ' + i);
  }
});

test('成功写入开新纪元：写入后的同参数读取不继承上一轮计数', () => {
  const g = loadGuard().create();
  const read = { success: true, paragraphs: [{ text: 'same' }] };
  assert.equal(g.observe('get_paragraphs', {}, read).action, 'continue');
  assert.equal(g.observe('insert_hyperlink', { find: '来源', occurrence: 1 }, { success: true, address: 'https://a.example' }).action, 'continue');
  assert.equal(g.observe('get_paragraphs', {}, read).action, 'continue', '写入后同参读取从 1 重数');
  assert.equal(g.observe('get_paragraphs', {}, read).action, 'correct');
});

test('失败记忆跨写入结转：失败→写入成功→同参失败按 1+1 计为提示，再失败停止', () => {
  const g = loadGuard().create();
  const args = { target: { index: 1 } };
  const fail = { success: false, error: 'locked' };
  assert.equal(g.observe('insert_textbox', args, fail).action, 'continue');
  assert.equal(g.observe('insert_textbox', args, { success: true, id: 'sh-1' }).action, 'continue');
  assert.equal(g.observe('insert_textbox', args, fail).action, 'correct', '本 epoch 1 次 + 结转 1 次 = 2');
  assert.equal(g.observe('insert_textbox', args, fail).action, 'stop');
});

test('分页游标变化不是重复；空转读取同结果 2 提示 3 停止', () => {
  const g = loadGuard().create();
  assert.equal(g.observe('get_paragraphs', { start: 0, count: 20 }, { success: true, hasMore: true, nextStart: 20 }).action, 'continue');
  assert.equal(g.observe('get_paragraphs', { start: 20, count: 20 }, { success: true, hasMore: false }).action, 'continue');
  const result = { success: true, paragraphs: [{ text: 'same' }] };
  assert.equal(g.observe('get_slide', { index: 2 }, result).action, 'continue');
  assert.equal(g.observe('get_slide', { index: 2 }, result).action, 'correct');
  assert.equal(g.observe('get_slide', { index: 2 }, result).action, 'stop');
});

test('isSuccessfulWrite：propose_edits 按 applied 判定，零写入不算成功', () => {
  const PG = loadGuard();
  assert.equal(PG.isSuccessfulWrite('propose_edits', { success: true, applied: 2 }), true);
  assert.equal(PG.isSuccessfulWrite('propose_edits', { success: true, applied: 0 }), false);
  assert.equal(PG.isSuccessfulWrite('insert_textbox', { success: true, id: 'sh-1' }), true);
  assert.equal(PG.isSuccessfulWrite('insert_textbox', { success: false, id: null }), false);
  assert.equal(PG.isSuccessfulWrite('get_slide', { success: true }), false, '读取工具永不算成功写入');
});

test('observeBatch：同一批内 3 个同参同结果失败直接给 stop', () => {
  const g = loadGuard().create();
  const fail = { success: false, error: 'nope' };
  const verdict = g.observeBatch([
    { name: 'set_cell_range', args: { range: 'B2' }, result: fail },
    { name: 'set_cell_range', args: { range: 'B2' }, result: fail },
    { name: 'set_cell_range', args: { range: 'B2' }, result: fail }
  ]);
  assert.equal(verdict.action, 'stop');
  assert.equal(verdict.reason, 'repeat-fail');
});
