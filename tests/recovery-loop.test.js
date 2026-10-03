'use strict';
// 【recovery-loop】公开版（旧版 Agent Loop）的 PPT/Excel 同参同结果失败停止保护测试。
// 被测对象：真实 app/assets/js/api.js（VM 原样加载）+ app/assets/js/progress-guard.js
// （按目标接线顺序先于 api.js 加载）。验收口径：
//   1) PPT/Excel 同参同结果失败：第 2 次提示（系统纠偏）、第 3 次强制停止，不再有第 4 次模型请求；
//   2) 参数改变不中止；
//   3) 恢复后能继续；
//   4) 停止场景零写入；
//   5) Word 行为不扩大（保持旧版 3 提示 6 停止指引、不强制中止）；
//   6) progress-guard.js 未接线时循环不报错、行为与旧版一致（接线前的安全降级）。
const assert = require('node:assert/strict');
const test = require('node:test');
const { createRecoveryLoopSandbox } = require('./recovery-loop-harness.js');

const IDENTICAL_PPT_ARGS = { target: { index: 1, shapeId: 'sh-9' }, text: '结论' };
// 锚点校验失败发生在任何宿主写入之前：handler 不记录 writes，即真实语义的「未写入」。
const anchorMissing = () => ({ success: false, error: '锚点文本「旧标题」在第 2 页不存在，未执行任何写入' });

function lastAssistant(box) {
  const list = box.messages();
  return list[list.length - 1];
}

test('PPT：同参同结果失败第 2 次提示、第 3 次停止，无第 4 次请求，零写入', async () => {
  const box = createRecoveryLoopSandbox({
    hostType: 'powerpoint',
    tools: { insert_textbox: anchorMissing },
    script: [
      { tool_calls: [{ name: 'insert_textbox', args: IDENTICAL_PPT_ARGS }] },
      { tool_calls: [{ name: 'insert_textbox', args: IDENTICAL_PPT_ARGS }] },
      { tool_calls: [{ name: 'insert_textbox', args: IDENTICAL_PPT_ARGS }] }
      // 故意不写第 4 步：若循环没有在第 3 次失败后收尾，脚本化客户端会抛「不该发生第 4 次请求」
    ]
  });
  await box.run();
  assert.equal(box.requests.length, 3, '第 3 次同参同结果失败后必须收尾，不得有第 4 次模型请求');
  assert.equal(box.toolCalls.length, 3);
  // 参数在 VM 沙箱内构造（原型属沙箱 realm），deepStrictEqual 会因原型不同误报，用 JSON 比较
  for (const call of box.toolCalls) assert.equal(JSON.stringify(call.args), JSON.stringify(IDENTICAL_PPT_ARGS), '三次调用参数完全一致');
  assert.equal(box.writes.length, 0, '锚点校验失败发生在写入之前，全程零写入');
  // 第 2 次失败后的系统纠偏（提示）应出现在第 3 次请求的上下文里
  const correction = box.requests[2].messages.find(
    message => message.role === 'user' && String(message.content).includes('【系统纠偏，不是新任务】')
  );
  assert.ok(correction, '第 3 次请求前应已注入第 2 次失败的纠偏提示');
  assert.match(correction.content, /insert_textbox/);
  assert.match(correction.content, /相同参数/);
  // 第 3 次失败后强制停止：最后一条消息是面向用户的停止说明
  const last = lastAssistant(box);
  assert.equal(last.role, 'assistant');
  assert.match(last.content, /insert_textbox/);
  assert.match(last.content, /相同参数/);
});

test('Excel：同参同结果失败第 2 次提示、第 3 次停止，无第 4 次请求，零写入', async () => {
  const args = { range: 'Sheet1!B2', values: [[42]] };
  const box = createRecoveryLoopSandbox({
    hostType: 'excel',
    prompt: '把 B2 的值改成 42。',
    tools: { set_cell_range: () => ({ success: false, error: '工作表受保护，单元格未写入' }) },
    script: [
      { tool_calls: [{ name: 'set_cell_range', args }] },
      { tool_calls: [{ name: 'set_cell_range', args }] },
      { tool_calls: [{ name: 'set_cell_range', args }] }
    ]
  });
  await box.run();
  assert.equal(box.requests.length, 3, '第 3 次同参同结果失败后必须收尾，不得有第 4 次模型请求');
  assert.equal(box.toolCalls.length, 3);
  assert.equal(box.writes.length, 0, '保护工作表拒绝写入，全程零写入');
  const correction = box.requests[2].messages.find(
    message => message.role === 'user' && String(message.content).includes('【系统纠偏，不是新任务】')
  );
  assert.ok(correction, '第 3 次请求前应已注入第 2 次失败的纠偏提示');
  const last = lastAssistant(box);
  assert.equal(last.role, 'assistant');
  assert.match(last.content, /set_cell_range/);
  assert.match(last.content, /相同参数/);
});

test('PPT：参数改变不算同参重复，连续失败也不中止', async () => {
  const variants = [1, 2, 3, 4].map(n => ({ target: { index: n, shapeId: 'sh-' + n }, text: '结论' + n }));
  const box = createRecoveryLoopSandbox({
    hostType: 'powerpoint',
    tools: { insert_textbox: anchorMissing },
    script: [
      ...variants.map(args => ({ tool_calls: [{ name: 'insert_textbox', args }] })),
      { content: '四处锚点都无法命中，建议先手工核对标题文本。' }
    ]
  });
  await box.run();
  assert.equal(box.requests.length, 5, '参数每次变化：不触发同参停止，循环执行到最终回复');
  assert.equal(box.toolCalls.length, 4);
  assert.equal(box.writes.length, 0);
  const last = lastAssistant(box);
  assert.equal(last.role, 'assistant');
  assert.equal(last.content, '四处锚点都无法命中，建议先手工核对标题文本。');
  for (const request of box.requests) {
    for (const message of request.messages) {
      assert.doesNotMatch(String(message.content || ''), /【系统纠偏，不是新任务】/, '不同参数不应触发纠偏');
    }
  }
});

test('PPT：同参失败两次后工具恢复，任务继续完成（恢复后继续）', async () => {
  const box = createRecoveryLoopSandbox({
    hostType: 'powerpoint',
    tools: {
      insert_textbox: (args, ctx) => ctx.callIndex <= 2
        ? { success: false, error: '文档被临时锁定，未写入' }
        : (ctx.writes.push({ name: 'insert_textbox', args }), { success: true, id: 'shape-7' })
    },
    script: [
      { tool_calls: [{ name: 'insert_textbox', args: IDENTICAL_PPT_ARGS }] },
      { tool_calls: [{ name: 'insert_textbox', args: IDENTICAL_PPT_ARGS }] },
      { tool_calls: [{ name: 'insert_textbox', args: IDENTICAL_PPT_ARGS }] },
      { content: '已完成：第 2 页标题更新为「结论」。' }
    ]
  });
  await box.run();
  assert.equal(box.requests.length, 4, '第 3 次调用成功开新纪元，循环继续到最终回复');
  assert.equal(box.writes.length, 1, '恢复后完成一次真实写入');
  const last = lastAssistant(box);
  assert.equal(last.role, 'assistant');
  assert.equal(last.content, '已完成：第 2 页标题更新为「结论」。');
});

test('Word：同参同结果失败不接入新守卫，行为与旧版一致（不中止）', async () => {
  const args = { paragraphIndex: 3, text: '更明确的说法' };
  const box = createRecoveryLoopSandbox({
    hostType: 'word',
    prompt: '把第一段末句再改明确一点。',
    tools: { insert_text: anchorMissing },
    script: [
      ...[1, 2, 3, 4].map(() => ({ tool_calls: [{ name: 'insert_text', args }] })),
      { content: '这段暂时改不动，我先汇报原因。' }
    ]
  });
  await box.run();
  assert.equal(box.requests.length, 5, 'Word 不接入 2/3 守卫：同参失败按旧版继续由步数预算兜底');
  assert.equal(box.writes.length, 0);
  const last = lastAssistant(box);
  assert.equal(last.role, 'assistant');
  assert.equal(last.content, '这段暂时改不动，我先汇报原因。');
  for (const request of box.requests) {
    for (const message of request.messages) {
      assert.doesNotMatch(String(message.content || ''), /相同参数连续失败/, 'Word 不得出现 PPT/Excel 的新停止语义');
    }
  }
});

test('Word：抛错路径保持旧版 3 提示 6 停止指引，不强制中止', async () => {
  const args = { find: '旧句子', replacement: '新句子' };
  const box = createRecoveryLoopSandbox({
    hostType: 'word',
    prompt: '把「旧句子」替换成「新句子」。',
    tools: { replace_text: () => { throw new Error('段落被锁定'); } },
    script: [
      ...[1, 2, 3, 4, 5, 6].map(() => ({ tool_calls: [{ name: 'replace_text', args }] })),
      { content: '连续多次失败，我先停下来向你说明。' }
    ]
  });
  await box.run();
  assert.equal(box.requests.length, 7);
  const afterThird = JSON.stringify(box.requests[3].messages);
  assert.match(afterThird, /已经连续失败 3 次/, '第 3 次失败挂上换路径提示（旧版 3 提示）');
  const afterSixth = JSON.stringify(box.requests[6].messages);
  assert.match(afterSixth, /已经连续失败 6 次/, '第 6 次失败挂上停止指引（旧版 6 停止）');
  assert.match(afterSixth, /不要再用同样的方式重试/);
  const last = lastAssistant(box);
  assert.equal(last.role, 'assistant');
  assert.equal(last.content, '连续多次失败，我先停下来向你说明。', '旧版语义：stopRetrying 是给模型的指引，不强制中止循环');
});

test('未接线降级：progress-guard.js 未加载时，PPT 循环不报错且保持旧版行为', async () => {
  const box = createRecoveryLoopSandbox({
    hostType: 'powerpoint',
    loadGuard: false,
    tools: { insert_textbox: anchorMissing },
    script: [
      { tool_calls: [{ name: 'insert_textbox', args: IDENTICAL_PPT_ARGS }] },
      { tool_calls: [{ name: 'insert_textbox', args: IDENTICAL_PPT_ARGS }] },
      { tool_calls: [{ name: 'insert_textbox', args: IDENTICAL_PPT_ARGS }] },
      { content: '多次尝试未成功，我先说明情况。' }
    ]
  });
  await box.run(); // 正常 resolve 即证明：守卫缺失时 api.js 优雅降级，不抛错
  assert.equal(box.requests.length, 4, '未接线时同参失败不触发 2/3 停止（与接线前行为一致）');
  assert.equal(box.writes.length, 0);
  const last = lastAssistant(box);
  assert.equal(last.role, 'assistant');
  assert.equal(last.content, '多次尝试未成功，我先说明情况。');
});
