'use strict';
// 【recovery-loop · proposal 分支】修复验收：api.js 的 propose_edits 分支此前对
// App.presentEditProposal 无 try/catch——PPT/Excel 出卡前 enrich 抛错（错字 find、过期
// shapeId、expectedCells 失配）会直接炸出整个 Agent Loop，无重读回喂、无 3 次保护。
// 本文件用 propose_edits 真实分支验收：
//   1) PPT 走真实 host-powerpoint enrich/apply（仿真机假宿主）：enrich 抛错 → 转成
//      success:false + retryable:true 回执回喂 → get_slide 真码重读 → 修正卡应用成功；
//   2) 同参 proposal 连续抛错 3 次：第 2 次提示、第 3 次停止，无第 4 次请求，零写入；
//   3) Word 保持原行为：异常照常上抛，不转换回执；
//   4) Excel 用与 host-excel.js enrichEditProposal 同结构的失配异常走同一闭环
//      （read_range 重读 → 修正 expectedCells → 应用）。
const assert = require('node:assert/strict');
const test = require('node:test');
const { createRecoveryLoopSandbox } = require('./recovery-loop-harness.js');

// 与 ppt-anchor-recovery.test.js 同款写操作流水判定：enrich 只读，只有 apply 才落盘
const WRITE_OPS = /^(text|insertText|geometry|name|fill\.|line\.|font\.|bulletFormat\.|delete|setZOrder|setHyperlink|addTextBox|addGeometricShape|addSlide|deleteSlide|background\.|adjustments\.|duplicate)/;
const writeCount = fake => fake.log.filter(e => WRITE_OPS.test(e.op)).length;

// fake 宿主不深拷贝 spec：必须每次给全新对象，避免前一个测试的写入污染后一个测试的初始文本
const slideSpec = () => ({ slides: [{ shapes: [{ id: 'sh1', name: '标题', text: '竞品分析显示份额领先。' }] }] });
const badChange = () => ({
  label: '改写标题',
  find: '竞业分析显示份额领先',   // 错字：真实文本是「竞品」且句末有句号
  replacement: '对手分析显示份额领先',
  target: { index: 0, shapeId: 'sh1' }
});
const goodChange = () => ({
  label: '改写标题',
  find: '竞品分析',
  replacement: '对手分析',
  target: { index: 0, shapeId: 'sh1' }
});
const propose = change => ({ tool_calls: [{ name: 'propose_edits', args: { explanation: '按用户要求改写标题', changes: [change] } }] });

function lastAssistant(box) {
  const list = box.messages();
  return list[list.length - 1];
}

test('PPT 真实 enrich：抛错回喂 retryable 回执 → get_slide 重读 → 修正卡应用，全程只写一次', async () => {
  const box = createRecoveryLoopSandbox({
    hostType: 'powerpoint',
    accessMode: 'confirm',
    powerPointFake: slideSpec(),
    script: [
      propose(badChange()),
      { tool_calls: [{ name: 'get_slide', args: { index: 0 } }] },
      propose(goodChange()),
      { content: '已把第 1 页标题改写为「对手分析显示份额领先」。' }
    ]
  });
  await box.run();
  assert.equal(box.requests.length, 4, '失败回喂 → 重读 → 修正卡 → 最终回复，共 4 次请求');
  // 第 2 次请求的上下文里有第 1 次失败的结构化回执：可重试 + 完整 enrich 原因（提交 find、当前文本、get_slide 指引）
  const fedTool = box.requests[1].messages.find(message => message.role === 'tool' && message.name === 'propose_edits');
  assert.ok(fedTool, '失败回执以 tool 结果进入第 2 次请求上下文');
  assert.match(fedTool.content, /"retryable":true/, 'enrich 异常被转成 success:false + retryable:true 回执');
  assert.match(fedTool.content, /竞业分析显示份额领先/, '回执引用模型提交的 find');
  assert.match(fedTool.content, /竞品分析显示份额领先。/, '回执给出目标当前真实文本');
  assert.match(fedTool.content, /get_slide/, '回执带重读指引');
  // 第 3 次请求前模型确实用真码 get_slide 读到了真实内容
  const reread = box.requests[2].messages.filter(message => message.name === 'get_slide').pop();
  assert.ok(reread, 'get_slide 结果进入第 3 次请求上下文');
  assert.match(reread.content, /竞品分析显示份额领先/, '重读返回的是宿主当前真实文本');
  // 修正卡真实应用：fake 落盘一次文本替换，最终文本正确
  assert.equal(writeCount(box.fake), 1, 'enrich 失败零写入，只有修正卡应用写一次');
  assert.equal(box.fake.slideStates[0].shapes[0].__st.text, '对手分析显示份额领先。');
  assert.equal(lastAssistant(box).content, '已把第 1 页标题改写为「对手分析显示份额领先」。');
});

test('PPT 真实 enrich：同参 proposal 连续抛错 3 次，第 2 次提示第 3 次停止，零写入', async () => {
  const box = createRecoveryLoopSandbox({
    hostType: 'powerpoint',
    accessMode: 'confirm',
    powerPointFake: slideSpec(),
    script: [
      propose(badChange()),
      propose(badChange()),
      propose(badChange())
      // 故意不写第 4 步：若未在第 3 次失败后收尾，脚本化客户端会抛「不该发生第 4 次请求」
    ]
  });
  await box.run();
  assert.equal(box.requests.length, 3, '第 3 次同参抛错后必须收尾，不得有第 4 次模型请求');
  assert.equal(writeCount(box.fake), 0, '三次都停在出卡前核验，零写入');
  assert.equal(box.fake.slideStates[0].shapes[0].__st.text, '竞品分析显示份额领先。', '文档原样');
  const correction = box.requests[2].messages.find(
    message => message.role === 'user' && String(message.content).includes('【系统纠偏，不是新任务】')
  );
  assert.ok(correction, '第 3 次请求前应已注入第 2 次失败的纠偏提示');
  assert.match(correction.content, /propose_edits/);
  assert.match(correction.content, /相同参数/);
  const last = lastAssistant(box);
  assert.equal(last.role, 'assistant');
  assert.match(last.content, /propose_edits/);
  assert.match(last.content, /相同参数/);
});

test('Word：proposal 异常保持原行为，照常上抛不转回执', async () => {
  const box = createRecoveryLoopSandbox({
    hostType: 'word',
    accessMode: 'confirm',
    prompt: '把「旧标题」改成「新标题」。',
    presentEditProposal: async () => { throw new Error('锚点文本「旧标题」不存在，未写入'); },
    script: [
      propose({ label: '改标题', find: '旧标题', replacement: '新标题', target: { paragraphIndex: 1 } })
    ]
  });
  await assert.rejects(box.run(), /锚点文本「旧标题」不存在/, 'Word 不做 PPT/Excel 的回执转换，异常原样逃出循环');
  assert.equal(box.requests.length, 1, '没有回喂，不再有第 2 次请求');
  assert.equal(box.writes.length, 0);
});

// 与 host-excel.js enrichEditProposal 同结构的失配异常 + read_range 闭环（仿真内存工作表）
function createExcelProposalStub(initialCells) {
  const cells = Object.assign({}, initialCells);
  const writes = [];
  return {
    writes,
    cells,
    read_range: args => ({ success: true, sheetId: args.sheetId, range: args.range, values: [[cells[args.range]]] }),
    presentEditProposal: async (uiCall, args) => {
      const items = Array.isArray(args.changes) ? args.changes : (Array.isArray(args.edits) ? args.edits : []);
      for (const item of items) {
        const target = item.target || {};
        const expected = target.expectedCells && target.expectedCells[0] && target.expectedCells[0][0];
        const expectedValue = expected && typeof expected === 'object' ? expected.value : expected;
        if (expectedValue !== cells[target.range]) {
          throw new Error('目标单元格已变化：提案「' + (item.label || '') + '」的 expectedCells 快照与当前工作表不符（sheetId ' + target.sheetId + ' 区域 ' + target.range + '，失配 1 格）。内容可能在出卡后被修改。请用 read_range(sheetId=' + target.sheetId + ', range="' + target.range + '") 重新读取该区域，按真实内容重填 expectedCells 后重新提交。');
        }
      }
      let applied = 0;
      for (const item of items) {
        const target = item.target || {};
        const next = target.cells && target.cells[0] && target.cells[0][0];
        cells[target.range] = next && typeof next === 'object' ? next.value : next;
        writes.push({ range: target.range });
        applied++;
      }
      return { success: true, applied, results: items.map(() => ({ success: true, applied: true })) };
    }
  };
}

test('Excel：同结构失配异常 3 次，第 2 次提示第 3 次停止，零写入', async () => {
  const stub = createExcelProposalStub({ B2: 135 });
  const stale = () => propose({
    label: '更新报价',
    find: 'B2: 120',
    replacement: 'B2: 150',
    target: { sheetId: 1, range: 'B2', expectedCells: [[{ value: 120 }]], cells: [[{ value: 150 }]] }
  });
  const box = createRecoveryLoopSandbox({
    hostType: 'excel',
    accessMode: 'confirm',
    prompt: '把 B2 的报价改成 150。',
    presentEditProposal: stub.presentEditProposal,
    script: [stale(), stale(), stale()]
  });
  await box.run();
  assert.equal(box.requests.length, 3, '第 3 次同参失配后必须收尾，不得有第 4 次模型请求');
  assert.equal(stub.writes.length, 0, '失配停在出卡前核验，零写入');
  assert.equal(stub.cells.B2, 135, '工作表原样');
  const fedTool = box.requests[1].messages.find(message => message.role === 'tool' && message.name === 'propose_edits');
  assert.ok(fedTool, '失败回执以 tool 结果进入第 2 次请求上下文');
  assert.match(fedTool.content, /"retryable":true/);
  assert.match(fedTool.content, /read_range\(sheetId=1/, '回执带 read_range 重读指引');
  assert.match(fedTool.content, /重新读取该区域/);
  const correction = box.requests[2].messages.find(
    message => message.role === 'user' && String(message.content).includes('【系统纠偏，不是新任务】')
  );
  assert.ok(correction, '第 3 次请求前应已注入第 2 次失败的纠偏提示');
  const last = lastAssistant(box);
  assert.equal(last.role, 'assistant');
  assert.match(last.content, /propose_edits/);
  assert.match(last.content, /相同参数/);
});

test('Excel：失配回喂 → read_range 重读 → 修正 expectedCells 应用成功后收轮', async () => {
  const stub = createExcelProposalStub({ B2: 135 });
  const box = createRecoveryLoopSandbox({
    hostType: 'excel',
    accessMode: 'confirm',
    prompt: '把 B2 的报价改成 150。',
    presentEditProposal: stub.presentEditProposal,
    tools: { read_range: args => stub.read_range(args) },
    script: [
      propose({ label: '更新报价', find: 'B2: 120', replacement: 'B2: 150', target: { sheetId: 1, range: 'B2', expectedCells: [[{ value: 120 }]], cells: [[{ value: 150 }]] } }),
      { tool_calls: [{ name: 'read_range', args: { sheetId: 1, range: 'B2' } }] },
      propose({ label: '更新报价', find: 'B2: 135', replacement: 'B2: 150', target: { sheetId: 1, range: 'B2', expectedCells: [[{ value: 135 }]], cells: [[{ value: 150 }]] } })
      // Excel 提案应用后按既有设计收轮（finishAfterProposal），不再追加最终回复请求
    ]
  });
  await box.run();
  assert.equal(box.requests.length, 3, '失配回喂 → 重读 → 修正卡应用后收轮');
  const reread = box.requests[2].messages.filter(message => message.name === 'read_range').pop();
  assert.ok(reread, 'read_range 结果进入第 3 次请求上下文');
  assert.match(reread.content, /135/, '重读返回工作表当前真实值');
  assert.equal(stub.writes.length, 1, '只有修正卡应用写一次');
  assert.equal(stub.cells.B2, 150);
});

// ---- 停止/取消语义：取消类异常必须原样上抛终止本轮，不得转成 retryable 回喂诱发继续写 ----

test('PPT：WRITE_CANCELLED 原样上抛，不转 retryable、不回喂、零写入', async () => {
  const box = createRecoveryLoopSandbox({
    hostType: 'powerpoint',
    accessMode: 'confirm',
    powerPointFake: slideSpec(),
    presentEditProposal: async (uiCall, args) => {
      // 真实数据接线：先走真实 enrich（真实加载页面文本，fake 的形状状态随批 sync 建立），
      // 再在写入阶段抛出与 host.js docWriteBegin.throwIfCancelled 同形状的取消异常——
      // WRITE_CANCELLED 本来就发生在 apply 期，不是读取期。
      await box.App.host.enrichEditProposal(args);
      const err = new Error('应用操作已被停止取消：本次写入不再执行。');
      err.code = 'WRITE_CANCELLED';
      throw err;
    },
    script: [propose(goodChange())]
  });
  await assert.rejects(box.run(), e => e.code === 'WRITE_CANCELLED' && /停止取消/.test(e.message), '取消异常保留原始身份上抛');
  assert.equal(box.requests.length, 1, '取消不产生回喂，无第 2 次请求');
  assert.equal(writeCount(box.fake), 0, '零写入');
  assert.equal(box.fake.slideStates[0].shapes[0].__st.text, '竞品分析显示份额领先。', 'enrich 已加载真实文本，取消后原文未变');
});

test('PPT：AbortError 原样上抛终止本轮，不转 retryable、零写入', async () => {
  const box = createRecoveryLoopSandbox({
    hostType: 'powerpoint',
    accessMode: 'confirm',
    powerPointFake: slideSpec(),
    presentEditProposal: async () => {
      const err = new Error('Request aborted');
      err.name = 'AbortError';
      throw err;
    },
    script: [propose(goodChange())]
  });
  await assert.rejects(box.run(), e => e.name === 'AbortError', 'AbortError 保留原始身份上抛');
  assert.equal(box.requests.length, 1, '中断不产生回喂，无第 2 次请求');
  assert.equal(writeCount(box.fake), 0);
});

test('PPT：stopRequested 置位时抛错按停止收尾，不转 retryable、零写入', async () => {
  let boxRef = null;
  const box = createRecoveryLoopSandbox({
    hostType: 'powerpoint',
    accessMode: 'confirm',
    powerPointFake: slideSpec(),
    presentEditProposal: async () => {
      boxRef.App.state.stopRequested = true;  // 模拟用户在出卡/应用期间点了停止
      throw new Error('写入中途被打断');
    },
    script: [propose(goodChange())]
  });
  boxRef = box;
  await assert.rejects(box.run(), e => e.name === 'AbortError', 'stopRequested 优先按统一 AbortError 收尾');
  assert.equal(box.requests.length, 1, '停止不产生回喂，无第 2 次请求');
  assert.equal(writeCount(box.fake), 0);
});

test('Excel：WRITE_CANCELLED 同样原样上抛，不转 retryable', async () => {
  const box = createRecoveryLoopSandbox({
    hostType: 'excel',
    accessMode: 'confirm',
    prompt: '把 B2 的报价改成 150。',
    presentEditProposal: async () => {
      const err = new Error('应用操作已被停止取消：本次写入不再执行。');
      err.code = 'WRITE_CANCELLED';
      throw err;
    },
    script: [
      propose({ label: '更新报价', find: 'B2: 135', replacement: 'B2: 150', target: { sheetId: 1, range: 'B2', expectedCells: [[{ value: 135 }]], cells: [[{ value: 150 }]] } })
    ]
  });
  await assert.rejects(box.run(), e => e.code === 'WRITE_CANCELLED');
  assert.equal(box.requests.length, 1, '取消不产生回喂，无第 2 次请求');
  assert.equal(box.writes.length, 0);
});
