'use strict';
// 【recovery-loop】真实旧版 Agent Loop 测试台：把公开版 app/assets/js/api.js 原样装进 VM
// 沙箱，用脚本化的模型客户端与假宿主工具驱动 runAgentLoop。不 mock 循环本身——被测代码
// 就是发布产物那份 api.js；默认先加载 progress-guard.js，模拟 taskpane 的目标接线顺序
// （progress-guard.js 必须在 api.js 之前）。loadGuard:false 可测未接线的降级路径。
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ROOT = path.resolve(__dirname, '..');
const { createFakeHost } = require('./fake-powerpoint-host.js');

function createRecoveryLoopSandbox(options = {}) {
  const hostType = options.hostType || 'powerpoint';
  const script = options.script || [];
  const tools = options.tools || {};
  const requests = [];
  const toolCalls = [];
  const writes = [];

  const sandbox = { console };
  sandbox.window = sandbox;
  const App = sandbox.App = {};
  App.HOSTS = {};
  App.hasOffice = () => false;
  // powerPointFake：装入真实 host.js + host-powerpoint.js 与「仿真机」假宿主，
  // presentEditProposal 走真实 enrich/apply 契约（enrich 抛错原样上抛，由 api.js 的 catch 转回执）。
  let fake = null;
  if (options.powerPointFake) {
    fake = createFakeHost(options.powerPointFake);
    sandbox.Office = { context: { requirements: { isSetSupported: () => true } } };
    sandbox.PowerPoint = { run: async fn => fn(fake.context) };
  }
  App.host = { hostType, systemPrompt: 'You are the ' + hostType + ' assistant.' };
  App.state = {
    messages: [{ role: 'user', content: options.prompt || '把第 2 页标题改成「结论」。' }],
    settings: { accessMode: options.accessMode || 'auto', apiKey: 'recovery-loop-test-key' },
    sessions: [],
    currentSessionId: null,
    stopRequested: false,
    abortController: null,
    steerQueue: [],
    activeSkillId: null,
    pptTaskScope: null,
    activePptTaskScope: null,
    workPhase: null
  };
  App.now = () => Date.now();
  App.t = key => ({ stopped: '已停止', stepLimitReached: '本任务步骤预算已用完。', truncatedNote: '输出被截断' }[key] || key);
  App.currentDialect = () => ({ needsReasoningEcho: false });
  App.currentVisionKey = () => 'test';
  App.render = () => {};
  App.patchStreamingMessage = () => {};
  App.persistCurrentSession = () => {};

  let scriptIndex = 0;
  App.aiClient = {
    reasonEchoRequired: () => false,
    async callChatCompletions(messages) {
      const stepNo = scriptIndex + 1;
      if (scriptIndex >= script.length) {
        throw new Error('不该发生第 ' + stepNo + ' 次模型请求（脚本共 ' + script.length + ' 步）');
      }
      const step = script[scriptIndex++];
      requests.push({
        step: stepNo,
        messages: messages.map(message => ({
          role: message.role,
          name: message.name,
          content: typeof message.content === 'string' ? message.content : JSON.stringify(message.content)
        }))
      });
      const message = step.tool_calls
        ? {
            tool_calls: step.tool_calls.map((call, i) => ({
              id: 'call_' + stepNo + '_' + i,
              type: 'function',
              function: { name: call.name, arguments: JSON.stringify(call.args || {}) }
            }))
          }
        : { content: String(step.content || '') };
      return { choices: [{ message, finish_reason: step.tool_calls ? 'tool_calls' : 'stop' }] };
    }
  };

  vm.createContext(sandbox);
  if (fake) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, 'app/assets/js/host.js'), 'utf8'), sandbox, { filename: 'host.js' });
    vm.runInContext(fs.readFileSync(path.join(ROOT, 'app/assets/js/host-powerpoint.js'), 'utf8'), sandbox, { filename: 'host-powerpoint.js' });
  }
  if (options.loadGuard !== false) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, 'app/assets/js/progress-guard.js'), 'utf8'), sandbox, { filename: 'progress-guard.js' });
  }
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'app/assets/js/api.js'), 'utf8'), sandbox, { filename: 'api.js' });

  if (fake) {
    // 真实宿主工具执行器要求 App.hasOffice() 为真；api.js 的全文/元数据注入失败会自行降级，
    // 假宿主的 PowerPoint.run 全程在内存中，不会触网。
    App.hasOffice = () => true;
    App.host = App.HOSTS.powerpoint;  // 真实宿主：hostType/systemPrompt/enrich/apply/toolExecutors 齐备
    // 与 ui.js presentEditProposal 同契约的仿真：真实 enrich（抛错原样上抛）→ 逐项 applyEdit →
    // 卡片应用结果（逐项带 index/applied，形状对齐宿主提案路径，供 proposalTouchedSlideIndexes
    // 等消费方使用）。不模拟用户拒绝；本测试台只关心 enrich 失败回喂与应用成功两条路径。
    App.presentEditProposal = async (uiCall, args) => {
      const normalized = await App.host.enrichEditProposal(args);
      const items = Array.isArray(normalized.changes) ? normalized.changes : (Array.isArray(normalized.edits) ? normalized.edits : []);
      const results = [];
      let applied = 0;
      for (let i = 0; i < items.length; i++) {
        const appliedResult = await App.host.applyEdit(items[i]);
        if (appliedResult && appliedResult.success) applied++;
        results.push(Object.assign({ index: i, applied: Boolean(appliedResult && appliedResult.success) }, appliedResult));
      }
      return { success: applied > 0, applied, results };
    };
  }
  if (typeof options.presentEditProposal === 'function') {
    App.presentEditProposal = options.presentEditProposal;
  }

  // 工具分发在所有脚本加载后接线：真实宿主时包一层 host.js 分发器（get_slide 等走真码），
  // 否则用假宿主工具表。handler(args, ctx) 返回结果或抛错；ctx.callIndex 是本 run 第几次调用，
  // ctx.writes 由 handler 在真正落盘时显式记录——失败场景不记录即「零写入」。
  const realDispatch = App.executeToolByName;
  const useRealDispatch = Boolean(fake);
  App.executeToolByName = async (name, args) => {
    const callIndex = toolCalls.length + 1;
    toolCalls.push({ index: callIndex, name, args });
    if (useRealDispatch) return await realDispatch(name, args);
    const handler = tools[name];
    if (typeof handler !== 'function') return { success: false, error: '测试沙箱未定义工具 ' + name };
    return await handler(args, { callIndex, writes, fake });
  };

  return {
    App,
    requests,
    toolCalls,
    writes,
    fake,
    run: () => App.runAgentLoop(),
    messages: () => App.state.messages
  };
}

module.exports = { createRecoveryLoopSandbox };
