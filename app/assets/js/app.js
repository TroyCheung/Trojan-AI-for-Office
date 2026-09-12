(function () {
  'use strict';
  const App = (window.App = window.App || {});

  async function initOffice() {
    if (navigator.userAgent.indexOf('Trident') !== -1 || navigator.userAgent.indexOf('Edge') !== -1) {
      const legacy = document.getElementById('legacy-message');
      if (legacy) legacy.hidden = false;
    }
    if (typeof Office !== 'undefined' && Office.onReady) {
      try {
        // onReady 在无 Office 宿主或 CDN 异常时可能永不 resolve，3 秒兜底放行
        const info = await Promise.race([
          Office.onReady(),
          new Promise(resolve => setTimeout(() => resolve(null), 3000))
        ]);
        // 优先用 Office 报告的宿主类型选择提供者，回退到运行时探测。
        const hostType = App.hostTypeFromOffice(info && info.host) || App.detectHostType();
        App.selectHost(hostType);
        if (App.hasOffice()) {
          if (App.syncSharedSettings) await App.syncSharedSettings();
          App.state.workbookId = await App.getDocumentId();
          const md = await App.host.getMetadata().catch(() => null);
          // 仅 Excel 提供 workbookName；其余宿主保持空标签，避免展示过期状态。
          if (md && md.workbookName) App.state.workbookLabel = md.workbookName;
          // Excel 真机能力探测（只读、一次）：结果进 console 与「工具」页，为后续"按真机行为复刻夹具"提供实测依据
          if (hostType === 'excel' && typeof App.host.probeExcelCapabilities === 'function') {
            App.host.capabilities = await App.host.probeExcelCapabilities().catch(e => ({ probedAt: new Date().toISOString(), error: e.message }));
            console.log('[Excel caps]', App.host.capabilities);
          }
          // Word 真机能力探测（同 Excel，只读一次）：探测为 false 的批注工具由 capabilityToolGates 摘牌
          if (hostType === 'word' && typeof App.host.probeWordCapabilities === 'function') {
            App.host.capabilities = await App.host.probeWordCapabilities().catch(e => ({ probedAt: new Date().toISOString(), error: e.message }));
            console.log('[Word caps]', App.host.capabilities);
          }
        }
      } catch (e) {
        console.warn('[Office init]', e);
      }
    }
    if (App.applyHostAccessSettings) App.applyHostAccessSettings();
    App.ensureSession();
    App.render();
    // 学习复盘欠账补跑（v137）：延迟几秒避开启动竞态；用户很快发消息时复盘会自行让位
    if (App.review) setTimeout(() => { try { App.review.onStartup(); } catch {} }, 6000);
    // 【第 11 节】真机测试桥：宿主与会话就绪后登记。休眠态只做低频可用性探测，
    // 无活跃测试运行时零命令；服务不支持测试协议时安静退回，不影响正常使用。
    if (App.testBridge) { try { App.testBridge.start(); } catch (e) { console.warn('[test-bridge]', e); } }
  }

  initOffice();
})();
