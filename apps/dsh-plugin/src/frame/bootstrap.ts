import { createDshDesktopApi } from './dsh-api';
import { createFolderBridge } from './folder-bridge';
import { installDshAppMenu } from './app-menu';

const params = new URLSearchParams(location.search);
const isOfficeHost = params.get('hostTheme') === 'office';
if (isOfficeHost) {
  document.documentElement.dataset.hostTheme = 'office';
  // Paint the host chrome before the workbench bundle mounts, avoiding a light
  // flash while the typed desktop API resolves its initial system theme.
  document.documentElement.dataset.hostColorScheme =
    window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}
const lifecycle = new AbortController();
const api = await createDshDesktopApi({ openFolder: createFolderBridge(params, lifecycle.signal) });
const hostThemeSubscription = isOfficeHost ? api.app.onThemeChanged(theme => {
  document.documentElement.dataset.hostColorScheme = theme;
}) : { unsubscribe() {} };
if (isOfficeHost) void api.app.getTheme().then(result => {
  if (result.ok) document.documentElement.dataset.hostColorScheme = result.value;
});
Object.assign(window, { easyViewDesktop: api });
// Install before the original shell captures popupAppMenu for its button/Alt handler.
installDshAppMenu(api, {
  openFolder: () => api.workspace.openFolder(),
  switchWorkspace: () => api.workbench.pickWorkspace(),
  async showWelcome() { await workbench.ready; await workbench.flush(); workbench.showWelcome(); },
}, lifecycle.signal);
// Evaluate the original renderer only after the typed Host and locale exist.
const { createEasyViewWorkbench } = await import('@easyview/workbench-core');
// Keep the Office-only option additive so older shared workbench declarations
// remain consumable while the bundled renderer receives the system-following flag.
const workbenchOptions = { host: api, showAiChatOnWelcome: true, followSystemTheme: isOfficeHost,
  flush: api.workbench.flush, requestClose: api.workbench.requestClose };
const workbench = createEasyViewWorkbench(workbenchOptions);
await workbench.ready;
const channel = params.get('channel'), workbenchId = params.get('workbenchId');
const dispose = () => { workbench.dispose(); api.workbench.dispose(); hostThemeSubscription.unsubscribe(); lifecycle.abort(); };
if (window.parent !== window && channel && /^[a-f0-9-]{36}$/.test(channel) && workbenchId) {
  const send = (value: object) => window.parent.postMessage({ protocol: 'easyview.workbench.v1', channel, workbenchId, ...value }, location.origin);
  window.addEventListener('message', event => {
    const value = event.data;
    if (event.origin !== location.origin || event.source !== window.parent || !value || value.protocol !== 'easyview.workbench.v1'
      || value.channel !== channel || value.workbenchId !== workbenchId || typeof value.command !== 'string'
      || typeof value.requestId !== 'string' || !/^[a-f0-9-]{36}$/.test(value.requestId)) return;
    void (async () => {
      let result: unknown;
      switch (value.command) {
        case 'openDocument': if (typeof value.argument !== 'string') throw new Error('无效的文档路径'); await workbench.openDocument(value.argument); break;
        case 'flush': await workbench.flush(); break;
        case 'requestClose': result = await workbench.requestClose(); break;
        case 'suspend': workbench.suspend(); break;
        case 'resume': workbench.resume(); break;
        case 'dispose': dispose(); break;
        default: throw new Error('未知的工作台命令');
      }
      send({ type: 'result', requestId: value.requestId, ok: true, value: result });
    })().catch(error => send({ type: 'result', requestId: value.requestId, ok: false, error: String(error) }));
  }, { signal: lifecycle.signal });
  send({ type: 'ready' });
}
window.addEventListener('pagehide', dispose, { once: true });
