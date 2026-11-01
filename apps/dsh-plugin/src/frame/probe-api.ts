/** E01 shell probe. Unsupported commands fail visibly; this is not the document Provider. */
import { operationFailure, operationSuccess, DEFAULT_LOCAL_HISTORY_SETTINGS } from '@easyview/contracts';
import type { HostToEditorMessage, TimelineResultMessage } from '@easyview/contracts';
import type { EasyViewDesktopApi } from '@easyview/workbench-core/host';
import type { DesktopWorkspaceState, DesktopTabSnapshot } from '@easyview/workbench-core/contracts';
import type { DesktopThemeMode } from '@easyview/workbench-core/contracts';

export function createProbeDesktopApi(): EasyViewDesktopApi {
  const subscriptions = () => ({ unsubscribe() {}, dispose() {} });
  const unsupported = async () => operationFailure('DEPENDENCY_MISSING', '此开发预览尚未接通该宿主操作。');
  const notifyUnsupported = async () => { window.alert('此开发预览尚未接通该宿主操作。'); return unsupported(); };
  const editorListeners = new Set<(message: HostToEditorMessage) => void>();
  const timelineListeners = new Set<(message: TimelineResultMessage) => void>();
  const themeQuery = typeof window.matchMedia === 'function' ? window.matchMedia('(prefers-color-scheme: dark)') : null;
  const themeListeners = new Set<(theme: DesktopThemeMode) => void>();
  const currentTheme = (): DesktopThemeMode => themeQuery?.matches ? 'dark' : 'light';
  const emitTheme = () => { const theme = currentTheme(); for (const listener of themeListeners) listener(theme); };
  const subscribeTheme = (listener: (theme: DesktopThemeMode) => void) => {
    themeListeners.add(listener);
    if (themeListeners.size === 1 && themeQuery) {
      if (typeof themeQuery.addEventListener === 'function') themeQuery.addEventListener('change', emitTheme);
      else themeQuery.addListener(emitTheme);
    }
    return { unsubscribe() {
      themeListeners.delete(listener);
      if (!themeListeners.size && themeQuery) {
        if (typeof themeQuery.removeEventListener === 'function') themeQuery.removeEventListener('change', emitTheme);
        else themeQuery.removeListener(emitTheme);
      }
    } };
  };
  let state: DesktopWorkspaceState = { rootPath: null, openTabs: [], activeTabId: null,
    expandedRelativePaths: [], explorerVisible: true, outlineVisible: true, explorerWidth: 280, outlineWidth: 280 };
  const tabs: DesktopTabSnapshot = { windowId: 'dsh-probe', groups: [{ id: 'group-1', tabs: [], activeTabId: null }],
    activeGroupId: 'group-1', layout: { kind: 'group', groupId: 'group-1' }, tabs: [], activeTabId: null, splitRenderingAvailable: true };
  return {
    versionGuard: { request: unsupported, subscribe: subscriptions, onDiscovery: subscriptions },
    timeline: {
      post(message) { if (message.type === 'timeline.request') for (const listener of timelineListeners) listener({ type: 'timeline.result', id: message.id, error: 'Timeline 未适配' }); },
      subscribe(listener) { timelineListeners.add(listener); return { dispose: () => { timelineListeners.delete(listener); } }; },
      diagnostics: unsupported, comparison: unsupported,
      getSettings: async () => operationSuccess({ ...DEFAULT_LOCAL_HISTORY_SETTINGS, enabled: false }), setSettings: unsupported,
    },
    document: { open: notifyUnsupported, save: unsupported, saveAs: unsupported, rename: unsupported, onChanged: subscriptions },
    window: { requestClose: notifyUnsupported, titleBarDoubleClick: unsupported, popupAppMenu: notifyUnsupported, setTitleBarOverlay: unsupported },
    external: { open: notifyUnsupported }, clipboard: { writeText: async text => { await navigator.clipboard.writeText(text); return operationSuccess(true); } },
    app: { getTheme: async () => operationSuccess(currentTheme()), newWindow: notifyUnsupported, getZoomFactor: () => 1,
      locale: 'zh-CN', openedFileAtLaunch: false, onThemeChanged: subscribeTheme, onRevealDocument: subscriptions },
    workspace: {
      openFolder: notifyUnsupported, openRecentFolder: notifyUnsupported, getRecentFolders: async () => operationSuccess([]), removeRecentFolder: unsupported,
      onRootChanged: subscriptions, getState: async () => operationSuccess(state), setState: async next => { state = next; return operationSuccess(state); },
      readDirectory: unsupported, search: unsupported, getSortMode: async () => operationSuccess('name'), setSortMode: unsupported,
      getShowCreatedAt: async () => operationSuccess(false), setShowCreatedAt: unsupported,
      getShowUpdatedAt: async () => operationSuccess(false), setShowUpdatedAt: unsupported,
      getShowDotEntries: async () => operationSuccess(false), setShowDotEntries: unsupported,
      getShowTimestampHover: async () => operationSuccess(true), setShowTimestampHover: unsupported,
      create: unsupported, rename: unsupported, move: unsupported, reorder: unsupported, delete: unsupported,
      copyClipboard: unsupported, pasteClipboard: unsupported, importExternal: unsupported, getResourcePaths: unsupported,
      showContextMenu: notifyUnsupported, popupNewFileMenu: notifyUnsupported, openEntry: notifyUnsupported,
      onChanged: subscriptions, onContextCommand: subscriptions,
    },
    tabs: { getState: async () => operationSuccess(tabs), openWorkspaceEntry: unsupported, openPreview: unsupported,
      activate: unsupported, close: unsupported, closeOthers: unsupported, closeAll: unsupported,
      showContextMenu: unsupported, executeContextCommand: unsupported, onChanged: subscriptions },
    preview: { open: unsupported, close: unsupported, readText: unsupported, writeBytes: unsupported,
      sendHttp: unsupported, decompileJava: unsupported, onFileChanged: subscriptions },
    system: { openWithDefaultApp: notifyUnsupported, revealInFolder: notifyUnsupported },
    archive: { list: unsupported, readEntry: unsupported, exportEntry: unsupported },
    menu: { publishState() {}, onCommand: subscriptions, onRevealInIndexArea: subscriptions },
    editor: {
      postMessage(message) {
        if (message.type === 'aiChat.send') for (const listener of editorListeners) listener({ type: 'aiChat.error', requestId: message.requestId, message: 'DSH AI 执行链尚未接入，此预览不会调用模型。' });
      },
      subscribe(listener) { editorListeners.add(listener); return { unsubscribe: () => { editorListeners.delete(listener); } }; },
    },
  };
}
