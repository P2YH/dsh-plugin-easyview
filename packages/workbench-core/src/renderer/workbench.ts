import { VersionGuardPanel } from '@easyview/editor-core/version-guard';
import { maxPanelWidth } from '@easyview/editor-core/workspace';
import type { GuardUiState } from '@easyview/contracts';
import './bootLocale';
import { TimelinePanel, RpcTimelineClient, showHistorySettingsDialog } from '@easyview/editor-core/timeline';
import type { TimelineUiState } from '@easyview/contracts';
import type { HostToEditorMessage, EditorToHostMessage } from '@easyview/contracts';
import { isMarkdownFilePath, readLocalProductTheme } from '@easyview/contracts';
import { applyDesktopProductAppearance } from './productThemeControl';
import { AI_CHAT_ICON_HTML, createEasyViewEditor, createLazyAiChatPanel, EditorAppearanceStore, type AiChatPanel, type EasyViewEditorInstance } from '@easyview/editor-core';
import { defaultDesktopPanelLayout, isDesktopAppearanceCommand, normalizeDesktopPanelLayout } from '../contracts';
import type { DesktopAppearanceMode, DesktopEditorTab, DesktopPanelLayout, DesktopTab, DesktopTabSnapshot, DesktopThemeMode } from '../contracts';
import { applyTitleBarToggles, applyWorkbenchPanelLayout, installPanelDock, type PanelDockElements, type TitleBarToggleElements } from './panelDock';
import type { EasyViewDesktopApi } from '../host';
import type { EasyViewAccentTheme } from '@easyview/editor-core';
import { createDesktopEditorHostTransport } from './desktopHostTransport';
import { DesktopEditorInstanceRegistry } from './DesktopEditorInstanceRegistry';
import { WorkspaceExplorer } from './workspace/WorkspaceExplorer';
import { ActiveViewController } from './ActiveViewController';
import { isAgentEditablePreview } from './officePreview';
import { TabController } from './TabController';
import { WelcomePage } from './WelcomePage';
import { tDesktop } from './desktopRenderer.i18n';
import { readToolbarPinned, subscribeToolbarVisibility } from '@easyview/preview-ui/toolbar';
import { installWindowChrome } from './windowChrome';
import { installAppMenuButton } from './appMenuButton';
import { installTitleBarDoubleClick } from './titleBarDoubleClick';
import { installTitleBarOverlaySync } from './titleBarOverlaySync';
import { installToolbarWheelScroll } from './toolbarWheelScroll';
import '@easyview/preview-ui/toolbar-styles';

interface WindowWithDesktopApi extends Window {
  easyViewDesktop: EasyViewDesktopApi;
  systemLocale?: string;
  csvDelimiterSetting?: string;
  __easyviewPdfFonts?: {
    normal: string;
    bold: string;
    symbols: string;
  };
}

declare const window: WindowWithDesktopApi;

const mountedDocuments = new WeakSet<Document>();

/** Mount the original shell explicitly; each frame owns one complete workbench lifetime. */
export function mountDesktopWorkbench(api: EasyViewDesktopApi, options: {
  showAiChatOnWelcome?: boolean;
  followSystemTheme?: boolean;
} = {}) {
if (mountedDocuments.has(document)) throw new Error('A workbench is already mounted in this document');
mountedDocuments.add(document);
const lifetime = new AbortController();
let disposed = false;
if (!api?.editor || !api.app || !api.document) {
  document.body.innerHTML = '<main style="padding:24px;font-family:system-ui">EasyView_Md 宿主接口加载失败。Failed to load the EasyView_Md host API.</main>'; // i18n-ignore
  throw new Error('EasyView_Md preload API is unavailable');
}
document.body.classList.toggle('desktop-platform-other', !navigator.platform.toLowerCase().includes('mac'));

const sharedEditorHost = createDesktopEditorHostTransport({
  capabilities: api.capabilities?.editorHost,
  postMessage: (message: EditorToHostMessage) => api.editor.postMessage(message),
  subscribe: (listener: (message: HostToEditorMessage) => void) =>
    api.editor.subscribe(listener),
});

const productAppearance = new EditorAppearanceStore(
  window.localStorage,
  window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light',
);
applyDesktopProductAppearance(productAppearance.state, document);

function applyOsTheme(theme: DesktopThemeMode): void {
  const systemTheme = theme === 'dark' ? 'dark' : 'light';
  productAppearance.setMode(options.followSystemTheme ? systemTheme : readLocalProductTheme(window.localStorage, systemTheme), false);
}

window.systemLocale = navigator.language || 'en-US';
window.csvDelimiterSetting = 'auto';
window.__easyviewPdfFonts = {
  normal: new URL('./fonts/SourceHanSansCN-Normal.otf', window.location.href).toString(),
  bold: new URL('./fonts/SourceHanSansCN-Heavy.otf', window.location.href).toString(),
  symbols: new URL('./fonts/NotoSansSymbols2-Regular.ttf', window.location.href).toString(),
};

function localizeShell(): void {
  const aiChatToggleButton = document.getElementById('ai-chat-toggle');
  if (aiChatToggleButton) aiChatToggleButton.innerHTML = AI_CHAT_ICON_HTML;
  aiChatToggleButton?.setAttribute('title', tDesktop('shell.aiChatToggle'));
  aiChatToggleButton?.setAttribute('aria-label', tDesktop('shell.aiChat'));
  document.getElementById('desktop-tab-bar')?.setAttribute('aria-label', tDesktop('shell.openTabs'));
  document.getElementById('desktop-welcome')?.setAttribute('aria-label', tDesktop('shell.welcome'));
  document.getElementById('desktop-ai-chat-host')?.setAttribute('aria-label', tDesktop('shell.aiChat'));
  const guardToggle = document.getElementById('version-guard-toggle');
  guardToggle?.setAttribute('title', tDesktop('dock.panel.versionGuard'));
  guardToggle?.setAttribute('aria-label', tDesktop('dock.panel.versionGuard'));
  document.getElementById('version-guard-host')?.setAttribute('aria-label', tDesktop('dock.panel.versionGuard'));
}
localizeShell();

// Keep native window controls and the tab strip aligned when the document UI is zoomed.
const syncTitleBarZoom = (): void => {
  document.documentElement.style.setProperty('--desktop-page-zoom', String(api.app.getZoomFactor()));
};
window.addEventListener('resize', syncTitleBarZoom, { signal: lifetime.signal });
syncTitleBarZoom();
const explorerElement = document.getElementById('workspace-explorer')!;
const tabBar = document.getElementById('desktop-tab-bar')!;
const editorMount = document.getElementById('desktop-editor-mount')!;
const previewBody = document.getElementById('file-preview-root')!;
const splitRoot = document.getElementById('desktop-split-root')!;
let currentTab: DesktopTab | null = null;
let tabController: TabController;
let lastOutlineVisible = true;
let lastOutlineWidth = 280;
let panelLayout: DesktopPanelLayout = defaultDesktopPanelLayout();
const panelDockElements: PanelDockElements = {
  centerPane: document.getElementById('desktop-center-pane')!,
  panels: { explorer: explorerElement, aiChat: document.getElementById('desktop-ai-chat-host')!, versionGuard: document.getElementById('version-guard-host')! },
};
let editorStateSubscription: { unsubscribe(): void } | undefined;
const previewApi = api;
const editors = new DesktopEditorInstanceRegistry({
  mount: editorMount,
  capabilities: sharedEditorHost.capabilities,
  postMessage: (message) => sharedEditorHost.postMessage(message),
  subscribe: (listener) => sharedEditorHost.subscribe(listener),
  createEditor: ({ root, host, initialMessage }) => createEasyViewEditor({
    host,
    hostActions: { registerDocumentEditor: api.documentEditors?.register },
    initialMessage,
    uiMode: 'desktop',
    appearance: productAppearance,
    outlinePosition: panelLayout.outline,
    aiChatContainer: panelDockElements.panels.aiChat,
    root,
  }),
});

function activeEditor(): EasyViewEditorInstance | null {
  return editors.active();
}

const titleBarToggleElements: TitleBarToggleElements = {
  left: document.getElementById('desktop-title-left')!,
  right: document.getElementById('desktop-title-actions')!,
  toggles: {
    explorer: document.getElementById('workspace-toggle')!,
    aiChat: document.getElementById('ai-chat-toggle')!,
    versionGuard: document.getElementById('version-guard-toggle')!,
  },
};

function applyPanelLayout(layout: DesktopPanelLayout): void {
  panelLayout = layout;
  applyWorkbenchPanelLayout(panelDockElements, layout);
  applyTitleBarToggles(layout, titleBarToggleElements);
  for (const editor of editors.all()) editor.setOutlinePosition(layout.outline);
}
applyPanelLayout(panelLayout);

let aiChatVisible = false;

function updateAiChatVisible(visible: boolean): void {
  if (visible !== aiChatVisible) {
    aiChatVisible = visible;
    workspaceExplorer.setAiChatVisible(visible);
  }
  syncAiChatToggle(visible);
}

// Serves the AI chat while no Markdown editor is active, so it does not depend on an open file.
let standaloneAiChat: AiChatPanel | undefined;
let standaloneAiSubscription: { unsubscribe(): void } | undefined;
function getStandaloneAiChat(): AiChatPanel {
  if (standaloneAiChat) return standaloneAiChat;
  const panel = createLazyAiChatPanel({
    historyProvider: sharedEditorHost.capabilities.aiChatHistory,
    postMessage: (message) => sharedEditorHost.postMessage(message),
    container: panelDockElements.panels.aiChat,
    getDocumentContent: () => '',
    getFilePath: () => activeOfficePreview()?.relativePath ?? '',
    getFileName: () => activeOfficePreview()?.fileName ?? '',
    applyDocumentEdit: () => {},
    copyText: (text, successMessage) => sharedEditorHost.postMessage({ type: 'copyTextToClipboard', text, successMessage }),
    onVisibilityChange: (visible) => { if (!activeEditor()) updateAiChatVisible(visible); },
  });
  panel.setDocumentContextAvailable(Boolean(activeOfficePreview()));
  standaloneAiSubscription = sharedEditorHost.subscribe((message) => {
    if (!message.type.startsWith('aiChat.')) return;
    panel.handleMessage(message);
    if ((message.type === 'aiChat.officeCommitted' || (message.type === 'aiChat.officeUndone' && !message.error)) && message.filePath) {
      void activeView.reloadPreviewMatching(message.filePath);
    }
  });
  standaloneAiChat = panel;
  // The preview may already be active when its lazy AI panel is first opened.
  // Initialize the same conversation path used by later view-change events.
  syncStandaloneOfficeContext();
  return panel;
}

/** The focused preview tab when Agent mode can edit it (Phase 1: .xlsx only). */
function activeOfficePreview(): { relativePath: string; fileName: string } | null {
  const view = activeView.getState();
  if (view.kind !== 'preview' || !isAgentEditablePreview(view.relativePath)) return null;
  return { relativePath: view.relativePath, fileName: view.relativePath.split(/[\\/]/).pop() ?? view.relativePath };
}

function syncStandaloneOfficeContext(): void {
  if (!standaloneAiChat) return;
  const office = activeOfficePreview();
  standaloneAiChat.setFilePath(office ? `preview:${office.relativePath}` : '');
  standaloneAiChat.setDocumentContextAvailable(Boolean(office));
}

function syncStandaloneAiChat(): void {
  const hasEditor = Boolean(activeEditor());
  if (!standaloneAiChat && (hasEditor || !aiChatVisible)) return;
  const panel = getStandaloneAiChat();
  panel.setSuspended(hasEditor);
  if (!hasEditor) { if (aiChatVisible) panel.open(); else panel.close(); }
}

function bindActiveEditor(editor: EasyViewEditorInstance | null): void {
  editorStateSubscription?.unsubscribe();
  // Apply the persisted outline layout before subscribing: subscribeUiState emits
  // the editor's current state immediately, and a freshly created editor starts
  // with its outline closed, which would otherwise overwrite the saved preference.
  if (editor) {
    editor.setOutlineVisible(lastOutlineVisible);
    editor.setOutlineWidth(lastOutlineWidth);
    editor.setAiChatVisible(aiChatVisible);
  }
  syncStandaloneAiChat();
  syncAiChatToggle(aiChatVisible);
  editorStateSubscription = editor?.subscribeUiState((state) => {
    workspaceExplorer.setOutlineVisible(state.outlineVisible);
    updateAiChatVisible(state.aiChatVisible);
    publishMenuState();
  });
}

/** Hosts can keep AI accessible on the welcome page; visibility still follows the user's toggle. */
let aiChatOnWelcome = options.showAiChatOnWelcome === true;
function aiChatSuppressed(): boolean {
  return welcomeCoversWorkspace && !aiChatOnWelcome;
}

function toggleAiChat(): void {
  if (aiChatSuppressed()) {
    aiChatOnWelcome = true;
    document.body.classList.add('desktop-ai-chat-on-welcome');
    if (aiChatVisible) { syncAiChatToggle(true); return; }
  }
  const editor = activeEditor();
  if (editor) { editor.setAiChatVisible(!editor.getUiState().aiChatVisible); return; }
  updateAiChatVisible(!aiChatVisible);
  syncStandaloneAiChat();
}

function syncAiChatToggle(requested: boolean): void {
  const visible = requested && !aiChatSuppressed();
  const toggle = document.getElementById('ai-chat-toggle');
  toggle?.classList.toggle('active', visible);
  toggle?.setAttribute('aria-pressed', String(visible));
}

const timelineClient = new RpcTimelineClient({
  configureSettings: () => showHistorySettingsDialog(async () => { const result = await api.timeline.getSettings(); if (!result.ok) throw new Error(result.message); return result.value; }, async (settings) => { const result = await api.timeline.setSettings(settings); if (!result.ok) throw new Error(result.message); }),
  post: (message) => api.timeline.post(message as import('@easyview/contracts').TimelineRequestMessage),
  subscribe: (listener) => api.timeline.subscribe(listener),
  loadState: () => { const raw = localStorage.getItem('easyview.timeline.state'); return raw ? JSON.parse(raw) as TimelineUiState : undefined; },
  saveState: (state) => localStorage.setItem('easyview.timeline.state', JSON.stringify(state)),
});
const timelinePanel = new TimelinePanel(timelineClient);
const welcomePage = new WelcomePage({
  api,
  container: document.getElementById('desktop-welcome')!,
  onOpenFolder: () => { void workspaceExplorer.openFolder().then((opened) => { if (opened) dismissStartupWelcome(); }); },
  onOpenRecentFolder: (folderPath) => { void workspaceExplorer.openRecentFolder(folderPath).then((opened) => { if (opened) dismissStartupWelcome(); }); },
  onResume: () => dismissStartupWelcome(),
  onShowWelcome: () => {
    welcomeRequested = true;
    if (lastTabSnapshot) syncWelcome(lastTabSnapshot);
  },
});
/** Launch shows the welcome page over restored tabs until the user picks something. */
let startupWelcome = !api.app.openedFileAtLaunch;
/** With a folder open and no tab, the pane shows a light hint unless the user asks for the full welcome page. */
let welcomeRequested = false;
let workspaceRootPath: string | null = null;

/** The full welcome page hides the restored folder until the user resumes or picks one. */
let welcomeCoversWorkspace = false;

function setWorkspaceRootPath(rootPath: string | null): void {
  if (workspaceRootPath === rootPath) return;
  workspaceRootPath = rootPath;
  guardPanel.reset();
  welcomeRequested = false;
  if (lastTabSnapshot) syncWelcome(lastTabSnapshot);
  else applyGuardVisibility();
}
let startupTabsKey: string | null = null;
let lastTabSnapshot: DesktopTabSnapshot | null = null;

function syncWelcome(snapshot: DesktopTabSnapshot): void {
  lastTabSnapshot = snapshot;
  const tabs = snapshot.groups.flatMap((group) => group.tabs);
  const tabCount = Math.max(tabs.length, snapshot.tabs.length);
  if (startupWelcome) {
    const key = `${snapshot.activeTabId}|${tabs.map((item) => item.id).join(',')}`;
    if (startupTabsKey === null) startupTabsKey = key;
    else if (key !== startupTabsKey) startupWelcome = false;
  }
  const showWelcome = tabCount === 0 || startupWelcome;
  if (tabCount > 0) welcomeRequested = false;
  const fullWelcome = (startupWelcome && tabCount > 0) || !workspaceRootPath || welcomeRequested;
  welcomePage.setMode(fullWelcome ? 'welcome' : 'empty');
  const active = snapshot.tabs.find((item) => item.id === snapshot.activeTabId) ?? tabs.find((item) => item.id === snapshot.activeTabId);
  welcomePage.setResume(startupWelcome && tabCount > 0
    ? { folderPath: workspaceRootPath, tabCount, activeFileName: active?.fileName ?? null }
    : null);
  welcomePage.setCurrentRoot(startupWelcome ? null : workspaceRootPath);
  splitRoot.hidden = showWelcome;
  welcomeCoversWorkspace = showWelcome && fullWelcome;
  if (!welcomeCoversWorkspace) aiChatOnWelcome = options.showAiChatOnWelcome === true;
  document.body.classList.toggle('desktop-welcome-active', showWelcome);
  document.body.classList.toggle('desktop-welcome-full', welcomeCoversWorkspace);
  document.body.classList.toggle('desktop-ai-chat-on-welcome', aiChatOnWelcome);
  syncAiChatToggle(aiChatVisible);
  welcomePage.setVisible(showWelcome);
  applyGuardVisibility();
}

function dismissStartupWelcome(): void {
  if (!startupWelcome) return;
  startupWelcome = false;
  if (lastTabSnapshot) syncWelcome(lastTabSnapshot);
}

const revealSubscription = api.app.onRevealDocument(() => dismissStartupWelcome());

window.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape' || !welcomePage.isVisible()) return;
  if (welcomeRequested) {
    event.preventDefault();
    welcomeRequested = false;
    if (lastTabSnapshot) syncWelcome(lastTabSnapshot);
    return;
  }
  if (!startupWelcome) return;
  event.preventDefault();
  dismissStartupWelcome();
}, { signal: lifetime.signal });
const workspaceExplorer = new WorkspaceExplorer({
  footer: timelinePanel.element,
  onSelectionChanged: (filePath, visible) => { timelinePanel.setVisible(visible); timelinePanel.setFile(filePath); },
  api,
  container: explorerElement,
  onStateChange: (state) => {
    setWorkspaceRootPath(state.rootPath);
    lastOutlineVisible = state.outlineVisible;
    lastOutlineWidth = state.outlineWidth;
    activeEditor()?.setOutlineVisible(state.outlineVisible);
    document.getElementById('workspace-toggle')?.classList.toggle('active', state.explorerVisible);
    publishMenuState();
  },
  onOpenFile: async (relativePath) => {
    const opened = await tabController.openWorkspaceEntry(relativePath);
    if (opened) dismissStartupWelcome();
    return opened;
  },
  onOpenWithDefaultApp: (relativePath) => { void previewApi.system.openWithDefaultApp(relativePath); },
  onRevealInFolder: (relativePath) => { void previewApi.system.revealInFolder(relativePath); },
});

const guardHost = panelDockElements.panels.versionGuard;
let guardVisible = false;
let guardWidth = 360;
const guardPanel = new VersionGuardPanel({
  grip: 'versionGuard',
  client: {
    onDiscovery: listener => api.versionGuard.onDiscovery(listener),
    request: async request => { const result = await api.versionGuard.request(request); if (!result.ok) throw new Error(result.message); return result.value; },
    subscribe: listener => { const metadata = api.versionGuard.subscribe(listener); const files = api.workspace.onChanged(paths => listener(paths?.map(file => `${workspaceRootPath}/${file}`) ?? null)); return { dispose: () => { metadata.dispose(); files.unsubscribe(); } }; },
    loadState: () => { try { return JSON.parse(localStorage.getItem('easyview.versionGuard.ui') ?? 'null') as GuardUiState | undefined; } catch { return undefined; } },
    saveState: state => localStorage.setItem('easyview.versionGuard.ui', JSON.stringify(state)),
  },
});
guardHost.append(guardPanel.element);
const guardResizer = document.createElement('div'); guardResizer.className = 'guard-width-resizer'; guardHost.append(guardResizer);
/**
 * The panel only shows for an open workspace that the full welcome page is not covering, so the
 * welcome page never lists repositories of a previously opened folder.
 */
function applyGuardVisibility(): void {
  const shown = guardVisible && workspaceRootPath !== null && !welcomeCoversWorkspace;
  guardHost.hidden = !shown; guardHost.style.width = `${guardWidth}px`; guardPanel.setVisible(shown);
}
function setGuardVisible(visible: boolean, persist = true): void {
  guardVisible = visible; applyGuardVisibility();
  const toggle = document.getElementById('version-guard-toggle')!; toggle.classList.toggle('active', visible); toggle.setAttribute('aria-pressed', String(visible));
  if (persist) workspaceExplorer.setVersionGuardState(visible, guardWidth);
}
document.getElementById('version-guard-toggle')!.onclick = () => { dismissStartupWelcome(); setGuardVisible(!guardVisible); };
guardResizer.onpointerdown = event => {
  event.preventDefault(); guardHost.classList.add('resizing'); const start = event.clientX; const width = guardWidth; const max = maxPanelWidth(guardHost, 280);
  const move = (next: PointerEvent) => { guardWidth = Math.max(280, Math.min(max, width + (next.clientX - start) * (guardHost.dataset.dockSide === 'right' ? -1 : 1))); guardHost.style.width = `${guardWidth}px`; };
  const end = () => { guardHost.classList.remove('resizing'); window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', end); workspaceExplorer.setVersionGuardState(guardVisible, guardWidth); };
  window.addEventListener('pointermove', move, { signal: lifetime.signal }); window.addEventListener('pointerup', end, { once: true, signal: lifetime.signal });
};

const activeView = new ActiveViewController({
  api, splitRoot, editorBody: editorMount, previewBody, tabBar,
  onViewChanged: () => { publishMenuState(); syncStandaloneOfficeContext(); },
});
const fileChangedSubscription = api.preview.onFileChanged((relativePaths) => {
  for (const relativePath of relativePaths) void activeView.reloadPreviewMatching(relativePath);
});

function visibleEditorTabs(snapshot: { groups: Array<{ tabs: DesktopTab[]; activeTabId: string | null }> }): DesktopEditorTab[] {
  const visible: DesktopEditorTab[] = [];
  for (const group of snapshot.groups) {
    const tab = group.tabs.find((item) => item.id === group.activeTabId);
    if (tab?.kind === 'editor') visible.push(tab);
  }
  return visible;
}

function bindVisibleEditors(snapshot: { groups: Array<{ id: string; tabs: DesktopTab[]; activeTabId: string | null }> }, focused: DesktopTab | null): void {
  const visible = visibleEditorTabs(snapshot);
  const groupMounts = activeView.editorMounts();
  const mounts = new Map<string, HTMLElement>();
  for (const group of snapshot.groups) {
    const mount = groupMounts.get(group.id);
    const tab = group.tabs.find((item) => item.id === group.activeTabId);
    if (mount && tab?.kind === 'editor') mounts.set(tab.id, mount);
  }
  editors.setVisible(visible, mounts);
  if (focused?.kind === 'editor') {
    editors.focus(focused.id);
    bindActiveEditor(editors.active());
  } else {
    editors.clearActive();
    bindActiveEditor(null);
  }
  editors.syncOpenTabs(snapshot.groups.flatMap((group) => group.tabs).filter((item) => item.kind === 'editor').map((item) => item.id));
}

tabController = new TabController({
  api,
  onActiveTabChanged: (tab, snapshot) => {
    currentTab = tab;
    syncWelcome(snapshot);
    document.body.classList.toggle('desktop-history-only', tab?.kind === 'history' && visibleEditorTabs(snapshot).length === 0);
    void activeView.showSnapshot(snapshot);
    bindVisibleEditors(snapshot, tab);
    if (!tab) workspaceExplorer.setActiveRelativePath(null);
    else if (tab.kind === 'editor') workspaceExplorer.setActiveDocument(tab.filePath);
    else if (tab.kind !== 'history') workspaceExplorer.setActiveRelativePath(tab.relativePath);
    publishMenuState();
  },
});

function publishMenuState(): void {
  const state = activeEditor()?.getUiState() ?? {
    sourceMode: false,
    outlineVisible: lastOutlineVisible,
    fullWidth: true,
    tableWrap: false,
    viewChanges: false,
  };
  const appearance = productAppearance.state;
  api.menu.publishState({
    ...state,
    hasActiveDocument: currentTab?.kind === 'editor',
    isMarkdownDocument: currentTab?.kind === 'editor' && (!currentTab.filePath || isMarkdownFilePath(currentTab.filePath)),
    themeMode: appearance.mode,
    themeDepth: Math.round(appearance.depth * 100),
    accent: appearance.accent,
  });
}

document.getElementById('workspace-toggle')?.addEventListener('click', () => workspaceExplorer.toggle(), { signal: lifetime.signal });
function syncToolbarVisibility(): void {
  const visible = readToolbarPinned();
  document.body.classList.toggle('desktop-toolbar-hidden', !visible);
}
const unsubscribeToolbarVisibility = subscribeToolbarVisibility(syncToolbarVisibility);
syncToolbarVisibility();
const disposeAppMenuButton = document.body.classList.contains('desktop-platform-other')
  ? installAppMenuButton({ container: titleBarToggleElements.left, label: tDesktop('shell.appMenu'), popup: api.window.popupAppMenu })
  : () => undefined;
const disposeTitleBarOverlaySync = document.body.classList.contains('desktop-platform-other')
  ? installTitleBarOverlaySync({
    source: titleBarToggleElements.right,
    apply: api.window.setTitleBarOverlay,
    subscribe: (listener) => {
      const subscription = productAppearance.subscribe(listener);
      // The title-control background switches with the welcome page; repaint the native caption buttons too.
      const welcomeObserver = new MutationObserver(listener);
      welcomeObserver.observe(document.body, { attributes: true, attributeFilter: ['class'] });
      return () => { subscription.unsubscribe(); welcomeObserver.disconnect(); };
    },
  })
  : () => undefined;
const disposeTitleBarDoubleClick = document.body.classList.contains('desktop-platform-other')
  ? () => undefined
  : installTitleBarDoubleClick({
    root: document.getElementById('desktop-root')!,
    chromeHeight: () => 32 / (Number(document.documentElement.style.getPropertyValue('--desktop-page-zoom')) || 1),
    onDoubleClick: () => { void api.window.titleBarDoubleClick(); },
  });
const disposeToolbarWheelScroll = installToolbarWheelScroll(document);
const disposeWindowChrome = installWindowChrome({
  workbench: document.getElementById('desktop-workbench')!,
  leading: titleBarToggleElements.left,
  trailing: titleBarToggleElements.right,
});
const aiChatToggle = document.getElementById('ai-chat-toggle');
aiChatToggle?.addEventListener('click', toggleAiChat, { signal: lifetime.signal });
const documentChangedSubscription = api.document.onChanged((documentState) => {
  if (currentTab?.kind !== 'editor') return;
  workspaceExplorer.setActiveDocument(documentState.filePath);
  publishMenuState();
});

window.addEventListener('easyview-toc-width-change', (event) => {
  const width = (event as CustomEvent<number>).detail;
  if (!Number.isFinite(width)) return;
  void api.workspace.getState().then((result) => {
    if (!disposed && result.ok) void api.workspace.setState({ ...result.value, outlineWidth: width });
  });
}, { signal: lifetime.signal });

const indexRevealSubscription = api.menu.onRevealInIndexArea((relativePath) => {
  void workspaceExplorer.revealRelativePath(relativePath);
});

const menuSubscription = api.menu.onCommand((command) => {
  if (command === 'toggleTheme') {
    productAppearance.cycleMode();
    return;
  }
  if (isDesktopAppearanceCommand(command)) {
    const [kind, value] = command.split(':');
    if (kind === 'setThemeMode') productAppearance.setMode(value as DesktopAppearanceMode);
    else if (kind === 'setThemeDepth') productAppearance.setDepth(Number(value) / 100);
    else productAppearance.setAccent(value as EasyViewAccentTheme);
    return;
  }
  if (command === 'toggleExplorer') {
    workspaceExplorer.toggle();
    return;
  }
  if (command === 'rename') {
    const currentName = currentTab?.fileName ?? '';
    const nextName = window.prompt(tDesktop('rename.prompt'), currentName);
    if (nextName?.trim() && currentTab?.kind === 'editor') {
      void api.document.rename(nextName.trim(), currentTab.id);
    }
    return;
  }
  if (command === 'toggleAiChat') {
    toggleAiChat();
    return;
  }
  activeEditor()?.executeCommand(command);
});

const appearanceSubscription = productAppearance.subscribe((state) => {
  applyDesktopProductAppearance(state, document);
  publishMenuState();
});

const disposePanelDock = installPanelDock({
  ...panelDockElements,
  getLayout: () => panelLayout,
  onLayoutChange: (layout) => {
    applyPanelLayout(layout);
    workspaceExplorer.setPanelLayout(layout);
  },
});

const ready = workspaceExplorer.initialize().then(async (state) => {
  if (disposed) return;
  applyPanelLayout(normalizeDesktopPanelLayout(state.panelLayout));
  guardWidth = state.versionGuardWidth ?? 360;
  setGuardVisible(state.versionGuardVisible === true, false);
  setWorkspaceRootPath(state.rootPath);
  lastOutlineVisible = state.outlineVisible;
  lastOutlineWidth = state.outlineWidth;
  document.getElementById('workspace-toggle')?.classList.toggle('active', state.explorerVisible);
  activeEditor()?.setOutlineVisible(state.outlineVisible);
  activeEditor()?.setOutlineWidth(state.outlineWidth);
  aiChatVisible = state.aiChatVisible === true;
  activeEditor()?.setAiChatVisible(aiChatVisible);
  syncStandaloneAiChat();
  syncAiChatToggle(aiChatVisible);
  await tabController.initialize();
  if (disposed) return;
  publishMenuState();
});

const themeSubscription = api.app.onThemeChanged(applyOsTheme);
void api.app.getTheme().then((result) => {
  if (!disposed && result.ok) applyOsTheme(result.value);
});

function dispose(): void {
  if (disposed) return;
  disposed = true; lifetime.abort(); mountedDocuments.delete(document);
  revealSubscription.unsubscribe(); fileChangedSubscription.unsubscribe(); documentChangedSubscription.unsubscribe();
  indexRevealSubscription.unsubscribe(); menuSubscription.unsubscribe(); standaloneAiSubscription?.unsubscribe();
  standaloneAiChat?.destroy();
  document.getElementById('version-guard-toggle')!.onclick = null; guardResizer.onpointerdown = null;
  themeSubscription.unsubscribe();
  appearanceSubscription.unsubscribe();
  editorStateSubscription?.unsubscribe();
  disposePanelDock();
  disposeWindowChrome();
  disposeAppMenuButton();
  disposeTitleBarOverlaySync();
  disposeTitleBarDoubleClick();
  disposeToolbarWheelScroll();
  unsubscribeToolbarVisibility();
  guardPanel.dispose();
  tabController.dispose();
  workspaceExplorer.dispose();
  activeView.dispose();
  editors.dispose();
  timelinePanel.dispose(); timelineClient.dispose();
}

window.addEventListener('unload', dispose, { once: true, signal: lifetime.signal });
window.addEventListener('pagehide', dispose, { once: true, signal: lifetime.signal });
return { ready, dispose,
  showWelcome() {
    if (disposed) return;
    startupWelcome = true; startupTabsKey = null; welcomeRequested = true;
    if (lastTabSnapshot) syncWelcome(lastTabSnapshot);
  },
  suspend() { for (const editor of editors.all()) editor.setDocumentActive(false); standaloneAiChat?.setSuspended(true); },
  resume() { if (disposed) return; activeEditor()?.setDocumentActive(true); syncStandaloneAiChat(); window.dispatchEvent(new Event('resize')); } };
}
