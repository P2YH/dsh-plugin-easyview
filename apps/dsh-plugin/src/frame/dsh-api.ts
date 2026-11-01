/** Typed adapter for the original shell. Mutations remain unavailable until durable editing is admitted. */
import { operationSuccess, operationFailure, isMarkdownFilePath, type OperationResult, type PreviewDescriptor, type HostToEditorMessage, type WorkspaceEntry } from '@easyview/contracts';
import type { EasyViewDesktopApi } from '@easyview/workbench-core/host';
import type { DesktopTabSnapshot, DesktopWorkspaceState, WorkspaceContextCommand } from '@easyview/workbench-core/contracts';
import type { EasyViewBootstrap } from '../shared/contracts';
import { createProbeDesktopApi } from './probe-api';
import { createDraftOwner } from './draft-owner';
import { createTextOwner } from './text-owner';
import { createAiOwner } from './ai-owner';
import { createFileHistoryOwner } from './file-history-owner';

export type DshDesktopApi = EasyViewDesktopApi & { workbench: {
  flush(): Promise<void>; openDocument(path: string): Promise<void>; requestClose(): Promise<boolean>; dispose(): void;
  pickWorkspace(): ReturnType<EasyViewDesktopApi['workspace']['openFolder']>;
} };
export async function createDshDesktopApi(options: { openFolder?: () => Promise<string | null> } = {}): Promise<DshDesktopApi> {
  const api = createProbeDesktopApi();
  const bootstrap: EasyViewBootstrap = await (await fetch('/api/easyview/v1/bootstrap', { credentials: 'same-origin' })).json();
  if (bootstrap.apiVersion !== 1) throw new Error('EasyView protocol mismatch');
  const workbenchId = new URLSearchParams(location.search).get('workbenchId') ?? (() => {
    const key = 'easyview:direct-workbench';
    let id = sessionStorage.getItem(key); if (!id) { id = crypto.randomUUID(); sessionStorage.setItem(key, id); } return id;
  })();
  if (!/^[a-f0-9-]{36}$/.test(workbenchId)) throw new Error('Invalid workbench identity');
  const drafts = createDraftOwner(bootstrap, workbenchId);
  const editorListeners = new Set<(message: HostToEditorMessage) => void>();
  const emit = (message: HostToEditorMessage) => { for (const listener of editorListeners) listener(message); };
  const text = createTextOwner(bootstrap, workbenchId, emit, drafts.show);
  const fallbackEditor = api.editor;
  api.editor = {
    postMessage(message) {
      if (ai.handle(message)) return;
      if (text.handle(message)) return;
      if (message.type === 'save') {
        const active = tabs.tabs.find(tab => tab.id === tabs.activeTabId);
        if (active?.kind === 'editor' && active.documentId) void text.save(active.documentId).catch(error => drafts.show('failed', String(error)));
        return;
      }
      return fallbackEditor.postMessage(message);
    },
    subscribe(listener) { editorListeners.add(listener); const original = fallbackEditor.subscribe(listener);
      return { unsubscribe() { editorListeners.delete(listener); original.unsubscribe(); } }; },
  };
  api.documentEditors = { ...drafts, register(editor) {
    return text.has(editor.snapshot().documentId) ? text.register(editor) : drafts.register(editor);
  } };
  window.addEventListener('pagehide', () => { text.dispose(); drafts.dispose(); }, { once: true });
  const flush = async () => {
    try { await text.flush(); await drafts.flush(); return true; }
    catch (error) { window.alert(`恢复草稿尚未确认，已取消当前导航：${String(error)}`); return false; }
  };
  const rootListeners = new Set<(state: DesktopWorkspaceState) => void>();
  const fileListeners = new Set<(paths: string[] | null) => void>();
  const contextListeners = new Set<(command: WorkspaceContextCommand, path: string, detail?: string) => void>();
  const tabListeners = new Set<(state: DesktopTabSnapshot) => void>();
  let workspaceId = '';
  let state: DesktopWorkspaceState = { rootPath: null, openTabs: [], activeTabId: null,
    expandedRelativePaths: [], explorerVisible: true, outlineVisible: true, aiChatVisible: true, explorerWidth: 280, outlineWidth: 280 };
  const tabs: DesktopTabSnapshot = { windowId: crypto.randomUUID(), groups: [], activeGroupId: 'group-1',
    layout: { kind: 'group', groupId: 'group-1' }, tabs: [], activeTabId: null, splitRenderingAvailable: false };
  const descriptors = new Map<string, PreviewDescriptor>();
  const activePath = () => {
    const tab = tabs.tabs.find(tab => tab.id === tabs.activeTabId);
    return tab?.kind === 'preview' ? tab.relativePath : tab?.kind === 'editor' && tab.documentId
      ? [...descriptors.values()].find(value => value.sessionId === tab.documentId)?.relativePath ?? '' : '';
  };
  const officeUndoRequests = new Map<string, string>();
  const ai = createAiOwner({ emit, filePath: activePath,
    async prepareContext(path) {
      const targetPath = (path || activePath()).replace(/^preview:/, ''), descriptor = descriptors.get(targetPath);
      if (descriptor && /\.xlsx$/i.test(targetPath)) await drafts.ready(descriptor.sessionId);
      if (descriptor && /\.docx$/i.test(targetPath) && drafts.canEdit(descriptor.sessionId)) await drafts.ready(descriptor.sessionId);
    },
    filesChanged: () => { for (const listener of fileListeners) listener(null); },
    flush: async () => { await text.flush(); await drafts.flush(); },
    address: path => {
      const targetPath = (path ?? activePath()).replace(/^preview:/, '');
      const descriptor = descriptors.get(targetPath);
      return { workbenchId, workspaceId, ...(descriptor && { documentId: descriptor.sessionId,
        ...(text.has(descriptor.sessionId) ? text.context(descriptor.sessionId) : drafts.context(descriptor.sessionId)) }) };
    },
    officeLease(address, action) {
      if (!address.documentId || descriptors.get(activePath())?.sessionId !== address.documentId) throw new Error('请先激活提案对应的工作簿。');
      return drafts.withCleanLease(address.documentId, action);
    },
    async officeUndo(address, commitId, requestId) {
      if (!address.documentId || !address.generation || descriptors.get(activePath())?.sessionId !== address.documentId) throw new Error('请先激活对应工作簿。');
      let id = officeUndoRequests.get(requestId); if (!id) { id = crypto.randomUUID(); officeUndoRequests.set(requestId, id); }
      await drafts.withCleanLease(address.documentId, async revision => {
        const response = await fetch(`/api/easyview/v1/undo?${new URLSearchParams({ documentId: address.documentId!, workbenchId, generation: address.generation! })}`, {
          method: 'POST', credentials: 'same-origin', body: JSON.stringify({ operationId: commitId, requestId: id, editorRevision: revision }) });
        const result = await response.json(); if (!response.ok || !result.ok) throw new Error(result.message || '文件撤销未确认。');
      });
    },
    preparePatch(address, message) {
      if (!address.documentId || activePath() !== message.path) throw new Error('目标 Markdown 编辑器未激活，已暂停 AI 修改。');
      text.beginAiOperation(address.documentId, message);
    },
  });
  let disposed = false;
  let closeDialog: HTMLDialogElement | null = null;
  const tabDocumentId = (tab: DesktopTabSnapshot['tabs'][number]) => tab.kind === 'editor' ? tab.documentId : tab.kind === 'preview' ? descriptors.get(tab.relativePath)?.sessionId : undefined;
  const history = createFileHistoryOwner({
    context() {
      const path = activePath(), descriptor = descriptors.get(path);
      if (!descriptor) return null;
      const context = text.has(descriptor.sessionId) ? text.context(descriptor.sessionId) : drafts.context(descriptor.sessionId);
      return context.generation ? { documentId: descriptor.sessionId, workbenchId, generation: context.generation, path } : null;
    },
    withCleanLease: (id, action) => text.has(id) ? text.withCleanLease(id, action) : drafts.withCleanLease(id, action),
    async retire(documentId) {
      const path = [...descriptors.values()].find(value => value.sessionId === documentId)?.relativePath;
      await forgetTabs(tab => tabDocumentId(tab) !== documentId);
      if (path) descriptors.delete(path);
      tabs.activeTabId = tabs.tabs.at(-1)?.id ?? null; publishTabs();
      for (const listener of fileListeners) listener(null);
    },
    async reopen(path) { const result = await api.tabs.openWorkspaceEntry(path); if (!result.ok) throw new Error(result.message); },
  });
  async function requestClose(retain: (tab: DesktopTabSnapshot['tabs'][number]) => boolean): Promise<boolean> {
    if (disposed || closeDialog || !await flush()) return false;
    const dirty: Array<{ tab: DesktopTabSnapshot['tabs'][number]; documentId: string; generation: string; checkpointRevision: number }> = [];
    for (const tab of tabs.tabs.filter(tab => !retain(tab))) {
      const documentId = tabDocumentId(tab); if (!documentId) continue;
      const result = await request<{ dirty: boolean; generation: string; checkpointRevision: number }>('draft-state', { documentId, workbenchId });
      if (!result.ok) { drafts.show('failed', result.message); return false; }
      if (result.value.dirty) dirty.push({ tab, documentId, generation: result.value.generation, checkpointRevision: result.value.checkpointRevision });
    }
    if (!dirty.length) return true;
    const dialog = document.createElement('dialog'); closeDialog = dialog;
    const title = document.createElement('h3'); title.textContent = `${dirty.length} 个文档有未保存到原件的修改`; dialog.append(title);
    const message = document.createElement('p'); message.textContent = '选择如何处理这些文档。'; dialog.append(message);
    const action = await new Promise<'save' | 'keep' | 'discard' | null>(resolve => {
      dialog.addEventListener('cancel', () => resolve(null), { once: true });
      dialog.addEventListener('close', () => resolve(null), { once: true });
      for (const [value, label] of [['save', '保存原件并关闭'], ['keep', '保留恢复草稿并关闭'], ['discard', '丢弃修改并关闭'], ['cancel', '取消']] as const) {
        const button = document.createElement('button'); button.textContent = label;
        button.onclick = () => resolve(value === 'cancel' ? null : value); dialog.append(button);
      }
      document.body.append(dialog); dialog.showModal();
    });
    try {
      if (!action || disposed) return false;
      if (action === 'save') for (const item of dirty) {
        if (disposed) return false;
        if (item.tab.kind === 'editor') { tabs.activeTabId = item.tab.id; text.activate(item.documentId); publishTabs(); await text.save(item.documentId); }
        else { tabs.activeTabId = item.tab.id; publishTabs(); await drafts.saveCurrent(item.documentId); }
      }
      if (action === 'discard') for (const item of dirty) {
        const response = await fetch(`/api/easyview/v1/discard?${new URLSearchParams({ documentId: item.documentId, workbenchId, generation: item.generation })}`, {
          method: 'POST', credentials: 'same-origin', body: JSON.stringify({ requestId: crypto.randomUUID(), expectedCheckpointRevision: item.checkpointRevision }) });
        const body = await response.json(); if (!response.ok || !body.ok) throw new Error(body.message || '丢弃草稿失败。');
      }
      return true;
    } catch (error) { drafts.show('failed', `关闭已取消：${String(error)}`); return false; }
    finally { dialog.close(); dialog.remove(); closeDialog = null; }
  }
  async function forgetTabs(retain: (tab: DesktopTabSnapshot['tabs'][number]) => boolean) {
    const ids: string[] = [];
    for (const tab of tabs.tabs) {
      if (retain(tab)) continue;
      if (tab.kind === 'editor' && tab.documentId) ids.push(tab.documentId);
      if (tab.kind === 'preview') { const descriptor = descriptors.get(tab.relativePath); if (descriptor) ids.push(descriptor.sessionId); descriptors.delete(tab.relativePath); }
    }
    text.forget(ids);
    if (ids.length) await fetch('/api/easyview/v1/close', { method: 'POST', credentials: 'same-origin', body: JSON.stringify(ids) });
    tabs.tabs = tabs.tabs.filter(retain);
  }
  const subscription = <T>(listeners: Set<(value: T) => void>, listener: (value: T) => void) => {
    listeners.add(listener); return { unsubscribe() { listeners.delete(listener); } };
  };
  async function request<T>(endpoint: string, params: Record<string, string>): Promise<OperationResult<T>> {
    const url = new URL(`/api/easyview/v1/${endpoint}`, location.origin);
    for (const [key, value] of Object.entries(params)) url.searchParams.set(key, value);
    try {
      const response = await fetch(url, { credentials: 'same-origin' });
      const body: OperationResult<T> = await response.json();
      if (typeof body.ok !== 'boolean') throw new Error('Invalid EasyView response');
      return body;
    } catch { return operationFailure('DEPENDENCY_MISSING', '文档宿主连接失败，请重新进入工作台。'); }
  }
  function publishTabs() {
    tabs.groups = [{ id: 'group-1', tabs: [...tabs.tabs], activeTabId: tabs.activeTabId }];
    for (const listener of tabListeners) listener(structuredClone(tabs));
    return operationSuccess(structuredClone(tabs));
  }
  async function selectWorkspace(id: string) {
    const workspace = bootstrap.workspaces.find(value => value.id === id);
    if (!workspace) return operationFailure('PERMISSION_DENIED', '请先在 DSH 添加工作区，再进入文档工作台。');
    if (workspaceId === id) return operationSuccess(structuredClone(state));
    if (!await requestClose(() => false)) return operationFailure('CONFLICT', '已取消切换工作区。');
    await forgetTabs(() => false);
    workspaceId = id; tabs.tabs = []; tabs.activeTabId = null; descriptors.clear(); publishTabs();
    state = { ...state, rootPath: `dsh-workspace/${workspace.title}`, openTabs: [], activeTabId: null };
    for (const listener of rootListeners) listener(structuredClone(state));
    return operationSuccess(structuredClone(state));
  }
  async function pickWorkspace() {
    // Refresh the public registry: DSH may have added a workspace since this frame opened.
    try {
      const response = await fetch('/api/easyview/v1/bootstrap', { credentials: 'same-origin' });
      if (!response.ok) throw new Error('Workspace list unavailable');
      const latest: EasyViewBootstrap = await response.json();
      if (latest.apiVersion !== 1 || !Array.isArray(latest.workspaces)) throw new Error('Invalid workspace list');
      bootstrap.workspaces = latest.workspaces;
    } catch { return operationFailure('DEPENDENCY_MISSING', '工作区列表加载失败，请重试。'); }
    const dialog = document.createElement('dialog');
    dialog.className = 'easyview-workspace-picker';
    dialog.setAttribute('aria-label', '切换工作区');
    const heading = document.createElement('h2'); heading.textContent = '切换工作区'; dialog.appendChild(heading);
    const hint = document.createElement('p');
    hint.textContent = '选择 DSH 已添加的工作区。添加或管理工作区，请先返回 DSH。'; dialog.append(hint);
    const list = document.createElement('div'); list.className = 'easyview-workspace-list'; dialog.append(list);
    if (!bootstrap.workspaces.length) {
      const empty = document.createElement('p'); empty.textContent = '暂无工作区，请先在 DSH 中添加。'; list.append(empty);
    }
    const selected = new Promise<string | null>(resolve => {
      dialog.addEventListener('cancel', () => resolve(null), { once: true });
      dialog.addEventListener('close', () => resolve(null), { once: true });
      for (const workspace of bootstrap.workspaces) {
        const button = document.createElement('button'); button.type = 'button'; button.textContent = workspace.title;
        button.setAttribute('aria-label', workspace.title);
        if (workspace.id === workspaceId) {
          button.setAttribute('aria-current', 'true');
          const badge = document.createElement('span'); badge.textContent = '当前'; button.append(badge);
        }
        button.onclick = () => { resolve(workspace.id); dialog.close(); }; list.appendChild(button);
      }
      const cancel = document.createElement('button'); cancel.type = 'button'; cancel.textContent = '取消';
      cancel.className = 'easyview-workspace-cancel'; cancel.onclick = () => { resolve(null); dialog.close(); }; dialog.appendChild(cancel);
    });
    document.body.appendChild(dialog); dialog.showModal();
    const id = await selected; dialog.remove(); return id ? selectWorkspace(id) : operationSuccess(null);
  }
  api.capabilities = { newWindow: false, previewEditing: bootstrap.capabilities.edit, previewSaving: bootstrap.capabilities.save, editorHost: { sourceMode: 'embedded', git: false, terminal: false, aiCommitMessage: false,
    aiChat: true, aiChatHistory: 'host', documentConversion: false, shortcutPersistence: false } };
  api.workspace.getState = async () => operationSuccess(structuredClone(state));
  api.workspace.setState = async next => { state = { ...next, rootPath: state.rootPath }; return operationSuccess(structuredClone(state)); };
  api.workspace.onRootChanged = listener => subscription(rootListeners, listener);
  api.workspace.openFolder = async () => {
    if (!options.openFolder) return pickWorkspace();
    try {
      const id = await options.openFolder();
      if (!id) return operationSuccess(null);
      const response = await fetch('/api/easyview/v1/bootstrap', { credentials: 'same-origin' });
      if (!response.ok) throw new Error('工作区列表加载失败。');
      const latest: EasyViewBootstrap = await response.json();
      if (latest.apiVersion !== 1 || !Array.isArray(latest.workspaces)) throw new Error('工作区列表无效。');
      bootstrap.workspaces = latest.workspaces;
      return selectWorkspace(id);
    } catch { return operationFailure('DEPENDENCY_MISSING', '打开文件夹失败，请重试或检查 DSH 目录权限。'); }
  };
  api.workspace.openRecentFolder = selectWorkspace;
  api.workspace.getRecentFolders = async () => operationSuccess(bootstrap.workspaces.map(workspace => ({ path: workspace.id, name: workspace.title, displayPath: workspace.title, exists: true })));
  api.workspace.readDirectory = path => request('tree', { workspaceId, path });
  api.workspace.onChanged = listener => subscription(fileListeners, listener);
  api.workspace.onContextCommand = listener => { contextListeners.add(listener); return { unsubscribe: () => { contextListeners.delete(listener); } }; };
  api.workspace.popupNewFileMenu = async parent => {
    for (const listener of contextListeners) listener('beginCreateFile', parent, '.md');
    return operationSuccess(true);
  };
  const createRequests = new Map<string, string>();
  api.workspace.create = async input => {
    if (input.kind !== 'file' || !isMarkdownFilePath(input.name)) return operationFailure('INVALID_ARGUMENT', '当前仅支持新建 Markdown 文件。');
    const path = input.parentRelativePath ? `${input.parentRelativePath}/${input.name}` : input.name;
    const key = `${workspaceId}/${path}`, requestId = createRequests.get(key) ?? crypto.randomUUID(); createRequests.set(key, requestId);
    try {
      const response = await fetch('/api/easyview/v1/create-markdown', { method: 'POST', credentials: 'same-origin',
        body: JSON.stringify({ workspaceId, path, content: '', requestId }) });
      const result = await response.json();
      if (!response.ok || !result.ok) return operationFailure(result.code || 'CONFLICT', result.message || '新建文件失败。');
      createRequests.delete(key);
      return operationSuccess<WorkspaceEntry>({ id: result.value.operationId, rootId: workspaceId, name: input.name, relativePath: path, kind: 'file' });
    } catch { return operationFailure('CONFLICT', '新建结果尚未确认，请核对文件列表后重试。'); }
  };
  api.tabs.getState = async () => publishTabs();
  api.tabs.onChanged = listener => subscription(tabListeners, listener);
  api.tabs.openWorkspaceEntry = async path => {
    if (!await flush()) return operationFailure('CONFLICT', '草稿尚未确认。');
    const prior = tabs.tabs.find(tab => tab.kind === 'preview' ? tab.relativePath === path : tab.kind === 'editor' && tab.filePath === `${state.rootPath}/${path}`);
    if (prior) { if (prior.kind === 'editor' && prior.documentId) text.activate(prior.documentId); tabs.activeTabId = prior.id; return publishTabs(); }
    const opened = await request<PreviewDescriptor>('open', { workspaceId, path });
    if (!opened.ok) return opened;
    descriptors.set(path, opened.value);
    const existing = tabs.tabs.find(tab => tab.kind === 'preview' && tab.relativePath === path);
    if (bootstrap.capabilities.edit && (isMarkdownFilePath(path) || /\.txt$/i.test(path))) {
      try {
        const prior = tabs.tabs.find(tab => tab.kind === 'editor' && tab.filePath === `${state.rootPath}/${path}`);
        if (prior?.kind === 'editor' && prior.documentId) { text.activate(prior.documentId); tabs.activeTabId = prior.id; return publishTabs(); }
        if (await text.open(opened.value)) {
          const id = crypto.randomUUID();
          tabs.tabs.push({ id, kind: 'editor', filePath: `${state.rootPath}/${path}`, fileName: opened.value.fileName,
            documentId: opened.value.sessionId, documentKey: opened.value.sessionId, dirty: false });
          tabs.activeTabId = id; return publishTabs();
        }
      } catch (error) { return operationFailure('CONFLICT', String(error)); }
    }
    const id = existing?.id ?? crypto.randomUUID();
    if (!existing) tabs.tabs.push({ id, kind: 'preview', relativePath: path, fileName: opened.value.fileName, route: opened.value.route });
    tabs.activeTabId = id; return publishTabs();
  };
  api.tabs.openPreview = api.tabs.openWorkspaceEntry;
  api.tabs.activate = async id => { const tab = tabs.tabs.find(tab => tab.id === id); if (!tab) return operationFailure('NOT_FOUND', '标签不存在。'); if (!await flush()) return operationFailure('CONFLICT', '草稿尚未确认。'); if (tab.kind === 'editor' && tab.documentId) text.activate(tab.documentId); tabs.activeTabId = id; return publishTabs(); };
  api.tabs.close = async id => { if (!await requestClose(tab => tab.id !== id)) return operationFailure('CONFLICT', '已取消关闭。'); await forgetTabs(tab => tab.id !== id); if (tabs.activeTabId === id) tabs.activeTabId = tabs.tabs.at(-1)?.id ?? null; return publishTabs(); };
  api.tabs.closeAll = async () => { if (!await requestClose(() => false)) return operationFailure('CONFLICT', '已取消关闭。'); await forgetTabs(() => false); tabs.activeTabId = null; return publishTabs(); };
  api.tabs.closeOthers = async id => { if (!await requestClose(tab => tab.id === id)) return operationFailure('CONFLICT', '已取消关闭。'); await forgetTabs(tab => tab.id === id); tabs.activeTabId = tabs.tabs[0]?.id ?? null; return publishTabs(); };
  api.preview.open = async path => {
    const opened = descriptors.has(path) ? operationSuccess(descriptors.get(path)!) : await request<PreviewDescriptor>('open', { workspaceId, path });
    if (!opened.ok) return opened;
    try { await drafts.prepare(opened.value.sessionId, opened.value.fileName); }
    catch (error) { return operationFailure('CONFLICT', String(error)); }
    return opened;
  };
  api.preview.close = async () => operationSuccess(true);
  api.preview.readText = async id => {
    const descriptor = [...descriptors.values()].find(value => value.sessionId === id);
    if (!descriptor) return operationFailure('NOT_FOUND', '文档会话不存在。');
    const response = await fetch(descriptor.contentUrl, { credentials: 'same-origin' });
    if (!response.ok) return operationFailure('CONFLICT', '原件发生变化，请重新打开。');
    return operationSuccess({ content: await response.text(), truncated: false });
  };
  const requested = new URLSearchParams(location.search).get('workspaceId');
  if (requested) await selectWorkspace(requested);
  return Object.assign(api, { workbench: {
    pickWorkspace,
    async flush() { await text.flush(); await drafts.flush(); },
    async openDocument(path: string) { const result = await api.tabs.openWorkspaceEntry(path); if (!result.ok) throw new Error(result.message); },
    requestClose: () => requestClose(() => false),
    dispose() {
      if (disposed) return; disposed = true; closeDialog?.close();
      const ids = tabs.tabs.map(tabDocumentId).filter((id): id is string => Boolean(id));
      if (ids.length) void fetch('/api/easyview/v1/close', { method: 'POST', credentials: 'same-origin', keepalive: true, body: JSON.stringify(ids) }).catch(() => {});
      history.dispose(); ai.dispose(); text.dispose(); drafts.dispose(); editorListeners.clear(); rootListeners.clear(); tabListeners.clear(); fileListeners.clear(); contextListeners.clear();
    },
  } });
}
