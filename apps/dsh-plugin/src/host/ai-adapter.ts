/** DSH owns model execution/history; SQLite owns only the workbench's visible projection. */
import { DatabaseSync } from 'node:sqlite';
import { resolve } from 'node:path';
import { randomUUID } from 'node:crypto';
import { applyAiTextPatches, type HostToEditorMessage } from '@easyview/contracts';
import { createCanonicalDocument } from '@easyview/editor-sync';
import { contentHash, DocumentError, type DocumentProvider } from './documents';
import type { DraftStore } from './draft-store';
import type { FileCommands } from './file-commands';
import type { XlsxProposals } from './xlsx-proposals';
import { AI_PRESETS, AI_TOOLS, type AiAddress, type AiExecution, type AiTools, type DshSessions, type VisibleAiMessage } from '../shared/ai';
import { readDocxContext } from './docx-context';

type Mode = 'chat' | 'agent';
interface SessionBinding { id: string; cursor: number }
interface RequestRecord {
  id: string; hash: string; sessionId: string; mode: Mode; user: string; assistant: string;
  state: 'running' | 'done' | 'failed' | 'interrupted'; error?: string;
}
interface Conversation {
  schema: 1; key: string; sessions: Partial<Record<Mode, SessionBinding>>;
  messages: VisibleAiMessage[]; requests: RequestRecord[];
}
interface Running {
  conversation: Conversation; request: RequestRecord; address: AiAddress; sessionId: string;
  events: HostToEditorMessage[]; text: string; abort: AbortController; toolsAbort: AbortController; stopped: boolean;
  completion?: Promise<void>; patch?: { operationId: string; resolve(error?: string): void };
}
export interface AiDependencies {
  sessions(): Partial<DshSessions> | undefined;
  tools(): AiTools | undefined;
  agent(id: string): object | undefined;
  modelSchemas?(agent: object): Promise<Array<{ name: string }>>;
  workspaceExists(id: string): boolean;
}

export class DshAiAdapter {
  private readonly database: DatabaseSync;
  private readonly running = new Map<string, Running>();
  private readonly admissions = new Set<string>();
  private readonly admissionTasks = new Set<Promise<unknown>>();
  private readonly conversations = new Map<string, Conversation>();
  private closed = false;
  constructor(directory: string, private readonly drafts: DraftStore, private readonly documents: DocumentProvider, private readonly dependencies: AiDependencies, private readonly files?: FileCommands, private readonly xlsx?: XlsxProposals) {
    this.database = new DatabaseSync(resolve(directory, 'ai.sqlite'));
    this.database.exec(`PRAGMA journal_mode=WAL; PRAGMA synchronous=FULL;
      CREATE TABLE IF NOT EXISTS conversations (key TEXT PRIMARY KEY, record TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS cleared_conversations (requestId TEXT PRIMARY KEY, key TEXT NOT NULL, record TEXT NOT NULL);`);
    for (const row of this.database.prepare('SELECT record FROM conversations').all()) {
      const conversation: Conversation = JSON.parse(String(row.record));
      if (conversation.schema !== 1) throw new Error('Unsupported EasyView AI projection schema');
      for (const request of conversation.requests) if (request.state === 'running') {
        request.state = 'interrupted'; request.error = '执行已中断；请核对 DSH 原会话和恢复草稿后再继续。';
        const message = conversation.messages.find(item => item.id === request.assistant);
        if (message) { message.error = request.error; message.badge = 'error'; }
      }
      this.write(conversation);
    }
  }
  get available(): boolean {
    const sessions = this.dependencies.sessions();
    return ['modelCatalog', 'create', 'selectSessionModel', 'prompt', 'follow', 'cancel'].every(key => typeof Reflect.get(sessions ?? {}, key) === 'function')
      && Boolean(this.dependencies.tools());
  }
  async models() {
    const catalog = await this.controller().modelCatalog();
    return catalog.groups.flatMap(group => group.models.map(model => ({ label: `${group.id} / ${model.id}`, provider: group.id, model: model.id })));
  }
  async state(address: AiAddress) {
    const conversation = await this.conversation(address);
    const active = [...this.running.values()].find(job => job.conversation.key === conversation.key);
    return { messages: conversation.messages.map(message => ({ ...message, ...(message.role === 'assistant' && this.xlsx && {
      officeEdits: this.xlsx.projection(conversation.key, message.requestId) }) })), activeRequestId: active?.request.id,
      streamingMessageId: active?.request.assistant, content: active?.text ?? '',
      events: active?.events.filter(event => event.type !== 'aiChat.applyPatches' || this.drafts.aiPatch(event.operationId)?.state === 'pending') ?? [] };
  }
  async decideOffice(address: AiAddress, input: { proposalId: string; decisionId: string; accepted: boolean; editorRevision: number }) {
    if (!this.xlsx) throw new DocumentError('DEPENDENCY_MISSING', 'XLSX 提案尚未启用。');
    const conversation = await this.conversation(address);
    return this.xlsx.decide(conversation.key, input);
  }
  async clear(address: AiAddress, requestId: string) {
    const conversation = await this.conversation(address);
    if (typeof requestId !== 'string' || !/^[A-Za-z0-9:_-]{1,128}$/.test(requestId)) throw new DocumentError('INVALID_ARGUMENT', '无效的清空请求身份。');
    const receipt = this.database.prepare('SELECT key FROM cleared_conversations WHERE requestId=?').get(requestId);
    if (receipt) { if (receipt.key !== conversation.key) throw new DocumentError('CONFLICT', '清空请求身份冲突。'); return; }
    if (this.admissions.has(conversation.key) || [...this.running.values()].some(job => job.conversation.key === conversation.key)
      || conversation.requests.some(request => request.state === 'interrupted')) throw new DocumentError('CONFLICT', '请先停止并核实当前执行，再清空对话。');
    this.database.exec('BEGIN IMMEDIATE');
    try {
      this.database.prepare('INSERT INTO cleared_conversations VALUES (?, ?, ?)').run(requestId, conversation.key, JSON.stringify(conversation));
      this.write({ ...conversation, messages: [], requests: [], sessions: {} });
      this.database.exec('COMMIT');
      conversation.messages = []; conversation.requests = []; conversation.sessions = {};
    } catch (error) { this.database.exec('ROLLBACK'); throw error; }
  }
  async send(address: AiAddress, input: { requestId: string; mode: Mode; model: string; userMessage: string }) {
    const task = this.admit(address, input); this.admissionTasks.add(task);
    try { return await task; } finally { this.admissionTasks.delete(task); }
  }
  private async admit(address: AiAddress, input: { requestId: string; mode: Mode; model: string; userMessage: string }) {
    if (this.closed || this.running.size >= 4) throw new DocumentError('CONFLICT', 'AI 执行队列不可用，请稍后重试。');
    if (!input || !['chat', 'agent'].includes(input.mode) || typeof input.requestId !== 'string' || !/^[A-Za-z0-9:_-]{1,128}$/.test(input.requestId)
      || typeof input.userMessage !== 'string' || !input.userMessage.trim() || input.userMessage.length > 64_000) throw new DocumentError('INVALID_ARGUMENT', '无效的 AI 请求。');
    const conversation = await this.conversation(address);
    const hash = contentHash(Buffer.from(JSON.stringify({ ...input, address })));
    const previous = conversation.requests.find(request => request.id === input.requestId);
    if (previous) {
      if (previous.hash !== hash) throw new DocumentError('CONFLICT', '同一 AI 请求身份对应不同内容。');
      return { requestId: previous.id, state: previous.state };
    }
    if ([...this.running.values()].some(job => job.conversation.key === conversation.key)) throw new DocumentError('CONFLICT', '当前文档已有 AI 请求执行中。');
    // An interrupted request may still own a DSH turn. Do not silently submit a
    // second command or treat a disconnected prompt as safe to replay.
    if (conversation.requests.some(request => request.state === 'interrupted')) throw new DocumentError('CONFLICT', '存在中断的 AI 请求，请先在 DSH 原会话核对执行结果。');
    if (this.admissions.has(conversation.key)) throw new DocumentError('CONFLICT', '当前文档正在提交另一个 AI 请求。');
    this.admissions.add(conversation.key);
    try {
    if (conversation.messages.length >= 198 || conversation.messages.reduce((sum, message) => sum + message.content.length, 0) > 900_000) throw new DocumentError('CONFLICT', '本次文档会话已达到容量上限。');
    const route = (await this.models()).find(model => model.label === input.model);
    if (!route) throw new DocumentError('DEPENDENCY_MISSING', 'DSH 当前没有可用的所选模型，请在 DSH 配置模型后重试。');
    if (input.mode === 'agent' && (!address.documentId || !address.generation)) throw new DocumentError('DEPENDENCY_MISSING', '请先打开 Markdown 或 XLSX 文档。');
    if (address.documentId) {
      const identity = await this.documents.draftIdentity(address.documentId);
      if (input.mode === 'agent' && identity.format !== 'markdown' && identity.format !== 'xlsx') throw new DocumentError('DEPENDENCY_MISSING', '此格式的 AI 编辑尚未接通。');
      if (input.mode === 'agent' && identity.format === 'xlsx') {
        if (!this.xlsx) throw new DocumentError('DEPENDENCY_MISSING', 'XLSX 提案尚未启用。');
        this.xlsx.assertClean(address);
      }
    }
    const controller = this.controller();
    let session = conversation.sessions[input.mode];
    if (!session) {
      const created = await controller.create({ workspaceId: address.workspaceId, agentPreset: AI_PRESETS[input.mode] });
      if (created.agentPreset !== AI_PRESETS[input.mode]) throw new DocumentError('DEPENDENCY_MISSING', 'DSH 未激活指定 EasyView preset。');
      session = { id: created.sessionId, cursor: 0 }; conversation.sessions[input.mode] = session; this.write(conversation);
    }
    await controller.selectSessionModel({ sessionId: session.id, provider: route.provider, model: route.model });
    await this.verifyTools(session.id, input.mode);
    if (this.closed) throw new DocumentError('CONFLICT', '插件正在卸载，AI 请求尚未发送。');
    const request: RequestRecord = { id: input.requestId, hash, sessionId: session.id, mode: input.mode,
      user: randomUUID(), assistant: randomUUID(), state: 'running' };
    const delta = conversation.messages.slice(session.cursor).filter(message => !message.error);
    const context = delta.length ? `Visible workbench history after cursor ${session.cursor}; source IDs must not be treated as new requests:\n${JSON.stringify(delta.map(({ id, role, content }) => ({ id, role, content })))}` : '';
    conversation.messages.push({ id: request.user, role: 'user', content: input.userMessage, requestId: request.id },
      { id: request.assistant, role: 'assistant', content: '', requestId: request.id });
    conversation.requests.push(request); this.write(conversation);
    const job: Running = { conversation, request, address: { ...address }, sessionId: session.id, events: [], text: '', abort: new AbortController(), toolsAbort: new AbortController(), stopped: false };
    this.running.set(session.id, job);
    job.completion = this.execute(job, [context, 'You are the EasyView document assistant. Use only the authorized document tools. Chat is read-only. Markdown patches change recovery drafts; saving the original requires a separate user save action. Workbook edits produce a cell-diff proposal and write the original only when the user explicitly applies that card. Read the active document first to obtain baseRevision or workbook fingerprint.', input.userMessage].filter(Boolean).join('\n\n'));
    void job.completion.catch(() => {});
    return { requestId: request.id, state: request.state };
    } finally { this.admissions.delete(conversation.key); }
  }
  async cancel(address: AiAddress, requestId: string) {
    const conversation = await this.conversation(address);
    const job = [...this.running.values()].find(item => item.conversation.key === conversation.key && item.request.id === requestId);
    if (!job) return;
    job.stopped = true; job.toolsAbort.abort(); job.patch?.resolve('本次请求已停止。');
    await this.controller().cancel({ sessionId: job.sessionId });
    // Cancel admission is not a claim that a tool or turn has already stopped.
  }
  async acknowledge(address: AiAddress, requestId: string, operationId: string, error?: string) {
    const conversation = await this.conversation(address);
    const job = [...this.running.values()].find(item => item.conversation.key === conversation.key && item.request.id === requestId);
    if (job?.patch?.operationId === operationId) job.patch.resolve(error);
  }
  async tool(name: string, args: unknown, execution: AiExecution): Promise<string> {
    const job = execution.agent && this.running.get(execution.agent.id);
    if (!job || job.stopped || !AI_TOOLS[job.request.mode].includes(name)) throw new Error('No active authorized EasyView request');
    execution.signal.throwIfAborted(); await this.verifyTools(job.sessionId, job.request.mode);
    const operationId = contentHash(Buffer.from(`${job.sessionId}\0${job.request.id}\0${execution.callId}`));
    const inputHash = contentHash(Buffer.from(JSON.stringify(args)));
    if (name === 'easyview_edit_workbook') {
      if (!this.xlsx) throw new Error('Workbook proposals are unavailable');
      return JSON.stringify(await this.xlsx.propose({ address: job.address, conversationKey: job.conversation.key,
        requestId: job.request.id, toolId: operationId, args, signal: AbortSignal.any([execution.signal, job.toolsAbort.signal]) }));
    }
    if (name === 'easyview_create_markdown') {
      if (!this.files || !args || typeof args !== 'object') throw new Error('File creation is unavailable');
      const input = { workspaceId: job.address.workspaceId, path: Reflect.get(args, 'path'), content: Reflect.get(args, 'content'), requestId: operationId };
      execution.signal.throwIfAborted();
      if (job.stopped || this.closed) throw new Error('This AI request has stopped');
      const result = await this.files.createMarkdown(input, AbortSignal.any([execution.signal, job.toolsAbort.signal]));
      if (result.state !== 'committed') throw new Error(`File creation requires reconciliation: ${result.state}`);
      return JSON.stringify(result);
    }
    const previous = name === 'easyview_edit_markdown' ? this.drafts.aiPatch(operationId) : null;
    if (previous) {
      if (previous.inputHash !== inputHash) throw new Error('AI operation payload conflict');
      if (previous.state !== 'applied') throw new Error('This patch operation cannot be replayed');
      return JSON.stringify({ operationId, state: 'applied', checkpointRevision: previous.checkpointRevision });
    }
    const documentId = job.address.documentId;
    if (!documentId) throw new Error('Open an authorized document first');
    const identity = await this.documents.draftIdentity(documentId, execution.signal);
    if (name === 'easyview_read_document' && identity.format === 'xlsx' && this.xlsx) return JSON.stringify(await this.xlsx.inspect(job.address, execution.signal));
    if (name === 'easyview_read_document' && identity.format === 'docx') {
      const key = contentHash(Buffer.from(`${job.address.workbenchId}\0${identity.fileIdentity}`));
      if (job.address.generation) this.drafts.assertGeneration(key, job.address.generation);
      const original = await this.documents.read(documentId, execution.signal), record = this.drafts.load(key);
      if (record && record.sourceDiskHash !== identity.diskHash) throw new Error('Original/draft baseline conflict');
      const context = await readDocxContext(record ? await this.drafts.bytes(record) : original, execution.signal);
      return JSON.stringify({ path: this.documents.reference(documentId).relativePath, format: 'docx', scope: 'document',
        source: record ? 'recovery_checkpoint' : 'saved_document', baseRevision: record?.editorRevision ?? 0, ...context });
    }
    if (!job.address.generation) throw new Error('The text editor has not finished mounting');
    if (!['markdown', 'text'].includes(identity.format ?? '')) throw new Error('This format’s AI context adapter is not available');
    const key = contentHash(Buffer.from(`${job.address.workbenchId}\0${identity.fileIdentity}`));
    this.drafts.assertGeneration(key, job.address.generation);
    // Re-authorize and check the original even when a recovery draft exists.
    const original = await this.documents.read(documentId, execution.signal), record = this.drafts.load(key);
    if (record && record.sourceDiskHash !== identity.diskHash) throw new Error('Original/draft baseline conflict');
    const content = createCanonicalDocument(new TextDecoder('utf-8', { fatal: true }).decode(record ? await this.drafts.bytes(record) : original)).content;
    const baseRevision = record?.editorRevision ?? 0;
    if (name === 'easyview_read_document') return JSON.stringify({ path: this.documents.reference(documentId).relativePath, baseRevision, content });
    if (identity.format !== 'markdown' || !args || typeof args !== 'object') throw new Error('Only Markdown supports draft patches');
    const patch = { description: Reflect.get(args, 'description'), expectedText: Reflect.get(args, 'expectedText'), replacement: Reflect.get(args, 'replacement') };
    if (Reflect.get(args, 'baseRevision') !== baseRevision || !Object.values(patch).every(value => typeof value === 'string')
      || patch.expectedText.length + patch.replacement.length > 512_000) throw new Error('Patch arguments or baseRevision are invalid');
    const next = applyAiTextPatches(content, [patch]);
    execution.signal.throwIfAborted();
    if (job.stopped || this.closed) throw new Error('This AI request has stopped');
    if (next === content) return JSON.stringify({ operationId, state: 'no_change', baseRevision });
    this.drafts.prepareAiPatch({ operationId, key, generation: job.address.generation, baseRevision,
      outputHash: contentHash(new TextEncoder().encode(next)), inputHash, state: 'pending' });
    let complete!: (error?: string) => void;
    const decision = new Promise<string | undefined>(resolve => { complete = resolve; });
    job.patch = { operationId, resolve: complete };
    const aborted = () => complete('AI 修改已中断。');
    execution.signal.addEventListener('abort', aborted, { once: true });
    const timer = setTimeout(() => complete('编辑器尚未确认 AI 修改，请核对恢复草稿。'), 30_000);
    job.events.push({ type: 'aiChat.applyPatches', requestId: job.request.id, operationId,
      path: this.documents.reference(documentId).relativePath, patches: [patch] });
    try {
      const error = await decision;
      const receipt = this.drafts.aiPatch(operationId);
      // The durable checkpoint receipt wins over a lost browser ACK/cancel race.
      if (receipt?.state === 'applied') return JSON.stringify({ operationId, state: 'applied', checkpointRevision: receipt.checkpointRevision });
      throw new Error(error ?? 'The editor did not confirm a durable patch receipt');
    } finally {
      clearTimeout(timer); execution.signal.removeEventListener('abort', aborted); this.drafts.expireAiPatch(operationId); job.patch = undefined;
    }
  }
  async close() {
    if (this.closed) return; this.closed = true;
    await Promise.allSettled(this.admissionTasks);
    for (const job of this.running.values()) { job.stopped = true; job.toolsAbort.abort(); job.patch?.resolve('插件已卸载。'); job.abort.abort(); }
    await Promise.allSettled([...this.running.values()].map(job => this.controller().cancel({ sessionId: job.sessionId })));
    await Promise.allSettled([...this.running.values()].map(job => job.completion));
    this.database.close();
  }
  private controller(): DshSessions {
    if (!this.available) throw new DocumentError('DEPENDENCY_MISSING', '目标 DSH 缺少会话级模型或工具接口。');
    return this.dependencies.sessions() as DshSessions;
  }
  private async verifyTools(sessionId: string, mode: Mode) {
    const agent = this.dependencies.agent(sessionId), tools = this.dependencies.tools();
    if (!agent || !tools) throw new DocumentError('DEPENDENCY_MISSING', '无法验证 DSH Agent 工具 scope。');
    const schemas = this.dependencies.modelSchemas ? await this.dependencies.modelSchemas(agent) : tools.schemas(agent);
    const names = schemas.map(tool => tool.name).sort();
    if (JSON.stringify(names) !== JSON.stringify([...AI_TOOLS[mode]].sort())) throw new DocumentError('PERMISSION_DENIED', `EasyView preset 工具集合与预期不一致，已阻止发送。实际工具：${names.join(', ') || '(empty)'}`);
  }
  private async conversation(address: AiAddress): Promise<Conversation> {
    if (this.closed || !/^[a-f0-9-]{36}$/.test(address.workbenchId) || !this.dependencies.workspaceExists(address.workspaceId)) throw new DocumentError('PERMISSION_DENIED', 'AI 工作区或工作台身份已失效。');
    const identity = address.documentId ? await this.documents.reference(address.documentId) : null;
    if (identity && identity.workspaceId !== address.workspaceId) throw new DocumentError('PERMISSION_DENIED', 'AI 文档不属于当前工作区。');
    if (address.documentId) await this.documents.read(address.documentId);
    const key = contentHash(Buffer.from(`${address.workbenchId}\0${address.workspaceId}\0${identity?.fileIdentity ?? 'standalone'}`));
    const active = [...this.running.values()].find(job => job.conversation.key === key);
    if (active) return active.conversation;
    const cached = this.conversations.get(key);
    if (cached) { await this.reconcileInterrupted(cached); return cached; }
    if (this.conversations.size >= 512) throw new DocumentError('CONFLICT', 'AI 文档会话已达到容量上限。');
    const row = this.database.prepare('SELECT record FROM conversations WHERE key=?').get(key);
    if (row) {
      const conversation: Conversation = JSON.parse(String(row.record));
      this.conversations.set(key, conversation);
      await this.reconcileInterrupted(conversation); return conversation;
    }
    if (Number(this.database.prepare('SELECT count(*) AS count FROM conversations').get()?.count) >= 512) throw new DocumentError('CONFLICT', 'AI 文档会话已达到容量上限。');
    const conversation: Conversation = { schema: 1, key, sessions: {}, messages: [], requests: [] };
    this.conversations.set(key, conversation); return conversation;
  }
  private write(conversation: Conversation) {
    this.database.prepare('INSERT INTO conversations VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET record=excluded.record').run(conversation.key, JSON.stringify(conversation));
  }
  private async reconcileInterrupted(conversation: Conversation) {
    const controller = this.dependencies.sessions();
    if (!controller?.inspect) return;
    let changed = false;
    for (const request of conversation.requests.filter(request => request.state === 'interrupted')) {
      // Read DSH's durable log; never restart an old prompt/tool round or append
      // synthetic tool results to it. An unfinished/absent turn stays blocked.
      const inspection = await controller.inspect(request.sessionId).catch(() => null);
      if (!inspection) continue;
      let currentTurn: number | undefined, turn: number | undefined, content = '';
      for (const event of inspection.events) {
        if (event.type === 'turn/start') currentTurn = event.data.turn;
        if (event.type === 'user/message' && event.data.source?.rpcId === request.id) turn = currentTurn;
        if (turn === undefined || event.data.turn !== turn) continue;
        if (event.type === 'assistant/message') content = (event.data.message?.content ?? []).filter((part: any) => part.type === 'text').map((part: any) => part.text).join('\n').slice(0, 256_000);
        if (event.type === 'turn/end') {
          request.state = event.data.reason?.kind === 'completed' ? 'done' : 'failed';
          request.error = request.state === 'done' ? undefined : '已从 DSH 日志核实：本次执行已中断，请用新请求继续。';
          const message = conversation.messages.find(message => message.id === request.assistant);
          if (message) { message.content = content; message.error = request.error; message.badge = request.error ? 'error' : undefined; }
          // The originating Session already contains this pair, including its
          // assistant response; other modes receive it through their own cursor.
          const position = conversation.messages.findIndex(message => message.id === request.assistant);
          const session = conversation.sessions[request.mode]; if (session) session.cursor = Math.max(session.cursor, position + 1);
          changed = true; break;
        }
      }
    }
    if (changed) this.write(conversation);
  }
  private async execute(job: Running, prompt: string) {
    const controller = this.controller(), request = job.request;
    let turn: number | undefined, currentTurn: number | undefined, ended = false, submitted = false, interference = false;
    let attempt = '', index = 0;
    const blocks = new Map<number, string>();
    const observe = async () => {
      for await (const frame of controller.follow({ address: { kind: 'session', sessionId: job.sessionId }, maxMessages: 100, assistantStream: true }, job.abort.signal)) {
        const live = frame.frame;
        if (frame.type === 'assistant-stream' && live && turn !== undefined) {
          if (live.type === 'start' && live.turn === turn) { attempt = live.attemptId; index = 0; blocks.clear(); }
          if (live.type === 'chunk' && live.attemptId === attempt && live.index === index++) {
            const chunk = live.chunk;
            if (chunk?.type === 'text-delta') blocks.set(chunk.index, (blocks.get(chunk.index) ?? '') + chunk.text);
            if (chunk?.type === 'block-end' && chunk.block?.type === 'text') blocks.set(chunk.index, chunk.block.text);
            job.text = [...blocks.values()].join('\n').slice(0, 256_000);
          }
        }
        for (const event of frame.type === 'snapshot' ? frame.records?.map(row => row.event) ?? [] : frame.event ? [frame.event] : []) {
          const data = event.data;
          if (event.type === 'turn/start') currentTurn = data.turn;
          if (event.type === 'user/message' && data.source?.rpcId === request.id) {
            turn = currentTurn;
            job.conversation.sessions[request.mode]!.cursor = job.conversation.messages.length - 1;
            this.write(job.conversation);
          } else if (event.type === 'user/message' && turn !== undefined && currentTurn === turn
            && !['agent-instructions', 'skill-catalog', 'runtime-context', 'time-context', 'compact-checkpoint', 'model-selection', 'tool-registry'].includes(data.source?.kind)) {
            interference = true; job.stopped = true; job.toolsAbort.abort(); job.patch?.resolve('执行轮次收到外部消息，修改已停止。');
          }
          if (turn === undefined || data.turn !== turn) continue;
          if (event.type === 'assistant/message') {
            job.text = (data.message?.content ?? []).filter((part: any) => part.type === 'text').map((part: any) => part.text).join('\n').slice(0, 256_000);
          }
          if (event.type === 'tool/call' || event.type === 'tool/result') {
            const message = job.conversation.messages.find(item => item.id === request.assistant)!;
            const calls = message.toolCalls ??= [];
            const callId = event.type === 'tool/call' ? data.callId : data.message?.toolCallId;
            const existing = calls.find(call => call.id === callId);
            if (typeof callId !== 'string') continue;
            if (event.type === 'tool/call' && !existing && typeof data.name === 'string' && calls.length < 256) {
              calls.push({ id: callId, name: data.name, done: false });
              job.events.push({ type: 'aiChat.tool', requestId: request.id, callId, phase: 'call', toolName: data.name, input: {} });
            } else if (event.type === 'tool/result' && existing && !existing.done) {
              existing.done = true;
              job.events.push({ type: 'aiChat.tool', requestId: request.id, callId, phase: 'result', toolName: existing.name, input: {} });
            }
            this.write(job.conversation);
          }
          if (event.type === 'turn/end') {
            ended = true;
            job.conversation.sessions[request.mode]!.cursor = job.conversation.messages.length;
            if (data.reason?.kind !== 'completed' || interference) throw new Error('DSH 执行已终止，请核对会话结果。');
            return;
          }
        }
      }
      throw new Error('DSH 会话连接中断，执行结果尚未确认。');
    };
    // Start the subscription before the prompt can enter the inbox.
    const observation = observe(); void observation.catch(() => {});
    const timeout = setTimeout(() => { job.toolsAbort.abort(); job.patch?.resolve('AI 请求超时。'); job.abort.abort(); }, 180_000);
    try {
      submitted = true;
      await controller.prompt({ sessionId: job.sessionId, requestId: request.id, mode: 'queue', content: [{ type: 'text', text: prompt }], clientTimeZone: 'Asia/Shanghai' }, job.abort.signal);
      await observation; request.state = 'done';
    } catch {
      request.state = ended ? 'failed' : 'interrupted';
      request.error = job.stopped ? '已请求停止；请核对 DSH 会话及恢复草稿。' : 'DSH 执行未完成；请核对原会话后继续。';
    } finally {
      clearTimeout(timeout); job.toolsAbort.abort(); job.abort.abort();
      if (submitted && !ended) await controller.cancel({ sessionId: job.sessionId }).catch(() => {});
      await observation.catch(() => {});
      const message = job.conversation.messages.find(item => item.id === request.assistant)!;
      message.content = job.text; message.error = request.error;
      if (request.error) message.badge = job.stopped ? 'stopped' : 'error';
      this.write(job.conversation); this.running.delete(job.sessionId);
    }
  }
}
