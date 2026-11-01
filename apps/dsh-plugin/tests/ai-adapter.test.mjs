import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { tmpdir } from 'node:os';
import { randomUUID } from 'node:crypto';
import { Context, runtimeRoot } from './runtime.mjs';
import * as preset from '../src/ai-preset.ts';
import { DshAiAdapter } from '../src/host/ai-adapter.ts';
import { DraftStore } from '../src/host/draft-store.ts';
import { DocumentProvider, contentHash } from '../src/host/documents.ts';
import { AI_TOOLS } from '../src/shared/ai.ts';
import { OperationJournal } from '../src/host/operation-journal.ts';
import { FileCommands } from '../src/host/file-commands.ts';
import { XlsxProposals } from '../src/host/xlsx-proposals.ts';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const runtime = path => import(pathToFileURL(resolve(runtimeRoot, path, 'lib/index.js')).href);

test('inherited preset exposes only document schemas and denies Agent-local team tools, including late additions', async () => {
  const { SystemPrompt } = await runtime('packages/core/system-prompt');
  const { ToolRuntime } = await runtime('packages/core/tools');
  const { createScope, bindScopeParent } = await runtime('packages/core/scope');
  const ctx = new Context();
  await ctx.plugin(SystemPrompt, {}).await(); await ctx.plugin(ToolRuntime).await();
  await ctx.plugin(inner => inner.provide('easyview', { ai: { tool: async name => name } })).await();
  const generic = name => ({ name, description: name, parameters: { type: 'object', properties: {} },
    output: { schema: { type: 'string' }, render: (_args, value) => [{ type: 'text', text: value }] }, execute: async () => 'unsafe' });
  await ctx.plugin(Object.assign(inner => inner.tools.register(generic('shell')), { inject: ['tools'] })).await();
  try {
    for (const mode of ['chat', 'agent']) {
      const key = {}, agent = { id: randomUUID() };
      let standing, child;
      const owner = ctx.plugin(Object.assign(inner => {
        standing = createScope(inner, key); child = createScope(inner, agent); bindScopeParent(agent, key);
      }, { inject: ['tools', 'systemPrompt', 'easyview'] })); await owner.await();
      await standing.ctx.plugin(preset, { mode }).await();
      child.ctx.tools.register(generic('spawn_teammate'));
      assert.deepEqual((await ctx.systemPrompt.assemble({ scope: agent })).tools.map(tool => tool.name).sort(), [...AI_TOOLS[mode]].sort());
      child.ctx.tools.register(generic('late_writer'));
      assert.deepEqual((await ctx.systemPrompt.assemble({ scope: agent })).tools.map(tool => tool.name).sort(), [...AI_TOOLS[mode]].sort());
      for (const name of ['shell', 'spawn_teammate', 'late_writer', ...(mode === 'chat' ? ['easyview_edit_markdown'] : [])]) {
        const result = await ctx.tools.execute({ name, agent, callId: randomUUID(), signal: new AbortController().signal, arguments: {} });
        assert.equal(result.isError, true); assert.doesNotMatch(JSON.stringify(result.content), /unsafe/);
      }
      assert.deepEqual(ctx.tools.schemas().map(tool => tool.name), ['shell']);
      await owner.dispose();
    }
  } finally { await ctx.fiber.dispose(); }
});

test('real DSH scopes exclude shell/subagents, Chat has no writer, and the execution guard rejects later injected tools', async () => {
  const { SystemPrompt } = await runtime('packages/core/system-prompt');
  const { ToolRuntime } = await runtime('packages/core/tools');
  const { createScope } = await runtime('packages/core/scope');
  const ctx = new Context();
  const prompt = ctx.plugin(SystemPrompt, {}); await prompt.await();
  const tools = ctx.plugin(ToolRuntime); await tools.await();
  const service = ctx.plugin(inner => inner.provide('easyview', { ai: { tool: async name => name } })); await service.await();
  const generic = name => ({ name, description: name, parameters: { type: 'object', properties: {} },
    output: { schema: { type: 'string' }, render: (_args, value) => [{ type: 'text', text: value }] }, execute: async () => 'unsafe' });
  const rootTools = ctx.plugin(Object.assign(inner => {
    inner.tools.register(generic('shell')); inner.tools.register(generic('subagent'));
  }, { inject: ['tools'] })); await rootTools.await();
  try {
    for (const mode of ['chat', 'agent']) {
      const agent = { id: randomUUID() }; let scope;
      const minter = ctx.plugin(Object.assign(inner => { scope = createScope(inner, agent); }, { inject: ['tools', 'systemPrompt', 'easyview'] })); await minter.await();
      const binding = scope.ctx.plugin(preset, { mode }); await binding.await();
      assert.deepEqual(ctx.tools.schemas(agent).map(tool => tool.name).sort(), [...AI_TOOLS[mode]].sort());
      assert.deepEqual(ctx.tools.schemas().map(tool => tool.name).sort(), ['shell', 'subagent']);
      for (const name of ['shell', 'subagent', ...(mode === 'chat' ? ['easyview_edit_markdown', 'easyview_create_markdown'] : [])]) {
        const result = await ctx.tools.execute({ name, agent, callId: randomUUID(), signal: new AbortController().signal, arguments: {} });
        assert.equal(result.isError, true);
      }
      scope.ctx.tools.register(generic('injected_writer'));
      const result = await ctx.tools.execute({ name: 'injected_writer', agent, callId: randomUUID(), signal: new AbortController().signal, arguments: {} });
      assert.equal(result.isError, true); assert.doesNotMatch(JSON.stringify(result.content), /unsafe/);
      await minter.dispose();
      assert.deepEqual(ctx.tools.schemas(agent).map(tool => tool.name).sort(), ['shell', 'subagent']);
    }
  } finally { await rootTools.dispose(); await service.dispose(); await tools.dispose(); await prompt.dispose(); }
});

function controlledSessions() {
  const rows = new Map(), prompts = [], selections = []; let onPrompt = async () => {};
  const controller = {
    modelCatalog: async () => ({ default: null, groups: [{ id: 'controlled', models: [{ id: 'fixture' }] }] }),
    create: async ({ agentPreset }) => { const sessionId = randomUUID(); rows.set(sessionId, { agentPreset, queue: [], wake: null, followed: false }); return { sessionId, agentPreset }; },
    selectSessionModel: async selection => { selections.push(selection); },
    async *follow({ address }, signal) {
      const row = rows.get(address.sessionId); row.followed = true;
      yield { type: 'snapshot', records: [] };
      while (!signal.aborted) {
        if (row.queue.length) { yield { type: 'event', event: row.queue.shift() }; continue; }
        await new Promise(resolve => {
          const finish = () => { signal.removeEventListener('abort', finish); row.wake = null; resolve(); };
          row.wake = finish; signal.addEventListener('abort', finish, { once: true });
        });
      }
    },
    async prompt(input) {
      const row = rows.get(input.sessionId); assert.equal(row.followed, true, 'Must follow before admission'); prompts.push(input);
      const emit = (type, data) => { row.queue.push({ type, data }); row.wake?.(); };
      emit('turn/start', { turn: prompts.length });
      emit('user/message', { turn: prompts.length, source: { rpcId: input.requestId } });
      await onPrompt(input, emit);
      emit('assistant/message', { turn: prompts.length, message: { content: [{ type: 'text', text: `Controlled reply ${prompts.length}` }] } });
      emit('turn/end', { turn: prompts.length, reason: { kind: 'completed' } }); return { accepted: true };
    },
    cancel: async () => ({ accepted: true }),
  };
  return { controller, rows, prompts, selections, onPrompt: callback => { onPrompt = callback; } };
}
async function fixture() {
  const directory = await mkdtemp(resolve(tmpdir(), 'easyview-ai-')); const workspace = resolve(directory, 'workspace'); await mkdir(workspace);
  await writeFile(resolve(workspace, 'note.md'), '# Original\nKeep this paragraph.');
  const { LocalFileSystem } = await runtime('packages/fs/fs-local'); const ctx = new Context();
  const fsPlugin = ctx.plugin(LocalFileSystem, { cwd: workspace }); await fsPlugin.await();
  const documents = new DocumentProvider(() => ctx.get('fs'), () => [{ id: 'test', title: 'test', path: workspace }], 1024 * 1024);
  const document = await documents.open('test', 'note.md'), identity = await documents.draftIdentity(document.sessionId);
  const store = await DraftStore.open(directory), workbenchId = randomUUID(), generation = randomUUID();
  const key = contentHash(Buffer.from(`${workbenchId}\0${identity.fileIdentity}`)); store.begin(key, generation);
  const sessions = controlledSessions();
  const dependencies = { sessions: () => sessions.controller, tools: () => ({ schemas: agent => AI_TOOLS[sessions.rows.get(agent.id).agentPreset.includes('agent') ? 'agent' : 'chat'].map(name => ({ name })) }),
    agent: id => sessions.rows.has(id) ? { id } : undefined, workspaceExists: id => id === 'test' };
  const journal = new OperationJournal(directory, store, documents), files = new FileCommands(documents, journal, 1024 * 1024);
  const ai = new DshAiAdapter(directory, store, documents, dependencies, files);
  const address = { workspaceId: 'test', workbenchId, generation, documentId: document.sessionId };
  return { ai, address, directory, workspace, store, key, sessions, documents, dependencies,
    journal, close: async () => { await ai.close(); await journal.close(); await store.close(); await fsPlugin.dispose(); } };
}
async function settled(ai, address) {
  for (let attempt = 0; attempt < 200; attempt++) {
    const state = await ai.state(address); if (!state.activeRequestId) return state;
    await new Promise(resolve => setTimeout(resolve, 5));
  }
  throw new Error('Controlled execution did not settle');
}

for (const mounted of [false, true]) for (const draft of [false, true]) test(`Chat reads Word body for both summary prompts (editor mounted: ${mounted}, draft: ${draft})`, async () => {
  const f = await fixture();
  const zip = new JSZip();
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>项目总结 &amp; 下一步</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>本季度完成交付。</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>');
  await writeFile(resolve(f.workspace, 'report.docx'), await zip.generateAsync({ type: 'nodebuffer' }));
  const opened = await f.documents.open('test', 'report.docx');
  const address = { ...f.address, documentId: opened.sessionId, generation: mounted ? randomUUID() : undefined };
  const identity = await f.documents.draftIdentity(opened.sessionId);
  const key = contentHash(Buffer.from(`${address.workbenchId}\0${identity.fileIdentity}`));
  const generation = address.generation ?? randomUUID();
  if (mounted || draft) f.store.begin(key, generation);
  if (draft) {
    zip.file('word/document.xml', '<x:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:body><x:p><x:r><x:t>项目总结 &amp; 下一步</x:t></x:r></x:p><x:p><x:del><x:r><x:delText>旧的错误结论</x:delText></x:r></x:del><x:r><x:t>本季度完成交付。未保存的最新进展。</x:t></x:r></x:p></x:body></x:document>');
    await f.store.checkpoint({ key, generation, requestId: randomUUID(), editorRevision: 1, expectedCheckpointRevision: 0,
      format: 'docx', sourceDiskHash: identity.diskHash, bytes: await zip.generateAsync({ type: 'nodebuffer' }) });
  }
  let readError;
  f.sessions.onPrompt(async input => {
    try {
      const context = JSON.parse(await f.ai.tool('easyview_read_document', {}, {
        agent: { id: input.sessionId }, callId: randomUUID(), signal: new AbortController().signal,
      }));
      assert.equal(context.path, 'report.docx');
      assert.equal(context.source, draft ? 'recovery_checkpoint' : 'saved_document');
      assert.equal(context.scope, 'document');
      assert.equal(context.baseRevision, draft ? 1 : 0);
      assert.match(context.content, /项目总结 & 下一步/);
      assert.match(context.content, /本季度完成交付。/);
      assert.equal(context.content.includes('未保存的最新进展。'), draft);
      assert.doesNotMatch(context.content, /旧的错误结论/);
    } catch (error) { readError = error; throw error; }
  });
  try {
    for (const userMessage of ['总结当前页内容', '总结当前文档内容']) {
      await f.ai.send(address, { requestId: randomUUID(), mode: 'chat', model: 'controlled / fixture', userMessage });
      const state = await settled(f.ai, address);
      assert.equal(readError, undefined);
      assert.equal(state.messages.at(-1).error, undefined, state.messages.at(-1).error);
    }
    assert.equal(f.journal.list().length, 0);
    assert.equal(contentHash(await readFile(resolve(f.workspace, 'report.docx'))), identity.diskHash);
  } finally { await f.close(); }
});

test('Chat inspects an authorized saved XLSX before its editor generation exists, without exposing a write tool', async () => {
  const f = await fixture(); await f.ai.close();
  const workbook = new ExcelJS.Workbook(); workbook.addWorksheet('Sheet1').getCell('A1').value = 'Read-only context';
  await workbook.xlsx.writeFile(resolve(f.workspace, 'book.xlsx'));
  const opened = await f.documents.open('test', 'book.xlsx'), address = { ...f.address, documentId: opened.sessionId, generation: undefined };
  const xlsx = new XlsxProposals(f.directory, f.store, f.journal, f.documents, { clean() { throw new Error('Chat must not enter write admission'); } });
  const ai = new DshAiAdapter(f.directory, f.store, f.documents, f.dependencies, undefined, xlsx);
  f.sessions.onPrompt(async input => {
    const exec = { agent: { id: input.sessionId }, callId: 'inspect-xlsx', signal: new AbortController().signal };
    const inspection = JSON.parse(await ai.tool('easyview_read_document', {}, exec));
    assert.equal(inspection.sheets[0].headers[0], 'Read-only context');
    await assert.rejects(ai.tool('easyview_edit_workbook', {}, exec), /No active authorized/);
  });
  try {
    await ai.send(address, { requestId: 'read-before-editor-ready', mode: 'chat', model: 'controlled / fixture', userMessage: 'Read this workbook' });
    const state = await settled(ai, address); assert.equal(state.messages.at(-1).error, undefined); assert.equal(f.journal.list().length, 0);
  } finally { await ai.close(); await xlsx.close(); await f.close(); }
});

test('DSH adapter replays receipts, selects only the Session model, and injects Host history increments between modes', async () => {
  const f = await fixture();
  f.sessions.onPrompt(async (_input, emit) => {
    const turn = f.sessions.prompts.length;
    emit('tool/call', { turn, step: 1, callId: `read-${turn}`, name: 'easyview_read_document', arguments: '{}' });
    emit('tool/call', { turn, step: 1, callId: `read-${turn}`, name: 'easyview_read_document', arguments: '{}' });
    emit('tool/result', { turn, step: 1, message: { toolCallId: `read-${turn}` } });
  });
  try {
    const input = { requestId: 'chat-request', mode: 'chat', model: 'controlled / fixture', userMessage: 'Review this document' };
    await f.ai.send(f.address, input); let state = await settled(f.ai, f.address);
    assert.match(state.messages[1].content, /Controlled reply 1/); assert.equal(state.messages[1].error, undefined);
    assert.deepEqual(state.messages[1].toolCalls, [{ id: 'read-1', name: 'easyview_read_document', done: true }]);
    await f.ai.send(f.address, input); assert.equal(f.sessions.prompts.length, 1);
    await assert.rejects(f.ai.send(f.address, { ...input, userMessage: 'different' }), { code: 'CONFLICT' });
    await f.ai.send(f.address, { ...input, requestId: 'agent-request', mode: 'agent', userMessage: 'Improve it' });
    state = await settled(f.ai, f.address);
    assert.equal(f.sessions.rows.size, 2); assert.equal(f.sessions.selections.length, 2);
    assert.match(f.sessions.prompts[1].content[0].text, /Review this document/);
    assert.match(f.sessions.prompts[1].content[0].text, new RegExp(state.messages[0].id));
    await f.ai.send(f.address, { ...input, requestId: 'same-mode-request', mode: 'agent', userMessage: 'Continue' });
    state = await settled(f.ai, f.address);
    assert.doesNotMatch(f.sessions.prompts[2].content[0].text, /Controlled reply 2/, 'Same-session history must not be sent twice');
    assert.equal((await readFile(resolve(f.workspace, 'note.md'), 'utf8')), '# Original\nKeep this paragraph.');
    const ids = state.messages.map(message => message.id);
    await f.ai.close();
    const reopened = new DshAiAdapter(f.directory, f.store, f.documents, { sessions: () => f.sessions.controller, tools: () => undefined, agent: () => undefined, workspaceExists: () => true });
    const hydrated = await reopened.state(f.address);
    assert.deepEqual(hydrated.messages.map(message => message.id), ids);
    assert.deepEqual(hydrated.messages[1].toolCalls, state.messages[1].toolCalls); await reopened.close();
  } finally { await f.close(); }
});

test('an empty Markdown draft can be started once; no-change patches never wait for a nonexistent editor checkpoint', async () => {
  const f = await fixture();
  await writeFile(resolve(f.workspace, 'note.md'), '');
  f.documents.saved(f.address.documentId, contentHash(Buffer.from('')));
  f.sessions.onPrompt(async input => {
    const execution = { agent: { id: input.sessionId }, callId: 'initial-content', signal: new AbortController().signal };
    const patch = { baseRevision: 0, expectedText: '', replacement: '# First draft\n', description: 'Start the document' };
    const operation = f.ai.tool('easyview_edit_markdown', patch, execution);
    for (let attempt = 0; attempt < 100; attempt++) {
      const event = (await f.ai.state(f.address)).events.find(event => event.type === 'aiChat.applyPatches');
      if (event) {
        await f.store.checkpoint({ key: f.key, generation: f.address.generation, requestId: randomUUID(), editorRevision: 1,
          expectedCheckpointRevision: 0, format: 'markdown', sourceDiskHash: contentHash(Buffer.from('')),
          bytes: Buffer.from('# First draft\n'), aiOperationId: event.operationId });
        await f.ai.acknowledge(f.address, input.requestId, event.operationId);
        assert.equal(JSON.parse(await operation).state, 'applied');
        assert.equal(JSON.parse(await f.ai.tool('easyview_edit_markdown', { baseRevision: 1, expectedText: '# First draft', replacement: '# First draft', description: 'No change' },
          { ...execution, callId: 'no-change' })).state, 'no_change');
        return;
      }
      await new Promise(resolve => setTimeout(resolve, 5));
    }
    throw new Error('No initial-content patch');
  });
  try {
    await f.ai.send(f.address, { requestId: 'empty-draft', mode: 'agent', model: 'controlled / fixture', userMessage: 'Start a document' });
    const state = await settled(f.ai, f.address); assert.equal(state.messages[1].error, undefined);
    assert.equal(await readFile(resolve(f.workspace, 'note.md'), 'utf8'), '');
    assert.equal(new TextDecoder().decode(await f.store.bytes(f.store.load(f.key))), '# First draft\n');
  } finally { await f.close(); }
});

test('AI checkpoint receipts expire on restart; mismatched bytes cannot acknowledge a model patch', async () => {
  const directory = await mkdtemp(resolve(tmpdir(), 'easyview-ai-receipt-'));
  let store = await DraftStore.open(directory);
  const key = 'draft', generation = randomUUID(), bytes = new TextEncoder().encode('expected');
  const operationId = '1'.repeat(64);
  store.begin(key, generation);
  store.prepareAiPatch({ operationId, key, generation, baseRevision: 0, outputHash: contentHash(bytes), inputHash: '2'.repeat(64), state: 'pending' });
  await assert.rejects(store.checkpoint({ key, generation, requestId: randomUUID(), editorRevision: 1, expectedCheckpointRevision: 0,
    format: 'markdown', sourceDiskHash: '3'.repeat(64), bytes: new TextEncoder().encode('concurrent input'), aiOperationId: operationId }), { code: 'CONFLICT' });
  assert.equal(store.load(key), null); assert.equal(store.aiPatch(operationId).state, 'pending');
  await store.close(); store = await DraftStore.open(directory);
  try {
    assert.equal(store.aiPatch(operationId).state, 'expired');
    await assert.rejects(store.checkpoint({ key, generation, requestId: randomUUID(), editorRevision: 1, expectedCheckpointRevision: 0,
      format: 'markdown', sourceDiskHash: '3'.repeat(64), bytes, aiOperationId: operationId }), { code: 'CONFLICT' });
  } finally { await store.close(); }
});

test('the scoped AI creation tool reuses a durable receipt and refuses existing paths or a changed retry payload', async () => {
  const f = await fixture();
  f.sessions.onPrompt(async input => {
    const exec = { agent: { id: input.sessionId }, callId: 'new-document', signal: new AbortController().signal };
    const args = { path: 'new.md', content: '# New document\n' };
    const first = JSON.parse(await f.ai.tool('easyview_create_markdown', args, exec));
    assert.equal(first.state, 'committed'); assert.equal(first.canUndo, false);
    assert.deepEqual(JSON.parse(await f.ai.tool('easyview_create_markdown', args, exec)), first);
    await assert.rejects(f.ai.tool('easyview_create_markdown', { ...args, content: 'overwrite' }, exec), { code: 'CONFLICT' });
    await assert.rejects(f.ai.tool('easyview_create_markdown', { path: 'note.md', content: 'overwrite' }, { ...exec, callId: 'existing-document' }), { code: 'CONFLICT' });
  });
  try {
    await f.ai.send(f.address, { requestId: 'new-document-request', mode: 'agent', model: 'controlled / fixture', userMessage: 'Create new.md' });
    assert.equal((await settled(f.ai, f.address)).messages[1].error, undefined);
    assert.equal(await readFile(resolve(f.workspace, 'new.md'), 'utf8'), '# New document\n');
    assert.match(await readFile(resolve(f.workspace, 'note.md'), 'utf8'), /Original/);
    assert.equal(f.journal.list().length, 1);
  } finally { await f.close(); }
});

test('cancelling while the document is being read prevents a late patch from entering the UI', async () => {
  const f = await fixture();
  let entered, release, observed;
  const reading = new Promise(resolve => { entered = resolve; }), barrier = new Promise(resolve => { release = resolve; });
  const outcome = new Promise(resolve => { observed = resolve; });
  const read = f.documents.read.bind(f.documents); let pause = false;
  f.documents.read = async (...args) => {
    const bytes = await read(...args);
    if (pause) { pause = false; entered(); await barrier; }
    return bytes;
  };
  f.sessions.onPrompt(async input => {
    pause = true;
    try {
      await f.ai.tool('easyview_edit_markdown', { baseRevision: 0, expectedText: 'Original', replacement: 'Too late', description: 'Late update' },
        { agent: { id: input.sessionId }, callId: 'cancelled-call', signal: new AbortController().signal });
      observed('unexpected success');
    } catch (error) { observed(String(error)); }
  });
  try {
    await f.ai.send(f.address, { requestId: 'cancelled-request', mode: 'agent', model: 'controlled / fixture', userMessage: 'Improve it' });
    await reading; await f.ai.cancel(f.address, 'cancelled-request'); release();
    assert.match(await outcome, /stopped/);
    const state = await settled(f.ai, f.address);
    assert.deepEqual(state.events, []); assert.equal(f.store.load(f.key), null);
    assert.match(await readFile(resolve(f.workspace, 'note.md'), 'utf8'), /Original/);
  } finally { release(); await f.close(); }
});

test('AI Markdown success requires the same durable checkpoint transaction; a lost UI ACK does not apply patches twice', async () => {
  const f = await fixture();
  f.sessions.onPrompt(async (input) => {
    const operation = f.ai.tool('easyview_edit_markdown', { baseRevision: 0, expectedText: 'Original', replacement: 'Revised', description: 'Improve heading' },
      { agent: { id: input.sessionId }, callId: 'tool-call', signal: new AbortController().signal });
    for (let attempt = 0; attempt < 100; attempt++) {
      const state = await f.ai.state(f.address); const event = state.events.find(event => event.type === 'aiChat.applyPatches');
      if (event) {
        const record = await f.store.checkpoint({ key: f.key, generation: f.address.generation, requestId: randomUUID(), editorRevision: 1,
          expectedCheckpointRevision: 0, format: 'markdown', sourceDiskHash: (await f.documents.draftIdentity(f.address.documentId)).diskHash,
          bytes: new TextEncoder().encode('# Revised\nKeep this paragraph.'), aiOperationId: event.operationId });
        assert.equal(f.store.aiPatch(event.operationId).checkpointRevision, record.checkpointRevision);
        // Report a lost ACK/cancel race: the Host must consult the receipt.
        await f.ai.acknowledge(f.address, input.requestId, event.operationId, 'ACK transport lost');
        assert.equal(JSON.parse(await operation).state, 'applied');
        assert.equal(JSON.parse(await f.ai.tool('easyview_edit_markdown', { baseRevision: 0, expectedText: 'Original', replacement: 'Revised', description: 'Improve heading' },
          { agent: { id: input.sessionId }, callId: 'tool-call', signal: new AbortController().signal })).state, 'applied');
        await assert.rejects(f.ai.tool('easyview_edit_markdown', { baseRevision: 1, expectedText: 'Revised', replacement: 'Other', description: 'Changed payload' },
          { agent: { id: input.sessionId }, callId: 'tool-call', signal: new AbortController().signal }), /payload conflict/);
        return;
      }
      await new Promise(resolve => setTimeout(resolve, 5));
    }
    throw new Error('No patch event');
  });
  try {
    await f.ai.send(f.address, { requestId: 'patch-request', mode: 'agent', model: 'controlled / fixture', userMessage: 'Improve heading' });
    const state = await settled(f.ai, f.address); assert.equal(state.messages[1].error, undefined);
    assert.equal(new TextDecoder().decode(await f.store.bytes(f.store.load(f.key))), '# Revised\nKeep this paragraph.');
    assert.match(await readFile(resolve(f.workspace, 'note.md'), 'utf8'), /Original/);
  } finally { await f.close(); }
});
