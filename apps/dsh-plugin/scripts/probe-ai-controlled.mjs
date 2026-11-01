/** Deterministic UI fault probe; explicitly NOT real-model E05/A07 acceptance. */
import { createServer } from 'node:http';
import { Readable } from 'node:stream';
import { mkdtemp, mkdir, writeFile, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { Document, Packer, Paragraph } from 'docx';
import { chromium } from '@playwright/test';
import { Context, runtimeRoot } from '../tests/runtime.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));
const installationRoot = process.env.DSH_EASYVIEW_PACKAGE_ROOT ? resolve(process.env.DSH_EASYVIEW_PACKAGE_ROOT) : root;
const host = await import(pathToFileURL(resolve(installationRoot, 'lib/index.js')).href);
const evidenceName = process.env.DSH_EASYVIEW_PACKAGE_ROOT ? 'e05-package-controlled-ai' : 'e05-controlled-ai';
const directory = await mkdtemp(resolve(tmpdir(), 'easyview-ai-ui-')), workspace = resolve(directory, 'workspace');
await mkdir(workspace); await writeFile(resolve(workspace, 'note.md'), '# Original\n\nKeep this paragraph.');
const originalWord = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('Word summary fixture'), new Paragraph('本季度完成交付。')] }] }));
await writeFile(resolve(workspace, 'report.docx'), originalWord);
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Sheet1'); sheet.getCell('A1').value = 'Workbook original'; sheet.getCell('B2').value = { formula: '1+2', result: 3 };
await workbook.xlsx.writeFile(resolve(workspace, 'book.xlsx'));
const originalWorkbook = await readFile(resolve(workspace, 'book.xlsx'));
const { LocalFileSystem } = await import(pathToFileURL(resolve(runtimeRoot, 'packages/fs/fs-local/lib/index.js')).href);
const ctx = new Context(), routes = new Map(), sessions = new Map(), prompts = [], errors = [];
ctx.logger.exporter({ export(message) { if (message.type === 'error') errors.push(message.args.map(String).join(' ')); } });
const fs = ctx.plugin(LocalFileSystem, { cwd: workspace }); await fs.await();
const dependencies = ctx.plugin(inner => {
  inner.provide('profileContext', { home: directory });
  inner.provide('workspaceRegistry', { list: () => [{ id: 'test', title: 'AI fixture', path: workspace }] });
  inner.provide('connection', { fetch: { register(route) { routes.set(route.path, route); return async () => { routes.delete(route.path); }; } } });
  inner.provide('agents', { get: id => sessions.has(id) ? { id } : undefined });
  inner.provide('tools', { schemas: agent => ['easyview_read_document', ...(sessions.get(agent.id).preset.includes('agent') ? ['easyview_edit_markdown', 'easyview_create_markdown', 'easyview_edit_workbook'] : [])].map(name => ({ name })) });
  inner.provide('systemPrompt', { assemble: async ({ scope }) => ({ tools: inner.get('tools').schemas(scope) }) });
  inner.provide('sessionController', {
    modelCatalog: async () => ({ default: null, groups: [{ id: 'controlled', models: [{ id: 'fixture' }] }] }),
    create: async ({ agentPreset }) => { const sessionId = randomUUID(); sessions.set(sessionId, { preset: agentPreset, queue: [], wake: null }); return { sessionId, agentPreset }; },
    selectSessionModel: async () => ({}),
    async *follow({ address }, signal) {
      const row = sessions.get(address.sessionId); yield { type: 'snapshot', records: [] };
      while (!signal.aborted) {
        if (row.queue.length) { yield { type: 'event', event: row.queue.shift() }; continue; }
        await new Promise(resolve => {
          const done = () => { signal.removeEventListener('abort', done); row.wake = null; resolve(); };
          row.wake = done; signal.addEventListener('abort', done, { once: true });
        });
      }
    },
    async prompt(input) {
      prompts.push(input); const turn = prompts.length, row = sessions.get(input.sessionId);
      const emit = (type, data) => { row.queue.push({ type, data: { ...data, turn } }); row.wake?.(); };
      emit('turn/start', {}); emit('user/message', { source: { rpcId: input.requestId } });
      const exec = { agent: { id: input.sessionId }, callId: `call-${turn}`, signal: new AbortController().signal };
      const content = JSON.parse(await ctx.get('easyview').ai.tool('easyview_read_document', {}, exec).catch(error => { console.error('Controlled read failed:', error.stack); throw error; }));
      if (content.format === 'docx') {
        assert.match(content.content, /Word summary fixture/);
        assert.match(content.content, /本季度完成交付。/);
        assert.equal(content.scope, 'document');
        if (input.content[0].text.includes('总结当前文档内容')) {
          assert.match(content.content, /WORD_UNSAVED_SUMMARY/);
          assert.equal(content.source, 'recovery_checkpoint');
        }
        emit('assistant/message', { message: { content: [{ type: 'text', text: `CONTROLLED_WORD_SUMMARY_${turn}` }] } });
        emit('turn/end', { reason: { kind: 'completed' } }); return { accepted: true };
      }
      if (content.sheets) {
        emit('tool/call', { step: 1, callId: exec.callId, name: 'easyview_edit_workbook', arguments: '{}' });
        const result = JSON.parse(await ctx.get('easyview').ai.tool('easyview_edit_workbook', {
          fingerprint: content.fingerprint, ops: [{ op: 'setCell', sheet: 'Sheet1', ref: 'A1', value: 'WORKBOOK_AI_CHANGED' }],
        }, exec).catch(error => { console.error('Controlled workbook tool failed:', error.stack); throw error; }));
        emit('tool/result', { step: 1, message: { toolCallId: exec.callId } });
        emit('assistant/message', { message: { content: [{ type: 'text', text: `CONTROLLED_XLSX_${result.state}` }] } });
        emit('turn/end', { reason: { kind: 'completed' } }); return { accepted: true };
      }
      assert.match(content.content, /Original/);
      if (row.preset.includes('agent')) {
        emit('tool/call', { step: 1, callId: exec.callId, name: 'easyview_edit_markdown', arguments: '{}' });
        const result = JSON.parse(await ctx.get('easyview').ai.tool('easyview_edit_markdown', {
          baseRevision: content.baseRevision, expectedText: 'Original', replacement: 'Revised by AI', description: 'Improve heading',
        }, exec)); assert.equal(result.state, 'applied');
        emit('tool/result', { step: 1, message: { toolCallId: exec.callId } });
      }
      emit('assistant/message', { message: { content: [{ type: 'text', text: 'CONTROLLED_UI_COMPLETE' }] } });
      emit('turn/end', { reason: { kind: 'completed' } }); return { accepted: true };
    },
    cancel: async () => ({ accepted: true }),
  });
}); await dependencies.await();
const plugin = ctx.plugin(host); await plugin.await();
for (let attempt = 0; !routes.has('/api/easyview/v1/bootstrap') && attempt < 100; attempt++) await new Promise(resolve => setTimeout(resolve, 10));
assert.ok(routes.size > 3000, errors.join('\n'));
const server = createServer(async (incoming, outgoing) => {
  try {
    const url = new URL(incoming.url, `http://${incoming.headers.host}`), route = routes.get(url.pathname);
    if (!route) { outgoing.writeHead(404).end(); return; }
    const request = new Request(url, { method: incoming.method, headers: incoming.headers,
      ...(incoming.method !== 'GET' && incoming.method !== 'HEAD' && { body: Readable.toWeb(incoming), duplex: 'half' }) });
    const response = await route.fetch(request); outgoing.writeHead(response.status, Object.fromEntries(response.headers));
    if (response.body) for await (const chunk of response.body) outgoing.write(chunk);
    outgoing.end();
  } catch (error) { outgoing.writeHead(500).end(error.message); }
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const browser = await chromium.launch({ channel: process.env.EASYVIEW_BROWSER_CHANNEL || 'msedge', headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
page.setDefaultTimeout(30_000);
const browserErrors = []; page.on('pageerror', error => browserErrors.push(error.message));
try {
  await page.goto(`http://127.0.0.1:${server.address().port}/api/easyview/v1/assets/index.html`);
  const welcomeAi = page.locator('.ai-chat-panel:not(.hidden):not(.suspended)');
  await welcomeAi.waitFor({ state: 'visible' });
  await page.locator('#ai-chat-toggle').click();
  await welcomeAi.waitFor({ state: 'hidden' });
  await page.locator('#ai-chat-toggle').click();
  await welcomeAi.waitFor({ state: 'visible' });
  await page.locator('.desktop-welcome-recent-open').first().click();
  await page.locator('[data-relative-path="note.md"] .workspace-entry-name').click();
  const prose = page.locator('.desktop-editor-session:not([hidden]) .ProseMirror');
  await prose.getByText('Original', { exact: false }).waitFor();
  assert.equal(await page.locator('#ai-chat-toggle').getAttribute('aria-pressed'), 'true', 'AI sidebar opens by default');
  await page.locator('.ai-chat-panel:not(.hidden):not(.suspended)').waitFor();
  await page.locator('.ai-chat-panel:not(.hidden):not(.suspended) .ai-chat-input').fill('Improve the heading only');
  await page.locator('.ai-chat-panel:not(.hidden):not(.suspended) .ai-chat-input').press('Enter');
  await prose.getByText('Revised by AI', { exact: false }).waitFor();
  await page.getByText('CONTROLLED_UI_COMPLETE', { exact: true }).waitFor();
  await page.locator('.ai-chat-tool-line').waitFor();
  await page.locator('.ai-chat-tool-line.running').waitFor({ state: 'detached' });
  assert.equal(await page.locator('.ai-chat-tool-line.running').count(), 0);
  assert.equal(await readFile(resolve(workspace, 'note.md'), 'utf8'), '# Original\n\nKeep this paragraph.');
  assert.match(await prose.innerText(), /Keep this paragraph/);
  await page.screenshot({ path: resolve(root, `docs/evidence/${evidenceName}.png`), fullPage: true });
  await prose.click(); await page.keyboard.press('ControlOrMeta+z');
  await prose.getByText('Original', { exact: false }).waitFor();
  await page.locator('[data-easyview-draft-status="confirmed"]').waitFor();
  await page.reload();
  await page.locator('.desktop-welcome-recent-open').first().click();
  await page.locator('[data-relative-path="note.md"] .workspace-entry-name').click();
  await prose.getByText('Original', { exact: false }).waitFor();
  assert.equal(await page.locator('#ai-chat-toggle').getAttribute('aria-pressed'), 'true', 'AI sidebar reopens on a fresh visit');
  await page.getByText('CONTROLLED_UI_COMPLETE', { exact: true }).waitFor();
  assert.equal(prompts.length, 1, 'Hydration must not prompt the model again');
  const markdownPanel = await page.locator('.ai-chat-panel:not(.suspended):not(.hidden)').elementHandle();
  await page.locator('[data-relative-path="book.xlsx"] .workspace-entry-name').click();
  await page.locator('.preview-spreadsheet-loading').waitFor({ state: 'detached' });
  await page.waitForFunction(panel => panel.classList.contains('suspended'), markdownPanel);
  await page.locator('.ai-chat-panel:not(.suspended) .ai-chat-mode-label').getByText('Agent', { exact: true }).waitFor();
  const officeInput = page.locator('.ai-chat-panel:not(.suspended) .ai-chat-input:visible');
  await officeInput.fill('Propose a workbook change for rejection'); await officeInput.press('Enter');
  await page.locator('.ai-chat-office-edit.pending').waitFor();
  assert.deepEqual(await readFile(resolve(workspace, 'book.xlsx')), originalWorkbook);
  await page.locator('.ai-chat-office-edit.pending .ai-chat-office-btn.reject').click();
  await page.getByText('CONTROLLED_XLSX_rejected', { exact: true }).waitFor();
  assert.deepEqual(await readFile(resolve(workspace, 'book.xlsx')), originalWorkbook);
  await officeInput.fill('Propose a workbook change for application'); await officeInput.press('Enter');
  await page.locator('.ai-chat-office-edit.pending').waitFor();
  await page.locator('.ai-chat-office-edit.pending .ai-chat-office-btn.apply').click();
  await page.getByText('CONTROLLED_XLSX_committed', { exact: true }).waitFor();
  await page.locator('.ai-chat-office-edit.applied .ai-chat-office-btn.undo').waitFor();
  const changedWorkbook = await JSZip.loadAsync(await readFile(resolve(workspace, 'book.xlsx')));
  assert.match(await changedWorkbook.file('xl/sharedStrings.xml').async('string'), /WORKBOOK_AI_CHANGED/);
  assert.match(await changedWorkbook.file('xl/worksheets/sheet1.xml').async('string'), /<f[^>]*>1\+2<\/f>/);
  await page.locator('.preview-spreadsheet-loading').waitFor({ state: 'detached' });
  await page.locator('.ai-chat-office-edit.applied .ai-chat-office-btn.undo').click();
  await page.locator('.ai-chat-office-edit.undone').waitFor();
  assert.deepEqual(await readFile(resolve(workspace, 'book.xlsx')), originalWorkbook);
  await page.screenshot({ path: resolve(root, `docs/evidence/${process.env.DSH_EASYVIEW_PACKAGE_ROOT ? 'e05-package-controlled-xlsx' : 'e05-controlled-xlsx'}.png`), fullPage: true });
  await page.reload();
  await page.locator('.ai-chat-panel:not(.hidden):not(.suspended)').waitFor({ state: 'visible' });
  await page.locator('#ai-chat-toggle').click();
  await page.locator('.desktop-welcome-recent-open').first().click();
  await page.locator('[data-relative-path="book.xlsx"] .workspace-entry-name').click();
  // Open the lazy AI panel after the preview is fully active, as on a fresh visit.
  await page.locator('.x-spreadsheet-overlayer').waitFor({ state: 'visible' });
  await page.locator('.preview-spreadsheet-loading').waitFor({ state: 'detached' });
  assert.equal(await page.locator('#ai-chat-toggle').getAttribute('aria-pressed'), 'false', 'Manual close survives document activation');
  await page.locator('#ai-chat-toggle').click(); await page.locator('.ai-chat-office-edit.undone').waitFor();
  assert.equal(prompts.length, 3, 'Office projection hydration must not replay any model request');
  await page.locator('[data-relative-path="report.docx"] .workspace-entry-name').click();
  await page.locator('.preview-word-editor .layout-paragraph').filter({ hasText: 'Word summary fixture' }).first().waitFor();
  const wordPanel = page.locator('.ai-chat-panel:not(.suspended):not(.hidden)');
  await wordPanel.locator('.ai-chat-mode-current').click();
  await wordPanel.locator('.ai-chat-mode-option').filter({ hasText: 'Chat' }).click();
  for (const [index, prompt] of ['总结当前页内容', '总结当前文档内容'].entries()) {
    if (index === 1) {
      await page.locator('.preview-word-editor .layout-paragraph').filter({ hasText: 'Word summary fixture' }).first().click();
      await page.keyboard.press('End'); await page.keyboard.type(' WORD_UNSAVED_SUMMARY');
    }
    await wordPanel.locator('.ai-chat-input').fill(prompt);
    await wordPanel.locator('.ai-chat-input').press('Enter');
    await wordPanel.getByText(`CONTROLLED_WORD_SUMMARY_${index + 4}`, { exact: true }).waitFor();
  }
  assert.equal(prompts.length, 5);
  assert.deepEqual(await readFile(resolve(workspace, 'report.docx')), originalWord);
  assert.deepEqual(browserErrors, []); assert.deepEqual(errors, []);
  const evidence = { status: 'passed', kind: 'controlled-response UI integration; NOT real-model acceptance', browser: browser.version(),
    installationRoot, aiSidebarDefaultOpen: true, welcomeAiVisible: true, manualSidebarToggle: true,
    originalAiPanel: true, originalMarkdownEditor: true, localPatchDurableReceipt: true, originalUnchanged: true,
    batchUndo: true, reloadHydrationNoReprompt: true, originalToolFeedback: true,
    docx: { currentTabSummary: true, documentSummary: true, bodyTextRead: true, latestUnsavedBodyRead: true, originalUnchanged: true },
    xlsx: { originalProposalCard: true, rejectDoesNotWrite: true, acceptWritesOnce: true, formulaRetained: true, originalUndoCard: true, exactUndoBytes: true,
      reloadHydrationNoReprompt: true, lazyHistoryAfterPreviewReady: true }, browserErrors, errors, workspace };
  await writeFile(resolve(root, `docs/evidence/${evidenceName}.json`), JSON.stringify(evidence, null, 2) + '\n');
  console.log(JSON.stringify(evidence, null, 2));
} catch (error) {
  console.error(JSON.stringify(await page.locator('.ai-chat-input').evaluateAll(elements => elements.map(element => ({
    rect: element.getBoundingClientRect().toJSON(), connected: element.isConnected,
    ancestors: [...(function* () { for (let node = element; node; node = node.parentElement) yield node; })()].map(node => ({ tag: node.tagName, id: node.id, class: node.className, inert: node.inert,
      display: getComputedStyle(node).display, visibility: getComputedStyle(node).visibility, contentVisibility: getComputedStyle(node).contentVisibility })),
  }))), null, 2));
  await page.screenshot({ path: resolve(root, `docs/evidence/${evidenceName}-failure.png`), fullPage: true });
  await writeFile(resolve(root, `docs/evidence/${evidenceName}.json`), JSON.stringify({ status: 'failed', error: error.stack, errors, browserErrors }, null, 2) + '\n');
  throw error;
} finally {
  await browser.close(); await new Promise(resolve => server.close(resolve));
  await plugin.dispose(); await dependencies.dispose(); await fs.dispose(); assert.equal(routes.size, 0);
}
