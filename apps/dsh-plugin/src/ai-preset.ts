/** Loaded inside a real preset scope. Never register these tools globally. */
import { AI_TOOLS, type AiTools, type AiExecution } from './shared/ai';
interface PresetContext {
  tools: AiTools;
  easyview: { ai?: { tool(name: string, args: unknown, execution: AiExecution): Promise<string> } | null };
  on(event: 'system-prompt/assemble', handler: (assembly: { tools: Array<{ name: string }> }, context: unknown,
    next: () => Promise<{ tools: Array<{ name: string }> }>) => Promise<{ tools: Array<{ name: string }> }>): unknown;
}
export const name = 'easyview-ai-preset';
export const inject = ['tools', 'easyview'];
export function apply(ctx: PresetContext, config: { mode: 'chat' | 'agent' }) {
  if (config.mode !== 'chat' && config.mode !== 'agent') throw new Error('Invalid EasyView preset mode');
  const names = AI_TOOLS[config.mode];
  // Presets are standing parent scopes: allow:[] also masks their own tools
  // when an Agent inherits them. Mask current globals and enforce the model
  // surface after assembly; the monotonic guard also denies later/local tools.
  ctx.tools.restrict({ deny: ctx.tools.schemas().map(tool => tool.name) });
  ctx.tools.presentAs('native');
  ctx.tools.guard(exec => names.includes(exec.name) ? undefined : 'Only this workbench’s authorized document tools are available.');
  ctx.on('system-prompt/assemble', async (_assembly, _context, next) => {
    const assembly = await next();
    return { ...assembly, tools: assembly.tools.filter(tool => names.includes(tool.name)) };
  });
  const output = { schema: { type: 'string' as const }, render: (_args: unknown, value: unknown) => [{ type: 'text' as const, text: String(value) }] };
  for (const toolName of names) ctx.tools.register({
    name: toolName,
    description: toolName === 'easyview_create_markdown'
      ? 'Create a new Markdown file in an existing authorized workspace directory when the user asks for a new document. Never overwrites an existing file. Returns a durable operation ID; file creation cannot be undone by deleting files.'
      : toolName === 'easyview_edit_workbook'
      ? 'Propose XLSX cell value/formula updates using the fingerprint from easyview_read_document. The original EasyView diff card waits for the user to Apply or Reject. Never claim success before the returned committed state. Dirty workbooks must be saved first; after a committed proposal, ask the user to start a new request for further edits.'
      : toolName === 'easyview_read_document'
      ? 'Read the active authorized document: Word (DOCX) body text from its confirmed recovery checkpoint or saved original, Markdown/text recovery checkpoint with baseRevision, or saved XLSX workbook with fingerprint, sheets, headers and rows. Word scope is the whole active document, not a rendered physical page; do not claim a page-specific summary. Respect the truncated flag when present.'
      : 'Edit the active Markdown draft using one exact, unique expectedText anchor at baseRevision. An empty expectedText is allowed only when the entire document is empty. For continuation, include a unique suffix and append to it. Changes appear in the original editor with highlights and undo; this never saves the original file.',
    parameters: toolName === 'easyview_edit_workbook' ? { type: 'object', properties: {
      fingerprint: { type: 'string', minLength: 1 }, ops: { type: 'array', minItems: 1, maxItems: 1000, items: {
        type: 'object', properties: { op: { const: 'setCell', type: 'string' }, sheet: { type: 'string', minLength: 1, maxLength: 31 },
          ref: { type: 'string', pattern: '^[A-Z]{1,3}[1-9][0-9]{0,6}$' }, value: { type: ['string', 'number', 'boolean', 'null'] }, formula: { type: 'string', minLength: 1, maxLength: 8192 } },
        required: ['op', 'sheet', 'ref'], additionalProperties: false,
        oneOf: [{ required: ['value'], not: { required: ['formula'] } }, { required: ['formula'], not: { required: ['value'] } }],
      } } }, required: ['fingerprint', 'ops'], additionalProperties: false }
      : toolName === 'easyview_create_markdown' ? { type: 'object', properties: {
      path: { type: 'string', minLength: 1, description: 'Workspace-relative path ending in .md, in an existing directory.' }, content: { type: 'string' } },
      required: ['path', 'content'], additionalProperties: false }
      : toolName === 'easyview_read_document' ? { type: 'object', properties: {}, additionalProperties: false } : {
      type: 'object', properties: { baseRevision: { type: 'integer', minimum: 0 },
        expectedText: { type: 'string' }, replacement: { type: 'string' }, description: { type: 'string', minLength: 1 } },
      required: ['baseRevision', 'expectedText', 'replacement', 'description'], additionalProperties: false },
    output,
    execute: async (args, exec) => {
      if (!ctx.easyview.ai) throw new Error('EasyView AI service is unavailable');
      return ctx.easyview.ai.tool(toolName, args, exec);
    },
  });
}
