// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { ByteEditorLifecycle } from './editorLifecycle';
import { EditorCheckpointQueue } from './editorCheckpointQueue';

function fixture() {
  return new ByteEditorLifecycle({ documentId: 'document', format: 'docx', sourceDiskHash: 'base',
    root: () => null, flushInput() {}, async serialize() { return new Uint8Array([1]); }, async restore() {} });
}
afterEach(() => vi.useRealTimers());
describe('durable checkpoint queue', () => {
  for (const background of [false, true]) it(`flush recaptures when Word input changes during serialization (background: ${background})`, async () => {
    vi.useFakeTimers();
    let release!: () => void, captures = 0;
    const editor = new ByteEditorLifecycle({ documentId: 'document', format: 'docx', sourceDiskHash: 'base',
      root: () => null, flushInput() {}, async restore() {},
      async serialize() {
        captures++;
        if (captures === 1) await new Promise<void>(resolve => { release = resolve; });
        return new Uint8Array([editor.snapshot().editorRevision]);
      } });
    const persisted: number[] = [];
    const queue = new EditorCheckpointQueue(editor, async snapshot => {
      persisted.push(snapshot.editorRevision); return { checkpointRevision: persisted.length };
    }, { delayMs: 10, maxRetries: 0, status() {} });
    editor.changed();
    if (background) await vi.advanceTimersByTimeAsync(10);
    const flush = queue.flush();
    // Observe the rejection on the old implementation while controlling the capture.
    void flush.catch(() => {});
    await vi.advanceTimersByTimeAsync(0);
    editor.changed(); release();
    await flush;
    expect(captures).toBe(2); expect(persisted).toEqual([2]);
    expect(queue.confirmed()?.bytes).toEqual(new Uint8Array([2]));
    queue.dispose(); editor.dispose();
  });
  it('waits for an in-flight ACK before coalescing newer input', async () => {
    vi.useFakeTimers(); const editor = fixture(); const revisions: number[] = [];
    let acknowledge!: () => void;
    const queue = new EditorCheckpointQueue(editor, async snapshot => {
      revisions.push(snapshot.editorRevision);
      if (revisions.length === 1) await new Promise<void>(resolve => { acknowledge = resolve; });
      return { checkpointRevision: revisions.length };
    }, { delayMs: 10, maxRetries: 0, status() {} });
    editor.changed(); await vi.advanceTimersByTimeAsync(10);
    editor.changed(); editor.changed(); await vi.advanceTimersByTimeAsync(50);
    expect(revisions).toEqual([1]); acknowledge(); await vi.advanceTimersByTimeAsync(10);
    expect(revisions).toEqual([1, 3]); expect(editor.snapshot().dirty).toBe(true);
    queue.dispose(); editor.dispose();
  });
  it('honors zero retries, then explicitly retries the identical snapshot after a lost ACK', async () => {
    vi.useFakeTimers(); const editor = fixture(); const captures: unknown[] = [];
    const queue = new EditorCheckpointQueue(editor, async snapshot => {
      captures.push(snapshot); if (captures.length === 1) throw new Error('response_lost');
      return { checkpointRevision: 1 };
    }, { delayMs: 10, maxRetries: 0, status() {} });
    editor.changed(); await vi.advanceTimersByTimeAsync(100);
    expect(captures).toHaveLength(1); await queue.flush();
    expect(captures).toHaveLength(2); expect(captures[0]).toBe(captures[1]);
    queue.dispose(); editor.dispose();
  });
  it('does not issue a late ACK or another capture after external disposal', async () => {
    vi.useFakeTimers(); const editor = fixture(); const states: string[] = [];
    let acknowledge!: () => void; let calls = 0;
    const queue = new EditorCheckpointQueue(editor, async () => {
      calls++; await new Promise<void>(resolve => { acknowledge = resolve; }); return { checkpointRevision: 1 };
    }, { delayMs: 10, maxRetries: 3, status(value) { states.push(value.state); } });
    editor.changed(); await vi.advanceTimersByTimeAsync(10); editor.changed(); queue.dispose(); editor.dispose();
    acknowledge(); await vi.advanceTimersByTimeAsync(100);
    expect(calls).toBe(1); expect(states).not.toContain('confirmed');
  });
});
