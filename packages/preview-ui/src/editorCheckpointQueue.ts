import type { DocumentEditorLifecycle, DraftSnapshot } from './editorLifecycle';

export type CheckpointStatus =
  | { state: 'pending' }
  | { state: 'confirmed'; editorRevision: number; checkpointRevision: number }
  | { state: 'failed'; message: string };

/** Coalesces edits without dropping an in-flight durable ACK; disposal never saves the original. */
export class EditorCheckpointQueue {
  private readonly abort = new AbortController();
  private readonly unsubscribe: () => void;
  private timer: ReturnType<typeof setTimeout> | null = null;
  private running: Promise<void> | null = null;
  private confirmedRevision = -1;
  private failures = 0;
  private pendingSnapshot: DraftSnapshot | null = null;
  private flushing = false;
  private confirmedSnapshot: DraftSnapshot | null = null;

  confirmed(): DraftSnapshot | null {
    return this.confirmedSnapshot ? { ...this.confirmedSnapshot, bytes: this.confirmedSnapshot.bytes.slice() } : null;
  }
  rebaseSaved(diskHash: string): void {
    if (this.confirmedSnapshot) this.confirmedSnapshot = { ...this.confirmedSnapshot, sourceDiskHash: diskHash };
  }

  constructor(private readonly editor: DocumentEditorLifecycle,
    private readonly persist: (snapshot: DraftSnapshot, signal: AbortSignal) => Promise<{ checkpointRevision: number }>,
    private readonly options: { delayMs: number; maxRetries: number; status(value: CheckpointStatus): void }) {
    if (!Number.isSafeInteger(options.delayMs) || options.delayMs < 1 || !Number.isSafeInteger(options.maxRetries) || options.maxRetries < 0) throw new Error('Invalid checkpoint policy');
    this.unsubscribe = editor.subscribe(state => { if (state.dirty && state.editorRevision > this.confirmedRevision) { this.failures = 0; this.schedule(); } });
    if (editor.snapshot().dirty) this.schedule();
  }

  async flush(): Promise<void> {
    this.abort.signal.throwIfAborted();
    this.flushing = true;
    if (this.timer) { clearTimeout(this.timer); this.timer = null; }
    try {
      while (true) {
        try {
          await this.running;
          if (!this.editor.snapshot().dirty || this.editor.snapshot().editorRevision <= this.confirmedRevision) return;
          await this.start();
        } catch (error) {
          // Input can invalidate a background capture while Chat waits for its ACK.
          // Recapture the current revision, as the background scheduler already does.
          if (!(error instanceof Error) || error.message !== 'editor_revision_changed') throw error;
        }
        this.abort.signal.throwIfAborted();
      }
    } finally { this.flushing = false; }
  }

  dispose(): void {
    if (this.abort.signal.aborted) return;
    this.unsubscribe(); this.abort.abort(new Error('checkpoint_owner_disposed'));
    if (this.timer) clearTimeout(this.timer); this.timer = null;
  }

  private schedule(): void {
    if (this.abort.signal.aborted || this.timer || this.running || this.flushing) return;
    this.options.status({ state: 'pending' });
    this.timer = setTimeout(() => {
      this.timer = null;
      void this.start().catch(error => {
        if (this.abort.signal.aborted) return;
        if (error instanceof Error && error.message === 'editor_revision_changed') { this.schedule(); return; }
        this.failures++;
        this.options.status({ state: 'failed', message: error instanceof Error ? error.message : String(error) });
        if (this.failures <= this.options.maxRetries) this.schedule();
      });
    }, this.options.delayMs);
  }

  private start(): Promise<void> {
    if (this.running) return this.running;
    let succeeded = false;
    const work = async () => {
      // Retry the same capture after an uncertain ACK. The transport associates a stable
      // requestId with this object, so a lost response cannot create a conflicting write.
      const snapshot = this.pendingSnapshot ?? await this.editor.capture(this.abort.signal);
      this.pendingSnapshot = snapshot;
      const result = await this.persist(snapshot, this.abort.signal);
      this.abort.signal.throwIfAborted();
      this.confirmedRevision = snapshot.editorRevision; this.failures = 0; this.pendingSnapshot = null;
      this.confirmedSnapshot = snapshot;
      succeeded = true;
      this.options.status({ state: 'confirmed', editorRevision: snapshot.editorRevision, checkpointRevision: result.checkpointRevision });
    };
    this.running = work().finally(() => {
      this.running = null;
      if (succeeded && !this.abort.signal.aborted && this.editor.snapshot().dirty && this.editor.snapshot().editorRevision > this.confirmedRevision) this.schedule();
    });
    return this.running;
  }
}
