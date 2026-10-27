/** A paused hidden reader waits for one signal, rather than polling. */
export class ActivityGate {
  private enabled = true;
  private waiting = new Set<() => void>();
  get active() { return this.enabled; }
  setActive(active: boolean) {
    this.enabled = active;
    if (active) { for (const resume of this.waiting) resume(); this.waiting.clear(); }
  }
  wait(): Promise<void> {
    return this.enabled ? Promise.resolve() : new Promise(resolve => this.waiting.add(resolve));
  }
}

export type IndexedText = { page: number; text: string; limited: boolean };
/** Bounded IPC batches avoid copying the growing search index for every page. */
export class IndexBatch {
  private pending: IndexedText[] = [];
  private characters = 0;
  private emit: (entries: IndexedText[]) => void;
  constructor(emit: (entries: IndexedText[]) => void) { this.emit = emit; }
  add(entry: IndexedText) {
    this.pending.push(entry); this.characters += entry.text.length;
    if (this.pending.length >= 8 || this.characters >= 100_000) this.flush();
  }
  flush() {
    if (!this.pending.length) return;
    const ready = this.pending; this.pending = []; this.characters = 0; this.emit(ready);
  }
}
