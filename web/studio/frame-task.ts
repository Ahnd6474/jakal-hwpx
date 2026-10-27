type Clock = { request: (run: () => void) => number; cancel: (id: number) => void };

/** Mutations remain synchronous; only their derived display work is coalesced. */
export class FrameTask {
  private frame: number | null = null;
  private latest: (() => void) | null = null;
  private clock: Clock;
  constructor(clock: Clock = { request: run => requestAnimationFrame(run), cancel: id => cancelAnimationFrame(id) }) { this.clock = clock; }
  schedule(run: () => void) {
    this.latest = run;
    if (this.frame !== null) return;
    this.frame = this.clock.request(() => {
      this.frame = null;
      const latest = this.latest; this.latest = null;
      latest?.();
    });
  }
  cancel() {
    if (this.frame !== null) this.clock.cancel(this.frame);
    this.frame = null; this.latest = null;
  }
}
