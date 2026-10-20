type Clock = {
  now: () => number;
  delay: (run: () => void, ms: number) => number;
  cancel: (id: number) => void;
  idle: (run: () => void) => number;
  cancelIdle: (id: number) => void;
};
const browserClock: Clock = {
  now: Date.now,
  delay: (run, ms) => window.setTimeout(run, ms),
  cancel: id => window.clearTimeout(id),
  idle: run => window.requestIdleCallback ? window.requestIdleCallback(run, { timeout: 8000 }) : window.setTimeout(run, 100),
  cancelIdle: id => window.cancelIdleCallback ? window.cancelIdleCallback(id) : window.clearTimeout(id),
};

/** Serialize drafts after an input pause, at idle, without overlapping exports. */
export class IdleBackup {
  private version = 0;
  private pending = false;
  private running = false;
  private lastInput = 0;
  private lastRun = -Infinity;
  private timer: number | undefined;
  private idle: number | undefined;
  private task: (isCurrent: () => boolean) => Promise<void>;
  private clock: Clock;
  constructor(task: (isCurrent: () => boolean) => Promise<void>, clock: Clock = browserClock) { this.task = task; this.clock = clock; }
  schedule() {
    this.version++; this.pending = true; this.lastInput = this.clock.now(); this.arm();
  }
  cancel() {
    this.version++; this.pending = false; this.clear();
  }
  private clear() {
    if (this.timer !== undefined) this.clock.cancel(this.timer);
    if (this.idle !== undefined) this.clock.cancelIdle(this.idle);
    this.timer = this.idle = undefined;
  }
  private arm() {
    this.clear();
    if (this.running || !this.pending) return;
    const wait = Math.max(0, this.lastInput + 2500 - this.clock.now(), this.lastRun + 10_000 - this.clock.now());
    this.timer = this.clock.delay(() => {
      this.timer = undefined;
      this.idle = this.clock.idle(() => { this.idle = undefined; void this.run(); });
    }, wait);
  }
  async flush() {
    this.clear();
    if (this.running || !this.pending) return;
    await this.run();
  }
  private async run() {
    if (this.running || !this.pending) return;
    this.pending = false; this.running = true;
    const at = this.version;
    try { await this.task(() => at === this.version); }
    finally { this.running = false; this.lastRun = this.clock.now(); this.arm(); }
  }
}
