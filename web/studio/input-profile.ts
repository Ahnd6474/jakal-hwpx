// Opt-in local diagnostics. Records timings only, never document text or keys.
export function installInputProfile(targets: Record<string, object | null>) {
  if (!new URLSearchParams(location.search).has('profile-input')) return;
  const stages: Record<string, { count: number; total: number; max: number }> = {};
  const frames: number[] = [];
  let active = false, pending = false, began = 0, inputs = 0;
  for (const [scope, target] of Object.entries(targets)) {
    if (!target) continue;
    const prototype = Object.getPrototypeOf(target);
    for (const name of Object.getOwnPropertyNames(prototype)) {
      const descriptor = Object.getOwnPropertyDescriptor(prototype, name);
      if (name === 'constructor' || typeof descriptor?.value !== 'function') continue;
      const original = descriptor.value;
      Object.defineProperty(target, name, { configurable: true, value: function (...args: unknown[]) {
        if (!active && !name.startsWith('export')) return original.apply(this, args);
        const start = performance.now();
        try { return original.apply(this, args); }
        finally {
          const elapsed = performance.now() - start;
          const value = stages[`${scope}.${name}`] ||= { count: 0, total: 0, max: 0 };
          value.count++; value.total += elapsed; value.max = Math.max(value.max, elapsed);
          if (!active) queueMicrotask(publish);
        }
      } });
    }
  }
  const publish = () => {
    document.documentElement.dataset.jakalInputProfile = JSON.stringify({ inputs, frames, stages });
  };
  function capture(event: Event) {
    if (!(event.target instanceof Element) || !event.target.closest('textarea, [contenteditable="true"]')) return;
    inputs++;
    if (pending) return;
    pending = true; active = true; began = Math.min(event.timeStamp, performance.now());
    requestAnimationFrame(() => requestAnimationFrame(() => {
      frames.push(Math.round((performance.now() - began) * 10) / 10);
      if (frames.length > 128) frames.shift();
      active = pending = false; publish();
    }));
  }
  for (const event of ['keydown', 'beforeinput', 'input', 'compositionstart', 'compositionend']) document.addEventListener(event, capture, true);
  publish();
}
