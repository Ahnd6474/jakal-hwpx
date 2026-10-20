import test from 'node:test';
import assert from 'node:assert/strict';
import { IdleBackup } from '../src/idle-backup.ts';

function fakeClock() {
  let now = 0, serial = 0;
  const timers = new Map(), idle = new Map();
  return {
    clock: { now: () => now, delay: (fn, ms) => { const id = ++serial; timers.set(id, { at: now + ms, fn }); return id; }, cancel: id => timers.delete(id), idle: fn => { const id = ++serial; idle.set(id, fn); return id; }, cancelIdle: id => idle.delete(id) },
    async tick(ms) {
      const end = now + ms;
      while ([...timers.values()].some(t => t.at <= end)) {
        const [id, t] = [...timers].filter(([,t]) => t.at <= end).sort((a,b) => a[1].at - b[1].at)[0];
        now = t.at; timers.delete(id); t.fn();
      }
      now = end; await Promise.resolve();
    },
    async idle() { const jobs = [...idle.values()]; idle.clear(); jobs.forEach(fn => fn()); await Promise.resolve(); await Promise.resolve(); },
  };
}
test('continuous input postpones serialization until a pause and idle; frequent pauses are throttled', async () => {
  const clock = fakeClock(); let calls = 0;
  const queue = new IdleBackup(async () => { calls++; }, clock.clock);
  for (let i = 0; i < 12; i++) { queue.schedule(); await clock.tick(500); await clock.idle(); }
  assert.equal(calls, 0);
  await clock.tick(2000); assert.equal(calls, 0); await clock.idle(); assert.equal(calls, 1);
  queue.schedule(); await clock.tick(2500); await clock.idle(); assert.equal(calls, 1);
  await clock.tick(7500); await clock.idle(); assert.equal(calls, 2);
});
test('typing cancels a queued idle task and never overlaps an export already running', async () => {
  const clock = fakeClock(); let calls = 0, finish, valid;
  const queue = new IdleBackup(async isCurrent => { calls++; valid = isCurrent; await new Promise(resolve => { finish = resolve; }); }, clock.clock);
  queue.schedule(); await clock.tick(2500); queue.schedule(); await clock.idle(); assert.equal(calls, 0);
  await clock.tick(2500); await clock.idle(); assert.equal(calls, 1); assert.equal(valid(), true);
  queue.schedule(); assert.equal(valid(), false);
  await clock.tick(20000); await clock.idle(); assert.equal(calls, 1);
  finish(); await Promise.resolve(); await Promise.resolve();
  await clock.tick(10000); await clock.idle(); assert.equal(calls, 2);
  queue.cancel(); finish();
});
test('manual save cancels pending backup and invalidates an in-flight result; background flush is immediate', async () => {
  const clock = fakeClock(); let calls = 0, finish, valid;
  const queue = new IdleBackup(async isCurrent => { calls++; valid = isCurrent; await new Promise(resolve => { finish = resolve; }); }, clock.clock);
  queue.schedule(); queue.cancel(); await clock.tick(20000); await clock.idle(); assert.equal(calls, 0);
  queue.schedule(); const flush = queue.flush(); assert.equal(calls, 1); assert.equal(valid(), true);
  queue.cancel(); assert.equal(valid(), false); finish(); await flush;
  await clock.tick(20000); await clock.idle(); assert.equal(calls, 1);
});
