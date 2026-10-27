import assert from 'node:assert/strict';
import test from 'node:test';
import { ActivityGate, IndexBatch } from '../src/viewer-index.ts';

test('hidden reader indexing sleeps until explicitly resumed', async () => {
  const gate = new ActivityGate(); await gate.wait();
  gate.setActive(false); const completed = [];
  const first = gate.wait().then(() => completed.push(1));
  const second = gate.wait().then(() => completed.push(2));
  await Promise.resolve(); assert.deepEqual(completed, []);
  gate.setActive(true); await Promise.all([first, second]);
  assert.deepEqual(completed, [1, 2]);
  gate.setActive(false); assert.equal(gate.active, false);
  let resumed = false; const next = gate.wait().then(() => { resumed = true; });
  await Promise.resolve(); assert.equal(resumed, false);
  gate.setActive(true); await next;
});
test('500 pages arrive in bounded batches with identical text and limited flags', () => {
  const sent = [], batch = new IndexBatch(entries => sent.push(entries));
  const input = Array.from({ length: 500 }, (_, page) => ({ page, text: `문서 ${page} 😀`, limited: page >= 490 }));
  input.forEach(entry => batch.add(entry)); batch.flush(); batch.flush();
  assert.equal(sent.length, 63);
  assert.ok(sent.every(entries => entries.length <= 8));
  assert.deepEqual(sent.flat(), input);
});
test('large search text flushes early and pause flush preserves partial progress', () => {
  const sent = [], batch = new IndexBatch(entries => sent.push(entries));
  batch.add({ page: 0, text: 'x'.repeat(100_000), limited: false });
  assert.equal(sent.length, 1);
  batch.add({ page: 1, text: 'last page', limited: true });
  batch.flush(); assert.equal(sent.length, 2);
  assert.equal(sent[1][0].page, 1);
});
