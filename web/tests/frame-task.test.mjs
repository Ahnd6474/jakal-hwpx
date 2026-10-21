import assert from 'node:assert/strict';
import test from 'node:test';
import { FrameTask } from '../studio/frame-task.ts';
function fixture() {
  let serial = 0;
  const frames = new Map();
  const task = new FrameTask({ request: run => { const id = serial++; frames.set(id, run); return id; }, cancel: id => frames.delete(id) });
  return { task, frames, flush() { const jobs = [...frames.values()]; frames.clear(); jobs.forEach(run => run()); } };
}
test('an input burst paints only the latest document state once per frame', () => {
  const f = fixture(), painted = [];
  for (let n = 0; n < 20; n++) f.task.schedule(() => painted.push(n));
  assert.equal(f.frames.size, 1);
  assert.deepEqual(painted, []);
  f.flush();
  assert.deepEqual(painted, [19]);
  f.task.schedule(() => painted.push(20)); f.flush();
  assert.deepEqual(painted, [19, 20]);
});
test('document reset cancels pending work and does not paint a stale document', () => {
  const f = fixture(), painted = [];
  f.task.schedule(() => painted.push('old'));
  f.task.cancel(); f.flush(); assert.deepEqual(painted, []);
  f.task.schedule(() => painted.push('new')); f.flush(); assert.deepEqual(painted, ['new']);
});
test('a callback can schedule the next frame and an error does not lock the queue', () => {
  const f = fixture(), painted = [];
  f.task.schedule(() => { painted.push(1); f.task.schedule(() => painted.push(2)); });
  f.flush(); assert.deepEqual(painted, [1]); f.flush(); assert.deepEqual(painted, [1, 2]);
  f.task.schedule(() => { throw Error('paint failed'); }); assert.throws(() => f.flush());
  f.task.schedule(() => painted.push(3)); f.flush(); assert.deepEqual(painted, [1, 2, 3]);
});
