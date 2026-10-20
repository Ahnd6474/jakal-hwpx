import assert from 'node:assert/strict';
import test from 'node:test';
import { draftReducer } from '../src/draft-history.ts';

const initial = () => ({ past: [], present: {}, future: [] });
const change = (text, key = 's0-p0') => ({ key, original: '원문', operation: { type: 'paragraph', section: 0, index: 0, text } });
const edit = (state, text) => draftReducer(state, { type: 'change', ...change(text) });

test('unchanged inputs and empty batches preserve the redo branch', () => {
  const first = edit(initial(), '수정 1');
  const second = edit(first, '수정 2');
  const undone = draftReducer(second, { type: 'undo' });
  assert.equal(edit(undone, '수정 1'), undone);
  assert.equal(draftReducer(undone, { type: 'batch', changes: [] }), undone);
  assert.equal(draftReducer(undone, { type: 'redo' }).present['s0-p0'].text, '수정 2');
});

test('returning to original removes the draft and remains undoable', () => {
  const empty = initial();
  assert.equal(edit(empty, '원문'), empty);
  const changed = edit(empty, '수정');
  const reverted = edit(changed, '원문');
  assert.deepEqual(reverted.present, {});
  assert.equal(draftReducer(reverted, { type: 'undo' }).present, changed.present);
  assert.equal(changed.present['s0-p0'].text, '수정');
});

test('a replacement batch is one undo step and cancelled batches are no-ops', () => {
  const before = edit(initial(), '첫 수정');
  const cancelled = draftReducer(before, { type: 'batch', changes: [change('중간'), change('첫 수정')] });
  assert.equal(cancelled, before);
  const after = draftReducer(before, { type: 'batch', changes: [change('본문'), change('둘째', 's0-p1')] });
  assert.equal(after.past.length, before.past.length + 1);
  assert.deepEqual(Object.keys(after.present), ['s0-p0', 's0-p1']);
  assert.equal(draftReducer(after, { type: 'undo' }).present, before.present);
  assert.deepEqual(Object.keys(before.present), ['s0-p0']);
});

test('a real edit after undo clears redo, while history keeps the most recent 60 edits', () => {
  let state = initial();
  for (let n = 1; n <= 70; n++) state = edit(state, String(n));
  assert.equal(state.past.length, 60);
  for (let n = 0; n < 60; n++) state = draftReducer(state, { type: 'undo' });
  assert.equal(state.present['s0-p0'].text, '10');
  assert.equal(draftReducer(state, { type: 'undo' }), state);
  state = edit(state, '새 분기');
  assert.equal(state.future.length, 0);
  assert.equal(draftReducer(state, { type: 'redo' }), state);
});

test('cell drafts compare target coordinates as well as text', () => {
  const operation = { type: 'cell', section: 0, table: 0, row: 1, column: 2, text: '값' };
  const state = draftReducer(initial(), { type: 'change', key: 'cell', original: '', operation });
  assert.equal(draftReducer(state, { type: 'change', key: 'cell', original: '', operation: { ...operation } }), state);
  const moved = draftReducer(state, { type: 'change', key: 'cell', original: '', operation: { ...operation, column: 3 } });
  assert.equal(moved.present.cell.column, 3);
  assert.equal(state.present.cell.column, 2);
});

test('reset discards the previous document’s drafts and both history branches', () => {
  const state = draftReducer(edit(initial(), '수정'), { type: 'undo' });
  assert.deepEqual(draftReducer(state, { type: 'reset' }), initial());
});
