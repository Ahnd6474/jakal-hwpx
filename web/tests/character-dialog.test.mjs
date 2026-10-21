import assert from 'node:assert/strict';
import test from 'node:test';
import { applyDialogCharacterProps } from '../studio/character-dialog.ts';

function fixture() {
  const ranges = [], pending = [];
  return { ranges, pending, input: { applyCharPropsToRange: (...args) => ranges.push(args) }, events: { emit: (...args) => pending.push(args) } };
}
const compare = (a, b) => a.paragraph - b.paragraph || a.offset - b.offset;
test('character dialog at a caret stages the next typing format without changing nearby text', () => {
  for (const selection of [null, { start: { paragraph: 2, offset: 3 }, end: { paragraph: 2, offset: 3 } }]) {
    const f = fixture(), properties = { bold: true, height: 1400 };
    applyDialogCharacterProps(f.input, f.events, selection, properties, compare);
    assert.deepEqual(f.ranges, []);
    assert.deepEqual(f.pending, [['format-char', properties]]);
  }
});
test('character dialog applies to the range captured before focus moved into the dialog', () => {
  const f = fixture(), start = { paragraph: 2, offset: 3 }, end = { paragraph: 3, offset: 1 }, properties = { italic: true };
  applyDialogCharacterProps(f.input, f.events, { start, end }, properties, compare);
  assert.deepEqual(f.ranges, [[start, end, properties]]);
  assert.deepEqual(f.pending, []);
});
