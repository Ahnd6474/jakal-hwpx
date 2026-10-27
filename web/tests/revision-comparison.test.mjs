import assert from 'node:assert/strict';
import test from 'node:test';
import { compareRevisions } from '../src/revision-comparison.ts';

const document = (texts, tables = []) => ({ sections: [{ index: 0, paragraphs: texts.map((text, index) => ({ id: `p${index}`, section: 0, index, text, editable: true })), tables }] });
const table = (text, span = 1) => ({ id: 't0', section: 0, index: 0, rows: span, columns: 1, cells: [{ row: 0, column: 0, row_span: span, col_span: 1, text, editable: true }] });

test('unchanged snapshots have no text or table changes', () => {
  assert.equal(compareRevisions(document(['본문'], [table('표')]), document(['본문'], [table('표')])).total, 0);
});
test('inserted paragraphs do not mark following paragraphs as edited', () => {
  const result = compareRevisions(document(['첫째', '둘째', '셋째']), document(['새 문단', '첫째', '둘째', '중간 삽입', '셋째']));
  assert.equal(result.total, 2);
  assert.deepEqual(result.items.map(i => [i.beforeCount, i.after]), [[0, '새 문단'], [0, '중간 삽입']]);
});
test('anchors separate independent replacements and deletions', () => {
  const result = compareRevisions(document(['A', 'B', 'C', 'D', 'E']), document(['변경 A', 'B', 'C', 'E']));
  assert.equal(result.total, 2);
  assert.equal(result.items[0].before, 'A');
  assert.equal(result.items[0].after, '변경 A');
  assert.equal(result.items[1].before, 'D');
  assert.equal(result.items[1].afterCount, 0);
});
test('repeated passages and blank paragraphs retain additions and deletions', () => {
  const result = compareRevisions(document(['', '반복', '반복', '끝']), document(['', '반복', '', '반복', '끝']));
  assert.equal(result.total, 1);
  assert.equal(result.items[0].afterCount, 1);
  assert.equal(result.items[0].after, '');
});
test('table contents and merge geometry are both compared', () => {
  assert.equal(compareRevisions(document([], [table('이전')]), document([], [table('현재')])).tables, 1);
  const geometry = compareRevisions(document([], [table('같은 내용')]), document([], [table('같은 내용', 2)]));
  assert.equal(geometry.tables, 1);
  assert.match(geometry.items[0].after, /병합/);
});
test('section additions and table deletions appear even without paragraphs', () => {
  const before = document([], [table('표')]);
  const after = { sections: [{ index: 1, paragraphs: [{ index: 0, text: '새 섹션' }], tables: [] }] };
  const result = compareRevisions(before, after);
  assert.equal(result.tables, 1);
  assert.equal(result.paragraphs, 1);
});
test('large histories bound rendered comparisons and long text', () => {
  const old = [], current = [];
  for (let n = 0; n < 700; n++) { old.push(`이전 ${n}`, `기준 ${n}`); current.push(`현재 ${n}`, `기준 ${n}`); }
  const result = compareRevisions(document(old), document(current));
  assert.equal(result.total, 700);
  assert.equal(result.items.length, 500);
  assert.equal(result.omitted, 200);
  const long = compareRevisions(document(['x'.repeat(20000)]), document(['y'.repeat(20000)]));
  assert.equal(long.items[0].before.length, 16000);
  assert.equal(long.items[0].clipped, true);
});
