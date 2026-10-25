import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs';
import { HwpDocument, initSync } from '@rhwp/core';
import { insertNestedTable, parentTableCell, editNestedTable } from '../studio/nested-table.ts';

initSync({ module: fs.readFileSync(new URL('../node_modules/@rhwp/core/rhwp_bg.wasm', import.meta.url)) });
const source = fs.readFileSync(new URL('../public/sample.hwpx', import.meta.url));
function fixture(t) {
  const doc = new HwpDocument(source);
  t.after(() => doc.free());
  const paragraph = doc.getParagraphCount(0) - 1;
  const table = JSON.parse(doc.createTable(0, paragraph, doc.getParagraphLength(0, paragraph), 2, 2));
  assert.equal(table.ok, true);
  const position = { sectionIndex: 0, paragraphIndex: 0, charOffset: 0,
    parentParaIndex: table.paraIdx, controlIndex: table.controlIdx, cellIndex: 0, cellParaIndex: 0 };
  return { doc, position };
}
function dimensions(doc, position) {
  return JSON.parse(doc.getTableDimensionsByPath(position.sectionIndex, position.parentParaIndex, JSON.stringify(position.cellPath)));
}
function exportBytes(doc) {
  const artifact = doc.exportHwpxWithReport();
  try {
    assert.equal(JSON.parse(artifact.contentLoss()).count, 0);
    return artifact.takeBytes();
  } finally { artifact.free(); }
}

test('nested tables fit the cell, accept text, persist across export, undo and redo', t => {
  const { doc, position } = fixture(t);
  doc.insertTextInCell(0, position.parentParaIndex, position.controlIndex, 0, 0, 0, '앞😀뒤');
  position.charOffset = 2;
  const count = doc.getParagraphCount(0), before = doc.saveSnapshot(), beforeBytes = exportBytes(doc);
  const cursor = insertNestedTable(doc, position, 2, 3);
  assert.equal(doc.getTextInCell(0, position.parentParaIndex, position.controlIndex, 0, 0, 0, 20), '앞😀');
  assert.equal(doc.getTextInCell(0, position.parentParaIndex, position.controlIndex, 0, 2, 0, 20), '뒤');
  assert.deepEqual(dimensions(doc, cursor), { rowCount: 2, colCount: 3, cellCount: 6 });
  assert.equal(doc.getParagraphCount(0), count, 'scratch paragraphs are removed');
  assert.equal(doc.hasInternalClipboard(), false);
  doc.insertTextInCellByPath(0, cursor.parentParaIndex, JSON.stringify(cursor.cellPath), 0, '안쪽 셀😀');
  const third = insertNestedTable(doc, cursor, 1, 2);
  assert.equal(third.cellPath.length, 3);
  assert.equal(dimensions(doc, third).colCount, 2);
  const bytes = exportBytes(doc), after = doc.saveSnapshot();
  const reloaded = new HwpDocument(bytes);
  try {
    assert.equal(dimensions(reloaded, third).colCount, 2);
    const inner = JSON.parse(reloaded.getCellPropertiesByPath(0, cursor.parentParaIndex, JSON.stringify(cursor.cellPath), 0));
    const outer = JSON.parse(reloaded.getCellProperties(0, position.parentParaIndex, position.controlIndex, 0));
    assert.ok(inner.width * 3 <= outer.width - outer.paddingLeft - outer.paddingRight);
    assert.ok(reloaded.getControls().includes('tbl'));
  } finally { reloaded.free(); }
  doc.restoreSnapshot(before);
  assert.deepEqual(exportBytes(doc), beforeBytes);
  doc.restoreSnapshot(after);
  assert.deepEqual(exportBytes(doc), bytes);
  doc.discardSnapshot(before); doc.discardSnapshot(after);
});

test('inserting a table preserves a rich internal copy buffer', t => {
  const { doc, position } = fixture(t);
  doc.copyControl(0, position.parentParaIndex, '', position.controlIndex);
  const text = doc.getClipboardText();
  insertNestedTable(doc, position, 2, 2);
  assert.equal(doc.getClipboardText(), text);
  assert.equal(doc.clipboardHasControl(), true);
  doc.pasteInternalInCell(0, position.parentParaIndex, position.controlIndex, 1, 0, 0);
  const path = [{ controlIndex: position.controlIndex, cellIndex: 1, cellParaIndex: 0 }, { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }];
  assert.equal(JSON.parse(doc.getTableDimensionsByPath(0, position.parentParaIndex, JSON.stringify(path))).cellCount, 4);
  exportBytes(doc);
});
test('row and column edits apply to the inner table and preserve the outer grid', t => {
  const { doc, position } = fixture(t);
  let inner = insertNestedTable(doc, position, 2, 2);
  doc.insertTextInCellByPath(0, inner.parentParaIndex, JSON.stringify(inner.cellPath), 0, '행열 유지😀');
  doc.copyControl(0, position.parentParaIndex, '', position.controlIndex);
  const clipboard = doc.getClipboardText(), originalCount = doc.getParagraphCount(0);
  for (const operation of ['insertRowAbove', 'insertRowBelow', 'insertColLeft', 'insertColRight', 'deleteRow', 'deleteCol']) {
    inner = editNestedTable(doc, inner, operation);
  }
  assert.equal(dimensions(doc, inner).rowCount, 3);
  assert.equal(dimensions(doc, inner).colCount, 3);
  const moved = inner.cellPath.map(entry => ({ ...entry })); moved.at(-1).cellIndex = 4;
  assert.equal(doc.getTextInCellByPath(0, inner.parentParaIndex, JSON.stringify(moved), 0, 50), '행열 유지😀');
  assert.equal(doc.getClipboardText(), clipboard);
  assert.equal(doc.getParagraphCount(0), originalCount);
  assert.deepEqual(JSON.parse(doc.getTableDimensions(0, position.parentParaIndex, position.controlIndex)), { rowCount: 2, colCount: 2, cellCount: 4 });
  exportBytes(doc);
});

test('invalid sizes do not mutate the document and leaving a nested table keeps its parent path', t => {
  const { doc, position } = fixture(t);
  const before = exportBytes(doc);
  for (const [rows, cols] of [[0, 2], [2, -1], [NaN, 2], [1.5, 2], [100, 100]]) {
    assert.throws(() => insertNestedTable(doc, position, rows, cols));
  }
  assert.deepEqual(exportBytes(doc), before);
  const inner = insertNestedTable(doc, position, 1, 1);
  const parent = parentTableCell(inner);
  assert.equal(parent.cellPath.length, 1);
  assert.equal(parent.parentParaIndex, position.parentParaIndex);
  assert.equal(parent.cellPath[0].cellParaIndex, 1);
  assert.throws(() => parentTableCell(position));
});

test('deleting the last inner row is rejected without changing either table', t => {
  const { doc, position } = fixture(t);
  const inner = insertNestedTable(doc, position, 1, 1);
  assert.throws(() => editNestedTable(doc, inner, 'deleteRow'));
  assert.equal(dimensions(doc, inner).cellCount, 1);
  assert.equal(JSON.parse(doc.getTableDimensions(0, position.parentParaIndex, position.controlIndex)).cellCount, 4);
  exportBytes(doc);
});

test('nested edits in the second section resolve the right parent table', t => {
  const doc = new HwpDocument(fs.readFileSync(new URL('../../examples/samples/hwpx/math_template_2_formula_fixed.hwpx', import.meta.url)));
  t.after(() => doc.free());
  const paragraph = doc.getParagraphCount(1) - 1;
  const table = JSON.parse(doc.createTable(1, paragraph, doc.getParagraphLength(1, paragraph), 2, 2));
  const position = { sectionIndex: 1, paragraphIndex: 0, charOffset: 0, parentParaIndex: table.paraIdx,
    controlIndex: table.controlIdx, cellIndex: 0, cellParaIndex: 0 };
  let inner = insertNestedTable(doc, position, 2, 2);
  inner = editNestedTable(doc, inner, 'insertRowBelow');
  assert.equal(dimensions(doc, inner).rowCount, 3);
  assert.equal(JSON.parse(doc.getTableDimensions(1, table.paraIdx, table.controlIdx)).rowCount, 2);
  exportBytes(doc);
});

test('multiline text copy buffers survive insertion and failures can be rolled back without leaked paragraphs', t => {
  const { doc, position } = fixture(t);
  const last = doc.getParagraphCount(0) - 1;
  doc.insertText(0, last, 0, '복사😀');
  const split = JSON.parse(doc.splitParagraph(0, last, doc.getParagraphLength(0, last)));
  doc.insertText(0, split.paraIdx, 0, '둘째 줄');
  doc.copySelection(0, last, 0, split.paraIdx, doc.getParagraphLength(0, split.paraIdx));
  const copied = doc.getClipboardText();
  insertNestedTable(doc, position, 1, 2);
  assert.equal(doc.getClipboardText(), copied);
  const before = exportBytes(doc), snapshot = doc.saveSnapshot(), count = doc.getParagraphCount(0);
  const failing = new Proxy(doc, { get(target, name) {
    if (name === 'pasteInternalInCellByPath') return () => '{"ok":false}';
    const member = target[name]; return typeof member === 'function' ? member.bind(target) : member;
  } });
  assert.throws(() => insertNestedTable(failing, position, 1, 1));
  assert.equal(doc.getClipboardText(), copied);
  assert.equal(doc.getParagraphCount(0), count);
  doc.restoreSnapshot(snapshot); doc.discardSnapshot(snapshot);
  assert.deepEqual(exportBytes(doc), before);
});
