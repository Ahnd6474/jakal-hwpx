import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs';
import { HwpDocument, initSync } from '@rhwp/core';
import { insertBlock } from '../studio/block-insertion.ts';

initSync({ module: fs.readFileSync(new URL('../node_modules/@rhwp/core/rhwp_bg.wasm', import.meta.url)) });
test('multi-line block insertion preserves real math controls and restores a snapshot in one step', () => {
  const source = fs.readFileSync(new URL('../../examples/samples/hwpx/math_template_2_formula_fixed.hwpx', import.meta.url));
  const doc = new HwpDocument(source);
  try {
    const original = doc.exportHwpx();
    const before = doc.saveSnapshot();
    const position = insertBlock(doc, { sectionIndex: 0, paragraphIndex: 0, charOffset: 0 }, '\n① 한글😀\n② x < 3\n');
    assert.equal(position.paragraphIndex, 3);
    const exported = doc.exportHwpxWithReport();
    try {
      assert.equal(JSON.parse(exported.contentLoss()).count, 0);
      const reloaded = new HwpDocument(exported.takeBytes());
      try { assert.ok(reloaded.pageCount() >= 1); } finally { reloaded.free(); }
    } finally { exported.free(); }
    const after = doc.saveSnapshot();
    doc.restoreSnapshot(before);
    assert.deepEqual(doc.exportHwpx(), original);
    doc.restoreSnapshot(after);
    assert.notDeepEqual(doc.exportHwpx(), original);
    doc.discardSnapshot(before); doc.discardSnapshot(after);
  } finally { doc.free(); }
});
test('failed native edits throw for the surrounding snapshot transaction to roll back', () => {
  const position = { sectionIndex: 0, paragraphIndex: 0, charOffset: 0 };
  const reject = { insertText: () => '{"ok":false}', splitParagraph: () => '{"ok":false}' };
  assert.throws(() => insertBlock(reject, position, 'text'));
  assert.throws(() => insertBlock(reject, position, '\ntext'));
  assert.throws(() => insertBlock(reject, position, '\u0000'));
});
