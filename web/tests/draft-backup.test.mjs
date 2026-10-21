import assert from 'node:assert/strict';
import test from 'node:test';
import { backupText, clearDraftBackup, decodeBackup, readDraftBackup, writeDraftBackup } from '../src/draft-backup.ts';
import { textTargetIndex } from '../src/editing.ts';

const now = Date.now();
const doc = { id: 'test-document', revision: 3, expires: now / 1000 + 86400, sections: [{ index: 0,
  paragraphs: [{ id: 's0-p0', section: 0, index: 0, text: '원문', editable: true }, { id: 's0-p1', section: 0, index: 1, text: '복합 개체', editable: false }],
  tables: [{ id: 's0-t0', section: 0, index: 0, cells: [{ row: 1, column: 2, text: '셀', editable: true }] }],
}] };
const drafts = { 's0-p0': { type: 'paragraph', section: 0, index: 0, text: '한글 & <안내> 😀\n둘째 줄' },
  's0-t0-1-2': { type: 'cell', section: 0, table: 0, row: 1, column: 2, text: '' } };
const backup = () => ({ version: 1, documentId: doc.id, revision: doc.revision, expires: doc.expires, updatedAt: now,
  changes: [{ id: 's0-p0', original: '원문', operation: drafts['s0-p0'] }, { id: 's0-t0-1-2', original: '셀', operation: drafts['s0-t0-1-2'] }] });

test('empty drafts do not traverse document content and sparse changes stay canonical', t => {
  const saved = new Map(); storage(t, { setItem: (key, value) => saved.set(key, value), removeItem: key => saved.delete(key) });
  assert.equal(writeDraftBackup({ id: 'empty', get sections() { throw new Error('unneeded content access'); } }, {}), true);
  const reversed = { 'unknown': { type: 'paragraph', section: 0, index: 0, text: 'ignored' }, 's0-t0-1-2': drafts['s0-t0-1-2'], 's0-p0': { ...drafts['s0-p0'], index: 999 } };
  assert.equal(writeDraftBackup(doc, reversed), true);
  const record = JSON.parse(saved.get('jakal-draft:' + doc.id));
  assert.deepEqual(record.changes.map(item => item.id), ['s0-p0', 's0-t0-1-2']);
  assert.equal(record.changes[0].operation.index, 0);
  assert.equal(record.changes[1].operation.text, '');
  assert.equal(writeDraftBackup(doc, { ...drafts, 's0-p0': { ...drafts['s0-p0'], text: null } }), true);
  assert.deepEqual(JSON.parse(saved.get('jakal-draft:' + doc.id)).changes.map(item => item.id), ['s0-t0-1-2']);
});
test('target index follows immutable content rather than document id or settings identity', () => {
  const first = textTargetIndex(doc);
  assert.equal(textTargetIndex({ ...doc, settings: { instructions: 'changed' } }), first);
  const next = { ...doc, sections: [{ ...doc.sections[0], paragraphs: [{ ...doc.sections[0].paragraphs[0], text: 'new source' }] }] };
  assert.notEqual(textTargetIndex(next), first);
  assert.equal(textTargetIndex(next).get('s0-p0').original, 'new source');
  assert.equal(first.get('s0-p0').original, '원문');
});

function storage(t, value) {
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'sessionStorage');
  Object.defineProperty(globalThis, 'sessionStorage', { configurable: true, value });
  t.after(() => { if (previous) Object.defineProperty(globalThis, 'sessionStorage', previous); else delete globalThis.sessionStorage; });
}

test('same-version recovery preserves Unicode, line breaks and a cleared cell', () => {
  const result = decodeBackup(JSON.stringify(backup()), doc, now);
  assert.equal(result.kind, 'restore');
  assert.deepEqual(result.drafts, drafts);
  assert.ok(backupText(result.backup).includes(drafts['s0-p0'].text));
});

test('different server versions or changed source text are kept as conflicts, never applied', () => {
  for (const changed of [{ ...doc, revision: 4 }, { ...doc, sections: [{ ...doc.sections[0], paragraphs: [{ ...doc.sections[0].paragraphs[0], text: '다른 원문' }] }] }]) {
    const result = decodeBackup(JSON.stringify(backup()), changed, now);
    assert.equal(result.kind, 'conflict');
    assert.equal(result.drafts, undefined);
    assert.equal(result.backup.changes[0].operation.text, drafts['s0-p0'].text);
  }
});

test('removed, read-only, or mismatched targets cannot be restored', () => {
  for (const mutate of [value => { value.changes[0].id = 'missing'; }, value => { value.changes[0].id = 's0-p1'; }, value => { value.changes[0].operation = { ...value.changes[0].operation, index: 999 }; }]) {
    const value = structuredClone(backup()); mutate(value);
    assert.equal(decodeBackup(JSON.stringify(value), doc, now).kind, 'conflict');
  }
});

test('expired, wrong-document, duplicate and malformed records are ignored', () => {
  const records = [null, '', '{', 'x'.repeat(2000001), JSON.stringify({ ...backup(), documentId: 'another-document' }), JSON.stringify({ ...backup(), expires: now / 1000 - 1 }), JSON.stringify({ ...backup(), changes: [backup().changes[0], backup().changes[0]] }), JSON.stringify({ ...backup(), changes: [null] })];
  for (const raw of records) assert.equal(decodeBackup(raw, doc, now).kind, 'none');
});

test('restored operations use canonical target fields, ignoring extra stored properties', () => {
  const value = backup();
  value.changes[0].operation = { ...value.changes[0].operation, unexpected: 'value' };
  assert.deepEqual(decodeBackup(JSON.stringify(value), doc, now).drafts['s0-p0'], drafts['s0-p0']);
});

test('writing and clearing drafts is independent for each document', t => {
  const entries = new Map();
  storage(t, { getItem: key => entries.get(key) ?? null, setItem: (key, value) => entries.set(key, value), removeItem: key => entries.delete(key) });
  assert.ok(writeDraftBackup(doc, drafts));
  assert.deepEqual(readDraftBackup(doc).drafts, drafts);
  assert.ok(writeDraftBackup({ ...doc, id: 'second-document' }, drafts));
  assert.ok(writeDraftBackup(doc, {}));
  assert.equal(readDraftBackup(doc).kind, 'none');
  assert.equal(entries.size, 1);
  assert.ok(clearDraftBackup('second-document'));
  assert.equal(entries.size, 0);
});

test('unavailable or full storage reports failure without interrupting editing', t => {
  storage(t, { getItem() { throw new Error('Blocked'); }, setItem() { throw new Error('Full'); }, removeItem() { throw new Error('Blocked'); } });
  assert.equal(readDraftBackup(doc).kind, 'unavailable');
  assert.equal(writeDraftBackup(doc, drafts), false);
  assert.equal(clearDraftBackup(doc.id), false);
});
