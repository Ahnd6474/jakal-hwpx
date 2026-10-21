import assert from 'node:assert/strict';
import test from 'node:test';
import { installHostFileActions } from '../studio/file-actions.ts';

test('embedded file lifecycle routes through the host without marking native edits saved', () => {
  let native = 0;
  const definitions = new Map(['file:open', 'file:new-doc', 'file:save', 'file:save-as', 'file:save-as-hwpx', 'file:save-as-hwp', 'file:print-to-pdf', 'edit:document-history'].map(id => [id, { id, label: id, canExecute: () => true, execute: () => native++ }]));
  const actions = [];
  installHostFileActions({ get: id => definitions.get(id), register: item => definitions.set(item.id, item) }, action => actions.push(action));
  for (const id of ['file:open', 'file:new-doc', 'file:save', 'file:save-as', 'file:save-as-hwpx', 'edit:document-history']) definitions.get(id).execute();
  assert.deepEqual(actions, ['open', 'new', 'save', 'copy', 'download', 'history']);
  assert.equal(native, 0);
  assert.equal(definitions.get('file:save-as-hwp').canExecute(), false);
  definitions.get('file:print-to-pdf').execute(); assert.equal(native, 1);
});
