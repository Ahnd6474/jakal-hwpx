import assert from 'node:assert/strict';
import test from 'node:test';
import { installHostFileActions } from '../studio/file-actions.ts';
import { readFileSync } from 'node:fs';
import { restoreHostFileCommands, restoredFileCommands } from '../studio/file-command-policy.mjs';

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

test('pinned embed policy restores only host lifecycle and native output commands', () => {
  const source = readFileSync('../build/rhwp-studio-source/rhwp-studio/src/ui/chrome-mode.ts', 'utf8').replaceAll('\r\n', '\n');
  const changed = restoreHostFileCommands(source);
  const block = changed.match(/EMBED_HIDDEN_FILE_COMMAND_IDS:[\s\S]*?\];/)[0];
  for (const id of restoredFileCommands) assert.equal(block.includes(`'${id}'`), false);
  for (const id of ['file:save-as-hwp', 'file:export-html', 'file:export-doc', 'file:clear-recent']) assert.equal(block.includes(`'${id}'`), true);
  assert.ok(changed.includes("'edit:compare-documents'"));
  assert.throws(() => restoreHostFileCommands(source.replace("'file:print',", "'file:new-print',")), /Pinned embed command changed/);
});
