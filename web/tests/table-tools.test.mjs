import assert from 'node:assert/strict';
import test from 'node:test';
import { JSDOM } from 'jsdom';
import { showTableDialog, installTableTools } from '../studio/table-tools.ts';

function fixture(t) {
  const dom = new JSDOM('', { url: 'https://editor.test' });
  const previous = { document: globalThis.document, localStorage: globalThis.localStorage };
  globalThis.document = dom.window.document; globalThis.localStorage = dom.window.localStorage;
  dom.window.HTMLDialogElement.prototype.showModal = function () { this.open = true; };
  dom.window.HTMLDialogElement.prototype.close = function () { this.open = false; };
  t.after(() => { dom.window.close(); for (const key in previous) { if (previous[key] === undefined) delete globalThis[key]; else globalThis[key] = previous[key]; } });
  const submit = dialog => dialog.querySelector('form').dispatchEvent(new dom.window.Event('submit', { bubbles: true, cancelable: true }));
  return { dom, submit };
}

test('table dialog applies presets, remembers sizes and returns focus on apply or cancel', t => {
  const { dom, submit } = fixture(t);
  const calls = []; let focuses = 0;
  let dialog = showTableDialog(true, (...args) => calls.push(args), () => focuses++);
  assert.match(dialog.textContent, /셀 안에 표 만들기/);
  assert.equal(document.activeElement.name, 'rows');
  dialog.querySelector('[data-rows="5"]').click();
  assert.match(dialog.querySelector('[role="status"]').textContent, /5줄 × 2칸 · 10셀/);
  submit(dialog);
  assert.deepEqual(calls, [[5, 2]]); assert.equal(focuses, 1); assert.equal(dialog.isConnected, false);
  dialog = showTableDialog(false, () => assert.fail('cancel must not apply'), () => focuses++);
  assert.equal(dialog.querySelector('[name="rows"]').value, '5');
  dialog.dispatchEvent(new dom.window.Event('cancel', { cancelable: true }));
  assert.equal(dialog.isConnected, false); assert.equal(focuses, 2);
});

test('invalid dimensions stay blocked and a failed insertion keeps the dialog and error visible', t => {
  const { dom, submit } = fixture(t);
  let calls = 0;
  const dialog = showTableDialog(true, () => { calls++; throw new Error('보호된 셀'); }, () => {});
  const rows = dialog.querySelector('[name="rows"]');
  rows.value = '0'; rows.dispatchEvent(new dom.window.Event('input', { bubbles: true }));
  assert.equal(dialog.querySelector('[type="submit"]').disabled, true);
  submit(dialog); assert.equal(calls, 0);
  rows.value = '2'; rows.dispatchEvent(new dom.window.Event('input', { bubbles: true }));
  submit(dialog);
  assert.equal(calls, 1); assert.equal(dialog.isConnected, true);
  assert.equal(dialog.querySelector('[role="alert"]').textContent, '보호된 셀');
});

test('the create command keeps a frozen cell path and uses one snapshot transaction', t => {
  const { submit } = fixture(t);
  const commands = new Map([['table:create', { id: 'table:create', label: '표 만들기' }]]);
  const registry = { get: id => commands.get(id), register: command => commands.set(command.id, command) };
  const position = { sectionIndex: 0, paragraphIndex: 0, charOffset: 3, parentParaIndex: 1, controlIndex: 0,
    cellIndex: 2, cellParaIndex: 0, cellPath: [{ controlIndex: 0, cellIndex: 2, cellParaIndex: 0 }] };
  const operations = [], inserted = [];
  const engine = { documentGeneration: 1, createNestedTable: (...args) => { inserted.push(args); return args[0]; } };
  const input = { cursor: { isInHeaderFooter: () => false, isInFootnote: () => false, isInTextBox: () => false },
    getCursorPosition: () => position, focus: () => {}, executeOperation: operation => { operations.push(operation); operation.operation(engine); } };
  installTableTools(registry, () => input);
  const context = { hasDocument: true, isEditable: true, inTable: true };
  assert.equal(commands.get('table:create').canExecute(context), true);
  assert.equal(commands.get('table:create').canExecute({ ...context, isEditable: false }), false);
  assert.equal(commands.get('table:create').canExecute({ ...context, inCellSelectionMode: true }), false);
  assert.equal(commands.get('table:parent-cell').canExecute(context), false);
  commands.get('table:create-nested').execute({ getInputHandler: () => input, wasm: engine });
  position.cellPath[0].cellIndex = 9;
  submit(document.querySelector('dialog'));
  assert.equal(operations.length, 1); assert.equal(operations[0].kind, 'snapshot');
  assert.equal(inserted[0][0].cellPath[0].cellIndex, 2);
});
