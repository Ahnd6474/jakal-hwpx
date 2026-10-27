import assert from 'node:assert/strict';
import test from 'node:test';
import { resolve } from 'node:path';
import { build } from 'esbuild';
import { JSDOM } from 'jsdom';

const module = await build({ entryPoints: ['studio/object-layout.ts'], bundle: true, format: 'esm', platform: 'node', write: false,
  alias: { '@': resolve('../build/rhwp-studio-source/rhwp-studio/src') } });
const { installObjectLayout, resolveLayoutSelection } = await import(`data:text/javascript;base64,${Buffer.from(module.outputFiles[0].text).toString('base64')}`);

function fixture(t) {
  const dom = new JSDOM('<header id="studio-header"><nav id="menu-bar"><div class="menu-item" data-menu="tool"><span class="menu-title">도구</span></div><button id="toolbox-basic-toggle"></button></nav><div id="style-bar"></div></header>');
  const previous = globalThis.document; globalThis.document = dom.window.document;
  t.after(() => { dom.window.close(); if (previous) globalThis.document = previous; else delete globalThis.document; });
  const context = { hasDocument: true, isEditable: true, inTable: false, inPictureObjectSelection: false };
  const position = { sectionIndex: 0, parentParaIndex: 1, controlIndex: 2 };
  let ref = null, refs = [], charts = [], enumerations = 0, focused = 0;
  const input = { getCursorPosition: () => position, getSelectedPictureRef: () => ref, getSelectedPictureRefs: () => refs,
    getSelectedTableRef: () => null, focus: () => focused++ };
  const wasm = { listCharts: () => { enumerations++; return charts; } };
  const commands = new Map(), calls = [], listeners = new Map();
  const tableIds = ['cell-props', 'border-each', 'cell-split', 'cell-merge', 'cell-width-equal', 'cell-height-equal', 'insert-row-above', 'insert-row-below', 'insert-col-left', 'insert-col-right', 'create-nested', 'parent-cell'];
  for (const id of tableIds) commands.set(`table:${id}`, { canExecute: ctx => id === 'cell-merge' ? ctx.hasMultiCellSelection : ctx.inTable });
  for (const id of ['picture-props', 'chart-data-edit', 'rotate-cw', 'rotate-ccw', 'flip-horz', 'flip-vert', 'group-shapes']) commands.set(`insert:${id}`, { canExecute: ctx => ctx.inPictureObjectSelection });
  const registry = { get: id => commands.get(id), register: command => commands.set(command.id, command) };
  const dispatcher = { dispatch(id) { calls.push(id); commands.get(id)?.execute?.(); return true; } };
  // The existing MenuBar delegates direct command buttons on the menu row.
  document.getElementById('menu-bar').addEventListener('click', event => {
    const command = event.target.closest('.menu-command[data-cmd]');
    if (command) dispatcher.dispatch(command.dataset.cmd);
  });
  installObjectLayout(registry, dispatcher, { on: (event, callback) => listeners.set(event, callback) }, () => context, () => input, wasm);
  const emit = (event = 'picture-object-selection-changed') => listeners.get(event)();
  const choose = (type, extra = {}) => { context.inPictureObjectSelection = true; ref = { sec: 0, ppi: 2, ci: 1, type, ...extra }; refs = [ref]; emit(); };
  return { dom, context, input, wasm, commands, calls, emit, choose, setCharts: value => charts = value,
    setRefs: value => refs = value, enumerations: () => enumerations, focused: () => focused,
    region: document.querySelector('.jakal-object-layout'), panel: document.querySelector('[role=tabpanel]'),
    tab: document.querySelector('[role=tab]'), entry: document.querySelector('[role=tab]'),
    button: id => document.querySelector(`[data-command="${id}"]`) };
}

test('table selection shows layout actions, preserves selection on mouse down and dispatches split/merge separately', t => {
  const f = fixture(t); assert.equal(f.region.hidden, true);
  f.context.inTable = true; f.emit('cursor-cell-changed');
  assert.equal(f.tab.textContent, '표 레이아웃'); assert.equal(f.region.hidden, false);
  const tabs = f.tab.parentElement;
  assert.equal(tabs.previousElementSibling.dataset.menu, 'tool');
  assert.equal(tabs.nextElementSibling.id, 'toolbox-basic-toggle');
  assert.equal(f.region.querySelector('[role=tablist]'), null, 'no second tab row below the menu');
  assert.equal(document.getElementById('studio-header').classList.contains('jakal-layout-active'), true);
  assert.equal(f.button('table:cell-merge').disabled, true);
  const down = new f.dom.window.MouseEvent('mousedown', { bubbles: true, cancelable: true });
  f.button('table:cell-split').dispatchEvent(down); assert.equal(down.defaultPrevented, true);
  f.context.hasMultiCellSelection = true; f.emit('command-state-changed');
  for (const id of ['table:cell-split', 'table:cell-merge', 'table:create-nested', 'table:cell-props']) f.button(id).click();
  assert.deepEqual(f.calls, ['table:cell-split', 'table:cell-merge', 'table:create-nested', 'table:cell-props']);
  f.context.inTable = false; f.emit('selection-changed'); assert.equal(f.region.hidden, true);
  assert.equal(tabs.hidden, true);
  assert.equal(document.getElementById('studio-header').classList.contains('jakal-layout-active'), false);
});

test('a normal menu restores standard tools while keeping the contextual tab available', t => {
  const f = fixture(t); f.choose('image');
  document.querySelector('.menu-title').dispatchEvent(new f.dom.window.MouseEvent('mousedown', { bubbles: true }));
  assert.equal(f.region.hidden, true); assert.equal(f.tab.parentElement.hidden, false);
  assert.equal(f.tab.getAttribute('aria-selected'), 'false');
  f.emit('cursor-format-changed'); assert.equal(f.region.hidden, true);
  f.tab.click(); assert.equal(f.region.hidden, false);
  assert.equal(f.tab.getAttribute('aria-selected'), 'true');
  assert.deepEqual(f.calls, ['view:object-layout']);
});

test('image selection inside a table takes priority and offers size, position, rotation and flipping', t => {
  const f = fixture(t); f.context.inTable = true; f.choose('image');
  assert.equal(f.tab.textContent, '그림 레이아웃'); assert.equal(f.button('table:cell-split'), null);
  f.button('insert:picture-props').click(); f.button('insert:rotate-cw').click(); f.button('insert:flip-horz').click();
  assert.deepEqual(f.calls, ['insert:picture-props', 'insert:rotate-cw', 'insert:flip-horz']);
  f.context.isFormMode = true; f.emit('command-state-changed');
  assert.equal(f.button('insert:picture-props').disabled, true);
  f.context.isFormMode = false; f.context.isEditable = false; f.emit('command-state-changed');
  assert.equal(f.button('insert:rotate-cw').disabled, true);
});

test('nested table tools do not route property dialogs to the outer table', t => {
  const f = fixture(t); f.context.inTable = true;
  f.input.getCursorPosition().cellPath = [{ controlIndex: 2, cellIndex: 0, cellParaIndex: 1 }, { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }];
  f.emit('cursor-cell-changed');
  assert.equal(f.button('table:cell-props').disabled, true);
  assert.equal(f.button('table:border-each').disabled, true);
  assert.equal(f.button('table:cell-split').disabled, false);
  assert.equal(f.button('table:create-nested').disabled, false);
});

test('chart matching uses the full container path, caches enumeration and never exposes chart tools for unrelated OLEs', t => {
  const f = fixture(t);
  f.setCharts([{ index: 0, section: 0, paragraph: 2, control: 1, container: [{ kind: 'tableCell', control: 3, paragraph: 0, cell: 1 }] }]);
  f.choose('ole'); assert.equal(f.tab.textContent, '개체 레이아웃'); assert.equal(f.button('insert:chart-data-edit'), null);
  f.choose('ole', { cellPath: [{ controlIndex: 3, cellIndex: 1, cellParaIndex: 0 }] });
  assert.equal(f.tab.textContent, '차트 레이아웃'); f.button('insert:chart-data-edit').click();
  assert.deepEqual(f.calls, ['insert:chart-data-edit']);
  const count = f.enumerations(); f.emit('cursor-format-changed'); f.emit('command-state-changed'); assert.equal(f.enumerations(), count);
  f.setCharts([]); f.emit('document-mutated'); assert.equal(f.tab.textContent, '개체 레이아웃');
});

test('layout command works without a selection, folds and reopens, and never edits the document', t => {
  const f = fixture(t); f.commands.get('view:object-layout').execute();
  assert.equal(f.region.hidden, false); assert.match(f.panel.textContent, /클릭하세요/);
  f.choose('image');
  document.querySelector('.jakal-layout-collapse').click(); assert.equal(f.panel.hidden, true);
  f.emit('cursor-format-changed'); assert.equal(f.panel.hidden, true);
  f.tab.click(); assert.equal(f.panel.hidden, false);
  f.region.dispatchEvent(new f.dom.window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  assert.equal(f.panel.hidden, true); assert.equal(f.focused(), 1);
  f.entry.click(); assert.equal(f.panel.hidden, false);
  f.context.hasDocument = false; f.emit('document-loaded'); assert.equal(f.region.hidden, true); assert.equal(f.entry.disabled, true);
  assert.equal(f.commands.get('view:object-layout').canExecute(f.context), false);
});

test('multiple images offer group tools instead of editing an arbitrary first image', t => {
  const f = fixture(t); f.choose('image');
  f.setRefs([{ sec: 0, ppi: 2, ci: 1, type: 'image' }, { sec: 0, ppi: 2, ci: 2, type: 'image' }]); f.emit();
  assert.equal(f.tab.textContent, '개체 레이아웃'); assert.equal(f.button('insert:picture-props'), null);
  assert.match(f.panel.textContent, /여러 개체/);
  assert.equal(resolveLayoutSelection({ hasDocument: false }, f.input, f.wasm), null);
});
