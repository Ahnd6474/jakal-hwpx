import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import { resolve } from 'node:path';
import { runInNewContext } from 'node:vm';
import { build } from 'esbuild';

// Execute only the integration module against mocks, without a browser or site.
const bridge = await fs.readFile('studio/host-bridge.ts', 'utf8');
const bundled = await build({
  stdin: { contents: `const {window,document,location,registry,inputHandler,wasm,canvasView,eventBus,documentState,setChromeVisibility,setThemeMode,initPromise} = env;\n${bridge}`, loader: 'ts', resolveDir: resolve('.') },
  bundle: true, format: 'iife', platform: 'node', write: false, logLevel: 'silent',
  alias: { '@jakal/input-profile': resolve('studio/input-profile.ts'), '@jakal/block-insertion': resolve('studio/block-insertion.ts'), '@jakal/file-actions': resolve('studio/file-actions.ts'), '@jakal/table-tools': resolve('studio/table-tools.ts') },
});
function fixture() {
  const definitions = new Map(), handlers = new Map(), messages = [], edits = [];
  definitions.set('table:create', { id: 'table:create', label: '표 만들기' });
  let focused = 0, offset = 4, paragraph = 7;
  const cursor = { isInHeaderFooter: () => false, isInFootnote: () => false, isInTextBox: () => false };
  const engine = { insertText: (_s, _p, at, text) => JSON.stringify({ ok: true, charOffset: offset = at + Array.from(text).length }), splitParagraph: () => JSON.stringify({ ok: true, paraIdx: ++paragraph }) };
  const parent = { postMessage: (message, origin) => messages.push({ message, origin }) };
  const env = {
    window: { parent, addEventListener() {} },
    document: { getElementById: () => null, addEventListener: (event, run) => handlers.set(event, run), documentElement: { classList: { contains: () => false, toggle() {}, add() {} } } },
    location: { origin: 'https://unit.test', search: '' }, registry: { get: id => definitions.get(id), register: value => definitions.set(value.id, value) },
    inputHandler: { cursor, focus: () => focused++, getPosition: () => ({ sectionIndex: 2, paragraphIndex: 7, charOffset: 4 }), executeOperation: descriptor => { edits.push(descriptor); descriptor.operation(engine); } },
    wasm: {}, canvasView: null, eventBus: { on() {} }, documentState: { isDirty: () => true }, setChromeVisibility() {}, setThemeMode() {}, initPromise: Promise.resolve(),
  };
  runInNewContext(bundled.outputFiles[0].text, { env, URLSearchParams, location: env.location });
  return { definitions, handlers, messages, edits, cursor, focusCount: () => focused };
}
test('block command gates body context and uses the current caret in one snapshot operation', () => {
  const f = fixture(), command = f.definitions.get('jakal:insert-block');
  const context = { hasDocument: true, isEditable: true };
  assert.equal(command.canExecute(context), true);
  for (const key of ['isFormMode', 'hasSelection', 'inTable', 'inField', 'inPictureObjectSelection', 'inTableObjectSelection']) assert.equal(command.canExecute({ ...context, [key]: true }), false);
  for (const key of Object.keys(f.cursor)) { f.cursor[key] = () => true; assert.equal(command.canExecute(context), false); f.cursor[key] = () => false; }
  command.execute({}, { text: '\n문항😀\n' });
  assert.equal(f.focusCount(), 1);
  assert.equal(f.edits.length, 1);
  assert.equal(f.edits[0].kind, 'snapshot');
  assert.equal(f.edits[0].operationType, 'jakalInsertBlock');
});
test('workbench shortcut is handled before the embedded editor and posts no text', () => {
  const f = fixture(); let prevented = 0;
  f.handlers.get('keydown')({ ctrlKey: true, shiftKey: true, code: 'KeyK', preventDefault: () => prevented++, stopImmediatePropagation: () => prevented++ });
  assert.equal(prevented, 2);
  assert.equal(f.messages[0].message.type, 'workbench');
  assert.equal(f.messages[0].origin, 'https://unit.test');
  assert.equal('text' in f.messages[0].message, false);
});

test('copy shortcut and document drops reach the host; image drops stay native', () => {
  const f = fixture(); let stopped = 0;
  f.handlers.get('keydown')({ ctrlKey: true, shiftKey: true, code: 'KeyS', preventDefault() {}, stopImmediatePropagation() {} });
  assert.equal(f.messages[0].message.type, 'file-action');
  assert.equal(f.messages[0].message.action, 'copy');
  const drop = f.handlers.get('drop');
  const file = { name: '문서.hwpx' };
  drop({ dataTransfer: { files: [file] }, preventDefault() { stopped++; }, stopImmediatePropagation() { stopped++; } });
  assert.equal(stopped, 2);
  assert.equal(f.messages[1].message.file, file);
  drop({ dataTransfer: { files: [file, file] }, preventDefault() {}, stopImmediatePropagation() {} });
  assert.equal(f.messages[2].message.count, 2);
  assert.equal(Object.hasOwn(f.messages[2].message, 'file'), false);
  drop({ dataTransfer: { files: [{ name: 'image.png' }] }, preventDefault() {}, stopImmediatePropagation() {} });
  assert.equal(f.messages.length, 3);
});
