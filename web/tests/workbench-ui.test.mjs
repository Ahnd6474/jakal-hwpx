// Isolated component tests: no browser, network requests or live documents.
import assert from 'node:assert/strict';
import test, { after } from 'node:test';
import { build } from 'esbuild';
import { JSDOM } from 'jsdom';
import { mkdtemp, rm } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import React, { act } from 'react';
// React detects input event support at import time, after a DOM is available.
const bootstrap = new JSDOM('<html></html>');
globalThis.window = bootstrap.window; globalThis.document = bootstrap.window.document;
const { createRoot } = await import('react-dom/client');
delete globalThis.window; delete globalThis.document; bootstrap.window.close();

const temp = await mkdtemp(resolve('node_modules', '.jakal-ui-'));
after(() => rm(temp, { recursive: true, force: true }));
async function component(name) {
  const out = resolve(temp, name + '.mjs');
  const plugins = name === 'PageEditor' ? [{ name: 'mock-editor', setup(builder) {
    builder.onResolve({ filter: /^@rhwp\/editor$/ }, () => ({ path: 'editor', namespace: 'mock-editor' }));
    builder.onLoad({ filter: /.*/, namespace: 'mock-editor' }, () => ({ contents: 'export const createStudio = (...args) => globalThis.__jakalCreateStudio(...args);' }));
  } }] : [];
  await build({ entryPoints: [resolve('src', name + '.tsx')], outfile: out, bundle: true, platform: 'node', format: 'esm', packages: 'external', logLevel: 'silent', plugins });
  return (await import(pathToFileURL(out).href))[name];
}
const EditorWorkbench = await component('EditorWorkbench');
const NewDocument = await component('NewDocument');
const ToolBoundary = await component('ToolBoundary');
const DocumentReader = await component('DocumentReader');
const AdvancedSettings = await component('AdvancedSettings');
const ConnectionStatus = await component('ConnectionStatus');
const ApplicationBoundary = await component('ApplicationBoundary');
const DocumentActions = await component('DocumentActions');
const BackupShelf = await component('BackupShelf');
const PageEditor = await component('PageEditor');
const DocumentSyncNotice = await component('DocumentSyncNotice');
function fixture(t) {
  const dom = new JSDOM('<button id="previous">Previous</button><div id="root"></div>', { url: 'https://unit.test/' });
  const errors = []; dom.window.addEventListener('error', event => errors.push(event.error));
  const old = {};
  for (const [key, value] of Object.entries({ window: dom.window, document: dom.window.document, HTMLElement: dom.window.HTMLElement, Node: dom.window.Node, localStorage: dom.window.localStorage, IS_REACT_ACT_ENVIRONMENT: true })) {
    old[key] = Object.getOwnPropertyDescriptor(globalThis, key);
    Object.defineProperty(globalThis, key, { configurable: true, value, writable: true });
  }
  const caught = [];
  const root = createRoot(document.querySelector('#root'), { onCaughtError: error => caught.push(error) });
  t.after(async () => { await act(async () => root.unmount()); dom.window.close(); for (const key of Object.keys(old)) { if (old[key]) Object.defineProperty(globalThis, key, old[key]); else delete globalThis[key]; } assert.deepEqual(errors, []); });
  return { root, dom, caught };
}
const click = async element => { assert.ok(element); await act(async () => element.click()); };
const button = name => [...document.querySelectorAll('button')].find(e => e.textContent.includes(name));

test('general editing tools respect context, pin commands and close with Escape', async t => {
  const { root, dom } = fixture(t);
  let closed = 0, ran = '';
  await act(async () => root.render(React.createElement(EditorWorkbench, {
    onClose: () => closed++, onRun: async id => { ran = id; },
    commands: async () => [{ id: 'insert:equation', label: '수식', enabled: true }, { id: 'table:merge', label: '셀 병합', enabled: false }, { id: 'edit:undo', label: '실행 취소', enabled: false }, { id: 'edit:find-replace', label: '찾아 바꾸기', enabled: true, opensDialog: true }],
  })));
  assert.equal(document.querySelector('[aria-label="작업 도구 검색"]'), document.activeElement);
  assert.equal(button('셀 병합').disabled, true);
  await click(document.querySelector('[aria-label="수식 즐겨찾기"]'));
  assert.deepEqual(JSON.parse(localStorage.getItem('jakal-tool-favorites')), ['insert:equation']);
  await click(button('수식')); assert.equal(ran, 'insert:equation');
  assert.equal(document.querySelector('[aria-label="기본 편집 도구"] button').disabled, true);
  await click(button('찾아 바꾸기')); assert.equal(ran, 'edit:find-replace');
  assert.equal(button('문항 블록'), undefined);
  await act(async () => document.querySelector('[role=dialog]').dispatchEvent(new dom.window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true })));
  assert.equal(closed, 1);
});

test('new document chooses a general template and submits its title', async t => {
  const { root } = fixture(t);
  const oldFetch = globalThis.fetch;
  t.after(() => { globalThis.fetch = oldFetch; });
  globalThis.fetch = async () => ({ ok: true, json: async () => [{ id: 'blank', name: '빈 문서', title: '새 문서', description: 'A4' }, { id: 'meeting', name: '회의록', title: '회의록', description: '일시와 안건' }] });
  let created;
  await act(async () => root.render(React.createElement(NewDocument, { busy: false, error: '', onClose() {}, onCreate: value => { created = value; } })));
  await click(button('회의록'));
  assert.equal(document.querySelector('.creation-field input').value, '회의록');
  assert.equal(document.querySelector('input[type=number]'), null);
  await click(button('만들기'));
  assert.deepEqual(created, { template: 'meeting', title: '회의록' });
});

test('an optional tool failure preserves the editor and offers a close action', async t => {
  const { root, caught } = fixture(t);
  let closed = false;
  function BrokenTool() { throw new Error('missing optional chunk'); }
  await act(async () => root.render(React.createElement(React.Fragment, null,
    React.createElement('div', { id: 'editor-proof' }, 'document remains open'),
    React.createElement(ToolBoundary, { onClose: () => { closed = true; } }, React.createElement(BrokenTool)),
  )));
  assert.equal(caught.length, 1);
  assert.equal(document.querySelector('#editor-proof').textContent, 'document remains open');
  assert.match(document.querySelector('[role=alert]').textContent, /저장한 뒤/);
  await click(button('닫고 편집 계속')); assert.equal(closed, true);
});

test('reader keeps its worker on hide and defers a hidden revision reload until reopened', async t => {
  const { root, dom, caught } = fixture(t);
  const workers = [], fetched = [];
  class WorkerMock {
    messages = []; terminated = false;
    constructor() { workers.push(this); }
    postMessage(message) { this.messages.push(message); }
    terminate() { this.terminated = true; }
  }
  const mocks = {
    Worker: WorkerMock,
    ResizeObserver: class { observe() {} disconnect() {} },
    getComputedStyle: dom.window.getComputedStyle.bind(dom.window),
    requestAnimationFrame: () => 0, cancelAnimationFrame() {},
    fetch: async url => { fetched.push(url); return { ok: true, arrayBuffer: async () => new ArrayBuffer(0) }; },
  };
  const previous = {};
  for (const [key, value] of Object.entries(mocks)) { previous[key] = Object.getOwnPropertyDescriptor(globalThis, key); Object.defineProperty(globalThis, key, { configurable: true, value, writable: true }); }
  t.after(() => { for (const key of Object.keys(previous)) { if (previous[key]) Object.defineProperty(globalThis, key, previous[key]); else delete globalThis[key]; } });
  dom.window.HTMLElement.prototype.scrollTo = () => {};
  Object.defineProperty(document, 'hidden', { configurable: true, value: false });
  const doc = { id: 'reader-unit', revision: 0, format: 'hwpx', name: 'unit.hwpx', sections: [] };
  const props = { doc, active: true, mode: 'original', dirty: 0, onMode() {}, onEdit() {}, onEditTable() {}, toolsCollapsed: true, onToggleTools() {}, onReadingScroll() {} };
  await act(async () => root.render(React.createElement(DocumentReader, props)));
  assert.equal(workers.length, 1); assert.equal(fetched.length, 1);
  await act(async () => workers[0].onmessage({ data: { type: 'ready', pages: [], figures: [], tables: [] } }));
  await act(async () => root.render(React.createElement(DocumentReader, { ...props, active: false })));
  assert.equal(workers.length, 1); assert.equal(workers[0].terminated, false);
  assert.deepEqual(workers[0].messages.at(-1), { type: 'activity', active: false });
  const changed = { ...props, active: false, doc: { ...doc, revision: 1 } };
  await act(async () => root.render(React.createElement(DocumentReader, changed)));
  assert.equal(workers[0].terminated, true); assert.equal(workers.length, 1); assert.equal(fetched.length, 1);
  await act(async () => root.render(React.createElement(DocumentReader, { ...changed, active: true })));
  assert.equal(workers.length, 2); assert.equal(fetched.length, 2); assert.match(fetched[1], /revision=1/);
  await act(async () => workers[0].onmessage({ data: { type: 'error', message: 'stale worker response' } }));
  assert.equal(document.body.textContent.includes('stale worker response'), false);
  assert.equal(caught.length, 0);
});

test('advanced settings uses roving tab focus with arrows, Home and End', async t => {
  const { root, dom } = fixture(t);
  const doc = { id: 'test', name: '문서', revision: 0, settings_version: 0, sections: [], settings: { instructions: '', protected_terms: [], strict_validation: false } };
  await act(async () => root.render(React.createElement(AdvancedSettings, { doc, drafts: {}, onClose() {}, onApply() {}, onSettings() {} })));
  const tabs = [...document.querySelectorAll('[role=tab]')];
  assert.deepEqual(tabs.map(tab => tab.tabIndex), [0, -1]);
  tabs[0].focus();
  const key = async name => act(async () => document.activeElement.dispatchEvent(new dom.window.KeyboardEvent('keydown', { key: name, bubbles: true, cancelable: true })));
  await key('ArrowRight');
  assert.equal(document.activeElement, tabs[1]); assert.equal(tabs[1].getAttribute('aria-selected'), 'true');
  assert.equal(document.querySelector('[role=tabpanel]').id, tabs[1].getAttribute('aria-controls'));
  await key('Home'); assert.equal(document.activeElement, tabs[0]);
  await key('End'); assert.equal(document.activeElement, tabs[1]);
  await key('ArrowRight'); assert.equal(document.activeElement, tabs[0]);
  assert.deepEqual(tabs.map(tab => tab.tabIndex), [0, -1]);
});

test('offline notice follows connectivity events without remounting the editor', async t => {
  const { root, dom } = fixture(t);
  let online = true;
  Object.defineProperty(dom.window.navigator, 'onLine', { get: () => online, configurable: true });
  await act(async () => root.render(React.createElement(React.Fragment, null,
    React.createElement('textarea', { defaultValue: 'unsaved work' }), React.createElement(ConnectionStatus))));
  const editor = document.querySelector('textarea');
  assert.equal(document.querySelector('[role=status]'), null);
  online = false;
  await act(async () => window.dispatchEvent(new dom.window.Event('offline')));
  assert.match(document.querySelector('[role=status]').textContent, /편집은 계속/);
  assert.equal(document.querySelector('textarea'), editor);
  assert.equal(editor.value, 'unsaved work');
  online = true;
  await act(async () => window.dispatchEvent(new dom.window.Event('online')));
  assert.equal(document.querySelector('[role=status]'), null);
});

test('application failure shows recovery guidance without removing stored drafts', async t => {
  const { root, caught } = fixture(t);
  localStorage.setItem('draft-proof', 'preserved');
  function BrokenApp() { throw new Error('render failed'); }
  await act(async () => root.render(React.createElement(ApplicationBoundary, null, React.createElement(BrokenApp))));
  assert.equal(caught.length, 1);
  assert.match(document.querySelector('[role=alert]').textContent, /임시보관되지 않았을 수/);
  assert.equal(localStorage.getItem('draft-proof'), 'preserved');
  assert.ok(button('페이지 다시 불러오기'));
});

async function settleUntil(check) {
  for (let attempt = 0; attempt < 100 && !check(); attempt++) await act(async () => new Promise(resolve => setImmediate(resolve)));
  assert.ok(check(), 'component did not settle');
}

test('copy dialog explains saving current changes and preserves the name on a conflict', async t => {
  const { root } = fixture(t);
  const doc = { id: 'doc', name: '회의록.hwpx' };
  const submissions = [];
  await act(async () => root.render(React.createElement(DocumentActions, {
    doc, initialMode: 'copy', dirty: true, onClose() {},
    onSubmit: async (mode, name) => { submissions.push({ mode, name }); throw new Error('다른 탭에서 이름이 변경되었습니다.'); },
  })));
  assert.equal(document.querySelector('input').value, '회의록_복사본');
  assert.match(document.querySelector('.document-action-note').textContent, /현재 문서에도 수정 내용이 저장/);
  await click(button('현재 문서 저장 후 복사'));
  assert.deepEqual(submissions, [{ mode: 'copy', name: '회의록_복사본' }]);
  assert.equal(document.querySelector('input').value, '회의록_복사본');
  assert.match(document.querySelector('[role=alert]').textContent, /다른 탭/);
  await click(button('이름 변경'));
  assert.equal(document.querySelector('input').value, '회의록');
});

test('recovery shelf imports a separate file, keeps its backup and protects the active draft', async t => {
  const { root } = fixture(t);
  const original = Object.getOwnPropertyDescriptor(globalThis, 'indexedDB');
  const { IDBFactory } = await import('fake-indexeddb');
  globalThis.indexedDB = new IDBFactory();
  t.after(() => { if (original) Object.defineProperty(globalThis, 'indexedDB', original); else delete globalThis.indexedDB; });
  const { studioBackup } = await import('../src/studio-backup.ts');
  await studioBackup.write({ id: 'orphan', revision: 3, saved: 2, name: '고아 임시본.hwpx', bytes: new Uint8Array([80, 75, 3, 4]) });
  await studioBackup.write({ id: 'active', revision: 0, saved: 1, name: '현재.hwpx', bytes: new Uint8Array([80, 75]) });
  let imported;
  await act(async () => root.render(React.createElement(BackupShelf, { activeId: 'active', names: {}, onClose() {}, onOpen: async file => { imported = file; } })));
  await settleUntil(() => document.querySelectorAll('.backup-shelf-list li').length === 2);
  assert.equal(document.querySelector('[aria-label="현재.hwpx 임시본 삭제"]').disabled, true);
  await click(button('새 문서로 열기'));
  await settleUntil(() => !!imported);
  assert.equal(imported.name, '고아 임시본_임시편집.hwpx');
  assert.deepEqual(new Uint8Array(await imported.arrayBuffer()), new Uint8Array([80, 75, 3, 4]));
  assert.ok(await studioBackup.read('orphan'));
  await click(document.querySelector('[aria-label="고아 임시본.hwpx 임시본 삭제"]'));
  assert.ok(await studioBackup.read('orphan'));
  await click(button('삭제 확인'));
  await settleUntil(() => document.querySelectorAll('.backup-shelf-list li').length === 1);
  assert.equal(await studioBackup.read('orphan'), undefined);
  assert.ok(await studioBackup.read('active'));
});

test('task groups surface printing, links and table tools without bypassing document saves', async t => {
  const { root } = fixture(t);
  let ran = '';
  await act(async () => root.render(React.createElement(EditorWorkbench, {
    onClose() {}, onRun: async id => { ran = id; }, commands: async () => [
      { id: 'file:print-to-pdf', label: 'PDF로 저장', enabled: true },
      { id: 'file:save', label: '내부 저장', enabled: true },
      { id: 'table:cell-merge', label: '셀 병합', enabled: false },
      { id: 'hyperlink:edit', label: '하이퍼링크 수정', enabled: true },
    ],
  })));
  assert.equal(button('내부 저장'), undefined);
  assert.ok(button('하이퍼링크 수정'));
  await click(button('표 다듬기'));
  assert.equal(button('셀 병합').disabled, true);
  assert.equal(button('PDF로 저장'), undefined);
  await click(button('검토·출력'));
  await click(button('PDF로 저장'));
  assert.equal(ran, 'file:print-to-pdf');
});

test('editor keeps unsaved content on rename and accepts file actions only from its own frame', async t => {
  const { root, dom, caught } = fixture(t);
  const { IDBFactory } = await import('fake-indexeddb');
  let frame, loads = 0, notifications = 0;
  const sent = [], actions = [], dirty = [], commands = [];
  const previous = {};
  const mocks = {
    indexedDB: new IDBFactory(), location: dom.window.location,
    requestAnimationFrame: run => { run(0); return 1; },
    fetch: async () => ({ ok: true, arrayBuffer: async () => new ArrayBuffer(0) }),
    __jakalCreateStudio: async mount => {
      frame = document.createElement('iframe'); mount.appendChild(frame);
      frame.contentWindow.postMessage = (data, origin) => sent.push({ data, origin });
      frame.contentWindow.focus = () => {};
      return { element: frame, loadFile: async () => { loads++; }, destroy: () => frame.remove(), notifySaved: async () => { notifications++; }, commands: { isEnabled: async () => true, execute: async id => { commands.push(id); return { ok: true }; } } };
    },
  };
  for (const [key, value] of Object.entries(mocks)) { previous[key] = Object.getOwnPropertyDescriptor(globalThis, key); Object.defineProperty(globalThis, key, { value, configurable: true, writable: true }); }
  dom.window.matchMedia = () => ({ matches: false, addEventListener() {}, removeEventListener() {} });
  t.after(() => { for (const key in previous) { if (previous[key]) Object.defineProperty(globalThis, key, previous[key]); else delete globalThis[key]; } });
  const handle = React.createRef();
  const props = { ref: handle, doc: { id: 'doc', name: '이전.hwpx', revision: 0 }, active: true, disabled: false, onSave() {}, onFileAction: (...args) => actions.push(args), onDirty: value => dirty.push(value) };
  await act(async () => root.render(React.createElement(PageEditor, props)));
  await settleUntil(() => !document.querySelector('.studio-loading'));
  const emit = async (data, origin = 'https://unit.test', source = frame.contentWindow) => act(async () => window.dispatchEvent(new dom.window.MessageEvent('message', { data: { channel: 'jakal-studio-v1', ...data }, origin, source })));
  await emit({ type: 'file-action', action: 'copy' }, 'https://other.test');
  await emit({ type: 'file-action', action: 'copy' }, 'https://unit.test', window);
  assert.deepEqual(actions, []);
  await emit({ type: 'file-action', action: 'copy' });
  assert.deepEqual(actions, [['copy']]);
  await emit({ type: 'changed', sequence: 1, dirty: true });
  assert.equal(dirty.at(-1), true);
  await act(async () => root.render(React.createElement(PageEditor, { ...props, doc: { ...props.doc, name: '바뀐이름.hwpx' } })));
  assert.equal(loads, 1);
  assert.equal(notifications, 0);
  assert.equal(dirty.at(-1), true);
  assert.deepEqual(sent.at(-1), { data: { channel: 'jakal-studio-v1', type: 'filename', name: '바뀐이름.hwpx' }, origin: 'https://unit.test' });
  await act(async () => root.render(React.createElement(PageEditor, { ...props, doc: { ...props.doc, name: '변환 원본.hwp' } })));
  assert.equal(sent.at(-1).data.name, '변환 원본.hwpx');
  assert.equal(loads, 1);
  const file = new File(['test'], '열기.hwpx');
  await emit({ type: 'file-drop', count: 1, file });
  assert.equal(actions.at(-1)[0], 'drop'); assert.equal(actions.at(-1)[1], file);
  await act(async () => handle.current.print());
  assert.deepEqual(commands, ['file:print-to-pdf']);
  assert.equal(notifications, 0);
  assert.equal(dirty.at(-1), true);
  assert.equal(caught.length, 0);
});

test('saved-state changes notify without replacing edits and only explicit actions reload or export', async t => {
  const { root, dom } = fixture(t);
  const priorFetch = globalThis.fetch;
  const priorNavigator = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: dom.window.navigator });
  Object.defineProperty(document, 'visibilityState', { configurable: true, value: 'visible' });
  t.after(() => { globalThis.fetch = priorFetch; if (priorNavigator) Object.defineProperty(globalThis, 'navigator', priorNavigator); else delete globalThis.navigator; });
  const doc = { id: 'document', revision: 1, settings_version: 0, name: '보고서.hwpx', status: 'ready' };
  let saved = { ...doc }, reloads = 0, downloads = 0;
  const paths = [];
  globalThis.fetch = async url => { paths.push(url); return new Response(JSON.stringify(saved)); };
  const props = { doc, disabled: false, dirty: true, onReload: () => reloads++, onBackup: () => downloads++ };
  await act(async () => root.render(React.createElement(DocumentSyncNotice, props)));
  assert.equal(document.querySelector('[role=status]'), null);
  saved = { ...doc, revision: 2, settings_version: 1 };
  await act(async () => window.dispatchEvent(new dom.window.Event('focus')));
  assert.match(document.querySelector('[role=status]').textContent, /현재 편집은 유지/);
  assert.equal(reloads, 0); assert.equal(downloads, 0);
  assert.equal(doc.revision, 1);
  await click(button('현재 편집 임시본 받기')); assert.equal(downloads, 1); assert.equal(reloads, 0);
  await click(button('편집 보관 후 최신 열기')); assert.equal(reloads, 1);
  assert.ok(paths.length >= 2 && paths.every(path => path === '/api/documents/document/state'));
  await act(async () => root.render(React.createElement(DocumentSyncNotice, { ...props, doc: saved })));
  assert.equal(document.querySelector('[role=status]'), null);
});

test('state checks pause during mutations and hidden pages and abort when the document closes', async t => {
  const { root, dom } = fixture(t);
  const priorFetch = globalThis.fetch;
  const priorNavigator = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: dom.window.navigator });
  let visibility = 'hidden', requests = 0, signal;
  Object.defineProperty(document, 'visibilityState', { configurable: true, get: () => visibility });
  t.after(() => { globalThis.fetch = priorFetch; if (priorNavigator) Object.defineProperty(globalThis, 'navigator', priorNavigator); else delete globalThis.navigator; });
  globalThis.fetch = async (_url, options) => {
    requests++; signal = options.signal;
    return new Promise((_resolve, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true }));
  };
  const props = { doc: { id: 'document', revision: 1, settings_version: 0, name: '파일.hwpx' }, disabled: false, dirty: true, onReload() {}, onBackup() {} };
  await act(async () => root.render(React.createElement(DocumentSyncNotice, props)));
  assert.equal(requests, 0);
  visibility = 'visible';
  await act(async () => root.render(React.createElement(DocumentSyncNotice, { ...props, disabled: true })));
  await act(async () => window.dispatchEvent(new dom.window.Event('focus')));
  assert.equal(requests, 0);
  await act(async () => root.render(React.createElement(DocumentSyncNotice, props)));
  assert.equal(requests, 1);
  await act(async () => window.dispatchEvent(new dom.window.Event('focus')));
  assert.equal(requests, 1); // Do not overlap another read while this one waits.
  await act(async () => root.render(null));
  assert.equal(signal.aborted, true);
});
