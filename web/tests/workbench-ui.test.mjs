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
  await build({ entryPoints: [resolve('src', name + '.tsx')], outfile: out, bundle: true, platform: 'node', format: 'esm', packages: 'external', logLevel: 'silent' });
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
