import assert from 'node:assert/strict';
import test from 'node:test';
import { IDBFactory } from 'fake-indexeddb';
import { JSDOM } from 'jsdom';
import { setWorkspaceContext, privateStorageKey, lockWorkspace, clearPrivateStorage } from '../src/privacy-context.ts';
import { studioBackup } from '../src/studio-backup.ts';
import { pageTemplateAssets } from '../src/page-template-assets.ts';
import { defaultLayout, savePagePreset, readPagePresets } from '../src/page-layout.ts';
import { lastDocument } from '../src/preferences.ts';
import { requestJson, requestBytes } from '../src/request.ts';

test('workspace switches isolate browser document bytes, form backgrounds, names and preferences; explicit exit clears old copies', async t => {
  const dom = new JSDOM('', { url: 'https://unit.test' });
  t.after(() => dom.window.close());
  const old = {};
  for (const [key, value] of Object.entries({ window: dom.window, localStorage: dom.window.localStorage, sessionStorage: dom.window.sessionStorage, indexedDB: new IDBFactory() })) {
    old[key] = Object.getOwnPropertyDescriptor(globalThis, key); Object.defineProperty(globalThis, key, { configurable: true, value });
  }
  t.after(() => { for (const key in old) { if (old[key]) Object.defineProperty(globalThis, key, old[key]); else delete globalThis[key]; } });
  const page = { width_mm: 210, height_mm: 297, image: 'data:image/jpeg;base64,YWJjZA==' };
  const a = 'a'.repeat(32), b = 'b'.repeat(32);
  setWorkspaceContext(a);
  await studioBackup.write({ id: 'document', name: 'private-name.hwpx', revision: 0, saved: 1, bytes: new Uint8Array([1, 2]) });
  await pageTemplateAssets.save('private template', defaultLayout, [page]);
  savePagePreset('private heading', { ...defaultLayout, header_text: 'private contents' });
  lastDocument.write('document');
  setWorkspaceContext(b);
  assert.deepEqual(await studioBackup.list(), []); assert.equal(await studioBackup.read('document'), undefined);
  assert.deepEqual(await pageTemplateAssets.list(), []); assert.deepEqual(readPagePresets(), []); assert.equal(lastDocument.read(), null);
  assert.notEqual(privateStorageKey('jakal-document-editor'), `jakal-document-editor:${a}`);
  setWorkspaceContext(a);
  assert.equal((await studioBackup.list())[0].name, 'private-name.hwpx'); assert.equal(readPagePresets()[0].name, 'private heading');
  localStorage.setItem('unrelated-theme', 'dark'); localStorage.setItem('jakal-page-presets-v1', 'legacy private data');
  let ended = 0; window.addEventListener('jakal-session-ended', () => ended++);
  lockWorkspace(); lockWorkspace(); assert.equal(ended, 1);
  await assert.rejects(studioBackup.write({ id: 'late', bytes: new Uint8Array([3]), revision: 0, saved: 2 }), /세션/);
  await clearPrivateStorage();
  assert.equal(localStorage.getItem('jakal-page-presets-v1'), null); assert.equal(localStorage.getItem('unrelated-theme'), 'dark');
  const names = (await indexedDB.databases()).map(db => db.name);
  assert.ok(names.every(name => !name.startsWith('jakal-')));
  setWorkspaceContext(a);
  assert.deepEqual(await studioBackup.list(), []); assert.deepEqual(await pageTemplateAssets.list(), []);
});

test('stale-tab fence is sent and a server session change locks further requests without retry', async t => {
  setWorkspaceContext('c'.repeat(32));
  const original = globalThis.fetch; t.after(() => { globalThis.fetch = original; });
  let calls = 0;
  globalThis.fetch = async (_url, options) => {
    calls++; assert.equal(options.headers.get('X-Jakal-Workspace'), 'c'.repeat(32));
    return new Response(JSON.stringify({ code: 'session_changed', detail: '세션 변경' }), { status: 401 });
  };
  await assert.rejects(requestJson('/api/documents', { method: 'POST' }), error => error.code === 'session_changed');
  await assert.rejects(requestJson('/api/documents'), error => error.code === 'session_changed');
  assert.equal(calls, 1);
});

test('session exit aborts document transfers already waiting for a response', async t => {
  setWorkspaceContext('d'.repeat(32));
  const original = globalThis.fetch; t.after(() => { globalThis.fetch = original; });
  let active;
  globalThis.fetch = (_url, options) => {
    active = options.signal;
    assert.equal(options.cache, 'no-store'); assert.equal(options.redirect, 'error'); assert.equal(options.credentials, 'same-origin');
    return new Promise((_, reject) => active.addEventListener('abort', () => reject(active.reason), { once: true }));
  };
  const pending = requestBytes('/api/documents/document/source?revision=0');
  lockWorkspace();
  await assert.rejects(pending, error => error.code === 'session_changed');
  assert.equal(active.aborted, true);
});

test('late document content is rejected if the workspace changes while reading its body', async t => {
  const a = 'e'.repeat(32), b = 'f'.repeat(32);
  setWorkspaceContext(a);
  const original = globalThis.fetch; t.after(() => { globalThis.fetch = original; });
  let finish, started;
  const reading = new Promise(resolve => { started = resolve; });
  globalThis.fetch = async () => ({ ok: true, headers: new Headers({ 'X-Jakal-Workspace': a }), json: () => {
    started(); return new Promise(resolve => { finish = resolve; });
  } });
  const pending = requestJson('/api/documents/document');
  await reading;
  setWorkspaceContext(b);
  finish({ text: 'old private contents' });
  await assert.rejects(pending, error => error.code === 'session_changed');
});

test('successful private responses must confirm the captured workspace', async t => {
  const original = globalThis.fetch; t.after(() => { globalThis.fetch = original; });
  const a = 'a'.repeat(32);
  for (const workspace of ['', 'b'.repeat(32)]) {
    setWorkspaceContext(a);
    let read = false;
    globalThis.fetch = async () => ({ ok: true, headers: new Headers(workspace ? { 'X-Jakal-Workspace': workspace } : {}), json: async () => { read = true; return { text: 'wrong workspace' }; } });
    await assert.rejects(requestJson('/api/documents/document'), error => error.code === 'session_changed');
    assert.equal(read, false);
  }
  setWorkspaceContext(a);
  globalThis.fetch = async () => new Response('{"revision":1}', { headers: { 'X-Jakal-Workspace': a } });
  assert.deepEqual(await requestJson('/api/documents/document'), { revision: 1 });
});

test('public template and health responses do not require a private workspace header', async t => {
  setWorkspaceContext('a'.repeat(32));
  const original = globalThis.fetch; t.after(() => { globalThis.fetch = original; });
  globalThis.fetch = async (_url, options) => {
    assert.equal(options.headers.has('X-Jakal-Workspace'), false);
    return new Response('{"public":true}');
  };
  for (const path of ['/api/templates', '/api/health?probe=1']) assert.deepEqual(await requestJson(path), { public: true });
});
