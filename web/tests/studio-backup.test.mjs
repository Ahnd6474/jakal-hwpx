import assert from 'node:assert/strict';
import test, { afterEach } from 'node:test';
import { indexedDB } from 'fake-indexeddb';
import { studioBackup } from '../src/studio-backup.ts';

globalThis.indexedDB = indexedDB;
const name = 'jakal-document-editor';
afterEach(() => new Promise((resolve, reject) => { const request = indexedDB.deleteDatabase(name); request.onsuccess = resolve; request.onerror = () => reject(request.error); }));

test('draft metadata and bytes commit together, list newest first and clear together', async () => {
  await studioBackup.write({ id: 'older', revision: 2, saved: 100, name: '한글.hwpx', bytes: new Uint8Array([80, 75, 1]) });
  await studioBackup.write({ id: 'newer', revision: 5, saved: 200, name: '검토.hwpx', bytes: new Uint8Array([80, 75, 2, 3]) });
  const list = await studioBackup.list();
  assert.deepEqual(list.map(item => item.id), ['newer', 'older']);
  assert.equal(list[0].size, 4);
  assert.equal(Object.hasOwn(list[0], 'bytes'), false);
  assert.deepEqual((await studioBackup.read('newer')).bytes, new Uint8Array([80, 75, 2, 3]));
  await studioBackup.clear('newer');
  assert.equal(await studioBackup.read('newer'), undefined);
  assert.deepEqual((await studioBackup.list()).map(item => item.id), ['older']);
});

test('upgrading the previous database keeps old bytes and makes uncatalogued drafts discoverable', async () => {
  await new Promise((resolve, reject) => {
    const request = indexedDB.open(name, 1);
    request.onupgradeneeded = () => request.result.createObjectStore('drafts', { keyPath: 'id' });
    request.onerror = () => reject(request.error);
    request.onsuccess = () => {
      const db = request.result, tx = db.transaction('drafts', 'readwrite');
      tx.objectStore('drafts').put({ id: 'legacy', revision: 3, saved: 1, bytes: new Uint8Array([1, 2, 3]) });
      tx.oncomplete = () => { db.close(); resolve(); };
    };
  });
  assert.deepEqual(await studioBackup.list(), [{ id: 'legacy', revision: -1, saved: 0, size: 0 }]);
  assert.deepEqual((await studioBackup.read('legacy')).bytes, new Uint8Array([1, 2, 3]));
  await studioBackup.write({ id: 'legacy', revision: 4, saved: 2, name: '복원.hwpx', bytes: new Uint8Array([4]) });
  assert.equal((await studioBackup.list())[0].revision, 4);
});

test('a cloning failure aborts both stores and keeps the previous recovery bytes', async () => {
  await studioBackup.write({ id: 'safe', revision: 0, saved: 1, bytes: new Uint8Array([7]) });
  await assert.rejects(studioBackup.write({ id: 'safe', revision: 1, saved: 2, bytes: new Uint8Array([8]), name: () => {} }));
  assert.equal((await studioBackup.list())[0].revision, 0);
  assert.deepEqual((await studioBackup.read('safe')).bytes, new Uint8Array([7]));
});
