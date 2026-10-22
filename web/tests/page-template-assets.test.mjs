import assert from 'node:assert/strict';
import test, { afterEach } from 'node:test';
import { indexedDB } from 'fake-indexeddb';
import { pageTemplateAssets } from '../src/page-template-assets.ts';
import { defaultLayout, backgroundProblem } from '../src/page-layout.ts';

globalThis.indexedDB = indexedDB;
const page = { width_mm: 210, height_mm: 297, image: 'data:image/jpeg;base64,YWJjZGVmZ2hpamtsbW5vcA==' };
afterEach(() => new Promise((resolve, reject) => { const request = indexedDB.deleteDatabase('jakal-page-template-assets:locked'); request.onsuccess = resolve; request.onerror = () => reject(request.error); }));

test('background templates store metadata without cloning page images on list and preserve independent copies', async () => {
  const input = [{ ...page }];
  const saved = await pageTemplateAssets.save('사진 양식', defaultLayout, input);
  input[0].image = 'changed';
  const listing = await pageTemplateAssets.list();
  assert.equal(listing.length, 1); assert.equal(listing[0].page_count, 1);
  assert.equal('background_pages' in listing[0], false);
  assert.deepEqual((await pageTemplateAssets.read(saved.id)).background_pages, [page]);
  await pageTemplateAssets.remove(saved.id);
  assert.equal(await pageTemplateAssets.read(saved.id), undefined);
  assert.deepEqual(await pageTemplateAssets.list(), []);
});

test('simultaneous saves respect the durable preset quota without orphan images', async () => {
  for (let index = 0; index < 19; index++) await pageTemplateAssets.save(`양식 ${index}`, defaultLayout, [page]);
  const result = await Promise.allSettled([pageTemplateAssets.save('마지막 A', defaultLayout, [page]), pageTemplateAssets.save('마지막 B', defaultLayout, [page])]);
  assert.equal(result.filter(item => item.status === 'fulfilled').length, 1);
  assert.equal((await pageTemplateAssets.list()).length, 20);
});

test('active content, wrong geometry and excessive page collections never enter the shelf', async () => {
  assert.ok(backgroundProblem([{ ...page, image: 'data:image/svg+xml,<svg onload="alert(1)"/>' }]));
  assert.ok(backgroundProblem([{ ...page, width_mm: Infinity }]));
  assert.ok(backgroundProblem(Array.from({ length: 11 }, () => page)));
  await assert.rejects(pageTemplateAssets.save('잘못된 양식', defaultLayout, [{ ...page, image: 'https://outside.test/image.jpg' }]));
  assert.deepEqual(await pageTemplateAssets.list(), []);
});
