import assert from 'node:assert/strict';
import test from 'node:test';
import { withFontResolutionCache, resolveCachedFont } from '../studio/font-cache.ts';

test('repeated font descriptors resolve once per paint without changing results', () => {
  const owner = {};
  let calls = 0;
  const resolve = () => { calls++; return 'bold 12px "대체 글꼴", serif'; };
  withFontResolutionCache(owner, () => {
    for (let i = 0; i < 600; i++) assert.equal(resolveCachedFont('bold 12px "원본 글꼴", serif', resolve), 'bold 12px "대체 글꼴", serif');
    withFontResolutionCache(owner, () => resolveCachedFont('bold 12px "원본 글꼴", serif', resolve));
  });
  assert.equal(calls, 1);
  withFontResolutionCache(owner, () => resolveCachedFont('bold 12px "원본 글꼴", serif', resolve));
  assert.equal(calls, 2);
});

test('font changes and different document owners never reuse stale resolutions', () => {
  const a = {}, b = {};
  withFontResolutionCache(a, () => {
    assert.equal(resolveCachedFont('12px font', () => 'A'), 'A');
    withFontResolutionCache(b, () => assert.equal(resolveCachedFont('12px font', () => 'B'), 'B'));
    assert.equal(resolveCachedFont('12px font', () => 'wrong'), 'A');
  });
  withFontResolutionCache(a, () => assert.equal(resolveCachedFont('12px font', () => 'loaded font'), 'loaded font'));
  assert.equal(resolveCachedFont('12px font', () => 'outside scope'), 'outside scope');
});

test('a failed nested paint restores the outer scope and releases caches', () => {
  const owner = {};
  withFontResolutionCache(owner, () => {
    resolveCachedFont('font', () => 'outer');
    assert.throws(() => withFontResolutionCache({}, () => { resolveCachedFont('font', () => 'inner'); throw Error('paint failed'); }));
    assert.equal(resolveCachedFont('font', () => 'wrong'), 'outer');
  });
  assert.throws(() => withFontResolutionCache(owner, () => { throw Error('paint failed'); }));
  assert.equal(resolveCachedFont('font', () => 'fresh'), 'fresh');
});

test('cache stays bounded for documents with many distinct font descriptors', () => {
  let calls = 0;
  withFontResolutionCache({}, () => {
    for (let i = 0; i < 1200; i++) resolveCachedFont(`${i}px font`, () => { calls++; return 'value'; });
    resolveCachedFont('0px font', () => { calls++; return 'value'; });
    resolveCachedFont('1199px font', () => { calls++; return 'value'; });
  });
  assert.equal(calls, 1201);
});
