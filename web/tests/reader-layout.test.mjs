import test from 'node:test';
import assert from 'node:assert/strict';
import { sheetLayout, pageAtOffset, pageWindow, readingPage } from '../src/reader-layout.ts';

test('continuous pages retain their own portrait and landscape proportions', () => {
  const pages = [{ width: 800, height: 1100 }, { width: 1100, height: 800 }, { width: 800, height: 1100 }];
  const layout = sheetLayout(pages, 'fit', { width: 360, height: 600 });
  assert.equal(layout[0].width, 360);
  assert.equal(layout[1].width, 360);
  assert.equal(layout[1].top, layout[0].height);
  assert.equal(layout[2].top, layout[0].height + layout[1].height);
  assert.ok(layout[1].paperHeight < layout[0].paperHeight);
});

test('scroll and direct navigation agree at each page boundary', () => {
  const layout = sheetLayout(Array.from({ length: 500 }, () => ({ width: 800, height: 1100 })), 150, { width: 720, height: 600 });
  assert.equal(pageAtOffset(layout, -500), 0);
  for (const page of [0, 1, 18, 250, 499]) {
    assert.equal(pageAtOffset(layout, layout[page].top), page);
    assert.equal(pageAtOffset(layout, layout[page].top + layout[page].height - .01), page);
  }
  assert.equal(pageAtOffset(layout, 1e12), 499);
});

test('live pages include visible pages and their neighbors as the user scrolls', () => {
  const layout = sheetLayout(Array.from({ length: 30 }, () => ({ width: 800, height: 1100 })), 100, { width: 720, height: 600 });
  assert.deepEqual(pageWindow(layout, 0, 600), { start: 0, end: 1 });
  assert.deepEqual(pageWindow(layout, layout[8].top + 1000, 600), { start: 7, end: 10 });
  assert.deepEqual(pageWindow(layout, layout[29].top, 600), { start: 28, end: 29 });
});

test('hundreds of tiny pages do not mount hundreds of SVG trees', () => {
  const layout = sheetLayout(Array.from({ length: 500 }, () => ({ width: 800, height: 1 })), 10, { width: 360, height: 2000 });
  const range = pageWindow(layout, layout[200].top, 2000);
  assert.ok(range.start <= 200 && range.end >= 200);
  assert.ok(range.end - range.start + 1 <= 24);
});

test('current page follows the reading area when the next page occupies most of the screen', () => {
  const layout = sheetLayout(Array.from({ length: 19 }, () => ({ width: 794, height: 1123 })), 'fit', { width: 290, height: 693 });
  assert.equal(readingPage(layout, layout[17].top, 693), 17);
  const nearBottom = layout[18].top - 200;
  assert.equal(readingPage(layout, nearBottom, 693), 18);
});
