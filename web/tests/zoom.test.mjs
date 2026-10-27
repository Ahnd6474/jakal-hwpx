import test from 'node:test';
import assert from 'node:assert/strict';
import { pageScale, stepZoom } from '../src/viewer-zoom.ts';

test('width fit follows narrow and wide pages without forcing horizontal overflow', () => {
  const page = { width: 794, height: 1123 };
  for (const width of [280, 340, 720, 1200]) {
    assert.ok(Math.abs(pageScale('fit', { width, height: 400 }, page) * page.width - width) < .001);
  }
});
test('whole page fit respects both available dimensions in landscape', () => {
  const page = { width: 794, height: 1123 };
  const scale = pageScale('page', { width: 700, height: 200 }, page);
  assert.ok(page.width * scale <= 700 && page.height * scale <= 200);
});
test('zoom buttons always move toward the requested direction and stop at limits', () => {
  for (const percent of [10, 37.8, 50, 99.8, 100, 125, 399, 400]) {
    assert.ok(stepZoom(percent, 1) > percent || percent === 400);
    assert.ok(stepZoom(percent, -1) < percent || percent === 10);
  }
  assert.equal(pageScale(999, { width: 300, height: 400 }, { width: 794, height: 1123 }), 4);
});
