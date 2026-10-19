import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { JSDOM } from 'jsdom';
import { HwpDocument, initSync } from '@rhwp/core';
import { renderedText } from '../src/viewer-text.ts';

const dom = new JSDOM('');
globalThis.window = dom.window;
globalThis.DOMParser = dom.window.DOMParser;
globalThis.XMLSerializer = dom.window.XMLSerializer;
const { safeSvg, cropSvg, highlightSvg } = await import('../src/viewer-svg.ts');

test('rendered SVG keeps text, mathematics and clip paths while removing active content', () => {
  const result = safeSvg('<svg xmlns="http://www.w3.org/2000/svg"><defs><clipPath id="cell-1"><rect width="10" height="10"/></clipPath></defs><g clip-path="url(#cell-1)"><text font-family="바탕" x="3" y="5">한글 &amp; 수식</text></g><script>alert(1)</script><foreignObject><iframe src="https://outside.test"/></foreignObject><image href="https://outside.test/private" onload="alert(1)"/><text style="fill:url(https://outside.test)">x</text></svg>');
  const xml = new DOMParser().parseFromString(result, 'image/svg+xml');
  assert.equal(xml.querySelector('text').textContent, '한글 & 수식');
  assert.equal(xml.querySelector('g').getAttribute('clip-path'), 'url(#cell-1)');
  assert.equal(xml.querySelector('script,foreignObject,iframe'), null);
  assert.equal(xml.querySelector('[onload],[style]'), null);
  assert.equal(xml.querySelector('image')?.hasAttribute('href') ?? false, false);
  assert.ok(!result.includes('outside.test'));
});

test('self-contained raster images remain available in graph previews', () => {
  const input = '<svg xmlns="http://www.w3.org/2000/svg"><image href="data:image/png;base64,aGVsbG8=" width="10" height="10"/></svg>';
  assert.match(safeSvg(input), /data:image\/png;base64,aGVsbG8=/);
});

test('object crop preserves vector content, changes only the view box', () => {
  const input = safeSvg('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 1100"><path d="M1 1L10 10"/><text x="20" y="30">그래프</text></svg>');
  const output = cropSvg(input, { x: 10, y: 20, w: 100, h: 80 });
  const xml = new DOMParser().parseFromString(output, 'image/svg+xml');
  assert.equal(xml.documentElement.getAttribute('viewBox'), '5 15 110 90');
  assert.equal(xml.querySelector('text').textContent, '그래프');
  assert.ok(xml.querySelector('path'));
});

test('search highlighting spans glyph-sized SVG text nodes without injecting search markup', () => {
  const svg = safeSvg('<svg xmlns="http://www.w3.org/2000/svg"><text>원</text><text>점</text><text>에서</text></svg>');
  const highlighted = highlightSvg(svg, '원점');
  const xml = new DOMParser().parseFromString(highlighted, 'image/svg+xml');
  assert.equal(xml.querySelectorAll('text[stroke]').length, 2);
  assert.ok(!highlightSvg(svg, '<script>alert(1)</script>').includes('<script>'));
});

test('search text follows painted page content and decodes XML entities', () => {
  const page = '<svg><metadata>다른 페이지의 곡선</metadata><text>원</text><text>점</text><text><tspan> A&amp;B &#xAC00; &#45208;</tspan></text><path aria-label="포함하지 않을 텍스트"/></svg>';
  assert.equal(renderedText(page), '원점A&B가나');
  assert.equal(renderedText(page).includes('곡선'), false);
});

// Run the actual bundled renderer against repository-owned files, independent
// of browser automation. Width approximation is used only by this Node test.
globalThis.measureTextWidth = (_font, text) => Array.from(text).reduce((sum, c) => sum + (c.codePointAt(0) > 255 ? 12 : 6), 0);
initSync({ module: fs.readFileSync(new URL('../node_modules/@rhwp/core/rhwp_bg.wasm', import.meta.url)) });

for (const relative of ['../../examples/samples/hwp/기출킬러양식.hwp', '../../examples/samples/hwpx/math_template_2_formula_fixed.hwpx']) {
  test(`real renderer opens ${relative.split('/').at(-1)} without changing source bytes`, () => {
    const path = new URL(relative, import.meta.url);
    const raw = fs.readFileSync(path);
    const doc = new HwpDocument(raw);
    try {
      assert.ok(doc.pageCount() > 0);
      const svg = safeSvg(doc.renderPageSvg(0));
      const xml = new DOMParser().parseFromString(svg, 'image/svg+xml');
      assert.ok(xml.querySelectorAll('text,path').length > 10);
      const page = JSON.parse(doc.getPageInfo(0));
      assert.ok(page.width > 100 && page.height > 100);
      const controls = JSON.parse(doc.getPageControlLayout(0)).controls;
      assert.ok(controls.some(c => ['equation', 'table'].includes(c.type)));
      assert.deepEqual(fs.readFileSync(path), raw);
    } finally { doc.free(); }
  });
}
