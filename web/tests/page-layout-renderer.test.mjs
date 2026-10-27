import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { JSDOM } from 'jsdom';
import { HwpDocument, initSync } from '@rhwp/core';
import { renderedText } from '../src/viewer-text.ts';

const root = path.resolve('..');
const python = path.join(root, '.venv', process.platform === 'win32' ? 'Scripts/python.exe' : 'bin/python');
globalThis.measureTextWidth = (_font, text) => Array.from(text).reduce((sum, char) => sum + (char.codePointAt(0) > 255 ? 12 : 6), 0);
initSync({ module: fs.readFileSync(new URL('../node_modules/@rhwp/core/rhwp_bg.wasm', import.meta.url)) });

test('real renderer displays columns, divider, header and footer after native text edit and HWPX export', { skip: !fs.existsSync(python) }, t => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'jakal-layout-'));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'layout.hwpx');
  execFileSync(python, ['-c', [
    'import sys', 'from pathlib import Path', 'from jakal_hwpx import HwpxDocument',
    'from server.schemas import PageLayout', 'from server.page_layout import apply_layout',
    'd=HwpxDocument.blank()', 'd.set_paragraph_text(0,0,"BODY")',
    '[d.append_paragraph("Document body paragraph " + str(n)) for n in range(10)]',
    'apply_layout(d,PageLayout(columns=2,separator=True,header_text="HEADER",footer_text="FOOTER",page_numbers="bottom-right"))',
    'Path(sys.argv[1]).write_bytes(d.compile(validate=False,auto_repair=False))',
    'other=HwpxDocument.blank()',
    'apply_layout(other,PageLayout(width_mm=297,height_mm=210,columns=3,separator=True))',
    'Path(sys.argv[1]+".landscape.hwpx").write_bytes(other.compile(validate=False,auto_repair=False))',
  ].join('\n'), file], { cwd: root });
  const dom = new JSDOM(); t.after(() => dom.window.close());
  function check(document) {
    assert.equal(JSON.parse(document.getPageInfo(0)).columns.length, 2);
    const svg = document.renderPageSvg(0);
    const xml = new dom.window.DOMParser().parseFromString(svg, 'image/svg+xml');
    assert.match(renderedText(svg), /HEADER/); assert.match(renderedText(svg), /FOOTER/);
    assert.ok([...xml.querySelectorAll('line')].some(line => line.getAttribute('x1') === line.getAttribute('x2') && Number(line.getAttribute('y2')) > Number(line.getAttribute('y1')) + 10));
  }
  const doc = new HwpDocument(fs.readFileSync(file));
  try {
    check(doc); doc.insertText(0, 1, 0, 'EDITED ');
    const roundtrip = new HwpDocument(doc.exportHwpx());
    try { check(roundtrip); assert.match(renderedText(roundtrip.renderPageSvg(0)), /EDITED/); }
    finally { roundtrip.free(); }
  } finally { doc.free(); }
  const landscape = new HwpDocument(fs.readFileSync(file + '.landscape.hwpx'));
  try {
    const info = JSON.parse(landscape.getPageInfo(0));
    assert.ok(info.width > info.height); assert.equal(info.columns.length, 3);
  } finally { landscape.free(); }
});

test('page backgrounds cover physical pages and survive body edits and native export', { skip: !fs.existsSync(python) }, t => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'jakal-background-'));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'background.hwpx');
  execFileSync(python, ['-c', [
    'import sys,base64,io', 'from pathlib import Path', 'from PIL import Image',
    'from server.schemas import PageLayout,BackgroundPage',
    'from server.page_background import make_background_document',
    'output=io.BytesIO()', 'Image.new("RGB",(210,297),"white").save(output,format="JPEG")',
    'page=BackgroundPage(width_mm=210,height_mm=297,image="data:image/jpeg;base64,"+base64.b64encode(output.getvalue()).decode())',
    'd=make_background_document(PageLayout(left_mm=25,top_mm=25,header_mm=7,header_text="HEADER"),[page,page])',
    'd.set_paragraph_text(0,0,"WRITABLE BODY")',
    'Path(sys.argv[1]).write_bytes(d.compile(validate=False,auto_repair=False))',
    'from server.page_layout import apply_layout',
    'apply_layout(d,PageLayout(width_mm=420,height_mm=594,left_mm=35,top_mm=30,header_mm=10,header_text="NEW HEADER"))',
    'Path(sys.argv[1]+".resized.hwpx").write_bytes(d.compile(validate=False,auto_repair=False))',
  ].join('\n'), file], { cwd: root });
  const dom = new JSDOM(); t.after(() => dom.window.close());
  function check(document) {
    assert.equal(document.pageCount(), 2);
    for (let page = 0; page < 2; page++) {
      const xml = new dom.window.DOMParser().parseFromString(document.renderPageSvg(page), 'image/svg+xml');
      const images = [...xml.querySelectorAll('image')];
      assert.equal(images.length, 1);
      const image = images[0], info = JSON.parse(document.getPageInfo(page));
      assert.ok(Math.abs(Number(image.getAttribute('x'))) < 0.1, image.outerHTML.slice(0, 100));
      assert.ok(Math.abs(Number(image.getAttribute('y'))) < 0.1, image.outerHTML.slice(0, 100));
      assert.ok(Math.abs(Number(image.getAttribute('width')) - info.width) < 0.1);
      assert.ok(Math.abs(Number(image.getAttribute('height')) - info.height) < 0.1);
      if (page === 0) {
        assert.match(renderedText(document.renderPageSvg(page)), /WRITABLE/);
        const body = [...xml.querySelectorAll('text')].find(node => node.textContent.includes('W'));
        assert.ok(body, 'body text renders above background');
        assert.ok(image.compareDocumentPosition(body) & dom.window.Node.DOCUMENT_POSITION_FOLLOWING, 'background paints before body text');
      }
    }
  }
  const doc = new HwpDocument(fs.readFileSync(file));
  try {
    check(doc); doc.insertText(0, 0, 0, 'EDITED ');
    const roundtrip = new HwpDocument(doc.exportHwpx());
    try { check(roundtrip); assert.match(renderedText(roundtrip.renderPageSvg(0)), /EDITED/); }
    finally { roundtrip.free(); }
  } finally { doc.free(); }
  const resized = new HwpDocument(fs.readFileSync(file + '.resized.hwpx'));
  try { check(resized); assert.match(renderedText(resized.renderPageSvg(0)), /NEW\s*HEADER/); }
  finally { resized.free(); }
});
