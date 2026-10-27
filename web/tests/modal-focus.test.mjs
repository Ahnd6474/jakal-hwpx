import assert from 'node:assert/strict';
import test from 'node:test';
import { JSDOM } from 'jsdom';
import { trapModalFocus } from '../src/modal-focus.ts';

function fixture(t, html) {
  const dom = new JSDOM(`<button id="trigger">Open</button><section id="dialog">${html}</section>`);
  const old = {};
  for (const key of ['window', 'document', 'Node']) { old[key] = globalThis[key]; globalThis[key] = key === 'window' ? dom.window : dom.window[key]; }
  const releases = [];
  const trap = (element, close) => { const release = trapModalFocus(element, close); releases.push(release); return release; };
  t.after(() => { for (const release of releases.reverse()) release(); dom.window.close(); for (const key in old) { if (old[key] === undefined) delete globalThis[key]; else globalThis[key] = old[key]; } });
  document.querySelector('#trigger').focus();
  const dialog = document.querySelector('#dialog');
  const key = (name, shiftKey = false) => document.activeElement.dispatchEvent(new dom.window.KeyboardEvent('keydown', { key: name, shiftKey, bubbles: true, cancelable: true }));
  return { dialog, key, trap };
}

test('modal cycles only visible enabled controls, includes links and restores focus', t => {
  const { dialog, key, trap } = fixture(t, '<div hidden><button>Hidden</button></div><button disabled tabindex="0">Disabled</button><fieldset disabled><input></fieldset><a href="#a" id="first">Download</a><div style="display:none"><input></div><input type="hidden"><button id="last">Close</button>');
  let closed = 0;
  const release = trap(dialog, () => closed++);
  assert.equal(document.activeElement.id, 'first');
  key('Tab', true); assert.equal(document.activeElement.id, 'last');
  key('Tab'); assert.equal(document.activeElement.id, 'first');
  document.querySelector('#trigger').focus(); assert.equal(document.activeElement.id, 'first');
  key('Escape'); assert.equal(closed, 1);
  release();
  assert.equal(document.activeElement.id, 'trigger');
  assert.equal(document.body.style.overflow, '');
});

test('modal without enabled controls remains focusable while busy', t => {
  const { dialog, key, trap } = fixture(t, '<button disabled>Saving</button>');
  trap(dialog, () => {});
  assert.equal(document.activeElement, dialog);
  key('Tab'); assert.equal(document.activeElement, dialog);
  key('Tab', true); assert.equal(document.activeElement, dialog);
});

test('nested modal handles Escape once and preserves the first scroll lock', t => {
  const { dialog, key, trap } = fixture(t, '<button id="first">First</button><section id="nested"><button id="second">Second</button></section>');
  let outer = 0, inner = 0;
  const releaseOuter = trap(dialog, () => outer++);
  const releaseInner = trap(document.querySelector('#nested'), () => inner++);
  key('Escape'); assert.equal(outer, 0); assert.equal(inner, 1);
  releaseInner(); assert.equal(document.body.style.overflow, 'hidden');
  key('Escape'); assert.equal(outer, 1);
  releaseOuter(); assert.equal(document.body.style.overflow, '');
});
