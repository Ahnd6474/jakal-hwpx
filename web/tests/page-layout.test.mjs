import assert from 'node:assert/strict';
import test from 'node:test';
import { defaultLayout, pagePresets, layoutProblem, savePagePreset, readPagePresets, deletePagePreset } from '../src/page-layout.ts';

function storage(t) {
  const old = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  const map = new Map();
  Object.defineProperty(globalThis, 'localStorage', { configurable: true, value: {
    getItem: key => map.get(key) ?? null,
    setItem: (key, value) => map.set(key, value),
  } });
  t.after(() => { if (old) Object.defineProperty(globalThis, 'localStorage', old); else delete globalThis.localStorage; });
  return map;
}

test('page templates contain geometry and recurring text without body paragraphs', () => {
  for (const preset of pagePresets) {
    assert.equal(layoutProblem(preset.layout), null);
    assert.equal('paragraphs' in preset.layout, false);
  }
  const split = pagePresets.find(item => item.id === 'two-column').layout;
  assert.equal(split.columns, 2); assert.equal(split.separator, true);
  assert.equal(split.header_text, null); assert.equal(split.footer_text, null);
});

test('invalid geometry and unsafe running text are rejected before a document request', () => {
  for (const change of [{ width_mm: NaN }, { height_mm: Infinity }, { columns: 4 }, { columns: 2, left_mm: 99, right_mm: 99 }, { header_mm: 25 }, { header_text: '줄\n바꿈' }, { footer_text: '\ud800' }]) assert.ok(layoutProblem({ ...defaultLayout, ...change }));
  assert.equal(layoutProblem({ ...defaultLayout, header_text: '김영주 😀' }), null);
});

test('saved presets survive re-reading and deletion leaves the original definitions untouched', t => {
  storage(t);
  const layout = { ...defaultLayout, columns: 2, separator: true, footer_text: '쪽 양식' };
  const [saved] = savePagePreset('  나의 2단 양식  ', layout);
  layout.columns = 1;
  assert.equal(readPagePresets()[0].layout.columns, 2);
  assert.equal(saved.name, '나의 2단 양식');
  assert.deepEqual(deletePagePreset(saved.id), []);
  assert.equal(pagePresets.find(item => item.id === 'standard').layout.columns, 1);
});

test('a full preset shelf and a failed storage write preserve its existing entries', t => {
  const map = storage(t);
  for (let n = 0; n < 20; n++) savePagePreset(`양식 ${n}`, defaultLayout);
  const before = [...map.values()][0];
  assert.throws(() => savePagePreset('초과', defaultLayout), /최대 20개/);
  assert.equal([...map.values()][0], before);
  deletePagePreset(readPagePresets()[0].id);
  const previous = [...map.values()][0];
  localStorage.setItem = () => { throw new Error('quota exceeded'); };
  assert.throws(() => savePagePreset('저장 실패', defaultLayout), /quota/);
  assert.equal([...map.values()][0], previous);
});
