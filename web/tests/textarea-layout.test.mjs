import assert from 'node:assert/strict';
import test from 'node:test';
import { fitTextareas } from '../src/textarea-layout.ts';

test('many paragraphs reset all heights before measuring, then apply measured heights', () => {
  const events = [];
  const fields = Array.from({ length: 300 }, (_, id) => ({
    style: { set height(value) { events.push(['write', id, value]); } },
    get scrollHeight() { events.push(['read', id]); return 20 + id; },
  }));
  fitTextareas(fields.values());
  assert.equal(events.length, 900);
  assert.ok(events.slice(0, 300).every(event => event[0] === 'write' && event[2] === '0px'));
  assert.ok(events.slice(300, 600).every(event => event[0] === 'read'));
  assert.deepEqual(events.slice(600), fields.map((_, id) => ['write', id, `${20 + id}px`]));
});

test('re-measuring after wrapping changes can grow and shrink the same field', () => {
  let contentHeight = 500;
  const field = { style: { height: '900px' }, get scrollHeight() {
    assert.equal(this.style.height, '0px');
    return contentHeight;
  } };
  fitTextareas([field]);
  assert.equal(field.style.height, '500px');
  contentHeight = 30;
  fitTextareas([field]);
  assert.equal(field.style.height, '30px');
  fitTextareas([]);
});
