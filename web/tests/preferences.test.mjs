import assert from 'node:assert/strict';
import test from 'node:test';
import { lastDocument } from '../src/preferences.ts';

function storage(t, descriptor, name = 'localStorage') {
  const previous = Object.getOwnPropertyDescriptor(globalThis, name);
  Object.defineProperty(globalThis, name, { ...descriptor, configurable: true });
  t.after(() => {
    if (previous) Object.defineProperty(globalThis, name, previous);
    else delete globalThis[name];
  });
}

test('last document can be restored and cleared', t => {
  const entries = new Map();
  storage(t, { value: {
    getItem: key => entries.get(key) ?? null,
    setItem: (key, value) => entries.set(key, value),
    removeItem: key => entries.delete(key),
  }, configurable: true });
  assert.equal(lastDocument.read(), null);
  lastDocument.write('document-id');
  assert.equal(lastDocument.read(), 'document-id');
  lastDocument.write(null);
  assert.equal(lastDocument.read(), null);
});

test('blocked browser storage does not interrupt document actions', t => {
  storage(t, { get() { throw new DOMException('Blocked', 'SecurityError'); } });
  assert.equal(lastDocument.read(), null);
  assert.doesNotThrow(() => lastDocument.write('document-id'));
  assert.doesNotThrow(() => lastDocument.write(null));
});

test('full storage still allows reading the existing preference', t => {
  storage(t, { value: {
    getItem: () => 'previous-document',
    setItem() { throw new DOMException('Full', 'QuotaExceededError'); },
    removeItem() { throw new Error('Unavailable'); },
  }, configurable: true });
  assert.doesNotThrow(() => lastDocument.write('new-document'));
  assert.equal(lastDocument.read(), 'previous-document');
  assert.doesNotThrow(() => lastDocument.write(null));
});

test('the current tab keeps its own document when another tab updates shared storage', t => {
  storage(t, { value: { getItem: () => 'other-tab-document' } });
  storage(t, { value: { getItem: () => 'this-tab-document' } }, 'sessionStorage');
  assert.equal(lastDocument.read(), 'this-tab-document');
});

test('the tab preference remains usable when persistent storage is blocked', t => {
  const entries = new Map();
  storage(t, { get() { throw new Error('Blocked'); } });
  storage(t, { value: { getItem: key => entries.get(key) ?? null, setItem: (key, value) => entries.set(key, value), removeItem: key => entries.delete(key) } }, 'sessionStorage');
  lastDocument.write('document-id');
  assert.equal(lastDocument.read(), 'document-id');
  lastDocument.write(null);
  assert.equal(lastDocument.read(), null);
});
