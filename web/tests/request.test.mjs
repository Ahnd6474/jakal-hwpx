import assert from 'node:assert/strict';
import test from 'node:test';
import { requestJson, requestBytes } from '../src/request.ts';

function fetchMock(t, implementation) {
  const previous = globalThis.fetch;
  globalThis.fetch = implementation;
  t.after(() => { globalThis.fetch = previous; });
}

test('timeouts abort a stalled mutation once without retrying it', async t => {
  let count = 0, signal;
  fetchMock(t, (_url, init) => {
    count++; signal = init.signal;
    return new Promise((_, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true }));
  });
  await assert.rejects(requestJson('/api/save', { method: 'POST', body: '{}', timeoutMs: 15 }), /처리 여부를 확인/);
  assert.equal(count, 1);
  assert.equal(signal.aborted, true);
});

test('deadline includes a stalled response body, including document bytes', async t => {
  fetchMock(t, (_url, { signal }) => Promise.resolve({ ok: true, arrayBuffer: () => new Promise((_, reject) => {
    signal.addEventListener('abort', () => reject(signal.reason), { once: true });
  }) }));
  await assert.rejects(requestBytes('/api/source', { timeoutMs: 15 }), /대기 시간이/);
});

test('unmount cancellation preserves AbortError and does not issue an already cancelled request', async t => {
  let count = 0;
  fetchMock(t, () => { count++; throw new Error('must not fetch'); });
  const abort = new AbortController(); abort.abort();
  await assert.rejects(requestJson('/api/doc', { signal: abort.signal }), { name: 'AbortError' });
  assert.equal(count, 0);
});

test('caller cancellation during fetch is forwarded without a connection error', async t => {
  let requestSignal;
  fetchMock(t, (_url, { signal }) => new Promise((_, reject) => {
    requestSignal = signal;
    signal.addEventListener('abort', () => reject(signal.reason), { once: true });
  }));
  const abort = new AbortController();
  const pending = requestJson('/api/doc', { signal: abort.signal });
  abort.abort();
  await assert.rejects(pending, { name: 'AbortError' });
  assert.equal(requestSignal.aborted, true);
});

test('API conflicts keep their specific message; network, HTML and server failures are readable', async t => {
  fetchMock(t, async () => new Response(JSON.stringify({ detail: '최신 버전을 다시 열어 주세요.' }), { status: 409 }));
  await assert.rejects(requestJson('/api/doc'), /최신 버전/);
  globalThis.fetch = async () => { throw new TypeError('Failed to fetch'); };
  await assert.rejects(requestJson('/api/doc'), /네트워크와 서버/);
  globalThis.fetch = async () => new Response('<html>proxy error</html>');
  await assert.rejects(requestJson('/api/doc'), /응답을 읽을 수/);
  globalThis.fetch = async () => new Response('<html>unavailable</html>', { status: 503 });
  await assert.rejects(requestJson('/api/doc'), /서버에서 요청을 처리하지/);
  globalThis.fetch = async () => new Response('null', { status: 422 });
  await assert.rejects(requestJson('/api/doc'), /입력 내용을 확인/);
});

test('successful JSON and byte requests preserve their payloads', async t => {
  fetchMock(t, async () => new Response('{"revision":3}'));
  assert.deepEqual(await requestJson('/api/doc'), { revision: 3 });
  globalThis.fetch = async () => new Response(new Uint8Array([80, 75, 3, 4]));
  assert.deepEqual(new Uint8Array(await requestBytes('/api/source')), new Uint8Array([80, 75, 3, 4]));
});
