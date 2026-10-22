import assert from 'node:assert/strict';
import test from 'node:test';
import { DocumentCommands, DocumentCommandError, pause } from '../src/document-commands.ts';
import { ApiError } from '../src/request.ts';

const base = { id: 'document', revision: 3, settings_version: 2 };
const operation = { type: 'paragraph', section: 0, index: 0, text: '바꾼 본문' };
const command = { kind: 'edit', operations: [operation] };
const receipt = { job_id: 'job', document_id: base.id, base_revision: 3, kind: 'edit' };
const job = { id: 'job', document_id: base.id, kind: 'edit', status: 'completed', error: null, result: { revision: 4 } };
const fresh = { ...base, revision: 4, status: 'ready', validated: null };
const immediate = async (_ms, signal) => signal?.throwIfAborted();

test('page layout shares versioned command receipts and carries the selected section', async () => {
  const { defaultLayout } = await import('../src/page-layout.ts');
  const client = new DocumentCommands(async (url, options) => {
    if (options?.method === 'POST') {
      assert.equal(url, '/api/documents/document/page-layout');
      assert.deepEqual(JSON.parse(options.body), { revision: 3, settings_version: 2, layout: defaultLayout, section: 1 });
      return { ...receipt, kind: 'layout' };
    }
    return url.includes('/jobs/') ? { ...job, kind: 'layout' } : fresh;
  }, immediate);
  assert.equal(await client.execute(base, { kind: 'layout', layout: defaultLayout, section: 1 }), fresh);
});

test('lost command response retries the identical payload and key, then verifies job and saved state', async () => {
  const writes = [];
  let reconnects = 0;
  const client = new DocumentCommands(async (url, options) => {
    if (options?.method === 'POST') {
      writes.push(options);
      if (writes.length === 1) throw new ApiError('connection lost', 'network_error');
      return receipt;
    }
    if (url.includes('/jobs/')) return job;
    return fresh;
  }, immediate);
  const result = await client.execute(base, command, { onReconnect: () => reconnects++ });
  assert.equal(result, fresh);
  assert.equal(writes.length, 2);
  assert.deepEqual(writes[0], writes[1]);
  assert.match(writes[0].headers['Idempotency-Key'], /^[a-f0-9]{32}$/);
  assert.deepEqual(JSON.parse(writes[0].body), { revision: 3, settings_version: 2, operations: [operation] });
  assert.equal(reconnects, 1);
});

test('a later explicit Save after unresolved connection failure uses the same identity', async () => {
  const keys = [];
  let connected = false;
  const client = new DocumentCommands(async (url, options) => {
    if (options?.method === 'POST') {
      keys.push(options.headers['Idempotency-Key']);
      if (!connected) throw new ApiError('offline', 'request_timeout');
      return receipt;
    }
    return url.includes('/jobs/') ? job : fresh;
  }, immediate);
  await assert.rejects(client.execute(base, command), { code: 'request_timeout' });
  connected = true;
  assert.equal(await client.execute(base, command), fresh);
  assert.equal(keys.length, 3);
  assert.equal(new Set(keys).size, 1);
});

test('reconnect during polling only reads the original job; failures and conflicts are not retried', async () => {
  let writes = 0, reads = 0;
  const client = new DocumentCommands(async (url, options) => {
    if (options?.method === 'POST') { writes++; return receipt; }
    if (url.includes('/jobs/')) {
      if (++reads === 1) throw new ApiError('offline', 'network_error');
      return job;
    }
    return fresh;
  }, immediate);
  assert.equal(await client.execute(base, command), fresh);
  assert.equal(writes, 1); assert.equal(reads, 2);
  let calls = 0;
  const conflict = new DocumentCommands(async () => {
    calls++; throw new ApiError('stale settings', 'settings_conflict', 409);
  }, immediate);
  await assert.rejects(conflict.execute(base, command), { code: 'settings_conflict', status: 409 });
  assert.equal(calls, 1);
});

test('mismatched receipt, other document job, bad result and advanced revisions never return success', async () => {
  for (const change of [
    { receipt: { ...receipt, document_id: 'other' }, code: 'invalid_receipt' },
    { receipt: { ...receipt, kind: 'restore' }, code: 'invalid_receipt' },
    { job: { ...job, document_id: 'other' }, code: 'invalid_job' },
    { job: { ...job, kind: 'restore' }, code: 'invalid_job' },
    { job: { ...job, result: { revision: 5 } }, code: 'invalid_result' },
    { fresh: { ...fresh, revision: 5 }, code: 'revision_conflict' },
    { fresh: { ...fresh, settings_version: 3 }, code: 'revision_conflict' },
    { fresh: { ...fresh, id: 'other' }, code: 'revision_conflict' },
  ]) {
    const client = new DocumentCommands(async (url, options) => options?.method === 'POST'
      ? change.receipt ?? receipt : url.includes('/jobs/') ? change.job ?? job : change.fresh ?? fresh, immediate);
    await assert.rejects(client.execute(base, command), error => error instanceof DocumentCommandError && error.code === change.code);
  }
});

test('a worker rejection permits a new attempt while retaining its specific message', async () => {
  const keys = [];
  const client = new DocumentCommands(async (_url, options) => {
    if (options?.method === 'POST') { keys.push(options.headers['Idempotency-Key']); return receipt; }
    return { ...job, status: 'failed', error: '보호 문구를 유지해 주세요.', result: null };
  }, immediate);
  for (let n = 0; n < 2; n++) await assert.rejects(client.execute(base, command), { code: 'job_failed', message: '보호 문구를 유지해 주세요.', jobId: 'job' });
  assert.notEqual(keys[0], keys[1]);
});

test('cancellation before submission and during polling keeps AbortError and never resubmits', async () => {
  let count = 0;
  const cancelled = new AbortController(); cancelled.abort();
  const client = new DocumentCommands(async () => { count++; }, immediate);
  await assert.rejects(client.execute(base, command, { signal: cancelled.signal }), { name: 'AbortError' });
  assert.equal(count, 0);
  const abort = new AbortController();
  const polling = new DocumentCommands(async (_url, options) => {
    count++;
    if (options?.method === 'POST') return receipt;
    abort.abort(); throw abort.signal.reason;
  }, immediate);
  await assert.rejects(polling.execute(base, command, { signal: abort.signal }), { name: 'AbortError' });
  assert.equal(count, 2);
  const waiting = new AbortController();
  const delay = pause(20_000, waiting.signal); waiting.abort();
  await assert.rejects(delay, { name: 'AbortError' });
});

test('native edits snapshot their bytes and send the revision and settings fences', async () => {
  const bytes = new Uint8Array([80, 75, 3, 4]);
  let uploaded;
  const client = new DocumentCommands(async (url, options) => {
    if (options?.method === 'POST') {
      assert.match(url, /\/studio\?revision=3&settings_version=2$/);
      assert.equal(options.headers['Content-Type'], undefined);
      bytes.fill(0);
      uploaded = new Uint8Array(await options.body.get('file').arrayBuffer());
      return { ...receipt, kind: 'studio' };
    }
    return url.includes('/jobs/') ? { ...job, kind: 'studio' } : fresh;
  }, immediate);
  await client.execute(base, { kind: 'studio', bytes });
  assert.deepEqual(uploaded, new Uint8Array([80, 75, 3, 4]));
});

test('validation must complete on the same revision with a matching validation stamp', async () => {
  let validated = false;
  const client = new DocumentCommands(async (url, options) => {
    if (options?.method === 'POST') return { ...receipt, kind: 'validate' };
    if (url.includes('/jobs/')) return { ...job, kind: 'validate', result: { revision: 3 } };
    return { ...fresh, revision: 3, validated: validated ? 3 : null };
  }, immediate);
  await assert.rejects(client.execute(base, { kind: 'validate' }), { code: 'revision_conflict' });
  validated = true;
  assert.equal((await client.execute(base, { kind: 'validate' })).validated, 3);
});
