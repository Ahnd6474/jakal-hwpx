import assert from 'node:assert/strict';
import test from 'node:test';
import { prepareDownload } from '../src/download-workflow.ts';

const doc = { id: 'document', revision: 2, validated: 2 };
const operations = [{ type: 'paragraph', section: 0, index: 0, text: '수정' }];

test('dirty documents save, validate the new revision, then export that revision', async () => {
  const calls = [];
  const url = await prepareDownload(doc, operations, {
    async save(current, pending) { calls.push(['save', current.revision]); assert.equal(pending, operations); return { ...current, revision: 3, validated: null }; },
    async validate(current) { calls.push(['validate', current.revision]); return { ...current, validated: current.revision }; },
    async export(current) { calls.push(['export', current.revision]); return '/download?revision=3'; },
  });
  assert.deepEqual(calls, [['save', 2], ['validate', 3], ['export', 3]]);
  assert.equal(url, '/download?revision=3');
  assert.equal(doc.revision, 2);
});

test('already validated documents export directly; unvalidated ones only validate and export', async () => {
  const calls = [];
  const steps = {
    async save() { assert.fail('No edits to save'); },
    async validate(current) { calls.push('validate'); return { ...current, validated: current.revision }; },
    async export() { calls.push('export'); return '/download'; },
  };
  await prepareDownload(doc, [], steps);
  assert.deepEqual(calls, ['export']);
  calls.length = 0;
  await prepareDownload({ ...doc, validated: null }, [], steps);
  assert.deepEqual(calls, ['validate', 'export']);
});

test('save and validation failures never request an export', async () => {
  for (const failing of ['save', 'validate']) {
    const calls = [];
    const steps = Object.fromEntries(['save', 'validate', 'export'].map(step => [step, async current => {
      calls.push(step);
      if (step === failing) throw new Error('실패');
      return { ...current, revision: 3, validated: null };
    }]));
    await assert.rejects(prepareDownload(doc, operations, steps), /실패/);
    assert.deepEqual(calls, failing === 'save' ? ['save'] : ['save', 'validate']);
  }
});

test('a validation response for a different revision cannot reach export', async () => {
  await assert.rejects(prepareDownload({ ...doc, validated: null }, [], {
    async save() { assert.fail(); },
    async validate(current) { return { ...current, validated: 1 }; },
    async export() { assert.fail('Unvalidated revision must not be exported'); },
  }), /검증을 완료하지 못했습니다/);
});

test('another tab advancing the document during saving or validation stops the download', async () => {
  for (const changedDuring of ['save', 'validate']) {
    await assert.rejects(prepareDownload(doc, operations, {
      async save(current) { return { ...current, revision: changedDuring === 'save' ? 4 : 3, validated: null }; },
      async validate(current) { return { ...current, revision: 4, validated: 4 }; },
      async export() { assert.fail('A different tab’s revision must not be exported'); },
    }), /문서 버전이 바뀌었습니다/);
  }
});
