import assert from 'node:assert/strict';
import test from 'node:test';
import { protectionError, replaceProposals, textTargets } from '../src/editing.ts';

const settings = { instructions: '', protected_terms: [], strict_validation: false };
const item = (text, editable = true) => ({ id: 's0-p0', label: '문단 1', original: text, text, editable, operation: { type: 'paragraph', section: 0, index: 0, text } });

test('literal metacharacters and replacement dollars remain text', () => {
  const changes = replaceProposals([item('가격 [A.*] [A.*]')], '[A.*]', '$&\n새 값', true, settings);
  assert.equal(changes[0].after, '가격 $&\n새 값 $&\n새 값');
});

test('case-sensitive and insensitive choices, Korean and emoji', () => {
  assert.equal(replaceProposals([item('Aa aa 한글 😀')], 'aa', '변경', true, settings)[0].after, 'Aa 변경 한글 😀');
  assert.equal(replaceProposals([item('Aa aa 한글 😀')], 'aa', '변경', false, settings)[0].after, '변경 변경 한글 😀');
});

test('blank replacement deletes matches; blank search and read-only are ignored', () => {
  assert.equal(replaceProposals([item('abc abc')], 'abc', '', true, settings)[0].after, ' ');
  assert.deepEqual(replaceProposals([item('abc')], '', 'x', true, settings), []);
  assert.deepEqual(replaceProposals([item('abc', false)], 'abc', 'x', true, settings), []);
});

test('protected term deletion and insertion are blocked per element', () => {
  const rules = { ...settings, protected_terms: ['자칼'] };
  assert.ok(protectionError('자칼', '다른 이름', rules));
  assert.ok(protectionError('자칼', '자칼 자칼', rules));
  assert.equal(protectionError('자칼 소개', '소개: 자칼', rules), undefined);
  assert.ok(replaceProposals([item('자칼 소개')], '자칼', '다른 이름', true, rules)[0].blocked);
});

test('replacement amplification is rejected before allocating oversized text', () => {
  const proposal = replaceProposals([item('a'.repeat(100000))], 'a', 'b'.repeat(100000), true, settings)[0];
  assert.ok(proposal.blocked);
  assert.ok(proposal.after.length < 100);
});

test('flattened targets use local drafts and retain original text for protection', () => {
  const doc = { sections: [{ index: 0, paragraphs: [{ id: 's0-p0', index: 0, section: 0, text: '원본', editable: true }], tables: [{ id: 's0-t0', index: 0, section: 0, cells: [{ row: 1, column: 2, text: '셀', editable: false }] }] }] };
  const items = textTargets(doc, { 's0-p0': { type: 'paragraph', section: 0, index: 0, text: '초안' } });
  assert.equal(items[0].text, '초안');
  assert.equal(items[0].original, '원본');
  assert.equal(items[1].id, 's0-t0-1-2');
  assert.equal(items[1].editable, false);
});
