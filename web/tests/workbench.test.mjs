import assert from 'node:assert/strict';
import test from 'node:test';
import { visibleCommands, blockText, blockKinds } from '../src/workbench.ts';

const commands = [
  { id: 'insert:equation', label: '수식', enabled: true },
  { id: 'table:merge', label: '셀 병합', enabled: false },
  { id: 'format:bold', label: '굵게', shortcutLabel: 'Ctrl+B', enabled: true },
  { id: 'file:new', label: '새 문서', enabled: true },
  { id: 'jakal:insert-block', label: '블록', enabled: true },
  { id: 'edit:document-history', label: '내부 기록', enabled: true },
];
test('tool search respects scope, context and multiple keywords', () => {
  assert.deepEqual(visibleCommands(commands, '표 병합', '').map(x => x.id), ['table:merge']);
  assert.equal(visibleCommands(commands, 'ctrl+b', '')[0].id, 'format:bold');
  assert.equal(visibleCommands(commands, '수식', 'table').length, 0);
  assert.deepEqual(visibleCommands(commands, '', '').map(x => x.id), ['insert:equation', 'format:bold', 'table:merge']);
});
test('blocks preserve typed math, Korean and markup as plain text', () => {
  const text = 'ㄱ. x^2 < 3\r\nㄴ. <script>not executable</script> 😀';
  assert.equal(blockText('examples', text, false), '\n' + text.replace('\r\n', '\n'));
  assert.equal(blockText('choices', '① A\n② B\n', true), '\n① A    ② B\n');
  assert.equal(blockText('answer', '답안\n\n\n', false), '\n답안\n\n\n');
});
test('block size and control character gates reject malformed payloads', () => {
  for (const text of ['', '   ', 'x'.repeat(12001), '\u0000x', 'x\n'.repeat(121)]) assert.throws(() => blockText('question', text, false));
  assert.throws(() => blockText('unknown', 'text', false));
  assert.equal(blockText('question', 'x'.repeat(12000), false).length, 12001);
  assert.equal(new Set(blockKinds.map(x => x.id)).size, blockKinds.length);
});
