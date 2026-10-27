export type EditorCommand = { id: string; label: string; shortcutLabel?: string; enabled: boolean; opensDialog?: boolean };
export const quickCommands = ['edit:undo', 'edit:redo', 'edit:find', 'edit:find-replace', 'format:char-shape', 'format:para-shape', 'table:create', 'page:setup'];
export const commandGroups: Record<string, string> = { edit: '편집', format: '글자·문단', insert: '삽입', table: '표', page: '쪽', view: '보기', field: '입력 필드', hyperlink: '링크', file: '인쇄', tool: '도구' };
const outputCommands = new Set(['file:print', 'file:print-to-pdf']);
export const taskGroups = [
  { id: 'objects', label: '개체 레이아웃', commands: ['view:object-layout', 'table:cell-props', 'insert:picture-props', 'insert:chart-data-edit', 'insert:rotate-cw', 'insert:flip-horz'] },
  { id: 'insert', label: '내용 넣기', commands: ['table:create', 'insert:image', 'insert:equation', 'insert:symbols', 'insert:hyperlink', 'insert:footnote', 'insert:endnote', 'insert:textbox', 'insert:bookmark'] },
  { id: 'table', label: '표 다듬기', commands: ['table:create', 'table:create-nested', 'table:parent-cell', 'table:insert-row-above', 'table:insert-row-below', 'table:insert-col-left', 'table:insert-col-right', 'table:delete-row', 'table:delete-col', 'table:cell-merge', 'table:cell-split', 'table:cell-props', 'table:border-each', 'table:cell-height-equal', 'table:cell-width-equal', 'table:formula'] },
  { id: 'layout', label: '쪽 구성', commands: ['page:setup', 'page:break', 'page:column-break', 'page:col-settings', 'page:header-create', 'page:footer-create', 'page:insert-field-pagenum', 'page:insert-field-totalpage', 'page:page-border', 'page:section-settings'] },
  { id: 'finish', label: '검토·출력', commands: ['edit:find', 'edit:find-replace', 'edit:goto', 'view:para-mark', 'view:ctrl-mark', 'file:print-to-pdf', 'file:print'] },
];
const hidden = new Set(['edit:document-history', 'edit:compare-documents', 'jakal:insert-block']);
const searchAliases: Record<string, string> = {
  'view:object-layout': '표 그림 이미지 차트 도형 레이아웃 탭 개체 속성 배치 크기 여백 layout',
  'insert:picture-props': '그림 이미지 차트 도형 개체 속성 레이아웃 크기 위치 배치 여백',
  'insert:chart-data-edit': '차트 그래프 데이터 계열 값 편집',
  'table:create-nested': '중첩표 중첩 표 안에 표 내부 nested table',
  'table:cell-split': '칸 나누기 셀 분할 줄 열 행 split',
  'table:cell-merge': '칸 합치기 셀 병합 merge',
  'table:parent-cell': '상위 바깥 셀 나가기 중첩표 탈출',
  'table:insert-row-above': '행 추가 위', 'table:insert-row-below': '행 추가 아래',
  'table:insert-col-left': '열 추가 왼쪽', 'table:insert-col-right': '열 추가 오른쪽',
  'table:delete-row': '행 삭제', 'table:delete-col': '열 삭제',
};
export function visibleCommands(commands: EditorCommand[], query: string, group: string) {
  const words = query.normalize('NFKC').toLocaleLowerCase().trim().split(/\s+/).filter(Boolean);
  return commands.filter(item => {
    const category = item.id.split(':')[0];
    return !!commandGroups[category] && (category !== 'file' || outputCommands.has(item.id)) && !hidden.has(item.id) && (!group || category === group) && words.every(word => `${item.label} ${commandGroups[category]} ${item.id} ${item.shortcutLabel || ''} ${searchAliases[item.id] || ''}`.normalize('NFKC').toLocaleLowerCase().includes(word));
  }).sort((a, b) => Number(b.enabled) - Number(a.enabled));
}

export const blockKinds = [
  { id: 'question', name: '문항', description: '번호 · 내용 · 배점', text: '1. 문항 내용을 입력하세요. [3점]\n\n' },
  { id: 'choices', name: '5지 선다', description: '한 줄 또는 한 항목씩', text: '① 선지 1\n② 선지 2\n③ 선지 3\n④ 선지 4\n⑤ 선지 5\n' },
  { id: 'examples', name: '보기', description: 'ㄱ·ㄴ·ㄷ 항목', text: '<보기>\nㄱ. 내용을 입력하세요.\nㄴ. 내용을 입력하세요.\nㄷ. 내용을 입력하세요.\n' },
  { id: 'solution', name: '정답·해설', description: '정답과 풀이 작성', text: '[정답] \n[해설]\n풀이를 입력하세요.\n' },
  { id: 'answer', name: '서술형 답안', description: '답안 작성 영역', text: '[답안]\n\n\n\n' },
  { id: 'source', name: '출처', description: '자료 제목과 출처', text: '[출처] 자료 제목 / 작성자 / 발행일\n' },
] as const;
export function blockText(kind: string, value: string, horizontal: boolean): string {
  if (!blockKinds.some(item => item.id === kind)) throw new Error('지원하지 않는 블록입니다.');
  const text = value.replace(/\r\n?/g, '\n');
  if (!text.trim() || text.length > 12000 || text.split('\n').length > 120 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(text)) throw new Error('블록 내용을 12,000자·120줄 이내로 입력해 주세요.');
  return '\n' + (kind === 'choices' && horizontal ? text.trim().split('\n').filter(Boolean).join('    ') + '\n' : text);
}
