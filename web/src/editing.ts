import type { Document, Operation, Settings } from './api';

export type TextTarget = { id: string; label: string; original: string; text: string; editable: boolean; operation: Operation };
export type Proposal = { id: string; label: string; before: string; after: string; reason: string; operation: Operation; blocked?: string };

// API snapshots are immutable in React. Scope this index to the sections array,
// so settings/status updates reuse it and new document content always replaces it.
type IndexedTarget = Pick<TextTarget, 'id' | 'original' | 'editable' | 'operation'> & { order: number };
const targetIndexes = new WeakMap<Document['sections'], ReadonlyMap<string, IndexedTarget>>();
export function textTargetIndex(doc: Document) {
  let index = targetIndexes.get(doc.sections);
  if (!index) {
    const built = new Map<string, IndexedTarget>();
    for (const section of doc.sections) {
      for (const p of section.paragraphs) built.set(p.id, { id: p.id, original: p.text, editable: p.editable, order: built.size, operation: { type: 'paragraph', section: section.index, index: p.index, text: p.text } });
      for (const table of section.tables) for (const cell of table.cells) {
        const id = `${table.id}-${cell.row}-${cell.column}`;
        built.set(id, { id, original: cell.text, editable: cell.editable, order: built.size, operation: { type: 'cell', section: section.index, table: table.index, row: cell.row, column: cell.column, text: cell.text } });
      }
    }
    index = built;
    targetIndexes.set(doc.sections, index);
  }
  return index;
}

export function textTargets(doc: Document, drafts: Record<string, Operation>): TextTarget[] {
  return doc.sections.flatMap(s => [
    ...s.paragraphs.map(p => ({ id: p.id, label: `섹션 ${s.index + 1} · 문단 ${p.index + 1}`, original: p.text, text: drafts[p.id]?.text ?? p.text, editable: p.editable, operation: { type: 'paragraph' as const, section: s.index, index: p.index, text: p.text } })),
    ...s.tables.flatMap(t => t.cells.map(c => {
      const id = `${t.id}-${c.row}-${c.column}`;
      return { id, label: `섹션 ${s.index + 1} · 표 ${t.index + 1} · ${c.row + 1}행 ${c.column + 1}열`, original: c.text, text: drafts[id]?.text ?? c.text, editable: c.editable, operation: { type: 'cell' as const, section: s.index, table: t.index, row: c.row, column: c.column, text: c.text } };
    })),
  ]);
}

export function protectionError(before: string, after: string, settings: Settings): string | undefined {
  const term = settings.protected_terms.find(t => before.split(t).length !== after.split(t).length);
  if (term) return `보호 문구 ‘${term}’ 변경`;
  if (after.length > 100000) return '요소당 100,000자 초과';
}

export function replaceProposals(items: TextTarget[], find: string, replacement: string, caseSensitive: boolean, settings: Settings): Proposal[] {
  if (!find) return [];
  // Literal text only: user input never becomes an executable regular expression.
  const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(escaped, caseSensitive ? 'gu' : 'giu');
  return items.filter(item => item.editable).flatMap(item => {
    let length = item.text.length;
    let matches = 0;
    for (const match of item.text.matchAll(regex)) {
      length += replacement.length - match[0].length;
      matches++;
    }
    if (!matches) return [];
    if (length > 100000) return [{ id: item.id, label: item.label, before: item.text, after: '(결과가 100,000자를 초과하여 미리보기를 생략했습니다.)', reason: '찾기·바꾸기', operation: item.operation, blocked: '요소당 100,000자 초과' }];
    const after = item.text.replace(regex, () => replacement);
    return after === item.text ? [] : [{ id: item.id, label: item.label, before: item.text, after, reason: '찾기·바꾸기', operation: { ...item.operation, text: after }, blocked: protectionError(item.original, after, settings) }];
  });
}
