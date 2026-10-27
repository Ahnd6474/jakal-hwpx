import type { Section, Table } from './api';

type Snapshot = { sections: Section[] };
type Block = { key: string; text: string; index: number };
export type RevisionChange = { id: string; label: string; before: string; after: string; beforeCount: number; afterCount: number; clipped: boolean };
const MAX_ITEMS = 500;
const MAX_TEXT = 16_000;

// Patience anchors align insertions without treating every following paragraph
// as a replacement. Ambiguous repeated passages remain a single changed group.
function ranges(a: Block[], b: Block[], emit: (a: Block[], b: Block[]) => void) {
  function visit(al: number, ar: number, bl: number, br: number, depth: number) {
    while (al < ar && bl < br && a[al].key === b[bl].key) { al++; bl++; }
    while (al < ar && bl < br && a[ar - 1].key === b[br - 1].key) { ar--; br--; }
    if (al === ar && bl === br) return;
    if (al === ar || bl === br || depth >= 32) { emit(a.slice(al, ar), b.slice(bl, br)); return; }
    const unique = (items: Block[], left: number, right: number) => {
      const map = new Map<string, number>();
      for (let i = left; i < right; i++) map.set(items[i].key, map.has(items[i].key) ? -1 : i);
      return map;
    };
    const old = unique(a, al, ar), fresh = unique(b, bl, br);
    const candidates: [number, number][] = [];
    for (let i = al; i < ar; i++) {
      const j = fresh.get(a[i].key);
      if (old.get(a[i].key) === i && j !== undefined && j >= 0) candidates.push([i, j]);
    }
    const tails: number[] = [], previous: number[] = [];
    for (let i = 0; i < candidates.length; i++) {
      let low = 0, high = tails.length;
      while (low < high) { const mid = (low + high) >> 1; if (candidates[tails[mid]][1] < candidates[i][1]) low = mid + 1; else high = mid; }
      previous[i] = low ? tails[low - 1] : -1; tails[low] = i;
    }
    if (!tails.length) { emit(a.slice(al, ar), b.slice(bl, br)); return; }
    const anchors: [number, number][] = [];
    for (let i = tails[tails.length - 1]; i >= 0; i = previous[i]) anchors.push(candidates[i]);
    anchors.reverse();
    for (const [i, j] of anchors) { visit(al, i, bl, j, depth + 1); al = i + 1; bl = j + 1; }
    visit(al, ar, bl, br, depth + 1);
  }
  visit(0, a.length, 0, b.length, 0);
}

function tableBlock(table: Table): Block {
  const cells = [...table.cells].sort((a, b) => a.row - b.row || a.column - b.column);
  const geometry = [table.rows, table.columns, cells.map(c => [c.row, c.column, c.row_span, c.col_span])];
  const text = `${table.rows}행 × ${table.columns}열\n` + cells.map(c => `${c.row + 1}행 ${c.column + 1}열${c.row_span > 1 || c.col_span > 1 ? ` (${c.row_span}행 × ${c.col_span}열 병합)` : ''}: ${c.text}`).join('\n');
  return { key: JSON.stringify(geometry) + '\n' + text, text, index: table.index };
}

export function compareRevisions(before: Snapshot, after: Snapshot) {
  const items: RevisionChange[] = [];
  let total = 0, paragraphs = 0, tables = 0;
  const oldSections = new Map(before.sections.map(s => [s.index, s]));
  const newSections = new Map(after.sections.map(s => [s.index, s]));
  const sections = [...new Set([...oldSections.keys(), ...newSections.keys()])].sort((a, b) => a - b);
  for (const section of sections) {
    const old = oldSections.get(section), fresh = newSections.get(section);
    for (const kind of ['paragraph', 'table'] as const) {
      const blocks = (s: Section | undefined) => kind === 'table' ? s?.tables.map(tableBlock) || [] : s?.paragraphs.map(p => ({ key: p.text, text: p.text, index: p.index })) || [];
      ranges(blocks(old), blocks(fresh), (a, b) => {
        total++;
        if (kind === 'paragraph') paragraphs++; else tables++;
        if (items.length >= MAX_ITEMS) return;
        const oldText = a.map(item => item.text).join('\n'), newText = b.map(item => item.text).join('\n');
        const location = (values: Block[]) => values.length === 1 ? `${values[0].index + 1}` : `${values[0].index + 1}–${values.at(-1)!.index + 1}`;
        const source = a.length ? a : b;
        items.push({ id: `${section}-${kind}-${total}`, label: `섹션 ${section + 1} · ${kind === 'paragraph' ? '문단' : '표'} ${location(source)}`, before: oldText.slice(0, MAX_TEXT), after: newText.slice(0, MAX_TEXT), beforeCount: a.length, afterCount: b.length, clipped: oldText.length > MAX_TEXT || newText.length > MAX_TEXT });
      });
    }
  }
  return { items, total, paragraphs, tables, omitted: total - items.length };
}
