export type CellPathEntry = { controlIndex: number; cellIndex: number; cellParaIndex: number };
export type TablePosition = {
  sectionIndex: number; paragraphIndex: number; charOffset: number;
  parentParaIndex?: number; controlIndex?: number; cellIndex?: number; cellParaIndex?: number;
  cellPath?: CellPathEntry[];
};
export type TableOptions = { colWidths?: number[]; rowHeights?: number[]; treatAsChar?: boolean; widthPercent?: number };
export type CellRange = { startRow: number; startCol: number; endRow: number; endCol: number };
export type StructureOptions = { rows?: number; cols?: number; range?: CellRange };
type NativeDocument = {
  getParagraphCount(section: number): number;
  getParagraphLength(section: number, paragraph: number): number;
  splitParagraph(section: number, paragraph: number, offset: number): string;
  deleteParagraph(section: number, paragraph: number): string;
  hasInternalClipboard(): boolean;
  clipboardHasControl(): boolean;
  getClipboardText(): string;
  clearClipboard(): void;
  pasteInternal(section: number, paragraph: number, offset: number): string;
  copySelection(section: number, first: number, start: number, last: number, end: number): string;
  copyControl(section: number, paragraph: number, path: string, control: number): string;
  createTableEx(options: string): string;
  setTableProperties(section: number, paragraph: number, control: number, props: string): string;
  getCellPropertiesByPath(section: number, paragraph: number, path: string, cell: number): string;
  splitParagraphInCellByPath(section: number, paragraph: number, path: string, offset: number): string;
  pasteInternalInCellByPath(section: number, paragraph: number, path: string, offset: number): string;
  pasteControl(section: number, paragraph: number, offset: number): string;
  getCellInfoByPath(section: number, paragraph: number, path: string): string;
  getCellParagraphLengthByPath(section: number, paragraph: number, path: string): number;
  resizeTableCellsByPath(section: number, paragraph: number, path: string, changes: string): string;
  getCursorModel(): string;
  getControls(): string;
  getCharIndexAtStreamPos(list: number, paragraph: number, position: number): string;
  deleteCellTableControlByPath(section: number, paragraph: number, path: string, control: number): string;
  insertTableRow(section: number, paragraph: number, control: number, row: number, below: boolean): string;
  insertTableColumn(section: number, paragraph: number, control: number, col: number, right: boolean): string;
  deleteTableRow(section: number, paragraph: number, control: number, row: number): string;
  deleteTableColumn(section: number, paragraph: number, control: number, col: number): string;
  splitTableCellInto(section: number, paragraph: number, control: number, row: number, col: number, rows: number, cols: number, equal: boolean, mergeFirst: boolean): string;
  splitTableCellsInRange(section: number, paragraph: number, control: number, startRow: number, startCol: number, endRow: number, endCol: number, rows: number, cols: number, equal: boolean): string;
  mergeTableCells(section: number, paragraph: number, control: number, startRow: number, startCol: number, endRow: number, endCol: number): string;
  beginBatch(): string;
  endBatch(): string;
};

function checked(raw: string): Record<string, any> {
  const value = JSON.parse(raw);
  if (!value.ok) throw new Error('표를 삽입하지 못했습니다. 문서의 입력 위치를 확인해 주세요.');
  return value;
}

export function tableCellPath(position: TablePosition): CellPathEntry[] {
  if (position.parentParaIndex === undefined) throw new Error('표를 넣을 셀 안을 먼저 클릭하세요.');
  const path = position.cellPath?.length ? position.cellPath : [{
    controlIndex: position.controlIndex!, cellIndex: position.cellIndex!, cellParaIndex: position.cellParaIndex!,
  }];
  if (path.some(entry => Object.values(entry).some(value => !Number.isSafeInteger(value) || value < 0))) {
    throw new Error('표를 넣을 셀 안을 다시 클릭하세요.');
  }
  return path.map(entry => ({ ...entry }));
}

function withScratch<T>(doc: NativeDocument, section: number,
  operation: (append: () => number, copy: (paragraph: number, path: string, control: number) => void) => T): T {
  const originalCount = doc.getParagraphCount(section);
  const hadClipboard = doc.hasInternalClipboard();
  const copiedObject = hadClipboard && doc.clipboardHasControl() && !!doc.getClipboardText();
  let backup: { first: number; last: number; end: number } | undefined;
  let clipboardChanged = false;
  const append = () => {
    const last = doc.getParagraphCount(section) - 1;
    return checked(doc.splitParagraph(section, last, doc.getParagraphLength(section, last))).paraIdx as number;
  };
  doc.beginBatch();
  try {
    if (hadClipboard) {
      const first = append();
      const pasted = checked(doc.pasteInternal(section, first, 0));
      backup = { first, last: pasted.paraIdx, end: doc.getParagraphLength(section, pasted.paraIdx) };
    }
    return operation(append, (paragraph, path, control) => {
      clipboardChanged = true;
      checked(doc.copyControl(section, paragraph, path, control));
    });
  } finally {
    try {
      if (clipboardChanged) {
        if (backup) {
          if (copiedObject && backup.first === backup.last && backup.end === 0) checked(doc.copyControl(section, backup.first, '', 0));
          else checked(doc.copySelection(section, backup.first, 0, backup.last, backup.end));
        } else doc.clearClipboard();
      }
      while (doc.getParagraphCount(section) > originalCount) checked(doc.deleteParagraph(section, doc.getParagraphCount(section) - 1));
    } finally { doc.endBatch(); }
  }
}

/** Runs inside SnapshotCommand. Native controls preserve editable nested tables;
 * the engine's HTML-in-cell importer deliberately flattens table markup.
 * Scratch paragraphs stay at the section end, and the internal clipboard is
 * saved/restored there so creating a table never replaces the user's copy buffer.
 */
export function insertNestedTable(
  doc: NativeDocument, position: TablePosition, rows: number, cols: number, options: TableOptions = {},
): TablePosition {
  if (![rows, cols].every(value => Number.isSafeInteger(value) && value >= 1 && value <= 100) || rows * cols > 2500) {
    throw new Error('표는 1~100줄·칸, 총 2,500셀 이내로 만들어 주세요.');
  }
  const path = tableCellPath(position);
  const section = position.sectionIndex, parent = position.parentParaIndex!;
  const props = JSON.parse(doc.getCellPropertiesByPath(section, parent, JSON.stringify(path), path.at(-1)!.cellIndex));
  if (props.cellProtect) throw new Error('보호된 셀에는 표를 넣을 수 없습니다.');
  const available = Math.floor(props.width - props.paddingLeft - props.paddingRight);
  if (!Number.isFinite(available) || available < cols) throw new Error('셀의 너비가 부족합니다. 바깥 표의 너비를 넓혀 주세요.');
  const requested = options.colWidths ?? Array(cols).fill(1);
  if (requested.length !== cols || requested.some(width => !Number.isFinite(width) || width <= 0)) {
    throw new Error('칸 너비를 확인해 주세요.');
  }
  const total = requested.reduce((sum, width) => sum + width, 0);
  const percent = options.widthPercent ?? 75;
  if (!Number.isFinite(percent) || percent < 10 || percent > 100) throw new Error('표 너비를 10~100%로 입력하세요.');
  const targetWidth = Math.max(cols, Math.floor(available * percent / 100));
  const width = options.colWidths ? Math.min(total, targetWidth) : targetWidth;
  const colWidths = requested.map(value => Math.max(1, Math.floor(value / total * (width - cols)) + 1));
  if (options.rowHeights && (options.rowHeights.length !== rows || options.rowHeights.some(h => !Number.isSafeInteger(h) || h < 1 || h > 1_000_000))) {
    throw new Error('줄 높이를 확인해 주세요.');
  }
  const cursor = withScratch(doc, section, (appendScratch, copy) => {
    const scratch = appendScratch();
    const table = checked(doc.createTableEx(JSON.stringify({
      sectionIdx: section, paraIdx: scratch, charOffset: 0, rowCount: rows, colCount: cols,
      treatAsChar: true, colWidths, ...(options.rowHeights ? { rowHeights: options.rowHeights } : {}),
    })));
    checked(doc.setTableProperties(section, table.paraIdx, table.controlIdx,
      JSON.stringify({ outerLeft: 141, outerRight: 141, outerTop: 141, outerBottom: 141 })));
    copy(table.paraIdx, '', table.controlIdx);
    // Give the new table its own paragraph, preserving text and controls on
    // either side and leaving a paragraph for exiting the table afterwards.
    const split = checked(doc.splitParagraphInCellByPath(section, parent, JSON.stringify(path), position.charOffset));
    path[path.length - 1].cellParaIndex = split.cellParaIndex;
    checked(doc.splitParagraphInCellByPath(section, parent, JSON.stringify(path), 0));
    checked(doc.pasteInternalInCellByPath(section, parent, JSON.stringify(path), 0));
    return {
      sectionIndex: section, paragraphIndex: 0, charOffset: 0, parentParaIndex: parent,
      controlIndex: path[0].controlIndex, cellIndex: path[0].cellIndex, cellParaIndex: path[0].cellParaIndex,
      cellPath: [...path, { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }],
    };
  });
  refreshTableLayout(doc, position);
  return cursor;
}

function refreshTableLayout(doc: NativeDocument, position: TablePosition) {
  // Clipboard edits refresh cell paragraphs but leave ancestor row measurements
  // cached. A zero-size resize recomputes them without changing stored sizes.
  const outer = tableCellPath(position).slice(0, 1);
  checked(doc.resizeTableCellsByPath(position.sectionIndex, position.parentParaIndex!, JSON.stringify(outer),
    JSON.stringify([{ cellIdx: outer[0].cellIndex, heightDelta: 0 }])));
}

export function parentTableCell(position: TablePosition): TablePosition {
  const path = tableCellPath(position);
  if (path.length < 2) throw new Error('현재 표는 바깥쪽 표입니다.');
  path.pop();
  return { sectionIndex: position.sectionIndex, paragraphIndex: path.at(-1)!.cellParaIndex, charOffset: 0,
    parentParaIndex: position.parentParaIndex, controlIndex: path[0].controlIndex,
    cellIndex: path[0].cellIndex, cellParaIndex: path[0].cellParaIndex, cellPath: path };
}

/** Flat row/column APIs target the outer table; resolve the innermost cell list. */
export function editNestedTable(doc: NativeDocument, position: TablePosition, operation: string, options: StructureOptions = {}): TablePosition {
  if (!['insertRowAbove', 'insertRowBelow', 'insertColLeft', 'insertColRight', 'deleteRow', 'deleteCol', 'splitCells', 'mergeCells'].includes(operation)) {
    throw new Error('지원하지 않는 표 편집입니다.');
  }
  const path = tableCellPath(position), model = JSON.parse(doc.getCursorModel());
  if (path.length < 2) throw new Error('중첩 표의 셀을 선택하세요.');
  const info = JSON.parse(doc.getCellInfoByPath(position.sectionIndex, position.parentParaIndex!, JSON.stringify(path)));
  const parentPath = path.slice(0, -1), target = path.at(-1)!;
  let list = 0, paragraph = position.parentParaIndex!;
  for (const entry of parentPath) {
    const cell = model.lists.find((value: any) => value.isCell && value.hostListId === list
      && value.sectionIndex === position.sectionIndex && value.hostPara === paragraph
      && value.controlIndex === entry.controlIndex && value.cellIndex === entry.cellIndex);
    if (!cell) throw new Error('편집할 표 셀을 찾지 못했습니다. 셀을 다시 선택하세요.');
    list = cell.listId; paragraph = entry.cellParaIndex;
  }
  const anchor = JSON.parse(doc.getControls()).find((control: any) => control.list === list
    && control.para === paragraph && control.controlIndex === target.controlIndex && control.ctrlId === 'tbl');
  // Some engine versions omit nested controls from later sections in
  // getControls(). Our inserted tables occupy an otherwise empty paragraph.
  if (!anchor && doc.getCellParagraphLengthByPath(position.sectionIndex, position.parentParaIndex!, JSON.stringify(parentPath)) !== 0) {
    throw new Error('이 표의 입력 위치를 확인하지 못했습니다. 표를 별도 문단에 놓고 다시 시도하세요.');
  }
  const offset = anchor ? JSON.parse(doc.getCharIndexAtStreamPos(list, paragraph, anchor.pos)).charIndex : 0;
  const cursor = withScratch(doc, position.sectionIndex, (append, copy) => {
    copy(position.parentParaIndex!, JSON.stringify(parentPath), target.controlIndex);
    const scratch = append(), table = checked(doc.pasteControl(position.sectionIndex, scratch, 0));
    const args = [position.sectionIndex, table.paraIdx, table.controlIdx] as const;
    if (operation === 'insertRowAbove' || operation === 'insertRowBelow') checked(doc.insertTableRow(...args, info.row, operation === 'insertRowBelow'));
    if (operation === 'insertColLeft' || operation === 'insertColRight') checked(doc.insertTableColumn(...args, info.col, operation === 'insertColRight'));
    if (operation === 'deleteRow') checked(doc.deleteTableRow(...args, info.row));
    if (operation === 'deleteCol') checked(doc.deleteTableColumn(...args, info.col));
    if (operation === 'splitCells') {
      if (![options.rows, options.cols].every(value => Number.isSafeInteger(value) && value! >= 1 && value! <= 100)
        || options.rows! * options.cols! > 2500) throw new Error('나눌 줄과 칸 수를 확인하세요.');
      if (options.range) {
        const r = options.range;
        checked(doc.splitTableCellsInRange(...args, r.startRow, r.startCol, r.endRow, r.endCol, options.rows!, options.cols!, false));
      } else checked(doc.splitTableCellInto(...args, info.row, info.col, options.rows!, options.cols!, false, false));
    }
    if (operation === 'mergeCells') {
      const r = options.range;
      if (!r) throw new Error('합칠 칸을 먼저 선택하세요.');
      checked(doc.mergeTableCells(...args, r.startRow, r.startCol, r.endRow, r.endCol));
    }
    copy(table.paraIdx, '', table.controlIdx);
    checked(doc.deleteCellTableControlByPath(position.sectionIndex, position.parentParaIndex!, JSON.stringify(parentPath), target.controlIndex));
    checked(doc.pasteInternalInCellByPath(position.sectionIndex, position.parentParaIndex!, JSON.stringify(parentPath), offset));
    // Native paste appends the replacement control after any siblings.
    const newControl = JSON.parse(doc.getCursorModel()).lists.filter((cell: any) => cell.isCell && cell.hostListId === list
      && cell.sectionIndex === position.sectionIndex && cell.hostPara === paragraph)
      .reduce((max: number, cell: any) => Math.max(max, cell.controlIndex), 0);
    target.controlIndex = newControl; target.cellIndex = 0; target.cellParaIndex = 0;
    return { sectionIndex: position.sectionIndex, paragraphIndex: 0, charOffset: 0,
      parentParaIndex: position.parentParaIndex, controlIndex: path[0].controlIndex,
      cellIndex: path[0].cellIndex, cellParaIndex: path[0].cellParaIndex, cellPath: path };
  });
  refreshTableLayout(doc, position);
  return cursor;
}
