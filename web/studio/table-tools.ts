import type { CommandRegistry } from '@/command/registry';
import type { CommandServices, EditorContext } from '@/command/types';
import { parentTableCell, type TablePosition } from './nested-table.ts';

type Input = NonNullable<ReturnType<CommandServices['getInputHandler']>>;
type TableEngine = CommandServices['wasm'] & {
  createNestedTable(position: TablePosition, rows: number, cols: number): TablePosition;
  editNestedTable(position: TablePosition, operation: string): TablePosition;
};
const preference = 'jakal-table-size';
function validSize(rows: number, cols: number) {
  return [rows, cols].every(n => Number.isInteger(n) && n > 0 && n <= 100) && rows * cols <= 2500;
}

export function showTableDialog(nested: boolean, apply: (rows: number, cols: number) => void, focus: () => void) {
  let size = { rows: 2, cols: 2 };
  try {
    const previous = JSON.parse(localStorage.getItem(preference) || 'null');
    if (previous && validSize(previous.rows, previous.cols)) size = previous;
  } catch { /* Preferences are optional. */ }
  const dialog = document.createElement('dialog');
  dialog.className = 'jakal-table-dialog';
  dialog.setAttribute('aria-labelledby', 'jakal-table-title');
  dialog.innerHTML = `<form>
    <h2 id="jakal-table-title"></h2><p class="jakal-table-hint"></p>
    <div class="jakal-table-presets" role="group" aria-label="자주 쓰는 표 크기">
      <button type="button" data-rows="1" data-cols="1">1 × 1</button><button type="button" data-rows="2" data-cols="2">2 × 2</button>
      <button type="button" data-rows="3" data-cols="3">3 × 3</button><button type="button" data-rows="5" data-cols="2">5 × 2</button>
    </div>
    <div class="jakal-table-size"><label>줄 (행)<input name="rows" type="number" min="1" max="100" step="1" required></label>
    <span aria-hidden="true">×</span><label>칸 (열)<input name="cols" type="number" min="1" max="100" step="1" required></label></div>
    <p class="jakal-table-summary" role="status" aria-live="polite"></p><p class="jakal-table-error" role="alert" hidden></p>
    <footer><button type="button" data-cancel>취소</button><button type="submit">표 만들기</button></footer>
  </form>`;
  dialog.querySelector('h2')!.textContent = nested ? '셀 안에 표 만들기' : '표 만들기';
  dialog.querySelector('.jakal-table-hint')!.textContent = nested
    ? '기존 내용을 유지하고, 셀 너비에 맞춰 새 표를 넣습니다.' : '줄과 칸을 정하세요. Enter로 만들고 Esc로 닫을 수 있습니다.';
  const rows = dialog.querySelector<HTMLInputElement>('[name="rows"]')!;
  const cols = dialog.querySelector<HTMLInputElement>('[name="cols"]')!;
  rows.value = String(size.rows); cols.value = String(size.cols);
  const summary = dialog.querySelector<HTMLElement>('.jakal-table-summary')!;
  const submit = dialog.querySelector<HTMLButtonElement>('[type="submit"]')!;
  const update = () => {
    const valid = validSize(rows.valueAsNumber, cols.valueAsNumber);
    summary.textContent = valid ? `${rows.value}줄 × ${cols.value}칸 · ${rows.valueAsNumber * cols.valueAsNumber}셀` : '1~100줄·칸, 총 2,500셀 이내로 입력하세요.';
    submit.disabled = !valid;
  };
  const close = () => { dialog.close(); dialog.remove(); focus(); };
  dialog.addEventListener('cancel', event => { event.preventDefault(); close(); });
  dialog.querySelector('[data-cancel]')!.addEventListener('click', close);
  dialog.addEventListener('click', event => { if (event.target === dialog) close(); });
  dialog.addEventListener('input', update);
  for (const button of dialog.querySelectorAll<HTMLButtonElement>('[data-rows]')) {
    button.addEventListener('click', () => { rows.value = button.dataset.rows!; cols.value = button.dataset.cols!; update(); rows.focus(); rows.select(); });
  }
  dialog.querySelector('form')!.addEventListener('submit', event => {
    event.preventDefault();
    if (!validSize(rows.valueAsNumber, cols.valueAsNumber)) return;
    try {
      apply(rows.valueAsNumber, cols.valueAsNumber);
      try { localStorage.setItem(preference, JSON.stringify({ rows: rows.valueAsNumber, cols: cols.valueAsNumber })); } catch { /* optional preference */ }
      close();
    } catch (error) {
      const message = dialog.querySelector<HTMLElement>('.jakal-table-error')!;
      message.hidden = false; message.textContent = error instanceof Error ? error.message : String(error);
    }
  });
  update(); document.body.appendChild(dialog); dialog.showModal(); rows.focus(); rows.select();
  return dialog;
}

export function installTableTools(registry: CommandRegistry, getInput: () => Input | null) {
  const original = registry.get('table:create');
  if (!original) throw new Error('표 만들기 명령을 찾을 수 없습니다.');
  const canCreate = (context: EditorContext) => {
    const input = getInput(), cursor = (input as any)?.cursor;
    return !!input && !!cursor && context.hasDocument && context.isEditable && !context.isFormMode
      && !context.hasSelection && !context.inField && !context.inCellSelectionMode
      && !context.inTableObjectSelection && !context.inPictureObjectSelection
      && !cursor.isInHeaderFooter() && !cursor.isInFootnote() && !cursor.isInTextBox();
  };
  const execute = (services: CommandServices) => {
    const input = services.getInputHandler();
    if (!input) return;
    const position = structuredClone(input.getCursorPosition());
    const generation = services.wasm.documentGeneration;
    showTableDialog(position.parentParaIndex !== undefined, (rows, cols) => {
      if (generation !== services.wasm.documentGeneration) throw new Error('문서가 바뀌었습니다. 창을 닫고 입력 위치를 다시 선택하세요.');
      input.executeOperation({ kind: 'snapshot', operationType: 'createTable', operation: engine => {
        if (position.parentParaIndex !== undefined) return (engine as TableEngine).createNestedTable(position, rows, cols);
        const result = engine.createTableEx({ sectionIdx: position.sectionIndex, paraIdx: position.paragraphIndex,
          charOffset: position.charOffset, rowCount: rows, colCount: cols });
        if (!result.ok) throw new Error('표를 만들지 못했습니다. 입력 위치를 다시 선택하세요.');
        return { sectionIndex: position.sectionIndex, paragraphIndex: 0, charOffset: 0,
          parentParaIndex: result.paraIdx, controlIndex: result.controlIdx, cellIndex: 0, cellParaIndex: 0,
          cellPath: [{ controlIndex: result.controlIdx, cellIndex: 0, cellParaIndex: 0 }] };
      } });
    }, () => input.focus());
  };
  registry.register({ ...original, canExecute: canCreate, execute });
  registry.register({ id: 'table:create-nested', label: '셀 안에 표 만들기', opensDialog: true,
    canExecute: context => context.inTable && canCreate(context), execute });
  registry.register({ id: 'table:parent-cell', label: '상위 셀로 이동', canExecute: context =>
    context.hasDocument && context.inTable && (getInput()?.getCursorPosition().cellPath?.length ?? 0) > 1,
    execute: services => {
      const input = services.getInputHandler();
      if (input) { input.moveCursorTo(parentTableCell(input.getCursorPosition())); input.focus(); }
    } });
  for (const [id, operation] of Object.entries({
    'table:insert-row-above': 'insertRowAbove', 'table:insert-row-below': 'insertRowBelow',
    'table:insert-col-left': 'insertColLeft', 'table:insert-col-right': 'insertColRight',
    'table:delete-row': 'deleteRow', 'table:delete-col': 'deleteCol',
  })) {
    const command = registry.get(id);
    if (!command) continue;
    registry.register({ ...command, execute(services, params) {
      const input = services.getInputHandler(), position = input?.getCursorPosition();
      if (!input || !position || (position.cellPath?.length ?? 0) < 2) return command.execute(services, params);
      input.executeOperation({ kind: 'snapshot', operationType: operation,
        operation: engine => (engine as TableEngine).editNestedTable(position, operation) });
      input.focus();
    } });
  }
}
