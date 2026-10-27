import type { CommandRegistry } from '@/command/registry';
import type { CommandServices, EditorContext } from '@/command/types';
import { chartTargetFromSelection, matchChartRef } from '@/core/chart-data-target';

type Input = NonNullable<ReturnType<CommandServices['getInputHandler']>>;
type Kind = 'table' | 'image' | 'chart' | 'object';
type Selection = { kind: Kind; key: string; multiple?: boolean; type?: string; nested?: boolean };
type Action = [id: string, label: string];
type Group = { label: string; actions: Action[] };
const titles: Record<Kind, string> = { table: '표 레이아웃', image: '그림 레이아웃', chart: '차트 레이아웃', object: '개체 레이아웃' };

export function resolveLayoutSelection(context: EditorContext, input: Input | null, wasm: CommandServices['wasm']): Selection | null {
  if (!context.hasDocument || !input) return null;
  // A picture inside a table takes precedence over the enclosing cell.
  if (context.inPictureObjectSelection) {
    const ref = input.getSelectedPictureRef();
    if (!ref) return null;
    const multiple = (input.getSelectedPictureRefs?.().length ?? 1) > 1;
    let kind: Kind = ref.type === 'image' ? 'image' : 'object';
    if (multiple) kind = 'object';
    else if (ref.type === 'ole') {
      try {
        const target = chartTargetFromSelection(ref);
        if (target && matchChartRef(wasm.listCharts(), target)) kind = 'chart';
      } catch { /* Unrecognized OLEs keep the generic object tools. */ }
    }
    return { kind, key: JSON.stringify([ref, multiple]), multiple, type: ref.type };
  }
  if (context.inTable || context.inCellSelectionMode || context.inTableObjectSelection) {
    const ref = context.inTableObjectSelection ? input.getSelectedTableRef() : null;
    const pos = input.getCursorPosition();
    return { kind: 'table', nested: !ref && (pos.cellPath?.length ?? 0) > 1, key: JSON.stringify(ref ?? [pos.sectionIndex, pos.parentParaIndex, pos.controlIndex,
      pos.cellPath?.map((entry, index, path) => [entry.controlIndex, index < path.length - 1 ? entry.cellIndex : null, index < path.length - 1 ? entry.cellParaIndex : null])]) };
  }
  return null;
}

function groupsFor(selection: Selection): Group[] {
  if (selection.kind === 'table') return [
    { label: '속성', actions: [['table:cell-props', '표·셀 속성'], ['table:border-each', '테두리·배경']] },
    { label: '칸 편집', actions: [['table:cell-split', '칸 나누기'], ['table:cell-merge', '칸 합치기']] },
    { label: '줄·칸 추가', actions: [['table:insert-row-above', '위에 줄'], ['table:insert-row-below', '아래에 줄'], ['table:insert-col-left', '왼쪽에 칸'], ['table:insert-col-right', '오른쪽에 칸']] },
    { label: '표 안에 표', actions: [['table:create-nested', '표 안에 새 표 넣기'], ['table:parent-cell', '상위 칸으로']] },
  ];
  if (selection.multiple) return [{ label: '개체 묶기', actions: [['insert:group-shapes', '개체 묶기']] }];
  const groups: Group[] = [{ label: '크기·위치·배치', actions: [['insert:picture-props', selection.kind === 'chart' ? '차트 크기·배치' : '크기·배치·여백']] }];
  if (selection.kind === 'chart') groups.push({ label: '차트', actions: [['insert:chart-data-edit', '차트 데이터 편집']] });
  if (selection.type === 'image' || selection.type === 'shape' || selection.type === 'line') {
    groups.push({ label: '회전·대칭', actions: [['insert:rotate-ccw', '왼쪽 90°'], ['insert:rotate-cw', '오른쪽 90°'], ['insert:flip-horz', '좌우 뒤집기'], ['insert:flip-vert', '상하 뒤집기']] });
  }
  if (selection.type === 'shape') groups.push({ label: '순서', actions: [['insert:arrange-front', '맨 앞으로'], ['insert:arrange-back', '맨 뒤로']] });
  if (selection.type === 'equation') groups.push({ label: '수식', actions: [['insert:equation-edit', '수식 편집']] });
  return groups;
}

export function installObjectLayout(registry: CommandRegistry, dispatcher: { dispatch: (id: string) => boolean },
  events: { on: (event: string, listener: () => void) => unknown }, getContext: () => EditorContext,
  getInput: () => Input | null, wasm: CommandServices['wasm']) {
  const header = document.getElementById('studio-header');
  if (!header || header.querySelector('.jakal-object-layout')) return;
  const menu = document.getElementById('menu-bar');
  const region = document.createElement('section'); region.className = 'jakal-object-layout';
  region.setAttribute('aria-label', '개체 레이아웃');
  const tabs = document.createElement('div'); tabs.className = 'jakal-layout-tabs'; tabs.setAttribute('role', 'tablist'); tabs.setAttribute('aria-label', '선택한 개체 도구');
  const tab = document.createElement('button'); tab.type = 'button'; tab.id = 'jakal-layout-tab'; tab.setAttribute('role', 'tab'); tab.setAttribute('aria-controls', 'jakal-layout-panel');
  tab.className = 'jakal-layout-tab menu-command'; tab.dataset.cmd = 'view:object-layout';
  const toolMenu = menu?.querySelector('[data-menu="tool"]');
  if (toolMenu) toolMenu.after(tabs); else (menu ?? header).append(tabs);
  const collapse = document.createElement('button'); collapse.type = 'button'; collapse.className = 'jakal-layout-collapse';
  const panel = document.createElement('div'); panel.id = 'jakal-layout-panel'; panel.className = 'jakal-layout-panel'; panel.setAttribute('role', 'tabpanel'); panel.setAttribute('aria-labelledby', tab.id);
  tabs.append(tab); region.append(collapse, panel);
  // The contextual tools replace the usual toolbar, rather than adding a
  // second tab row below it. MenuBar dispatches this tab and closes its menus.
  if (menu) menu.after(region); else header.append(region);
  let selected: Selection | null = null, expanded = true, manual = false, rendered = '';
  let selectionCache = '', cachedSelection: Selection | null = null;
  const buttons = new Map<string, HTMLButtonElement>();
  function render(selection: Selection | null) {
    panel.replaceChildren(); buttons.clear();
    if (selection) for (const group of groupsFor(selection)) {
      const area = document.createElement('div'); area.className = 'jakal-layout-group'; area.setAttribute('role', 'group'); area.setAttribute('aria-label', group.label);
      const actions = document.createElement('div');
      for (const [id, label] of group.actions) {
        if (!registry.get(id)) continue;
        const button = document.createElement('button'); button.type = 'button'; button.textContent = label; button.dataset.command = id;
        button.addEventListener('click', () => { dispatcher.dispatch(id); update(); });
        actions.append(button); buttons.set(id, button);
      }
      if (!actions.childElementCount) continue;
      const label = document.createElement('small'); label.textContent = group.label; area.append(actions, label); panel.append(area);
    }
    const hint = document.createElement('p'); hint.className = 'jakal-layout-hint';
    hint.textContent = !selection ? '문서에서 표의 칸이나 테두리, 그림, 차트를 클릭하세요. 선택한 개체의 레이아웃 도구가 여기에 나타납니다.'
      : selection.kind === 'table' ? (selection.nested ? '중첩 표: 칸 나누기·합치기와 줄·칸 추가를 사용할 수 있습니다.' : '칸 합치기: 합칠 칸을 먼저 선택하세요.')
      : selection.kind === 'chart' ? '크기·위치·본문 배치는 속성에서, 계열과 값은 데이터 편집에서 변경합니다.'
      : selection.multiple ? '여러 개체가 선택되어 있습니다. 개별 크기와 배치를 바꾸려면 하나만 선택하세요.'
      : '속성에서 크기, 위치, 본문과의 배치 및 여백을 설정하세요.';
    panel.append(hint);
  }
  function update() {
    const context = getContext(), input = getInput();
    // Chart enumeration is needed only when selection or document content changes.
    const cache = JSON.stringify([context.hasDocument, context.inPictureObjectSelection, context.inTable, context.inCellSelectionMode,
      context.inTableObjectSelection, input?.getSelectedPictureRef?.(), input?.getSelectedPictureRefs?.(), input?.getSelectedTableRef?.(), input?.getCursorPosition?.(), wasm.documentGeneration]);
    if (cache !== selectionCache) { cachedSelection = resolveLayoutSelection(context, input, wasm); selectionCache = cache; }
    const next = cachedSelection;
    if (next?.key !== selected?.key || next?.kind !== selected?.kind) {
      expanded = true;
      if (next) manual = false;
    }
    selected = next;
    const signature = JSON.stringify([next?.kind, next?.multiple, next?.type, next?.nested]);
    if (signature !== rendered) { rendered = signature; render(next); }
    tabs.hidden = !context.hasDocument || (!next && !manual);
    region.hidden = tabs.hidden || !expanded;
    header.classList.toggle('jakal-layout-active', !region.hidden);
    panel.hidden = !expanded;
    tab.textContent = next ? titles[next.kind] : '개체 레이아웃';
    tab.setAttribute('aria-selected', String(expanded));
    tab.setAttribute('aria-expanded', String(expanded));
    tab.disabled = !context.hasDocument;
    collapse.textContent = expanded ? '접기' : '펼치기'; collapse.setAttribute('aria-label', `레이아웃 도구 ${expanded ? '접기' : '펼치기'}`);
    for (const [id, button] of buttons) {
      const command = registry.get(id);
      // These upstream dialogs accept only a root table address. Do not apply
      // a nested table's properties to its outer table by accident.
      const rootOnly = !!next?.nested && ['table:cell-props', 'table:border-each'].includes(id);
      button.disabled = rootOnly || !context.isEditable || context.isFormMode || !command || !!command.canExecute && !command.canExecute(context);
      button.title = rootOnly ? '이 속성 창은 바깥 표에서 사용할 수 있습니다.' : '';
    }
  }
  const open = () => { manual = true; expanded = true; update(); region.scrollIntoView?.({ block: 'nearest' }); };
  registry.register({ id: 'view:object-layout', label: '개체 레이아웃 열기', canExecute: context => context.hasDocument, execute: open });
  if (!menu) tab.addEventListener('click', () => dispatcher.dispatch('view:object-layout'));
  collapse.addEventListener('click', () => { expanded = !expanded; update(); });
  for (const element of [region, tabs]) element.addEventListener('mousedown', event => event.preventDefault());
  const restoreStandardToolbar = (event: Event) => {
    if (expanded && (event.target as Element).closest('.menu-title')) { expanded = false; update(); }
  };
  menu?.addEventListener('mousedown', restoreStandardToolbar);
  menu?.addEventListener('click', restoreStandardToolbar);
  const onKeydown = (event: KeyboardEvent) => {
    if (event.key === 'Escape') { event.preventDefault(); expanded = false; update(); getInput()?.focus(); }
  };
  region.addEventListener('keydown', onKeydown); tabs.addEventListener('keydown', onKeydown);
  for (const event of ['cursor-cell-changed', 'cursor-format-changed', 'selection-changed', 'command-state-changed',
    'table-object-selection-changed', 'picture-object-selection-changed', 'document-changed', 'document-mutated', 'document-loaded']) {
    events.on(event, () => { if (event.startsWith('document-')) selectionCache = ''; update(); });
  }
  update();
}
