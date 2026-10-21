import { useEffect, useRef, useState } from 'react';
import { BookOpenCheck, ChevronDown, FileText, History, Images, Info, SlidersHorizontal, ShieldCheck, Type, Table2, HelpCircle } from 'lucide-react';

export type WorkspaceView = 'studio' | 'original' | 'paragraphs' | 'tables' | 'figures';
type Props = { view: WorkspaceView; disabled: boolean; dirty: number; studioDirty: boolean; tableCount: number; onView: (view: WorkspaceView) => void; onValidate: () => void; onInfo: () => void; onSettings: () => void; onHelp: () => void; onReview: () => void; onHistory: () => void; onFiles: () => void };
const views = [
  { id: 'studio', name: '문서 편집', icon: FileText, description: '원문 위에서 직접 편집' },
  { id: 'original', name: '원문 보기', icon: BookOpenCheck, description: '저장된 문서 읽기' },
  { id: 'paragraphs', name: '텍스트 편집', icon: Type, description: '본문 이어서 편집' },
  { id: 'tables', name: '표', icon: Table2, description: '표만 모아 보기' },
  { id: 'figures', name: '그래프·그림', icon: Images, description: '그림과 개체 모아 보기' },
] as const;
export function WorkspaceMenu(props: Props) {
  const [open, setOpen] = useState(false);
  const root = useRef<HTMLDivElement>(null);
  const trigger = useRef<HTMLButtonElement>(null);
  const selected = views.find(v => v.id === props.view)!;
  useEffect(() => {
    if (!open) return;
    const items = () => [...root.current!.querySelectorAll<HTMLButtonElement>('[role="menu"] button:not(:disabled)')];
    items()[views.findIndex(v => v.id === props.view)]?.focus();
    function outside(e: PointerEvent) { if (!root.current?.contains(e.target as Node)) setOpen(false); }
    function frameFocus() { setOpen(false); }
    function key(e: KeyboardEvent) {
      if (e.key === 'Escape') { e.preventDefault(); setOpen(false); trigger.current?.focus(); }
      if (e.key === 'Tab') { setOpen(false); return; }
      if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) return;
      e.preventDefault(); const buttons = items(); const at = buttons.indexOf(document.activeElement as HTMLButtonElement);
      const next = e.key === 'Home' ? 0 : e.key === 'End' ? buttons.length - 1 : (at + (e.key === 'ArrowUp' ? -1 : 1) + buttons.length) % buttons.length;
      buttons[next]?.focus();
    }
    document.addEventListener('pointerdown', outside); window.addEventListener('blur', frameFocus); root.current?.addEventListener('keydown', key);
    return () => { document.removeEventListener('pointerdown', outside); window.removeEventListener('blur', frameFocus); root.current?.removeEventListener('keydown', key); };
  }, [open, props.view]);
  function choose(action: () => void) { setOpen(false); action(); }
  return <div className="workspace-menu" ref={root}>
    <button ref={trigger} className="workspace-view-trigger" aria-label="보기 및 도구" aria-haspopup="menu" aria-expanded={open} disabled={props.disabled} onClick={() => setOpen(!open)}><selected.icon size={15} /><span>{selected.name}</span><ChevronDown size={12} /></button>
    {open && <div className="workspace-popup" role="menu" aria-label="보기 및 도구 메뉴"><p>보기 방식</p>{views.map(v => <button key={v.id} role="menuitemradio" aria-checked={props.view === v.id} onClick={() => choose(() => props.onView(v.id))}><v.icon size={16} /><span><b>{v.name}{v.id === 'tables' && ` ${props.tableCount}`}</b><small>{v.description}</small></span>{props.view === v.id && <i />}</button>)}<hr /><button role="menuitem" disabled={!!props.dirty} onClick={() => choose(props.onValidate)}><ShieldCheck size={16} />검증하기</button><button role="menuitem" disabled={!props.dirty || props.studioDirty} onClick={() => choose(props.onReview)}><FileText size={16} />수정 내용 확인</button><button role="menuitem" onClick={() => choose(props.onHistory)}><History size={16} />버전 기록</button><button role="menuitem" onClick={() => choose(props.onFiles)}><FileText size={16} />이름 변경·복사</button><button role="menuitem" onClick={() => choose(props.onInfo)}><Info size={16} />문서 정보</button><button role="menuitem" disabled={props.studioDirty} onClick={() => choose(props.onSettings)}><SlidersHorizontal size={16} />고급 설정</button><button role="menuitem" onClick={() => choose(props.onHelp)}><HelpCircle size={16} />사용 안내</button></div>}
  </div>;
}
