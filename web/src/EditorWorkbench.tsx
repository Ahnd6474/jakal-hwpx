import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ArrowRight, LoaderCircle, Search, Star, X } from 'lucide-react';
import { commandGroups, quickCommands, taskGroups, visibleCommands, type EditorCommand } from './workbench';
import { useModalFocus } from './modal-focus';

type Props = { onClose: () => void; commands: () => Promise<EditorCommand[]>; onRun: (id: string) => Promise<void> };
function readFavorites(): string[] {
  try { const value: unknown = JSON.parse(localStorage.getItem('jakal-tool-favorites') || '[]'); return Array.isArray(value) ? value.filter((id): id is string => typeof id === 'string').slice(0, 12) : []; } catch { return []; }
}
export function EditorWorkbench({ onClose, commands, onRun }: Props) {
  const root = useRef<HTMLElement>(null);
  const [items, setItems] = useState<EditorCommand[]>([]);
  const [pending, setPending] = useState(true);
  const [error, setError] = useState('');
  const [query, setQuery] = useState('');
  const [group, setGroup] = useState('');
  const [task, setTask] = useState('');
  const [favorites, setFavorites] = useState(readFavorites);
  const close = useCallback(() => onClose(), [onClose]);
  useModalFocus(root, close);
  useEffect(() => {
    let cancelled = false;
    void commands().then(value => { if (!cancelled) setItems(value); }).catch(e => { if (!cancelled) setError(e.message); }).finally(() => { if (!cancelled) setPending(false); });
    return () => { cancelled = true; };
  }, [commands]);
  const filtered = useMemo(() => visibleCommands(items, query, group).filter(item => !task || taskGroups.find(value => value.id === task)?.commands.includes(item.id)), [items, query, group, task]);
  const pinned = visibleCommands(items, '', '').filter(item => favorites.includes(item.id));
  function pin(id: string) {
    const next = favorites.includes(id) ? favorites.filter(value => value !== id) : [...favorites, id].slice(-12);
    setFavorites(next); try { localStorage.setItem('jakal-tool-favorites', JSON.stringify(next)); } catch { /* Preferences are optional. */ }
  }
  function command(item: EditorCommand) {
    return <div className="workbench-command" key={item.id}><button className="workbench-run" disabled={!item.enabled || pending} onClick={() => void onRun(item.id).catch(e => setError(e.message))}><span>{item.label}<small>{commandGroups[item.id.split(':')[0]]}{item.opensDialog ? ' · 설정 창' : ''}{!item.enabled ? ' · 현재 사용할 수 없음' : ''}</small></span>{item.shortcutLabel && <kbd>{item.shortcutLabel}</kbd>}<ArrowRight size={13} /></button><button className="workbench-pin icon-button" aria-label={`${item.label} ${favorites.includes(item.id) ? '즐겨찾기 해제' : '즐겨찾기'}`} aria-pressed={favorites.includes(item.id)} onClick={() => pin(item.id)}><Star size={13} fill={favorites.includes(item.id) ? 'currentColor' : 'none'} /></button></div>;
  }
  return <div className="modal-backdrop workbench-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) close(); }}><section ref={root} className="modal workbench-modal" role="dialog" aria-modal="true" aria-labelledby="workbench-title">
    <header className="workbench-heading"><div><h2 id="workbench-title">작업 도구</h2><p>필요한 기능을 찾고, 문서의 현재 위치에서 실행하세요.</p></div><button className="icon-button" aria-label="작업 도구 닫기" onClick={close}><X size={19} /></button></header>
    {error && <p className="tool-error" role="alert">{error}</p>}
    {!query && !group && !task && <div className="workbench-quick" role="group" aria-label="기본 편집 도구">{quickCommands.map(id => items.find(item => item.id === id)).filter((item): item is EditorCommand => !!item).map(item => <button key={item.id} disabled={!item.enabled || pending} onClick={() => void onRun(item.id).catch(e => setError(e.message))}>{item.label}</button>)}</div>}
      <label className="workbench-search"><Search size={17} /><input data-autofocus aria-label="작업 도구 검색" placeholder="예: 수식, 셀 병합, 각주, 여백…" value={query} onChange={e => { setQuery(e.target.value); setTask(''); }} onKeyDown={e => {
        if (e.key === 'Enter') { const first = filtered.find(item => item.enabled); if (first) { e.preventDefault(); void onRun(first.id).catch(err => setError(err.message)); } }
        if (e.key === 'ArrowDown') { e.preventDefault(); root.current?.querySelector<HTMLButtonElement>('.workbench-results .workbench-run:not(:disabled)')?.focus(); }
      }} />{query && <button className="icon-button" aria-label="작업 검색 초기화" onClick={() => setQuery('')}><X size={14} /></button>}</label>
      <div className="workbench-tasks" role="group" aria-label="작업별 기능">{taskGroups.map(value => <button aria-pressed={task === value.id} key={value.id} onClick={() => { setTask(task === value.id ? '' : value.id); setGroup(''); }}>{value.label}</button>)}</div>
      <div className="workbench-groups" role="group" aria-label="기능 분류"><button aria-pressed={!group} onClick={() => { setGroup(''); setTask(''); }}>전체</button>{Object.entries(commandGroups).map(([id, name]) => <button key={id} aria-pressed={group === id} onClick={() => { setGroup(id); setTask(''); }}>{name}</button>)}</div>
      <div className="workbench-results" onKeyDown={e => {
        if (!['ArrowDown', 'ArrowUp'].includes(e.key)) return;
        const buttons = [...(root.current?.querySelectorAll<HTMLButtonElement>('.workbench-results .workbench-run:not(:disabled)') || [])];
        const at = buttons.indexOf(document.activeElement as HTMLButtonElement); if (at < 0 || !buttons.length) return;
        e.preventDefault(); buttons[(at + (e.key === 'ArrowUp' ? -1 : 1) + buttons.length) % buttons.length]?.focus();
      }}>
        {!!pinned.length && !query && !group && !task && <section><h3>즐겨찾기</h3>{pinned.map(command)}</section>}
        <section><h3>{query ? `검색 결과 ${filtered.length}` : `${taskGroups.find(value => value.id === task)?.label || '모든 기능'} ${filtered.length}`}</h3>{pending ? <p className="workbench-empty" role="status"><LoaderCircle size={17} className="spin" />현재 위치의 도구를 확인하고 있습니다.</p> : filtered.length ? filtered.map(command) : <p className="workbench-empty">일치하는 기능이 없습니다. 다른 이름으로 검색하세요.</p>}</section>
      </div>
      <footer className="workbench-footer"><span>↑ ↓ 이동 · Enter 실행 · Esc 닫기</span><span>별표로 자주 쓰는 도구를 모으세요.</span></footer>
  </section></div>;
}
