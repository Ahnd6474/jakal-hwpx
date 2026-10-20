import { useEffect, useMemo, useRef, useState } from 'react';
import { ArrowRight, RotateCcw, X } from 'lucide-react';
import type { Document, Operation } from './api';
import type { Drafts } from './draft-history';
import { textTargets, type TextTarget } from './editing';

export function ChangeReview({ doc, drafts, onClose, onRevert, onGo }: {
  doc: Document; drafts: Drafts; onClose: () => void; onRevert: (items: TextTarget[]) => void; onGo: (operation: Operation) => void;
}) {
  const dialog = useRef<HTMLElement>(null);
  const [search, setSearch] = useState('');
  const [limit, setLimit] = useState(30);
  const items = useMemo(() => textTargets(doc, drafts).filter(item => item.text !== item.original), [doc, drafts]);
  const filtered = useMemo(() => {
    const needle = search.trim().toLocaleLowerCase();
    return items.filter(item => !needle || `${item.label}\n${item.original}\n${item.text}`.toLocaleLowerCase().includes(needle));
  }, [items, search]);
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    const overflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    const controls = () => [...(dialog.current?.querySelectorAll<HTMLElement>('button:not([disabled]), input') || [])];
    controls()[0]?.focus();
    function keyboard(event: KeyboardEvent) {
      if (event.key === 'Escape') onClose();
      if (event.key !== 'Tab') return;
      const fields = controls();
      if (!dialog.current?.contains(document.activeElement)) { event.preventDefault(); (event.shiftKey ? fields.at(-1) : fields[0])?.focus(); }
      else if (event.shiftKey && document.activeElement === fields[0]) { event.preventDefault(); fields.at(-1)?.focus(); }
      else if (!event.shiftKey && document.activeElement === fields.at(-1)) { event.preventDefault(); fields[0]?.focus(); }
    }
    window.addEventListener('keydown', keyboard);
    return () => { document.body.style.overflow = overflow; window.removeEventListener('keydown', keyboard); previous?.focus(); };
  }, [onClose]);
  return <div className="modal-backdrop" onClick={onClose}><section className="modal changes-modal" ref={dialog} role="dialog" aria-modal="true" aria-labelledby="changes-title" onClick={event => event.stopPropagation()}>
    <div className="modal-heading"><div><h2 id="changes-title">수정 내용 <small>{items.length}</small></h2><p>아직 문서에 저장하지 않은 변경입니다.</p></div><button className="icon-button" aria-label="수정 내용 닫기" onClick={onClose}><X size={20} /></button></div>
    <label className="changes-search"><span className="sr-only">수정 내용 검색</span><input aria-label="수정 내용 검색" value={search} onChange={event => { setSearch(event.target.value); setLimit(30); }} placeholder="변경 전후의 내용에서 찾기" /></label>
    <div className="changes-list">{filtered.slice(0, limit).map(item => <article className="change-item" key={item.id}>
      <header><strong>{item.label}</strong><div><button className="text-button" onClick={() => onGo(drafts[item.id])}>이동<ArrowRight size={13} /></button><button className="text-button" aria-label={`${item.label} 변경 취소`} onClick={() => onRevert([item])}><RotateCcw size={13} />되돌리기</button></div></header>
      <div className="change-comparison"><div><span>변경 전</span><p>{item.original || '(빈 내용)'}</p></div><div><span>변경 후</span><p>{item.text || '(내용 삭제)'}</p></div></div>
    </article>)}{!filtered.length && <p className="changes-empty">{items.length ? '일치하는 수정 내용이 없습니다.' : '모든 변경을 되돌렸습니다.'}</p>}
    {filtered.length > limit && <button className="outline-button changes-more" onClick={() => setLimit(limit + 30)}>다음 30개 보기 ({filtered.length - limit}개 남음)</button>}</div>
    <div className="modal-actions"><button className="outline-button" disabled={!items.length} onClick={() => onRevert(items)}><RotateCcw size={14} />전체 되돌리기</button><button className="primary-button" onClick={onClose}>편집 계속</button></div>
  </section></div>;
}
