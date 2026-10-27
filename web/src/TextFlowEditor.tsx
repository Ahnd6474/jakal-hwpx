import { forwardRef, memo, useCallback, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { ArrowRight, ChevronDown, ChevronUp, LockKeyhole } from 'lucide-react';
import type { Operation, Paragraph, Section } from './api';
import { fitTextareas } from './textarea-layout';

type Props = {
  sections: Section[];
  drafts: Record<string, Operation>;
  search: string;
  disabled: boolean;
  onChange: (key: string, operation: Operation, original: string) => void;
  onSelect: (paragraph: Paragraph) => void;
  onOriginal: () => void;
};
export type TextFlowHandle = { find: (direction: number) => void };

function FlowField({ paragraph, text, disabled, onChange, onSelect, onMove }: {
  paragraph: Paragraph; text: string; disabled: boolean;
  onChange: (text: string) => void; onSelect: () => void; onMove: (direction: number) => void;
}) {
  const field = useRef<HTMLTextAreaElement>(null);
  useLayoutEffect(() => { if (field.current) fitTextareas([field.current]); }, [text]);
  return <textarea ref={field} className="flow-field" rows={1} aria-label={`${paragraph.section + 1}섹션 ${paragraph.index + 1}문단`} value={text} maxLength={100000} disabled={disabled} placeholder="내용을 입력하세요" onFocus={onSelect} onChange={event => onChange(event.target.value)} onKeyDown={event => {
    if (event.nativeEvent.isComposing || event.shiftKey || event.ctrlKey || event.metaKey || event.altKey) return;
    const element = event.currentTarget;
    if (element.selectionStart !== element.selectionEnd) return;
    if (event.key === 'ArrowUp' && element.selectionStart === 0) { event.preventDefault(); onMove(-1); }
    if (event.key === 'ArrowDown' && element.selectionEnd === text.length) { event.preventDefault(); onMove(1); }
  }} />;
}

function MarkedText({ text, query }: { text: string; query: string }) {
  if (!query) return <>{text}</>;
  const start = text.toLocaleLowerCase().indexOf(query.toLocaleLowerCase());
  if (start < 0) return <>{text}</>;
  return <>{text.slice(0, start)}<mark>{text.slice(start, start + query.length)}</mark>{text.slice(start + query.length)}</>;
}

const FlowParagraph = memo(function FlowParagraph({ paragraph, text, changed, disabled, query, matched, current, onChange, onSelect, onMove, register }: {
  paragraph: Paragraph; text: string; changed: boolean; disabled: boolean; query: string; matched: boolean; current: boolean;
  onChange: Props['onChange']; onSelect: Props['onSelect'];
  onMove: (id: string, direction: number) => void;
  register: (id: string, element: HTMLDivElement | null) => void;
}) {
  const paragraphRef = useCallback((element: HTMLDivElement | null) => register(paragraph.id, element), [paragraph.id, register]);
  return <div ref={paragraphRef} className={`flow-paragraph ${changed ? 'flow-changed' : ''} ${matched ? 'flow-match' : ''} ${current ? 'flow-current-match' : ''}`}>
    {paragraph.editable ? <FlowField paragraph={paragraph} text={text} disabled={disabled} onSelect={() => onSelect(paragraph)} onChange={next => onChange(paragraph.id, { type: 'paragraph', section: paragraph.section, index: paragraph.index, text: next }, paragraph.text)} onMove={direction => onMove(paragraph.id, direction)} /> : <p className="flow-readonly" aria-label={`읽기 전용 · ${paragraph.section + 1}섹션 ${paragraph.index + 1}문단`} title="복합 개체가 포함된 내용입니다. 원문 보기에서 확인하세요."><MarkedText text={text} query={query} /></p>}
  </div>;
});

export const TextFlowEditor = forwardRef<TextFlowHandle, Props>(function TextFlowEditor({ sections, drafts, search, disabled, onChange, onSelect, onOriginal }, ref) {
  const [fontSize, setFontSize] = useState(16);
  const [hit, setHit] = useState('');
  const canvas = useRef<HTMLDivElement>(null);
  const paper = useRef<HTMLDivElement>(null);
  const paragraphs = useRef(new Map<string, HTMLDivElement>());
  const anchor = useRef<{ id: string; top: number } | null>(null);
  const query = search.trim();
  // Keep read-only text in the reading flow. Empty controls and their repeated
  // fallback messages belong in the original view, not between every sentence.
  const groups = useMemo(() => sections.map(section => ({
    index: section.index,
    paragraphs: section.paragraphs.map(paragraph => ({ paragraph, text: drafts[paragraph.id]?.text ?? paragraph.text, changed: !!drafts[paragraph.id] }))
      .filter(item => item.text.trim() || item.changed),
  })).filter(section => section.paragraphs.length), [sections, drafts]);
  const all = useMemo(() => groups.flatMap(group => group.paragraphs), [groups]);
  const matches = useMemo(() => {
    const needle = query.toLocaleLowerCase();
    return needle ? all.filter(item => item.text.toLocaleLowerCase().includes(needle)) : [];
  }, [all, query]);
  const matchedIds = useMemo(() => new Set(matches.map(item => item.paragraph.id)), [matches]);
  const editableIds = useRef<string[]>([]);
  useLayoutEffect(() => { editableIds.current = all.filter(item => item.paragraph.editable).map(item => item.paragraph.id); }, [all]);
  const register = useCallback((id: string, element: HTMLDivElement | null) => {
    if (element) paragraphs.current.set(id, element);
    else paragraphs.current.delete(id);
  }, []);
  const hitIndex = matches.findIndex(item => item.paragraph.id === hit);
  function find(direction: number) {
    if (!matches.length) return;
    const index = hitIndex < 0 ? direction > 0 ? 0 : matches.length - 1 : (hitIndex + direction + matches.length) % matches.length;
    const id = matches[index].paragraph.id;
    setHit(id);
    paragraphs.current.get(id)?.scrollIntoView({ block: 'center' });
  }
  useImperativeHandle(ref, () => ({ find }));
  const move = useCallback((id: string, direction: number) => {
    const index = editableIds.current.indexOf(id);
    if (index < 0) return;
    const next = editableIds.current[index + direction];
    if (!next) return;
    const field = paragraphs.current.get(next)?.querySelector('textarea');
    field?.focus();
    const position = direction < 0 ? field?.value.length ?? 0 : 0;
    field?.setSelectionRange(position, position);
  }, []);
  useLayoutEffect(() => {
    if (paper.current) fitTextareas(paper.current.querySelectorAll<HTMLTextAreaElement>('textarea'));
    if (anchor.current && canvas.current) {
      const element = paragraphs.current.get(anchor.current.id);
      if (element) canvas.current.scrollTop += element.getBoundingClientRect().top - anchor.current.top;
      anchor.current = null;
    }
  }, [fontSize]);
  useLayoutEffect(() => {
    const element = paper.current;
    if (!element) return;
    let previousWidth = element.clientWidth;
    let frame = 0;
    let active = true;
    const scheduleFit = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => fitTextareas(element.querySelectorAll<HTMLTextAreaElement>('textarea')));
    };
    const observer = new ResizeObserver(() => {
      if (element.clientWidth === previousWidth) return;
      previousWidth = element.clientWidth;
      scheduleFit();
    });
    observer.observe(element);
    // Font swaps can change line wrapping without changing the paper's width.
    void document.fonts.ready.then(() => { if (active) scheduleFit(); });
    document.fonts.addEventListener('loadingdone', scheduleFit);
    return () => { active = false; observer.disconnect(); document.fonts.removeEventListener('loadingdone', scheduleFit); cancelAnimationFrame(frame); };
  }, []);

  return <div ref={canvas} tabIndex={0} className="document-canvas flow-canvas" aria-label="이어지는 본문 편집">
    <div className="flow-reading-tools">
      <span>텍스트를 이어서 읽고 수정하세요</span>
      <div><label><span className="sr-only">본문 글자 크기</span><select aria-label="본문 글자 크기" value={fontSize} onChange={event => {
        const top = canvas.current?.getBoundingClientRect().top ?? 0;
        const visible = [...paragraphs.current.entries()].find(([, element]) => element.getBoundingClientRect().bottom > top + 50);
        if (visible) anchor.current = { id: visible[0], top: visible[1].getBoundingClientRect().top };
        setFontSize(Number(event.target.value));
      }}>{[16, 18, 20].map(size => <option key={size} value={size}>{size === 16 ? '기본 글씨' : size === 18 ? '큰 글씨' : '더 큰 글씨'}</option>)}</select></label><button className="text-button" onClick={onOriginal}>원문 보기<ArrowRight size={13} /></button></div>
    </div>
    {query && <div className="flow-search-results" role="status"><span>{matches.length ? `${matches.length}개 구간 · ${hitIndex < 0 ? 'Enter로 이동' : `${hitIndex + 1} / ${matches.length}`}` : '일치하는 내용이 없습니다'}<small>주변 본문을 함께 표시합니다</small></span><div><button className="icon-button" aria-label="이전 본문 검색 결과" disabled={!matches.length} onClick={() => find(-1)}><ChevronUp size={16} /></button><button className="icon-button" aria-label="다음 본문 검색 결과" disabled={!matches.length} onClick={() => find(1)}><ChevronDown size={16} /></button></div></div>}
    <div ref={paper} className="document-paper flow-paper" style={{ '--flow-font-size': `${fontSize}px` } as React.CSSProperties}>
      <p className="flow-context-note"><LockKeyhole size={12} />수식·그림이 포함된 내용은 읽기 전용이며, 정확한 배치는 원문 보기에서 확인할 수 있습니다.</p>
      {groups.length ? groups.map(group => <section className="flow-section" key={group.index} aria-label={`${group.index + 1}섹션 본문`}>
        {groups.length > 1 && <div className="flow-section-label">섹션 {group.index + 1}</div>}
        {group.paragraphs.map(({ paragraph, text, changed }) => <FlowParagraph key={paragraph.id} paragraph={paragraph} text={text} changed={changed} disabled={disabled} query={query} matched={matchedIds.has(paragraph.id)} current={matchedIds.has(paragraph.id) && hit === paragraph.id} onChange={onChange} onSelect={onSelect} onMove={move} register={register} />)}
      </section>) : <div className="empty-state"><p>이 범위에는 편집할 본문 텍스트가 없습니다.</p><button className="text-button" onClick={onOriginal}>표·수식·그림을 원문으로 보기<ArrowRight size={14} /></button></div>}
    </div>
  </div>;
});
