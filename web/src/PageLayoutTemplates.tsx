import { lazy, Suspense, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Download, LayoutTemplate, Trash2, Upload, X } from 'lucide-react';
import { api, type Document } from './api';
import { useModalFocus } from './modal-focus';
import { backgroundProblem, defaultLayout, deletePagePreset, layoutProblem, pagePresets, readPagePresets, savePagePreset, type LayoutResponse, type LayoutSection, type PageLayout, type PagePreset, type BackgroundPage, type TemplateExtraction } from './page-layout';

import { pageTemplateAssets, type AssetPresetInfo } from './page-template-assets';
const ImportPageTemplate = lazy(() => import('./ImportPageTemplate').then(module => ({ default: module.ImportPageTemplate })));

export type LayoutApplication = { layout: PageLayout; section: number | null; create: boolean; title: string; background_pages?: BackgroundPage[] };
type Props = { doc: Document | null; dirty: boolean; busy: boolean; onClose: () => void; onApply: (value: LayoutApplication) => Promise<void> };

export function PageLayoutTemplates({ doc, dirty, busy, onClose, onApply }: Props) {
  const root = useRef<HTMLElement>(null), running = useRef(false);
  const [layout, setLayout] = useState<PageLayout>({ ...defaultLayout });
  const [sections, setSections] = useState<LayoutSection[]>([]);
  const [scope, setScope] = useState(-1);
  const [create, setCreate] = useState(!doc);
  const [title, setTitle] = useState('새 문서');
  const [name, setName] = useState('내 쪽 양식');
  const [custom, setCustom] = useState<(PagePreset | AssetPresetInfo)[]>([]);
  const [selected, setSelected] = useState('');
  const [importOpen, setImportOpen] = useState(false);
  const closeImport = useCallback(() => setImportOpen(false), []);
  const [extraction, setExtraction] = useState<TemplateExtraction | null>(null);
  const [importChoice, setImportChoice] = useState('all');
  const [rawPages, setRawPages] = useState<BackgroundPage[] | undefined>();
  const backgrounds = useMemo(() => rawPages === undefined ? undefined : importChoice === 'all' ? rawPages : rawPages.filter((_, i) => i === Number(importChoice)), [rawPages, importChoice]);
  const multiplePages = !!backgrounds && backgrounds.length > 1;
  const backgroundError = useMemo(() => backgrounds ? backgroundProblem(backgrounds) : null, [backgrounds]);
  const [pending, setPending] = useState(false);
  const [loading, setLoading] = useState(!!doc);
  const [failure, setFailure] = useState('');
  const [storageError, setStorageError] = useState('');
  const [notice, setNotice] = useState('');
  const [confirmDelete, setConfirmDelete] = useState(false);
  const disabled = busy || pending;
  const close = useCallback(() => { if (!disabled) onClose(); }, [disabled, onClose]);
  useModalFocus(root, close);
  useEffect(() => {
    let cancelled = false;
    let legacy: PagePreset[] = [];
    try { legacy = readPagePresets(); setCustom(legacy); }
    catch { setStorageError('내 쪽 양식 보관함을 읽지 못했습니다. 기본 양식은 사용할 수 있습니다.'); }
    void pageTemplateAssets.list().then(items => { if (!cancelled) setCustom([...legacy, ...items]); })
      .catch(() => { if (!cancelled) setStorageError('배경 양식 보관함을 열지 못했습니다. 기본 양식과 양식 파일은 계속 사용할 수 있습니다.'); });
    return () => { cancelled = true; };
  }, []);
  useEffect(() => {
    if (!doc) return;
    const abort = new AbortController();
    void api<LayoutResponse>(`/documents/${doc.id}/page-layout?revision=${doc.revision}`, { signal: abort.signal }).then(value => {
      if (!abort.signal.aborted) { setSections(value.sections); if (value.sections[0]) setLayout(value.sections[0].layout); }
    }).catch(e => { if (!abort.signal.aborted) setFailure(e.message); }).finally(() => { if (!abort.signal.aborted) setLoading(false); });
    return () => abort.abort();
  }, [doc?.id]);
  const current = sections.find(item => item.section === scope) || sections[0];
  const problem = layoutProblem(layout) || backgroundError;
  function change<K extends keyof PageLayout>(key: K, value: PageLayout[K]) { setLayout(previous => ({ ...previous, [key]: value })); setSelected(''); setConfirmDelete(false); }
  function imported(value: TemplateExtraction) {
    setExtraction(value); setRawPages(value.background_pages.length ? value.background_pages : undefined);
    setImportChoice(value.background_pages.length > 1 ? 'all' : '0');
    setLayout(value.choices[0].layout); setName(value.name); setSelected(''); setFailure('');
    if (value.background_pages.length > 1) setCreate(true);
    setNotice('양식을 가져왔습니다. 구역·쪽을 선택하고 미리보기를 확인한 뒤 적용하거나 보관하세요.');
  }
  function selectImport(value: string) {
    setImportChoice(value); setSelected('');
    const choice = extraction?.choices[value === 'all' ? 0 : Number(value)];
    if (choice) setLayout(choice.layout);
    if (value === 'all' && rawPages && rawPages.length > 1) setCreate(true);
  }
  async function choose(preset: PagePreset | AssetPresetInfo) {
    setFailure(''); setConfirmDelete(false);
    if ('asset' in preset) {
      setPending(true);
      try {
        const record = await pageTemplateAssets.read(preset.id);
        if (!record) throw new Error('보관한 배경 양식 파일을 찾지 못했습니다. 목록을 다시 열어 주세요.');
        imported({ name: record.name, source: 'json', choices: record.background_pages.map((page, i) => ({ label: `${i + 1}쪽`, layout: { ...record.layout, width_mm: page.width_mm, height_mm: page.height_mm }, notes: [] })), background_pages: record.background_pages, thumbnails: [], notes: ['원본의 글자와 그림을 포함한 배경 양식입니다. 배경 위에 새 내용을 작성합니다.'] });
        setLayout(record.layout); setSelected(preset.id);
      } catch (e) { setFailure(e instanceof Error ? e.message : '배경 양식을 읽지 못했습니다.'); }
      finally { setPending(false); }
    } else {
      setLayout({ ...preset.layout }); setSelected(preset.id); setName(preset.name); setRawPages(undefined); setExtraction(null);
    }
  }
  async function removePreset() {
    setPending(true);
    try {
      const item = custom.find(preset => preset.id === selected);
      if (item && 'asset' in item) await pageTemplateAssets.remove(item.id);
      else deletePagePreset(selected);
      setCustom(items => items.filter(item => item.id !== selected)); setSelected(''); setConfirmDelete(false);
    } catch { setStorageError('양식을 삭제하지 못했습니다.'); }
    finally { setPending(false); }
  }
  function numeric(key: keyof PageLayout, label: string, max = 100, min = 0) {
    const value = layout[key] as number;
    return <label className="page-layout-number">{label}<span><input aria-label={label} type="number" min={min} max={max} step="0.1" disabled={disabled || loading || (multiplePages && (key === 'width_mm' || key === 'height_mm'))} value={Number.isFinite(value) ? value : ''} onChange={event => change(key, event.target.valueAsNumber)} /><small>mm</small></span></label>;
  }
  async function submit(event: React.FormEvent) {
    event.preventDefault();
    if (disabled || running.current || problem || loading || (create && !title.trim())) return;
    running.current = true; setPending(true); setFailure('');
    try { await onApply({ layout, section: create || scope < 0 ? null : scope, create, title: title.trim(), ...(backgrounds !== undefined ? { background_pages: backgrounds } : {}) }); onClose(); }
    catch (e) { setFailure(e instanceof Error ? e.message : '쪽 양식을 적용하지 못했습니다.'); }
    finally { running.current = false; setPending(false); }
  }
  async function savePreset() {
    setPending(true);
    try {
      if (custom.length >= 20) throw new Error('내 쪽 양식은 최대 20개까지 보관할 수 있습니다. 사용하지 않는 양식을 정리해 주세요.');
      if (backgrounds?.length) {
        const saved = await pageTemplateAssets.save(name, layout, backgrounds);
        setCustom(items => [...items, saved]); setSelected(saved.id);
      } else setCustom([...savePagePreset(name, layout), ...custom.filter(item => 'asset' in item)]);
      setStorageError(''); setNotice(backgrounds?.length ? '배경 페이지를 포함한 양식을 이 브라우저에 보관했습니다.' : '쪽 양식을 이 브라우저에 보관했습니다. 본문 내용은 포함하지 않습니다.');
    } catch (e) { setStorageError(e instanceof Error ? e.message : '쪽 양식을 보관하지 못했습니다.'); }
    finally { setPending(false); }
  }
  function exportPreset() {
    if (problem) return;
    const url = URL.createObjectURL(new Blob([JSON.stringify({ version: backgrounds?.length ? 2 : 1, name: name.trim() || '쪽 양식', layout, ...(backgrounds?.length ? { background_pages: backgrounds } : {}) }, null, 2)], { type: 'application/json' }));
    const link = document.createElement('a'); link.href = url; link.download = `${(name.trim() || '쪽 양식').replace(/[\\/:*?"<>|]/g, '_')}.jakal-layout.json`; link.click();
    setTimeout(() => URL.revokeObjectURL(url), 30_000);
  }
  // Unsupported source columns remain editable, but never allocate an unbounded preview.
  const previewColumns = [1, 2, 3].includes(layout.columns) ? layout.columns : 1;
  const orientation = layout.width_mm > layout.height_mm;
  const paper = [[210, 297], [297, 420], [216, 279]].find(([w, h]) => Math.abs(Math.min(layout.width_mm, layout.height_mm) - w) < 0.1 && Math.abs(Math.max(layout.width_mm, layout.height_mm) - h) < 0.1);
  return <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) close(); }}><section className="modal page-layout-modal" ref={root} role="dialog" aria-modal="true" aria-labelledby="page-layout-title">
    <div className="modal-heading"><div><h2 id="page-layout-title">쪽 양식·템플릿</h2><p>용지, 여백, 단 구분선과 머리말·꼬리말을 설정합니다.</p></div><button className="icon-button" aria-label="쪽 양식 닫기" disabled={disabled} onClick={close}><X size={19} /></button></div>
    <div className="page-layout-workspace"><aside className="page-preset-list" aria-label="쪽 양식 목록"><h3>기본 양식</h3>{pagePresets.map(preset => <button type="button" key={preset.id} disabled={disabled || loading} aria-pressed={selected === preset.id} onClick={() => void choose(preset)}><LayoutTemplate size={16} />{preset.name}</button>)}<h3>내 쪽 양식 <small>{custom.length}/20</small></h3>
      {custom.length ? custom.map(preset => <button key={preset.id} type="button" disabled={disabled || loading} aria-pressed={selected === preset.id} onClick={() => void choose(preset)}>{preset.name}</button>) : <p>현재 설정을 이름을 붙여 보관할 수 있습니다.</p>}
      {custom.some(item => item.id === selected) && (confirmDelete ? <div className="page-preset-confirm"><span>이 양식을 보관함에서 삭제할까요?</span><button type="button" disabled={disabled} onClick={() => void removePreset()}>삭제 확인</button><button type="button" disabled={disabled} onClick={() => setConfirmDelete(false)}>취소</button></div> : <button type="button" disabled={disabled} onClick={() => setConfirmDelete(true)}><Trash2 size={14} />선택한 내 양식 삭제</button>)}
      <button type="button" disabled={disabled || loading} onClick={() => setImportOpen(true)}><Upload size={14} />양식 파일 가져오기</button>

    </aside><form className="page-layout-form" onSubmit={event => void submit(event)}>
      {loading && <p role="status">현재 쪽 설정을 읽고 있습니다.</p>}
      {extraction && <section className="imported-page-template" aria-label="가져온 양식 정보"><strong>{extraction.name}</strong><label>{rawPages?.length ? '사용할 쪽' : '가져올 구역'}<select aria-label="가져온 양식 구역·쪽 선택" disabled={disabled} value={importChoice} onChange={e => selectImport(e.target.value)}>{rawPages && rawPages.length > 1 && <option value="all">전체 {rawPages.length}쪽 · 새 문서로 만들기</option>}{extraction.choices.map((choice, index) => <option key={index} value={String(index)}>{choice.label}</option>)}</select></label>
        {multiplePages && <p>여러 쪽은 각각의 용지 크기를 유지해 새 문서로 만듭니다. 현재 문서에 적용하려면 한 쪽을 선택하세요.</p>}
        {extraction.notes.map((note, index) => <p key={index}>{note}</p>)}{(extraction.choices[importChoice === 'all' ? 0 : Number(importChoice)]?.notes || []).map((note, index) => <p key={`choice-${index}`}>{note}</p>)}
        {!!rawPages?.length && <div className="imported-page-thumbnails">{rawPages.map((page, index) => <button key={index} type="button" disabled={disabled} aria-pressed={importChoice === String(index)} onClick={() => selectImport(String(index))}><img loading="lazy" decoding="async" src={extraction.thumbnails[index] || page.image} alt={`${index + 1}쪽 배경 미리보기`} /><span>{index + 1}쪽</span></button>)}</div>}
      </section>}
      {(backgrounds?.length || (backgrounds === undefined && current?.has_background)) && <div className="page-background-actions"><span>{backgrounds?.length ? '페이지 배경이 포함됩니다. 배경 글자는 개별 편집되지 않습니다.' : '현재 문서의 배경 양식을 유지합니다.'}</span><button type="button" className="text-button" disabled={disabled} onClick={() => { setRawPages([]); setExtraction(null); setSelected(''); }}>배경 양식 제거</button></div>}
      {backgrounds?.length === 0 && <p className="page-layout-note">적용하면 기존에 가져온 배경 양식을 제거합니다. 본문과 다른 머리말·꼬리말 개체는 유지합니다.</p>}
      <div className="page-layout-settings"><div className="page-layout-fields"><fieldset disabled={disabled || loading}><legend>용지와 여백</legend><div className="page-layout-grid"><label>용지<select aria-label="용지 종류" disabled={multiplePages} value={paper ? paper.join(',') : 'custom'} onChange={e => { if (e.target.value === 'custom') return; const [w, h] = e.target.value.split(',').map(Number); setLayout(value => ({ ...value, width_mm: orientation ? h : w, height_mm: orientation ? w : h })); setSelected(''); }}><option value="210,297">A4</option><option value="297,420">A3</option><option value="216,279">Letter (216×279)</option><option value="custom">직접 지정</option></select></label><label>방향<select aria-label="용지 방향" disabled={multiplePages} value={orientation ? 'landscape' : 'portrait'} onChange={e => { if ((e.target.value === 'landscape') !== orientation) { setLayout(value => ({ ...value, width_mm: value.height_mm, height_mm: value.width_mm })); setSelected(''); } }}><option value="portrait">세로</option><option value="landscape">가로</option></select></label>
      {numeric('width_mm', '용지 너비', 600, 80)}{numeric('height_mm', '용지 높이', 800, 80)}{numeric('left_mm', '왼쪽 여백')}{numeric('right_mm', '오른쪽 여백')}{numeric('top_mm', '위쪽 여백')}{numeric('bottom_mm', '아래쪽 여백')}</div></fieldset>
      <fieldset disabled={disabled || loading}><legend>단 나누기</legend><div className="page-layout-grid"><label>단 수<select aria-label="단 수" value={layout.columns} onChange={e => change('columns', Number(e.target.value) as 1 | 2 | 3)}><option value="1">1단</option><option value="2">2단</option><option value="3">3단</option></select></label>{numeric('gap_mm', '단 사이 간격', 40)}</div><label className="page-layout-check"><input type="checkbox" checked={layout.separator} disabled={layout.columns === 1} onChange={e => change('separator', e.target.checked)} />단 사이에 세로 구분선 표시</label></fieldset>
      <fieldset disabled={disabled || loading}><legend>머리말·꼬리말과 쪽 번호</legend>{(['header_text', 'footer_text'] as const).map((key, index) => <div className="page-running-text" key={key}><label>{index ? '꼬리말' : '머리말'}<select aria-label={index ? '꼬리말 처리' : '머리말 처리'} value={layout[key] === null ? 'keep' : 'replace'} onChange={e => change(key, e.target.value === 'keep' ? null : (index ? current?.footer_text : current?.header_text)?.slice(0, 200) || '')}><option value="keep">기존 내용·서식 유지</option><option value="replace">직접 입력 / 비우면 제거</option></select></label>{layout[key] !== null && <input aria-label={index ? '꼬리말 내용' : '머리말 내용'} maxLength={200} value={layout[key] || ''} placeholder={index ? '예: 학교명 · 작성일' : '예: 문서 제목'} onChange={e => change(key, e.target.value)} />}</div>)}
      <div className="page-layout-grid">{numeric('header_mm', '머리말 여백', 60)}{numeric('footer_mm', '꼬리말 여백', 60)}</div><label className="page-number-label">쪽 번호<select aria-label="쪽 번호" value={layout.page_numbers} onChange={e => change('page_numbers', e.target.value as PageLayout['page_numbers'])}><option value="keep">기존 설정 유지</option><option value="none">없음</option><option value="bottom-center">아래 가운데</option><option value="bottom-right">아래 오른쪽</option></select></label></fieldset></div>
      <aside className="page-layout-preview"><div className={`page-layout-sheet ${backgrounds?.length ? 'has-background' : ''}`} style={{ aspectRatio: `${layout.width_mm || 210} / ${layout.height_mm || 297}` }} aria-label="쪽 구조 예시">{backgrounds?.[0] && <img className="page-preview-image" src={backgrounds[0].image} alt="가져온 페이지 배경" />}<div className="page-preview-header">{layout.header_text ?? current?.header_text}</div><div className="page-preview-body" style={{ left: `${layout.left_mm / layout.width_mm * 100}%`, right: `${layout.right_mm / layout.width_mm * 100}%`, top: `${(layout.top_mm + layout.header_mm) / layout.height_mm * 100}%`, bottom: `${(layout.bottom_mm + layout.footer_mm) / layout.height_mm * 100}%`, gridTemplateColumns: `repeat(${previewColumns}, minmax(0, 1fr))`, gap: `${layout.gap_mm / layout.width_mm * 100}%` }}>{Array.from({ length: previewColumns }, (_, n) => <div className="page-preview-column" key={n}>{Array.from({ length: 14 }, (_, i) => <i key={i} />)}</div>)}{layout.separator && previewColumns > 1 && Array.from({ length: previewColumns - 1 }, (_, n) => <b className="page-preview-divider" key={n} style={{ left: `${(n + 1) / previewColumns * 100}%` }} />)}</div><div className="page-preview-footer">{layout.footer_text ?? current?.footer_text}</div>{layout.page_numbers !== 'none' && layout.page_numbers !== 'keep' && <span className={`page-preview-number ${layout.page_numbers}`}>1</span>}</div><p>쪽 구조 예시<br />실제 내용과 줄바꿈은 적용 후 문서에서 확인하세요.</p><p>머리말·꼬리말을 직접 입력하면 기존 해당 영역의 서식을 바꿉니다. 본문의 표·그림 크기는 유지됩니다.</p></aside></div>
      <div className="page-preset-save"><label>양식 이름<input aria-label="보관할 쪽 양식 이름" maxLength={60} value={name} disabled={disabled} onChange={e => setName(e.target.value)} /></label><button type="button" className="outline-button" disabled={disabled || !!problem || !name.trim() || loading} onClick={() => void savePreset()}>내 양식으로 보관</button><button type="button" className="icon-button" aria-label="쪽 양식 파일 받기" disabled={disabled || !!problem || loading} onClick={exportPreset}><Download size={16} /></button></div>
      {notice && <p className="page-layout-note" role="status">{notice}</p>}{storageError && <p className="tool-error" role="status">{storageError}</p>}{failure && <p className="tool-error" role="alert">{failure}</p>}{problem && <p className="tool-error" role="alert">{problem}</p>}
      <div className="page-layout-target"><label>적용 대상<select aria-label="쪽 양식 적용 대상" disabled={disabled} value={create ? 'new' : 'current'} onChange={e => setCreate(e.target.value === 'new')}>{doc && <option value="current" disabled={multiplePages}>현재 문서</option>}<option value="new">빈 새 문서</option></select></label>{create ? <label>새 문서 제목<input aria-label="새 양식 문서 제목" required maxLength={100} value={title} disabled={disabled} onChange={e => setTitle(e.target.value)} /></label> : <label>적용 범위<select aria-label="쪽 양식 적용 범위" disabled={disabled || loading} value={scope} onChange={e => { const next = Number(e.target.value); setScope(next); const existing = sections.find(item => item.section === next); if (existing) { setLayout(existing.layout); setSelected(''); } }}><option value="-1">전체 구역</option>{sections.map(item => <option key={item.section} value={item.section}>구역 {item.section + 1}</option>)}</select></label>}</div>
      <p className="page-layout-note">{dirty ? '현재 미저장 편집을 먼저 원래 문서에 저장한 뒤 진행합니다. ' : ''}선택한 구역 안의 단 나누기 설정은 이 양식으로 통일합니다. ‘기존 유지’ 항목은 적용받는 문서의 현재 내용을 유지합니다.{!create && ' 적용하면 새 버전으로 저장되며, 이전 버전은 기록에서 복원할 수 있습니다.'}</p>
      <div className="modal-actions"><button type="button" className="outline-button" disabled={disabled} onClick={close}>취소</button><button className="primary-button" disabled={disabled || !!problem || loading || (create && !title.trim())}>{disabled ? '양식 적용 중…' : create ? '이 양식으로 새 문서' : dirty ? '문서 저장 후 양식 적용' : '현재 문서에 적용'}</button></div>
    </form></div>
    {importOpen && <Suspense fallback={<p role="status">양식 가져오기 도구를 열고 있습니다.</p>}><ImportPageTemplate onClose={closeImport} onImport={imported} /></Suspense>}
  </section></div>;
}
