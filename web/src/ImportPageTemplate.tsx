import { useCallback, useRef, useState, useEffect } from 'react';
import { ArrowDown, ArrowUp, LoaderCircle, Plus, Trash2, X } from 'lucide-react';
import { api } from './api';
import { useModalFocus } from './modal-focus';
import type { TemplateExtraction } from './page-layout';

export function ImportPageTemplate({ onClose, onImport }: { onClose: () => void; onImport: (value: TemplateExtraction) => void }) {
  const root = useRef<HTMLElement>(null), input = useRef<HTMLInputElement>(null), abort = useRef<AbortController | null>(null);
  const [files, setFiles] = useState<File[]>([]), [busy, setBusy] = useState(false), [error, setError] = useState('');
  const close = useCallback(() => { abort.current?.abort(); onClose(); }, [onClose]);
  useModalFocus(root, close);
  useEffect(() => () => abort.current?.abort(), []);
  function choose(selected: File[]) {
    if (!selected.length) return;
    setError('');
    if (selected.length > 10 || selected.some(file => !file.size) || selected.reduce((sum, file) => sum + file.size, 0) > 20 * 1024 * 1024) { setError('파일은 최대 10개, 합계 20MB 이하여야 하며 빈 파일은 선택할 수 없습니다.'); return; }
    const images = selected.every(file => /\.(jpe?g|png)$/i.test(file.name));
    if ((!images && selected.length !== 1) || selected.some(file => !/\.(hwpx|pdf|jpe?g|png|json)$/i.test(file.name))) { setError('HWPX·PDF·JSON은 한 파일씩, JPG·PNG는 여러 장을 함께 선택해 주세요.'); return; }
    setFiles(selected);
  }
  function move(at: number, by: number) {
    setFiles(current => { const result = [...current]; [result[at], result[at + by]] = [result[at + by], result[at]]; return result; });
  }
  async function extract() {
    if (busy || !files.length) return;
    setBusy(true); setError('');
    const controller = new AbortController(); abort.current = controller;
    try {
      const form = new FormData(); files.forEach(file => form.append('files', file));
      const value = await api<TemplateExtraction>('/page-templates/extract', { method: 'POST', body: form, signal: controller.signal });
      if (!controller.signal.aborted) { onImport(value); onClose(); }
    } catch (e) { if (!controller.signal.aborted) setError(e instanceof Error ? e.message : '양식을 추출하지 못했습니다.'); }
    finally { if (!controller.signal.aborted) setBusy(false); }
  }
  return <div className="modal-backdrop template-import-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) close(); }}><section className="modal template-import-modal" ref={root} role="dialog" aria-modal="true" aria-labelledby="template-import-title">
    <div className="modal-heading"><div><h2 id="template-import-title">파일에서 쪽 양식 가져오기</h2><p>HWPX · PDF · JPG/PNG 모음 · 양식 JSON</p></div><button className="icon-button" aria-label="양식 가져오기 닫기" onClick={close}><X size={18} /></button></div>
    <p className="page-layout-note">HWPX는 쪽 설정과 머리말·꼬리말 텍스트를 추출합니다. PDF·이미지는 원래 글자와 그림이 포함된 배경 양식으로 가져오며 그 위에 새 내용을 작성합니다.</p>
    <input ref={input} type="file" multiple hidden aria-label="가져올 양식 파일 선택" accept=".hwpx,.pdf,.jpg,.jpeg,.png,.json" onChange={event => { choose(Array.from(event.target.files || [])); event.target.value = ''; }} />
    <button className="outline-button" disabled={busy} onClick={() => input.current?.click()}><Plus size={15} />{files.length ? '파일 다시 선택' : '파일 선택'}</button>
    {files.length > 1 && <button className="text-button" disabled={busy} onClick={() => setFiles(current => [...current].sort((a, b) => a.name.localeCompare(b.name, 'ko', { numeric: true })))}>파일명 순으로 정렬</button>}
    <ol className="template-import-files" aria-label="가져올 쪽 순서">{files.map((file, index) => <li key={`${index}-${file.name}`}><span><b>{index + 1}.</b> {file.name}<small>{(file.size / 1024).toLocaleString('ko-KR', { maximumFractionDigits: 0 })} KB</small></span><button className="icon-button" aria-label={`${file.name} 위로`} disabled={busy || index === 0} onClick={() => move(index, -1)}><ArrowUp size={14} /></button><button className="icon-button" aria-label={`${file.name} 아래로`} disabled={busy || index === files.length - 1} onClick={() => move(index, 1)}><ArrowDown size={14} /></button><button className="icon-button" aria-label={`${file.name} 제외`} disabled={busy} onClick={() => setFiles(current => current.filter((_, at) => at !== index))}><Trash2 size={14} /></button></li>)}</ol>
    <p className="page-layout-note">파일 합계 20MB · PDF/이미지 최대 10쪽. 여러 이미지의 표시 순서가 쪽 순서가 됩니다. 추출 완료 전에는 현재 문서를 변경하지 않습니다.</p>
    {error && <p className="tool-error" role="alert">{error}</p>}{busy && <p role="status"><LoaderCircle size={14} className="spin" /> 용지와 양식을 분석하고 미리보기를 준비하고 있습니다.</p>}
    <div className="modal-actions"><button className="outline-button" onClick={close}>{busy ? '가져오기 취소' : '취소'}</button><button className="primary-button" disabled={busy || !files.length} onClick={() => void extract()}>{busy ? '추출 중…' : '양식 추출'}</button></div>
  </section></div>;
}
