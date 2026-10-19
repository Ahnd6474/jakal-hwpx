import { useEffect, useMemo, useRef, useState } from 'react';
import { ArrowRight, Check, LockKeyhole, Save, X } from 'lucide-react';
import { api, type Document, type Operation, type Settings } from './api';
import { replaceProposals, textTargets, type Proposal } from './editing';

type Props = { doc: Document; drafts: Record<string, Operation>; onClose: () => void; onApply: (proposals: Proposal[]) => void; onSettings: (doc: Document) => void };

export function AdvancedSettings({ doc, drafts, onClose, onApply, onSettings }: Props) {
  const [tab, setTab] = useState<'rules' | 'replace'>('rules');
  const [settings, setSettings] = useState<Settings>(doc.settings);
  const [terms, setTerms] = useState(doc.settings.protected_terms.join('\n'));
  const [scope, setScope] = useState('all');
  const [section, setSection] = useState(-1);
  const [find, setFind] = useState('');
  const [replacement, setReplacement] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(true);
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [excluded, setExcluded] = useState<string[]>([]);
  const excludedIds = useMemo(() => new Set(excluded), [excluded]);
  const dialogRef = useRef<HTMLElement>(null);
  const settingsValue = { ...settings, protected_terms: [...new Set(terms.split('\n').map(t => t.trim()).filter(Boolean))] };
  const settingsDirty = JSON.stringify(settingsValue) !== JSON.stringify(doc.settings);
  const targets = useMemo(() => textTargets(doc, drafts), [doc, drafts]);
  const ranged = targets.filter(t => (section === -1 || t.operation.section === section) && (scope === 'all' || (scope === 'paragraphs') === (t.operation.type === 'paragraph')));
  const replacements = useMemo(() => replaceProposals(ranged, find, replacement, caseSensitive, settingsValue), [ranged, find, replacement, caseSensitive, settingsValue]);
  const proposals = tab === 'replace' ? replacements : [];
  const selected = proposals.filter(p => !p.blocked && !excludedIds.has(p.id));
  const closeRef = useRef(() => {});
  closeRef.current = () => {
    if (busy) return;
    if (settingsDirty && !window.confirm('저장하지 않은 설정을 버리고 닫을까요?')) return;
    onClose();
  };
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    const overflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    const controls = () => [...(dialogRef.current?.querySelectorAll<HTMLElement>('button:not([disabled]), a[href], input:not([disabled]), textarea:not([disabled]), select:not([disabled])') || [])];
    controls()[0]?.focus();
    function keyboard(e: KeyboardEvent) {
      if (e.key === 'Escape') closeRef.current();
      if (e.key !== 'Tab') return;
      const items = controls();
      if (e.shiftKey && document.activeElement === items[0]) { e.preventDefault(); items.at(-1)?.focus(); }
      else if (!e.shiftKey && document.activeElement === items.at(-1)) { e.preventDefault(); items[0]?.focus(); }
    }
    document.addEventListener('keydown', keyboard);
    return () => { document.body.style.overflow = overflow; document.removeEventListener('keydown', keyboard); previous?.focus(); };
  }, []);
  useEffect(() => { setExcluded([]); }, [scope, section, doc.revision, doc.settings_version]);
  useEffect(() => { setExcluded([]); }, [find, replacement, caseSensitive]);
  useEffect(() => {
    function beforeUnload(event: BeforeUnloadEvent) {
      if (settingsDirty || busy) { event.preventDefault(); event.returnValue = ''; }
    }
    window.addEventListener('beforeunload', beforeUnload);
    return () => window.removeEventListener('beforeunload', beforeUnload);
  }, [settingsDirty, busy]);
  async function saveSettings() {
    setBusy('설정을 저장하고 있습니다'); setError(''); setNotice('');
    try {
      await api(`/documents/${doc.id}/settings`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ version: doc.settings_version, settings: settingsValue }) });
      const fresh = await api<Document>(`/documents/${doc.id}`);
      onSettings(fresh); setSettings(fresh.settings); setTerms(fresh.settings.protected_terms.join('\n')); setNotice('이 문서의 설정을 저장했습니다.');
    } catch (e) { setError(e instanceof Error ? e.message : '설정을 저장하지 못했습니다.'); }
    finally { setBusy(''); }
  }
  function apply() {
    if (!selected.length || selected.length > 500 || settingsDirty) return;
    onApply(selected); onClose();
  }
  const scopeControls = <div className="advanced-scope"><label>대상<select aria-label="수정 대상" value={scope} onChange={e => setScope(e.target.value)} disabled={!!busy}><option value="all">문단과 표</option><option value="paragraphs">문단만</option><option value="tables">표만</option></select></label><label>범위<select aria-label="수정 섹션" value={section} onChange={e => setSection(Number(e.target.value))} disabled={!!busy}><option value={-1}>전체 섹션</option>{doc.sections.map(s => <option key={s.index} value={s.index}>섹션 {s.index + 1}</option>)}</select></label></div>;
  return <div className="advanced-backdrop" onClick={() => closeRef.current()}><section ref={dialogRef} className="advanced-panel" role="dialog" aria-modal="true" aria-labelledby="advanced-title" onClick={e => e.stopPropagation()}>
    <header className="advanced-heading"><div><span className="eyebrow">문서별 작업 도구</span><h2 id="advanced-title">고급 설정</h2><p title={doc.name}>{doc.name}</p></div><button className="icon-button" aria-label="고급 설정 닫기" disabled={!!busy} onClick={() => closeRef.current()}><X size={20} /></button></header>
    <div className="advanced-tabs" role="tablist" aria-label="고급 기능">{([['rules', '편집 규칙'], ['replace', '찾기·바꾸기']] as const).map(([id, label]) => <button key={id} role="tab" aria-selected={tab === id} aria-controls={`advanced-${id}`} id={`advanced-tab-${id}`} disabled={!!busy} onClick={() => { setTab(id); setError(''); setNotice(''); }} className={tab === id ? 'active' : ''}>{label}</button>)}</div>
    <div className="advanced-body" role="tabpanel" id={`advanced-${tab}`} aria-labelledby={`advanced-tab-${tab}`}>
      {error && <div className="advanced-message error" role="alert">{error}</div>}{notice && <div className="advanced-message" role="status"><Check size={15} />{notice}</div>}
      {tab === 'rules' ? <>
        <section className="setting-section"><h3>작업 메모</h3><p>문서의 목적, 문체, 유지할 내용 등을 적어 두세요.<br />작업 중 참고할 수 있도록 이 문서에 함께 저장합니다.</p><label className="sr-only" htmlFor="task-instructions">작업 메모</label><textarea id="task-instructions" rows={6} maxLength={8000} value={settings.instructions} onChange={e => setSettings({ ...settings, instructions: e.target.value })} placeholder={'예: 외부 고객에게 보내는 제안서입니다.\n합니다체로 통일하고, 제품명과 금액은 유지하세요.'} /><span className="field-counter">{settings.instructions.length.toLocaleString()} / 8,000</span></section>
        <section className="setting-section"><h3><LockKeyhole size={16} />보호 문구</h3><p>한 줄에 하나씩 입력하세요. 수동·일괄 수정 시 각 문단과 셀에서 해당 문구의 개수가 유지되어야 합니다. 대소문자를 구분합니다.</p><label className="sr-only" htmlFor="protected-terms">보호 문구</label><textarea id="protected-terms" rows={4} value={terms} maxLength={20100} onChange={e => setTerms(e.target.value)} placeholder={'주식회사 자칼\n1,000,000원'} /><span className="field-counter">최대 100개 · 문구당 200자</span></section>
        <section className="setting-section"><label className="setting-checkbox"><input type="checkbox" checked={settings.strict_validation} onChange={e => setSettings({ ...settings, strict_validation: e.target.checked })} /><span><strong>엄격 검증</strong><small>원본에 남아 있는 오류도 저장·검증 단계에서 차단합니다. 기본 설정에서는 새로 생긴 오류만 차단합니다.</small></span></label></section>
      </> : <>
        <section className="setting-section"><h3>텍스트 일괄 수정</h3><p>찾은 문구를 모두 바꿉니다. 아래에서 변경 전후를 확인하고 편집 내용에 추가하세요.</p>{scopeControls}<label className="setting-label" htmlFor="batch-find">찾을 내용</label><input id="batch-find" value={find} maxLength={10000} onChange={e => setFind(e.target.value)} placeholder="찾을 문구" /><label className="setting-label" htmlFor="batch-replace">바꿀 내용</label><textarea id="batch-replace" rows={2} maxLength={100000} value={replacement} onChange={e => setReplacement(e.target.value)} placeholder="비워 두면 찾은 문구를 삭제합니다" /><label className="inline-checkbox"><input type="checkbox" checked={caseSensitive} onChange={e => setCaseSensitive(e.target.checked)} />대소문자 구분</label></section>
        {find && !proposals.length && <p className="advanced-empty">선택 범위의 편집 가능한 텍스트에 일치하는 문구가 없습니다.</p>}

      </>}
      {tab !== 'rules' && proposals.length > 0 && <section className="proposal-section"><div className="proposal-heading"><h3>변경 미리보기 <span>{proposals.length}</span></h3><button onClick={() => setExcluded(selected.length ? proposals.map(p => p.id) : [])}>{selected.length ? '선택 해제' : '모두 선택'}</button></div>{proposals.slice(0, 500).map(p => <article className={`proposal ${p.blocked ? 'blocked' : ''}`} key={p.id}><label><input type="checkbox" checked={!p.blocked && !excludedIds.has(p.id)} disabled={!!p.blocked} onChange={e => setExcluded(e.target.checked ? excluded.filter(id => id !== p.id) : [...excluded, p.id])} /><strong>{p.label}</strong></label><div className="proposal-before"><span>변경 전</span><p>{p.before || '(빈 내용)'}</p></div><div className="proposal-after"><span>변경 후</span><p>{p.after || '(빈 내용)'}</p></div><p className="proposal-reason">{p.blocked || p.reason}</p></article>)}</section>}
      {settingsDirty && tab !== 'rules' && <p className="inline-warning">편집 규칙 탭의 설정을 먼저 저장해 주세요.</p>}
      {selected.length > 500 && <p className="inline-warning">한 번에 최대 500개 요소를 적용할 수 있습니다. 미리보기는 처음 500개만 표시합니다. 선택 범위를 좁혀 주세요.</p>}
    </div>
    <footer className="advanced-footer"><span>{tab === 'rules' ? settingsDirty ? '저장하지 않은 설정' : '이 문서에 저장됨' : `${selected.length}개 선택 · 적용 후 수정 저장 필요`}</span>{tab === 'rules' ? <button className="primary-button" disabled={!!busy || !settingsDirty} onClick={() => void saveSettings()}><Save size={15} />설정 저장</button> : <button className="primary-button" disabled={!!busy || !selected.length || selected.length > 500 || settingsDirty} onClick={apply}>편집에 추가<ArrowRight size={15} /></button>}</footer>
  </section></div>;
}
