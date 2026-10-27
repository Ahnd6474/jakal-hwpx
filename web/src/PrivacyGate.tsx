import { useEffect, useState, type ReactNode } from 'react';
import { api } from './api';
import { clearPrivateStorage, lockWorkspace, privacyLocked, setWorkspaceContext, workspaceContext } from './privacy-context';

let initialization: Promise<void> | undefined;
export function PrivacyGate({ children }: { children: ReactNode }) {
  const [state, setState] = useState<'loading' | 'ready' | 'ended' | 'error'>('loading');
  const [message, setMessage] = useState('');
  useEffect(() => {
    let active = true;
    initialization ||= api<{ workspace: string }>('/session').then(value => setWorkspaceContext(value.workspace));
    void initialization.then(() => { if (active) setState(privacyLocked() ? 'ended' : 'ready'); })
      .catch(() => { initialization = undefined; if (active) setState('error'); });
    const ended = () => { if (active) { setState('ended'); document.title = '사용 세션 종료 · Jakal'; } };
    window.addEventListener('jakal-session-ended', ended);
    let channel: BroadcastChannel | undefined;
    try { channel = new BroadcastChannel('jakal-private-session'); channel.onmessage = event => { if (event.data?.ended === workspaceContext()) lockWorkspace(); }; } catch { /* No cross-tab API. */ }
    const check = () => { if (!document.hidden && workspaceContext() && !privacyLocked()) void api('/session').catch(() => {}); };
    window.addEventListener('focus', check);
    document.addEventListener('visibilitychange', check);
    const timer = window.setInterval(check, 30_000);
    return () => { active = false; window.removeEventListener('jakal-session-ended', ended); window.removeEventListener('focus', check); document.removeEventListener('visibilitychange', check); clearInterval(timer); channel?.close(); };
  }, []);
  if (state === 'ready') return children;
  return <main className="privacy-gate" aria-live="polite"><h1>{state === 'loading' ? '비공개 문서 공간을 준비하고 있습니다' : state === 'ended' ? '사용 세션이 종료되었습니다' : '비공개 문서 공간을 열지 못했습니다'}</h1>
    <p>{state === 'ended' ? '문서 화면을 닫았습니다. 공용 기기라면 브라우저 보관 데이터도 삭제하세요. 새로 시작하면 별도의 문서 공간을 사용합니다.' : '세션 확인을 마쳐야 문서와 임시본을 표시합니다.'}</p>
    {message && <p role="status">{message}</p>}
    {state === 'ended' && <button className="outline-button" onClick={() => void clearPrivateStorage().then(() => setMessage('브라우저 임시본과 내 양식을 삭제했습니다.')).catch(e => setMessage(e.message))}>브라우저 보관 데이터 삭제</button>}
    {state !== 'loading' && <button className="primary-button" onClick={() => window.location.reload()}>새로 시작</button>}
  </main>;
}
