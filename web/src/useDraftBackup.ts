import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { Document } from './api';
import type { Drafts } from './draft-history';
import { writeDraftBackup } from './draft-backup';

export function useDraftBackup(doc: Document | null, drafts: Drafts, paused: boolean) {
  const [status, setStatus] = useState<'idle' | 'pending' | 'stored' | 'unavailable' | 'paused'>('idle');
  const latest = useRef({ doc, drafts, paused });
  useLayoutEffect(() => { latest.current = { doc, drafts, paused }; }, [doc, drafts, paused]);
  useEffect(() => {
    if (!doc) { setStatus('idle'); return; }
    if (paused) { setStatus('paused'); return; }
    if (!Object.keys(drafts).length) {
      setStatus(writeDraftBackup(doc, drafts) ? 'idle' : 'unavailable');
      return;
    }
    setStatus('pending');
    const timer = setTimeout(() => setStatus(writeDraftBackup(doc, drafts) ? 'stored' : 'unavailable'), 400);
    return () => clearTimeout(timer);
  }, [doc, drafts, paused]);
  useEffect(() => {
    function flush() {
      const value = latest.current;
      if (value.doc && !value.paused) writeDraftBackup(value.doc, value.drafts);
    }
    window.addEventListener('pagehide', flush);
    window.addEventListener('beforeunload', flush);
    return () => { window.removeEventListener('pagehide', flush); window.removeEventListener('beforeunload', flush); };
  }, []);
  return status;
}
