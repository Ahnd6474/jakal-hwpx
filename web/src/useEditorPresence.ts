import { useEffect, useRef } from 'react';
import { api } from './api';

/** A small heartbeat; document bytes never travel through this path. */
export function useEditorPresence(id: string | undefined, revision: number | undefined, dirty: boolean) {
  const tab = useRef(crypto.randomUUID().replaceAll('-', ''));
  const latest = useRef({ revision, dirty });
  latest.current = { revision, dirty };
  const notify = useRef<() => void>(() => {});
  useEffect(() => {
    if (!id) return;
    let stopped = false, pending = false;
    const publish = async () => {
      if (stopped || pending || !navigator.onLine) return;
      pending = true;
      const sent = latest.current;
      try {
        await api(`/documents/${id}/presence`, { method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tab: tab.current, revision: sent.revision ?? 0, dirty: sent.dirty }), timeoutMs: 5000 });
      } catch { /* Existing save/conflict checks remain authoritative if offline. */ }
      finally {
        pending = false;
        if (!stopped && (latest.current.dirty !== sent.dirty || latest.current.revision !== sent.revision)) void publish();
      }
    };
    notify.current = () => void publish();
    void publish();
    const timer = setInterval(() => void publish(), 10_000);
    window.addEventListener('focus', publish); window.addEventListener('online', publish);
    return () => {
      stopped = true; clearInterval(timer); window.removeEventListener('focus', publish); window.removeEventListener('online', publish);
      notify.current = () => {};
    };
  }, [id]);
  useEffect(() => { notify.current(); }, [dirty, revision]);
}
