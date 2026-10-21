import { useEffect, useState } from 'react';

export function ConnectionStatus() {
  const [offline, setOffline] = useState(() => !window.navigator.onLine);
  useEffect(() => {
    const changed = () => setOffline(!window.navigator.onLine);
    window.addEventListener('online', changed);
    window.addEventListener('offline', changed);
    return () => { window.removeEventListener('online', changed); window.removeEventListener('offline', changed); };
  }, []);
  return offline ? <p className="connection-status" role="status">네트워크 연결이 끊겼습니다. 편집은 계속할 수 있습니다. 저장·다운로드는 서버 연결 상태를 확인한 뒤 진행해 주세요.</p> : null;
}
