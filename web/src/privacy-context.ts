// Public namespace/fence only. Authentication stays in the HttpOnly cookie.
let workspace = '';
let locked = false;
const pendingRequests = new Set<AbortController>();
function abortPrivateRequests() {
  const pending = [...pendingRequests]; pendingRequests.clear();
  for (const controller of pending) controller.abort(new DOMException('사용 세션이 변경되었습니다.', 'AbortError'));
}
export function trackPrivateRequest(controller: AbortController) {
  if (locked) controller.abort(new DOMException('사용 세션이 종료되었습니다.', 'AbortError'));
  else pendingRequests.add(controller);
  return () => pendingRequests.delete(controller);
}
export const workspaceContext = () => workspace;
export function setWorkspaceContext(value: string) {
  if (!/^[a-f0-9]{32}$/.test(value)) throw new Error('사용 세션 정보를 확인하지 못했습니다.');
  if (workspace !== value) abortPrivateRequests();
  workspace = value; locked = false;
}
export const privateStorageKey = (base: string) => `${base}:${workspace || 'locked'}`;
export const privacyLocked = () => locked;
export function lockWorkspace() {
  if (locked) return;
  locked = true;
  abortPrivateRequests();
  if (typeof window !== 'undefined') window.dispatchEvent(new window.Event('jakal-session-ended'));
}

export async function clearPrivateStorage() {
  for (const storage of [localStorage, sessionStorage]) {
    for (let index = storage.length - 1; index >= 0; index--) {
      const key = storage.key(index);
      if (key?.startsWith('jakal-')) storage.removeItem(key);
    }
  }
  const bases = ['jakal-document-editor', 'jakal-page-template-assets'];
  const names = new Set(bases.flatMap(base => [base, privateStorageKey(base)]));
  // Remove older scoped/legacy copies on an explicitly chosen shared-device exit.
  if (indexedDB.databases) for (const db of await indexedDB.databases()) {
    const name = db.name;
    if (name && bases.some(base => name === base || name.startsWith(base + ':'))) names.add(name);
  }
  await Promise.all([...names].map(name => new Promise<void>((resolve, reject) => {
    const request = indexedDB.deleteDatabase(name);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(new Error('브라우저 임시본을 지우지 못했습니다. 사이트 데이터를 직접 삭제해 주세요.'));
    request.onblocked = () => reject(new Error('다른 탭을 닫고 브라우저 사이트 데이터를 삭제해 주세요.'));
  })));
}
