import { privateStorageKey, privacyLocked } from './privacy-context.ts';
export type StudioBackup = { id: string; revision: number; bytes: Uint8Array; saved: number; name?: string };
export type StudioBackupInfo = { id: string; revision: number; saved: number; size: number; name?: string };
const database = 'jakal-document-editor';
async function access<T>(mode: IDBTransactionMode, stores: string[], action: (transaction: IDBTransaction) => () => T): Promise<T> {
  if (privacyLocked()) return Promise.reject(new Error('사용 세션이 종료되었습니다.'));
  return new Promise((resolve, reject) => {
    const open = indexedDB.open(privateStorageKey(database), 2);
    let abandoned = false;
    open.onupgradeneeded = () => {
      if (!open.result.objectStoreNames.contains('drafts')) open.result.createObjectStore('drafts', { keyPath: 'id' });
      if (!open.result.objectStoreNames.contains('catalog')) open.result.createObjectStore('catalog', { keyPath: 'id' });
    };
    open.onblocked = () => { abandoned = true; reject(new Error('다른 탭이 임시보관함을 사용 중입니다. 다른 탭을 닫고 다시 시도해 주세요.')); };
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      if (abandoned) { db.close(); return; }
      db.onversionchange = () => db.close();
      let transaction: IDBTransaction | undefined;
      try {
        transaction = db.transaction(stores, mode);
        const result = action(transaction);
        transaction.oncomplete = () => { db.close(); try { resolve(result()); } catch (error) { reject(error); } };
        transaction.onabort = transaction.onerror = () => { db.close(); reject(transaction?.error || new Error('임시보관함 작업을 완료하지 못했습니다.')); };
      } catch (error) { transaction?.abort(); db.close(); reject(error); }
    };
  });
}
export const studioBackup = {
  read: (id: string) => access<StudioBackup | undefined>('readonly', ['drafts'], tx => { const request = tx.objectStore('drafts').get(id); return () => request.result; }),
  write: (backup: StudioBackup) => access('readwrite', ['drafts', 'catalog'], tx => {
    tx.objectStore('drafts').put(backup);
    tx.objectStore('catalog').put({ id: backup.id, revision: backup.revision, saved: backup.saved, name: backup.name, size: backup.bytes.byteLength });
    return () => undefined;
  }),
  clear: (id: string) => access('readwrite', ['drafts', 'catalog'], tx => {
    tx.objectStore('drafts').delete(id); tx.objectStore('catalog').delete(id); return () => undefined;
  }),
  // List small metadata records and keys, without cloning every document's bytes.
  // Older drafts remain visible even when they predate the catalog store.
  list: () => access<StudioBackupInfo[]>('readonly', ['drafts', 'catalog'], tx => {
    const keys = tx.objectStore('drafts').getAllKeys();
    const catalog = tx.objectStore('catalog').getAll();
    return () => {
      const details = new Map<string, StudioBackupInfo>(catalog.result.map((item: StudioBackupInfo) => [item.id, item]));
      return keys.result.filter((id): id is string => typeof id === 'string').map(id => details.get(id) || { id, revision: -1, saved: 0, size: 0 }).sort((a, b) => b.saved - a.saved);
    };
  }),
};
