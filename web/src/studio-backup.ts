export type StudioBackup = { id: string; revision: number; bytes: Uint8Array; saved: number };
const database = 'jakal-document-editor';
async function access<T>(mode: IDBTransactionMode, action: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    const open = indexedDB.open(database, 1);
    open.onupgradeneeded = () => open.result.createObjectStore('drafts', { keyPath: 'id' });
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const transaction = db.transaction('drafts', mode);
      const request = action(transaction.objectStore('drafts'));
      transaction.oncomplete = () => { resolve(request.result); db.close(); };
      transaction.onabort = transaction.onerror = () => { reject(transaction.error); db.close(); };
    };
  });
}
export const studioBackup = {
  read: (id: string) => access<StudioBackup | undefined>('readonly', store => store.get(id)),
  write: (backup: StudioBackup) => access('readwrite', store => store.put(backup)),
  clear: (id: string) => access('readwrite', store => store.delete(id)),
};
