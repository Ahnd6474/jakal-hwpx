import { privateStorageKey, privacyLocked } from './privacy-context.ts';
import { backgroundProblem, layoutProblem, type BackgroundPage, type PageLayout } from './page-layout.ts';

export type AssetPresetInfo = { id: string; name: string; layout: PageLayout; page_count: number; size: number; created: number; asset: true };
export type AssetPreset = AssetPresetInfo & { background_pages: BackgroundPage[] };
const database = 'jakal-page-template-assets';

function access<T>(mode: IDBTransactionMode, action: (tx: IDBTransaction, fail: (message: string) => void) => () => T): Promise<T> {
  if (privacyLocked()) return Promise.reject(new Error('사용 세션이 종료되었습니다.'));
  return new Promise((resolve, reject) => {
    const open = indexedDB.open(privateStorageKey(database), 1);
    let abandoned = false;
    open.onupgradeneeded = () => { for (const name of ['files', 'catalog']) if (!open.result.objectStoreNames.contains(name)) open.result.createObjectStore(name, { keyPath: 'id' }); };
    open.onerror = () => reject(new Error('배경 양식 보관함을 열지 못했습니다.'));
    open.onblocked = () => { abandoned = true; reject(new Error('다른 탭에서 배경 양식 보관함을 사용 중입니다. 잠시 후 다시 확인해 주세요.')); };
    open.onsuccess = () => {
      const db = open.result;
      if (abandoned) { db.close(); return; }
      db.onversionchange = () => db.close();
      let transaction: IDBTransaction | undefined;
      let reason = '';
      try {
        transaction = db.transaction(['files', 'catalog'], mode);
        const value = action(transaction, message => { reason = message; transaction!.abort(); });
        transaction.oncomplete = () => { db.close(); try { resolve(value()); } catch (error) { reject(error); } };
        transaction.onabort = transaction.onerror = () => { db.close(); reject(new Error(reason || '배경 양식 저장 공간이 부족하거나 보관 작업을 완료하지 못했습니다.')); };
      } catch (error) { transaction?.abort(); db.close(); reject(error); }
    };
  });
}

export const pageTemplateAssets = {
  list: () => access<AssetPresetInfo[]>('readonly', tx => {
    const request = tx.objectStore('catalog').getAll(); return () => request.result;
  }),
  read: (id: string) => access<AssetPreset | undefined>('readonly', tx => {
    const request = tx.objectStore('files').get(id); return () => request.result;
  }),
  async save(name: string, layout: PageLayout, pages: BackgroundPage[]) {
    const problem = layoutProblem(layout) || backgroundProblem(pages);
    if (problem) throw new Error(problem);
    if (!name.trim() || name.trim().length > 60 || !pages.length) throw new Error('양식 이름과 배경 페이지를 확인해 주세요.');
    const data = structuredClone(pages), definition = structuredClone(layout);
    const info: AssetPresetInfo = { id: crypto.randomUUID(), name: name.trim(), layout: definition,
      page_count: data.length, size: JSON.stringify(data).length, created: Date.now(), asset: true };
    return access<AssetPresetInfo>('readwrite', (tx, fail) => {
      const catalog = tx.objectStore('catalog'), request = catalog.getAll();
      request.onsuccess = () => {
        const current = request.result as AssetPresetInfo[];
        if (current.length >= 20 || current.reduce((total, item) => total + item.size, 0) + info.size > 100 * 1024 * 1024) return fail('배경 양식은 최대 20개, 합계 100MB까지 보관할 수 있습니다. 사용하지 않는 양식을 정리해 주세요.');
        tx.objectStore('files').add({ ...info, background_pages: data }); catalog.add(info);
      };
      return () => info;
    });
  },
  remove: (id: string) => access('readwrite', tx => {
    tx.objectStore('files').delete(id); tx.objectStore('catalog').delete(id); return () => undefined;
  }),
};
