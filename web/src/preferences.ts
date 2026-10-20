const LAST_DOCUMENT = 'jakal-document';

// Remembering the last document is optional; blocked or full browser storage
// must never prevent opening, editing, or deleting the document itself.
export const lastDocument = {
  read(): string | null {
    try { const id = sessionStorage.getItem(LAST_DOCUMENT); if (id) return id; }
    catch { /* Fall back to the last document from an earlier session. */ }
    try { return localStorage.getItem(LAST_DOCUMENT); }
    catch { return null; }
  },
  write(id: string | null): void {
    try {
      if (id === null) sessionStorage.removeItem(LAST_DOCUMENT);
      else sessionStorage.setItem(LAST_DOCUMENT, id);
    } catch { /* Browser storage is optional. */ }
    try {
      if (id === null) localStorage.removeItem(LAST_DOCUMENT);
      else localStorage.setItem(LAST_DOCUMENT, id);
    } catch { /* The current session remains usable without this preference. */ }
  },
};
