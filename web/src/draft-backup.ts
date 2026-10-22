import { privateStorageKey, privacyLocked } from './privacy-context.ts';
import type { Document, Operation } from './api';
import type { Drafts } from './draft-history';
import { textTargetIndex } from './editing.ts';

const PREFIX = 'jakal-draft:';
const MAX_BACKUP = 2_000_000;
export type DraftBackup = { version: 1; documentId: string; revision: number; expires: number; updatedAt: number; changes: { id: string; original: string; operation: Operation }[] };
export type BackupResult = { kind: 'none' | 'unavailable' } | { kind: 'restore'; backup: DraftBackup; drafts: Drafts } | { kind: 'conflict'; backup: DraftBackup };

function sameTarget(a: Operation, b: Operation): boolean {
  return a.type === b.type && a.section === b.section && (a.type === 'paragraph' && b.type === 'paragraph'
    ? a.index === b.index : a.type === 'cell' && b.type === 'cell' && a.table === b.table && a.row === b.row && a.column === b.column);
}

export function decodeBackup(raw: string | null, doc: Document, now = Date.now()): BackupResult {
  if (!raw || raw.length > MAX_BACKUP) return { kind: 'none' };
  try {
    const value = JSON.parse(raw) as DraftBackup;
    if (!value || value.version !== 1 || value.documentId !== doc.id || !Number.isInteger(value.revision) || value.revision < 0
      || !Number.isFinite(value.expires) || value.expires * 1000 <= now || !Number.isFinite(value.updatedAt)
      || !Array.isArray(value.changes) || !value.changes.length || value.changes.length > 20000) return { kind: 'none' };
    const known = textTargetIndex(doc);
    const drafts: Drafts = {};
    const seen = new Set<string>();
    let compatible = value.revision === doc.revision;
    for (const change of value.changes) {
      if (!change || typeof change.id !== 'string' || seen.has(change.id) || typeof change.original !== 'string' || change.original.length > 100000
        || !change.operation || !['paragraph', 'cell'].includes(change.operation.type) || typeof change.operation.text !== 'string' || change.operation.text.length > 100000) return { kind: 'none' };
      seen.add(change.id);
      const target = known.get(change.id);
      if (!target || !target.editable || target.original !== change.original || !sameTarget(target.operation, change.operation)) compatible = false;
      if (target && target.original !== change.operation.text) drafts[target.id] = { ...target.operation, text: change.operation.text };
    }
    return compatible ? { kind: 'restore', backup: value, drafts } : { kind: 'conflict', backup: value };
  } catch { return { kind: 'none' }; }
}

export function readDraftBackup(doc: Document): BackupResult {
  try { return decodeBackup(sessionStorage.getItem(privateStorageKey(PREFIX) + doc.id), doc); }
  catch { return { kind: 'unavailable' }; }
}

export function clearDraftBackup(id: string): boolean {
  try { sessionStorage.removeItem(privateStorageKey(PREFIX) + id); return true; }
  catch { return false; }
}

export function writeDraftBackup(doc: Document, drafts: Drafts): boolean {
  if (privacyLocked()) return false;
  const entries = Object.entries(drafts);
  if (!entries.length) return clearDraftBackup(doc.id);
  const index = textTargetIndex(doc);
  const changes = entries.flatMap(([id, operation]) => {
    const target = index.get(id);
    const text = operation?.text ?? target?.original;
    return target?.editable && text !== undefined && text !== target.original ? [{ target, text }] : [];
  }).sort((a, b) => a.target.order - b.target.order)
    .map(({ target, text }) => ({ id: target.id, original: target.original, operation: { ...target.operation, text } }));
  if (!changes.length) return clearDraftBackup(doc.id);
  const backup: DraftBackup = { version: 1, documentId: doc.id, revision: doc.revision, expires: doc.expires, updatedAt: Date.now(), changes };
  try {
    const raw = JSON.stringify(backup);
    if (raw.length > MAX_BACKUP) return false;
    sessionStorage.setItem(privateStorageKey(PREFIX) + doc.id, raw);
    return true;
  } catch { return false; }
}

export function backupText(backup: DraftBackup): string {
  return [`미저장 편집 · 기준 버전 ${backup.revision + 1}`, new Date(backup.updatedAt).toLocaleString('ko-KR'), '',
    ...backup.changes.flatMap(change => [change.id, '변경 전', change.original, '변경 후', change.operation.text, ''])].join('\n');
}
