import type { Operation } from './api';

export type Drafts = Record<string, Operation>;
export type History = { past: Drafts[]; present: Drafts; future: Drafts[] };
type Change = { key: string; operation: Operation; original: string };
export type HistoryAction = { type: 'change' } & Change
  | { type: 'batch'; changes: Change[] }
  | { type: 'undo' | 'redo' | 'reset' };

const HISTORY_LIMIT = 60;

function sameOperation(a: Operation, b: Operation) {
  if (a.type !== b.type || a.section !== b.section || a.text !== b.text) return false;
  if (a.type === 'paragraph' && b.type === 'paragraph') return a.index === b.index;
  return a.type === 'cell' && b.type === 'cell' && a.table === b.table && a.row === b.row && a.column === b.column;
}

export function draftReducer(state: History, action: HistoryAction): History {
  if (action.type === 'reset') return { past: [], present: {}, future: [] };
  if (action.type === 'undo') return state.past.length
    ? { past: state.past.slice(0, -1), present: state.past.at(-1)!, future: [state.present, ...state.future] } : state;
  if (action.type === 'redo') return state.future.length
    ? { past: [...state.past, state.present], present: state.future[0], future: state.future.slice(1) } : state;
  if (action.type !== 'change' && action.type !== 'batch') return state;

  const present = { ...state.present };
  for (const change of action.type === 'batch' ? action.changes : [action]) {
    if (change.operation.text === change.original) delete present[change.key];
    else present[change.key] = change.operation;
  }
  const keys = Object.keys(present);
  // An unchanged input (or a batch that cancels itself out) must preserve redo.
  if (keys.length === Object.keys(state.present).length
    && keys.every(key => state.present[key] && sameOperation(present[key], state.present[key]))) return state;
  return { past: [...state.past.slice(-(HISTORY_LIMIT - 1)), state.present], present, future: [] };
}
