import { hostFileActions } from './file-command-policy.mjs';

type Definition = { id: string; label: string; canExecute?: (context: any) => boolean; execute: (...args: any[]) => unknown; [key: string]: unknown };
type Registry = { get: (id: string) => Definition | undefined; register: (definition: Definition) => void };
export type HostFileAction = 'new' | 'open' | 'documents' | 'save' | 'copy' | 'download' | 'history';

export function installHostFileActions(registry: Registry, post: (action: HostFileAction) => void) {
  for (const [id, action] of Object.entries(hostFileActions)) {
    const definition = registry.get(id);
    if (definition) registry.register({ ...definition, execute: () => post(action as HostFileAction) });
  }
  // This workspace validates and downloads HWPX; a local HWP save must not
  // silently acknowledge or clear the host's unsaved revision.
  const hwpSave = registry.get('file:save-as-hwp');
  if (hwpSave) registry.register({ ...hwpSave, canExecute: () => false });
}
