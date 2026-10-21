type Definition = { id: string; label: string; canExecute?: (context: any) => boolean; execute: (...args: any[]) => unknown; [key: string]: unknown };
type Registry = { get: (id: string) => Definition | undefined; register: (definition: Definition) => void };
export type HostFileAction = 'new' | 'open' | 'documents' | 'save' | 'copy' | 'download' | 'history';

export function installHostFileActions(registry: Registry, post: (action: HostFileAction) => void) {
  const actions: [string, HostFileAction][] = [
    ['file:new-doc', 'new'], ['file:open', 'open'], ['file:open-recent', 'documents'],
    ['file:save', 'save'], ['file:save-as', 'copy'], ['file:save-as-hwpx', 'download'], ['edit:document-history', 'history'],
  ];
  for (const [id, action] of actions) {
    const definition = registry.get(id);
    if (definition) registry.register({ ...definition, execute: () => post(action) });
  }
  // This workspace validates and downloads HWPX; a local HWP save must not
  // silently acknowledge or clear the host's unsaved revision.
  const hwpSave = registry.get('file:save-as-hwp');
  if (hwpSave) registry.register({ ...hwpSave, canExecute: () => false });
}
