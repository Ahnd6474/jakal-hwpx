// Only these commands may regain a surface in the embedded runtime.
export const hostFileActions = {
  'file:new-doc': 'new', 'file:open': 'open', 'file:open-recent': 'documents',
  'file:save': 'save', 'file:save-as': 'copy', 'file:save-as-hwpx': 'download',
  'edit:document-history': 'history',
};
export const nativeOutputCommands = ['file:print', 'file:print-to-pdf'];
export const restoredFileCommands = [...Object.keys(hostFileActions).filter(id => id.startsWith('file:')), ...nativeOutputCommands];

export function restoreHostFileCommands(code) {
  const start = code.indexOf('export const EMBED_HIDDEN_FILE_COMMAND_IDS: readonly string[] = [');
  const end = code.indexOf('\n];', start);
  if (start < 0 || end < 0) throw new Error('Pinned embed file command policy changed');
  let block = code.slice(start, end + 1);
  for (const id of restoredFileCommands) {
    const line = `  '${id}',\n`;
    if (!block.includes(line)) throw new Error(`Pinned embed command changed: ${id}`);
    block = block.replace(line, '');
  }
  return code.slice(0, start) + block + code.slice(end + 1);
}
