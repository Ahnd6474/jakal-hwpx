import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
export async function integrationHash() {
  const hash = createHash('sha256');
  hash.update(await readFile(new URL('../studio/file-command-policy.mjs', import.meta.url)));
  for (const name of ['build-studio.mjs', '../studio/host-bridge.ts', '../studio/jakal.css', '../studio/input-profile.ts', '../studio/font-cache.ts', '../studio/frame-task.ts', '../studio/block-insertion.ts', '../studio/character-dialog.ts', '../studio/file-actions.ts']) hash.update(await readFile(new URL(name, import.meta.url)));
  return hash.digest('hex');
}
