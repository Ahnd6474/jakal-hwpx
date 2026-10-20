import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
export async function integrationHash() {
  const hash = createHash('sha256');
  for (const name of ['build-studio.mjs', '../studio/host-bridge.ts', '../studio/jakal.css']) hash.update(await readFile(new URL(name, import.meta.url)));
  return hash.digest('hex');
}
