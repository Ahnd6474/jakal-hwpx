import { readFile } from 'node:fs/promises';
import { integrationHash } from './studio-version.mjs';
try {
  const version = JSON.parse(await readFile(new URL('../public/studio/metadata/version.json', import.meta.url), 'utf8'));
  if (version.integrationHash !== await integrationHash()) throw new Error('Editor integration changed');
}
catch { await import('./build-studio.mjs'); }
