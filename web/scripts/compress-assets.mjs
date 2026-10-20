import { readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { gzipSync } from 'node:zlib';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export async function compressAssets(root) {
  let original = 0, compressed = 0;
  for (const entry of await readdir(root, { withFileTypes: true })) {
    const path = resolve(root, entry.name);
    if (entry.isDirectory()) {
      const size = await compressAssets(path); original += size.original; compressed += size.compressed;
    } else if (/\.(js|css|wasm|svg)$/.test(entry.name)) {
      const info = await stat(path);
      if (info.size < 1024) continue;
      const raw = await readFile(path);
      const bytes = gzipSync(raw, { level: 6 });
      if (bytes.length >= raw.length) continue;
      await writeFile(path + '.gz', bytes);
      original += raw.length; compressed += bytes.length;
    }
  }
  return { original, compressed };
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const size = await compressAssets(resolve(process.argv[2] || 'dist'));
  console.log(`Static compression: ${size.original} → ${size.compressed} bytes (${Math.round((1 - size.compressed / size.original) * 100)}% smaller)`);
}
