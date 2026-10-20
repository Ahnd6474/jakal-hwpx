// Pinned, self-hosted editor runtime. No document is sent to the upstream site.
import { build } from 'vite';
import { execFileSync } from 'node:child_process';
import { cp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { integrationHash } from './studio-version.mjs';
import { compressAssets } from './compress-assets.mjs';

const web = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repo = resolve(web, '..');
const source = resolve(repo, 'build/rhwp-studio-source');
const commit = '1a76570e833917d15817415a53c09ad61ab3203f'; // v0.8.7, MIT
try {
  await readFile(resolve(source, 'rhwp-studio/src/main.ts'));
} catch {
  execFileSync('git', ['clone', '--depth', '1', '--branch', 'v0.8.7', '--filter=blob:none', '--sparse', 'https://github.com/edwardkim/rhwp.git', source], { stdio: 'inherit' });
  execFileSync('git', ['sparse-checkout', 'set', 'rhwp-studio', 'rhwp-shared', 'npm/editor', 'assets/fonts'], { cwd: source, stdio: 'inherit' });
}
if (execFileSync('git', ['rev-parse', 'HEAD'], { cwd: source, encoding: 'utf8' }).trim() !== commit) throw new Error('Unexpected editor source revision');
const out = resolve(web, 'public/studio');
const bridge = await readFile(resolve(web, 'studio/host-bridge.ts'), 'utf8');
const css = await readFile(resolve(web, 'studio/jakal.css'), 'utf8');
await build({
  configFile: false, root: resolve(source, 'rhwp-studio'), base: '/studio/', publicDir: false,
  define: { __APP_VERSION__: JSON.stringify('0.8.7'), __RHWP_DISABLE_EXTERNAL_WEBFONTS__: 'true', __RHWP_HWPCTRL__: 'false' },
  resolve: { alias: {
    '@': resolve(source, 'rhwp-studio/src'), '@wasm': resolve(web, 'node_modules/@rhwp/core'),
    '@noble/hashes': resolve(web, 'node_modules/@noble/hashes'),
    'canvaskit-wasm': resolve(web, 'node_modules/canvaskit-wasm'),
  } },
  plugins: [{
    name: 'jakal-studio-host', enforce: 'pre',
    transform(code, id) {
      if (id.replaceAll('\\', '/').endsWith('/rhwp-studio/src/main.ts')) return code + '\n' + bridge;
      if (id.replaceAll('\\', '/').endsWith('/core/wasm-bridge.ts')) {
        // Byte-only SDK exports must also reject reported content loss.
        return code.replace('return this.doc.exportHwpx();', `const artifact = this.exportHwpxWithReport();
          if (artifact.contentLoss.count > 0) throw new Error('저장 과정에서 지원하지 않는 개체가 발견되었습니다. 원본을 유지하고 저장을 중단했습니다.');
          return artifact.bytes;`);
      }
    },
    transformIndexHtml(html) {
      return html.replace('<title>rhwp-studio</title>', '<title>Jakal 문서 편집기</title>').replace('src="/theme-init.js"', 'src="/studio/theme-init.js"').replace('src="/locale-init.js"', 'src="/studio/locale-init.js"').replace('</head>', '<link rel="stylesheet" href="/studio/jakal.css"></head>');
    },
  }],
  build: { outDir: out, emptyOutDir: true, chunkSizeWarningLimit: 2500 },
});
for (const name of ['theme-init.js', 'locale-init.js', 'favicon.ico', 'icons', 'images', 'print.html']) {
  await cp(resolve(source, 'rhwp-studio/public', name), resolve(out, name), { recursive: true });
}
await cp(resolve(source, 'assets/fonts'), resolve(out, 'fonts'), { recursive: true });
await cp(resolve(source, 'LICENSE'), resolve(out, 'LICENSE-rhwp.txt'));
await writeFile(resolve(out, 'jakal.css'), css);
await compressAssets(out);
await mkdir(resolve(out, 'metadata'), { recursive: true });
await writeFile(resolve(out, 'metadata/version.json'), JSON.stringify({ version: '0.8.7', commit, integrationHash: await integrationHash(), source: 'https://github.com/edwardkim/rhwp', license: 'MIT' }));
