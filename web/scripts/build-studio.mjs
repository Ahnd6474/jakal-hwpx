// Pinned, self-hosted editor runtime. No document is sent to the upstream site.
import { build } from 'vite';
import { execFileSync } from 'node:child_process';
import { cp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { integrationHash } from './studio-version.mjs';
import { compressAssets } from './compress-assets.mjs';
import { restoreHostFileCommands } from '../studio/file-command-policy.mjs';

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
    '@jakal/input-profile': resolve(web, 'studio/input-profile.ts'),
    '@jakal/font-cache': resolve(web, 'studio/font-cache.ts'),
    '@jakal/frame-task': resolve(web, 'studio/frame-task.ts'),
    '@jakal/block-insertion': resolve(web, 'studio/block-insertion.ts'),
    '@jakal/character-dialog': resolve(web, 'studio/character-dialog.ts'),
    '@jakal/file-actions': resolve(web, 'studio/file-actions.ts'),
  } },
  plugins: [{
    name: 'jakal-studio-host', enforce: 'pre',
    transform(code, id) {
      if (id.replaceAll('\\', '/').endsWith('/ui/chrome-mode.ts')) return restoreHostFileCommands(code.replaceAll('\r\n', '\n'));
      if (id.replaceAll('\\', '/').endsWith('/command/commands/format.ts')) {
        code = code.replaceAll('\r\n', '\n');
        const selection = 'const savedSel = ih.getSelection();\n      if (!savedSel) return;';
        const apply = 'ih.applyCharPropsToRange(savedSel.start, savedSel.end, mods);';
        if (!code.includes(selection) || !code.includes(apply)) throw new Error('Pinned character dialog selection handling changed');
        return `import { applyDialogCharacterProps } from '@jakal/character-dialog';\nimport { CursorState } from '@/engine/cursor';\n` + code.replace(selection, 'const savedSel = ih.getSelection();').replace(apply, 'applyDialogCharacterProps(ih, services.eventBus, savedSel, mods, CursorState.comparePositions);');
      }
      if (id.replaceAll('\\', '/').endsWith('/rhwp-studio/src/main.ts')) {
        code = code.replaceAll('\r\n', '\n');
        const schedule = 'function autosaveScheduleFromUserSettings(): AutosaveScheduleSettings {';
        if (!code.includes(schedule)) throw new Error('Pinned autosave settings changed');
        code = code.replace(schedule, `${schedule}
          // The enclosing app owns version-aware HWPX recovery. Avoid a second
          // synchronous HWP export while the user is entering text.
          if (window.parent !== window && new URLSearchParams(location.search).get('chrome') === 'embed') {
            return { recoveryEnabled: false, recoveryIntervalMs: 600000, idleEnabled: false, idleDelayMs: 10000 };
          }`);
        const fontRefresh = "document.fonts.addEventListener('loadingdone', () => {\n    if (!canvasView || wasm.pageCount === 0) return;";
        if (!code.includes(fontRefresh)) throw new Error('Pinned font refresh listener changed');
        code = code.replace(fontRefresh, "document.fonts.addEventListener('loadingdone', () => {\n    if (document.documentElement.classList.contains('rhwp-busy') || !canvasView || wasm.pageCount === 0) return;");
        const loadStart = 'await withBusyCursor(document.documentElement, async () => {\n    //';
        if (!code.includes(loadStart)) throw new Error('Pinned document loading boundary changed');
        code = code.replace(loadStart, 'await withBusyCursor(document.documentElement, async () => {\n    inputHandler?.deactivate();\n    //');
        return code + '\n' + bridge;
      }
      if (id.replaceAll('\\', '/').endsWith('/core/wasm-bridge.ts')) {
        // Byte-only SDK exports must also reject reported content loss.
        const fontSetter = 'scopedHostCanvasFont(String(value)) ?? substituteCssFontFamily(String(value))';
        if (!code.includes(fontSetter)) throw new Error('Pinned Canvas font setter changed');
        code = `import { resolveCachedFont } from '@jakal/font-cache';\n` + code.replace(fontSetter, `resolveCachedFont(String(value), () => ${fontSetter})`);
        return code.replace('return this.doc.exportHwpx();', `const artifact = this.exportHwpxWithReport();
          if (artifact.contentLoss.count > 0) throw new Error('저장 과정에서 지원하지 않는 개체가 발견되었습니다. 원본을 유지하고 저장을 중단했습니다.');
          return artifact.bytes;`);
      }
      if (id.replaceAll('\\', '/').endsWith('/core/host-canvas-fonts.ts')) {
        const scope = 'try { return operation(); } finally { currentSession = previous; }';
        if (!code.includes(scope)) throw new Error('Pinned Canvas font scope changed');
        return `import { withFontResolutionCache } from '@jakal/font-cache';\n` + code.replace(scope, 'try { return withFontResolutionCache(session, operation); } finally { currentSession = previous; }');
      }
      if (id.replaceAll('\\', '/').endsWith('/view/canvas-view.ts')) {
        code = code.replaceAll('\r\n', '\n');
        const start = '  private async refreshPagesForMutation(): Promise<void> {';
        const end = '\n  /** document-agent RPC';
        const a = code.indexOf(start), b = code.indexOf(end, a);
        if (a < 0 || b < 0 || !code.includes('export class CanvasView {')) throw new Error('Pinned mutation paint boundary changed');
        code = code.slice(0, a) + `  private async refreshPagesForMutation(): Promise<void> {
          const generation = this.wasm.documentGeneration;
          const selected = await this.selectMutationRevision();
          if (!selected || !this.rendererSession.isCurrent(selected.selection)) return;
          this.jakalMutationPaint.schedule(() => {
            if (this.disposed || generation !== this.wasm.documentGeneration) return;
            this.refreshPages();
            this.eventBus.emit('document-layout-refreshed', { source: 'mutation' });
          });
        }
` + code.slice(b);
        code = code.replace('export class CanvasView {', 'export class CanvasView {\n  private jakalMutationPaint = new FrameTask();');
        code = code.replace('private reset(): void {', 'private reset(): void {\n    this.jakalMutationPaint.cancel();');
        return `import { FrameTask } from '@jakal/frame-task';\n` + code;
      }
    },
    transformIndexHtml(html) {
      return html.replace('<title>rhwp-studio</title>', '<title>Jakal 문서 편집기</title>').replace('src="/theme-init.js"', 'src="/studio/theme-init.js"').replace('src="/locale-init.js"', 'src="/studio/locale-init.js"').replace('href="/favicon.ico"', 'href="/studio/favicon.ico"').replaceAll('href="/icons/', 'href="/studio/icons/').replace('</head>', '<link rel="stylesheet" href="/studio/jakal.css"></head>');
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
