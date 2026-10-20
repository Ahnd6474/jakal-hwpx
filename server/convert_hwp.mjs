// Run only inside the bounded document worker. Never accepts URLs or commands.
import fs from 'node:fs';
import { HwpDocument, initSync } from '../web/node_modules/@rhwp/core/rhwp.js';

initSync({ module: fs.readFileSync(new URL('../web/node_modules/@rhwp/core/rhwp_bg.wasm', import.meta.url)) });
const [source, target, report] = process.argv.slice(2);
const document = new HwpDocument(fs.readFileSync(source));
try {
  const exported = document.exportHwpxWithReport();
  try {
    const loss = JSON.parse(exported.contentLoss());
    // A successful export alone does not establish that unsupported content survived.
    if (loss.count !== 0 || !Array.isArray(loss.losses) || loss.losses.length) {
      throw new Error('unsupported content');
    }
    const bytes = exported.takeBytes();
    if (!bytes.length || bytes.length > 20 * 1024 * 1024) throw new Error('output size');
    const reopened = new HwpDocument(bytes);
    reopened.free();
    fs.writeFileSync(target, bytes, { flag: 'wx' });
    fs.writeFileSync(report, JSON.stringify({ engine: '@rhwp/core 0.8.7', ...loss }), { flag: 'wx' });
  } finally { exported.free(); }
} finally { document.free(); }
