import { copyFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(fileURLToPath(import.meta.url));
const vendor = resolve(root, '../../src/MultiPPTController/PlayerHost/vendor');
mkdirSync(vendor, { recursive: true });

const files = [
    ['node_modules/pptx-vanilla-viewer/dist/styles.css', 'styles.css'],
    ['node_modules/pptx-vanilla-viewer/LICENSE', 'LICENSE'],
    ['node_modules/pptx-vanilla-viewer/NOTICE', 'NOTICE'],
    ['node_modules/pdfjs-dist/build/pdf.worker.min.mjs', 'pdf.worker.min.mjs']
];

for (const [from, name] of files) {
    copyFileSync(resolve(root, from), resolve(vendor, name));
}
