// Rebuild local browser libraries with npm ci && npm run vendor.
import { build } from 'esbuild';
import { copyFileSync } from 'node:fs';
const copies = [
  ['jszip/dist/jszip.min.js', 'jszip.min.js'],
  ['jszip/LICENSE.markdown', 'JSZip-LICENSE.markdown'],
  ['docx/dist/index.iife.js', 'docx.js'],
  ['docx/LICENSE', 'DOCX-LICENSE'],
  ['pdf-lib/dist/pdf-lib.min.js', 'pdf-lib.min.js'],
  ['pdf-lib/LICENSE.md', 'PDF-lib-LICENSE.md'],
  ['pdfjs-dist/LICENSE', 'PDFjs-LICENSE'],
];
for (const [source, target] of copies) copyFileSync('node_modules/' + source, 'vendor/' + target);
for (const [file, globalName] of [['pdf', 'pdfjsLib'], ['pdf.worker', 'pdfjsWorker']]) {
  await build({
    entryPoints: ['node_modules/pdfjs-dist/legacy/build/' + file + '.mjs'],
    outfile: 'vendor/' + file + '.js',
    format: 'iife', globalName, minify: true, target: 'es2022',
    legalComments: 'inline',
  });
}
