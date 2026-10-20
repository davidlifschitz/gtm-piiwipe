import { readFileSync } from 'node:fs';
import { it, expect } from 'vitest';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { loadApp } from './harness.js';

it('extracts and wipes an actual PDF using the vendored offline parser', async () => {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage();
  page.drawText('jane@example.com', { x: 40, y: 700, font, size: 12 });
  const bytes = await pdf.save();
  const { window, run } = loadApp();
  Object.assign(window, {
    TextEncoder, TextDecoder, ReadableStream, TransformStream,
    Response, Request, Headers, structuredClone,
    URL: globalThis.URL, URLSearchParams,
  });
  // JSDOM exposes the actual VM context through its script execution.
  window.eval(readFileSync(new URL('../vendor/pdf.worker.js', import.meta.url), 'utf8'));
  window.eval(readFileSync(new URL('../vendor/pdf.js', import.meta.url), 'utf8'));
  window.fixture = bytes;
  const text = await run('extractPdfText(fixture)');
  expect(text).toContain('jane@example.com');
  expect(run(`wipeText(${JSON.stringify(text)}).wiped`)).toContain('[EMAIL]');
});
