import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { JSDOM } from 'jsdom';
import JSZip from 'jszip';

export function loadApp() {
  const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
  const dom = new JSDOM(html, { runScripts: 'outside-only', url: 'https://example.test/' });
  const context = dom.getInternalVMContext();
  context.JSZip = JSZip;
  context.TextDecoder = TextDecoder;
  vm.runInContext(readFileSync(new URL('../app.js', import.meta.url), 'utf8'), context);
  return { window: dom.window, run: code => vm.runInContext(code, context) };
}
