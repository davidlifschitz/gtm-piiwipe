import { it, expect } from 'vitest';
import { loadApp } from './harness.js';

it('redacts supported PII types with counts', () => {
  const { run } = loadApp();
  const result = run('wipeText("jane@example.com (415) 555-0199 078-05-1120 4111 1111 1111 1111 192.168.1.24")');
  expect(result.wiped).toBe('[EMAIL] [PHONE] [SSN] [CARD] [IP]');
  expect(Object.values(result.counts)).toEqual([1,1,1,1,1]);
});
it('rejects invalid card checksums', () => {
  expect(loadApp().run('luhnOk("4111111111111112")')).toBe(false);
});
it('loads the sample without exposing HTML', () => {
  const { window } = loadApp();
  window.document.getElementById('sample').click();
  expect(window.document.getElementById('preview').textContent).toContain('[EMAIL]');
  expect(window.document.getElementById('download').disabled).toBe(false);
});

it('extracts text across Word runs and decodes XML entities', async () => {
  const { run } = loadApp();
  const result = await run('(async()=>{const z=new JSZip();z.file("word/document.xml",\'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>jane&#64;</w:t></w:r><w:r><w:t>example.com</w:t></w:r></w:p><w:p><w:r><w:t>Next &amp; final</w:t></w:r></w:p></w:body></w:document>\');return extractDocxText(await z.generateAsync({type:"uint8array"}));})()');
  expect(result.text).toBe('jane@example.com\nNext & final');
});
it('exports a fresh DOCX containing only scrubbed body text', async () => {
  const { run } = loadApp();
  const parts = await run('(async()=>{const bytes=await createCleanDocx(wipeText("jane@example.com\\nSafe & sound").wiped);const z=await JSZip.loadAsync(bytes);const parts={};for(const [p,f] of Object.entries(z.files)){if(!f.dir)parts[p]=await f.async("string");}return parts;})()');
  expect(parts['word/document.xml']).toContain('[EMAIL]');
  expect(parts['word/document.xml']).toContain('Safe &amp; sound');
  expect(Object.values(parts).join('')).not.toContain('jane@example.com');
  expect(Object.keys(parts).some(p => /word\/media|header|footer/.test(p))).toBe(false);
});
it('rejects malformed DOCX XML instead of silently stripping tags', async () => {
  const { run } = loadApp();
  await expect(run('(async()=>{const z=new JSZip();z.file("word/document.xml","<broken>");return extractDocxText(await z.generateAsync({type:"uint8array"}));})()')).rejects.toThrow(/XML/);
});

it('clears downloadable output when the paste is emptied', () => {
  const { run, window } = loadApp();
  run('els.paste.value="jane@example.com";handlePaste();els.paste.value="";handlePaste()');
  expect(run('state.wiped')).toBe('');
  expect(window.document.getElementById('download').disabled).toBe(true);
  expect(window.document.getElementById('copy').disabled).toBe(true);
});
it('clears the old result when a new file is rejected', async () => {
  const { run, window } = loadApp();
  await run('els.paste.value="safe";handlePaste();handleFile({name:"huge.txt",size:MAX_BYTES+1})');
  expect(run('state.wiped')).toBe('');
  expect(window.document.getElementById('download').disabled).toBe(true);
  expect(window.document.getElementById('error').textContent).toContain('8 MB');
});
it('does not overwrite newer pasted text with a slow file', async () => {
  const { run } = loadApp();
  const pending = run('let release;readAsText=()=>new Promise(r=>release=r);handleFile({name:"slow.txt",size:1})');
  run('els.paste.value="new text";handlePaste();release("old text")');
  await pending;
  expect(run('state.wiped')).toBe('new text');
});

it('uses only classic local scripts for offline file-open operation', () => {
  const { window } = loadApp();
  const scripts = Array.from(window.document.scripts);
  expect(scripts.find(s => s.getAttribute('src') === './app.js').type).not.toBe('module');
  expect(scripts.every(s => !/^https?:/.test(s.getAttribute('src') || ''))).toBe(true);
});
it('uses the bundled PDF parser without a remote module import', async () => {
  const { run, window } = loadApp();
  let destroyed = false;
  window.pdfjsLib = { getDocument: options => {
    expect(options.isEvalSupported).toBe(false);
    return { promise: Promise.resolve({ numPages:1,
      getPage:async()=>({getTextContent:async()=>({items:[{str:'jane@example.com'}]})}),
      destroy:async()=>{destroyed=true;} }) };
  } };
  expect(await run('extractPdfText(new Uint8Array([1,2]))')).toBe('jane@example.com');
  expect(destroyed).toBe(true);
});
