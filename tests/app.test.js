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
