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
