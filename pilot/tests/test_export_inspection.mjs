import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { relocateExportInspection } from '../scripts/export_inspection.mjs';

const validation = fileURLToPath(new URL('../validation/', import.meta.url));
async function fixture(t) {
  await fs.mkdir(validation, { recursive: true });
  const root = await fs.mkdtemp(path.join(validation, 'inspection-move-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const source = path.join(root, 'report.xlsx.inspect.ndjson');
  const destination = path.join(root, 'latest-export-inspect.ndjson');
  await fs.writeFile(source, '{"kind":"workbook"}\n');
  await fs.writeFile(destination, 'previous run');
  return { source, destination };
}
const failure = code => Object.assign(new Error(code), { code });
const crossDevice = { ...fs, rename: async () => { throw failure('EXDEV'); } };

test('relocates the sidecar and replaces the previous diagnostic', async t => {
  const { source, destination } = await fixture(t);
  await relocateExportInspection(source, destination);
  assert.equal(await fs.readFile(destination, 'utf8'), '{"kind":"workbook"}\n');
  await assert.rejects(fs.access(source), { code: 'ENOENT' });
});

test('copies across a filesystem boundary before removing the sidecar', async t => {
  const { source, destination } = await fixture(t);
  await relocateExportInspection(source, destination, crossDevice);
  assert.equal(await fs.readFile(destination, 'utf8'), '{"kind":"workbook"}\n');
  await assert.rejects(fs.access(source), { code: 'ENOENT' });
});

test('accepts a runtime that emits no inspection sidecar', async t => {
  const { source, destination } = await fixture(t);
  await fs.unlink(source);
  await relocateExportInspection(source, destination);
  assert.equal(await fs.readFile(destination, 'utf8'), 'previous run');
});

test('preserves the source when copying fails and reports the error', async t => {
  const { source, destination } = await fixture(t);
  await assert.rejects(relocateExportInspection(source, destination, {
    ...crossDevice, copyFile: async () => { throw failure('ENOSPC'); },
  }), { code: 'ENOSPC' });
  assert.equal(await fs.readFile(source, 'utf8'), '{"kind":"workbook"}\n');
  assert.equal(await fs.readFile(destination, 'utf8'), 'previous run');
});

test('reports unrelated move errors without deleting diagnostics', async t => {
  const { source, destination } = await fixture(t);
  await assert.rejects(relocateExportInspection(source, destination, {
    ...fs, rename: async () => { throw failure('EACCES'); },
  }), { code: 'EACCES' });
  assert.equal(await fs.readFile(source, 'utf8'), '{"kind":"workbook"}\n');
  assert.equal(await fs.readFile(destination, 'utf8'), 'previous run');
});

test('reports source cleanup errors while retaining the successful copy', async t => {
  const { source, destination } = await fixture(t);
  await assert.rejects(relocateExportInspection(source, destination, {
    ...crossDevice, unlink: async () => { throw failure('EACCES'); },
  }), { code: 'EACCES' });
  assert.equal(await fs.readFile(source, 'utf8'), '{"kind":"workbook"}\n');
  assert.equal(await fs.readFile(destination, 'utf8'), '{"kind":"workbook"}\n');
});
