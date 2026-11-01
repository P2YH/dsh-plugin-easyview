/** Backed-up local delta for Word AI reading; does not reload an open document. */
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, writeFile, mkdir, rename, realpath, unlink } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
assert.ok(process.env.DSH_DESKTOP_PROFILE, 'Set explicit DSH_DESKTOP_PROFILE');
const profile = resolve(process.env.DSH_DESKTOP_PROFILE);
const installed = await realpath(resolve(profile, 'node_modules/dsh-plugin-easyview'));
assert.equal(JSON.parse(await readFile(resolve(installed, 'package.json'))).name, 'dsh-plugin-easyview');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const manifest = JSON.parse(await readFile(resolve(root, 'lib/workbench/asset-manifest.json')));
const previousManifest = JSON.parse(await readFile(resolve(installed, 'lib/workbench/asset-manifest.json')));
const staticFiles = files => files.filter(file => !file.endsWith('.js')).sort();
assert.deepEqual(staticFiles(manifest.files), staticFiles(previousManifest.files), 'Installed static assets need a full compatible build');
for (const file of staticFiles(manifest.files)) {
  const source = await readFile(resolve(root, 'lib/workbench', file)), target = await readFile(resolve(installed, 'lib/workbench', file));
  if (['vendor/file-viewer/file-viewer-copy-assets.manifest.json', 'vendor/file-viewer/file-viewer-copy-assets.receipt.json', 'vendor/file-viewer/flyfish-viewer-assets.json'].includes(file)) {
    const current = JSON.parse(source), previous = JSON.parse(target);
    delete current.generatedAt; delete previous.generatedAt;
    delete current.validation?.checkedAt; delete previous.validation?.checkedAt;
    delete current.installedAt; delete previous.installedAt;
    delete current.assetManifestSha256; delete previous.assetManifestSha256;
    assert.equal(hash(JSON.stringify(current)), hash(JSON.stringify(previous)), `Different installed engine: ${file}`);
  } else assert.equal(hash(source), hash(target), `Different installed engine: ${file}`);
}
const files = [];
for (const relative of ['lib/index.js', 'lib/ai-preset.js', ...manifest.files.filter(file => file.endsWith('.js')).map(file => `lib/workbench/${file}`), 'lib/workbench/asset-manifest.json']) {
  const before = await readFile(resolve(installed, relative)).catch(error => { if (error.code === 'ENOENT') return null; throw error; });
  const after = await readFile(resolve(root, relative));
  if (!before?.equals(after)) files.push({ relative, before, after });
}
assert.ok(files.length, 'Word AI update already installed');
const backup = resolve(profile, '.desktop-plugin-backups', `docx-ai-${Date.now()}`);
for (const file of files) {
  await mkdir(resolve(backup, file.relative, '..'), { recursive: true });
  if (file.before) await writeFile(resolve(backup, file.relative), file.before);
}
const changed = [];
try {
  for (const file of files) {
    const target = resolve(installed, file.relative), temporary = `${target}.docx-ai.tmp`;
    await mkdir(resolve(target, '..'), { recursive: true });
    await writeFile(temporary, file.after); await rename(temporary, target); changed.push(file);
    assert.equal(hash(await readFile(target)), hash(file.after));
  }
} catch (error) {
  for (const file of changed.reverse()) {
    if (file.before) await writeFile(resolve(installed, file.relative), file.before);
    else await unlink(resolve(installed, file.relative));
  }
  throw error;
}
const result = { status: 'installed', installedAt: new Date().toISOString(), profile, installed, backup,
  files: files.map(({ relative, before, after }) => ({ relative, beforeSha256: before ? hash(before) : null, afterSha256: hash(after) })),
  effect: 'Restart the desktop app to load the updated Host and renderer. Open documents are not reloaded.',
  scope: 'Word AI read adapter, tool description and editor readiness only. Profile configuration and document data are untouched.' };
await writeFile(resolve(backup, 'update.json'), JSON.stringify(result, null, 2) + '\n');
await mkdir(resolve(root, 'docs/evidence'), { recursive: true });
await writeFile(resolve(root, 'docs/evidence/docx-ai-desktop-install.json'), JSON.stringify(result, null, 2) + '\n');
console.log(JSON.stringify(result, null, 2));
