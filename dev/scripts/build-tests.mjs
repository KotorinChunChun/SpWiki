// Bundles the core modules into a single ESM file so the unit tests can import
// them from plain Node without a TypeScript loader.
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

export const bundlePath = path.join(root, 'temp', 'core.bundle.mjs');

export async function buildCoreBundle() {
  await build({
    entryPoints: [path.join(root, 'src', 'webparts', 'spWiki', 'core', 'index.ts')],
    outfile: bundlePath,
    bundle: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2019',
    logLevel: 'warning'
  });
  return bundlePath;
}

if (import.meta.url === `file://${process.argv[1]}` || process.argv[1]?.endsWith('build-tests.mjs')) {
  await buildCoreBundle();
}
