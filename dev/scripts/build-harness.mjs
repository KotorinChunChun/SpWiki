// Bundles the local preview harness with esbuild.
// The SPFx gulp toolchain is not involved here, so this runs on any Node 20+.
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

await build({
  entryPoints: [path.join(root, 'tests/harness', 'harness.ts')],
  outfile: path.join(root, 'tests/harness', 'dist', 'harness.js'),
  bundle: true,
  format: 'iife',
  platform: 'browser',
  target: 'es2019',
  sourcemap: true,
  logLevel: 'info'
});
