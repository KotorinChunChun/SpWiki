// Builds the core bundle, then runs the unit tests with the Node test runner.
import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { buildCoreBundle } from './build-tests.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

await buildCoreBundle();

const child = spawn(
  process.execPath,
  ['--test', '--test-reporter=spec', 'tests/*.test.mjs'],
  { stdio: 'inherit', cwd: root }
);

child.on('exit', (code) => process.exit(code ?? 1));
