import { mkdirSync, readdirSync, renameSync, rmSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const SRC_DIR = 'dist/src/pages';
const DEST_DIR = 'dist';

if (!existsSync(SRC_DIR)) {
  console.log(`[flatten] Fant ikke ${SRC_DIR}. Ingenting å gjøre.`);
  process.exit(0);
}

for (const file of readdirSync(SRC_DIR)) {
  if (file.endsWith('.html')) {
    const from = join(SRC_DIR, file);
    const to = join(DEST_DIR, file);
    console.log(`[flatten] flytter ${from} -> ${to}`);
    renameSync(from, to);
  }
}

console.log('[flatten] rydder dist/src');
rmSync('dist/src', { recursive: true, force: true });
