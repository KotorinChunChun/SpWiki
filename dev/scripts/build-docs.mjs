/**
 * README.md と docs/ の利用者向け文書を同梱ヘルプに変換します。
 * 開発者向け資料は含めません。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const target = path.join(root, 'src', 'webparts', 'spWiki', 'core', 'bundledDocs.ts');

/**
 * `key` is what appears in `?help=<key>`, so it is short, ASCII and stable —
 * these URLs get shared. `title` is what the reader sees.
 */
const DOCS = [
  { key: 'overview', file: 'README.md', title: '概要', titleEn: 'Overview' },
  { key: 'user', file: 'docs/USERS_GUIDE.md', title: '見る人向け', titleEn: 'For readers' },
  { key: 'writer', file: 'docs/WRITERS_GUIDE.md', title: '書く人向け', titleEn: 'For writers' },
  { key: 'admin', file: 'docs/ADMIN_GUIDE.md', title: '管理者向け', titleEn: 'For administrators' }
];

/** Maps a document file name onto the key that opens it, for cross links. */
const BY_FILE = {};
for (const doc of DOCS) {
  BY_FILE[doc.file] = doc.key;
}

function read(file) {
  return fs.readFileSync(path.join(root, file), 'utf8');
}

/**
 * Rewrites links between the bundled documents so they stay inside the help
 * page. A link to a document that is *not* bundled is left alone: it points at
 * the repository, and turning it into a dead help link would be worse.
 */
function rewriteCrossLinks(markdown, sourceFile) {
  return markdown.replace(/\]\(([^)\s]+\.md)(#[^)]*)?\)/g, (whole, file, hash) => {
    const resolved = path.posix.normalize(path.posix.join(path.posix.dirname(sourceFile), file));
    const key = BY_FILE[resolved];
    return key ? `](?help=${key}${hash || ''})` : whole;
  });
}

const entries = DOCS.map((doc) => {
  const markdown = rewriteCrossLinks(read(doc.file), doc.file);

  return (
    `  {\n` +
    `    key: ${JSON.stringify(doc.key)},\n` +
    `    title: ${JSON.stringify(doc.title)},\n` +
    `    titleEn: ${JSON.stringify(doc.titleEn)},\n` +
    `    markdown: ${JSON.stringify(markdown)}\n` +
    `  }`
  );
});

const source =
  `/*\n` +
  ` * GENERATED FILE — do not edit.\n` +
  ` *\n` +
  ` * Produced by dev/scripts/build-docs.mjs from README.md and docs/.\n` +
  ` * Edit those documents and run \`npm run docs:build\` (the release script does\n` +
  ` * it for you).\n` +
  ` */\n\n` +
  `/** One document shipped inside the bundle and shown on the help page. */\n` +
  `export interface IBundledDoc {\n` +
  `  /** Value of \`?help=\`; short and stable, because these URLs get shared. */\n` +
  `  key: string;\n` +
  `  title: string;\n` +
  `  titleEn: string;\n` +
  `  markdown: string;\n` +
  `}\n\n` +
  `export const BUNDLED_DOCS: IBundledDoc[] = [\n` +
  entries.join(',\n') +
  `\n];\n\n` +
  `/** The document \`?help=<key>\` asks for, or undefined when there is no such key. */\n` +
  `export function findBundledDoc(key: string): IBundledDoc | undefined {\n` +
  `  for (let i: number = 0; i < BUNDLED_DOCS.length; i++) {\n` +
  `    if (BUNDLED_DOCS[i].key === key) {\n` +
  `      return BUNDLED_DOCS[i];\n` +
  `    }\n` +
  `  }\n` +
  `  return undefined;\n` +
  `}\n`;

fs.writeFileSync(target, source, 'utf8');

const bytes = Buffer.byteLength(source, 'utf8');
console.log(`bundledDocs.ts  ${DOCS.length} documents, ${(bytes / 1024).toFixed(1)} KB`);
