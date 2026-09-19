// JavaScript's \s already covers every Unicode space separator, including
// U+3000 IDEOGRAPHIC SPACE which shows up in Japanese headings.
const WHITESPACE: RegExp = /\s+/g;
const LEADING_TRAILING_WHITESPACE: RegExp = /^\s+|\s+$/g;
const ASCII_PUNCTUATION: RegExp = /[!"#$%&'()*+,./:;<=>?@[\\\]^`{|}~]/g;
const LEADING_TRAILING_DASH: RegExp = /^-+|-+$/g;

/**
 * GitHub-ish heading slug: lower cased, ASCII punctuation dropped, whitespace
 * turned into dashes. Non-ASCII characters (Japanese headings, for instance)
 * are kept as-is, which is what GitHub does too.
 */
export function slugify(text: string): string {
  return text
    .replace(LEADING_TRAILING_WHITESPACE, '')
    .toLowerCase()
    .replace(ASCII_PUNCTUATION, '')
    .replace(WHITESPACE, '-')
    .replace(LEADING_TRAILING_DASH, '');
}
