import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core } from './setup.mjs';

const {
  dirname,
  documentFormat,
  encodePath,
  isDocumentPath,
  isExternalLink,
  isHtmlPath,
  isMarkdownPath,
  joinUrl,
  normalizePath,
  resolveRelativePath,
  sanitizeDocPath,
  splitLink
} = core;

describe('resolveRelativePath', () => {
  it('resolves a sibling document (design 6.2, first example)', () => {
    assert.equal(resolveRelativePath('guide/index.md', 'install.md'), 'guide/install.md');
  });

  it('resolves a parent reference (design 6.2, second example)', () => {
    assert.equal(resolveRelativePath('guide/install.md', '../README.md'), 'README.md');
  });

  it('resolves from a library root document', () => {
    assert.equal(resolveRelativePath('README.md', 'guide/install.md'), 'guide/install.md');
  });

  it('resolves an explicit ./ prefix', () => {
    assert.equal(resolveRelativePath('guide/index.md', './faq.md'), 'guide/faq.md');
  });

  it('resolves several levels up', () => {
    assert.equal(resolveRelativePath('a/b/c/d.md', '../../x.md'), 'a/x.md');
  });

  it('treats a leading slash as library relative', () => {
    assert.equal(resolveRelativePath('guide/deep/x.md', '/README.md'), 'README.md');
  });

  it('cannot climb above the library root', () => {
    assert.equal(resolveRelativePath('README.md', '../../../../etc/passwd'), 'etc/passwd');
    assert.equal(resolveRelativePath('guide/install.md', '../../../secret.md'), 'secret.md');
  });
});

describe('normalizePath', () => {
  it('collapses redundant separators and dot segments', () => {
    assert.equal(normalizePath('a//b/./c'), 'a/b/c');
    assert.equal(normalizePath('./a/b'), 'a/b');
    assert.equal(normalizePath('a/b/../c'), 'a/c');
  });

  it('normalizes backslashes', () => {
    assert.equal(normalizePath('guide\\install.md'), 'guide/install.md');
  });
});

describe('dirname', () => {
  it('returns the containing folder', () => {
    assert.equal(dirname('guide/install.md'), 'guide');
    assert.equal(dirname('a/b/c.md'), 'a/b');
  });

  it('returns an empty string for library root documents', () => {
    assert.equal(dirname('README.md'), '');
  });
});

describe('isExternalLink', () => {
  it('detects schemes and protocol relative URLs', () => {
    assert.equal(isExternalLink('https://example.com'), true);
    assert.equal(isExternalLink('http://example.com'), true);
    assert.equal(isExternalLink('mailto:a@example.com'), true);
    assert.equal(isExternalLink('data:image/png;base64,AAA'), true);
    assert.equal(isExternalLink('//example.com/x'), true);
  });

  it('treats library paths as internal', () => {
    assert.equal(isExternalLink('guide/install.md'), false);
    assert.equal(isExternalLink('../README.md'), false);
    assert.equal(isExternalLink('/README.md'), false);
  });
});

describe('splitLink', () => {
  it('separates path, query and hash', () => {
    assert.deepEqual(splitLink('guide/install.md?x=1#top'), {
      path: 'guide/install.md',
      query: '?x=1',
      hash: '#top'
    });
  });

  it('handles a bare path', () => {
    assert.deepEqual(splitLink('README.md'), { path: 'README.md', query: '', hash: '' });
  });

  it('handles a hash placed before a question mark', () => {
    assert.deepEqual(splitLink('a.md#s?x'), { path: 'a.md', query: '', hash: '#s?x' });
  });
});

describe('isMarkdownPath', () => {
  it('accepts .md and .markdown in any casing', () => {
    assert.equal(isMarkdownPath('README.md'), true);
    assert.equal(isMarkdownPath('a/b.MD'), true);
    assert.equal(isMarkdownPath('a.markdown'), true);
  });

  it('rejects everything else', () => {
    assert.equal(isMarkdownPath('images/logo.png'), false);
    assert.equal(isMarkdownPath('notes.txt'), false);
    assert.equal(isMarkdownPath('md'), false);
    assert.equal(isMarkdownPath('page.html'), false);
  });
});

describe('isHtmlPath', () => {
  it('accepts .html and .htm in any casing', () => {
    assert.equal(isHtmlPath('page.html'), true);
    assert.equal(isHtmlPath('guide/old.HTM'), true);
    assert.equal(isHtmlPath('a/b.Html'), true);
  });

  it('rejects everything else', () => {
    assert.equal(isHtmlPath('README.md'), false);
    assert.equal(isHtmlPath('page.aspx'), false);
    assert.equal(isHtmlPath('page.xhtml'), false);
    assert.equal(isHtmlPath('html'), false);
  });
});

describe('isDocumentPath', () => {
  it('covers both formats the viewer renders', () => {
    assert.equal(isDocumentPath('README.md'), true);
    assert.equal(isDocumentPath('a.markdown'), true);
    assert.equal(isDocumentPath('guide/page.html'), true);
    assert.equal(isDocumentPath('guide/page.htm'), true);
  });

  it('rejects anything else, so it can gate the query string', () => {
    assert.equal(isDocumentPath('web.config'), false);
    assert.equal(isDocumentPath('images/logo.png'), false);
    assert.equal(isDocumentPath('notes.txt'), false);
  });
});

describe('documentFormat', () => {
  it('reads the format off the extension', () => {
    assert.equal(documentFormat('README.md'), 'markdown');
    assert.equal(documentFormat('a.markdown'), 'markdown');
    assert.equal(documentFormat('guide/page.HTML'), 'html');
    assert.equal(documentFormat('guide/page.htm'), 'html');
  });

  it('falls back to markdown for anything unfamiliar', () => {
    assert.equal(documentFormat(''), 'markdown');
    assert.equal(documentFormat('notes.txt'), 'markdown');
  });
});

describe('encodePath / joinUrl', () => {
  it('encodes each segment but keeps the separators', () => {
    assert.equal(encodePath('Shared Documents/guide/install.md'), 'Shared%20Documents/guide/install.md');
    assert.equal(encodePath('日本語/手順.md'), `${encodeURIComponent('日本語')}/${encodeURIComponent('手順')}.md`);
  });

  it('joins without doubling slashes', () => {
    assert.equal(joinUrl('/sites/docs/Shared Documents/', '/guide/x.md'), '/sites/docs/Shared Documents/guide/x.md');
    assert.equal(joinUrl('/sites/docs', 'README.md'), '/sites/docs/README.md');
  });
});

describe('sanitizeDocPath', () => {
  it('accepts a normal markdown path', () => {
    assert.equal(sanitizeDocPath('guide/install.md'), 'guide/install.md');
  });

  it('decodes percent-encoded values', () => {
    assert.equal(sanitizeDocPath('guide%2Finstall.md'), 'guide/install.md');
  });

  it('drops the query string and fragment', () => {
    assert.equal(sanitizeDocPath('guide/install.md#step-1'), 'guide/install.md');
  });

  it('refuses absolute URLs', () => {
    assert.equal(sanitizeDocPath('https://evil.example.com/x.md'), '');
    assert.equal(sanitizeDocPath('//evil.example.com/x.md'), '');
  });

  it('accepts an html path', () => {
    assert.equal(sanitizeDocPath('guide/page.html'), 'guide/page.html');
    assert.equal(sanitizeDocPath('guide%2Fpage.htm'), 'guide/page.htm');
  });

  it('refuses files the viewer does not render', () => {
    assert.equal(sanitizeDocPath('web.config'), '');
    assert.equal(sanitizeDocPath('images/logo.png'), '');
    assert.equal(sanitizeDocPath('page.aspx'), '');
  });

  it('contains directory traversal inside the library', () => {
    assert.equal(sanitizeDocPath('../../../../README.md'), 'README.md');
  });

  it('returns an empty string for empty input', () => {
    assert.equal(sanitizeDocPath(''), '');
  });
});
