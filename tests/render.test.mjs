import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, render } from './setup.mjs';

const LIBRARY = 'https://tenant.sharepoint.com/sites/docs/Shared%20Documents/';

describe('markdown rendering', () => {
  it('renders headings, lists, tables and inline code', () => {
    const container = render('# 見出し\n\n- 一\n- 二\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n\n`code`');
    const html = container.innerHTML;

    // `textContent`, because the heading also carries the copy button.
    assert.equal(container.querySelector('h1').textContent, '見出し');
    assert.match(html, /<li>一<\/li>/);
    assert.match(html, /<table>/);
    assert.match(html, /<th>a<\/th>/);
    assert.match(html, /<code[^>]*>code<\/code>/);
  });

  it('gives headings an id so in-page links resolve', () => {
    const container = render('## Getting Started\n\n### 導入 手順');

    assert.equal(container.querySelector('h2').getAttribute('id'), 'getting-started');
    assert.equal(container.querySelector('h3').getAttribute('id'), '導入-手順');
  });

  it('de-duplicates repeated heading ids', () => {
    const headings = render('# Notes\n\n# Notes').querySelectorAll('h1');

    assert.equal(headings[0].getAttribute('id'), 'notes');
    assert.equal(headings[1].getAttribute('id'), 'notes-1');
  });
});

describe('code highlighting', () => {
  it('highlights a registered language', () => {
    const code = render('```typescript\nconst a: number = 1;\n```').querySelector('pre > code');

    assert.match(code.className, /hljs/);
    assert.match(code.innerHTML, /hljs-keyword/);
  });

  it('leaves an unregistered language as plain text', () => {
    const code = render('```brainfuck\n+++.\n```').querySelector('pre > code');

    assert.equal(code.innerHTML.indexOf('hljs-'), -1);
    assert.equal(code.textContent, '+++.\n');
  });

  it('supports every language named in the design document', () => {
    const languages = ['javascript', 'typescript', 'powershell', 'json', 'xml', 'css', 'sql', 'csharp'];

    for (const language of languages) {
      const code = render('```' + language + '\nx\n```').querySelector('pre > code');
      assert.match(code.className, /hljs/, `${language} should be registered`);
    }
  });
});

describe('link rewriting', () => {
  it('turns a sibling link into a page URL plus a navigation marker', () => {
    const anchor = render('[install](install.md)', { currentPath: 'guide/index.md' }).querySelector('a');

    assert.equal(anchor.getAttribute(core.DOC_LINK_ATTRIBUTE), 'guide/install.md');
    assert.equal(anchor.getAttribute('href'), '/sites/docs/SitePages/docs.aspx?file=guide/install.md');
  });

  it('resolves a parent reference', () => {
    const anchor = render('[top](../README.md)', { currentPath: 'guide/install.md' }).querySelector('a');

    assert.equal(anchor.getAttribute(core.DOC_LINK_ATTRIBUTE), 'README.md');
  });

  it('keeps the fragment of a markdown link', () => {
    const anchor = render('[step](install.md#step-1)', { currentPath: 'guide/index.md' }).querySelector('a');

    assert.equal(anchor.getAttribute(core.DOC_LINK_ATTRIBUTE), 'guide/install.md#step-1');
    assert.match(anchor.getAttribute('href'), /\?file=guide\/install\.md#step-1$/);
  });

  it('points non-markdown links at the file in the library', () => {
    const anchor = render('[manual](docs/manual.pdf)', { currentPath: 'guide/index.md' }).querySelector('a');

    assert.equal(anchor.getAttribute('href'), LIBRARY + 'guide/docs/manual.pdf');
    assert.equal(anchor.getAttribute(core.DOC_LINK_ATTRIBUTE), null);
  });

  it('opens external links in a new tab', () => {
    const anchor = render('[spfx](https://aka.ms/spfx)').querySelector('a');

    assert.equal(anchor.getAttribute('href'), 'https://aka.ms/spfx');
    assert.equal(anchor.getAttribute('target'), '_blank');
    assert.equal(anchor.getAttribute('rel'), 'noopener noreferrer');
  });

  it('leaves in-page anchors untouched', () => {
    const anchor = render('[top](#overview)').querySelector('a');

    assert.equal(anchor.getAttribute('href'), '#overview');
    assert.equal(anchor.getAttribute('target'), null);
  });
});

describe('image rewriting', () => {
  it('resolves an image next to the document (design 12)', () => {
    const image = render('![logo](images/logo.png)', { currentPath: 'README.md' }).querySelector('img');

    assert.equal(image.getAttribute('src'), LIBRARY + 'images/logo.png');
  });

  it('resolves an image in a sibling folder', () => {
    const image = render('![s](../images/sample.png)', { currentPath: 'guide/install.md' }).querySelector('img');

    assert.equal(image.getAttribute('src'), LIBRARY + 'images/sample.png');
  });

  it('leaves an absolute image URL alone', () => {
    const image = render('![x](https://cdn.example.com/x.png)').querySelector('img');

    assert.equal(image.getAttribute('src'), 'https://cdn.example.com/x.png');
  });
});

describe('sanitization', () => {
  it('removes script, iframe and object elements', () => {
    const container = render(
      '<script>alert(1)</script>\n<iframe src="https://evil.example"></iframe>\n<object data="x"></object>\n\nafter'
    );

    assert.equal(container.querySelector('script'), null);
    assert.equal(container.querySelector('iframe'), null);
    assert.equal(container.querySelector('object'), null);
    assert.match(container.textContent, /after/);
  });

  it('removes inline event handlers', () => {
    const container = render('<img src="x" onerror="alert(1)">');
    const image = container.querySelector('img');

    assert.equal(image === null || image.getAttribute('onerror') === null, true);
    assert.equal(container.innerHTML.indexOf('onerror'), -1);
  });

  it('blocks javascript: links', () => {
    const anchor = render("[x](javascript:alert('x'))").querySelector('a');

    assert.equal(anchor === null || (anchor.getAttribute('href') || '').indexOf('javascript:'), -1);
  });

  it('ignores a data-md-path forged by the document', () => {
    const container = render('<a href="README.md" data-md-path="../../../secret.md">x</a>', {
      currentPath: 'guide/index.md'
    });

    // The forged attribute is stripped by the sanitizer, and the real one is
    // written afterwards from the resolved href.
    assert.equal(container.querySelector('a').getAttribute(core.DOC_LINK_ATTRIBUTE), 'guide/README.md');
  });
});

describe('findDocLinkPath', () => {
  it('finds the marker on an ancestor of the clicked node', () => {
    const container = render('[**bold link**](install.md)', { currentPath: 'guide/index.md' });
    const strong = container.querySelector('strong');

    assert.equal(core.findDocLinkPath(strong, container), 'guide/install.md');
  });

  it('returns an empty string outside a document link', () => {
    const container = render('plain paragraph');

    assert.equal(core.findDocLinkPath(container.querySelector('p'), container), '');
  });
});
