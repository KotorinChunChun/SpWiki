import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, render, renderHtml } from './setup.mjs';

const LIBRARY = 'https://tenant.sharepoint.com/sites/docs/Shared%20Documents/';

describe('html rendering', () => {
  it('takes the markup as it is instead of parsing it as Markdown', () => {
    const container = renderHtml('<h1>見出し</h1><p>**not bold**</p>');

    assert.equal(container.querySelector('h1').textContent, '見出し');
    assert.equal(container.querySelector('strong'), null);
    assert.match(container.querySelector('p').textContent, /\*\*not bold\*\*/);
  });

  it('keeps the body of a whole document and drops the head', () => {
    const container = renderHtml(
      '<!doctype html><html lang="ja"><head><title>タイトル</title>' +
        '<style>body { color: red }</style></head>' +
        '<body><p>本文</p></body></html>'
    );

    assert.equal(container.querySelector('style'), null);
    assert.equal(container.textContent.indexOf('タイトル'), -1);
    assert.equal(container.textContent.indexOf('color: red'), -1);
    assert.match(container.querySelector('p').textContent, /本文/);
  });

  it('gives headings an id, like a Markdown document', () => {
    const container = renderHtml('<h2>Getting Started</h2><h2>Getting Started</h2>');
    const headings = container.querySelectorAll('h2');

    assert.equal(headings[0].getAttribute('id'), 'getting-started');
    assert.equal(headings[1].getAttribute('id'), 'getting-started-1');
  });

  it('highlights a fenced code block written as pre > code', () => {
    const container = renderHtml('<pre><code class="language-typescript">const a: number = 1;</code></pre>');
    const code = container.querySelector('pre > code');

    assert.match(code.className, /hljs/);
    assert.match(code.innerHTML, /hljs-keyword/);
  });

  it('keeps the structural elements an HTML page uses', () => {
    const container = renderHtml(
      '<section><figure><img src="images/logo.png" alt="logo">' +
        '<figcaption>図 1</figcaption></figure>' +
        '<details><summary>詳細</summary><p>中身</p></details></section>'
    );

    assert.notEqual(container.querySelector('section figure figcaption'), null);
    assert.notEqual(container.querySelector('details > summary'), null);
  });
});

describe('html link rewriting', () => {
  it('resolves a link to another document relative to the html file', () => {
    const anchor = renderHtml('<a href="install.md">install</a>', {
      currentPath: 'guide/index.html'
    }).querySelector('a');

    assert.equal(anchor.getAttribute(core.DOC_LINK_ATTRIBUTE), 'guide/install.md');
    assert.equal(anchor.getAttribute('href'), '/sites/docs/SitePages/docs.aspx?file=guide/install.md');
  });

  it('navigates in place between html documents', () => {
    const anchor = renderHtml('<a href="../top.htm">top</a>', {
      currentPath: 'guide/index.html'
    }).querySelector('a');

    assert.equal(anchor.getAttribute(core.DOC_LINK_ATTRIBUTE), 'top.htm');
  });

  it('resolves images relative to the html file', () => {
    const image = renderHtml('<img src="../images/sample.png" alt="s">', {
      currentPath: 'guide/index.html'
    }).querySelector('img');

    assert.equal(image.getAttribute('src'), LIBRARY + 'images/sample.png');
    assert.equal(image.getAttribute('loading'), 'lazy');
  });

  it('opens external links in a new tab', () => {
    const anchor = renderHtml('<a href="https://aka.ms/spfx">spfx</a>').querySelector('a');

    assert.equal(anchor.getAttribute('target'), '_blank');
    assert.equal(anchor.getAttribute('rel'), 'noopener noreferrer');
  });

  it('points a link to an html file inside a Markdown document at the viewer', () => {
    const anchor = render('[page](page.html)', { currentPath: 'guide/index.md' }).querySelector('a');

    assert.equal(anchor.getAttribute(core.DOC_LINK_ATTRIBUTE), 'guide/page.html');
  });
});

describe('html sanitization', () => {
  it('removes script, iframe, object and form elements', () => {
    const container = renderHtml(
      '<p>before</p><script>alert(1)</script><iframe src="https://evil.example"></iframe>' +
        '<object data="x"></object><form action="https://evil.example"><input name="a"></form><p>after</p>'
    );

    assert.equal(container.querySelector('script'), null);
    assert.equal(container.querySelector('iframe'), null);
    assert.equal(container.querySelector('object'), null);
    assert.equal(container.querySelector('form'), null);
    assert.match(container.textContent, /before/);
    assert.match(container.textContent, /after/);
  });

  it('removes inline event handlers', () => {
    const container = renderHtml('<img src="x.png" onerror="alert(1)"><div onclick="alert(2)">x</div>');

    assert.equal(container.innerHTML.indexOf('onerror'), -1);
    assert.equal(container.innerHTML.indexOf('onclick'), -1);
  });

  it('blocks javascript: links', () => {
    const anchor = renderHtml('<a href="javascript:alert(1)">x</a>').querySelector('a');

    assert.equal(anchor === null || (anchor.getAttribute('href') || '').indexOf('javascript:'), -1);
  });

  it('drops the style attribute so the document cannot restyle the page', () => {
    const container = renderHtml('<p style="position:fixed;top:0">x</p>');

    assert.equal(container.querySelector('p').getAttribute('style'), null);
  });

  it('ignores a data-md-path forged by the document', () => {
    const container = renderHtml('<a href="README.md" data-md-path="../../../secret.md">x</a>', {
      currentPath: 'guide/index.html'
    });

    assert.equal(container.querySelector('a').getAttribute(core.DOC_LINK_ATTRIBUTE), 'guide/README.md');
  });
});

describe('format selection', () => {
  it('follows the extension of the document being rendered', () => {
    assert.notEqual(render('<h1>x</h1>', { currentPath: 'a.html' }).querySelector('h1'), null);
    assert.notEqual(render('# x', { currentPath: 'a.md' }).querySelector('h1'), null);
  });

  it('can be overridden explicitly', () => {
    const asMarkdown = render('<h1>x</h1>\n\n# y', { currentPath: 'a.html' }, 'markdown');

    // Markdown keeps the raw block, and still turns the ATX heading into one.
    assert.equal(asMarkdown.querySelectorAll('h1').length, 2);
  });
});
