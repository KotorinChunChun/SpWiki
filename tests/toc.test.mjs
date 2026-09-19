import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, dom, render } from './setup.mjs';

/** The generated list as `level:text -> href` lines, in document order. */
function readToc(container) {
  const toc = container.querySelector('.mdv-toc');
  if (!toc) {
    return undefined;
  }

  return [...toc.querySelectorAll('li')].map((item) => {
    const link = item.querySelector('a');
    let depth = 0;
    for (let node = item.parentNode; node && node !== toc; node = node.parentNode) {
      if (node.nodeName.toLowerCase() === 'ul') {
        depth += 1;
      }
    }
    return `${depth}:${link.textContent} -> ${link.getAttribute('href')}`;
  });
}

describe('table of contents', () => {
  it('lists the headings that follow a 目次 heading', () => {
    const container = render('# 題名\n\n## 目次\n\n## 前提\n\n## 手順\n\n### 詳細\n\n## まとめ');

    assert.deepEqual(readToc(container), [
      '1:前提 -> #前提',
      '1:手順 -> #手順',
      '2:詳細 -> #詳細',
      '1:まとめ -> #まとめ'
    ]);
  });

  it('leaves out the heading the marker itself sits on, and anything above it', () => {
    const container = render('# 題名\n\n## はじめに\n\n## 目次\n\n## 本編');
    const entries = readToc(container);

    assert.equal(entries.length, 1);
    assert.match(entries[0], /本編/);
  });

  it('inserts the list right after the marker', () => {
    const container = render('## 目次\n\n## 本編');
    const marker = container.querySelector('h2');

    assert.equal(marker.nextElementSibling.className, 'mdv-toc');
  });

  it('does nothing to a document without the marker', () => {
    assert.equal(render('# 題名\n\n## 手順').querySelector('.mdv-toc'), null);
  });

  it('does nothing when no heading follows the marker', () => {
    assert.equal(render('# 題名\n\n## 目次').querySelector('.mdv-toc'), null);
  });

  it('copes with a skipped level', () => {
    const container = render('## 目次\n\n## 一\n\n#### 深い\n\n## 二');

    assert.deepEqual(readToc(container), ['1:一 -> #一', '2:深い -> #深い', '1:二 -> #二']);
  });

  it('links to the de-duplicated ids so repeated headings stay distinct', () => {
    const container = render('## 目次\n\n## 注意\n\n## 注意');

    assert.deepEqual(readToc(container), ['1:注意 -> #注意', '1:注意 -> #注意-1']);
  });

  it('works in an html document too', () => {
    const container = render('<h2>目次</h2><h2>本編</h2><h3>詳細</h3>', { currentPath: 'page.html' });

    assert.deepEqual(readToc(container), ['1:本編 -> #本編', '2:詳細 -> #詳細']);
  });

  it('ignores a heading that merely contains the word', () => {
    assert.equal(render('## 目次について\n\n## 本編').querySelector('.mdv-toc'), null);
  });
});

describe('buildHeadingNav', () => {
  const { buildHeadingNav } = core;

  const page = (html) => {
    const host = dom.window.document.createElement('div');
    host.innerHTML = html;
    return host;
  };

  it('lists every heading that can be linked to, without needing a marker', () => {
    // The outline button asks "what is in this page", which is a question about
    // any page — not only the ones whose author asked for a contents list.
    const host = page('<h1 id="a">A</h1><h2 id="b">B</h2><h3 id="c">C</h3>');
    const nav = buildHeadingNav(host);

    assert.deepEqual([...nav.querySelectorAll('a')].map((x) => x.textContent), ['A', 'B', 'C']);
    assert.deepEqual([...nav.querySelectorAll('a')].map((x) => x.getAttribute('href')), ['#a', '#b', '#c']);
  });

  it('nests the levels', () => {
    const host = page('<h1 id="a">A</h1><h2 id="b">B</h2>');
    const nav = buildHeadingNav(host);

    assert.equal(nav.querySelectorAll('ul ul').length, 1);
  });

  it('skips a heading with no id, since there is nothing to link to', () => {
    const host = page('<h1 id="a">A</h1><h2>B</h2>');

    assert.deepEqual([...buildHeadingNav(host).querySelectorAll('a')].map((x) => x.textContent), ['A']);
  });

  it('leaves the generated table of contents out of itself', () => {
    const host = page('<h1 id="a">A</h1><nav class="mdv-toc"><ul><li><h2 id="x">X</h2></li></ul></nav>');

    assert.deepEqual([...buildHeadingNav(host).querySelectorAll('a')].map((x) => x.textContent), ['A']);
  });

  it('returns undefined for a page with nothing to list', () => {
    assert.equal(buildHeadingNav(page('<p>本文だけ</p>')), undefined);
  });
});
