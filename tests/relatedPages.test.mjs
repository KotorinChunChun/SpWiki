import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, dom } from './setup.mjs';

const { buildProtocolUrl, filterPagesByTag, renderTagSearch } = core;

describe('buildProtocolUrl', () => {
  it('swaps the scheme and keeps the rest of the URL untouched', () => {
    assert.equal(
      buildProtocolUrl('spmd', 'https://contoso.sharepoint.com/sites/docs/Shared%20Documents/md/README.md'),
      'spmd://contoso.sharepoint.com/sites/docs/Shared%20Documents/md/README.md'
    );
  });

  it('handles an http URL too', () => {
    assert.equal(buildProtocolUrl('spmd', 'http://host/a.md'), 'spmd://host/a.md');
  });

  it('leaves the percent-encoding exactly as SharePoint produced it', () => {
    const url = buildProtocolUrl('spmd', 'https://host/a%20b/c%23d.md');

    assert.equal(url, 'spmd://host/a%20b/c%23d.md');
    assert.equal(url.indexOf(' '), -1);
  });

  it('strips anything that is not valid in a scheme name', () => {
    assert.equal(buildProtocolUrl('sp md:', 'https://host/a.md'), 'spmd://host/a.md');
  });

  it('falls back to a usable scheme when the setting is blank', () => {
    assert.match(buildProtocolUrl('', 'https://host/a.md'), /^spmd:\/\//);
  });
});

describe('filterPagesByTag', () => {
  const pages = [
    { path: 'README.md', name: 'README.md', tags: ['設計', 'SPFx'] },
    { path: 'guide/install.md', name: 'install.md', tags: ['SPFx'] },
    { path: 'guide/page.html', name: 'page.html', tags: ['spfx'] },
    { path: 'guide/faq.md', name: 'faq.md', tags: ['FAQ'] },
    { path: 'images/logo.png', name: 'logo.png', tags: ['SPFx'] }
  ];

  it('finds every document carrying the tag, anywhere in the library, in path order', () => {
    assert.deepEqual(
      filterPagesByTag(pages, 'SPFx').map((page) => page.path),
      ['guide/install.md', 'guide/page.html', 'README.md']
    );
  });

  it('ignores case and surrounding space', () => {
    assert.equal(filterPagesByTag(pages, '  spfx ').length, 3);
  });

  it('leaves out files the viewer cannot render', () => {
    assert.equal(
      filterPagesByTag(pages, 'SPFx').filter((page) => page.path === 'images/logo.png').length,
      0
    );
  });

  it('returns nothing for a tag nobody carries, or for no tag at all', () => {
    assert.deepEqual(filterPagesByTag(pages, '未使用'), []);
    assert.deepEqual(filterPagesByTag(pages, ''), []);
    assert.deepEqual(filterPagesByTag(pages, '   '), []);
  });
});

describe('renderTagSearch', () => {
  it('lists the matches with the folder they sit in, ready to navigate', () => {
    const host = dom.window.document.createElement('div');

    renderTagSearch({
      container: host,
      tag: '設計',
      pages: [
        { path: 'README.md', name: 'README.md', tags: ['設計'] },
        { path: 'guide/install.md', name: 'install.md', tags: ['設計'] }
      ],
      title: 'タグ:',
      emptyMessage: 'ありません。',
      linkAttribute: core.DOC_LINK_ATTRIBUTE,
      buildDocUrl: (p) => `/pages/docs.aspx?file=${p}`
    });

    const links = [...host.querySelectorAll('.mdv-related-link')];
    assert.deepEqual(links.map((a) => a.textContent), ['README.md', 'install.md']);
    assert.deepEqual(
      links.map((a) => a.getAttribute(core.DOC_LINK_ATTRIBUTE)),
      ['README.md', 'guide/install.md']
    );
    assert.equal(host.querySelector('.mdv-related-folder').textContent, 'guide/');
    assert.match(host.querySelector('.mdv-search-title').textContent, /タグ:\s*設計/);
  });

  it('says so when nothing carries the tag', () => {
    const host = dom.window.document.createElement('div');

    renderTagSearch({
      container: host,
      tag: '未使用',
      pages: [],
      title: 'タグ:',
      emptyMessage: 'このタグが付いたページはありません。',
      linkAttribute: core.DOC_LINK_ATTRIBUTE,
      buildDocUrl: (p) => p
    });

    assert.equal(host.querySelector('.mdv-related-list'), null);
    assert.match(host.textContent, /このタグが付いたページはありません/);
  });
});
