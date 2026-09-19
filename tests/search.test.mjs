import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, dom } from './setup.mjs';

const {
  commonRoot,
  extractSnippet,
  matchByName,
  pagesInScope,
  renderSearchPage
} = core;

const PAGES = [
  { path: 'README.md', name: 'README.md', tags: [] },
  { path: 'guide/index.md', name: 'index.md', tags: [] },
  { path: 'guide/install.md', name: 'install.md', tags: [] },
  { path: 'guide/api/rest.md', name: 'rest.md', tags: [] },
  { path: 'reference/codes.md', name: 'codes.md', tags: [] }
];

const paths = (pages) => pages.map((p) => p.path);
const hitPaths = (hits) => hits.map((h) => h.page.path);

describe('pagesInScope', () => {
  it('takes the whole library when no folder is given', () => {
    assert.equal(pagesInScope(PAGES, '').length, PAGES.length);
  });

  it('takes everything under the folder, not just its first level', () => {
    // A scope that stopped at the top level would miss guide/api/rest.md, which
    // is where most of a wiki's content actually lives.
    assert.deepEqual(paths(pagesInScope(PAGES, 'guide')), [
      'guide/index.md',
      'guide/install.md',
      'guide/api/rest.md'
    ]);
  });

  it('accepts a folder written with slashes around it', () => {
    assert.deepEqual(paths(pagesInScope(PAGES, '/guide/')), paths(pagesInScope(PAGES, 'guide')));
  });

  it('returns nothing for a folder that holds nothing', () => {
    assert.deepEqual(pagesInScope(PAGES, 'nowhere'), []);
  });
});

describe('matchByName', () => {
  it('matches the file name', () => {
    assert.deepEqual(hitPaths(matchByName(PAGES, 'install')), ['guide/install.md']);
  });

  it('matches the folder, which is how people look for a section', () => {
    assert.deepEqual(hitPaths(matchByName(PAGES, 'guide/')), [
      'guide/index.md',
      'guide/install.md',
      'guide/api/rest.md'
    ]);
  });

  it('ignores case', () => {
    assert.deepEqual(hitPaths(matchByName(PAGES, 'ReadMe')), ['README.md']);
  });

  it('returns nothing for an empty term rather than everything', () => {
    assert.deepEqual(matchByName(PAGES, ''), []);
    assert.deepEqual(matchByName(PAGES, '   '), []);
  });
});

describe('extractSnippet', () => {
  it('returns the text around the first match', () => {
    const snippet = extractSnippet('前置き。ここに設計の話が書いてあります。', '設計');

    assert.match(snippet, /設計/);
  });

  it('collapses newlines and runs of spaces onto one line', () => {
    const snippet = extractSnippet('- 一行目\n\n  - 設計   の話', '設計');

    assert.equal(snippet.indexOf('\n'), -1);
    assert.equal(snippet.indexOf('   '), -1);
  });

  it('marks where it cut the text', () => {
    const long = 'あ'.repeat(200) + '設計' + 'い'.repeat(200);
    const snippet = extractSnippet(long, '設計');

    assert.equal(snippet.charAt(0), '…');
    assert.equal(snippet.charAt(snippet.length - 1), '…');
  });

  it('does not mark an edge it did not cut', () => {
    const snippet = extractSnippet('設計の話', '設計');

    assert.equal(snippet, '設計の話');
  });

  it('returns an empty string when the term is absent, which is the match test', () => {
    assert.equal(extractSnippet('本文', '出てこない'), '');
    assert.equal(extractSnippet('本文', ''), '');
  });

  it('ignores case', () => {
    assert.match(extractSnippet('See the SPFx notes', 'spfx'), /SPFx/);
  });
});

describe('commonRoot', () => {
  it('is the deepest folder every hit sits under', () => {
    assert.equal(
      commonRoot([
        { path: 'guide/api/rest.md', name: 'rest.md', tags: [] },
        { path: 'guide/api/soap.md', name: 'soap.md', tags: [] }
      ]),
      'guide/api'
    );
  });

  it('falls back to the library root when the hits diverge', () => {
    assert.equal(commonRoot(PAGES), '');
  });

  it('is the root when a hit sits there', () => {
    assert.equal(
      commonRoot([
        { path: 'README.md', name: 'README.md', tags: [] },
        { path: 'guide/index.md', name: 'index.md', tags: [] }
      ]),
      ''
    );
  });

  it('copes with no hits at all', () => {
    assert.equal(commonRoot([]), '');
  });
});

describe('renderSearchPage', () => {
  const STRINGS = {
    title: '検索',
    termPlaceholder: '検索する文字列',
    modeName: 'ページ名で探す',
    modeContent: '本文で探す',
    scope: '検索するフォルダー',
    scopePlaceholder: 'フォルダー',
    submit: '検索',
    searching: '検索しています...',
    resultCount: '{0} 件見つかりました。',
    empty: '該当するページはありません。',
    prompt: '検索する文字列を入力してください。',
    restrictedOnly: '限定公開のみ',
    locked: '限定公開'
  };

  function page(request, onSearch = () => {}) {
    const host = dom.window.document.createElement('div');
    const handle = renderSearchPage({
      container: host,
      request,
      strings: STRINGS,
      buildDocUrl: (path) => `/pages/docs.aspx?file=${path}`,
      linkAttribute: 'data-mdv-path',
      onSearch
    });
    return { host, handle };
  }

  it('keeps what was asked for in the form', () => {
    const { host } = page({ term: '設計', mode: 'content', scope: 'guide' });

    assert.equal(host.querySelector('.mdv-search-term').value, '設計');
    assert.equal(host.querySelector('.mdv-search-scope').value, 'guide');
    assert.equal(host.querySelectorAll('input[type="radio"]')[1].checked, true);
  });

  it('asks for a term instead of showing an empty result set', () => {
    const { host } = page({ term: '', mode: 'name', scope: '' });

    assert.equal(host.querySelector('.mdv-related-empty').textContent, STRINGS.prompt);
  });

  it('reports what was typed, including the mode and the folder', () => {
    let asked;
    const { host } = page({ term: '', mode: 'name', scope: '' }, (r) => { asked = r; });

    host.querySelector('.mdv-search-term').value = '  設計  ';
    host.querySelector('.mdv-search-scope').value = ' guide ';
    host.querySelectorAll('input[type="radio"]')[1].checked = true;
    host.querySelector('.mdv-btn-primary').click();

    assert.deepEqual(asked, {
      term: '設計',
      mode: 'content',
      scope: 'guide',
      restrictedOnly: false
    });
  });

  it('searches again when the mode is switched, so the results match the question', () => {
    let calls = 0;
    const { host } = page({ term: '設計', mode: 'name', scope: '' }, () => { calls += 1; });

    const contentMode = host.querySelectorAll('input[type="radio"]')[1];
    contentMode.checked = true;
    contentMode.dispatchEvent(new dom.window.Event('change'));

    assert.equal(calls, 1);
  });

  it('leaves the form alone while it is working', () => {
    const { host, handle } = page({ term: '設計', mode: 'name', scope: 'guide' });

    handle.setBusy('検索しています...');

    // Blanking the form to show progress makes refining a query hostile.
    assert.equal(host.querySelector('.mdv-search-term').value, '設計');
    assert.match(host.querySelector('.mdv-search-results').textContent, /検索しています/);
  });

  it('shows the hits as a tree, with the count above it', () => {
    const { host, handle } = page({ term: 'md', mode: 'name', scope: '' });

    handle.setResults([
      { page: { path: 'guide/index.md', name: 'index.md', tags: [] }, snippet: '' },
      { page: { path: 'guide/api/rest.md', name: 'rest.md', tags: [] }, snippet: '' }
    ]);

    assert.equal(host.querySelector('.mdv-search-count').textContent, '2 件見つかりました。');
    assert.deepEqual(
      [...host.querySelectorAll('.mdv-tree-row')].map((r) => r.textContent),
      ['guide/', '├─ api/', '│  └─ rest.md', '└─ index.md']
    );
  });

  it('lists the matching line under the tree for a body search', () => {
    const { host, handle } = page({ term: '設計', mode: 'content', scope: '' });

    handle.setResults([
      { page: { path: 'guide/index.md', name: 'index.md', tags: [] }, snippet: '…設計の話…' }
    ]);

    const snippet = host.querySelector('.mdv-search-snippet');
    assert.equal(snippet.querySelector('.mdv-related-link').textContent, 'guide/index.md');
    assert.equal(snippet.querySelector('.mdv-search-snippet-text').textContent, '…設計の話…');
  });

  it('says so when nothing matched', () => {
    const { host, handle } = page({ term: 'かすりもしない', mode: 'name', scope: '' });

    handle.setResults([]);

    assert.equal(host.querySelector('.mdv-search-count').textContent, '0 件見つかりました。');
    assert.equal(host.querySelector('.mdv-related-empty').textContent, STRINGS.empty);
  });
});

describe('onlyRestricted', () => {
  const MIXED = [
    { path: 'README.md', name: 'README.md', tags: [] },
    { path: '人事/評価.md', name: '評価.md', tags: [], hasUniquePermissions: true },
    { path: 'guide/index.md', name: 'index.md', tags: [], hasUniquePermissions: false },
    { path: '人事/給与/表.md', name: '表.md', tags: [], hasUniquePermissions: true }
  ];

  it('keeps only the pages holding permissions of their own', () => {
    assert.deepEqual(
      core.onlyRestricted(MIXED).map((p) => p.path),
      ['人事/評価.md', '人事/給与/表.md']
    );
  });

  it('treats an unknown flag as not restricted', () => {
    // The flag is absent for a listing fetched before it was selected; guessing
    // "restricted" there would hide pages that are perfectly open.
    assert.deepEqual(core.onlyRestricted([{ path: 'a.md', name: 'a.md', tags: [] }]), []);
  });
});
