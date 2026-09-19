import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, dom } from './setup.mjs';

const { findDirectives, parseDirective, renderDirectives, resolveFolder, selectPages } = core;

const PAGES = [
  { path: 'README.md', name: 'README.md', tags: ['設計'] },
  { path: 'guide/index.md', name: 'index.md', tags: [] },
  { path: 'guide/install.md', name: 'install.md', tags: ['設計', '手順書'] },
  { path: 'guide/faq.md', name: 'faq.md', tags: ['FAQ'] },
  { path: 'guide/api/rest.md', name: 'rest.md', tags: ['設計'] },
  { path: 'reference/codes.md', name: 'codes.md', tags: [] }
];

function context(currentPath, overrides = {}) {
  return {
    pages: PAGES,
    currentPath,
    buildDocUrl: (path) => `/pages/docs.aspx?file=${path}`,
    linkAttribute: 'data-mdv-path',
    strings: { empty: '該当するページはありません。', rootLabel: '(ライブラリのルート)' },
    ...overrides
  };
}

function render(html, currentPath = 'guide/install.md', overrides = {}) {
  const host = dom.window.document.createElement('div');
  host.innerHTML = html;
  renderDirectives(host, findDirectives(host), context(currentPath, overrides));
  return host;
}

const linkText = (host) =>
  [...host.querySelectorAll('.mdv-related-link')].map((a) => a.textContent);

describe('parseDirective', () => {
  const heading = (text) => {
    const h = dom.window.document.createElement('h2');
    h.textContent = text;
    return h;
  };

  it('reads each marker with its argument', () => {
    assert.deepEqual(
      ['タグページ一覧:設計', 'サブページ一覧:guide', '関連ページ一覧:install'].map((t) => {
        const d = parseDirective(heading(t));
        return [d.kind, d.argument];
      }),
      [['tagPages', '設計'], ['subPages', 'guide'], ['relatedPages', 'install']]
    );
  });

  it('reads the marker that stands alone', () => {
    assert.equal(parseDirective(heading('ページツリー')).kind, 'pageTree');
  });

  it('accepts a full width colon and surrounding space', () => {
    const d = parseDirective(heading(' タグページ一覧：　設計 '));
    assert.deepEqual([d.kind, d.argument], ['tagPages', '設計']);
  });

  it('ignores a heading that only mentions a marker', () => {
    // Prose about the feature is not a request for it.
    assert.equal(parseDirective(heading('サブページ一覧について')), undefined);
    assert.equal(parseDirective(heading('この章ではページツリーを説明します')), undefined);
    assert.equal(parseDirective(heading('概要')), undefined);
  });

  it('ignores a marker whose argument is missing', () => {
    assert.equal(parseDirective(heading('タグページ一覧:')), undefined);
    assert.equal(parseDirective(heading('タグページ一覧: 　')), undefined);
  });
});

describe('resolveFolder', () => {
  it('reads a leading slash as the library root', () => {
    assert.equal(resolveFolder('/guide', 'reference/codes.md'), 'guide');
    assert.equal(resolveFolder('/guide/api/', 'reference/codes.md'), 'guide/api');
  });

  it('reads anything else as relative to the document', () => {
    assert.equal(resolveFolder('api', 'guide/install.md'), 'guide/api');
    assert.equal(resolveFolder('../reference', 'guide/install.md'), 'reference');
    assert.equal(resolveFolder('.', 'guide/install.md'), 'guide');
  });

  it('treats a document in the root as being in the root', () => {
    assert.equal(resolveFolder('guide', 'README.md'), 'guide');
  });
});

describe('selectPages', () => {
  const select = (kind, argument, currentPath = 'guide/install.md') =>
    selectPages({ kind, argument, heading: null }, context(currentPath)).map((p) => p.path);

  it('lists every page carrying a tag, in path order', () => {
    assert.deepEqual(select('tagPages', '設計'), [
      'guide/api/rest.md',
      'guide/install.md',
      'README.md'
    ]);
  });

  it('ignores case and space when matching a tag', () => {
    assert.deepEqual(select('tagPages', ' faq '), ['guide/faq.md']);
  });

  it('lists a folder and everything below it', () => {
    assert.deepEqual(select('subPages', '/guide'), [
      'guide/api/rest.md',
      'guide/faq.md',
      'guide/index.md',
      'guide/install.md'
    ]);
  });

  it('resolves a relative folder against the page it is written on', () => {
    assert.deepEqual(select('subPages', 'api', 'guide/install.md'), ['guide/api/rest.md']);
    assert.deepEqual(select('subPages', '../reference', 'guide/install.md'), [
      'reference/codes.md'
    ]);
  });

  it('matches a keyword against the path and the tags', () => {
    assert.deepEqual(select('relatedPages', 'install'), ['guide/install.md']);
    assert.deepEqual(select('relatedPages', '手順書'), ['guide/install.md']);
    assert.deepEqual(select('relatedPages', 'REST'), ['guide/api/rest.md']);
  });

  it('returns nothing rather than everything for a keyword nobody wrote', () => {
    assert.deepEqual(select('relatedPages', 'かすりもしない'), []);
  });
});

describe('renderDirectives', () => {
  it('inserts the list right after the heading', () => {
    const host = render('<h2>タグページ一覧:FAQ</h2><p>あと</p>');
    const block = host.querySelector('h2').nextElementSibling;

    assert.match(block.className, /mdv-directive/);
    assert.deepEqual(linkText(host), ['faq.md']);
    assert.equal(host.querySelector('p').textContent, 'あと');
  });

  it('gives every entry a real href and the attribute the host navigates on', () => {
    const host = render('<h2>タグページ一覧:FAQ</h2>');
    const link = host.querySelector('.mdv-related-link');

    assert.equal(link.getAttribute('href'), '/pages/docs.aspx?file=guide/faq.md');
    assert.equal(link.getAttribute('data-mdv-path'), 'guide/faq.md');
  });

  it('says so instead of leaving a blank gap when nothing matches', () => {
    const host = render('<h2>タグページ一覧:存在しないタグ</h2>');

    assert.equal(host.querySelector('.mdv-related-empty').textContent, '該当するページはありません。');
    assert.equal(host.querySelector('.mdv-related-link'), null);
  });

  it('fills in several directives in one document', () => {
    const host = render('<h2>タグページ一覧:FAQ</h2><h2>サブページ一覧:../reference</h2>');

    assert.deepEqual(linkText(host), ['faq.md', 'codes.md']);
  });

  it('builds the tree down to the folder holding the page', () => {
    const host = render('<h2>ページツリー</h2>', 'guide/install.md');

    // The root is a label; the folders on the way down are links, so the route
    // can be walked back up.
    assert.deepEqual(
      [...host.querySelectorAll('.mdv-tree-label')].map((s) => s.textContent),
      ['(ライブラリのルート)']
    );
    assert.deepEqual(
      [...host.querySelectorAll('.mdv-tree-folder-link')].map((a) => a.textContent),
      ['guide/']
    );
    // The folder's own documents, not the ones in guide/api below it.
    assert.deepEqual(linkText(host), ['faq.md', 'index.md', 'install.md']);
  });

  it('draws the rules so a row shows which folder it belongs to', () => {
    const host = render('<h2>ページツリー</h2>', 'guide/install.md');
    const prefixes = [...host.querySelectorAll('.mdv-tree-prefix')].map((s) => s.textContent);

    // Root, the folder under it, then its three documents; the last one closes.
    assert.deepEqual(prefixes, ['', '└─ ', '   ├─ ', '   ├─ ', '   └─ ']);
  });

  it('names the library root when the page sits there', () => {
    const host = render('<h2>ページツリー</h2>', 'README.md');

    assert.deepEqual(
      [...host.querySelectorAll('.mdv-tree-label')].map((s) => s.textContent),
      ['(ライブラリのルート)']
    );
    assert.deepEqual(linkText(host), ['README.md']);
  });

  it('draws サブページ一覧 as a tree rather than a flat run of paths', () => {
    const host = render('<h2>サブページ一覧:/guide</h2>', 'README.md');

    // Folders before files at each level, each file indented under its folder.
    assert.deepEqual(
      [...host.querySelectorAll('.mdv-tree-row')].map((r) => r.textContent),
      [
        'guide/',
        '├─ api/',
        '│  └─ rest.md',
        '├─ faq.md',
        '├─ index.md',
        '└─ install.md'
      ]
    );
  });

  it('links every folder in a サブページ一覧 so they open as pages', () => {
    const host = render('<h2>サブページ一覧:/guide</h2>', 'README.md');
    const folders = [...host.querySelectorAll('.mdv-tree-folder-link')];

    // The root of the listing is a link too, so it can be opened on its own.
    assert.deepEqual(folders.map((a) => a.textContent), ['guide/', 'api/']);
    assert.equal(folders[1].getAttribute('data-mdv-path'), 'guide/api');
    assert.equal(folders[1].getAttribute('href'), '/pages/docs.aspx?file=guide/api');
  });

  it('marks the page being read inside the list', () => {
    const host = render('<h2>ページツリー</h2>', 'guide/install.md');
    const current = host.querySelector('.mdv-directive-current .mdv-related-link');

    assert.equal(current.textContent, 'install.md');
  });

  it('replaces the previous block instead of stacking a second one', () => {
    const host = render('<h2>タグページ一覧:FAQ</h2>');
    renderDirectives(host, findDirectives(host), context('guide/install.md'));

    assert.equal(host.querySelectorAll('.mdv-directive').length, 1);
  });

  it('leaves a document without a marker completely alone', () => {
    const host = render('<h2>概要</h2><p>本文</p>');

    assert.equal(host.querySelector('.mdv-directive'), null);
    assert.equal(host.innerHTML, '<h2>概要</h2><p>本文</p>');
  });
});

describe('page tree padlocks', () => {
  const { buildTree, renderTree } = core;

  const locked = (host) =>
    [...host.querySelectorAll('.mdv-tree-row')]
      .filter((r) => r.querySelector('.mdv-tree-lock'))
      .map((r) => r.textContent);

  function tree(pages, restrictedFolders) {
    const host = dom.window.document.createElement('div');
    host.appendChild(
      renderTree(dom.window.document, buildTree(pages, ''), {
        buildDocUrl: (p) => `/pages/docs.aspx?file=${p}`,
        linkAttribute: 'data-mdv-path',
        currentPath: '',
        linkFolders: true,
        restrictedFolders,
        lockedLabel: '限定公開'
      })
    );
    return host;
  }

  it('marks the files that carry permissions of their own', () => {
    const host = tree(
      [
        { path: 'README.md', name: 'README.md', tags: [] },
        { path: '人事/評価.md', name: '評価.md', tags: [], hasUniquePermissions: true }
      ],
      []
    );

    assert.deepEqual(locked(host), ['│  └─ 評価.md']);
  });

  it('marks a restricted folder, which is the unit permissions should be set on', () => {
    const host = tree(
      [{ path: '人事/評価.md', name: '評価.md', tags: [] }],
      ['人事']
    );

    assert.deepEqual(locked(host), ['└─ 人事/']);
  });

  it('matches a folder regardless of how its path was cased', () => {
    const host = tree([{ path: '案件/a.md', name: 'a.md', tags: [] }], ['案件']);

    assert.equal(locked(host).length, 1);
  });

  it('leaves an open library completely unmarked', () => {
    const host = tree([{ path: 'guide/index.md', name: 'index.md', tags: [] }], []);

    assert.equal(host.querySelector('.mdv-tree-lock'), null);
  });

  it('gives the padlock a name, since its meaning is not in the text', () => {
    const host = tree([{ path: 'a.md', name: 'a.md', tags: [], hasUniquePermissions: true }], []);
    const mark = host.querySelector('.mdv-tree-lock');

    assert.equal(mark.getAttribute('aria-label'), '限定公開');
    assert.equal(mark.getAttribute('title'), '限定公開');
  });
});
