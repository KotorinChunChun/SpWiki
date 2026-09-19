import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, dom } from './setup.mjs';

const { buildAncestors, joinFolderAndName, renderNavPanel, tagColor, withExtension } = core;

const NAV_STRINGS = {
  newPage: '新しいページを作成',
  newPageName: 'ファイル名',
  format: '形式',
  openWith: '作成後',
  openInSharePoint: 'SharePoint で編集',
  openInApp: 'アプリで編集',
  create: '作成',
  cancel: 'キャンセル',
  location: '現在の場所',
  root: '(ライブラリのルート)',
  empty: 'ファイルはありません。',
  close: '閉じる'
};

function navPanel(overrides) {
  const host = dom.window.document.createElement('div');
  const created = [];

  renderNavPanel({
    container: host,
    ancestors: [{ name: 'guide', path: 'guide' }],
    folders: [],
    files: [{ name: 'install.md', path: 'guide/install.md' }],
    currentPath: 'guide/install.md',
    strings: NAV_STRINGS,
    buildDocUrl: (p) => `/pages/docs.aspx?file=${p}`,
    onOpen: () => {},
    onFolder: () => {},
    onCreate: (request) => created.push(request),
    onClose: () => {},
    ...overrides
  });

  const button = (text) => [...host.querySelectorAll('button')].filter((b) => b.textContent === text)[0];
  return { host, created, button };
}

describe('new page form', () => {
  it('asks for nothing until the writer says they want a page', () => {
    const { host } = navPanel();

    assert.match(host.querySelector('.mdv-nav-form').className, /mdv-hidden/);
  });

  it('shows the name, the format and where to edit once opened', () => {
    const { host, button } = navPanel();

    button('新しいページを作成').click();

    const form = host.querySelector('.mdv-nav-form');
    assert.equal(form.className.indexOf('mdv-hidden'), -1);
    assert.notEqual(form.querySelector('.mdv-nav-input'), null);
    assert.deepEqual(
      [...form.querySelectorAll('input[name="mdv-new-format"]')].map((i) => i.value),
      ['markdown', 'html']
    );
    assert.deepEqual(
      [...form.querySelectorAll('input[name="mdv-new-editor"]')].map((i) => i.value),
      ['sharepoint', 'app']
    );
  });

  it('reports the filled in form, defaulting to markdown in SharePoint', () => {
    const { host, created, button } = navPanel();

    button('新しいページを作成').click();
    host.querySelector('.mdv-nav-input').value = '手順';
    button('作成').click();

    assert.deepEqual(created, [{ name: '手順', format: 'markdown', editor: 'sharepoint' }]);
  });

  it('reports html and the desktop app when they are picked', () => {
    const { host, created, button } = navPanel();

    button('新しいページを作成').click();
    host.querySelector('.mdv-nav-input').value = '手順';
    host.querySelectorAll('input[name="mdv-new-format"]')[1].checked = true;
    host.querySelectorAll('input[name="mdv-new-editor"]')[1].checked = true;
    button('作成').click();

    assert.deepEqual(created, [{ name: '手順', format: 'html', editor: 'app' }]);
  });

  it('folds the form away again on cancel, without creating anything', () => {
    const { host, created, button } = navPanel();

    button('新しいページを作成').click();
    button('キャンセル').click();

    assert.match(host.querySelector('.mdv-nav-form').className, /mdv-hidden/);
    assert.deepEqual(created, []);
  });
});

describe('buildAncestors', () => {
  it('lists every folder down to the current one', () => {
    assert.deepEqual(buildAncestors('guide/api'), [
      { name: 'guide', path: 'guide' },
      { name: 'api', path: 'guide/api' }
    ]);
  });

  it('returns nothing for the library root', () => {
    assert.deepEqual(buildAncestors(''), []);
    assert.deepEqual(buildAncestors('/'), []);
  });

  it('normalizes the path it is given', () => {
    assert.deepEqual(buildAncestors('guide//./api/'), [
      { name: 'guide', path: 'guide' },
      { name: 'api', path: 'guide/api' }
    ]);
  });
});

describe('withExtension', () => {
  it('adds the extension of the chosen format', () => {
    assert.equal(withExtension('手順', 'markdown'), '手順.md');
    assert.equal(withExtension('手順', 'html'), '手順.html');
  });

  it('keeps an extension the viewer already renders', () => {
    assert.equal(withExtension('README.md', 'html'), 'README.md');
    assert.equal(withExtension('page.htm', 'markdown'), 'page.htm');
  });

  it('treats an unknown extension as part of the name', () => {
    assert.equal(withExtension('notes.txt', 'markdown'), 'notes.txt.md');
  });

  it('trims and refuses names that are not a single file name', () => {
    assert.equal(withExtension('  手順  ', 'markdown'), '手順.md');
    assert.equal(withExtension('', 'markdown'), '');
    assert.equal(withExtension('   ', 'markdown'), '');
    assert.equal(withExtension('guide/手順', 'markdown'), '');
    assert.equal(withExtension('guide\\手順', 'markdown'), '');
  });
});

describe('joinFolderAndName', () => {
  it('joins the two boxes of the rename panel', () => {
    assert.equal(joinFolderAndName('guide', 'install.md'), 'guide/install.md');
  });

  it('moves the file to the root when the folder is empty', () => {
    assert.equal(joinFolderAndName('', 'install.md'), 'install.md');
    assert.equal(joinFolderAndName('  ', 'install.md'), 'install.md');
  });

  it('tolerates stray slashes and back slashes', () => {
    assert.equal(joinFolderAndName('/guide/', '/install.md'), 'guide/install.md');
    assert.equal(joinFolderAndName('guide\\api', 'install.md'), 'guide/api/install.md');
  });

  it('returns an empty path when there is no file name', () => {
    assert.equal(joinFolderAndName('guide', ''), '');
  });
});

describe('tagColor', () => {
  it('gives the same tag the same colour every time', () => {
    assert.deepEqual(tagColor('設計'), tagColor('設計'));
  });

  it('ignores case and surrounding space, like tag comparison does', () => {
    assert.deepEqual(tagColor('SPFx'), tagColor('  spfx '));
  });

  it('gives different tags different hues', () => {
    const colours = ['設計', 'SPFx', '運用', 'FAQ', 'リリース'].map((tag) => tagColor(tag).background);

    assert.equal(new Set(colours).size, colours.length);
  });

  it('stays pale, so the chips never fight with the text around them', () => {
    for (const tag of ['設計', 'SPFx', '運用', 'FAQ', 'リリース', 'a', 'zzz']) {
      const lightness = Number(/,\s*([\d.]+)%\)$/.exec(tagColor(tag).background)[1]);
      assert.ok(lightness >= 75, `${tag} -> ${lightness}%`);
    }
  });

  it('picks a foreground that reads on that background', () => {
    for (const tag of ['設計', 'SPFx', '運用', 'FAQ', 'リリース']) {
      const colour = tagColor(tag);
      assert.ok(colour.foreground === '#1b1b1b' || colour.foreground === '#ffffff');
    }
  });
});
