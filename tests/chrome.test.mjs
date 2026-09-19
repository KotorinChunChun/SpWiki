import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, dom } from './setup.mjs';

const { ACTION_ATTRIBUTE, renderNotice, renderTagBar, renderTagEditor, renderToolbar } = core;

const STRINGS = {
  title: 'タグを編集',
  save: '保存',
  cancel: 'キャンセル',
  textHint: '複数のタグは ; か , で区切ってください。',
  noChoices: 'この列に選択肢が登録されていません。',
  addLabel: '新しいタグを追加',
  addPlaceholder: '例: 設計; 手順書',
  addDisabledPlaceholder: 'この列にはタグを追加できません',
  addHint: '一覧に無いタグはここに入力します。',
  addDisabledHint: 'この列は選択肢以外の値を受け付けません。'
};

function container() {
  return dom.window.document.createElement('div');
}

describe('renderToolbar', () => {
  it('builds the two rows the layout rules hook onto', () => {
    const host = container();
    renderToolbar({ container: host, path: 'guide/install.md', buttons: [] });

    assert.equal(host.children.length, 2);
    assert.equal(host.children[0].className, 'mdv-toolbar');
    assert.equal(host.children[1].className, 'mdv-toolbar-file');
  });

  it('shows the title on top and the file name with its folder underneath', () => {
    const host = container();
    renderToolbar({
      container: host,
      path: 'guide/install.md',
      title: 'ドキュメント',
      titleLabel: '開始ページに戻る',
      pathLabel: '変更',
      buttons: []
    });

    const title = host.querySelector('.mdv-toolbar-title');
    assert.equal(title.textContent, 'ドキュメント');
    assert.equal(title.getAttribute(ACTION_ATTRIBUTE), 'home');
    assert.equal(title.getAttribute('title'), '開始ページに戻る');

    assert.equal(host.querySelector('.mdv-toolbar-path').textContent, 'install.md');
    assert.equal(host.querySelector('.mdv-toolbar-folder').textContent, 'guide');
  });

  it('opens the rename panel from the file name only', () => {
    const host = container();
    renderToolbar({ container: host, path: 'guide/install.md', pathLabel: '変更', buttons: [] });

    // The folder is a breadcrumb now, so renaming hangs off the file name
    // alone — clicking a folder means "take me there", not "rename this".
    assert.deepEqual(
      [...host.querySelectorAll('.mdv-toolbar-path-action')].map((b) => b.getAttribute(ACTION_ATTRIBUTE)),
      ['path']
    );
  });

  it('makes each folder segment a link to that folder', () => {
    const host = container();
    renderToolbar({
      container: host,
      path: 'guide/api/rest.md',
      buttons: [],
      folderLinkAttribute: 'data-md-path',
      buildFolderUrl: (p) => `/pages/docs.aspx?file=${p}`
    });

    const links = [...host.querySelectorAll('.mdv-toolbar-folder-link')];

    // Each segment goes to itself, not to the whole path: that is what makes
    // the row usable for walking back up.
    assert.deepEqual(links.map((a) => a.textContent), ['guide', 'api']);
    assert.deepEqual(links.map((a) => a.getAttribute('data-md-path')), ['guide', 'guide/api']);
    assert.equal(links[1].getAttribute('href'), '/pages/docs.aspx?file=guide/api');
  });

  it('leaves the folder as plain text when the host offers no folder links', () => {
    const host = container();
    renderToolbar({ container: host, path: 'guide/api/rest.md', buttons: [] });

    assert.equal(host.querySelector('.mdv-toolbar-folder-link'), null);
    assert.equal(host.querySelector('.mdv-toolbar-folder').textContent, 'guide/api');
  });

  it('puts the outline button in front of everything else on the file row', () => {
    const host = container();
    renderToolbar({
      container: host,
      path: 'guide/install.md',
      buttons: [],
      outline: { label: '目次', pinned: false },
      permission: { locked: false, label: 'アクセス許可' }
    });

    const row = host.querySelector('.mdv-toolbar-file');
    assert.match(row.children[0].className, /mdv-outline-btn/);
    assert.match(row.children[1].className, /mdv-perm-btn/);
    assert.equal(row.children[0].getAttribute(ACTION_ATTRIBUTE), 'outline');
  });

  it('shows on the outline button whether it is pinned', () => {
    const pinned = container();
    renderToolbar({
      container: pinned,
      path: 'a.md',
      buttons: [],
      outline: { label: '目次', pinned: true }
    });

    const button = pinned.querySelector('.mdv-outline-btn');
    assert.match(button.className, /mdv-outline-btn-pinned/);
    assert.equal(button.getAttribute('aria-pressed'), 'true');
  });

  it('names the library root in place of the folder', () => {
    const host = container();
    renderToolbar({
      container: host,
      path: 'README.md',
      rootLabel: '(ライブラリのルート)',
      buttons: []
    });

    assert.equal(host.querySelector('.mdv-toolbar-path').textContent, 'README.md');
    assert.equal(host.querySelector('.mdv-toolbar-folder').textContent, '(ライブラリのルート)');
  });

  it('leaves the path as plain text when the host offers no rename', () => {
    const host = container();
    renderToolbar({ container: host, path: 'README.md', buttons: [] });

    assert.equal(host.querySelector('.mdv-toolbar-path').nodeName.toLowerCase(), 'span');
  });

  it('tags each action button with its action', () => {
    const host = container();
    renderToolbar({
      container: host,
      path: 'guide/install.md',
      buttons: [
        { action: 'edit', label: '編集' },
        { action: 'open-app', label: 'アプリで開く', disabled: true }
      ]
    });

    const buttons = [...host.querySelectorAll('.mdv-toolbar-actions .mdv-btn')];
    assert.deepEqual(buttons.map((b) => b.getAttribute(ACTION_ATTRIBUTE)), ['edit', 'open-app']);
    assert.deepEqual(buttons.map((b) => b.hasAttribute('disabled')), [false, true]);
  });

  it('replaces the previous toolbar instead of stacking a second one', () => {
    const host = container();
    const view = { container: host, path: 'a.md', buttons: [{ action: 'edit', label: '編集' }] };

    renderToolbar(view);
    renderToolbar({ ...view, path: 'b/c.md' });

    assert.equal(host.querySelectorAll('.mdv-toolbar').length, 1);
    assert.equal(host.querySelector('.mdv-toolbar-path').textContent, 'c.md');
    assert.equal(host.querySelector('.mdv-toolbar-folder').textContent, 'b');
  });
});

describe('renderTagBar', () => {
  const bar = (overrides) => {
    const host = container();
    const opened = [];
    const edited = [];

    renderTagBar({
      container: host,
      groups: [
        { field: { internalName: 'DocTags', title: 'タグ', kind: 'multiChoice', choices: [], allowFillIn: true }, tags: ['設計'] },
        { field: { internalName: 'DocOwner', title: '担当', kind: 'text', choices: [], allowFillIn: true }, tags: [] }
      ],
      buildTagUrl: (tag) => `/pages/docs.aspx?tag=${tag}`,
      onTag: (tag) => opened.push(tag),
      onEdit: (internalName) => { edited.push(internalName); },
      editLabel: 'クリックして編集',
      ...overrides
    });

    return { host, opened, edited };
  };

  it('gives every configured column a row, even one with no tags on it', () => {
    const { host } = bar();

    assert.deepEqual(
      [...host.querySelectorAll('.mdv-tag-row-label')].map((l) => l.textContent),
      ['タグ', '担当']
    );
  });

  it('opens the editor for the column whose name was clicked', () => {
    const { host, edited } = bar();

    host.querySelectorAll('.mdv-tag-row-label')[1].click();
    host.querySelectorAll('.mdv-tag-row-label')[0].click();

    // The editor works one column at a time, so which one was clicked is the
    // whole point of the callback.
    assert.deepEqual(edited, ['DocOwner', 'DocTags']);
  });

  it('leaves the column name as plain text when editing is off', () => {
    const { host } = bar({ onEdit: undefined });

    assert.equal(host.querySelector('.mdv-tag-row-label').nodeName.toLowerCase(), 'span');
  });

  it('gives each tag a real link and reports a plain click', () => {
    const { host, opened } = bar();
    const chip = host.querySelector('.mdv-tag-link');

    assert.equal(chip.getAttribute('href'), '/pages/docs.aspx?tag=設計');
    assert.notEqual(chip.style.backgroundColor, '');

    chip.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true, cancelable: true, button: 0 }));
    assert.deepEqual(opened, ['設計']);
  });
});

describe('renderTagEditor', () => {
  const group = (kind, choices, tags, allowFillIn = true) => ({
    field: { internalName: 'Tags', title: 'タグ', kind, choices, allowFillIn },
    tags
  });

  it('offers a text box for a text column and splits what is typed', () => {
    const host = container();
    let saved;

    renderTagEditor({
      container: host,
      groups: [group('text', [], ['設計', 'SPFx'])],
      strings: STRINGS,
      onSave: (groups) => { saved = groups; },
      onCancel: () => {}
    });

    const input = host.querySelector('.mdv-tag-input');
    assert.equal(input.value, '設計; SPFx');

    input.value = ' 運用,  リリース ;; 運用 ';
    host.querySelector('.mdv-btn-primary').click();
    assert.deepEqual(saved.map((g) => g.tags), [['運用', 'リリース']]);
    assert.equal(saved[0].field.internalName, 'Tags');
  });

  it('offers check boxes for a multi choice column and pre-checks the current tags', () => {
    const host = container();
    let saved;

    renderTagEditor({
      container: host,
      groups: [group('multiChoice', ['設計', 'SPFx', '運用'], ['SPFx'])],
      strings: STRINGS,
      onSave: (groups) => { saved = groups; },
      onCancel: () => {}
    });

    const boxes = [...host.querySelectorAll('input[type="checkbox"]')];
    assert.deepEqual(boxes.map((b) => b.checked), [false, true, false]);

    boxes[0].checked = true;
    host.querySelector('.mdv-btn-primary').click();
    assert.deepEqual(saved[0].tags, ['設計', 'SPFx']);
  });

  it('offers radio buttons for a single choice column', () => {
    const host = container();

    renderTagEditor({
      container: host,
      groups: [group('choice', ['設計', '運用'], ['運用'])],
      strings: STRINGS,
      onSave: () => {},
      onCancel: () => {}
    });

    const inputs = [...host.querySelectorAll('input[type="radio"]')];
    assert.deepEqual(inputs.map((b) => b.checked), [false, true]);
  });

  it('explains an empty choice column rather than showing nothing', () => {
    const host = container();

    renderTagEditor({
      container: host,
      groups: [group('choice', [], [])],
      strings: STRINGS,
      onSave: () => {},
      onCancel: () => {}
    });

    assert.match(host.textContent, /選択肢が登録されていません/);
  });

  it('reports a cancel without saving', () => {
    const host = container();
    let cancelled = false;

    renderTagEditor({
      container: host,
      groups: [group('text', [], [])],
      strings: STRINGS,
      onSave: () => { throw new Error('should not save'); },
      onCancel: () => { cancelled = true; }
    });

    [...host.querySelectorAll('.mdv-btn')].filter((b) => b.textContent === 'キャンセル')[0].click();
    assert.equal(cancelled, true);
  });

  it('escapes choice values instead of interpreting them as markup', () => {
    const host = container();

    renderTagEditor({
      container: host,
      groups: [group('multiChoice', ['<b>bold</b>'], [])],
      strings: STRINGS,
      onSave: () => {},
      onCancel: () => {}
    });

    assert.equal(host.querySelector('b'), null);
    assert.match(host.textContent, /<b>bold<\/b>/);
  });

  it('renders one section per column and saves them all at once', () => {
    const host = container();
    let saved;

    renderTagEditor({
      container: host,
      groups: [
        group('multiChoice', ['設計', 'SPFx'], ['設計']),
        group('choice', ['手順書'], []),
        group('text', [], ['山田'])
      ],
      strings: STRINGS,
      onSave: (groups) => { saved = groups; },
      onCancel: () => {}
    });

    assert.equal(host.querySelectorAll('.mdv-tag-section').length, 3);

    host.querySelector('.mdv-btn-primary').click();
    assert.deepEqual(saved.map((g) => g.tags), [['設計'], [], ['山田']]);
  });

  it('adds a tag the column does not offer yet', () => {
    const host = container();
    let saved;

    renderTagEditor({
      container: host,
      groups: [group('multiChoice', ['設計'], [])],
      strings: STRINGS,
      onSave: (groups) => { saved = groups; },
      onCancel: () => {}
    });

    host.querySelector('.mdv-tag-add').value = ' 新タグ ; もう一つ ';
    host.querySelector('.mdv-btn-primary').click();

    assert.deepEqual(saved[0].tags, ['新タグ', 'もう一つ']);
  });

  it('greys the add box out when the column refuses values outside its choices', () => {
    const host = container();
    let saved;

    renderTagEditor({
      container: host,
      groups: [group('multiChoice', ['設計'], [], false)],
      strings: STRINGS,
      onSave: (groups) => { saved = groups; },
      onCancel: () => {}
    });

    const add = host.querySelector('.mdv-tag-add');
    assert.equal(add.hasAttribute('disabled'), true);

    // Disabled has to read as disabled without a hover: its own placeholder
    // says so, and the styling hook is on the element.
    assert.equal(add.getAttribute('placeholder'), STRINGS.addDisabledPlaceholder);
    assert.match(add.className, /mdv-tag-add-disabled/);

    // The reason is a tooltip on the warning icon, not a paragraph of prose.
    // Its own element rather than the `title` attribute, so it appears at once
    // and a keyboard reader can reach it.
    const tip = host.querySelector('.mdv-tip-host');
    assert.equal(tip.getAttribute('aria-label'), STRINGS.addDisabledHint);
    assert.equal(tip.getAttribute('tabindex'), '0');
    assert.equal(tip.querySelector('.mdv-tag-warn').textContent, '!');
    assert.equal(tip.querySelector('.mdv-tip').textContent, STRINGS.addDisabledHint);

    // And it is not also printed under the box.
    assert.equal(host.querySelector('.mdv-tag-hint'), null);

    // Even if something did put a value in it, it must not reach the save.
    add.value = '無理なタグ';
    host.querySelector('.mdv-btn-primary').click();
    assert.deepEqual(saved[0].tags, []);
  });

  it('labels the add box and leaves the warning off when the column takes anything', () => {
    const host = container();

    renderTagEditor({
      container: host,
      groups: [group('multiChoice', ['設計'], [])],
      strings: STRINGS,
      onSave: () => {},
      onCancel: () => {}
    });

    assert.equal(host.querySelector('.mdv-tag-add-label').textContent, STRINGS.addLabel);
    assert.equal(host.querySelector('.mdv-tag-warn'), null);
    assert.equal(host.querySelector('.mdv-tag-add').getAttribute('placeholder'), STRINGS.addPlaceholder);
  });

  it('reports a failed save under the buttons and clears it on the next attempt', () => {
    const host = container();
    let saves = 0;

    const editor = renderTagEditor({
      container: host,
      groups: [group('text', [], ['設計'])],
      strings: STRINGS,
      onSave: () => { saves += 1; },
      onCancel: () => {}
    });

    const error = host.querySelector('.mdv-tag-error');
    assert.equal(error.textContent, '');

    editor.setError('タグを保存できませんでした。 HTTP 400');
    assert.equal(error.textContent, 'タグを保存できませんでした。 HTTP 400');

    // Pressing save again must not leave the previous verdict on screen.
    host.querySelector('.mdv-btn-primary').click();
    assert.equal(error.textContent, '');
    assert.equal(saves, 1);
  });

  it('clears the failure when the editor is cancelled', () => {
    const host = container();
    let cancelled = false;

    const editor = renderTagEditor({
      container: host,
      groups: [group('text', [], [])],
      strings: STRINGS,
      onSave: () => {},
      onCancel: () => { cancelled = true; }
    });

    editor.setError('タグを保存できませんでした。 HTTP 400');
    [...host.querySelectorAll('.mdv-btn')].filter((b) => b.textContent === STRINGS.cancel)[0].click();

    assert.equal(host.querySelector('.mdv-tag-error').textContent, '');
    assert.equal(cancelled, true);
  });

  it('leaves the add box usable for a text column', () => {
    const host = container();

    renderTagEditor({
      container: host,
      groups: [group('text', [], [], false)],
      strings: STRINGS,
      onSave: () => {},
      onCancel: () => {}
    });

    // A text column has the main box instead, and takes anything.
    assert.equal(host.querySelector('.mdv-tag-add'), null);
    assert.equal(host.querySelector('.mdv-tag-input').hasAttribute('disabled'), false);
  });

  it('keeps a tag the column does not know, instead of dropping it on save', () => {
    const host = container();
    let saved;

    renderTagEditor({
      container: host,
      groups: [group('multiChoice', ['設計'], ['設計', '手書きのタグ'])],
      strings: STRINGS,
      onSave: (groups) => { saved = groups; },
      onCancel: () => {}
    });

    const boxes = [...host.querySelectorAll('input[type="checkbox"]')];
    assert.deepEqual(boxes.map((b) => b.value), ['設計', '手書きのタグ']);
    assert.deepEqual(boxes.map((b) => b.checked), [true, true]);

    host.querySelector('.mdv-btn-primary').click();
    assert.deepEqual(saved[0].tags, ['設計', '手書きのタグ']);
  });
});

describe('renderNotice', () => {
  it('shows the message, the detail and a copy button', () => {
    const host = container();
    let copied = false;

    renderNotice({
      container: host,
      message: 'アプリを起動できませんでした。',
      detail: 'C:\\docs\\guide\\install.md',
      copyLabel: 'パスをコピー',
      onCopy: () => { copied = true; },
      closeLabel: '閉じる',
      onClose: () => {}
    });

    assert.match(host.querySelector('.mdv-notice-message').textContent, /起動できませんでした/);
    assert.equal(host.querySelector('.mdv-notice-detail').textContent, 'C:\\docs\\guide\\install.md');

    [...host.querySelectorAll('.mdv-btn')].filter((b) => b.textContent === 'パスをコピー')[0].click();
    assert.equal(copied, true);
  });

  it('omits the copy button when there is nothing to copy', () => {
    const host = container();

    renderNotice({
      container: host,
      message: 'タグを保存しました。',
      closeLabel: '閉じる',
      onClose: () => {}
    });

    assert.equal(host.querySelectorAll('.mdv-btn').length, 1);
    assert.equal(host.querySelector('.mdv-notice-detail'), null);
  });

  it('clears itself when closed', () => {
    const host = container();

    renderNotice({
      container: host,
      message: 'x',
      closeLabel: '閉じる',
      onClose: () => { host.innerHTML = ''; }
    });

    host.querySelector('.mdv-btn').click();
    assert.equal(host.innerHTML, '');
  });
});
