import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { core, dom } from './setup.mjs';

const { grantableRoles, renderPermissionPanel, visibleAssignments } = core;

const READ = { id: 1073741826, name: '閲覧' };
const EDIT = { id: 1073741830, name: '編集' };
const LIMITED = { id: 1073741825, name: '制限付きアクセス' };
const FULL = { id: 1073741829, name: 'フル コントロール' };

const group = (title) => ({
  id: 7,
  loginName: `c:0t.c|tenant|${title}`,
  title,
  kind: 'securityGroup',
  email: ''
});

const STRINGS = {
  title: 'アクセス許可',
  inheritedFrom: '{0}のアクセス許可を継承',
  unique: '固有のアクセス許可',
  team: 'チーム',
  folderSuffix: ' フォルダー',
  whoHasAccess: 'アクセスできる相手',
  restrict: '限定公開にする',
  restore: '親の設定に戻す',
  confirmTitle: '確認してください',
  confirmRestrict: '「{0}」を親フォルダーから切り離します。',
  confirmRestore: '「{0}」を親フォルダーの設定に戻します。',
  confirmYes: '実行する',
  confirmNo: 'キャンセル',
  addLabel: 'アクセスできる相手を追加',
  addPlaceholder: 'グループ名',
  search: '検索',
  searchFailed: '相手を検索できませんでした。',
  add: '追加',
  remove: '削除',
  searching: '検索しています...',
  noResults: '見つかりませんでした。',
  brokenTitle: '個別設定されている配下の項目',
  brokenIntro: 'ここでの変更は届きません。',
  fileScopeWarning: 'ファイル単位の設定は重くなります。',
  loading: '読み込んでいます...',
  close: '閉じる'
};

function panel(overrides = {}, calls = {}) {
  const host = dom.window.document.createElement('div');
  const handle = renderPermissionPanel({
    container: host,
    state: {
      path: '人事',
      isFolder: true,
      hasUnique: false,
      assignments: [],
      brokenDescendants: [],
      ...(overrides.state || {})
    },
    roles: overrides.roles || [READ, EDIT],
    strings: STRINGS,
    searchPrincipals: overrides.searchPrincipals || (() => Promise.resolve([])),
    onRestrict: calls.onRestrict || (() => {}),
    onRestore: calls.onRestore || (() => {}),
    onAdd: calls.onAdd || (() => {}),
    onRemove: calls.onRemove || (() => {}),
    onClose: calls.onClose || (() => {})
  });
  return { host, handle };
}

const buttonNamed = (host, text) =>
  [...host.querySelectorAll('button')].filter((b) => b.textContent === text)[0];

describe('visibleAssignments', () => {
  it('drops the bookkeeping SharePoint adds for traversal', () => {
    // "Limited Access" grants nothing; showing it makes the list unreadable.
    const kept = visibleAssignments([
      { principal: group('人事部'), roles: [READ] },
      { principal: group('通りすがり'), roles: [LIMITED] }
    ]);

    assert.deepEqual(kept.map((a) => a.principal.title), ['人事部']);
  });

  it('drops an assignment with no roles left on it', () => {
    assert.deepEqual(visibleAssignments([{ principal: group('空'), roles: [] }]), []);
  });

  it('keeps a principal that holds limited access alongside a real role', () => {
    const kept = visibleAssignments([{ principal: group('人事部'), roles: [LIMITED, READ] }]);

    assert.equal(kept.length, 1);
  });
});

describe('grantableRoles', () => {
  it('offers the ordinary levels', () => {
    assert.deepEqual(grantableRoles([READ, EDIT]).map((r) => r.name), ['閲覧', '編集']);
  });

  it('never offers full control or limited access', () => {
    // Limited access is assigned by SharePoint, and a wiki page should not be
    // handing out full control.
    assert.deepEqual(
      grantableRoles([LIMITED, READ, FULL, EDIT]).map((r) => r.name),
      ['閲覧', '編集']
    );
  });
});

describe('renderPermissionPanel', () => {
  it('names the parent it is inheriting from, rather than just saying it does', () => {
    const { host } = panel({ state: { path: '人事/評価.md' } });

    assert.equal(host.querySelector('.mdv-perm-badge').textContent, '人事 フォルダーのアクセス許可を継承');
    assert.equal(buttonNamed(host, STRINGS.restrict) !== undefined, true);
    assert.equal(buttonNamed(host, STRINGS.restore), undefined);
  });

  it('calls the library root the team, since there is no folder to name', () => {
    const { host } = panel({ state: { path: 'README.md' } });

    assert.equal(host.querySelector('.mdv-perm-badge').textContent, 'チームのアクセス許可を継承');
  });

  it('offers the way back once the item has its own permissions', () => {
    const { host } = panel({ state: { hasUnique: true } });

    assert.equal(host.querySelector('.mdv-perm-badge').textContent, STRINGS.unique);
    assert.equal(buttonNamed(host, STRINGS.restore) !== undefined, true);
    assert.equal(buttonNamed(host, STRINGS.restrict), undefined);
  });

  it('does nothing until the confirmation is accepted', () => {
    let restricted = 0;
    const { host } = panel({}, { onRestrict: () => { restricted += 1; } });

    buttonNamed(host, STRINGS.restrict).click();
    assert.equal(restricted, 0, 'pressing the button must only ask');

    const confirm = host.querySelector('.mdv-perm-confirm');
    assert.match(confirm.textContent, /人事/, 'the confirmation names the target');

    buttonNamed(host, STRINGS.confirmYes).click();
    assert.equal(restricted, 1);
  });

  it('lets the confirmation be backed out of', () => {
    let restricted = 0;
    const { host } = panel({}, { onRestrict: () => { restricted += 1; } });

    buttonNamed(host, STRINGS.restrict).click();
    buttonNamed(host, STRINGS.confirmNo).click();

    assert.equal(restricted, 0);
    assert.equal(host.querySelector('.mdv-perm-confirm'), null);
    // And the panel is usable again.
    assert.equal(buttonNamed(host, STRINGS.restrict).disabled, false);
  });

  it('confirms a restore separately, saying what will be lost', () => {
    let restored = 0;
    const { host } = panel(
      { state: { hasUnique: true } },
      { onRestore: () => { restored += 1; } }
    );

    buttonNamed(host, STRINGS.restore).click();
    assert.match(host.querySelector('.mdv-perm-confirm-body').textContent, /戻します/);

    buttonNamed(host, STRINGS.confirmYes).click();
    assert.equal(restored, 1);
  });

  it('warns about the descendants a change here will not reach', () => {
    const { host } = panel({
      state: { hasUnique: true, brokenDescendants: ['人事/公開資料', '人事/案内.md'] }
    });

    const broken = host.querySelector('.mdv-perm-broken');
    assert.match(broken.textContent, /届きません/);
    assert.deepEqual(
      [...broken.querySelectorAll('.mdv-perm-row')].map((r) => r.textContent),
      ['人事/公開資料', '人事/案内.md']
    );
  });

  it('leaves the warning out when there is nothing below to warn about', () => {
    const { host } = panel({ state: { hasUnique: true } });

    assert.equal(host.querySelector('.mdv-perm-broken'), null);
  });

  it('warns about per-file permissions only when one is about to be created', () => {
    const { host } = panel({ state: { isFolder: false, path: '人事/評価.md' } });

    // Not on every visit: nobody is making that decision just by looking.
    assert.doesNotMatch(host.textContent, /ファイル単位の設定は重くなります/);

    buttonNamed(host, STRINGS.restrict).click();
    assert.match(host.querySelector('.mdv-perm-confirm').textContent, /ファイル単位の設定は重くなります/);
  });

  it('leaves that warning out of a folder confirmation', () => {
    const { host } = panel({ state: { isFolder: true } });

    buttonNamed(host, STRINGS.restrict).click();
    assert.doesNotMatch(host.querySelector('.mdv-perm-confirm').textContent, /ファイル単位/);
  });

  it('only offers to remove someone once the permissions are its own', () => {
    const inherited = panel({ state: { assignments: [{ principal: group('人事部'), roles: [READ] }] } });
    assert.equal(buttonNamed(inherited.host, STRINGS.remove), undefined);

    const unique = panel({
      state: { hasUnique: true, assignments: [{ principal: group('人事部'), roles: [READ] }] }
    });
    assert.equal(buttonNamed(unique.host, STRINGS.remove) !== undefined, true);
  });

  it('reports who was removed', () => {
    let removed;
    const { host } = panel(
      { state: { hasUnique: true, assignments: [{ principal: group('人事部'), roles: [READ] }] } },
      { onRemove: (p) => { removed = p; } }
    );

    buttonNamed(host, STRINGS.remove).click();
    assert.equal(removed.title, '人事部');
  });

  it('only offers the add box once the permissions are its own', () => {
    assert.equal(panel().host.querySelector('.mdv-perm-add-row'), null);
    assert.notEqual(panel({ state: { hasUnique: true } }).host.querySelector('.mdv-perm-add-row'), null);
  });

  it('grants the found principal the chosen level', async () => {
    let granted;
    const { host } = panel(
      {
        state: { hasUnique: true },
        searchPrincipals: () => Promise.resolve([group('人事部')])
      },
      { onAdd: (principal, roleId) => { granted = { principal, roleId }; } }
    );

    host.querySelector('.mdv-perm-add-row input').value = '人事';
    buttonNamed(host, STRINGS.search).click();
    await Promise.resolve();
    await Promise.resolve();

    host.querySelector('.mdv-perm-role-select').value = String(EDIT.id);
    buttonNamed(host, STRINGS.add).click();

    assert.equal(granted.principal.title, '人事部');
    assert.equal(granted.roleId, EDIT.id);
  });

  it('says so when the directory had nothing to offer', async () => {
    const { host } = panel({
      state: { hasUnique: true },
      searchPrincipals: () => Promise.resolve([])
    });

    host.querySelector('.mdv-perm-add-row input').value = 'いない人';
    buttonNamed(host, STRINGS.search).click();
    await Promise.resolve();
    await Promise.resolve();

    assert.equal(host.querySelector('.mdv-perm-results').textContent, STRINGS.noResults);
  });

  it('reports a failed change inside the panel', () => {
    const { host, handle } = panel();

    handle.setError('アクセス許可を変更できませんでした。 HTTP 403');
    assert.match(host.querySelector('.mdv-tag-error').textContent, /HTTP 403/);
  });
});
