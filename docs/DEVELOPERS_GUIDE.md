# 開発者向け

SpWiki 1.0 の環境構築・設計・テスト・パッケージ作成を説明します。
構成・運用上の注意事項は [TECHNICAL_REFERENCE.md](TECHNICAL_REFERENCE.md) を参照してください。

製品バージョンは 1.0、npm は `1.0.0`、SPFx パッケージは `1.0.0.0` です。

---

## 1. 環境

| 項目 | バージョン |
| --- | --- |
| Node.js | **22.14.0 以上 23 未満**（SPFx 1.23.2 の要件） |
| SPFx | 1.23.2 |
| パッケージ マネージャー | npm |
| シェル | pwsh（PowerShell 7+） |

SPFx のビルド ツールは Node 24 を受け付けません。システムの Node を差し替えずに済むよう、
**ポータブル版の Node 22 を `.tools/node22` に置いて使います**（PATH は変更しません）。

```bash
npm install
```

```bash
npm run setup:node22
```

`dev/scripts/setup-node22.ps1` が nodejs.org の公式 ZIP を取得し、公開されている SHA256 と
照合してから展開します（`.tools` は Git 管理外）。

`npm test` と `npm run harness` は SPFx のツールチェーンを使わないので、Node 24 のままでも動きます。

---

## 2. 構成

```
SpWiki
├─ src/webparts/spWiki/
│   ├─ SpWikiWebPart.ts                 Web パーツ本体（描画・履歴・クリック処理）
│   ├─ SpWikiWebPart.manifest.json
│   ├─ core/                             SPFx に依存しない中核ロジック
│   │   ├─ pathUtils.ts                  相対パス解決・拡張子から形式を判定
│   │   ├─ urlUtils.ts                 ?file= の読み書き
│   │   ├─ markdown.ts                   marked 設定
│   │   ├─ sanitize.ts / sanitizeConfig.ts  DOMPurify 設定
│   │   ├─ highlight.ts                  highlight.js 設定
│   │   ├─ domRewrite.ts                 リンク・画像の書き換え
│   │   ├─ slugify.ts                    見出し id の生成
│   │   ├─ toc.ts                        `## 目次` からの目次生成とアウトライン
│   │   ├─ headings.ts                   見出しの折りたたみとリンクのコピー
│   │   ├─ renderer.ts                   上記をつなぐ描画パイプライン（md / html 共通）
│   │   ├─ tags.ts                       タグ列の読み書き（型ごとの差異を吸収）
│   │   ├─ tagColors.ts                  タグ名から決まる淡色と前景色
│   │   ├─ relatedPages.ts               タグ検索の絞り込みとページ一覧の描画
│   │   ├─ navPanel.ts                   ☰ メニューとリネーム パネルの描画
│   │   ├─ chrome.ts                     ツールバー・タグ バー・タグ編集・通知の描画
│   │   ├─ appLink.ts                    spmd:// URL の組み立て
│   │   ├─ pageTree.ts                   罫線付きのツリー描画
│   │   ├─ directives.ts                 特殊見出しの展開
│   │   ├─ search.ts                     検索ページ
│   │   ├─ permissions.ts                アクセス許可の画面
│   │   ├─ helpPage.ts                   ヘルプ ページ
│   │   └─ bundledDocs.ts                生成物: 同梱ドキュメント
│   ├─ services/
│   │   ├─ SharePointDocumentProvider.ts ファイル取得とライブラリ解決
│   │   └─ SharePointLibraryService.ts   列一覧・アイテム・タグ更新
│   └─ loc/                              en-us / ja-jp
├─ src/extensions/spWikiCommand/        ListView Command Set
│   ├─ SpWikiCommandSet.ts              コマンド定義と遷移
│   ├─ MarkdownPreviewDialog.ts          モーダル内のビューアー
│   ├─ NoticeDialog.ts                   コマンドからの案内ダイアログ
│   └─ loc/
├─ src/styles/viewer.module.scss         Web パーツとモーダルで共有する見た目
├─ docs/                                 公開ガイドと技術資料
├─ tests/                                単体テスト
│   ├─ fixtures/documents/               サンプル文書（Markdown と HTML）
│   └─ harness/                          ブラウザーでの確認用画面
├─ dev/scripts/
│   ├─ spfx.ps1                          ポータブル Node 22 で gulp を起動
│   ├─ setup-node22.ps1                  ポータブル Node 22 の取得
│   ├─ release.ps1                       リリース（バージョン管理込み）
│   ├─ build-docs.mjs                    README と docs を同梱ヘルプに変換
│   ├─ set-page-layout.js                PnP 無しでページ レイアウトを変更
│   └─ ...                               ハーネス／テストの補助
├─ config/                               SPFx ビルド設定
└─ packages/<バージョン>/SpWiki.sppkg          リリース成果物（Git 管理外）
```

### 設計方針

**`core/` は DOM だけに依存し、SPFx の API を一切使いません。**
これにより次の 3 つが成立します。

- 単体テスト（jsdom）で描画パイプラインまで検証できる
- ローカル ハーネス（実ブラウザー）で SharePoint 無しに確認できる
- Web パーツとコマンドセットのモーダルで**同じコードが動く**（挙動が食い違わない）

SharePoint に触るのは `services/` と各ホスト（Web パーツ / コマンドセット）だけです。

---

## 3. ローカルでの確認

### 3.1 ハーネス

`tests/fixtures/documents` をライブラリ代わりにして、実ブラウザーで確認できます。

```bash
npm run harness:build
```

```bash
npm run harness
```

`http://localhost:4322/` を開きます。確認できること:

- 見出し・表・箇条書き・コード・画像・リンクの表示
- `install.md` → `guide/install.md`、`../README.md` → `README.md` の解決
- リンク遷移時の `?file=` 書き換えと、戻る・進む
- 404 と読込中の表示
- `<script>` `<iframe>` `<object>` `onerror` の除去（`guide/faq.md`）
- `.html` の描画とサニタイズ、`.md` との相互リンク（`guide/html-sample.html`）
- `## 目次` からの目次生成（`guide/install.md`）、見出しの折りたたみとリンクのコピー
- タイトル（押すと開始ファイルへ）とファイル名／フォルダーの 2 段ヘッダー
- 3 段のタグ バー、列名を押してタグ編集、タグを押しての検索
- `#見出し` 付き URL を直接開いたときの自動スクロール
- 未定義タグの追加。`mock-metadata.json` の `allowFillIn` が `false` の列
  （既定では「分類」）では入力欄がグレーアウトします
- ☰ メニューのフォルダー ツリー、テンプレートからの新規ページ作成（形式と編集先を選ぶ）、
  ファイル名の変更（作成とリネームはハーネスではメモリ上だけで、実ファイルは書き換えません）
- ツールバー、タグ編集、「アプリで開く」の案内
- 目次ボタンと ☰ の**シングル＝ポップアップ / ダブル＝左に固定**、左ペインの入れ替わり
- 上部パネルの固定表示（`--mdv-top` の実測値が `#viewer` に載ります）
- フォルダー名のパンくず。遷移すると `?q=` などの検索パラメーターが落ちること
- 特殊見出し 4 種（`tests/fixtures/documents/guide/index.md` の末尾に確認用として置いてあります）
- コードブロックのコピー ボタン、右下のトースト、ヘルプ ページ

**ハーネスで確認できないもの**（SharePoint の REST が要るため）:
フォルダーをページとして開く、アクセス許可の画面、鍵アイコンの状態、検索ページ。
これらは単体テストとブラウザーでの手動確認の対象外なので、実機で見てください。

タグは `tests/harness/mock-metadata.json` から読み込みます（保存はメモリ上だけ）。
列の型による UI の違いを見たいときは `field.kind` を
`text` / `choice` / `multiChoice` に書き換えてください。

### 3.2 単体テスト

```bash
npm test
```

`core/` を esbuild でバンドルし、jsdom 上で Node のテストランナーが実行します
（ パス解決・URL 操作・描画・HTML 描画・サニタイズ・目次・アウトライン生成・
見出しの折りたたみ・タグ・タグ検索・タグの配色・新規ページのテンプレートと入力フォーム・
プロパティ ウィンドウの列読み込み判定・メニューとリネームの入力処理・ツールバー・
罫線付きツリー・特殊見出し・検索・アクセス許可の画面）。

実行件数と結果は `npm test` の出力で確認してください。

### 3.3 ワークベンチ

```bash
npm run serve
```

`https://<テナント>.sharepoint.com/_layouts/workbench.aspx` が開きます。
ワークベンチには `?file=` を付けられないため、URL 遷移の確認は実ページで行ってください。

---

## 4. リリース

```bash
npm run release
```

`dev/scripts/release.ps1` が clean → `bundle --ship` → `package-solution --ship` を通しで実行し、
`packages/<バージョン>/SpWiki.sppkg` を出力します。最後に SHA256 と手順を表示します。

| コマンド | 動作 |
| --- | --- |
| `npm run release` | 現在のバージョンのまま再パッケージ |
| `pwsh -File dev/scripts/release.ps1 -Bump Build` | パッチ番号を上げる |
| `pwsh -File dev/scripts/release.ps1 -Bump Minor` | マイナー番号を上げる |
| `pwsh -File dev/scripts/release.ps1 -Bump Major` | メジャー番号を上げる |
| `pwsh -File dev/scripts/release.ps1 -Version 1.0.0.0` | 直接指定 |
| `... -SkipClean` | `lib/` を再利用（クリーンを省略） |

`-Bump` / `-Version` は `config/package-solution.json` の `solution.version` と
`features[].version`、`package.json` の `version` を書き換え、書き込み後に読み直して検証します。

> ### ファイル名にバージョンを付けない
>
> アプリ カタログは**ファイル名**でパッケージを同定します。バージョン付きの名前にすると
> 置き換えにならず `Disabled` になります。バージョンはフォルダー名で持ちます。
>
> 置き場所が `release/` ではなく `packages/` なのは、**`gulp clean` が `release/` を
> 丸ごと削除する**ためです。

> ### リリース後はサイズを確認する
>
> `-SkipClean` は `lib/` を再利用しますが、`dist/` と `release/` は必ず削除します。
> デバッグ ビルドの非圧縮バンドルが同梱されると、パッケージが不必要に大きくなります。
> 配布前に成果物の内容とサイズを確認してください。

### 個別のタスクを叩く

```bash
pwsh -File dev/scripts/spfx.ps1 bundle --ship
```

---

## 5. 実装メモ

### セキュリティ

- **リンク書き換えはサニタイズの後**に行います。ドキュメント側が `data-md-path` を偽装しても
  DOMPurify に除去されるので、遷移先はこちらが計算した値だけになります
- **`.html` も Markdown から生成した HTML とまったく同じ設定を通します。**
  `renderer.ts` が分岐するのは marked を通すかどうかだけで、サニタイズ以降は共通です。
  `style` 属性と `<style>` は許可していないので、ドキュメント側がページの見た目に
  干渉したり CSS の `url()` で外部へ通信したりはできません
- **`?file=` は必ず正規化**します。`..` はライブラリのルートで止まり、絶対 URL は拒否します。
  受け付ける拡張子は `.md` / `.markdown` / `.html` / `.htm` だけです
  （ビューアー経由で任意のファイルを覗けないようにするため）。
  **拡張子の無いパスはフォルダーとして扱いますが、これは `sanitizeDocPath(raw, true)` と
  明示的に許可した呼び出しだけ**です。無条件に通すと `?file=web.config` を塞いでいた
  判定が緩みます
- **`data:` の画像は表示されません**。DOMPurify の既定で許可していないためです。
  必要なら `core/sanitizeConfig.ts` で緩和できます

### 描画

- **ハイライト対象は 8 言語のみ**登録しています（JavaScript / TypeScript / PowerShell /
  JSON / HTML / CSS / SQL / C#）。追加は `core/highlight.ts` に 2 行です
- **見出しに id を振っています**。ページ内リンクが動くほか、目次生成にそのまま使えます
- **`#見出し` へのスクロールは自前で行います**。本文が届くのはブラウザーが
  フラグメントを探し終えた後なので、描画直後に該当 id を探して `scrollIntoView()` します。
  id には `1-ビルド` のような CSS セレクターとして不正な値が入るため、
  `querySelector('#...')` ではなく `[id]` を走査して比較します
- **`mdv-` プレフィックスのクラスは `:global()`** で宣言しています。CSS モジュールの
  マングリングを避け、Web パーツとハーネスで同じマークアップを共有するためです

### データ

- **ライブラリ名は「①リスト タイトル → ②フォルダー パスの実在確認 → ③専用エラー」**の順で
  解決します。無言でフォールバックせず、設定ミスだと分かるようにしています
- **タグの絞り込みはクライアント側**です。列の型ごとに OData の `$filter` の書き方が
  違う（複数選択列は `eq` が効かない）ため、型に依存しない処理にしています。
  500 件を超えるライブラリでは `SharePointLibraryService.ts` の `ITEM_FETCH_LIMIT` を調整します
- **タグ列の型は実行時に判別**します（`Text` / `Note` / `Choice` / `MultiChoice`）。
  読み取りは文字列・配列・`{results:[]}` のどれでも受け付けます
- **一覧とツールバーは本文の描画をブロックしません**。メタデータ取得は Markdown 表示後に
  走らせ、届いた時点でボタンを有効化します

### ホスト連携

- **「編集」はライブラリの既定ビューに `id` と `parent` を付けた URL** を開きます
  （`.../Forms/AllItems.aspx?id=<ファイル>&parent=<フォルダー>`）。
  ライブラリでファイルをクリックしたときと同じ、SharePoint 標準のプレビュー／編集画面です。
  組み立ては `core/urlUtils.ts` の `buildLibraryViewUrl`（単体テストあり）。
  ビュー ページの場所は**ファイルの親リスト**から取ります（`ListItemAllFields/ParentList`）。
  Web パーツがライブラリ内のサブフォルダーを指している場合、
  `<設定したフォルダー>/Forms/AllItems.aspx` は存在しないためです。取得結果はキャッシュします
- 親リストを読めなかったときだけ `_layouts/15/Doc.aspx?sourcedoc={UniqueId}` に退避します。
  これは Office 用のビューアーなので `.md` / `.html` には向きません
- **ファイルの作成は `Files/add(url=...,overwrite=false)`、移動は `moveto(newurl=...,flags=1)`** です。
  どちらも上書きを禁止しているので、同名のファイルがあると成功したふりをせずに失敗します
- **未定義のタグを付けられるかは列の `FillInChoice` で決まります**。無効な列では入力欄を
  グレーアウトします。列定義を書き換えてしまうとリストの管理権限が要るので、
  列の設定に従います（[TECHNICAL_REFERENCE.md](TECHNICAL_REFERENCE.md)）
- **新規ページはテンプレート付きで作成し、指定された編集画面へ渡します**。
  SharePoint 編集は同じタブで遷移します。クリックから時間が経った `window.open` は
  ポップアップ ブロックに止められるためです。アプリ編集は `spmd:` を起動するだけなので
  ビューアーは新しいページを表示したまま残ります
- **コマンドのタイトルはローカライズ リソースから実行時に設定**しています。
  マニフェストの `ja-jp` が SharePoint に無視されるためです（[TECHNICAL_REFERENCE.md](TECHNICAL_REFERENCE.md)）
- **ダイアログのサイズは中身の要素に持たせています**。`IDialogConfiguration` に
  サイズ指定が無く、Fluent UI の DOM を直接触るのを避けたためです

---

## 6. 主要な識別子

変更するとアプリ カタログ上で別物になるので注意してください。

| 対象 | 値 |
| --- | --- |
| ソリューション ID | `b08b7dff-253e-41e8-bb5a-2c471b64b2ea` |
| フィーチャー ID | `a0735bdc-baab-49eb-8d1d-687e49d64589` |
| Web パーツ コンポーネント ID | `85e1ddfc-3c81-4cf2-b89f-2638b3c51b34` |
| コマンドセット コンポーネント ID | `7bf4c5c5-631b-48cf-b066-e92bfb4c91dc` |

---

## 7. Web パーツのプロパティ

| グループ | 項目 | プロパティ名 | 既定 |
| --- | --- | --- | --- |
| 参照元 | タイトル | `title` | (空 = `ドキュメント`) |
| | ライブラリ名 | `libraryName` | `Shared Documents` |
| | 開始ファイル | `startFile` | `README.md` |
| | サイト URL | `siteUrl` | (空) |
| タグ | タグの列 1 | `tagFieldName` | (空) |
| | タグの列 2 | `tagFieldName2` | (空) |
| | タグの列 3 | `tagFieldName3` | (空) |
| | タグ検索ページ | `tagSearchPageUrl` | (空 = 同じページ内で検索) |
| | 列名をクリックしてタグを編集できるようにする | `showTagEditor` | `true` |
| ツールバー | 「編集」ボタン | `showEditButton` | `true` |
| | 鍵アイコン（アクセス許可） | `showPermissionsButton` | `true` |
| | 「アプリで開く」ボタン | `showOpenInAppButton` | `false` |
| | URL スキーム | `appProtocol` | `spmd` |

「アプリで開く」は明示的に `true` の場合だけ表示します。新規配置時と未設定時は OFF です。
他のトグルは未設定を ON として扱います（`_isEnabled()`）。

---

## 8. 拡張するときの入口

| やりたいこと | 触る場所 |
| --- | --- |
| ハイライト言語を増やす | `core/highlight.ts` に import と `registerLanguage` を 1 行ずつ |
| 許可する HTML タグ・属性を変える | `core/sanitizeConfig.ts`（`.md` と `.html` の両方に効きます） |
| 表示できる拡張子を増やす | `core/pathUtils.ts` の `isDocumentPath` / `documentFormat` と `core/renderer.ts` の分岐 |
| 目次の見出し語を変える | `core/toc.ts` の `TOC_MARKER` |
| タグの配色を変える | `core/tagColors.ts`（色相はタグ名のハッシュ。乱数ではないので再読込でも変わりません） |
| タグの列を 4 つ以上にする | `SpWikiWebPart` の `_getTagFieldNames()` とプロパティ定義。以降は配列で流れます |
| 目次を自動生成する | `core/domRewrite.ts` の `addHeadingIds` が付けた id を使う |
| タグ検索の取得上限を変える | `services/SharePointLibraryService.ts` の `ITEM_FETCH_LIMIT` |
| タグ列の型を増やす | `core/tags.ts` の `toTagFieldKind` と `core/chrome.ts` の `renderTagEditor` |
| コマンドを増やす | `extensions/spWikiCommand/SpWikiCommandSet.manifest.json` の `items` と `onExecute` |
| 表示言語を増やす | `loc/` に `<locale>.js` を追加し、`config/config.json` の `localizedResources` は `{locale}` のままで可 |

---

## 9. 対応範囲と導入時の確認

Mermaid 描画と PDF 出力は 1.0 の提供機能に含まれません。

導入先では、次の項目を確認してください。ローカルテストだけでは SharePoint 上の動作を保証できません。

- アクセス許可の継承・解除・付与と、個別権限を持つ子項目の一覧
- ページ構造に応じた上部パネル・左ペインの固定表示
- フォルダーの表示とライブラリのコマンド
- 利用者の権限に応じたファイル・タグの読み書き
