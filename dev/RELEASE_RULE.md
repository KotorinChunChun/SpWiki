# リリースルール

## 初版の識別

- 製品表記: SpWiki 1.0
- npm バージョン: `1.0.0`
- SPFx ソリューション・フィーチャーバージョン: `1.0.0.0`
- 配布ファイル: `packages/1.0.0.0/SpWiki.sppkg`

## 作成と確認

1. `npm ci` と `npm run setup:node22` で開発環境を準備する。
2. `npm run docs:build`、`npm test`、`npm run harness:build` を実行する。
3. `npm run release` でクリーンな製品ビルドとパッケージを作る。
4. 出力のバージョン・サイズ・SHA256を確認する。
5. 導入先で必要な確認項目は [開発者ガイド](../docs/DEVELOPERS_GUIDE.md) に従う。

配布ファイル名は常に `SpWiki.sppkg` とし、バージョンは保存フォルダー名で管理する。
版を更新する場合は package.json、package-lock.json のルート情報、SPFx の設定を揃える。

## 公開対象

ソース、公開文書、テスト、開発ルール、スクリプト、依存関係のロックファイルを追跡する。
依存パッケージ、生成物、ローカル設定、dev/private と dev/temp は Git に含めない。
公開前に `git status` とコミット対象を確認する。配布物はソースとは別に扱う。

パッケージ作成・コミット・外部への公開は別の操作である。
アップロードの手順は [管理者ガイド](../docs/ADMIN_GUIDE.md) を参照する。
