# ビルドコマンド

プロジェクトルートで実行します。PowerShell は pwsh を使用してください。
環境要件は [開発者ガイド](DEVELOPERS_GUIDE.md)を参照してください。

| コマンド | 用途 |
| --- | --- |
| `npm ci` | ロックファイルに従って依存パッケージを導入 |
| `npm run setup:node22` | ポータブル Node 22 を準備 |
| `npm run docs:build` | README と docs から同梱ヘルプを生成 |
| `npm test` | 中核ロジックの単体テスト |
| `npm run harness:build` | ローカル確認画面をビルド |
| `npm run harness` | 確認用サーバーを起動（http://localhost:4322/） |
| `npm run build` | SPFx の開発用ビルド |
| `npm run serve` | SPFx ワークベンチを起動 |
| `npm run release` | クリーンビルドし配布物を生成 |

配布物は `packages/1.0.0.0/SpWiki.sppkg` に出力されます。
パッケージ生成はアップロードや公開を行いません。
