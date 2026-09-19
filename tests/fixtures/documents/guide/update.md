# アップデート手順

[← ガイドに戻る](index.md) ／ [← インストール手順](install.md)

## 概要

このページは、1.0 の導入後に更新版を配布する場合の手順例です。

既存の `.sppkg` を新しいものに差し替えます。Web パーツの設定は保持されます。

## 手順

1. `npm run package` で新しい `.sppkg` を作る
2. アプリ カタログで同名ファイルを上書きアップロードする
3. サイトのアプリを更新する

## バージョン番号

配布時は `npm run release -- -Bump Build` でバージョンを上げて作成します。
次の番号は更新時の例で、現在の初版は `1.0.0.0` です。

```json
{
  "solution": {
    "name": "SpWiki",
    "version": "1.0.1.0"
  }
}
```

## 確認クエリ

```sql
SELECT Title, Modified
FROM   Documents
WHERE  Title LIKE '%.md'
ORDER  BY Modified DESC;
```

## ロールバック

- 直前のバージョンの `.sppkg` を再アップロードします
- ブラウザーのキャッシュをクリアします
