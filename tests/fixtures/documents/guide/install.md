# インストール手順

[← ガイドに戻る](index.md) ／ [← トップに戻る](../README.md)

## 目次

## 前提条件

| 項目 | バージョン |
| --- | --- |
| Node.js | 22.14.0 以上 23 未満 |
| SPFx | 1.23.2 |
| ブラウザー | Microsoft Edge / Google Chrome |

## 手順

1. パッケージをビルドする
2. アプリ カタログにアップロードする
3. サイトにアプリを追加する
4. ページに Web パーツを配置する

### 1. ビルド

```powershell
npm install
npm run package
```

生成物は `sharepoint/solution/SpWiki.sppkg` です。

### 2. アップロード

アプリ カタログに `.sppkg` をアップロードし、展開を許可します。

### 3. Web パーツの設定

```json
{
  "libraryName": "Documents",
  "startFile": "README.md",
  "siteUrl": ""
}
```

### 4. 動作確認

以下の URL を開いて表示を確認します。

```
/sites/docs/SitePages/docs.aspx?file=guide/install.md
```

## 画像の参照

親フォルダーの画像も相対パスで参照できます。

![サンプル画像](../images/sample.png)

## 次の手順

- [アップデート手順](update.md)
- [よくある質問](faq.md)
- [HTML ページのサンプル](html-sample.html) — `.html` を置いたときの見え方
