# システム概要

このライブラリは **SpWiki** の動作確認用サンプルです。
本物の Markdown が用意されるまでの間、表示・リンク遷移・サニタイズの検証に使います。

![ロゴ](images/logo.png)

## サブページ一覧

## メニュー

- [インストール手順](guide/install.md)
- [アップデート手順](guide/update.md)
- [よくある質問](guide/faq.md)
- [ガイドの入口](guide/index.md)
- [HTML ページのサンプル](guide/html-sample.html) — `.md` から `.html` へも同じように移動できます

## このビューアーでできること

| 機能 | 説明 | 状態 |
| --- | --- | --- |
| Markdown 表示 | 見出し・表・箇条書き・コード・画像・リンク | 実装済み |
| 相対リンク解決 | `install.md` / `../README.md` を解決 | 実装済み |
| URL 更新 | `history.pushState()` で `?file=` を書き換え | 実装済み |
| 戻る・進む | `window.onpopstate` に対応 | 実装済み |
| 目次自動生成 | 見出しから TOC を生成 | 将来拡張 |

## 使い方

1. ドキュメントライブラリに `.md` を置く
2. ページに Web パーツを配置する
3. `?file=` にライブラリ相対パスを指定して共有する

例:

```
https://tenant.sharepoint.com/sites/docs/SitePages/docs.aspx?file=guide/install.md
```

## サンプルコード

```typescript
export function resolve(current: string, link: string): string {
  const folder = current.slice(0, current.lastIndexOf('/'));
  return folder === '' ? link : `${folder}/${link}`;
}
```

> 相対リンクは「いま表示しているファイルのあるフォルダー」を基準に解決されます。
> 詳しくは [インストール手順](guide/install.md) を参照してください。

---

外部サイトへのリンクは新しいタブで開きます: [SharePoint Framework](https://aka.ms/spfx)

## タグページ一覧:設計
	そのタグが付いた全ページ

## サブページ一覧:guide
	相対パス。書いたページのフォルダー基準

## サブページ一覧:/guide
	先頭 / でライブラリ ルート基準。../reference も可

## ページツリー	ルートから
	現在のフォルダーまでの階層と、そのフォルダーのページ一覧(現在のページは太字)

## 関連ページ一覧:install
	キーワードに一致するページ
