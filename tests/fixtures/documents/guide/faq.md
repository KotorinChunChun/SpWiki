# よくある質問

[← ガイドに戻る](index.md)

## 表示について

### 見出しリンクは使えますか

使えます。見出しには自動で id が振られるので、[サニタイズについて](#サニタイズについて) のような
ページ内リンクが機能します。

### 対応しているコードの言語は

JavaScript / TypeScript / PowerShell / JSON / HTML / CSS / SQL / C# の 8 種類です。
未対応の言語は色付けせず、そのまま表示します。

```javascript
const params = new URLSearchParams(location.search);
console.log(params.get('md'));
```

```csharp
public sealed class MarkdownPath
{
    public string Value { get; }
    public MarkdownPath(string value) => Value = value;
}
```

```css
.markdown-viewer pre {
  overflow-x: auto;
}
```

```html
<div class="markdown-viewer">
  <p>Hello</p>
</div>
```

```yaml
# 未登録の言語はハイライトされません
name: markdown-viewer
```

## サニタイズについて

Markdown に埋め込まれた危険なタグは DOMPurify が除去します。
以下は除去されるため、下の一覧には何も表示されないのが正常な状態です。

<script>alert('script は除去されます');</script>
<iframe src="https://example.com"></iframe>
<object data="malicious.swf"></object>
<img src="x" onerror="alert('onerror は除去されます')">

[javascript: スキームのリンク](javascript:alert('blocked'))

除去結果の確認:

- `<script>` … 消えていること
- `<iframe>` … 消えていること
- `<object>` … 消えていること
- `onerror` 属性 … 消えていること

## チェックリスト

- [x] 見出しが表示される
- [x] 表が表示される
- [x] コードがハイライトされる
- [ ] 本番用 Markdown への差し替え

## その他

| 質問 | 回答 |
| --- | --- |
| 画像は表示できますか | ライブラリ内の相対パスで表示できます |
| 検索はできますか | 将来拡張です |
| PDF 出力は | 将来拡張です |

用語は `インライン コード` のように書けます。~~取り消し線~~ も使えます。
