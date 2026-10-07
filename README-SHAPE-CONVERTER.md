# SHAPE CONVERTER

正方形・長方形と角丸四角形を相互変換するPowerPointアドインです。

## 機能

- 正方形／長方形から角丸四角形へ変換
- 角丸四角形から正方形／長方形へ変換
- 位置、サイズ、回転、重なり順を維持
- 単色または塗りなし、枠線を維持
- テキスト、文字色、フォント、サイズ、太字などを可能な範囲で維持
- 段落の配置、箇条書きを可能な範囲で維持

## GitHub Pagesへの追加

既存の `ppt-copy-addin` リポジトリ直下へ、次の6ファイルをアップロードします。

- `shape-converter.html`
- `shape-converter.css`
- `shape-converter-core.js`
- `shape-converter.js`
- `shape-converter-commands.html`
- `shape-converter-help.html`

`assets` フォルダは既存のものを使います。上書きは不要です。

アップロード後、次のURLが表示できることを確認します。

`https://kopeta-jp.github.io/ppt-copy-addin/shape-converter.html`

## PowerPointへの追加

1. PowerPointを完全に終了します。
2. Finderで `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` を開きます。
3. `manifest-shape-converter.xml` だけを `wef` フォルダへコピーします。
4. PowerPointを開き直します。
5. 「ホーム」から「アドイン」を開き、`SHAPE CONVERTER` を追加します。

## 使い方

1. 標準の正方形、長方形、または角丸四角形を1個選択します。
2. `SHAPE CONVERTER` のタスクペインを開きます。
3. 変換方向を選びます。
4. 必要な場合はPowerPointの「元に戻す」で復元します。

## 制限

- PowerPoint JavaScript APIでは図形タイプを直接変更できないため、新しい図形へ安全に置き換えます。
- グラデーション、パターン、画像塗りは保持できないため、変換前に処理を止めます。
- 影、光彩、3DなどはAPIで取得・再設定できないため対応していません。
- 文字数が800文字を超え、複数の文字書式が混在する場合、一部の書式が統一される可能性があります。
- PowerPointApi 1.10を使用します。
