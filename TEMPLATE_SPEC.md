# `assets/template.ppt` 仕様

このツールは `assets/template.ppt` という **動作実績ある PowerPoint 97-2003 (.ppt) ファイル** に対し、埋め込み JPEG だけをバイト置換することで会場プレイヤー (Panasonic PW_DPC01 / e-Signage Lite) 用の .ppt を生成する。

## バイト配置の前提（必須）

| 範囲 | サイズ | 内容 |
|---|---|---|
| `[0, 537)` | 537 B | OLE Compound File ヘッダ + PowerPoint Document ストリーム先頭（Office Art レコードヘッダ含む） |
| `[537, 537 + 224068)` | 224,068 B | 埋め込み JPEG 本体（`FF D8 FF` で開始、`FF D9` で終端） |
| `[224605, 337920)` | 113,315 B | OLE FAT、PowerPoint トレーラ、SummaryInformation 等 |

この前提が崩れたらツールは動かなくなる。テンプレを差し替えたら `ppt-builder.js` 冒頭の以下定数を必ず更新すること：

```js
const JPEG_OFFSET = 537;
const JPEG_SLOT_SIZE = 224068;
const TEMPLATE_TOTAL_SIZE = 337920;
```

## なぜこの方式か

候補比較：

| 方式 | 利点 | 欠点 |
|---|---|---|
| **採用：バイト置換 + JPEG COM パディング** | 外部ライブラリ不要、コード最小、構造を一切壊さない | 出力は常に 337,920 B、画像が小さくても容量を食う |
| js-cfb で stream 置換 + OfficeArt サイズフィールド再計算 | 任意サイズ JPEG 可、出力サイズ最適化 | 仕様3つ（MS-CFB, MS-ODRAW, JPEG）を正しく扱う必要、依存追加 |
| pptxgenjs から .ppt 変換 | 既存資産流用 | ブラウザでは .pptx → .ppt 変換手段が事実上ない |
| サーバ側 LibreOffice 変換 | 確実 | ブラウザ完結という設計方針を破壊 |

バイト置換方式の核は「**OLE は JPEG ストリームを物理的に連続配置している**」というサンプル分析の事実。2件の実機サンプル（337,920 B / 476,672 B）で JPEG の開始 offset が両方とも 537 だったことから確信した。

## JPEG パディングのメカニズム

新 JPEG が 224,068 B より小さい場合、JPEG 仕様の **COM マーカー（`FF FE LEN_HI LEN_LO data...`）** で末尾を埋める。COM はデコーダに無視される。実装は `ppt-builder.js` の `padJpegToExactSize()` 参照。

最大 COM サイズは 65,535 B（length field 含む）。複数 COM で任意長を埋められる。

## スライド背景色（黒）

会場 PW_DPC01 のモニターは 16:9 だが画像配置枠は 1.415:1 のため左右が余る。その余白を黒にするため、テンプレ `.ppt` 内の**背景シェイプ（fBackground=1 の OfficeArtFSP）に紐付く OfficeArtFOPT の `fillColor`（opid 0x0181）**を直接 RGB(0,0,0) に書換済み。

書換箇所（バイトオフセット）：
- `238821`〜`238826`: `81 01 00 00 00 08` → `81 01 00 00 00 00`（マスタースライドの背景）
- `275781`〜`275786`: `81 01 00 00 00 08` → `81 01 00 00 00 00`（タイトルマスターの背景）

元値の `0x08000000` はスキームカラー index 0（=bg、白）参照。最後のフラグバイト `0x08`（fSchemeIndex）を `0x00` に倒すと RGB 直値として解釈され、`00 00 00` = 黒になる。スキーム自体は変更していないので、テキスト等の他要素は影響を受けない。

総バイト数は変わらないので `JPEG_OFFSET` / `JPEG_SLOT_SIZE` / `TEMPLATE_TOTAL_SIZE` の更新は不要。

## 画像サイズの強制

入力画像（PDF レンダ or 元画像）を **1684 × 1190 px**（テンプレ画像の元サイズ）に**レターボックス**で押し込む。これはテンプレのスライドに記述されている画像配置矩形のアスペクト比に一致させるため。アスペクトが違うと再生時に歪む。

## 何が壊れたら何を疑うか

| 症状 | 疑い |
|---|---|
| プレイヤーで真っ黒・無音 | JPEG が破損（パディング処理バグ） |
| プレイヤーで前のサンプルが再生される | Schedule.csv の日付 / TimeTable.csv |
| スライドはあるが画像なし | JPEG offset / size の想定がずれた（CFB 構造変化） |
| 「対応形式でない」エラー | テンプレ .ppt のメタデータがプレイヤー互換でない |

## テンプレ更新時のチェックリスト

1. 新しいテンプレ .ppt の最初の JPEG オフセットと長さを確認：
   ```sh
   python3 -c "
   import re
   d=open('assets/template.ppt','rb').read()
   s=re.search(b'\\xff\\xd8\\xff', d).start()
   e=d.find(b'\\xff\\xd9', s)+2
   print('offset', s, 'size', e-s, 'total', len(d))
   "
   ```
2. その値で `ppt-builder.js` の3定数を更新
3. 実機（PW_DPC01 + e-Signage Lite）で生成パッケージを再生して目視確認

## サンプル

リポジトリ外の参照サンプル（個人情報を含むためコミット不可）：
- `5.16.ppt`: JPEG offset=537, size=224068, total=337920（← 現テンプレの元）
- `0509.ppt`: JPEG offset=537, size=349986, total=476672（参考、未使用）
