# pdf2pptx

PDFを編集可能なPPTXに変換するCLIツール（自社ツール・Phase 1 MVP）。

## 概要

PyMuPDFのredaction機能でPDFページの背景画像からテキストを削除し（編集可能テキストとの二重写りを防止）、
その上にPowerPointの編集可能なテキストボックスを重ねて配置する方式でPDF→PPTX変換を行う。

Phase 1の範囲:

- 編集対象は横書きテキストのみ。縦書き・回転テキストは背景画像側に残す（編集不可）。
- スキャンPDF（テキストレイヤーなしの画像PDF）は検出し、背景画像のみを配置した上で警告を出す。

## セットアップ

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install pymupdf python-pptx pillow
```

## 使い方

```bash
python convert.py input.pdf output.pptx
```

## ディレクトリ構成

| パス | 内容 |
|------|------|
| `src/`  | ソースコード |
| `tests/`| テスト |
| `docs/` | ドキュメント |
