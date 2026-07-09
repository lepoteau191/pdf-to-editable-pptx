# pdf2pptx

PDFを編集可能なPPTXに変換するCLIツール（自社ツール・Phase 1 MVP）。
PyMuPDFのredactionでテキストを消した背景画像を敷き、その上に編集可能テキストボックスを重ねる方式。

## コマンド

- 環境構築: `uv venv --python 3.12 .venv && uv pip install --python .venv/bin/python -r requirements.txt`
- 実行: `.venv/bin/python convert.py input.pdf output.pptx [--dpi 150] [--debug-dir DIR]`
- テスト: `.venv/bin/python -m pytest -q`

## 構成

- `convert.py` … 変換本体（抽出 → redaction → 背景画像化 → PPTX組み立て）
- `tests/` … pytest。フィクスチャは `tests/fixtures_gen.py` が生成する合成PDFのみ（ネットワーク不使用）

## 守ること

- テストは合成PDFのみ。実PDF（顧客資料等）や生成物（*.pptx）をコミットしない。
- コード変更後は `.venv/bin/python -m pytest -q` を通す。
- Phase 1 の範囲: 横書き・可視のみ編集対象。縦書き・回転テキスト・回転ページ・不可視(OCR)テキスト・
  文字化けspan・それらと重なる横書き行は背景に残す。スキャンPDFは背景のみ+警告。
- `--debug-dir` は検証用。機密PDFでは使わない（背景画像が平文PNGでディスクに残るため）。
