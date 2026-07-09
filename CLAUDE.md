# pdf2pptx

PDFを編集可能なPPTXに変換するCLIツール（自社ツール・Phase 1.1）。
PyMuPDFのredactionでテキストを消した背景画像を敷き、その上に編集可能テキストボックスを重ねる方式。

## コマンド

- 環境構築: `uv venv --python 3.12 .venv && uv pip install --python .venv/bin/python -r requirements.txt`
- 実行: `.venv/bin/python convert.py input.pdf output.pptx [--dpi 150] [--debug-dir DIR]`
- Web公開時のハードタイムアウト実行: `.venv/bin/python worker.py input.pdf output.pptx --hard-timeout 120`
- テスト: `.venv/bin/python -m pytest -q`

## 構成

- `convert.py` … 変換本体（抽出 → redaction → 背景画像化 → PPTX組み立て）。ソフトタイムアウトのみ
- `worker.py` … convert.pyを別プロセスで実行しhard_timeout秒でSIGKILLするラッパー（Web公開向け）
- `tests/` … pytest。フィクスチャは `tests/fixtures_gen.py` が生成する合成PDFのみ（ネットワーク不使用）

## 守ること

- テストは合成PDFのみ。実PDF（顧客資料等）や生成物（*.pptx）をコミットしない。
- コード変更後は `.venv/bin/python -m pytest -q` を通す。
- Phase 1 の範囲: 横書き・可視のみ編集対象。縦書き・回転テキスト・回転ページ・不可視(OCR)テキスト・
  文字化けspan・それらと重なる横書き行・全面画像ページ(既定85%以上を画像が覆うページの可視テキスト)
  は背景に残す。スキャンPDFは背景のみ+警告。
- `--debug-dir` は空ディレクトリのみ許可（既存ファイルがあるとエラー）。検証用であり、機密PDFや
  本番では使わない（背景画像が平文PNGでディスクに残るため）。
- convert.py の `--timeout` はソフトタイムアウト（ページ処理の合間のみチェック）。Web公開等で
  確実に打ち切りたい場合は worker.py のハードタイムアウト(SIGKILL)を使う。
- convert.py と worker.py の両方にCLIオプションを追加する場合は `convert.build_arg_parser()` を
  worker.py 側も再利用しているため、二重メンテにならないよう共有パーサ側に追加する。
