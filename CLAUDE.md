# pdf2pptx

PDFを編集可能なPPTXに変換するCLIツール（自社ツール・Phase 1.2）。
PyMuPDFのredactionでテキストを消した背景画像を敷き、その上に編集可能テキストボックスを重ねる方式。
CLIに加えて、ブラウザアップロード用のローカル専用FastAPI MVP（app.py）がある。

## コマンド

- 環境構築: `uv venv --python 3.12 .venv && uv pip install --python .venv/bin/python -r requirements.txt`
- CLI実行: `.venv/bin/python convert.py input.pdf output.pptx [--dpi 150] [--debug-dir DIR]`
- OCR付きCLI実行（Tesseract導入済みPCのみ）: `.venv/bin/python convert.py input.pdf output.pptx --dpi 300 --ocr tesseract`
- Web公開時のハードタイムアウト実行: `.venv/bin/python worker.py input.pdf output.pptx --hard-timeout 120`
- Webアプリ起動（ローカル）: `.venv/bin/uvicorn app:app --host 127.0.0.1 --port 8731 --reload`
- ジョブディレクトリ掃除: `.venv/bin/python janitor.py --max-age-hours 24`
- テスト: `.venv/bin/python -m pytest -q`

## 構成

- `convert.py` … 変換本体（抽出 → redaction → 背景画像化 → PPTX組み立て）。ソフトタイムアウトのみ
- `worker.py` … convert.pyを別プロセスで実行しhard_timeout秒でSIGKILLするラッパー（Web公開向け）
- `app.py` … ブラウザアップロード用FastAPI MVP（ローカル専用。認証・S3等は未実装）
- `static/index.html` … app.pyが配信する簡易アップロード画面
- `janitor.py` … `jobs/<job_id>/` の古いジョブディレクトリを削除する掃除スクリプト
- `tests/` … pytest。フィクスチャは `tests/fixtures_gen.py` が生成する合成PDFのみ（ネットワーク不使用）
- `.github/workflows/tests.yml` … push/PR時にpytestをUbuntu(Python 3.10〜3.12)で自動実行するCI

## 守ること

- テストは合成PDFのみ。実PDF（顧客資料等）や生成物（*.pptx）をコミットしない。
- コード変更後は `.venv/bin/python -m pytest -q` を通す。CIはUbuntu上でも動くので、
  macOS専用パスをテストフィクスチャに直書きしない（`tests/fixtures_gen.py`の
  `_resolve_jp_font()`のように候補パス+`fc-match`フォールバックで解決する）。
- Phase 1 の範囲: 横書き・可視のみ編集対象。縦書き・回転テキスト・回転ページ・不可視(OCR)テキスト・
  文字化けspan・それらと重なる横書き行・全面画像ページ(既定85%以上を画像が覆うページの可視テキスト)
  は背景に残す。スキャンPDFは背景のみ+警告。
- OCRはβの任意機能。`--ocr tesseract` またはWeb UIの「OCRを試す（β）」が明示された場合だけ、
  背景のみページへTesseract OCR結果を編集可能テキストとして重ねる。通常変換ではOCRしない。
  Tesseract本体と `jpn`/`eng` 言語データが無い環境ではOCRだけ失敗させ、通常変換には影響させない。
  PDF画像・OCR入力画像・OCR結果は外部送信しない。
- `--debug-dir` は空ディレクトリのみ許可（既存ファイルがあるとエラー）。検証用であり、機密PDFや
  本番では使わない（背景画像が平文PNGでディスクに残るため）。
- convert.py の `--timeout` はソフトタイムアウト（ページ処理の合間のみチェック）。Web公開等で
  確実に打ち切りたい場合は worker.py のハードタイムアウト(プロセスグループごとSIGKILL)を使う。
- convert.py と worker.py の両方にCLIオプションを追加する場合は `convert.build_arg_parser()` を
  worker.py 側も再利用しているため、二重メンテにならないよう共有パーサ側に追加する。
- 入出力パス同一チェック(`convert.check_distinct_input_output`)は convert.py と worker.py の
  **両方**で呼ぶ。worker.pyは子プロセスに一時パスを渡すため、convert.py側の検査だけでは
  worker.py経由の実行（実際の出力先への `.part`→置換）を保護できない。
- worker.pyをWebから使う場合は、ハードタイムアウトだけでは不十分。README の
  「Web公開時のリソース隔離方針」（コンテナ/cgroupでのRAM・CPU・ディスク上限、非特権・
  ネットワークなしワーカー、ユーザー指定パス禁止、janitorによる一時ファイル掃除）に従うこと。
- app.pyはローカル専用MVP。認証・決済・S3保存・同時実行数制限・本番公開は未実装で、
  意図的に後回しにしている（要件を先取りして実装しない）。
- app.pyの変換は必ず `worker.run_with_hard_timeout()` 経由で行い、`convert.py`を
  直接同一プロセスで呼ばない。`debug_dir`は常に`None`固定（Web経由で有効化する手段を作らない）。
- ジョブディレクトリ(`jobs/<job_id>/`)のパスはサーバー側で生成するUUIDのみに固定する。
  URLパスから来る`job_id`は、ファイルパスに使う前に必ず`uuid.UUID()`で検証する
  （パストラバーサル対策。app.pyの`_job_dir_for()`参照）。
- `jobs/`はgitignore対象（アップロードPDF・生成PPTXは秘密情報になりうる）。コミットしない。
- ダウンロード直前は必ず `app._validate_output_for_download()` を通す（symlink拒否・
  resolve()がjob_dir配下か確認・zipfile.is_zipfileで検証）。output.pptxを直接
  `FileResponse`に渡さない。
- 変換失敗時、`status.json`の`message`は常に固定文（`app.GENERIC_CONVERSION_ERROR_MESSAGE`）。
  worker.pyの生の標準エラー出力（内部パスを含みうる）は`internal_error`にのみ記録し、
  `app._public_status()`でAPIレスポンスから必ず除外する。新しくエラー系のフィールドを
  追加するときも、内部詳細を安易にAPIへ直接返さないこと。
- アップロード元ファイル名・ダウンロード提示名は必ず長さ上限
  （`MAX_ORIGINAL_FILENAME_LENGTH` / `MAX_DOWNLOAD_STEM_LENGTH`）で切り詰める
  （Content-Dispositionの肥大化防止）。
- janitor.pyはUUID形式でないディレクトリ・status.json内job_id不一致のディレクトリを
  skipし、queued/processingは`--stale-timeout-hours`を超えるまで削除しない。
  `shutil.rmtree`の失敗は`ignore_errors=True`で握り潰さず警告として返す。
- 公開用ジョブキュー（Celery/RQ/Arq等）・worker完全分離・rate limit・認証・
  同時実行数制限・jobs総容量制限はPhase 0.3以降。今のapp.pyのBackgroundTasksは
  ローカルMVP専用で、それらの代替にはならない。
