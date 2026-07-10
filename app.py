#!/usr/bin/env python3
"""pdf2pptx のWebアップロードMVP（ローカル専用）。

PDFをブラウザからアップロードし、worker.py（ハードタイムアウト付き別プロセス）
経由でPPTXに変換し、完了後にダウンロードできるようにするFastAPIアプリ。

エンドポイント:
  GET  /                       アップロード用の簡易HTML画面
  POST /upload                 PDFをアップロードしてジョブを開始する
  GET  /jobs/{job_id}          ジョブの状態(queued/processing/done/error)を返す
  GET  /jobs/{job_id}/download 完了したジョブのPPTXをダウンロードする

設計上の注意（README「Web公開時のリソース隔離方針」も参照）:
  - ユーザー指定パスは一切受け取らない。保存先は常にサーバーが生成する
    UUIDジョブディレクトリ配下の固定名（input.pdf / output.pptx）。
  - job_id はURLパスから来るため、ファイルパスに使う前に必ずUUIDとして
    検証する（パストラバーサル対策）。
  - debug_dir は常に None（Web経由では絶対に有効化しない）。
  - 変換は worker.run_with_hard_timeout() 経由でのみ行い、convert.py を
    直接同一プロセスで呼び出さない（ハードタイムアウトを必ず効かせるため）。
  - これはローカル専用MVP。認証・決済・S3保存・同時実行数の制限は無い。
    本番公開する場合はコンテナ/cgroupでのRAM・CPU・ディスク制限が必須
    （README参照）。
"""

from __future__ import annotations

import json
import os
import re
import shutil
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

import convert
import worker

APP_ROOT = Path(__file__).resolve().parent
JOBS_ROOT = APP_ROOT / "jobs"
STATIC_DIR = APP_ROOT / "static"

ALLOWED_EXTENSION = ".pdf"
MAX_UPLOAD_SIZE_MB = convert.DEFAULT_MAX_FILE_SIZE_MB
MAX_UPLOAD_BYTES = int(MAX_UPLOAD_SIZE_MB * 1024 * 1024)
UPLOAD_CHUNK_SIZE = 1024 * 1024
HARD_TIMEOUT_SECONDS = worker.DEFAULT_HARD_TIMEOUT

# Web UI向けの画質プリセット。スキャンPDFはページを背景画像としてPPTXに
# 貼るため、150dpiでは読みにくいケースがある。ローカルPC利用ではまず
# 250dpiを標準にし、細かい文字のPDFだけ300dpiを選べるようにする。
# 将来OCRを入れる場合も、このdpiで生成したページ画像をOCR入力に使う。
WEB_QUALITY_PRESETS = {
    "standard": 250,
    "high": 300,
}
DEFAULT_WEB_QUALITY = "standard"
WEB_OCR_LANG = convert.DEFAULT_OCR_LANG

# 保存するファイル名・提示するダウンロード名の長さ上限。
# Content-Dispositionヘッダの肥大化や、極端に長いファイル名によるOS依存の
# 問題を避けるため、表示用の値は常にこの範囲に切り詰める。
MAX_ORIGINAL_FILENAME_LENGTH = 255
MAX_DOWNLOAD_STEM_LENGTH = 180

# 変換失敗時にAPI/UIへ返す固定の安全なメッセージ。実際のエラー詳細
# （worker.pyの標準エラー出力。一時ディレクトリ等の内部パスを含みうる）は
# status.jsonの internal_error にのみ記録し、APIレスポンスには含めない。
GENERIC_CONVERSION_ERROR_MESSAGE = (
    "変換に失敗しました。ファイルを確認するか、時間をおいてもう一度お試しください。"
)

PPTX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)
STATUS_FILENAME = "status.json"
JOB_STATUSES = ("queued", "processing", "done", "error")

# status.json内にあってもAPIレスポンスには含めないフィールド
# （internal_errorには内部パス等を含みうる生のエラー詳細が入るため）。
_INTERNAL_ONLY_FIELDS = {"internal_error"}

app = FastAPI(title="pdf2pptx Web MVP (ローカル専用)")


# ---------------------------------------------------------------------------
# ジョブディレクトリ・状態ファイルの管理
# ---------------------------------------------------------------------------

def _job_dir_for(job_id: str) -> Path:
    """job_idを検証したうえでジョブディレクトリのパスを返す。

    job_idはURLパスパラメータ由来でありユーザーが自由な文字列を渡せるため、
    パストラバーサル対策として厳密にUUID形式であることを確認してから
    初めてパスを組み立てる（正規化されたUUID文字列のみを使う）。
    """
    try:
        parsed = uuid.UUID(job_id)
    except ValueError:
        raise HTTPException(status_code=404, detail="ジョブが見つかりません")
    return JOBS_ROOT / str(parsed)


def _status_path(job_dir: Path) -> Path:
    return job_dir / STATUS_FILENAME


def _read_status(job_dir: Path) -> dict:
    path = _status_path(job_dir)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        raise HTTPException(status_code=404, detail="ジョブが見つかりません")


def _public_status(data: dict) -> dict:
    """APIレスポンスとして返して良いフィールドだけに絞る。

    internal_error（worker.pyの生の標準エラー出力。一時ディレクトリ等の
    内部パスを含みうる）は絶対に外部へ返さない。
    """
    return {k: v for k, v in data.items() if k not in _INTERNAL_ONLY_FIELDS}


def _write_status(job_dir: Path, **updates) -> None:
    """status.jsonを既存内容とマージしてアトミックに書き込む。

    同じディレクトリ内で一時ファイルに書いてから os.replace() するため、
    GETリクエストが読み取り途中の不完全なJSONを見ることはない。
    """
    path = _status_path(job_dir)
    current: dict = {}
    if path.exists():
        try:
            current = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            current = {}
    current.update(updates)
    current["updated_at"] = datetime.now(timezone.utc).isoformat()

    tmp_path = path.with_name(path.name + ".tmp")
    tmp_path.write_text(
        json.dumps(current, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    os.replace(tmp_path, path)


_SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9._\-ぁ-んァ-ヶ一-龥ー]")


def _truncate_original_filename(name: str) -> str:
    """アップロード元のファイル名を保存用に切り詰める（拡張子は保持する）。

    Content-Dispositionヘッダの肥大化や、極端に長い値がstatus.jsonや
    ログに残り続けることを避けるための上限。
    """
    if len(name) <= MAX_ORIGINAL_FILENAME_LENGTH:
        return name
    p = Path(name)
    suffix = p.suffix
    stem_budget = max(MAX_ORIGINAL_FILENAME_LENGTH - len(suffix), 1)
    return p.stem[:stem_budget] + suffix


def _safe_download_name(original_filename: str | None) -> str:
    """ダウンロード時に提示するファイル名を作る（表示用のみ・パスには使わない）。"""
    stem = Path(original_filename or "output").stem
    stem = _SAFE_NAME_RE.sub("_", stem).strip("._") or "output"
    stem = stem[:MAX_DOWNLOAD_STEM_LENGTH].strip("._") or "output"
    return f"{stem}.pptx"


def _quality_to_dpi(quality: str) -> int:
    """Web UIの画質指定を安全なdpi値へ変換する。

    利用者から任意のdpi数値を受け取ると、極端な値でCPU/RAM/出力サイズを
    膨らませられるため、Web経由では固定プリセットだけを許可する。
    """
    try:
        return WEB_QUALITY_PRESETS[quality]
    except KeyError:
        allowed = ", ".join(sorted(WEB_QUALITY_PRESETS))
        raise HTTPException(
            status_code=400,
            detail=f"画質指定が不正です（指定可能: {allowed}）",
        )


# ---------------------------------------------------------------------------
# アップロード
# ---------------------------------------------------------------------------

async def _save_upload_streaming(file: UploadFile, dest: Path) -> None:
    """アップロードをストリーミングで保存し、サイズ上限をその場で強制する。

    一括読み込みだと上限を超える巨大ファイルが一度メモリ/ディスクに
    展開されてしまうため、チャンク単位で読みながら合計サイズを監視する。
    """
    total = 0
    with open(dest, "wb") as out:
        while True:
            chunk = await file.read(UPLOAD_CHUNK_SIZE)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_UPLOAD_BYTES:
                raise HTTPException(
                    status_code=413,
                    detail=f"ファイルが大きすぎます（上限 {MAX_UPLOAD_SIZE_MB:.0f}MB）",
                )
            out.write(chunk)


@app.post("/upload", status_code=202)
async def upload_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    quality: str = Form(DEFAULT_WEB_QUALITY),
    ocr: bool = Form(False),
):
    dpi = _quality_to_dpi(quality)
    if ocr:
        try:
            convert.check_ocr_available(convert.OCR_ENGINE_TESSERACT, WEB_OCR_LANG)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))
    original_name = file.filename or "upload.pdf"
    # 拡張子の判定は元のファイル名で行う（切り詰めた後だと拡張子を
    # 巻き込んでしまい、長いファイル名の正当なPDFを誤って拒否しうるため）。
    if Path(original_name).suffix.lower() != ALLOWED_EXTENSION:
        raise HTTPException(
            status_code=400, detail="PDFファイル(.pdf)のみアップロードできます"
        )
    original_name = _truncate_original_filename(original_name)

    job_id = str(uuid.uuid4())
    job_dir = JOBS_ROOT / job_id
    job_dir.mkdir(parents=True)
    input_path = job_dir / "input.pdf"

    try:
        await _save_upload_streaming(file, input_path)
        if not convert.looks_like_pdf(input_path):
            raise HTTPException(
                status_code=400,
                detail="PDFファイルとして認識できませんでした（%PDFヘッダが見つかりません）",
            )
    except HTTPException:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    finally:
        await file.close()

    _write_status(
        job_dir,
        job_id=job_id,
        status="queued",
        original_filename=original_name,
        quality=quality,
        dpi=dpi,
        ocr=ocr,
        ocr_lang=WEB_OCR_LANG if ocr else None,
        created_at=datetime.now(timezone.utc).isoformat(),
        message=None,
        warnings=[],
    )

    background_tasks.add_task(_run_conversion_job, job_id, dpi, ocr)
    return {
        "job_id": job_id,
        "status": "queued",
        "quality": quality,
        "dpi": dpi,
        "ocr": ocr,
    }


# ---------------------------------------------------------------------------
# 変換ジョブ本体（バックグラウンドで実行）
# ---------------------------------------------------------------------------

def _run_conversion_job(
    job_id: str,
    dpi: int | None = None,
    ocr: bool | None = None,
) -> None:
    job_dir = JOBS_ROOT / job_id
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "output.pptx"

    if dpi is None:
        status = _read_status(job_dir)
        dpi = int(status.get("dpi", WEB_QUALITY_PRESETS[DEFAULT_WEB_QUALITY]))
    if ocr is None:
        status = _read_status(job_dir)
        ocr = bool(status.get("ocr", False))

    _write_status(job_dir, status="processing")

    try:
        returncode, stdout, stderr = worker.run_with_hard_timeout(
            input_path,
            output_path,
            hard_timeout=HARD_TIMEOUT_SECONDS,
            debug_dir=None,  # Web経由では常に無効
            dpi=dpi,
            ocr=convert.OCR_ENGINE_TESSERACT if ocr else convert.OCR_ENGINE_OFF,
            ocr_lang=WEB_OCR_LANG,
        )
    except Exception as e:  # noqa: BLE001 - ジョブを必ずerror状態へ遷移させる
        _write_status(
            job_dir,
            status="error",
            message=GENERIC_CONVERSION_ERROR_MESSAGE,
            internal_error=f"{type(e).__name__}: {e}",
            warnings=[],
        )
        return

    # 警告(convert.pyが安全に設計した固定形式の文言。パスは含まない)は
    # そのまま利用者に見せてよい。一方でエラー詳細(stderr)は一時ディレクトリ
    # 等の内部パスを含みうるため、internal_errorにのみ記録しAPIには出さない。
    warning_lines = [
        line[len("警告: "):] for line in stderr.splitlines() if line.startswith("警告: ")
    ]

    if returncode == 0:
        _write_status(job_dir, status="done", message=None, warnings=warning_lines)
    else:
        error_lines = [line for line in stderr.splitlines() if line.startswith("エラー:")]
        internal_detail = error_lines[-1] if error_lines else (stderr.strip() or "変換に失敗しました")
        _write_status(
            job_dir,
            status="error",
            message=GENERIC_CONVERSION_ERROR_MESSAGE,
            internal_error=internal_detail,
            warnings=warning_lines,
        )


# ---------------------------------------------------------------------------
# ジョブ状態・ダウンロード
# ---------------------------------------------------------------------------

@app.get("/jobs/{job_id}")
def get_job(job_id: str):
    job_dir = _job_dir_for(job_id)
    if not job_dir.exists():
        raise HTTPException(status_code=404, detail="ジョブが見つかりません")
    return _public_status(_read_status(job_dir))


def _reject_unsafe_download(reason: str) -> None:
    """出力ファイルの安全性チェックに失敗した場合の共通エラー。

    reason（内部パスやsymlinkの指す先を含みうる）はサーバーログにのみ出し、
    APIレスポンスには含めない。呼び出し元からは常に例外として抜ける。
    """
    print(f"[download] 出力ファイルの検証に失敗しました: {reason}")
    raise HTTPException(
        status_code=500,
        detail="出力ファイルを確認できませんでした。時間をおいて再度お試しください。",
    )


def _validate_output_for_download(job_dir: Path, output_path: Path) -> Path:
    """ダウンロード対象ファイルの安全性を検査し、resolve済みの実パスを返す。

    以下はworker.pyの正常なアトミック置換(os.replace)では起こらないはずだが、
    万一の細工・競合・バグに備えた防御的チェック。ダウンロード直前に必ず行う。
      - symlinkでないこと
      - 実在すること
      - resolve()した実パスがjob_dirの配下であること
      - zipとして開けること（pptxはzip形式のため）
    """
    if output_path.is_symlink():
        _reject_unsafe_download(f"output.pptxがsymlinkです: {output_path}")
    if not output_path.exists():
        _reject_unsafe_download(f"出力ファイルが存在しません: {output_path}")

    try:
        resolved_output = output_path.resolve(strict=True)
        resolved_job_dir = job_dir.resolve(strict=True)
    except OSError:
        _reject_unsafe_download(f"出力ファイルのresolveに失敗しました: {output_path}")

    if not resolved_output.is_relative_to(resolved_job_dir):
        _reject_unsafe_download(
            f"出力ファイルがジョブディレクトリ外を指しています: {resolved_output}"
        )
    if not zipfile.is_zipfile(resolved_output):
        _reject_unsafe_download(f"出力ファイルがzipとして不正です: {resolved_output}")

    return resolved_output


@app.get("/jobs/{job_id}/download")
def download_job(job_id: str):
    job_dir = _job_dir_for(job_id)
    if not job_dir.exists():
        raise HTTPException(status_code=404, detail="ジョブが見つかりません")

    status = _read_status(job_dir)
    if status.get("status") != "done":
        raise HTTPException(
            status_code=409,
            detail=f"まだダウンロードできません（状態: {status.get('status')}）",
        )

    resolved_output = _validate_output_for_download(job_dir, job_dir / "output.pptx")

    return FileResponse(
        resolved_output,
        media_type=PPTX_MEDIA_TYPE,
        filename=_safe_download_name(status.get("original_filename")),
    )


# ---------------------------------------------------------------------------
# 簡易UI
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")
