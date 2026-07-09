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
from datetime import datetime, timezone
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, HTTPException, UploadFile
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

PPTX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)
STATUS_FILENAME = "status.json"
JOB_STATUSES = ("queued", "processing", "done", "error")

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


def _safe_download_name(original_filename: str | None) -> str:
    """ダウンロード時に提示するファイル名を作る（表示用のみ・パスには使わない）。"""
    stem = Path(original_filename or "output").stem
    stem = _SAFE_NAME_RE.sub("_", stem).strip("._") or "output"
    return f"{stem}.pptx"


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
async def upload_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    original_name = file.filename or "upload.pdf"
    if Path(original_name).suffix.lower() != ALLOWED_EXTENSION:
        raise HTTPException(
            status_code=400, detail="PDFファイル(.pdf)のみアップロードできます"
        )

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
        created_at=datetime.now(timezone.utc).isoformat(),
        message=None,
        warnings=[],
    )

    background_tasks.add_task(_run_conversion_job, job_id)
    return {"job_id": job_id, "status": "queued"}


# ---------------------------------------------------------------------------
# 変換ジョブ本体（バックグラウンドで実行）
# ---------------------------------------------------------------------------

def _run_conversion_job(job_id: str) -> None:
    job_dir = JOBS_ROOT / job_id
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "output.pptx"

    _write_status(job_dir, status="processing")

    try:
        returncode, stdout, stderr = worker.run_with_hard_timeout(
            input_path,
            output_path,
            hard_timeout=HARD_TIMEOUT_SECONDS,
            debug_dir=None,  # Web経由では常に無効
        )
    except Exception as e:  # noqa: BLE001 - ジョブを必ずerror状態へ遷移させる
        _write_status(job_dir, status="error", message=f"内部エラー: {e}", warnings=[])
        return

    warning_lines = [
        line[len("警告: "):] for line in stderr.splitlines() if line.startswith("警告: ")
    ]

    if returncode == 0:
        _write_status(job_dir, status="done", message=None, warnings=warning_lines)
    else:
        error_lines = [line for line in stderr.splitlines() if line.startswith("エラー:")]
        message = error_lines[-1] if error_lines else (stderr.strip() or "変換に失敗しました")
        _write_status(job_dir, status="error", message=message, warnings=warning_lines)


# ---------------------------------------------------------------------------
# ジョブ状態・ダウンロード
# ---------------------------------------------------------------------------

@app.get("/jobs/{job_id}")
def get_job(job_id: str):
    job_dir = _job_dir_for(job_id)
    if not job_dir.exists():
        raise HTTPException(status_code=404, detail="ジョブが見つかりません")
    return _read_status(job_dir)


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

    output_path = job_dir / "output.pptx"
    if not output_path.exists():
        raise HTTPException(status_code=500, detail="出力ファイルが見つかりません")

    return FileResponse(
        output_path,
        media_type=PPTX_MEDIA_TYPE,
        filename=_safe_download_name(status.get("original_filename")),
    )


# ---------------------------------------------------------------------------
# 簡易UI
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")
