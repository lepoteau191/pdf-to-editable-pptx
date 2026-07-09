"""app.py（Webアップロード MVP）のテスト。FastAPIのTestClientを使い、
実サーバーは起動せずリクエスト/レスポンスの往復を検証する。
"""

from __future__ import annotations

import io
import json
import time
import uuid
import zipfile
from pathlib import Path

import pytest
from fastapi import HTTPException
from fastapi.testclient import TestClient

import app as app_module
import fixtures_gen as fx

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


@pytest.fixture
def client(tmp_path, monkeypatch):
    """各テストごとにジョブディレクトリを隔離する（実プロジェクトのjobs/を汚さない）。"""
    monkeypatch.setattr(app_module, "JOBS_ROOT", tmp_path / "jobs")
    return TestClient(app_module.app)


def _pdf_bytes(tmp_path: Path) -> bytes:
    pdf_path = tmp_path / "src.pdf"
    fx.make_business_pdf(pdf_path)
    return pdf_path.read_bytes()


def _wait_for_terminal(client: TestClient, job_id: str, timeout: float = 30.0) -> dict:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        res = client.get(f"/jobs/{job_id}")
        assert res.status_code == 200
        data = res.json()
        if data["status"] in ("done", "error"):
            return data
        time.sleep(0.1)
    pytest.fail(f"job {job_id} がタイムアウト内に終了しなかった: {data}")


# ---------------------------------------------------------------------------
# 簡易UI
# ---------------------------------------------------------------------------

def test_index_page_serves_html(client):
    res = client.get("/")
    assert res.status_code == 200
    assert "text/html" in res.headers["content-type"]
    assert "pdf2pptx" in res.text


# ---------------------------------------------------------------------------
# アップロード検証
# ---------------------------------------------------------------------------

def test_upload_rejects_non_pdf_extension(client):
    res = client.post(
        "/upload",
        files={"file": ("note.txt", b"hello world", "text/plain")},
    )
    assert res.status_code == 400
    assert not (app_module.JOBS_ROOT).exists() or not any(app_module.JOBS_ROOT.iterdir())


def test_upload_rejects_non_pdf_content(client):
    """拡張子は.pdfでも中身が%PDF-で始まらないものは拒否する。"""
    res = client.post(
        "/upload",
        files={"file": ("fake.pdf", b"not actually a pdf" * 10, "application/pdf")},
    )
    assert res.status_code == 400
    assert "PDFファイルとして認識できません" in res.json()["detail"]
    # 失敗したジョブディレクトリが残っていないこと
    assert not app_module.JOBS_ROOT.exists() or not any(app_module.JOBS_ROOT.iterdir())


def test_upload_rejects_oversized_file(client, monkeypatch, tmp_path):
    monkeypatch.setattr(app_module, "MAX_UPLOAD_BYTES", 100)  # 100バイトに引き下げ
    content = _pdf_bytes(tmp_path)
    assert len(content) > 100
    res = client.post(
        "/upload", files={"file": ("in.pdf", content, "application/pdf")}
    )
    assert res.status_code == 413
    assert not app_module.JOBS_ROOT.exists() or not any(app_module.JOBS_ROOT.iterdir())


def test_upload_missing_file_field_returns_422(client):
    res = client.post("/upload")
    assert res.status_code == 422


# ---------------------------------------------------------------------------
# ジョブ状態・パス検証
# ---------------------------------------------------------------------------

def test_get_job_unknown_uuid_404(client):
    res = client.get(f"/jobs/{uuid.uuid4()}")
    assert res.status_code == 404


def test_get_job_non_uuid_rejected_404(client):
    """job_idがUUID形式でない場合はパスを組み立てずに404にする（パストラバーサル対策）。"""
    res = client.get("/jobs/not-a-uuid-at-all")
    assert res.status_code == 404


def test_download_unknown_job_404(client):
    res = client.get(f"/jobs/{uuid.uuid4()}/download")
    assert res.status_code == 404


# ---------------------------------------------------------------------------
# 変換フロー全体（アップロード → 状態確認 → ダウンロード）
# ---------------------------------------------------------------------------

def test_full_upload_convert_download_flow(client, tmp_path):
    content = _pdf_bytes(tmp_path)
    res = client.post(
        "/upload", files={"file": ("business.pdf", content, "application/pdf")}
    )
    assert res.status_code == 202
    body = res.json()
    job_id = body["job_id"]
    assert body["status"] == "queued"
    uuid.UUID(job_id)  # 発行されたjob_idが正しいUUIDであること

    data = _wait_for_terminal(client, job_id)
    assert data["status"] == "done", data
    assert isinstance(data.get("warnings"), list)

    dl = client.get(f"/jobs/{job_id}/download")
    assert dl.status_code == 200
    assert dl.headers["content-type"] == PPTX_MIME
    assert "business.pptx" in dl.headers["content-disposition"]
    assert zipfile.is_zipfile(io.BytesIO(dl.content))


def test_download_before_completion_returns_409(client):
    """変換が終わる前（queued/processing）はダウンロードできず409になる。

    FastAPIのTestClientはBackgroundTasksをレスポンス返却前に同期実行するため、
    アップロードエンドポイント経由では「処理中」の状態を観測できない。
    そのためジョブディレクトリを直接構築し、download側のゲート判定を検証する。
    """
    job_id = str(uuid.uuid4())
    job_dir = app_module.JOBS_ROOT / job_id
    job_dir.mkdir(parents=True)
    app_module._write_status(
        job_dir,
        job_id=job_id,
        status="processing",
        original_filename="in.pdf",
        created_at="2026-01-01T00:00:00+00:00",
        message=None,
        warnings=[],
    )

    dl = client.get(f"/jobs/{job_id}/download")
    assert dl.status_code == 409


def test_conversion_failure_sets_error_status(client, tmp_path, monkeypatch):
    def _fail(*args, **kwargs):
        return 1, "", "エラー: 疑似的な失敗"

    monkeypatch.setattr(app_module.worker, "run_with_hard_timeout", _fail)

    content = _pdf_bytes(tmp_path)
    res = client.post(
        "/upload", files={"file": ("in.pdf", content, "application/pdf")}
    )
    job_id = res.json()["job_id"]

    data = _wait_for_terminal(client, job_id)
    assert data["status"] == "error"
    assert data["message"] == app_module.GENERIC_CONVERSION_ERROR_MESSAGE

    dl = client.get(f"/jobs/{job_id}/download")
    assert dl.status_code == 409


def test_debug_dir_is_never_forwarded_to_worker(client, tmp_path, monkeypatch):
    """Web経由の変換ではdebug_dirを常にNoneで渡す(要件10)。"""
    captured = {}

    def _spy(input_pdf, output_pptx, hard_timeout=None, **kwargs):
        captured.update(kwargs)
        return 1, "", "エラー: スパイ用の失敗"

    monkeypatch.setattr(app_module.worker, "run_with_hard_timeout", _spy)

    content = _pdf_bytes(tmp_path)
    res = client.post(
        "/upload", files={"file": ("in.pdf", content, "application/pdf")}
    )
    job_id = res.json()["job_id"]
    _wait_for_terminal(client, job_id)

    assert "debug_dir" in captured
    assert captured["debug_dir"] is None


# ---------------------------------------------------------------------------
# エラーメッセージのパス秘匿
# ---------------------------------------------------------------------------

def test_error_message_does_not_leak_internal_details(client, tmp_path, monkeypatch):
    """内部パスやスタックトレースがAPIレスポンスに出ないこと。"""
    leaky_detail = "/private/var/tmp/pdf2pptx_worker_x7z/output.pptx で権限エラー"

    def _fail(*args, **kwargs):
        return 1, "", f"エラー: {leaky_detail}"

    monkeypatch.setattr(app_module.worker, "run_with_hard_timeout", _fail)

    content = _pdf_bytes(tmp_path)
    res = client.post(
        "/upload", files={"file": ("in.pdf", content, "application/pdf")}
    )
    job_id = res.json()["job_id"]
    data = _wait_for_terminal(client, job_id)

    assert data["status"] == "error"
    assert data["message"] == app_module.GENERIC_CONVERSION_ERROR_MESSAGE
    assert "internal_error" not in data
    assert "/private/var/tmp" not in json.dumps(data)
    assert "pdf2pptx_worker_x7z" not in json.dumps(data)

    # ジョブディレクトリ内には内部詳細が記録されている(運用者向け)ことも確認する
    job_dir = app_module.JOBS_ROOT / job_id
    raw = app_module._read_status(job_dir)
    assert leaky_detail in raw["internal_error"]


def test_unexpected_exception_error_does_not_leak_details(client, tmp_path, monkeypatch):
    def _boom(*args, **kwargs):
        raise RuntimeError("/etc/passwd を読めませんでした")

    monkeypatch.setattr(app_module.worker, "run_with_hard_timeout", _boom)

    content = _pdf_bytes(tmp_path)
    res = client.post(
        "/upload", files={"file": ("in.pdf", content, "application/pdf")}
    )
    job_id = res.json()["job_id"]
    data = _wait_for_terminal(client, job_id)

    assert data["status"] == "error"
    assert data["message"] == app_module.GENERIC_CONVERSION_ERROR_MESSAGE
    assert "internal_error" not in data
    assert "/etc/passwd" not in json.dumps(data)


# ---------------------------------------------------------------------------
# ダウンロードAPIの安全化
# ---------------------------------------------------------------------------

def _make_done_job(client, tmp_path, job_id: str | None = None) -> str:
    """status.jsonだけを直接書いて「完了」状態のジョブディレクトリを作る。

    output.pptxの内容/種類はテストごとに個別に用意するため、ここでは
    status.jsonの作成とディレクトリ準備だけを行う。
    """
    job_id = job_id or str(uuid.uuid4())
    job_dir = app_module.JOBS_ROOT / job_id
    job_dir.mkdir(parents=True)
    app_module._write_status(
        job_dir,
        job_id=job_id,
        status="done",
        original_filename="in.pdf",
        created_at="2026-01-01T00:00:00+00:00",
        message=None,
        warnings=[],
    )
    return job_id


def test_download_rejects_symlink_output(client, tmp_path):
    job_id = _make_done_job(client, tmp_path)
    job_dir = app_module.JOBS_ROOT / job_id

    real_target = tmp_path / "elsewhere.pptx"
    with zipfile.ZipFile(real_target, "w") as zf:
        zf.writestr("dummy.txt", "hello")
    (job_dir / "output.pptx").symlink_to(real_target)

    dl = client.get(f"/jobs/{job_id}/download")
    assert dl.status_code == 500
    assert "symlink" not in dl.json()["detail"]  # 詳細はログのみ、レスポンスは固定文


def test_download_rejects_output_outside_job_dir(tmp_path):
    """resolve()結果がjob_dirの配下でない場合は拒否する。

    通常のdownload_jobエンドポイントではjob_dir配下のoutput.pptxしか渡らないが
    （symlinkでない限りresolve結果がjob_dir外になることはない）、この検査自体を
    独立して確認するため _validate_output_for_download を直接呼び出す。
    """
    job_dir = tmp_path / "job"
    job_dir.mkdir()
    outside_file = tmp_path / "outside.pptx"
    with zipfile.ZipFile(outside_file, "w") as zf:
        zf.writestr("dummy.txt", "hello")

    with pytest.raises(HTTPException) as exc_info:
        app_module._validate_output_for_download(job_dir, outside_file)
    assert exc_info.value.status_code == 500


def test_download_rejects_invalid_zip(client, tmp_path):
    job_id = _make_done_job(client, tmp_path)
    job_dir = app_module.JOBS_ROOT / job_id
    (job_dir / "output.pptx").write_bytes(b"this is not a valid pptx/zip")

    dl = client.get(f"/jobs/{job_id}/download")
    assert dl.status_code == 500


def test_download_succeeds_for_valid_output(client, tmp_path):
    """正常系: symlinkでもなくjob_dir内の正しいzipなら通常通りダウンロードできる。"""
    job_id = _make_done_job(client, tmp_path)
    job_dir = app_module.JOBS_ROOT / job_id
    with zipfile.ZipFile(job_dir / "output.pptx", "w") as zf:
        zf.writestr("dummy.txt", "hello")

    dl = client.get(f"/jobs/{job_id}/download")
    assert dl.status_code == 200


# ---------------------------------------------------------------------------
# ファイル名長さ制限
# ---------------------------------------------------------------------------

def test_long_original_filename_is_truncated_on_upload(client, tmp_path):
    long_stem = "あ" * 400
    filename = f"{long_stem}.pdf"
    content = _pdf_bytes(tmp_path)

    res = client.post(
        "/upload", files={"file": (filename, content, "application/pdf")}
    )
    assert res.status_code == 202
    job_id = res.json()["job_id"]

    status = client.get(f"/jobs/{job_id}").json()
    assert len(status["original_filename"]) <= app_module.MAX_ORIGINAL_FILENAME_LENGTH
    assert status["original_filename"].endswith(".pdf")


def test_download_filename_stem_is_truncated(client):
    huge_name = "x" * 500 + ".pdf"
    truncated = app_module._safe_download_name(huge_name)
    stem = truncated.removesuffix(".pptx")
    assert len(stem) <= app_module.MAX_DOWNLOAD_STEM_LENGTH
    assert truncated.endswith(".pptx")


def test_truncate_original_filename_preserves_extension():
    name = "y" * 300 + ".pdf"
    truncated = app_module._truncate_original_filename(name)
    assert len(truncated) <= app_module.MAX_ORIGINAL_FILENAME_LENGTH
    assert truncated.endswith(".pdf")
