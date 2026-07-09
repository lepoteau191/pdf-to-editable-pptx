"""app.py（Webアップロード MVP）のテスト。FastAPIのTestClientを使い、
実サーバーは起動せずリクエスト/レスポンスの往復を検証する。
"""

from __future__ import annotations

import io
import time
import uuid
import zipfile
from pathlib import Path

import pytest
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
    assert "疑似的な失敗" in data["message"]

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
