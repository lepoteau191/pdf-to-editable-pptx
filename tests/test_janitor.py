"""janitor.py（ジョブディレクトリ掃除スクリプト）のテスト。"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

import janitor

ROOT = Path(__file__).resolve().parents[1]


def _make_job(
    jobs_root: Path,
    job_id: str,
    created_at: datetime | None,
    status: str = "done",
    status_job_id: str | None = "__same__",
) -> Path:
    """テスト用のジョブディレクトリを作る。

    status_job_id="__same__" ならディレクトリ名と同じjob_idを書く（既定・正常系）。
    None を渡すとstatus.json自体を作らない。それ以外の文字列を渡すと、
    ディレクトリ名とは異なるjob_idを書く（不一致ケースの検証用）。
    """
    job_dir = jobs_root / job_id
    job_dir.mkdir(parents=True)
    if status_job_id is None:
        return job_dir
    recorded_job_id = job_id if status_job_id == "__same__" else status_job_id
    data = {"job_id": recorded_job_id, "status": status}
    if created_at is not None:
        data["created_at"] = created_at.isoformat()
    (job_dir / "status.json").write_text(json.dumps(data), encoding="utf-8")
    return job_dir


def _uuid() -> str:
    return str(uuid.uuid4())


# ---------------------------------------------------------------------------
# 通常の掃除（done/error）
# ---------------------------------------------------------------------------

def test_cleanup_removes_only_old_done_jobs(tmp_path):
    jobs_root = tmp_path / "jobs"
    now = datetime.now(timezone.utc)
    old_id, new_id = _uuid(), _uuid()
    old_job = _make_job(jobs_root, old_id, now - timedelta(hours=48), status="done")
    new_job = _make_job(jobs_root, new_id, now - timedelta(minutes=5), status="done")

    removed, warnings = janitor.cleanup_old_jobs(jobs_root, max_age_hours=24)

    assert removed == [old_id]
    assert warnings == []
    assert not old_job.exists()
    assert new_job.exists()


def test_cleanup_falls_back_to_mtime_without_status_json(tmp_path):
    jobs_root = tmp_path / "jobs"
    job_id = _uuid()
    job_dir = _make_job(jobs_root, job_id, created_at=None, status_job_id=None)
    old_time = (datetime.now(timezone.utc) - timedelta(hours=48)).timestamp()
    os.utime(job_dir, (old_time, old_time))

    removed, warnings = janitor.cleanup_old_jobs(jobs_root, max_age_hours=24)

    assert removed == [job_id]
    assert warnings == []
    assert not job_dir.exists()


def test_cleanup_missing_jobs_root_returns_empty(tmp_path):
    assert janitor.cleanup_old_jobs(tmp_path / "nope", max_age_hours=24) == ([], [])


def test_cleanup_ignores_non_directory_entries(tmp_path):
    jobs_root = tmp_path / "jobs"
    jobs_root.mkdir()
    (jobs_root / "stray_file.txt").write_text("not a job dir")
    removed, warnings = janitor.cleanup_old_jobs(jobs_root, max_age_hours=0)
    assert removed == []
    assert warnings == []
    assert (jobs_root / "stray_file.txt").exists()


# ---------------------------------------------------------------------------
# queued/processing の特別扱い
# ---------------------------------------------------------------------------

def test_cleanup_keeps_processing_job_within_stale_timeout(tmp_path):
    """processingのジョブはmax_age_hoursを超えていてもstale_timeout内なら残す。"""
    jobs_root = tmp_path / "jobs"
    job_id = _uuid()
    now = datetime.now(timezone.utc)
    job_dir = _make_job(jobs_root, job_id, now - timedelta(hours=30), status="processing")

    removed, warnings = janitor.cleanup_old_jobs(
        jobs_root, max_age_hours=24, stale_timeout_hours=48
    )

    assert removed == []
    assert warnings == []
    assert job_dir.exists()


def test_cleanup_removes_stale_processing_job(tmp_path):
    """processingでもstale_timeout_hoursを超えていれば削除する（放置ジョブの掃除）。"""
    jobs_root = tmp_path / "jobs"
    job_id = _uuid()
    now = datetime.now(timezone.utc)
    job_dir = _make_job(jobs_root, job_id, now - timedelta(hours=72), status="processing")

    removed, warnings = janitor.cleanup_old_jobs(
        jobs_root, max_age_hours=24, stale_timeout_hours=48
    )

    assert removed == [job_id]
    assert warnings == []
    assert not job_dir.exists()


def test_cleanup_queued_job_uses_stale_timeout_too(tmp_path):
    jobs_root = tmp_path / "jobs"
    job_id = _uuid()
    now = datetime.now(timezone.utc)
    job_dir = _make_job(jobs_root, job_id, now - timedelta(hours=30), status="queued")

    removed, _ = janitor.cleanup_old_jobs(jobs_root, max_age_hours=24, stale_timeout_hours=48)
    assert removed == []
    assert job_dir.exists()


# ---------------------------------------------------------------------------
# 不正・不一致エントリのskip
# ---------------------------------------------------------------------------

def test_cleanup_skips_non_uuid_directory(tmp_path):
    jobs_root = tmp_path / "jobs"
    now = datetime.now(timezone.utc)
    job_dir = _make_job(jobs_root, "not-a-uuid", now - timedelta(hours=100), status="done")

    removed, warnings = janitor.cleanup_old_jobs(jobs_root, max_age_hours=24)

    assert removed == []
    assert warnings == []
    assert job_dir.exists()


def test_cleanup_skips_job_id_mismatch(tmp_path):
    """status.json内のjob_idがディレクトリ名と食い違う場合は削除しない。"""
    jobs_root = tmp_path / "jobs"
    job_id = _uuid()
    now = datetime.now(timezone.utc)
    job_dir = _make_job(
        jobs_root, job_id, now - timedelta(hours=100), status="done",
        status_job_id=_uuid(),  # ディレクトリ名とは異なるjob_id
    )

    removed, warnings = janitor.cleanup_old_jobs(jobs_root, max_age_hours=24)

    assert removed == []
    assert warnings == []
    assert job_dir.exists()


# ---------------------------------------------------------------------------
# 削除失敗の可視化
# ---------------------------------------------------------------------------

def test_cleanup_reports_rmtree_failure_as_warning(tmp_path, monkeypatch):
    jobs_root = tmp_path / "jobs"
    job_id = _uuid()
    now = datetime.now(timezone.utc)
    job_dir = _make_job(jobs_root, job_id, now - timedelta(hours=48), status="done")

    def _boom(path):
        raise OSError("permission denied (simulated)")

    monkeypatch.setattr(janitor.shutil, "rmtree", _boom)

    removed, warnings = janitor.cleanup_old_jobs(jobs_root, max_age_hours=24)

    assert removed == []
    assert len(warnings) == 1
    assert job_id in warnings[0]
    assert job_dir.exists()  # 削除に失敗したので残っている


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def test_janitor_cli_smoke(tmp_path):
    jobs_root = tmp_path / "jobs"
    now = datetime.now(timezone.utc)
    old_id, new_id = _uuid(), _uuid()
    _make_job(jobs_root, old_id, now - timedelta(hours=48), status="done")
    _make_job(jobs_root, new_id, now - timedelta(minutes=5), status="done")

    res = subprocess.run(
        [sys.executable, str(ROOT / "janitor.py"), "--jobs-root", str(jobs_root),
         "--max-age-hours", "24"],
        capture_output=True, text=True,
    )
    assert res.returncode == 0, res.stderr
    assert old_id in res.stdout
    assert not (jobs_root / old_id).exists()
    assert (jobs_root / new_id).exists()


def test_janitor_cli_nonzero_exit_on_warning(tmp_path, monkeypatch):
    jobs_root = tmp_path / "jobs"
    now = datetime.now(timezone.utc)
    job_id = _uuid()
    _make_job(jobs_root, job_id, now - timedelta(hours=48), status="done")

    def _boom(path):
        raise OSError("permission denied (simulated)")

    monkeypatch.setattr(janitor.shutil, "rmtree", _boom)
    rc = janitor.main(["--jobs-root", str(jobs_root), "--max-age-hours", "24"])
    assert rc == 1
