"""janitor.py（ジョブディレクトリ掃除スクリプト）のテスト。"""

from __future__ import annotations

import json
import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import janitor

ROOT = Path(__file__).resolve().parents[1]


def _make_job(jobs_root: Path, job_id: str, created_at: datetime | None) -> Path:
    job_dir = jobs_root / job_id
    job_dir.mkdir(parents=True)
    if created_at is not None:
        status = {"job_id": job_id, "status": "done", "created_at": created_at.isoformat()}
        (job_dir / "status.json").write_text(json.dumps(status), encoding="utf-8")
    return job_dir


def test_cleanup_removes_only_old_jobs(tmp_path):
    jobs_root = tmp_path / "jobs"
    now = datetime.now(timezone.utc)
    old_job = _make_job(jobs_root, "old-job", now - timedelta(hours=48))
    new_job = _make_job(jobs_root, "new-job", now - timedelta(minutes=5))

    removed = janitor.cleanup_old_jobs(jobs_root, max_age_hours=24)

    assert removed == ["old-job"]
    assert not old_job.exists()
    assert new_job.exists()


def test_cleanup_falls_back_to_mtime_without_status_json(tmp_path):
    jobs_root = tmp_path / "jobs"
    job_dir = _make_job(jobs_root, "no-status", created_at=None)
    old_time = (datetime.now(timezone.utc) - timedelta(hours=48)).timestamp()
    import os
    os.utime(job_dir, (old_time, old_time))

    removed = janitor.cleanup_old_jobs(jobs_root, max_age_hours=24)

    assert removed == ["no-status"]
    assert not job_dir.exists()


def test_cleanup_missing_jobs_root_returns_empty(tmp_path):
    assert janitor.cleanup_old_jobs(tmp_path / "nope", max_age_hours=24) == []


def test_cleanup_ignores_non_directory_entries(tmp_path):
    jobs_root = tmp_path / "jobs"
    jobs_root.mkdir()
    (jobs_root / "stray_file.txt").write_text("not a job dir")
    removed = janitor.cleanup_old_jobs(jobs_root, max_age_hours=0)
    assert removed == []
    assert (jobs_root / "stray_file.txt").exists()


def test_janitor_cli_smoke(tmp_path):
    jobs_root = tmp_path / "jobs"
    now = datetime.now(timezone.utc)
    _make_job(jobs_root, "old-job", now - timedelta(hours=48))
    _make_job(jobs_root, "new-job", now - timedelta(minutes=5))

    res = subprocess.run(
        [sys.executable, str(ROOT / "janitor.py"), "--jobs-root", str(jobs_root),
         "--max-age-hours", "24"],
        capture_output=True, text=True,
    )
    assert res.returncode == 0, res.stderr
    assert "old-job" in res.stdout
    assert not (jobs_root / "old-job").exists()
    assert (jobs_root / "new-job").exists()
