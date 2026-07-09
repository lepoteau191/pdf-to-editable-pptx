#!/usr/bin/env python3
"""pdf2pptx Webアプリ（app.py）が作るジョブディレクトリの掃除スクリプト。

app.py はアップロードされたPDFと生成したPPTXを `jobs/<job_id>/` に置いたままにする
（ダウンロードのため）。放置すると増え続けるため、cronや手動実行で定期的に
本スクリプトを呼び、一定時間より古いジョブディレクトリを削除する。

使い方:
  python janitor.py --max-age-hours 24
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

JOBS_ROOT = Path(__file__).resolve().parent / "jobs"
DEFAULT_MAX_AGE_HOURS = 24.0


def _job_created_at(job_dir: Path) -> datetime:
    """ジョブの作成時刻を返す。status.jsonが読めない場合はディレクトリのmtimeで代用する。"""
    status_path = job_dir / "status.json"
    if status_path.exists():
        try:
            data = json.loads(status_path.read_text(encoding="utf-8"))
            return datetime.fromisoformat(data["created_at"])
        except (OSError, ValueError, KeyError, json.JSONDecodeError):
            pass
    return datetime.fromtimestamp(job_dir.stat().st_mtime, tz=timezone.utc)


def cleanup_old_jobs(jobs_root: Path, max_age_hours: float) -> list[str]:
    """max_age_hours より古いジョブディレクトリを削除し、削除したjob_idの一覧を返す。"""
    if not jobs_root.exists():
        return []
    cutoff = datetime.now(timezone.utc) - timedelta(hours=max_age_hours)
    removed: list[str] = []
    for job_dir in sorted(jobs_root.iterdir()):
        if not job_dir.is_dir():
            continue
        if _job_created_at(job_dir) < cutoff:
            shutil.rmtree(job_dir, ignore_errors=True)
            removed.append(job_dir.name)
    return removed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="pdf2pptx Webアプリの古いジョブディレクトリを削除する"
    )
    parser.add_argument(
        "--max-age-hours", type=float, default=DEFAULT_MAX_AGE_HOURS,
        help=f"この時間(時)より古いジョブを削除する (既定: {DEFAULT_MAX_AGE_HOURS})",
    )
    parser.add_argument(
        "--jobs-root", type=Path, default=JOBS_ROOT,
        help="ジョブディレクトリの親パス (既定: jobs/)",
    )
    args = parser.parse_args(argv)

    removed = cleanup_old_jobs(args.jobs_root, args.max_age_hours)
    for job_id in removed:
        print(f"削除: {job_id}")
    print(f"{len(removed)}件のジョブディレクトリを削除しました")
    return 0


if __name__ == "__main__":
    sys.exit(main())
