#!/usr/bin/env python3
"""pdf2pptx Webアプリ（app.py）が作るジョブディレクトリの掃除スクリプト。

app.py はアップロードされたPDFと生成したPPTXを `jobs/<job_id>/` に置いたままにする
（ダウンロードのため）。放置すると増え続けるため、cronや手動実行で定期的に
本スクリプトを呼び、一定時間より古いジョブディレクトリを削除する。

安全のため次のジョブは削除対象から外す/扱いを変える:
  - ディレクトリ名がUUID形式でないもの（不正なエントリを誤って触らない）
  - status.json内のjob_idがディレクトリ名と一致しないもの（破損/改変の疑い）
  - status が queued/processing のもの: 変換がまだ実行中の可能性があるため、
    通常の max_age_hours ではなく、より長い stale_timeout_hours を超えた
    場合のみ「放置されたジョブ」とみなして削除する。

使い方:
  python janitor.py --max-age-hours 24
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path

JOBS_ROOT = Path(__file__).resolve().parent / "jobs"
DEFAULT_MAX_AGE_HOURS = 24.0
TERMINAL_STATUSES = ("done", "error")
IN_PROGRESS_STATUSES = ("queued", "processing")


def _read_job_info(job_dir: Path) -> dict | None:
    """status.jsonを読む。存在しない/壊れている場合はNoneを返す。"""
    status_path = job_dir / "status.json"
    if not status_path.exists():
        return None
    try:
        return json.loads(status_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _job_created_at(job_dir: Path, info: dict | None) -> datetime:
    """ジョブの作成時刻を返す。status.jsonが読めない場合はディレクトリのmtimeで代用する。"""
    if info is not None:
        try:
            return datetime.fromisoformat(info["created_at"])
        except (KeyError, ValueError):
            pass
    return datetime.fromtimestamp(job_dir.stat().st_mtime, tz=timezone.utc)


def cleanup_old_jobs(
    jobs_root: Path,
    max_age_hours: float,
    stale_timeout_hours: float | None = None,
) -> tuple[list[str], list[str]]:
    """ジョブディレクトリを掃除する。

    戻り値は (削除したjob_idのリスト, 削除に失敗した際の警告メッセージのリスト)。
    削除失敗は黙殺せず警告として返す（呼び出し元が表示・ログできるようにする）。
    """
    if stale_timeout_hours is None:
        stale_timeout_hours = max_age_hours
    if not jobs_root.exists():
        return [], []

    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(hours=max_age_hours)
    stale_cutoff = now - timedelta(hours=stale_timeout_hours)

    removed: list[str] = []
    warnings: list[str] = []

    for job_dir in sorted(jobs_root.iterdir()):
        if not job_dir.is_dir():
            continue

        # UUID形式でないディレクトリ名は対象外（不正なエントリを誤って触らない）
        try:
            uuid.UUID(job_dir.name)
        except ValueError:
            continue

        info = _read_job_info(job_dir)

        # status.json記載のjob_idとディレクトリ名が食い違う場合は触らない
        # （破損・改変・別プロセスによる書き込み競合の疑いがあるため安全側に倒す）
        if info is not None and info.get("job_id") and info["job_id"] != job_dir.name:
            continue

        status_value = info.get("status") if info else None
        created_at = _job_created_at(job_dir, info)
        threshold = stale_cutoff if status_value in IN_PROGRESS_STATUSES else cutoff

        if created_at >= threshold:
            continue

        try:
            shutil.rmtree(job_dir)
            removed.append(job_dir.name)
        except OSError as e:
            warnings.append(f"削除に失敗しました: {job_dir.name} ({e})")

    return removed, warnings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="pdf2pptx Webアプリの古いジョブディレクトリを削除する"
    )
    parser.add_argument(
        "--max-age-hours", type=float, default=DEFAULT_MAX_AGE_HOURS,
        help=f"完了(done)・失敗(error)のジョブをこの時間(時)より古ければ削除する "
             f"(既定: {DEFAULT_MAX_AGE_HOURS})",
    )
    parser.add_argument(
        "--stale-timeout-hours", type=float, default=None,
        help="queued/processingのまま放置されたジョブをこの時間(時)より古ければ削除する "
             "(既定: --max-age-hoursと同じ)",
    )
    parser.add_argument(
        "--jobs-root", type=Path, default=JOBS_ROOT,
        help="ジョブディレクトリの親パス (既定: jobs/)",
    )
    args = parser.parse_args(argv)

    removed, warnings = cleanup_old_jobs(
        args.jobs_root, args.max_age_hours, args.stale_timeout_hours
    )
    for job_id in removed:
        print(f"削除: {job_id}")
    for w in warnings:
        print(f"警告: {w}", file=sys.stderr)
    print(f"{len(removed)}件のジョブディレクトリを削除しました")
    return 1 if warnings else 0


if __name__ == "__main__":
    sys.exit(main())
