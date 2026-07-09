"""worker.py（ハードタイムアウト付き別プロセス実行）のテスト。"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from pptx import Presentation

import fixtures_gen as fx
import worker

ROOT = Path(__file__).resolve().parents[1]


def test_run_subprocess_with_hard_timeout_kills_hanging_process():
    """ハングするプロセスがhard_timeout秒で確実にSIGKILLされる。"""
    argv = [sys.executable, "-c", "import time; time.sleep(30)"]
    returncode, stdout, stderr, timed_out = worker.run_subprocess_with_hard_timeout(
        argv, hard_timeout=0.5
    )
    assert timed_out is True
    assert returncode != 0


def test_run_subprocess_with_hard_timeout_normal_exit():
    argv = [sys.executable, "-c", "print('hello')"]
    returncode, stdout, stderr, timed_out = worker.run_subprocess_with_hard_timeout(
        argv, hard_timeout=10.0
    )
    assert timed_out is False
    assert returncode == 0
    assert "hello" in stdout


def test_worker_success_moves_output_and_cleans_tmp(tmp_path):
    pdf = tmp_path / "in.pdf"
    out = tmp_path / "nested" / "out.pptx"  # 出力先ディレクトリも未作成の状態
    fx.make_business_pdf(pdf)

    before = set(Path(tempfile.gettempdir()).glob("pdf2pptx_worker_*"))
    returncode, stdout, stderr = worker.run_with_hard_timeout(
        pdf, out, hard_timeout=30.0
    )
    after = set(Path(tempfile.gettempdir()).glob("pdf2pptx_worker_*"))

    assert returncode == 0, stderr
    assert out.exists()
    assert Presentation(out).slides
    assert after == before  # 一時ディレクトリが残っていないこと


def test_worker_hard_timeout_leaves_no_output_and_cleans_tmp(tmp_path, monkeypatch):
    """ハードタイムアウト時、出力先にファイルが作られず一時ディレクトリも残らない。"""
    pdf = tmp_path / "in.pdf"
    out = tmp_path / "out.pptx"
    fx.make_business_pdf(pdf)

    # convert.py の代わりに「ハングするダミースクリプト」を子プロセスとして
    # 起動させ、実際のPDF処理時間に依存せず高速かつ確定的にテストする。
    hang_script = tmp_path / "hang.py"
    hang_script.write_text("import time\ntime.sleep(30)\n")
    monkeypatch.setattr(worker, "CONVERT_PY", hang_script)

    before = set(Path(tempfile.gettempdir()).glob("pdf2pptx_worker_*"))
    returncode, stdout, stderr = worker.run_with_hard_timeout(
        pdf, out, hard_timeout=0.5
    )
    after = set(Path(tempfile.gettempdir()).glob("pdf2pptx_worker_*"))

    assert returncode != 0
    assert "ハードタイムアウト" in stderr
    assert not out.exists()
    assert after == before


def test_worker_cli_smoke(tmp_path):
    pdf = tmp_path / "in.pdf"
    out = tmp_path / "out.pptx"
    fx.make_business_pdf(pdf)
    import subprocess

    res = subprocess.run(
        [sys.executable, str(ROOT / "worker.py"), str(pdf), str(out),
         "--hard-timeout", "30"],
        capture_output=True, text=True,
    )
    assert res.returncode == 0, res.stderr
    assert out.exists()
