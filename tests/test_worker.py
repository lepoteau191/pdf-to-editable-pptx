"""worker.py（ハードタイムアウト付き別プロセス実行）のテスト。"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import pymupdf
import pytest
from pptx import Presentation

import fixtures_gen as fx
import worker

ROOT = Path(__file__).resolve().parents[1]


def _part_path(output_pptx: Path) -> Path:
    return output_pptx.with_name(output_pptx.name + ".part")


def test_build_child_argv_includes_ocr_options(tmp_path):
    class Args:
        dpi = 300
        debug_dir = None
        max_pages = 10
        max_dpi = 300
        max_page_pixels = 50_000_000
        max_total_pixels = 100_000_000
        max_file_size_mb = 100
        max_output_size_mb = 300
        timeout = None
        ocr = "tesseract"
        ocr_lang = "jpn+eng"
        ocr_min_conf = 35
        ocr_timeout = 30

    argv = worker._build_child_argv(tmp_path / "in.pdf", tmp_path / "out.pptx", Args())
    assert "--ocr" in argv
    assert argv[argv.index("--ocr") + 1] == "tesseract"
    assert argv[argv.index("--ocr-lang") + 1] == "jpn+eng"


# ---------------------------------------------------------------------------
# ハードタイムアウトのコア機構
# ---------------------------------------------------------------------------

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


def test_hard_timeout_kills_whole_process_group(tmp_path):
    """タイムアウト時、子だけでなく孫プロセスも道連れに終了させる。"""
    pid_file = tmp_path / "grandchild.pid"
    script = tmp_path / "spawn_grandchild.py"
    script.write_text(
        "import subprocess, sys, time\n"
        "gc = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(30)'])\n"
        f"open({str(pid_file)!r}, 'w').write(str(gc.pid))\n"
        "time.sleep(30)\n"
    )
    argv = [sys.executable, str(script)]

    returncode, stdout, stderr, timed_out = worker.run_subprocess_with_hard_timeout(
        argv, hard_timeout=1.0
    )
    assert timed_out is True
    assert pid_file.exists()
    grandchild_pid = int(pid_file.read_text())

    deadline = time.monotonic() + 2.0
    alive = True
    while time.monotonic() < deadline:
        try:
            os.kill(grandchild_pid, 0)
        except ProcessLookupError:
            alive = False
            break
        time.sleep(0.05)
    assert not alive, "孫プロセスが終了せず残っている"


# ---------------------------------------------------------------------------
# 入出力パス防御（worker.py自身での検査）
# ---------------------------------------------------------------------------

def test_worker_rejects_same_input_output_path(tmp_path):
    pdf = tmp_path / "same.pdf"
    fx.make_business_pdf(pdf)

    before = set(Path(tempfile.gettempdir()).glob("pdf2pptx_worker_*"))
    with pytest.raises(ValueError, match="同じパス"):
        worker.run_with_hard_timeout(pdf, pdf, hard_timeout=30.0)
    after = set(Path(tempfile.gettempdir()).glob("pdf2pptx_worker_*"))

    assert after == before  # 子プロセスすら起動していない（一時ディレクトリ無し）
    assert pymupdf.open(pdf).page_count == 3  # 入力が破壊されていない


def test_worker_rejects_hardlink_input_output(tmp_path):
    """resolve()では区別できないハードリンクも samefile で拒否する。"""
    pdf = tmp_path / "in.pdf"
    linked = tmp_path / "out.pptx"
    fx.make_business_pdf(pdf)
    os.link(pdf, linked)
    assert pdf.resolve() != linked.resolve()

    before = set(Path(tempfile.gettempdir()).glob("pdf2pptx_worker_*"))
    with pytest.raises(ValueError, match="同一ファイル"):
        worker.run_with_hard_timeout(pdf, linked, hard_timeout=30.0)
    after = set(Path(tempfile.gettempdir()).glob("pdf2pptx_worker_*"))

    assert after == before
    # ハードリンクなので pdf 経由で読んでも入力の内容が壊れていないこと
    assert pymupdf.open(pdf).page_count == 3
    assert pymupdf.open(linked).page_count == 3


# ---------------------------------------------------------------------------
# アトミックな出力反映（.part + os.replace）
# ---------------------------------------------------------------------------

def test_publish_atomically_replaces_final_output(tmp_path):
    src = tmp_path / "src.pptx"
    with zipfile.ZipFile(src, "w") as zf:
        zf.writestr("dummy.txt", "hello")
    dest = tmp_path / "out" / "final.pptx"

    worker._publish_atomically(src, dest)

    assert dest.exists()
    assert not _part_path(dest).exists()


def test_publish_atomically_invalid_zip_leaves_no_part(tmp_path):
    src = tmp_path / "bad.pptx"
    src.write_bytes(b"this is not a zip file at all")
    dest = tmp_path / "final.pptx"

    with pytest.raises(ValueError, match="zip"):
        worker._publish_atomically(src, dest)

    assert not dest.exists()
    assert not _part_path(dest).exists()


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
    assert not _part_path(out).exists()  # .part が残っていないこと


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
    assert not _part_path(out).exists()
    assert after == before


def test_worker_cli_smoke(tmp_path):
    pdf = tmp_path / "in.pdf"
    out = tmp_path / "out.pptx"
    fx.make_business_pdf(pdf)

    res = subprocess.run(
        [sys.executable, str(ROOT / "worker.py"), str(pdf), str(out),
         "--hard-timeout", "30"],
        capture_output=True, text=True,
    )
    assert res.returncode == 0, res.stderr
    assert out.exists()


def test_worker_cli_rejects_same_input_output_path(tmp_path):
    pdf = tmp_path / "same.pdf"
    fx.make_business_pdf(pdf)

    res = subprocess.run(
        [sys.executable, str(ROOT / "worker.py"), str(pdf), str(pdf),
         "--hard-timeout", "30"],
        capture_output=True, text=True,
    )
    assert res.returncode == 1
    assert pymupdf.open(pdf).page_count == 3
