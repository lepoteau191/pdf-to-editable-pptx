#!/usr/bin/env python3
"""pdf2pptx の Web公開等を見据えたハードタイムアウト付きワーカー。

convert.py の `--timeout` はプロセス内のソフトタイムアウトで、ページ処理の
合間にしかチェックできない（1ページの描画がネイティブコード内でハングした
場合には効かない）。本ワーカーは convert.py を別の OS プロセスとして起動し、
親プロセスから hard_timeout 秒でSIGKILLする。SIGKILLはOSレベルの強制終了の
ため、ネイティブコード内でハングしていても確実に停止できる。

出力は一時ディレクトリ内に作らせ、子プロセスが正常終了した場合のみ本来の
出力先へ移動する（タイムアウトやエラーで不完全なファイルが出力先に残らない
ようにするため）。一時ディレクトリは成功・失敗にかかわらず必ず削除する。

使い方:
  python worker.py input.pdf output.pptx --hard-timeout 120 [convert.pyと同じオプション]
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import convert

DEFAULT_HARD_TIMEOUT = 120.0
CONVERT_PY = Path(__file__).resolve().parent / "convert.py"


def _build_child_argv(input_pdf: Path, output_pptx: Path, args) -> list[str]:
    argv = [sys.executable, str(CONVERT_PY), str(input_pdf), str(output_pptx)]
    argv += ["--dpi", str(args.dpi)]
    if args.debug_dir is not None:
        argv += ["--debug-dir", str(args.debug_dir)]
    argv += ["--max-pages", str(args.max_pages)]
    argv += ["--max-dpi", str(args.max_dpi)]
    argv += ["--max-page-pixels", str(args.max_page_pixels)]
    argv += ["--max-total-pixels", str(args.max_total_pixels)]
    argv += ["--max-file-size-mb", str(args.max_file_size_mb)]
    argv += ["--max-output-size-mb", str(args.max_output_size_mb)]
    if args.timeout is not None:
        argv += ["--timeout", str(args.timeout)]
    return argv


def run_subprocess_with_hard_timeout(
    argv: list[str], hard_timeout: float
) -> tuple[int, str, str, bool]:
    """任意のコマンドをhard_timeout秒でSIGKILLしながら実行する。

    convert.py固有の引数組み立てに依存しない、再利用可能なコア機構。
    戻り値は (returncode, stdout, stderr, timed_out)。
    """
    proc = subprocess.Popen(
        argv, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
    )
    timed_out = False
    try:
        stdout, stderr = proc.communicate(timeout=hard_timeout)
    except subprocess.TimeoutExpired:
        timed_out = True
        proc.kill()  # SIGKILL: ネイティブコード内でハングしていても確実に止める
        stdout, stderr = proc.communicate()

    returncode = proc.returncode
    if timed_out and returncode == 0:
        returncode = 1
    return returncode, stdout or "", stderr or "", timed_out


def run_with_hard_timeout(
    input_pdf: Path,
    output_pptx: Path,
    hard_timeout: float = DEFAULT_HARD_TIMEOUT,
    **convert_kwargs,
) -> tuple[int, str, str]:
    """convert.py を別プロセスで実行し、hard_timeout秒でSIGKILLする。

    戻り値は (returncode, stdout, stderr)。returncode==0 のときのみ
    output_pptx が書き込まれている。一時ディレクトリは常に削除する。
    """
    if not input_pdf.exists():
        raise FileNotFoundError(f"入力ファイルがありません: {input_pdf}")

    class _Args:
        dpi = convert_kwargs.get("dpi", 150)
        debug_dir = convert_kwargs.get("debug_dir")
        max_pages = convert_kwargs.get("max_pages", convert.DEFAULT_MAX_PAGES)
        max_dpi = convert_kwargs.get("max_dpi", convert.DEFAULT_MAX_DPI)
        max_page_pixels = convert_kwargs.get("max_page_pixels", convert.DEFAULT_MAX_PAGE_PIXELS)
        max_total_pixels = convert_kwargs.get("max_total_pixels", convert.DEFAULT_MAX_TOTAL_PIXELS)
        max_file_size_mb = convert_kwargs.get("max_file_size_mb", convert.DEFAULT_MAX_FILE_SIZE_MB)
        max_output_size_mb = convert_kwargs.get("max_output_size_mb", convert.DEFAULT_MAX_OUTPUT_SIZE_MB)
        timeout = convert_kwargs.get("timeout_seconds")

    tmp_dir = Path(tempfile.mkdtemp(prefix="pdf2pptx_worker_"))
    try:
        tmp_output = tmp_dir / "output.pptx"
        argv = _build_child_argv(input_pdf, tmp_output, _Args())

        returncode, stdout, stderr, timed_out = run_subprocess_with_hard_timeout(
            argv, hard_timeout
        )
        if timed_out:
            stderr += (
                f"\nエラー: ハードタイムアウト({hard_timeout}秒)のため処理を強制終了しました"
            )

        if returncode == 0 and tmp_output.exists():
            output_pptx.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(tmp_output), str(output_pptx))

        return returncode, stdout, stderr
    finally:
        # 一時ファイル（子プロセスが未完成のまま残した出力・debug成果物等）は
        # 常に削除する。debug_dir はユーザーが明示指定した外部パスのため対象外。
        shutil.rmtree(tmp_dir, ignore_errors=True)


def main(argv: list[str] | None = None) -> int:
    parser = convert.build_arg_parser()
    parser.description = "pdf2pptx をハードタイムアウト付きの別プロセスで実行する"
    parser.add_argument(
        "--hard-timeout", type=float, default=DEFAULT_HARD_TIMEOUT,
        help=f"子プロセスを強制終了するまでの秒数 (既定: {DEFAULT_HARD_TIMEOUT})",
    )
    args = parser.parse_args(argv)

    if not args.input.exists():
        print(f"エラー: 入力ファイルがありません: {args.input}", file=sys.stderr)
        return 1

    returncode, stdout, stderr = run_with_hard_timeout(
        args.input, args.output, hard_timeout=args.hard_timeout,
        dpi=args.dpi, debug_dir=args.debug_dir, max_pages=args.max_pages,
        max_dpi=args.max_dpi, max_page_pixels=args.max_page_pixels,
        max_total_pixels=args.max_total_pixels,
        max_file_size_mb=args.max_file_size_mb,
        max_output_size_mb=args.max_output_size_mb,
        timeout_seconds=args.timeout,
    )
    # 子プロセスの「完了: <一時パス>」行は内部の一時出力先を指すため隠し、
    # 呼び出し元から見える最終パスで報告し直す。
    filtered_stdout = "\n".join(
        line for line in stdout.splitlines() if not line.startswith("完了:")
    )
    if filtered_stdout:
        print(filtered_stdout)
    if stderr:
        print(stderr, end="", file=sys.stderr)
    if returncode == 0:
        print(f"完了: {args.output}")
    return 0 if returncode == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
