#!/usr/bin/env python3
"""pdf2pptx の Web公開等を見据えたハードタイムアウト付きワーカー。

convert.py の `--timeout` はプロセス内のソフトタイムアウトで、ページ処理の
合間にしかチェックできない（1ページの描画がネイティブコード内でハングした
場合には効かない）。本ワーカーは convert.py を別の OS プロセスとして起動し、
親プロセスから hard_timeout 秒でプロセスグループごとSIGKILLする。SIGKILLは
OSレベルの強制終了のため、ネイティブコード内でハングしていても確実に停止
できる。プロセスグループごと止めるため、子プロセスがさらに孫プロセスを
起動していても道連れにできる。

出力は一時ディレクトリ内に作らせ、子プロセスが正常終了した場合のみ
「出力先と同じディレクトリの .part ファイル」へコピーし、検証したうえで
os.replace() により最終出力へアトミックに反映する（コピー先ファイル
システムをまたぐ可能性のある shutil.move() の非アトミックなフォールバック
を避けるため）。一時ディレクトリは成功・失敗にかかわらず削除を試み、
削除に失敗した場合は（黙殺せず）警告を返す。

このワーカー自身も、実行前に入力と出力が同一ファイルでないことを検査する
（convert.py にも同じ検査があるが、子プロセスには一時ディレクトリ内の
パスを渡すため、convert.py側の検査だけではworker.py経由の実行を保護
できない。詳細は convert.check_distinct_input_output のdocstring参照）。

**Web公開時の注意**: ハードタイムアウトだけでは不十分。本番でこのワーカーを
Webから使う場合のリソース隔離方針は README の
「Web公開時のリソース隔離方針」を必ず参照すること。

使い方:
  python worker.py input.pdf output.pptx --hard-timeout 120 [convert.pyと同じオプション]
"""

from __future__ import annotations

import os
import shutil
import signal
import subprocess
import sys
import tempfile
import zipfile
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
    argv += ["--ocr", str(args.ocr)]
    argv += ["--ocr-lang", str(args.ocr_lang)]
    argv += ["--ocr-min-conf", str(args.ocr_min_conf)]
    argv += ["--ocr-timeout", str(args.ocr_timeout)]
    return argv


def run_subprocess_with_hard_timeout(
    argv: list[str], hard_timeout: float
) -> tuple[int, str, str, bool]:
    """任意のコマンドをhard_timeout秒でプロセスグループごとSIGKILLしながら実行する。

    convert.py固有の引数組み立てに依存しない、再利用可能なコア機構。
    `start_new_session=True` で子を新しいプロセスグループのリーダーにし、
    タイムアウト時は `os.killpg` でグループ全体（孫プロセスを含む）を
    強制終了する。戻り値は (returncode, stdout, stderr, timed_out)。
    """
    proc = subprocess.Popen(
        argv, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
        start_new_session=True,
    )
    timed_out = False
    try:
        stdout, stderr = proc.communicate(timeout=hard_timeout)
    except subprocess.TimeoutExpired:
        timed_out = True
        try:
            os.killpg(proc.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass  # 既に終了している
        stdout, stderr = proc.communicate()

    returncode = proc.returncode
    if timed_out and returncode == 0:
        returncode = 1
    return returncode, stdout or "", stderr or "", timed_out


def _publish_atomically(tmp_output: Path, output_pptx: Path) -> None:
    """一時出力を検証したうえで、出力先へアトミックに反映する。

    出力先と同じディレクトリに .part ファイルを作ってから os.replace()
    することで、コピー元とコピー先が異なるファイルシステムにまたがる場合の
    shutil.move() の非アトミックなフォールバック（コピー+削除）を避ける。
    検証に失敗した場合、.part ファイルは残さず削除する。
    """
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    part_path = output_pptx.with_name(output_pptx.name + ".part")
    try:
        shutil.copyfile(str(tmp_output), str(part_path))
        if not zipfile.is_zipfile(part_path):
            raise ValueError("生成されたPPTXが不正です（zip形式として認識できません）")
        os.replace(str(part_path), str(output_pptx))
    except BaseException:
        part_path.unlink(missing_ok=True)
        raise


def run_with_hard_timeout(
    input_pdf: Path,
    output_pptx: Path,
    hard_timeout: float = DEFAULT_HARD_TIMEOUT,
    **convert_kwargs,
) -> tuple[int, str, str]:
    """convert.py を別プロセスで実行し、hard_timeout秒でSIGKILLする。

    戻り値は (returncode, stdout, stderr)。returncode==0 のときのみ
    output_pptx が書き込まれている。一時ディレクトリは常に削除を試み、
    失敗時はstderrに警告を追記する（黙殺しない）。
    """
    if not input_pdf.exists():
        raise FileNotFoundError(f"入力ファイルがありません: {input_pdf}")
    # convert.py にも同じ検査があるが、子プロセスには一時ディレクトリ内の
    # パスを渡すため、実際の input_pdf/output_pptx に対してはここで検査する
    # 必要がある（子プロセス側の検査だけでは保護できない）。
    convert.check_distinct_input_output(input_pdf, output_pptx)

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
        ocr = convert_kwargs.get("ocr", convert.DEFAULT_OCR_ENGINE)
        ocr_lang = convert_kwargs.get("ocr_lang", convert.DEFAULT_OCR_LANG)
        ocr_min_conf = convert_kwargs.get("ocr_min_conf", convert.DEFAULT_OCR_MIN_CONF)
        ocr_timeout = convert_kwargs.get("ocr_timeout", convert.DEFAULT_OCR_TIMEOUT)

    tmp_dir = Path(tempfile.mkdtemp(prefix="pdf2pptx_worker_"))
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
        try:
            _publish_atomically(tmp_output, output_pptx)
        except Exception as e:  # noqa: BLE001 - 呼び出し元へ (returncode, stderr) で伝える
            returncode = 1
            stderr += f"\nエラー: 出力の確定に失敗しました: {e}"

    # 一時ディレクトリの削除は試みるが、ignore_errors=True で完全に黙殺は
    # しない。削除に失敗した場合はstderrに警告として残す（janitorでの
    # 定期的な掃除が必要になる可能性があるため）。
    try:
        shutil.rmtree(tmp_dir)
    except OSError as e:
        stderr += (
            f"\n警告: 一時ディレクトリの削除に失敗しました: {tmp_dir} ({e})。"
            "定期的なjanitorでの掃除を検討してください"
        )

    return returncode, stdout, stderr


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

    try:
        returncode, stdout, stderr = run_with_hard_timeout(
            args.input, args.output, hard_timeout=args.hard_timeout,
            dpi=args.dpi, debug_dir=args.debug_dir, max_pages=args.max_pages,
            max_dpi=args.max_dpi, max_page_pixels=args.max_page_pixels,
            max_total_pixels=args.max_total_pixels,
            max_file_size_mb=args.max_file_size_mb,
            max_output_size_mb=args.max_output_size_mb,
            timeout_seconds=args.timeout,
            ocr=args.ocr,
            ocr_lang=args.ocr_lang,
            ocr_min_conf=args.ocr_min_conf,
            ocr_timeout=args.ocr_timeout,
        )
    except (ValueError, FileNotFoundError) as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1

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
