"""テスト用の合成PDFを生成する（ネットワーク不使用・完全ローカル）。"""

from __future__ import annotations

import io
import os
import shutil
import subprocess
from pathlib import Path

import pymupdf
from PIL import Image

A4_W, A4_H = 595.0, 842.0

TITLE = "2026年度 事業計画書"
CONFIDENTIAL = "Confidential - Internal Use Only"
BODY_LINES = [
    "本資料は事業計画の概要をまとめたものです。",
    "対象期間は2026年4月から2027年3月までとします。",
    "数値はすべて概算であり、確定値ではありません。",
]
NOTE = "※ 社外秘。取り扱いに注意してください。"
FOOTER = "pdf2pptx test fixture - page 1"
TABLE_HEADER = ["項目", "上期", "下期", "合計"]
TABLE_ROWS = [
    ["売上高", "1,234", "1,567", "2,801"],
    ["営業利益", "234", "312", "546"],
]
ROTATED_TEXT = "回転テキストは編集対象外"
PAGE2_BODY = "2ページ目の通常テキストです。"
ROTATED_PAGE_TEXT = "回転ページのテキスト"
MIXED_P1_TEXT = "縦ページのテキスト"
MIXED_P2_TEXT = "横ページのテキスト"
SMALL_PAGE_TEXT = "小さいページのテキスト"
OCR_INVISIBLE_TEXT = "これはOCRで生成された不可視テキストです"
OVERLAP_ROTATED_TEXT = "縦向き文字列"
OVERLAP_HORIZONTAL_TEXT = "重複候補テキスト"
OVERLAP_NORMAL_TEXT = "通常の横書き行"
FULL_IMAGE_VISIBLE_TEXT = "画像の上に載った可視テキスト"

TITLE_COLOR = (0x1F / 255, 0x38 / 255, 0x64 / 255)  # 0x1F3864
RED = (0.8, 0.1, 0.1)
GRAY = (0.45, 0.45, 0.45)

# PyMuPDFのpip wheelにはCJK内蔵フォントがないため、OS側のフォントを使う。
# .ttc（TrueTypeコレクション）でもPyMuPDFはface 0を正しく読み込める（動作確認済み）。
_JP_FONT_CANDIDATES = [
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",         # macOS
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",        # Ubuntu: fonts-noto-cjk
    "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",            # Ubuntu: fonts-ipafont-gothic
    "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",           # Debian/Ubuntu (旧パッケージ名)
]


def _resolve_jp_font() -> str:
    # 環境変数での明示的な上書き（CI差異の再現・診断や、候補にないフォント
    # 環境での手動指定に使う）。存在しないパスが指定された場合は無視して
    # 通常の解決フローに進む。
    override = os.environ.get("PDF2PPTX_TEST_JP_FONT")
    if override and os.path.exists(override):
        return override
    for path in _JP_FONT_CANDIDATES:
        if os.path.exists(path):
            return path
    if shutil.which("fc-match"):
        try:
            result = subprocess.run(
                ["fc-match", "-f", "%{file}", ":lang=ja"],
                capture_output=True, text=True, timeout=5, check=True,
            )
            if result.stdout.strip():
                return result.stdout.strip()
        except (subprocess.SubprocessError, OSError):
            pass
    raise RuntimeError(
        "日本語(CJK)フォントが見つかりません。Linux/CIでは "
        "`sudo apt-get install -y fonts-noto-cjk fontconfig` 等でインストールしてください。"
    )


JP_FONT_FILE = _resolve_jp_font()


def _jp_text(page, point, text, size, color=(0, 0, 0), rotate=0):
    page.insert_text(point, text, fontname="jpfx", fontfile=JP_FONT_FILE,
                     fontsize=size, color=color, rotate=rotate)


def _invisible_jp_text(page, point, text, size):
    """render_mode=3 (invisible) でテキストを挿入する。OCRの不可視テキスト層の模擬。"""
    page.insert_text(point, text, fontname="jpfx", fontfile=JP_FONT_FILE,
                     fontsize=size, render_mode=3)


def _logo_png() -> bytes:
    """Pillowで簡単なロゴ風画像を作る。"""
    img = Image.new("RGB", (240, 120))
    px = img.load()
    for y in range(120):
        for x in range(240):
            px[x, y] = (30 + x // 2, 80 + y, 160)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _draw_table(page: pymupdf.Page, x: float, y: float) -> None:
    col_w, row_h = 100.0, 24.0
    n_cols = len(TABLE_HEADER)
    n_rows = 1 + len(TABLE_ROWS)
    # ヘッダ行の塗り
    page.draw_rect(
        pymupdf.Rect(x, y, x + col_w * n_cols, y + row_h),
        fill=(0.90, 0.92, 0.96), color=None,
    )
    # 罫線
    for i in range(n_rows + 1):
        page.draw_line((x, y + i * row_h), (x + col_w * n_cols, y + i * row_h),
                       color=GRAY, width=0.7)
    for j in range(n_cols + 1):
        page.draw_line((x + j * col_w, y), (x + j * col_w, y + n_rows * row_h),
                       color=GRAY, width=0.7)
    # セルの文字
    rows = [TABLE_HEADER] + TABLE_ROWS
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            _jp_text(page, (x + c * col_w + 6, y + r * row_h + row_h - 8),
                     text, 10)


def make_business_pdf(path: Path) -> None:
    """3ページ構成: ビジネス文書 / 回転テキスト入り / スキャン風（画像のみ）。"""
    doc = pymupdf.open()

    # --- page 1: 横書きビジネス文書（表・画像・色・太字入り） ---
    page = doc.new_page(width=A4_W, height=A4_H)
    _jp_text(page, (60, 80), TITLE, 20, color=TITLE_COLOR)
    page.insert_text((60, 110), CONFIDENTIAL, fontname="hebo",  # Helvetica-Bold
                     fontsize=11, color=RED)
    for i, line in enumerate(BODY_LINES):
        _jp_text(page, (60, 160 + i * 18), line, 10.5)
    _jp_text(page, (60, 230), NOTE, 9, color=RED)
    page.insert_image(pymupdf.Rect(420, 50, 540, 110), stream=_logo_png())
    _draw_table(page, 60, 280)
    page.insert_text((60, 810), FOOTER, fontname="helv", fontsize=8, color=GRAY)

    # --- page 2: 通常テキスト + 回転テキスト ---
    page2 = doc.new_page(width=A4_W, height=A4_H)
    _jp_text(page2, (60, 80), PAGE2_BODY, 12)
    _jp_text(page2, (100, 500), ROTATED_TEXT, 12, rotate=90)

    # --- page 3: スキャン風（page1のレンダリング画像のみ・テキスト層なし） ---
    pix = doc[0].get_pixmap(dpi=100)
    page3 = doc.new_page(width=A4_W, height=A4_H)
    page3.insert_image(page3.rect, pixmap=pix)

    doc.save(path)
    doc.close()


def make_mixed_pdf(path: Path) -> None:
    """ページサイズ混在: A4縦 + A4横。"""
    doc = pymupdf.open()
    p1 = doc.new_page(width=A4_W, height=A4_H)
    _jp_text(p1, (60, 80), MIXED_P1_TEXT, 12)
    p2 = doc.new_page(width=A4_H, height=A4_W)
    _jp_text(p2, (60, 80), MIXED_P2_TEXT, 12)
    doc.save(path)
    doc.close()


def make_rotated_page_pdf(path: Path) -> None:
    """ページ自体に /Rotate 90 が付いたPDF（Phase 1の背景フォールバック検査用）。"""
    doc = pymupdf.open()
    page = doc.new_page(width=A4_W, height=A4_H)
    _jp_text(page, (60, 80), ROTATED_PAGE_TEXT, 12)
    page.set_rotation(90)
    doc.save(path)
    doc.close()


def make_rotated_page_scan_pdf(path: Path) -> None:
    """回転ページ(/Rotate 90) + 可視テキストなし（画像のみ）。

    OCR対象（any_visible=False）になる回転ページを作るための、
    make_rotated_page_pdf とは別のフィクスチャ。OCR結果の座標が
    回転後のページ座標系（page.rectは回転を反映した幅・高さを返す）で
    正しく扱われることを検証するのに使う。
    """
    doc = pymupdf.open()
    page = doc.new_page(width=A4_W, height=A4_H)
    img = Image.new("RGB", (400, 566), (235, 235, 230))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    page.insert_image(page.rect, stream=buf.getvalue())
    page.set_rotation(90)
    doc.save(path)
    doc.close()


def make_small_mixed_pdf(path: Path) -> None:
    """ページサイズ混在: 通常ページ + より小さいページ（拡大されないことの検証用）。"""
    doc = pymupdf.open()
    p1 = doc.new_page(width=A4_W, height=A4_H)
    _jp_text(p1, (60, 80), MIXED_P1_TEXT, 12)
    p2 = doc.new_page(width=200, height=100)
    _jp_text(p2, (10, 50), SMALL_PAGE_TEXT, 10)
    doc.save(path)
    doc.close()


def make_invisible_ocr_pdf(path: Path) -> None:
    """全面画像 + 不可視テキスト層のみのページ（OCR済みスキャンPDFの模擬）。"""
    doc = pymupdf.open()
    page = doc.new_page(width=A4_W, height=A4_H)
    img = Image.new("RGB", (400, 566), (245, 245, 240))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    page.insert_image(page.rect, stream=buf.getvalue())
    _invisible_jp_text(page, (60, 80), OCR_INVISIBLE_TEXT, 12)
    doc.save(path)
    doc.close()


def make_overlap_pdf(path: Path) -> None:
    """回転テキストと横書きテキストのbboxが重なるページ（欠損防止検査用）。

    重なっていない通常の横書き行も同一ページに置き、重なり判定が
    行単位で正しく限定されることを確認できるようにする。
    """
    doc = pymupdf.open()
    page = doc.new_page(width=A4_W, height=A4_H)
    _jp_text(page, (100, 150), OVERLAP_ROTATED_TEXT, 16, rotate=90)
    _jp_text(page, (95, 120), OVERLAP_HORIZONTAL_TEXT, 16)
    _jp_text(page, (60, 300), OVERLAP_NORMAL_TEXT, 12)
    doc.save(path)
    doc.close()


def make_full_image_visible_text_pdf(path: Path) -> None:
    """全面画像 + 可視テキストのページ（画像被覆率フォールバック検査用）。

    不可視OCR層ではなく、通常の可視テキストが全面画像の上に乗っているケース。
    """
    doc = pymupdf.open()
    page = doc.new_page(width=A4_W, height=A4_H)
    img = Image.new("RGB", (400, 566), (230, 230, 225))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    page.insert_image(page.rect, stream=buf.getvalue())
    _jp_text(page, (60, 80), FULL_IMAGE_VISIBLE_TEXT, 12)
    doc.save(path)
    doc.close()


def make_extreme_page_size_pdf(path: Path) -> None:
    """PowerPointのスライドサイズ制限(1〜56インチ)を超える極端なページサイズ。"""
    doc = pymupdf.open()
    # 高さ 5000pt ≈ 69.4インチ (> 56インチ上限)
    doc.new_page(width=A4_W, height=5000.0)
    doc.save(path)
    doc.close()


def make_tiny_page_size_pdf(path: Path) -> None:
    """PowerPointのスライドサイズ制限(1〜56インチ)の下限を下回る極小ページ。"""
    doc = pymupdf.open()
    # 幅 36pt = 0.5インチ (< 1インチ下限)
    doc.new_page(width=36.0, height=100.0)
    doc.save(path)
    doc.close()


def make_image_heavy_pdf(path: Path, n_images: int = 15) -> None:
    """画像を多数埋め込んだPDF（get_textの画像抑制の検証用）。"""
    doc = pymupdf.open()
    page = doc.new_page(width=A4_W, height=A4_H)
    img = Image.new("RGB", (600, 600), (80, 120, 200))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    stream = buf.getvalue()
    cols = 5
    cell = 100.0
    for i in range(n_images):
        x = 20 + (i % cols) * cell
        y = 20 + (i // cols) * cell
        page.insert_image(pymupdf.Rect(x, y, x + cell - 10, y + cell - 10), stream=stream)
    _jp_text(page, (60, 780), PAGE2_BODY, 12)
    doc.save(path)
    doc.close()
