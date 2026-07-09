#!/usr/bin/env python3
"""pdf2pptx — PDFを編集可能なPPTXに変換する (Phase 1 MVP).

方式:
  1. PyMuPDF でテキストを span/line/block 単位で抽出する（横書きのみ編集対象）
  2. 編集対象テキストの領域を redaction で背景から削除する（二重写り防止。
     画像・罫線・図形は残す）
  3. redaction 後のページを画像化してスライド背景に敷く
  4. 抽出テキストを編集可能なテキストボックスとして同じ座標に重ねる

使い方:
  python convert.py input.pdf output.pptx [--dpi 150] [--debug-dir DIR]
"""

from __future__ import annotations

import argparse
import io
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

EMU_PER_PT = 12700

# 横書き判定: line の書字方向ベクトルが (1, 0) にほぼ一致すること
HORIZONTAL_DIR_TOL = 0.02
# redaction 矩形を少し縮めて隣接コンテンツの巻き込みを防ぐ (pt)
REDACT_SHRINK = 0.2
# span 中の置換文字 (U+FFFD) がこの割合を超えたら文字化けとみなし背景に残す
GARBLED_RATIO = 0.3
# テキストボックスの幅の余裕。代替フォントの字幅差で右端が欠けるのを防ぐ (pt)
BOX_PAD = 1.0
# ページサイズ差をこの値 (pt) まで「同一サイズ」とみなす
PAGE_SIZE_TOL = 1.0


# ---------------------------------------------------------------------------
# フォントマッピング
# ---------------------------------------------------------------------------

# PDF内フォント名（サブセット接頭辞除去・小文字・空白/ハイフン除去後）の
# 完全一致マップ。ヒューリスティックより優先される。
FONT_MAP = {
    "msgothic": "ＭＳ ゴシック",
    "mspgothic": "ＭＳ Ｐゴシック",
    "msmincho": "ＭＳ 明朝",
    "mspmincho": "ＭＳ Ｐ明朝",
    "meiryo": "Meiryo",
    "meiryoui": "Meiryo UI",
    "yugothic": "Yu Gothic",
    "yumincho": "Yu Mincho",
    "arial": "Arial",
    "helvetica": "Arial",
    "timesnewroman": "Times New Roman",
    "times": "Times New Roman",
    "couriernew": "Courier New",
    "courier": "Courier New",
    "calibri": "Calibri",
}

JP_HINTS = (
    "gothic", "mincho", "meiryo", "hiragino", "hira", "yugo", "yumin",
    "kozuka", "kozgo", "kozmin", "sourcehan", "notosanscjk", "notosansjp",
    "notoserifcjk", "notoserifjp", "ipaex", "ipag", "ipam", "biz", "udshingo",
    "ryumin", "shingo", "midashi", "maru", "kaku", "japan",
)
SERIF_HINTS = (
    "mincho", "ryumin", "kozmin", "notoserif", "sourcehanserif", "ipam",
    "times", "georgia", "garamond", "serif", "roman", "century", "cambria",
    "palatino", "minion", "book",
)
MONO_HINTS = ("courier", "mono", "consolas", "menlo")

FALLBACK_JP_SANS = "Yu Gothic"
FALLBACK_JP_SERIF = "Yu Mincho"
FALLBACK_LATIN_SANS = "Arial"
FALLBACK_LATIN_SERIF = "Times New Roman"
FALLBACK_MONO = "Courier New"

# span flags (PyMuPDF)
FLAG_ITALIC = 1 << 1
FLAG_MONO = 1 << 3
FLAG_BOLD = 1 << 4

_SUBSET_PREFIX = re.compile(r"^[A-Z]{6}\+")


def _contains_cjk(text: str) -> bool:
    return any(
        "　" <= ch <= "ヿ"      # 記号・かな
        or "一" <= ch <= "鿿"   # 漢字
        or "＀" <= ch <= "￯"   # 全角英数・半角カナ
        for ch in text
    )


_LATIN_ONLY_FAMILIES = {"Arial", "Times New Roman", "Courier New", "Calibri"}


def map_font(pdf_font_name: str, text: str) -> str:
    """PDF内のフォント名を PowerPoint で使えるファミリー名に変換する。"""
    name = _SUBSET_PREFIX.sub("", pdf_font_name or "")
    key = name.lower().replace(" ", "").replace("-", "").replace(",", "")
    for exact, family in FONT_MAP.items():
        if key.startswith(exact):
            # 日本語テキストが欧文専用フォント名 (Arial Unicode MS 等) に
            # 載っている場合は日本語フォールバックを優先する
            if family in _LATIN_ONLY_FAMILIES and _contains_cjk(text):
                break
            return family
    is_jp = _contains_cjk(text) or any(h in key for h in JP_HINTS)
    is_serif = any(h in key for h in SERIF_HINTS)
    if is_jp:
        return FALLBACK_JP_SERIF if is_serif else FALLBACK_JP_SANS
    if any(h in key for h in MONO_HINTS):
        return FALLBACK_MONO
    return FALLBACK_LATIN_SERIF if is_serif else FALLBACK_LATIN_SANS


def _is_bold(span: dict) -> bool:
    if span["flags"] & FLAG_BOLD:
        return True
    name = (span.get("font") or "").lower()
    if any(h in name for h in ("bold", "heavy", "black", "semibold", "demibold")):
        return True
    # ヒラギノ等のウェイト表記 (W6以上を太字扱い)
    m = re.search(r"w(\d)\b", name)
    return bool(m and int(m.group(1)) >= 6)


def _is_italic(span: dict) -> bool:
    if span["flags"] & FLAG_ITALIC:
        return True
    name = (span.get("font") or "").lower()
    return "italic" in name or "oblique" in name


# ---------------------------------------------------------------------------
# 抽出
# ---------------------------------------------------------------------------

@dataclass
class Line:
    """編集対象として抽出した1行分のテキスト。"""
    bbox: pymupdf.Rect
    spans: list[dict]


@dataclass
class PageResult:
    width: float                # pt
    height: float               # pt
    lines: list[Line] = field(default_factory=list)
    image_png: bytes = b""
    warnings: list[str] = field(default_factory=list)


def _is_garbled(text: str) -> bool:
    if not text:
        return False
    return text.count("�") / len(text) > GARBLED_RATIO


def extract_editable_lines(page: pymupdf.Page) -> tuple[list[Line], list[str]]:
    """横書きで文字化けしていない行だけを編集対象として抽出する。

    縦書き (wmode=1)・回転テキスト・文字化けspanは背景画像に残す。
    """
    lines: list[Line] = []
    warnings: list[str] = []
    skipped_vertical = 0
    skipped_garbled = 0

    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for raw_line in block["lines"]:
            if raw_line.get("wmode", 0) != 0:
                skipped_vertical += 1
                continue
            dx, dy = raw_line["dir"]
            if not (dx > 1.0 - HORIZONTAL_DIR_TOL and abs(dy) < HORIZONTAL_DIR_TOL):
                skipped_vertical += 1
                continue
            spans = []
            for span in raw_line["spans"]:
                if not span["text"].strip():
                    continue
                if _is_garbled(span["text"]):
                    skipped_garbled += 1
                    continue
                spans.append(span)
            if not spans:
                continue
            bbox = pymupdf.Rect(spans[0]["bbox"])
            for span in spans[1:]:
                bbox |= pymupdf.Rect(span["bbox"])
            lines.append(Line(bbox=bbox, spans=spans))

    if skipped_vertical:
        warnings.append(
            f"縦書き・回転テキスト {skipped_vertical} 行は編集対象外です（背景画像に残します）"
        )
    if skipped_garbled:
        warnings.append(
            f"文字コードを復元できないテキスト {skipped_garbled} 箇所を背景画像に残しました"
        )
    return lines, warnings


def redact_text(page: pymupdf.Page, lines: list[Line]) -> None:
    """編集対象テキストを背景から削除する。画像・罫線・図形は残す。"""
    s = REDACT_SHRINK
    for line in lines:
        for span in line.spans:
            r = pymupdf.Rect(span["bbox"])
            if r.width > 2 * s and r.height > 2 * s:
                r = pymupdf.Rect(r.x0 + s, r.y0 + s, r.x1 - s, r.y1 - s)
            # fill=False: 塗り潰さず、下にある画像・図形をそのまま見せる
            page.add_redact_annot(r, fill=False)
    if not lines:
        return
    try:
        page.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_NONE,
            graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
        )
    except TypeError:  # 古いPyMuPDF (graphics引数なし)
        page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE)


def render_background(page: pymupdf.Page, dpi: int) -> bytes:
    """redaction適用後のページをPNG化する。"""
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def process_page(page: pymupdf.Page, page_no: int, dpi: int) -> PageResult:
    result = PageResult(width=page.rect.width, height=page.rect.height)

    has_any_text = bool(page.get_text().strip())
    if not has_any_text:
        if page.get_images():
            result.warnings.append(
                "テキスト層がありません（スキャンPDFの可能性）。背景画像のみ出力します"
            )
        else:
            result.warnings.append("テキストがないページです。背景画像のみ出力します")
    else:
        result.lines, warns = extract_editable_lines(page)
        result.warnings.extend(warns)
        redact_text(page, result.lines)
        # redaction漏れの自己検査: 編集対象の文字が背景に残っていれば座標系の不整合
        leftover = page.get_text()
        ghosts = sum(
            1
            for line in result.lines
            for span in line.spans
            if span["text"].strip() and span["text"].strip() in leftover
        )
        if ghosts:
            result.warnings.append(
                f"redactionで消えなかったテキストが {ghosts} 箇所あります（二重写りの可能性）"
            )

    result.image_png = render_background(page, dpi)
    return result


# ---------------------------------------------------------------------------
# PPTX 生成
# ---------------------------------------------------------------------------

def _pt_to_emu(v: float) -> int:
    return round(v * EMU_PER_PT)


def _set_run_font(run, family: str) -> None:
    """欧文 (a:latin) に加えて日本語 (a:ea) のフォントも設定する。"""
    run.font.name = family  # a:latin
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    anchor = latin
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            anchor.addnext(el)
        el.set("typeface", family)
        anchor = el


def _add_line_textbox(slide, line: Line, scale: float, dx: float, dy: float) -> None:
    """1行分の編集可能テキストボックスをスライドに追加する。"""
    x = line.bbox.x0 * scale + dx
    y = line.bbox.y0 * scale + dy
    w = line.bbox.width * scale + BOX_PAD
    h = line.bbox.height * scale + BOX_PAD

    box = slide.shapes.add_textbox(
        _pt_to_emu(x), _pt_to_emu(y), _pt_to_emu(w), _pt_to_emu(h)
    )
    tf = box.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.TOP

    para = tf.paragraphs[0]
    para.space_before = Pt(0)
    para.space_after = Pt(0)

    for span in line.spans:
        run = para.add_run()
        run.text = span["text"]
        font = run.font
        font.size = Pt(max(span["size"] * scale, 1.0))
        c = span.get("color", 0)
        font.color.rgb = RGBColor((c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF)
        font.bold = _is_bold(span)
        font.italic = _is_italic(span)
        _set_run_font(run, map_font(span.get("font", ""), span["text"]))


def build_pptx(pages: list[PageResult], output: Path) -> list[str]:
    """ページ処理結果からPPTXを組み立てる。戻り値は警告リスト。"""
    warnings: list[str] = []
    slide_w, slide_h = pages[0].width, pages[0].height

    prs = Presentation()
    prs.slide_width = Emu(_pt_to_emu(slide_w))
    prs.slide_height = Emu(_pt_to_emu(slide_h))
    blank = prs.slide_layouts[6]

    for i, pr in enumerate(pages):
        slide = prs.slides.add_slide(blank)

        # ページサイズ混在: 1ページ目基準のスライドに等倍縮小・中央配置で収める
        if (
            abs(pr.width - slide_w) > PAGE_SIZE_TOL
            or abs(pr.height - slide_h) > PAGE_SIZE_TOL
        ):
            scale = min(slide_w / pr.width, slide_h / pr.height)
            warnings.append(
                f"ページ{i + 1}: サイズが1ページ目と異なります"
                f"（{pr.width:.0f}x{pr.height:.0f}pt）。{scale:.0%}に縮小して中央配置します"
            )
        else:
            scale = 1.0
        dx = (slide_w - pr.width * scale) / 2
        dy = (slide_h - pr.height * scale) / 2

        slide.shapes.add_picture(
            io.BytesIO(pr.image_png),
            _pt_to_emu(dx),
            _pt_to_emu(dy),
            width=_pt_to_emu(pr.width * scale),
            height=_pt_to_emu(pr.height * scale),
        )
        for line in pr.lines:
            _add_line_textbox(slide, line, scale, dx, dy)

    prs.save(output)
    return warnings


# ---------------------------------------------------------------------------
# 変換パイプライン / CLI
# ---------------------------------------------------------------------------

def convert(
    input_pdf: Path,
    output_pptx: Path,
    dpi: int = 150,
    debug_dir: Path | None = None,
) -> list[str]:
    """PDFをPPTXに変換する。戻り値は警告メッセージのリスト。"""
    doc = pymupdf.open(input_pdf)
    if doc.needs_pass:
        raise ValueError(f"パスワード付きPDFは扱えません: {input_pdf}")
    if doc.page_count == 0:
        raise ValueError(f"ページがありません: {input_pdf}")

    warnings: list[str] = []
    pages: list[PageResult] = []
    for i, page in enumerate(doc):
        pr = process_page(page, i, dpi)
        warnings.extend(f"ページ{i + 1}: {w}" for w in pr.warnings)
        if debug_dir:
            debug_dir.mkdir(parents=True, exist_ok=True)
            (debug_dir / f"page{i + 1:03d}_bg.png").write_bytes(pr.image_png)
        pages.append(pr)
    doc.close()

    warnings.extend(build_pptx(pages, output_pptx))
    return warnings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="PDFを編集可能なPPTXに変換する (Phase 1 MVP)"
    )
    parser.add_argument("input", type=Path, help="入力PDF")
    parser.add_argument("output", type=Path, help="出力PPTX")
    parser.add_argument("--dpi", type=int, default=150, help="背景画像の解像度 (既定: 150)")
    parser.add_argument(
        "--debug-dir", type=Path, default=None,
        help="redaction後の背景PNGを保存するディレクトリ（検証用）",
    )
    args = parser.parse_args(argv)

    if not args.input.exists():
        print(f"エラー: 入力ファイルがありません: {args.input}", file=sys.stderr)
        return 1

    try:
        warnings = convert(args.input, args.output, dpi=args.dpi, debug_dir=args.debug_dir)
    except Exception as e:  # noqa: BLE001 - CLIの最上位でまとめて報告する
        print(f"エラー: {e}", file=sys.stderr)
        return 1

    for w in warnings:
        print(f"警告: {w}", file=sys.stderr)
    print(f"完了: {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
