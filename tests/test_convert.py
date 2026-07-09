"""convert.py のテスト（合成PDFのみ・ネットワーク不使用）。"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pymupdf
import pytest
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Emu

import convert
import fixtures_gen as fx

ROOT = Path(__file__).resolve().parents[1]
EMU_PER_PT = 12700


# ---------------------------------------------------------------------------
# fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def business(tmp_path_factory):
    """business.pdf を1回だけ変換し、(pdf, pptx, warnings) を返す。"""
    d = tmp_path_factory.mktemp("business")
    pdf = d / "business.pdf"
    pptx = d / "business.pptx"
    fx.make_business_pdf(pdf)
    warnings = convert.convert(pdf, pptx)
    return pdf, pptx, warnings


def _textbox_map(slide):
    """スライド内のテキストボックスを {テキスト: shape} で返す。"""
    return {
        s.text_frame.text: s
        for s in slide.shapes
        if s.has_text_frame and s.text_frame.text
    }


def _find_box(slide, text):
    boxes = _textbox_map(slide)
    assert text in boxes, f"テキストボックスが見つからない: {text!r} in {list(boxes)}"
    return boxes[text]


# ---------------------------------------------------------------------------
# 基本構造
# ---------------------------------------------------------------------------

def test_slide_size_matches_pdf_page(business):
    pdf, pptx, _ = business
    doc = pymupdf.open(pdf)
    prs = Presentation(pptx)
    assert len(prs.slides._sldIdLst) == doc.page_count == 3
    assert prs.slide_width == round(doc[0].rect.width * EMU_PER_PT)
    assert prs.slide_height == round(doc[0].rect.height * EMU_PER_PT)
    doc.close()


def test_every_slide_has_background_picture(business):
    _, pptx, _ = business
    prs = Presentation(pptx)
    for slide in prs.slides:
        pics = [s for s in slide.shapes if s.shape_type == 13]  # PICTURE
        assert len(pics) == 1
        assert pics[0].left == 0 and pics[0].top == 0


# ---------------------------------------------------------------------------
# テキストの再現
# ---------------------------------------------------------------------------

def test_expected_texts_are_editable(business):
    _, pptx, _ = business
    slide = Presentation(pptx).slides[0]
    boxes = _textbox_map(slide)
    for expected in [fx.TITLE, fx.CONFIDENTIAL, fx.NOTE, *fx.BODY_LINES,
                     "項目", "売上高", "1,234", "営業利益"]:
        assert any(expected in t for t in boxes), f"編集不可: {expected!r}"


def test_position_within_tolerance(business):
    """PPTX上のボックス位置がPDFの行bboxと±3pt以内で一致する。"""
    pdf, pptx, _ = business
    doc = pymupdf.open(pdf)
    lines, _ = convert.extract_editable_lines(doc[0])
    doc.close()
    slide = Presentation(pptx).slides[0]
    boxes = _textbox_map(slide)

    checked = 0
    for line in lines:
        text = "".join(s["text"] for s in line.spans)
        if text not in boxes:
            continue
        shape = boxes[text]
        assert abs(shape.left / EMU_PER_PT - line.bbox.x0) < 3.0, text
        assert abs(shape.top / EMU_PER_PT - line.bbox.y0) < 3.0, text
        checked += 1
    assert checked >= 8  # 主要な行が検査されていること


def test_font_size_color_bold(business):
    _, pptx, _ = business
    slide = Presentation(pptx).slides[0]

    title = _find_box(slide, fx.TITLE).text_frame.paragraphs[0].runs[0]
    assert abs(title.font.size.pt - 20) < 0.5
    assert str(title.font.color.rgb) == "1F3864"
    assert title.font.bold is False

    conf = _find_box(slide, fx.CONFIDENTIAL).text_frame.paragraphs[0].runs[0]
    assert conf.font.bold is True
    assert abs(conf.font.size.pt - 11) < 0.5


def test_textbox_settings(business):
    """内部マージン0・word_wrap無効・上詰めアンカーになっている。"""
    _, pptx, _ = business
    slide = Presentation(pptx).slides[0]
    tf = _find_box(slide, fx.TITLE).text_frame
    assert (tf.margin_left, tf.margin_right, tf.margin_top, tf.margin_bottom) \
        == (Emu(0),) * 4
    assert tf.word_wrap is False
    assert tf.vertical_anchor == MSO_ANCHOR.TOP


def test_ea_font_is_set_for_japanese(business):
    """日本語run に a:ea フォントが設定されている。"""
    _, pptx, _ = business
    from pptx.oxml.ns import qn
    slide = Presentation(pptx).slides[0]
    run = _find_box(slide, fx.TITLE).text_frame.paragraphs[0].runs[0]
    ea = run._r.get_or_add_rPr().find(qn("a:ea"))
    assert ea is not None and ea.get("typeface")


# ---------------------------------------------------------------------------
# 除外対象・検出
# ---------------------------------------------------------------------------

def test_rotated_text_left_in_background(business):
    """回転テキストは編集ボックスにならず、警告が出る。"""
    _, pptx, warnings = business
    slide2 = Presentation(pptx).slides[1]
    boxes = _textbox_map(slide2)
    assert not any(fx.ROTATED_TEXT in t for t in boxes)
    assert any(fx.PAGE2_BODY in t for t in boxes)
    assert any("縦書き・回転テキスト" in w for w in warnings)


def test_scanned_page_detected(business):
    """スキャン風ページは背景画像のみ+警告。"""
    _, pptx, warnings = business
    slide3 = Presentation(pptx).slides[2]
    assert len(slide3.shapes) == 1  # 背景画像のみ
    assert any("スキャン" in w for w in warnings)


def test_no_redaction_ghosts(business):
    """redaction漏れ（二重写り）の自己検査警告が出ていない。"""
    _, _, warnings = business
    assert not any("redaction" in w for w in warnings)


def test_mixed_page_sizes_warns(tmp_path):
    pdf = tmp_path / "mixed.pdf"
    pptx = tmp_path / "mixed.pptx"
    fx.make_mixed_pdf(pdf)
    warnings = convert.convert(pdf, pptx)
    assert any("サイズが1ページ目と異なります" in w for w in warnings)
    prs = Presentation(pptx)
    assert len(prs.slides._sldIdLst) == 2
    # 2ページ目のテキストも（縮小されつつ）編集可能であること
    assert any(fx.MIXED_P2_TEXT in t for t in _textbox_map(prs.slides[1]))


def test_page_with_rotate_attribute(tmp_path):
    """/Rotate 90 付きページでも座標系が破綻しない（redaction漏れなし）。"""
    pdf = tmp_path / "rot.pdf"
    pptx = tmp_path / "rot.pptx"
    fx.make_rotated_page_pdf(pdf)
    warnings = convert.convert(pdf, pptx)
    assert not any("redaction" in w for w in warnings)
    slide = Presentation(pptx).slides[0]
    assert any(fx.ROTATED_PAGE_TEXT in t for t in _textbox_map(slide))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def test_cli_smoke(tmp_path):
    pdf = tmp_path / "in.pdf"
    out = tmp_path / "out.pptx"
    fx.make_business_pdf(pdf)
    res = subprocess.run(
        [sys.executable, str(ROOT / "convert.py"), str(pdf), str(out),
         "--debug-dir", str(tmp_path / "debug")],
        capture_output=True, text=True,
    )
    assert res.returncode == 0, res.stderr
    assert out.exists()
    assert (tmp_path / "debug" / "page001_bg.png").exists()
    assert "完了" in res.stdout


def test_cli_missing_input(tmp_path):
    res = subprocess.run(
        [sys.executable, str(ROOT / "convert.py"),
         str(tmp_path / "nai.pdf"), str(tmp_path / "out.pptx")],
        capture_output=True, text=True,
    )
    assert res.returncode == 1
