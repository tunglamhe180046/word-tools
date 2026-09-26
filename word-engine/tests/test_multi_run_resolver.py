"""
Tests for core/text_resolver.py - Multi-Run Text Range Resolver (see
docs/plans/DEDICATED_WORD_ENGINE_TOOL_PLAN.md section 3.2 and core/text_resolver.py's module
docstring for the design, especially the length-equality condition chosen for Edit Fidelity).

Paragraphs are hand-built as bare lxml w:p elements (not via python-docx) because
resolve_and_replace_text() operates on a single already-located w:p element, and several cases
need OOXML structures python-docx has no API for (w:hyperlink, w:sdt, w:tab as a bare run child).
No mocking, no files on disk - matches tests/test_locators.py's fixture style, adapted to this
module's narrower (single-paragraph-element) input.
"""
from __future__ import annotations

import pytest
from lxml import etree

from core.inspector import WORD_NS
from core.text_resolver import (
    ComplexSpanError,
    StyleBoundaryConflictError,
    TextNotFoundError,
    resolve_and_replace_text,
)

W = f"{{{WORD_NS}}}"


def _p(*children: etree._Element) -> etree._Element:
    p = etree.Element(f"{W}p")
    for child in children:
        p.append(child)
    return p


def _run(text: str, bold: bool = False) -> etree._Element:
    r = etree.Element(f"{W}r")
    if bold:
        rpr = etree.SubElement(r, f"{W}rPr")
        etree.SubElement(rpr, f"{W}b")
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    return r


def _run_with_extra_child(text: str, extra_tag: str) -> etree._Element:
    """A run containing an extra child besides w:rPr/w:t (w:tab, w:br, w:drawing...) - violates
    SAFE_TEXT_RUN on purpose."""
    r = _run(text)
    etree.SubElement(r, f"{W}{extra_tag}")
    return r


def _run_texts(p: etree._Element) -> list[str]:
    return ["".join(t.text or "" for t in r.findall(f"{W}t")) for r in p.findall(f"{W}r")]


def _paragraph_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


# ---------------------------------------------------------------------------
# Single-run replace
# ---------------------------------------------------------------------------
def test_single_run_replace_keeps_style_and_sets_xml_space_preserve():
    run = _run("Ha Noi", bold=True)
    p = _p(run)

    result = resolve_and_replace_text(p, "Ha Noi", " Sai Gon")

    assert result.run_mode == "single_run"
    assert result.start_run_index == 0
    assert result.end_run_index == 0
    assert result.runs_removed == 0

    runs = p.findall(f"{W}r")
    assert len(runs) == 1
    t = runs[0].find(f"{W}t")
    assert t.text == " Sai Gon"
    assert t.get(f"{{http://www.w3.org/XML/1998/namespace}}space") == "preserve"
    # style preserved
    assert runs[0].find(f"{W}rPr/{W}b") is not None


def test_single_run_replace_without_edge_whitespace_has_no_preserve_attr():
    run = _run("Ha Noi")
    p = _p(run)

    resolve_and_replace_text(p, "Noi", "Phong")

    t = p.find(f"{W}r/{W}t")
    assert t.text == "Ha Phong"
    assert t.get(f"{{http://www.w3.org/XML/1998/namespace}}space") is None


# ---------------------------------------------------------------------------
# Multi-run, same style
# ---------------------------------------------------------------------------
def test_multi_run_same_style_keeps_prefix_and_suffix():
    run_a = _run("Hello ")
    run_b = _run("World")
    run_c = _run(" Bye")
    p = _p(run_a, run_b, run_c)

    # "lo Wor" spans the tail of run_a ("lo ") and the head of run_b ("Wor").
    result = resolve_and_replace_text(p, "lo Wor", "XY")

    assert result.run_mode == "same_style_multi_run"
    assert result.start_run_index == 0
    assert result.end_run_index == 1
    assert result.runs_removed == 0
    assert _run_texts(p) == ["HelXY", "ld", " Bye"]
    assert _paragraph_text(p) == "HelXYld Bye"


def test_multi_run_same_style_removes_middle_runs():
    run1 = _run("Xin ")
    run2 = _run("chao ")
    run3 = _run("cac ")
    run4 = _run("ban")
    p = _p(run1, run2, run3, run4)

    # full text: "Xin chao cac ban" -> span [2:15) = "n chao cac ba", crossing all 4 runs.
    full_text = "Xin chao cac ban"
    search_text = full_text[2:15]
    assert search_text == "n chao cac ba"

    result = resolve_and_replace_text(p, search_text, "Y")

    assert result.run_mode == "same_style_multi_run"
    assert result.runs_removed == 2
    remaining = p.findall(f"{W}r")
    assert len(remaining) == 2
    assert _run_texts(p) == ["XiY", "n"]
    assert _paragraph_text(p) == "XiYn"


def test_multi_run_same_style_removes_end_run_when_suffix_empty():
    run_a = _run("Chao ")
    run_b = _run("ban")
    p = _p(run_a, run_b)

    result = resolve_and_replace_text(p, "o ban", "!")

    assert result.run_mode == "same_style_multi_run"
    assert result.runs_removed == 1
    assert _run_texts(p) == ["Cha!"]


# ---------------------------------------------------------------------------
# COMPLEX_SPAN fail-closed
# ---------------------------------------------------------------------------
def test_complex_span_hyperlink_raises():
    before = _run("Xem ")
    linked = _run("tai day")
    hyperlink = etree.Element(f"{W}hyperlink")
    hyperlink.append(linked)
    after = _run(" nhe")
    p = _p(before, hyperlink, after)

    with pytest.raises(ComplexSpanError):
        resolve_and_replace_text(p, "tai day", "here")


def test_complex_span_sdt_raises():
    before = _run("Ten: ")
    content_run = _run("Nguyen Van A")
    sdt = etree.Element(f"{W}sdt")
    sdt_content = etree.SubElement(sdt, f"{W}sdtContent")
    sdt_content.append(content_run)
    p = _p(before, sdt)

    with pytest.raises(ComplexSpanError):
        resolve_and_replace_text(p, "Nguyen Van A", "Tran Thi B")


@pytest.mark.parametrize("extra_tag", ["br", "tab", "drawing"])
def test_complex_span_special_run_children_raise(extra_tag):
    before = _run("Truoc ")
    unsafe = _run_with_extra_child("giua", extra_tag)
    after = _run(" sau")
    p = _p(before, unsafe, after)

    with pytest.raises(ComplexSpanError):
        resolve_and_replace_text(p, "giua", "GIUA")


# ---------------------------------------------------------------------------
# Heterogeneous rPr - Edit Fidelity vs STYLE_BOUNDARY_CONFLICT vs style_policy
# ---------------------------------------------------------------------------
def _ha_loi_paragraph() -> etree._Element:
    # "Ha L" normal (includes the "L" of "Loi") + "oi" bold -> full text "Ha Loi", "Loi" spans
    # both runs with different w:rPr.
    run_normal = _run("Ha L")
    run_bold = _run("oi", bold=True)
    return _p(run_normal, run_bold)


def test_heterogeneous_style_edit_fidelity_preserves_bold_tail():
    p = _ha_loi_paragraph()

    result = resolve_and_replace_text(p, "Loi", "Noi")

    assert result.run_mode == "edit_fidelity"
    assert _run_texts(p) == ["Ha N", "oi"]
    assert _paragraph_text(p) == "Ha Noi"
    # bold run keeps its own rPr untouched
    runs = p.findall(f"{W}r")
    assert runs[1].find(f"{W}rPr/{W}b") is not None
    assert runs[0].find(f"{W}rPr") is None


def test_heterogeneous_style_length_mismatch_without_policy_raises():
    p = _ha_loi_paragraph()

    with pytest.raises(StyleBoundaryConflictError):
        resolve_and_replace_text(p, "Loi", "Noii")


def test_heterogeneous_style_prefer_start_collapses_to_start_style():
    p = _ha_loi_paragraph()

    result = resolve_and_replace_text(p, "Loi", "Noii", style_policy="prefer-start")

    assert result.run_mode == "collapsed_prefer_start"
    assert result.runs_removed == 1
    assert result.style_policy_applied == "prefer-start"
    runs = p.findall(f"{W}r")
    assert len(runs) == 1
    assert "".join(t.text or "" for t in runs[0].findall(f"{W}t")) == "Ha Noii"
    assert runs[0].find(f"{W}rPr/{W}b") is None


def test_heterogeneous_style_prefer_end_collapses_to_end_style():
    p = _ha_loi_paragraph()

    result = resolve_and_replace_text(p, "Loi", "Noii", style_policy="prefer-end")

    assert result.run_mode == "collapsed_prefer_end"
    assert result.runs_removed == 1
    assert result.style_policy_applied == "prefer-end"
    runs = p.findall(f"{W}r")
    assert len(runs) == 1
    assert "".join(t.text or "" for t in runs[0].findall(f"{W}t")) == "Ha Noii"
    assert runs[0].find(f"{W}rPr/{W}b") is not None


# ---------------------------------------------------------------------------
# Not found
# ---------------------------------------------------------------------------
def test_text_not_found_raises():
    p = _p(_run("Xin chao"))

    with pytest.raises(TextNotFoundError):
        resolve_and_replace_text(p, "khong ton tai", "gi do")
