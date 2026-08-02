"""健壮性测试：错误路径/边界/异常输入（测试驱动维护）。"""

import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.batch import batch_build
from docx_toolkit.builder import build as build_docx
from docx_toolkit.parser import parse as parse_docx
from excel_toolkit.builder import build as build_excel
from excel_toolkit.parser import parse as parse_excel
from pdf_toolkit.converter import convert
from pptx_toolkit.builder import build as build_pptx
from pptx_toolkit.parser import parse as parse_pptx


# ── docx 错误路径 ────────────────────────────────────────────
def test_docx_invalid_output_suffix(tmp_path):
    r = build_docx({"doc_type": "general", "title": "t", "sections": []}, str(tmp_path / "bad.txt"))
    assert not r["ok"]


def test_docx_empty_sections(tmp_path):
    r = build_docx({"doc_type": "general", "title": "t", "sections": []}, str(tmp_path / "e.docx"))
    assert r["ok"]  # 空 sections 应允许（仅标题）


def test_parse_docx_missing_file():
    """parse 库函数抛异常（server 层负责 {ok:false} 封装）。"""
    with pytest.raises(Exception):
        parse_docx("Z:/不存在/文件.docx")


def test_parse_docx_not_docx(tmp_path):
    f = tmp_path / "fake.docx"
    f.write_bytes(b"not a zip file")
    with pytest.raises(Exception):
        parse_docx(str(f))


def test_docx_special_chars(tmp_path):
    """特殊字符（引号/换行/emoji）往返。"""
    out = str(tmp_path / "sp.docx")
    r = build_docx({"doc_type": "general", "title": '标题"引号"',
                    "sections": [{"type": "paragraph", "text": "换行\n第二行 emoji 🚀 符号&<>"}]}, out)
    assert r["ok"]
    d = parse_docx(out)
    assert '"引号"' in d["structure"][0]["text"]


# ── excel 错误路径 ───────────────────────────────────────────
def test_excel_empty_sheets(tmp_path):
    r = build_excel({"sheets": []}, str(tmp_path / "e.xlsx"))
    assert not r["ok"]


def test_excel_invalid_suffix(tmp_path):
    r = build_excel({"sheets": [{"rows": [["a"]]}]}, str(tmp_path / "b.xls"))
    assert not r["ok"]


def test_excel_empty_first_row(tmp_path):
    """空首行不崩溃（列数取最大行宽）。"""
    r = build_excel({"sheets": [{"rows": [[], ["a", "b"]]}]}, str(tmp_path / "e.xlsx"))
    assert r["ok"]


def test_parse_excel_missing():
    with pytest.raises(Exception):
        parse_excel("Z:/不存在.xlsx")


# ── pptx 错误路径 ────────────────────────────────────────────
def test_pptx_empty_slides(tmp_path):
    r = build_pptx({"title": "t", "slides": []}, str(tmp_path / "e.pptx"))
    assert not r["ok"]


def test_pptx_invalid_suffix(tmp_path):
    r = build_pptx({"slides": [{"type": "cover"}]}, str(tmp_path / "b.ppt"))
    assert not r["ok"]


def test_pptx_bad_theme_fallback(tmp_path):
    """非法主题回退默认，不崩溃。"""
    r = build_pptx({"theme": "不存在", "slides": [{"type": "cover", "title": "t"}]},
                   str(tmp_path / "t.pptx"))
    assert r["ok"] and r["theme"] == "corporate"


def test_pptx_image_missing_warning(tmp_path):
    r = build_pptx({"slides": [{"type": "image", "title": "图", "path": "Z:/无.png"}]},
                   str(tmp_path / "i.pptx"))
    assert r["ok"] and r.get("warnings")


def test_pptx_super_long_title(tmp_path):
    """超长标题不崩溃（自动缩字号）。"""
    long_title = "超长标题" * 30
    r = build_pptx({"slides": [{"type": "content", "title": long_title, "bullets": ["a"]}]},
                   str(tmp_path / "l.pptx"))
    assert r["ok"]


def test_parse_pptx_missing():
    with pytest.raises(Exception):
        parse_pptx("Z:/不存在.pptx")


# ── batch 错误路径 ───────────────────────────────────────────
def test_batch_empty_rows(tmp_path):
    r = batch_build({"title": "t"}, [], str(tmp_path / "b"))
    assert not r["ok"]


def test_batch_non_dict_row(tmp_path):
    r = batch_build({"title": "t"}, [["不是对象"]], str(tmp_path / "b"))
    assert not r["ok"]


def test_batch_missing_var_warning(tmp_path):
    """缺变量 → warnings 提示，不崩溃。"""
    tpl = {"title": "关于{会议}的通知",
           "sections": [{"type": "paragraph", "text": "正文{未知变量}"}]}
    r = batch_build(tpl, [{"会议": "安全"}], str(tmp_path / "b"))
    assert r["ok"] and r["succeeded"] == 1
    assert r["warnings"] and "未知变量" in r["warnings"][0]


# ── pdf 错误路径 ─────────────────────────────────────────────
def test_pdf_missing_source():
    r = convert("Z:/不存在.docx")
    assert not r["ok"]


def test_pdf_not_docx(tmp_path):
    f = tmp_path / "t.txt"
    f.write_text("hello")
    r = convert(str(f))
    assert not r["ok"]


def test_pdf_invalid_output_suffix(tmp_path):
    docx = str(tmp_path / "s.docx")
    build_docx({"title": "t", "sections": []}, docx)
    r = convert(docx, str(tmp_path / "out.txt"))
    assert not r["ok"]
