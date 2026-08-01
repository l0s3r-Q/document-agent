"""Excel / PDF 功能测试。PDF 转换在无 Office 环境（CI 用 LibreOffice）自动降级。"""

import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.builder import build as build_docx
from excel_toolkit.builder import build as build_excel
from excel_toolkit.parser import parse, to_data
from pdf_toolkit.converter import convert, detect_engines


def test_excel_build_parse(tmp_path):
    """生成带美化的 xlsx 并解析回读。"""
    spec = {"sheets": [{"name": "花名册", "rows": [
        ["姓名", "部门"], ["张三", "研发部"], ["李四", "市场部"]]}]}
    out = str(tmp_path / "花名册.xlsx")
    assert build_excel(spec, out)["ok"]
    data = parse(out)
    assert data["sheets"][0]["name"] == "花名册"
    assert data["sheets"][0]["rows"][0] == ["姓名", "部门"]
    assert len(data["sheets"][0]["rows"]) == 3


def test_excel_to_data(tmp_path):
    """Excel 数据表 → JSON 行数组（首行字段名）。"""
    out = str(tmp_path / "d.xlsx")
    build_excel({"sheets": [{"name": "S", "rows": [
        ["姓名", "部门"], ["张三", "研发部"], ["", ""]]}]}, out)
    d = to_data(out)
    assert d["ok"] and len(d["rows"]) == 1
    assert d["rows"][0] == {"姓名": "张三", "部门": "研发部"}
    assert d["skipped_rows"] == 1


def test_excel_merges_and_widths(tmp_path):
    """合并单元格与列宽。"""
    spec = {"sheets": [{"name": "S", "rows": [["标题", ""], ["a", "b"]],
                        "merges": [{"range": "A1:B1"}],
                        "col_widths": {"A": 20}}]}
    out = str(tmp_path / "m.xlsx")
    assert build_excel(spec, out)["ok"]
    data = parse(out)
    assert "A1:B1" in data["sheets"][0]["merged"]
    assert data["sheets"][0]["col_widths"].get("A") == 20


def test_pdf_convert(tmp_path):
    """docx → PDF 转换 + pdf_info 读取。"""
    if not detect_engines():
        pytest.skip("无可用 PDF 引擎")
    docx = str(tmp_path / "源.docx")
    assert build_docx({"doc_type": "general", "title": "PDF 测试",
                       "sections": [{"type": "paragraph", "text": "PDF 转换测试正文。"}]}, docx)["ok"]
    r = convert(docx)
    assert r["ok"], r
    assert os.path.exists(r["path"]) and os.path.getsize(r["path"]) > 500
    # 用 pypdf 读页数
    from pypdf import PdfReader
    reader = PdfReader(r["path"])
    assert len(reader.pages) >= 1
