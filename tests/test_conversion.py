"""互转测试：docx↔xlsx 表格、docx→Markdown。"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.builder import build as build_docx
from docx_toolkit.markdown import to_markdown
from docx_toolkit.parser import parse as parse_docx
from excel_toolkit.builder import build as build_excel
from excel_toolkit.parser import parse as parse_excel
from server import docx_tables_to_excel, docx_to_markdown, excel_to_docx


def test_docx_tables_to_excel(tmp_path):
    docx = str(tmp_path / "d.docx")
    build_docx({"doc_type": "general", "title": "t", "sections": [
        {"type": "table", "rows": [["月份", "工单数"], ["4月", "91"]]}]}, docx)
    out = str(tmp_path / "o.xlsx")
    r = docx_tables_to_excel(docx, out)
    assert '"ok": true' in r
    d = parse_excel(out)
    assert d["sheets"][0]["rows"][0] == ["月份", "工单数"]


def test_excel_to_docx(tmp_path):
    xlsx = str(tmp_path / "s.xlsx")
    build_excel({"sheets": [{"name": "数据", "rows": [["A", "B"], ["1", "2"]]}]}, xlsx)
    out = str(tmp_path / "o.docx")
    r = excel_to_docx(xlsx, out)
    assert '"ok": true' in r
    d = parse_docx(out)
    tbl = [s for s in d["structure"] if s["type"] == "table"]
    assert tbl and tbl[0]["rows"][0] == ["A", "B"]


def test_docx_to_markdown(tmp_path):
    docx = str(tmp_path / "d.docx")
    build_docx({"doc_type": "general", "title": "标题", "sections": [
        {"type": "heading1", "text": "一、章"},
        {"type": "table", "rows": [["列1", "列2"], ["v1", "v2"]]},
        {"type": "paragraph", "text": "正文"}]}, docx)
    out = str(tmp_path / "o.md")
    r = docx_to_markdown(docx, out)
    assert '"ok": true' in r
    md = open(out, encoding="utf-8").read()
    assert "# 标题" in md and "## 一、章" in md and "| 列1 | 列2 |" in md
