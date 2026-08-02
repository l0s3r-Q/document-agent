"""新功能测试：TOC/图片/Excel 图表/PDF 合并/表格合并。"""

import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.builder import build as build_docx
from excel_toolkit.builder import build as build_excel
from pdf_toolkit.converter import convert, merge_pdfs


def test_docx_toc(tmp_path):
    out = str(tmp_path / "toc.docx")
    r = build_docx({"doc_type": "general", "title": "t", "sections": [
        {"type": "toc"}, {"type": "heading1", "text": "第一章"}]}, out)
    assert r["ok"]
    with zipfile.ZipFile(out) as z:
        assert "TOC" in z.read("word/document.xml").decode("utf-8")


def test_docx_image(tmp_path):
    """图片插入 + 缺失路径警告。"""
    from PIL import Image
    img = tmp_path / "t.png"
    Image.new("RGB", (100, 50), "blue").save(img)
    out = str(tmp_path / "img.docx")
    r = build_docx({"doc_type": "general", "title": "t", "sections": [
        {"type": "image", "path": str(img), "caption": "图1"},
        {"type": "image", "path": "Z:/无.png"}]}, out)
    assert r["ok"] and r.get("warnings")
    with zipfile.ZipFile(out) as z:
        assert any("media" in n for n in z.namelist())


def test_excel_chart(tmp_path):
    out = str(tmp_path / "c.xlsx")
    r = build_excel({"sheets": [{"name": "S", "rows": [["月", "数"], ["1", 10], ["2", 20]],
                                 "charts": [{"type": "bar", "title": "T"}]}]}, out)
    assert r["ok"]
    with zipfile.ZipFile(out) as z:
        assert any("charts" in n for n in z.namelist())


def test_pdf_merge(tmp_path):
    """合并两个 PDF + 缺失文件警告。"""
    pdfs = []
    for name in ("a", "b"):
        docx = str(tmp_path / f"{name}.docx")
        build_docx({"doc_type": "general", "title": name, "sections": []}, docx)
        assert convert(docx)["ok"]
        pdfs.append(str(tmp_path / f"{name}.pdf"))
    out = str(tmp_path / "m.pdf")
    r = merge_pdfs(pdfs, out)
    assert r["ok"] and r["pages"] == 2
    r2 = merge_pdfs([pdfs[0], "Z:/无.pdf"], str(tmp_path / "m2.pdf"))
    assert r2["ok"] and r2.get("warnings")


def test_docx_table_merge(tmp_path):
    out = str(tmp_path / "t.docx")
    r = build_docx({"doc_type": "general", "title": "t", "sections": [
        {"type": "table", "rows": [["标题", ""], ["A", "B"]],
         "merges": [{"from": "A1", "to": "B1"}]}]}, out)
    assert r["ok"]
    with zipfile.ZipFile(out) as z:
        assert "gridSpan" in z.read("word/document.xml").decode("utf-8")
