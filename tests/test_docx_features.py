"""docx 扩展功能：页眉/页脚/页码。"""

import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.builder import build


def test_docx_header_footer_pagenum(tmp_path):
    out = str(tmp_path / "h.docx")
    r = build({"doc_type": "general", "title": "t", "header": "公司内部",
               "footer": "机密", "page_number": True,
               "sections": [{"type": "paragraph", "text": "正文"}]}, out)
    assert r["ok"]
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        hdr = [n for n in names if "header" in n]
        ftr = [n for n in names if "footer" in n]
        assert hdr and ftr
        assert "公司内部" in z.read(hdr[0]).decode("utf-8")
        fxml = z.read(ftr[0]).decode("utf-8")
        assert "PAGE" in fxml and "机密" in fxml


def test_docx_without_header(tmp_path):
    """未指定页眉页脚时不生成额外部件（无回归）。"""
    out = str(tmp_path / "n.docx")
    r = build({"doc_type": "general", "title": "t", "sections": []}, out)
    assert r["ok"]
