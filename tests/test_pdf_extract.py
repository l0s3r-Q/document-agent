"""PDF 文本提取测试：pdf_extract_text 提取文本 + 与 build_docx 闭环。"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.builder import build as build_docx
from pdf_toolkit.converter import convert
from pdf_toolkit.converter import extract_text as pdf_extract_text
from pdf_toolkit.converter import pdf_info

TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out_pdf")
os.makedirs(TMP, exist_ok=True)


def test_pdf_extract_text_roundtrip():
    """docx → PDF → 提取文本 → 验证内容存在。"""
    # 1. 生成 docx
    spec = {
        "doc_type": "notice",
        "title": "PDF提取测试",
        "sections": [
            {"type": "paragraph", "text": "这是用于验证PDF文本提取功能的第一段正文，包含消防安全、电气安全等关键词。"},
            {"type": "paragraph", "text": "第二段：各部门应按照检查要求逐项落实，确保隐患整改到位。"},
        ],
    }
    docx = os.path.join(TMP, "src.docx")
    build_docx(spec, docx)
    # 2. 转 PDF
    pdf = os.path.join(TMP, "src.pdf")
    r = convert(docx, pdf)
    assert r.get("ok"), f"PDF 转换失败: {r}"
    # 3. pdf_info
    info = pdf_info(pdf)
    assert info.get("ok") and info["pages"] > 0
    # 4. pdf_extract_text
    res = pdf_extract_text(pdf)
    data = res
    assert data["pages"] > 0
    assert data["total_chars"] > 0
    texts = "".join(p["text"] for p in data.get("text_pages", []))
    assert "PDF提取测试" in texts or "测试" in texts, f"文本未提取到内容: {texts[:100]}"
    # 5. max_pages 限制
    res1 = pdf_extract_text(pdf, max_pages=1)
    assert res1["extracted_pages"] <= 1


def test_pdf_extract_text_missing_file():
    res = pdf_extract_text(os.path.join(TMP, "nope.pdf"))
    assert not res.get("ok")