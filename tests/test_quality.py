"""质量检查引擎测试：AIGC 痕迹/占位符/emoji/结构。"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.builder import build as build_docx
from docx_toolkit.quality import check_docx, check_excel, check_pptx, quality_check
from excel_toolkit.builder import build as build_excel
from pptx_toolkit.builder import build as build_pptx


def test_quality_detects_issues(tmp_path):
    """问题文档应检出 AIGC 痕迹/占位符/emoji/标点/跳级。"""
    out = str(tmp_path / "bad.docx")
    build_docx({"doc_type": "general", "title": "t", "sections": [
        {"type": "paragraph", "text": "综上所述，方案可行。"},
        {"type": "paragraph", "text": "值得注意的是，{变量}未填。"},
        {"type": "paragraph", "text": "效果显著！！！🎉"},
        {"type": "heading1", "text": "一、标题"},
        {"type": "heading3", "text": "1.1 跳级"}]}, out)
    r = check_docx(out)
    types = {x["type"] for x in r["issues"]}
    assert r["error_count"] >= 1
    assert {"aigc_trace", "placeholder", "emoji", "punctuation", "heading_skip"} <= types


def test_quality_clean_docx(tmp_path):
    out = str(tmp_path / "clean.docx")
    build_docx({"doc_type": "general", "title": "t", "sections": [
        {"type": "paragraph", "text": "本方案已完成实施，效果符合预期。"},
        {"type": "heading1", "text": "一、概述"},
        {"type": "heading2", "text": "1.1 背景"}]}, out)
    r = check_docx(out)
    assert r["error_count"] == 0 and r["pass"]


def test_quality_pptx(tmp_path):
    out = str(tmp_path / "p.pptx")
    build_pptx({"theme": "corporate", "slides": [
        {"type": "cover", "title": "封面"},
        {"type": "content", "title": "内容", "bullets": ["要点"]}]}, out)
    r = check_pptx(out)
    assert r["slides"] == 2 and r["error_count"] == 0


def test_quality_excel(tmp_path):
    out = str(tmp_path / "e.xlsx")
    build_excel({"sheets": [{"name": "S", "rows": [["重复", "重复"], ["1", "2"]]}]}, out)
    r = check_excel(out)
    assert any(x["type"] == "dup_header" for x in r["issues"])


def test_quality_unsupported(tmp_path):
    f = tmp_path / "x.txt"
    f.write_text("hello")
    r = quality_check(str(f))
    assert not r["ok"]
