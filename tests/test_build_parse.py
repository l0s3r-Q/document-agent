"""构建/解析回读测试：生成 → 解析 → 断言对称性。"""

import json
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "docx-toolkit"))

from docx_toolkit.builder import build
from docx_toolkit.parser import extract_structure, parse

SPEC = {
    "doc_type": "official",
    "title": "关于开展安全生产大检查的通知",
    "sections": [
        {"type": "heading1", "text": "一、检查范围"},
        {"type": "paragraph", "text": "本次检查覆盖全市所有在建工程项目。"},
        {"type": "heading2", "text": "（一）重点领域"},
        {"type": "paragraph", "text": "建筑施工、交通运输、消防安全。"},
        {"type": "table", "rows": [["阶段", "时间"], ["自查", "8月"], ["督查", "9月"]]},
    ],
}


@pytest.fixture()
def docx_path(tmp_path):
    out = str(tmp_path / "test.docx")
    assert build(SPEC, out)["ok"]
    return out


def test_build_and_parse_roundtrip(docx_path):
    data = parse(docx_path)
    heads = [s["text"] for s in data["structure"] if s["type"].startswith("heading")]
    assert heads == ["一、检查范围", "（一）重点领域"]


def test_page_settings(docx_path):
    data = parse(docx_path)
    assert data["page"]["top_cm"] == 3.7
    assert data["page"]["bottom_cm"] == 3.5
    assert data["page"]["left_cm"] == 2.8
    assert data["page"]["right_cm"] == 2.6


def test_extract_structure(docx_path):
    tree = extract_structure(docx_path)
    assert [x["level"] for x in tree["outline"]] == [1, 2]
    assert tree["outline"][0]["text"] == "一、检查范围"


def test_table_roundtrip(docx_path):
    data = parse(docx_path)
    tables = [s for s in data["structure"] if s["type"] == "table"]
    assert len(tables) == 1
    assert tables[0]["rows"][0] == ["阶段", "时间"]


def test_custom_styles_override(docx_path):
    """spec.styles 按 role 覆盖应生效。"""
    spec = {
        "doc_type": "general",
        "title": "测试",
        "styles": {"body": {"font_name": "楷体", "size_pt": 14, "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5}},
        "sections": [{"type": "paragraph", "text": "正文测试"}],
    }
    import tempfile
    out = os.path.join(tempfile.gettempdir(), "test_styles_override.docx")
    build(spec, out)
    data = parse(out)
    body = [s for s in data["structure"] if s["type"] == "paragraph" and s["text"] == "正文测试"]
    assert body, "未找到正文段落"
    assert body[0].get("size_pt") == 14.0
    assert body[0].get("font_name") == "楷体"
