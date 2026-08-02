"""PPT 功能测试：build_pptx / parse_pptx。"""

import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from pptx_toolkit.builder import build
from pptx_toolkit.parser import parse

SPEC = {
    "title": "测试演示", "subtitle": "副标题", "author": "测试",
    "theme": "corporate", "size": "16:9",
    "slides": [
        {"type": "cover", "title": "测试演示", "subtitle": "副标题"},
        {"type": "agenda", "items": ["一、概述", "二、详情"]},
        {"type": "section", "title": "一、概述"},
        {"type": "content", "title": "要点页", "bullets": ["要点一", "> 子要点"]},
        {"type": "two_column", "title": "对比", "left": {"title": "A", "bullets": ["a1"]},
         "right": {"title": "B", "bullets": ["b1"]}},
        {"type": "table", "title": "数据", "rows": [["列1", "列2"], ["v1", "v2"]]},
        {"type": "closing", "title": "谢谢"},
    ],
}


def test_pptx_build_parse_roundtrip(tmp_path):
    out = str(tmp_path / "test.pptx")
    r = build(SPEC, out)
    assert r["ok"] and r["slides"] == 7
    data = parse(out)
    assert len(data["slides"]) == 7
    assert data["slide_size"]["width_inches"] == 13.33


def test_pptx_themes(tmp_path):
    """4 个主题均可生成。"""
    for theme in ("corporate", "academic", "launch", "minimal"):
        spec = dict(SPEC, theme=theme)
        out = str(tmp_path / f"{theme}.pptx")
        r = build(spec, out)
        assert r["ok"] and r["theme"] == theme


def test_pptx_table_roundtrip(tmp_path):
    out = str(tmp_path / "t.pptx")
    build(SPEC, out)
    data = parse(out)
    tbl_page = data["slides"][5]
    tbl = [sh["table"] for sh in tbl_page["shapes"] if "table" in sh]
    assert tbl and tbl[0][0] == ["列1", "列2"]


def test_pptx_4x3_and_errors(tmp_path):
    """4:3 尺寸 + 非法后缀拒绝。"""
    spec = dict(SPEC, size="4:3")
    out = str(tmp_path / "s43.pptx")
    assert build(spec, out)["ok"]
    data = parse(out)
    assert data["slide_size"]["width_inches"] == 10.0
    # 非法后缀
    r = build(SPEC, str(tmp_path / "bad.xls"))
    assert not r["ok"]


def test_pptx_enhanced_layouts(tmp_path):
    """视觉增强版式：stats 卡片 / cards 卡片化 / section 编号。"""
    spec = {"title": "增强", "theme": "corporate", "slides": [
        {"type": "section", "title": "一、数据", "index": "01"},
        {"type": "stats", "title": "数据页", "stats": [
            {"value": "287", "label": "工单", "sub": "月均 48"},
            {"value": "17", "label": "客户"}]},
        {"type": "content", "title": "职责", "cards": [
            {"title": "A", "bullets": ["a1"], "color": "2E75B6"},
            {"title": "B", "bullets": ["b1"]}]},
    ]}
    out = str(tmp_path / "enhanced.pptx")
    r = build(spec, out)
    assert r["ok"] and r["slides"] == 3
    data = parse(out)
    # stats 页应含大数字
    stats_page = data["slides"][1]
    texts = [t for sh in stats_page["shapes"] if sh.get("texts") for t in sh["texts"]]
    assert "287" in texts and "17" in texts
    # section 页含编号
    sec_page = data["slides"][0]
    sec_texts = [t for sh in sec_page["shapes"] if sh.get("texts") for t in sh["texts"]]
    assert "01" in sec_texts
