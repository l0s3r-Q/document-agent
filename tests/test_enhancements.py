"""增强功能测试：suggest_restructure / batch_build / 模板管理 / 新类型。"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "docx-toolkit"))

from docx_toolkit.batch import batch_build
from docx_toolkit.builder import build
from docx_toolkit.parser import parse
from docx_toolkit.restructure import suggest_restructure
from docx_toolkit.templates_store import (
    compare_templates,
    delete_user_template,
    export_template,
    list_templates,
    make_template_meta,
    rename_user_template,
    save_user_template,
)


def test_new_doc_types_roundtrip(tmp_path):
    """5 个新类型均可生成并回读。"""
    for dt in ("legal", "government_report", "techdoc", "resume", "notice"):
        out = str(tmp_path / f"{dt}.docx")
        assert build({"doc_type": dt, "title": "测试", "sections": [
            {"type": "heading1", "text": "第一章 概述"},
            {"type": "paragraph", "text": "正文。"}]}, out)["ok"]
        data = parse(out)
        assert [s["text"] for s in data["structure"] if s["type"] == "heading1"] == ["第一章 概述"]


def test_suggest_restructure_keep_match(tmp_path):
    """近合规文档应识别 keep 章节。"""
    src = str(tmp_path / "src.docx")
    build({"doc_type": "general", "title": "报告", "sections": [
        {"type": "heading1", "text": "一、过去一年工作回顾"},
        {"type": "paragraph", "text": "经济稳步增长。"},
        {"type": "heading1", "text": "三、重点工作安排"},
        {"type": "paragraph", "text": "推进产业升级。"}]}, src)
    r = suggest_restructure(src, "government_report")
    assert r["ok"]
    assert r["summary"]["keep"] >= 2
    assert r["target_doc_type"] == "government_report"


def test_batch_build(tmp_path):
    """批量生成：3 组数据 → 3 个文档，变量替换正确。"""
    tpl = {"doc_type": "notice", "title": "关于召开{meeting}会议的通知",
           "sections": [{"type": "paragraph", "text": "各{dept}：定于{date}召开会议。"}]}
    rows = [{"meeting": "安全生产", "dept": "生产部", "date": "8月10日"},
            {"meeting": "质量评审", "dept": "质检部", "date": "8月15日"},
            {"meeting": "年终总结", "dept": "行政部", "date": "8月20日"}]
    r = batch_build(tpl, rows, str(tmp_path / "batch"), filename_field="meeting")
    assert r["ok"] and r["succeeded"] == 3
    data = parse(os.path.join(str(tmp_path / "batch"), "安全生产.docx"))
    assert any("生产部" in s["text"] for s in data["structure"])


def test_template_management(tmp_path):
    """rename/export/compare/delete 全流程。"""
    save_user_template({"meta": make_template_meta("T-A", "user", "user"), "page": {},
                        "styles": {}, "skeleton": []}, "T-A")
    save_user_template({"meta": make_template_meta("T-B", "user", "user"),
                        "page": {"top_cm": 1.0}, "styles": {"body": {"size_pt": 16}},
                        "skeleton": [{"level": 1, "text": "甲章"}]}, "T-B")
    assert rename_user_template("T-A", "T-A2") is not None
    assert any(t["name"] == "T-A2" for t in list_templates())
    assert export_template("T-A2", str(tmp_path / "t.json")) is not None
    c = compare_templates("T-A2", "T-B")
    assert c["ok"] and c["diff"]["page"] and c["diff"]["skeleton"]["b_count"] == 1
    assert delete_user_template("T-A2") and delete_user_template("T-B")
    assert not any(t["name"] in ("T-A2", "T-B") for t in list_templates())
