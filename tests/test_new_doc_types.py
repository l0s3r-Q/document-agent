"""新增文档类型测试：meeting_minutes / speech / proposal / invitation 模板可生成 + 解析。"""

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.builder import build as build_docx
from docx_toolkit.parser import parse

TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out_new_types")
os.makedirs(TMP, exist_ok=True)

NEW_TYPES = ["meeting_minutes", "speech", "proposal", "invitation"]


def _spec(doc_type: str) -> dict:
    return {
        "doc_type": doc_type,
        "title": f"测试{doc_type}文档",
        "sections": [
            {"type": "heading1", "text": "一、测试章节"},
            {"type": "paragraph", "text": "这是一段用于验证新文档类型生成的测试正文内容，确保排版正确。"},
            {"type": "list", "items": ["要点一", "要点二"]},
        ],
    }


def test_new_types_build_and_parse():
    for dt in NEW_TYPES:
        out = os.path.join(TMP, f"{dt}.docx")
        build_docx(_spec(dt), out)
        assert os.path.exists(out) and os.path.getsize(out) > 1000, f"{dt} 生成失败"
        # 解析回读：应包含标题与正文
        data = parse(out)
        text = json.dumps(data, ensure_ascii=False)
        assert "测试" in text and "测试章节" in text, f"{dt} 解析内容缺失"


def test_new_types_in_template_list():
    from docx_toolkit.templates_store import get_builtin, list_templates
    for dt in NEW_TYPES:
        t = get_builtin(dt)
        assert t is not None, f"{dt} 模板未注册"
        assert t["meta"]["doc_type"] == dt
    # list_templates 应包含新类型
    listed = {t["doc_type"] for t in list_templates()}
    assert set(NEW_TYPES) <= listed, f"list_templates 缺少新类型: {set(NEW_TYPES) - listed}"
