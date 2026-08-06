"""AI 三格式生成测试：generate_excel_spec / generate_pptx_spec（mock LLM）。"""

import json
import os
import sys
from unittest.mock import patch

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.ai_generator import generate_excel_spec, generate_pptx_spec
from excel_toolkit.builder import build as build_excel
from pptx_toolkit.builder import build as build_pptx

TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out_formats")
os.makedirs(TMP, exist_ok=True)


def test_generate_excel_spec():
    """mock LLM 返回 Excel spec → 构建 xlsx。"""
    mock_spec = {
        "sheets": [
            {"name": "数据明细", "rows": [["部门", "人数"], ["生产部", "50"], ["质检部", "30"]]},
            {"name": "汇总", "rows": [["指标", "数值"], ["总人数", "80"]]},
        ]
    }
    with patch("docx_toolkit.ai_generator._chat_completion", return_value=json.dumps(mock_spec, ensure_ascii=False)):
        spec = generate_excel_spec("部门统计")
        assert "sheets" in spec and len(spec["sheets"]) == 2
        out = os.path.join(TMP, "test.xlsx")
        r = build_excel(spec, out)
        assert r.get("ok") and os.path.exists(out)


def test_generate_pptx_spec():
    """mock LLM 返回 PPT spec → 构建 pptx。"""
    mock_spec = {
        "title": "年度总结",
        "theme": "corporate",
        "slides": [
            {"type": "cover", "title": "年度总结汇报"},
            {"type": "content", "title": "业务回顾", "bullets": ["营收增长", "> 同比+15%"]},
            {"type": "chart", "title": "数据", "chart_type": "column", "categories": ["Q1", "Q2"], "series": [{"name": "营收", "values": [100, 120]}]},
            {"type": "closing", "title": "谢谢"},
        ],
    }
    with patch("docx_toolkit.ai_generator._chat_completion", return_value=json.dumps(mock_spec, ensure_ascii=False)):
        spec = generate_pptx_spec("年度总结")
        assert "slides" in spec and len(spec["slides"]) == 4
        out = os.path.join(TMP, "test.pptx")
        r = build_pptx(spec, out)
        assert r.get("ok") and os.path.exists(out)


def test_generate_spec_ai_flavor_retry():
    """Excel spec 含 AI 腔时重试。"""
    bad = {"sheets": [{"rows": [["a"], ["总而言之重要"]]}]}
    good = {"sheets": [{"rows": [["a"], ["具体数据"]]}]}
    with patch("docx_toolkit.ai_generator._chat_completion", side_effect=[json.dumps(bad, ensure_ascii=False), json.dumps(good, ensure_ascii=False)]):
        spec = generate_excel_spec("t")
        assert spec["sheets"][0]["rows"][1][0] == "具体数据"
