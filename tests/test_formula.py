"""Excel 公式建议测试（mock LLM）。"""

import json
import os
import sys
from unittest.mock import patch

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.ai_generator import suggest_formula


def test_suggest_formula_json():
    mock = json.dumps({"formula": "=VLOOKUP(A1,B:C,2,0)", "explanation": "按产品名查价格", "alternatives": ["=XLOOKUP(A1,B:B,C:C)"]}, ensure_ascii=False)
    with patch("docx_toolkit.ai_generator._chat_completion", return_value=mock):
        r = suggest_formula("查找产品名对应的价格")
        assert r["formula"].startswith("=VLOOKUP")
        assert r["alternatives"]


def test_suggest_formula_plain():
    """LLM 返回纯公式文本时兜底。"""
    with patch("docx_toolkit.ai_generator._chat_completion", return_value="=SUM(A1:A10)"):
        r = suggest_formula("求A1到A10的和")
        assert r["formula"] == "=SUM(A1:A10)"
