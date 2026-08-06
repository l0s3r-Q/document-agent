"""AI 改写/润色测试（mock LLM）。"""

import os
import sys
from unittest.mock import patch

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.ai_generator import rewrite_text


def test_rewrite_polish():
    with patch("docx_toolkit.ai_generator._chat_completion", return_value="优化后的文本"):
        r = rewrite_text("原文本内容", "polish")
        assert r == "优化后的文本"


def test_rewrite_summary():
    with patch("docx_toolkit.ai_generator._chat_completion", return_value="核心要点摘要"):
        r = rewrite_text("很长的原文" * 10, "summary")
        assert r == "核心要点摘要"


def test_rewrite_modes():
    with patch("docx_toolkit.ai_generator._chat_completion", return_value="x"):
        for mode in ["polish", "rewrite", "summary", "expand"]:
            rewrite_text("测试", mode)
