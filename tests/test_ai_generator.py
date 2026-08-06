"""AI 生成模块测试：mock LLM 调用，验证生成/防 AI 味/解析。"""

import json
import os
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit import ai_generator


def _sample_spec() -> dict:
    return {
        "doc_type": "notice",
        "title": "关于开展安全生产检查的通知",
        "sections": [
            {"type": "paragraph", "text": "各部门、各车间：为落实安全生产主体责任，现决定开展专项检查。检查时间为8月10日至20日。请各部门对照检查清单逐项自查，发现隐患立即整改，并于8月21日前将检查结果报送安全办。"},
            {"type": "heading1", "text": "一、检查内容"},
            {"type": "list", "items": ["消防设施设备完好情况", "用电用气安全", "特种设备运行状态"]},
            {"type": "paragraph", "text": "各部门要高度重视，确保检查取得实效。"},
        ],
    }


class TestAiGenerator(unittest.TestCase):

    def test_extract_json_plain(self):
        obj = ai_generator._extract_json(json.dumps(_sample_spec(), ensure_ascii=False))
        self.assertEqual(obj["doc_type"], "notice")

    def test_extract_json_markdown(self):
        raw = "```json\n" + json.dumps(_sample_spec(), ensure_ascii=False) + "\n```"
        obj = ai_generator._extract_json(raw)
        self.assertEqual(obj["doc_type"], "notice")

    def test_extract_json_noisy(self):
        raw = "好的，以下是生成的文档：\n" + json.dumps(_sample_spec(), ensure_ascii=False) + "\n希望有帮助！"
        obj = ai_generator._extract_json(raw)
        self.assertEqual(obj["title"], "关于开展安全生产检查的通知")

    def test_extract_json_invalid(self):
        with self.assertRaises(ValueError):
            ai_generator._extract_json("这不是 JSON")

    def test_has_ai_flavor_detects(self):
        bad = _sample_spec()
        bad["sections"][0]["text"] = "总而言之，安全工作十分重要。"
        hits = ai_generator._has_ai_flavor(bad)
        self.assertTrue(any("总结套话" in h for h in hits))

    def test_has_ai_flavor_clean(self):
        hits = ai_generator._has_ai_flavor(_sample_spec())
        self.assertEqual(hits, [])

    def test_generate_spec_with_mock(self):
        """mock _chat_completion 返回干净 spec，验证 generate_spec 正常。"""
        with patch.object(ai_generator, "_chat_completion", return_value=json.dumps(_sample_spec(), ensure_ascii=False)):
            spec = ai_generator.generate_spec("notice", "安全生产检查")
            self.assertEqual(spec["doc_type"], "notice")

    def test_generate_spec_retry_on_ai_flavor(self):
        """第一次返回含 AI 腔，重试后返回干净 spec。"""
        bad = json.dumps({"doc_type": "notice", "title": "t", "sections": [{"type": "paragraph", "text": "综上所述，很重要。"}]}, ensure_ascii=False)
        good = json.dumps(_sample_spec(), ensure_ascii=False)
        with patch.object(ai_generator, "_chat_completion", side_effect=[bad, good]):
            spec = ai_generator.generate_spec("notice", "安全生产检查", retries=2)
            self.assertEqual(spec["doc_type"], "notice")

    def test_is_configured(self):
        with patch.dict(os.environ, {"DEEPSEEK_API_KEY": "test-key"}, clear=False):
            self.assertTrue(ai_generator.is_configured())


if __name__ == "__main__":
    unittest.main()
