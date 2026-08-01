"""模板一致性测试：styles.py 的 BUILTIN_STYLES 与 templates/*.json 必须等价（防止双数据源漂移）。"""

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.styles import BUILTIN_STYLES

TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit", "docx_toolkit", "templates")

KEY_FIELDS = ("font_name", "size_pt", "bold", "align", "line_spacing_rule", "line_spacing_pt", "line_spacing_multiple", "first_line_indent_chars")


def test_builtin_styles_match_templates():
    for doc_type, cfg in BUILTIN_STYLES.items():
        with open(os.path.join(TEMPLATES_DIR, f"{doc_type}.json"), encoding="utf-8") as f:
            tpl = json.load(f)
        tpl_by_role = {s["role"]: s for s in tpl["styles"]}
        assert set(tpl_by_role) == set(cfg["roles"]), f"{doc_type}: role 集合不一致"
        for role, style in cfg["roles"].items():
            for k in KEY_FIELDS:
                assert style.get(k) == tpl_by_role[role].get(k), f"{doc_type}.{role}.{k}: {style.get(k)} != {tpl_by_role[role].get(k)}"
        # 页面设置
        page = cfg["page"]
        assert tpl["page"] == page, f"{doc_type}: page 不一致"
    # 反向断言：JSON 存在但 styles.py 缺键 → 漂移
    for fn in os.listdir(TEMPLATES_DIR):
        if fn.endswith(".json"):
            dt = fn[:-5]
            assert dt in BUILTIN_STYLES, f"模板 {dt} 在 styles.py 中缺失"
