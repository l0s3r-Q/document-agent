"""styles.py —— 预置 10 类文档排版规范（doc_type → 页面/字体/段落参数）。"""

# 排版参数说明：
#   line_spacing_rule: "EXACTLY" = 固定值行距（单位磅），"MULTIPLE" = 倍数行距
#   line_spacing_pt / line_spacing_multiple: 对应上述两种规则的值
#   first_line_indent: 首行缩进（字符数，按字号计算）

BUILTIN_STYLES = {
    # ── 党政公文 GB/T 9704-2012 ────────────────────────────────────────────
    "official": {
        "name": "党政公文（GB/T 9704-2012）",
        "page": {"top_cm": 3.7, "bottom_cm": 3.5, "left_cm": 2.8, "right_cm": 2.6},
        "roles": {
            "title":     {"font_name": "方正小标宋简体", "size_pt": 22, "bold": False, "align": "CENTER", "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "heading1":  {"font_name": "黑体",           "size_pt": 16, "bold": False, "align": "LEFT",   "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "heading2":  {"font_name": "楷体_GB2312",    "size_pt": 16, "bold": False, "align": "LEFT",   "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "heading3":  {"font_name": "仿宋_GB2312",    "size_pt": 16, "bold": True,  "align": "LEFT",   "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "body":      {"font_name": "仿宋_GB2312",    "size_pt": 16, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28, "first_line_indent_chars": 2},
            "table":     {"font_name": "仿宋_GB2312",    "size_pt": 14, "bold": False, "align": "CENTER",  "line_spacing_rule": "EXACTLY", "line_spacing_pt": 22},
        },
    },
    # ── 学位论文（通用高校规范）────────────────────────────────────────────
    "thesis": {
        "name": "学位论文（通用规范）",
        "page": {"top_cm": 3.0, "bottom_cm": 2.5, "left_cm": 3.0, "right_cm": 2.5},
        "roles": {
            "title":     {"font_name": "黑体", "size_pt": 16, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "黑体", "size_pt": 16, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "黑体", "size_pt": 14, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "黑体", "size_pt": 12, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2},
            "table":     {"font_name": "宋体", "size_pt": 10.5, "bold": False, "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    # ── 合同协议 ───────────────────────────────────────────────────────────
    "contract": {
        "name": "合同协议",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
        "roles": {
            "title":     {"font_name": "宋体", "size_pt": 22, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "黑体", "size_pt": 14, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "宋体", "size_pt": 14, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "宋体", "size_pt": 12, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "宋体", "size_pt": 14, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2},
            "table":     {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "CENTER",  "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    # ── 招投标文书 ─────────────────────────────────────────────────────────
    "bidding": {
        "name": "招投标文书",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
        "roles": {
            "title":     {"font_name": "黑体", "size_pt": 22, "bold": True,  "align": "CENTER", "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "heading1":  {"font_name": "黑体", "size_pt": 16, "bold": False, "align": "LEFT",   "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "heading2":  {"font_name": "楷体_GB2312", "size_pt": 16, "bold": False, "align": "LEFT", "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "heading3":  {"font_name": "仿宋_GB2312", "size_pt": 16, "bold": True, "align": "LEFT", "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "body":      {"font_name": "仿宋_GB2312", "size_pt": 16, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28, "first_line_indent_chars": 2},
            "table":     {"font_name": "仿宋_GB2312", "size_pt": 14, "bold": False, "align": "CENTER", "line_spacing_rule": "EXACTLY", "line_spacing_pt": 22},
        },
    },
    # ── 通用文档 ───────────────────────────────────────────────────────────
    "general": {
        "name": "通用文档",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
        "roles": {
            "title":     {"font_name": "黑体", "size_pt": 16, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "黑体", "size_pt": 16, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "黑体", "size_pt": 14, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "黑体", "size_pt": 12, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2},
            "table":     {"font_name": "宋体", "size_pt": 10.5, "bold": False, "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    # ── 法律文书 ───────────────────────────────────────────────────────────
    "legal": {
        "name": "法律文书",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
        "roles": {
            "title":     {"font_name": "宋体", "size_pt": 22, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "黑体", "size_pt": 14, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "宋体", "size_pt": 14, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "宋体", "size_pt": 12, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "仿宋_GB2312", "size_pt": 14, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2},
            "table":     {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "CENTER",  "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    # ── 政府工作报告 ───────────────────────────────────────────────────────
    "government_report": {
        "name": "政府工作报告",
        "page": {"top_cm": 3.7, "bottom_cm": 3.5, "left_cm": 2.8, "right_cm": 2.6},
        "roles": {
            "title":     {"font_name": "方正小标宋简体", "size_pt": 22, "bold": False, "align": "CENTER", "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "heading1":  {"font_name": "黑体",           "size_pt": 16, "bold": False, "align": "LEFT",   "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "heading2":  {"font_name": "楷体_GB2312",    "size_pt": 16, "bold": False, "align": "LEFT",   "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "heading3":  {"font_name": "仿宋_GB2312",    "size_pt": 16, "bold": True,  "align": "LEFT",   "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28},
            "body":      {"font_name": "仿宋_GB2312",    "size_pt": 16, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "EXACTLY", "line_spacing_pt": 28, "first_line_indent_chars": 2},
            "table":     {"font_name": "仿宋_GB2312",    "size_pt": 14, "bold": False, "align": "CENTER",  "line_spacing_rule": "EXACTLY", "line_spacing_pt": 22},
        },
    },
    # ── 技术文档 ───────────────────────────────────────────────────────────
    "techdoc": {
        "name": "技术文档",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 2.5, "right_cm": 2.5},
        "roles": {
            "title":     {"font_name": "微软雅黑", "size_pt": 18, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "微软雅黑", "size_pt": 16, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "微软雅黑", "size_pt": 14, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "微软雅黑", "size_pt": 12, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "table":     {"font_name": "宋体", "size_pt": 10.5, "bold": False, "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    # ── 简历 ───────────────────────────────────────────────────────────────
    "resume": {
        "name": "简历",
        "page": {"top_cm": 2.0, "bottom_cm": 2.0, "left_cm": 2.5, "right_cm": 2.5},
        "roles": {
            "title":     {"font_name": "微软雅黑", "size_pt": 24, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.2},
            "heading1":  {"font_name": "微软雅黑", "size_pt": 14, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.2},
            "heading2":  {"font_name": "微软雅黑", "size_pt": 12, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.2},
            "heading3":  {"font_name": "微软雅黑", "size_pt": 11, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.2},
            "body":      {"font_name": "宋体", "size_pt": 11, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.2},
            "table":     {"font_name": "宋体", "size_pt": 10.5, "bold": False, "align": "LEFT", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    # ── 通知公告 ───────────────────────────────────────────────────────────
    "notice": {
        "name": "通知公告",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
        "roles": {
            "title":     {"font_name": "黑体", "size_pt": 18, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "黑体", "size_pt": 14, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "黑体", "size_pt": 13, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "宋体", "size_pt": 12, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2},
            "table":     {"font_name": "宋体", "size_pt": 10.5, "bold": False, "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    "meeting_minutes": {
        "name": "会议纪要",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
        "roles": {
            "title":     {"font_name": "黑体", "size_pt": 18, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "黑体", "size_pt": 14, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "黑体", "size_pt": 13, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "宋体", "size_pt": 12, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2},
            "table":     {"font_name": "宋体", "size_pt": 10.5, "bold": False, "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    "speech": {
        "name": "演讲稿",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
        "roles": {
            "title":     {"font_name": "黑体", "size_pt": 18, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "黑体", "size_pt": 14, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "黑体", "size_pt": 13, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "宋体", "size_pt": 12, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2},
            "table":     {"font_name": "宋体", "size_pt": 10.5, "bold": False, "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    "proposal": {
        "name": "方案建议书",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
        "roles": {
            "title":     {"font_name": "黑体", "size_pt": 18, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "黑体", "size_pt": 14, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "黑体", "size_pt": 13, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "宋体", "size_pt": 12, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2},
            "table":     {"font_name": "宋体", "size_pt": 10.5, "bold": False, "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
    "invitation": {
        "name": "邀请函",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
        "roles": {
            "title":     {"font_name": "宋体", "size_pt": 18, "bold": True,  "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading1":  {"font_name": "黑体", "size_pt": 14, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading2":  {"font_name": "黑体", "size_pt": 13, "bold": False, "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "heading3":  {"font_name": "宋体", "size_pt": 12, "bold": True,  "align": "LEFT",   "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5},
            "body":      {"font_name": "宋体", "size_pt": 12, "bold": False, "align": "JUSTIFY", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2},
            "table":     {"font_name": "宋体", "size_pt": 10.5, "bold": False, "align": "CENTER", "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.0},
        },
    },
}