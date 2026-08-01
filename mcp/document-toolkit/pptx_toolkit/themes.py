"""themes.py —— PPT 主题风格预设（色彩/字体体系）。"""

# 主题：主色/副色/文字色/背景/强调色（RGB hex）
THEMES = {
    # 商务：深蓝 + 浅蓝
    "corporate": {
        "name": "商务",
        "primary": "1F4E79", "secondary": "2E75B6", "accent": "F2C94C",
        "text": "333333", "background": "FFFFFF", "title_size": 40, "body_size": 18,
    },
    # 学术：深红 + 橙
    "academic": {
        "name": "学术",
        "primary": "8C1D18", "secondary": "C55A11", "accent": "D9B310",
        "text": "333333", "background": "FFFFFF", "title_size": 40, "body_size": 18,
    },
    # 发布会：黑底白字 + 黄强调
    "launch": {
        "name": "发布会",
        "primary": "111111", "secondary": "333333", "accent": "FFC000",
        "text": "FFFFFF", "background": "111111", "title_size": 44, "body_size": 20,
    },
    # 极简：灰白 + 细线感
    "minimal": {
        "name": "极简",
        "primary": "595959", "secondary": "8C8C8C", "accent": "2E75B6",
        "text": "404040", "background": "FFFFFF", "title_size": 40, "body_size": 18,
    },
}

DEFAULT_THEME = "corporate"

# 中文字体（同时设置 latin + eastAsia）
CN_FONT = "微软雅黑"
EN_FONT = "Calibri"
