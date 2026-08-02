"""themes.py —— PPT 主题风格预设（色彩/字体体系）。"""

# 主题：主色/副色/文字色/标题色/背景/强调色（RGB hex）
# 约定：标题文字默认黑色（title 字段），主题色仅用于装饰（标题条背景/表格表头/强调线）
THEMES = {
    # 商务：深蓝装饰 + 黑字标题
    "corporate": {
        "name": "商务",
        "primary": "1F4E79", "secondary": "2E75B6", "accent": "F2C94C",
        "text": "333333", "title": "000000", "background": "FFFFFF",
        "title_size": 40, "body_size": 18,
    },
    # 学术：深红装饰 + 黑字标题
    "academic": {
        "name": "学术",
        "primary": "8C1D18", "secondary": "C55A11", "accent": "D9B310",
        "text": "333333", "title": "000000", "background": "FFFFFF",
        "title_size": 40, "body_size": 18,
    },
    # 发布会：黑底白字（标题白色，深色背景下可读）
    "launch": {
        "name": "发布会",
        "primary": "111111", "secondary": "333333", "accent": "FFC000",
        "text": "FFFFFF", "title": "FFFFFF", "background": "111111",
        "title_size": 44, "body_size": 20,
    },
    # 极简：灰白 + 黑字标题
    "minimal": {
        "name": "极简",
        "primary": "595959", "secondary": "8C8C8C", "accent": "2E75B6",
        "text": "404040", "title": "000000", "background": "FFFFFF",
        "title_size": 40, "body_size": 18,
    },
}

DEFAULT_THEME = "corporate"

# 中文字体（同时设置 latin + eastAsia）
CN_FONT = "微软雅黑"
EN_FONT = "Calibri"
