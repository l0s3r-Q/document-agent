"""生成 10 类示例文档到 examples/output/。"""

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "docx-toolkit"))

from docx_toolkit.builder import build

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

SPECS = {
    "official": {
        "doc_type": "official",
        "title": "关于开展安全生产大检查的通知",
        "sections": [
            {"type": "heading1", "text": "一、检查范围"},
            {"type": "paragraph", "text": "本次检查覆盖全市所有在建工程项目、危化品生产经营单位。"},
            {"type": "heading1", "text": "二、时间安排"},
            {"type": "paragraph", "text": "自本通知印发之日起至2026年9月30日止。"},
            {"type": "paragraph", "text": "XX市人民政府办公室", "align": "RIGHT"},
            {"type": "paragraph", "text": "2026年8月2日", "align": "RIGHT"},
        ],
    },
    "thesis": {
        "doc_type": "thesis",
        "title": "基于深度学习的入侵检测研究",
        "sections": [
            {"type": "heading1", "text": "第一章 绪论"},
            {"type": "paragraph", "text": "网络入侵检测是网络安全防护的核心手段。"},
            {"type": "heading2", "text": "1.1 研究背景"},
            {"type": "paragraph", "text": "随着网络规模扩大，攻击手段日益复杂。"},
            {"type": "heading1", "text": "第二章 相关工作"},
            {"type": "paragraph", "text": "本章综述现有入侵检测方法。"},
        ],
    },
    "contract": {
        "doc_type": "contract",
        "title": "软件委托开发合同",
        "sections": [
            {"type": "paragraph", "text": "甲方：XX科技有限公司"},
            {"type": "paragraph", "text": "乙方：YY软件开发有限公司"},
            {"type": "heading1", "text": "第一条 合同标的"},
            {"type": "paragraph", "text": "乙方为甲方开发客户管理系统一套。"},
            {"type": "heading1", "text": "第二条 价款与支付"},
            {"type": "paragraph", "text": "合同总价款人民币壹拾万元整。"},
        ],
    },
    "bidding": {
        "doc_type": "bidding",
        "title": "办公设备采购项目招标公告",
        "sections": [
            {"type": "heading1", "text": "一、项目概况"},
            {"type": "paragraph", "text": "采购办公电脑 50 台、打印机 10 台。"},
            {"type": "heading1", "text": "二、投标人资格要求"},
            {"type": "paragraph", "text": "具有独立法人资格，近三年无重大违法记录。"},
        ],
    },
    "general": {
        "doc_type": "general",
        "title": "2026 年上半年工作总结",
        "sections": [
            {"type": "heading1", "text": "一、工作回顾"},
            {"type": "paragraph", "text": "上半年完成重点项目 12 项。"},
            {"type": "heading1", "text": "二、存在问题"},
            {"type": "paragraph", "text": "跨部门协作效率有待提升。"},
        ],
    },
    "legal": {
        "doc_type": "legal",
        "title": "民事起诉状",
        "sections": [
            {"type": "paragraph", "text": "原告：张三，男，1990年1月生，住北京市海淀区XX路X号。"},
            {"type": "paragraph", "text": "被告：李四，男，1988年5月生，住北京市朝阳区XX街X号。"},
            {"type": "heading1", "text": "诉讼请求"},
            {"type": "paragraph", "text": "一、判令被告偿还借款人民币 10 万元；二、判令被告承担本案诉讼费用。"},
            {"type": "heading1", "text": "事实与理由"},
            {"type": "paragraph", "text": "2025年3月，被告向原告借款 10 万元，约定一年内归还，至今未还。"},
            {"type": "paragraph", "text": "此致北京市海淀区人民法院", "align": "RIGHT"},
            {"type": "paragraph", "text": "具状人：张三    2026年8月2日", "align": "RIGHT"},
        ],
    },
    "government_report": {
        "doc_type": "government_report",
        "title": "2026年XX市政府工作报告",
        "sections": [
            {"type": "paragraph", "text": "各位代表："},
            {"type": "heading1", "text": "一、过去一年工作回顾"},
            {"type": "paragraph", "text": "全年地区生产总值增长 6.5%，民生支出占比持续提升。"},
            {"type": "heading1", "text": "二、总体要求和主要目标"},
            {"type": "paragraph", "text": "2026 年地区生产总值预期增长 6% 左右。"},
        ],
    },
    "techdoc": {
        "doc_type": "techdoc",
        "title": "客户管理系统需求说明书",
        "sections": [
            {"type": "heading1", "text": "1 引言"},
            {"type": "heading2", "text": "1.1 编写目的"},
            {"type": "paragraph", "text": "本文档定义客户管理系统的功能需求与验收标准。"},
            {"type": "heading1", "text": "2 功能需求"},
            {"type": "paragraph", "text": "FR-01 客户信息增删改查；FR-02 跟进记录管理；FR-03 数据看板。"},
        ],
    },
    "resume": {
        "doc_type": "resume",
        "title": "王小明",
        "sections": [
            {"type": "paragraph", "text": "求职意向：前端开发工程师", "align": "CENTER"},
            {"type": "heading1", "text": "教育背景"},
            {"type": "paragraph", "text": "2019-2023  XX大学  计算机科学与技术  本科"},
            {"type": "heading1", "text": "专业技能"},
            {"type": "paragraph", "text": "熟练掌握 HTML/CSS/JavaScript、Vue/React 框架。"},
        ],
    },
    "notice": {
        "doc_type": "notice",
        "title": "关于中秋节放假的通知",
        "sections": [
            {"type": "paragraph", "text": "全体员工："},
            {"type": "heading1", "text": "一、放假时间"},
            {"type": "paragraph", "text": "9月15日至9月17日放假调休，共 3 天。"},
            {"type": "heading1", "text": "二、注意事项"},
            {"type": "paragraph", "text": "请各部门做好值班安排，注意防火防盗。"},
            {"type": "paragraph", "text": "XX科技有限公司行政部    2026年8月2日", "align": "RIGHT"},
        ],
    },
}


def main():
    os.makedirs(OUT, exist_ok=True)
    for name, spec in SPECS.items():
        path = os.path.join(OUT, f"demo_{name}.docx")
        result = build(spec, path)
        assert result["ok"], result
        print(f"generated: {path}")


if __name__ == "__main__":
    main()
