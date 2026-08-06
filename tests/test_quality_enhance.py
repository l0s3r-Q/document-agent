"""质量检查增强测试：重复段落 / Excel 空行与重复行。"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.builder import build as build_docx
from docx_toolkit.quality import check_docx, check_excel

TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out_quality")
os.makedirs(TMP, exist_ok=True)


def test_dup_paragraph_detected():
    """重复段落应被检测为 dup_paragraph warning。"""
    spec = {
        "doc_type": "general",
        "title": "重复测试",
        "sections": [
            {"type": "paragraph", "text": "安全生产检查是保障企业稳定运行的基础性工作，需要各部门高度重视并认真落实。"},
            {"type": "paragraph", "text": "安全生产检查是保障企业稳定运行的基础性工作，需要各部门高度重视并认真落实。"},
            {"type": "paragraph", "text": "各部门应当对照检查清单逐项排查隐患。"},
        ],
    }
    out = os.path.join(TMP, "dup.docx")
    build_docx(spec, out)
    res = check_docx(out)
    dup = [i for i in res["issues"] if i["type"] == "dup_paragraph"]
    assert dup, f"未检测到重复段落: {res['issues']}"


def test_no_dup_paragraph_clean():
    """不同段落不应误报。"""
    spec = {
        "doc_type": "general",
        "title": "正常文档",
        "sections": [
            {"type": "paragraph", "text": "第一段：介绍项目背景与目标，明确本次工作的范围与意义。"},
            {"type": "paragraph", "text": "第二段：说明实施步骤，包括需求分析、方案设计与验收等环节。"},
            {"type": "paragraph", "text": "第三段：强调时间节点与责任人，确保各项工作按时保质完成。"},
        ],
    }
    out = os.path.join(TMP, "clean.docx")
    build_docx(spec, out)
    res = check_docx(out)
    dup = [i for i in res["issues"] if i["type"] == "dup_paragraph"]
    assert not dup, f"误报重复段落: {dup}"


def test_excel_dup_rows_detected():
    """Excel 完全重复行应被检测。"""
    from openpyxl import Workbook
    out = os.path.join(TMP, "dup.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.append(["部门", "人数", "负责人"])
    for _ in range(4):
        ws.append(["生产部", "50", "张三"])  # 4 行完全相同
    wb.save(out)
    res = check_excel(out)
    dup = [i for i in res["issues"] if i["type"] == "dup_rows"]
    assert dup, f"未检测到重复行: {res['issues']}"


def test_excel_many_empty_rows_detected():
    """Excel 空行占比过高应被检测（用部分空单元格模拟真实空行）。"""
    from openpyxl import Workbook
    out = os.path.join(TMP, "empty.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.append(["部门", "人数"])
    ws.append(["生产部", "50"])
    # 模拟空行：只有前几列有空格占位（openpyxl 全 None 行不落盘）
    for _ in range(6):
        ws.append([" ", ""])
    wb.save(out)
    res = check_excel(out)
    empty = [i for i in res["issues"] if i["type"] == "many_empty_rows"]
    assert empty, f"未检测到空行占比问题: {res['issues']}"
