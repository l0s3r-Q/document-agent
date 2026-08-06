"""多模板串联批量生成测试：batch_build_multi。"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "mcp", "document-toolkit"))

from docx_toolkit.batch import batch_build_multi

TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out_multi")
import shutil
shutil.rmtree(TMP, ignore_errors=True)
os.makedirs(TMP, exist_ok=True)


def test_batch_multi_basic():
    """2 模板 × 2 数据行 = 4 份文档。"""
    templates = [
        {"doc_type": "notice", "title": "{name}欢迎信", "sections": [{"type": "paragraph", "text": "{name}您好，欢迎加入{company}。"}]},
        {"doc_type": "general", "title": "{name}入职名单", "sections": [{"type": "paragraph", "text": "{name}，{dept}部门，工号{id}。"}]},
    ]
    rows = [
        {"name": "张三", "company": "某某公司", "department": "技术", "id": "001"},
        {"name": "李四", "company": "某某公司", "department": "市场", "id": "002"},
    ]
    r = batch_build_multi(templates, rows, TMP)
    assert r["ok"], r
    assert r["total"] == 4 and r["succeeded"] == 4
    files = [x["file"] for x in r["results"]]
    assert len(files) == 4
    assert all(os.path.exists(f) for f in files)


def test_batch_multi_filename_fields():
    """指定 filename_fields。"""
    templates = [{"doc_type": "general", "title": "{name}文档", "sections": []}]
    rows = [{"name": "张三"}, {"name": "李四"}]
    r = batch_build_multi(templates, rows, TMP, filename_fields=["name"])
    assert r["succeeded"] == 2
    assert any("张三" in os.path.basename(x["file"]) for x in r["results"])


def test_batch_multi_empty():
    r = batch_build_multi([], [], TMP)
    assert not r.get("ok")
