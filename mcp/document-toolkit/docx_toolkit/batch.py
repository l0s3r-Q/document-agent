"""batch.py —— 批量生成：spec 模板 + 多组数据 → 批量产出文档。"""

from __future__ import annotations

import copy
import json
import os

from .builder import build


def _render(value, data: dict, warnings: list[str]) -> str:
    """替换字符串中的 {var} 占位符。"""
    if not isinstance(value, str):
        return value
    import re
    def repl(m):
        key = m.group(1)
        if key in data:
            return str(data[key])
        warnings.append(f"未提供的变量: {{{key}}}（已替换为空）")
        return ""
    return re.sub(r"\{([^{}]+)\}", repl, value)


def _render_tree(node, data: dict, warnings: list[str]):
    """递归渲染 dict/list 中的所有字符串。"""
    if isinstance(node, dict):
        return {k: _render_tree(v, data, warnings) for k, v in node.items()}
    if isinstance(node, list):
        return [_render_tree(v, data, warnings) for v in node]
    if isinstance(node, str):
        return _render(node, data, warnings)
    return node


_WIN_RESERVED = {"CON", "PRN", "AUX", "NUL"} | {f"COM{n}" for n in range(1, 10)} | {f"LPT{n}" for n in range(1, 10)}


def _safe_filename(output_dir: str, fname: str, index: int) -> str:
    """文件名消毒：去非法字符/尾随空格点/保留名/超长，并处理重名。"""
    stem = "".join(c for c in fname if c.isprintable() and c not in '\/:*?"<>|').strip().rstrip(".")
    if stem.lower().endswith(".docx"):
        stem = stem[:-5]  # 去扩展名，统一在末尾拼接
    if not stem:
        stem = f"document_{index}"
    if stem.upper() in _WIN_RESERVED:
        stem = f"doc_{stem}"
    if len(stem) > 120:
        stem = stem[:120]
    path = os.path.join(output_dir, f"{stem}.docx")
    n = 2
    while os.path.exists(path):
        path = os.path.join(output_dir, f"{stem}_{n}.docx")
        n += 1
    return path


def batch_build(spec_template: dict, data_rows: list[dict], output_dir: str,
                filename_field: str | None = None) -> dict:
    """按 spec 模板批量生成文档。

    模板中字符串可用 {变量名} 占位；data_rows 每行提供一组变量。
    filename_field 指定用哪行数据的字段作为文件名（缺省用 index）。
    """
    if not isinstance(data_rows, list) or not data_rows:
        return {"ok": False, "error": "data_rows 必须是非空数组"}
    if not isinstance(spec_template, dict):
        return {"ok": False, "error": "spec_template 必须是对象"}
    if not output_dir:
        return {"ok": False, "error": "output_dir 不能为空"}
    # 预校验所有行（避免中途失败残留半成品）
    for i, row in enumerate(data_rows, 1):
        if not isinstance(row, dict):
            return {"ok": False, "error": f"第 {i} 行数据不是对象"}
    os.makedirs(output_dir, exist_ok=True)

    results, warnings = [], []
    for i, row in enumerate(data_rows, 1):
        spec = _render_tree(copy.deepcopy(spec_template), row, warnings)
        if filename_field and filename_field in row:
            fname = f"{str(row[filename_field])}.docx"
        else:
            fname = ((spec.get("title") or f"document_{i}")[:50]) + f"_{i}.docx"
        out = _safe_filename(output_dir, fname, i)
        try:
            r = build(spec, out)
            results.append({"index": i, "file": out, "ok": r["ok"]})
        except Exception as e:  # noqa: BLE001
            results.append({"index": i, "file": out, "ok": False, "error": f"{type(e).__name__}: {e}"})

    return {
        "ok": True,
        "total": len(data_rows),
        "succeeded": sum(1 for r in results if r["ok"]),
        "failed": sum(1 for r in results if not r["ok"]),
        "output_dir": output_dir,
        "warnings": warnings,
        "results": results,
    }
