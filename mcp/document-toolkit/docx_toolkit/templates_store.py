"""templates_store.py —— 模板存取：内置预置模板 + 用户导入模板。"""

from __future__ import annotations

import json
import os
import time

# 内置模板目录（包内 templates/）
BUILTIN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
# 用户导入模板目录（server.py 同级 user_templates/，便于用户查看）
SERVER_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
USER_DIR = os.path.join(SERVER_DIR, "user_templates")

_DOC_TYPES = {"general", "thesis", "official", "contract", "bidding", "legal", "government_report", "techdoc", "resume", "notice"}


def _ensure_user_dir():
    os.makedirs(USER_DIR, exist_ok=True)


def get_builtin(doc_type: str) -> dict | None:
    """读取内置模板 JSON。"""
    if doc_type not in _DOC_TYPES:
        return None
    path = os.path.join(BUILTIN_DIR, f"{doc_type}.json")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _template_summary(t: dict) -> dict:
    """模板摘要：meta + 页面 + 样式角色数 + 骨架条数。"""
    styles = t.get("styles", {})
    return {
        "name": t.get("meta", {}).get("name", ""),
        "doc_type": t.get("meta", {}).get("doc_type", ""),
        "source": t.get("meta", {}).get("source", ""),
        "page": t.get("page", {}),
        "style_roles": list(styles.keys()) if isinstance(styles, dict) else f"{len(styles)} 条样式",
        "skeleton_count": len(t.get("skeleton", [])),
    }


def list_templates() -> list[dict]:
    """列出全部模板（内置 + 用户导入），含排版摘要。"""
    result = []
    for dt in sorted(_DOC_TYPES):
        t = get_builtin(dt)
        if t:
            result.append(_template_summary(t))
    _ensure_user_dir()
    for fn in sorted(os.listdir(USER_DIR)):
        if fn.endswith(".json"):
            try:
                with open(os.path.join(USER_DIR, fn), encoding="utf-8") as f:
                    t = json.load(f)
                result.append({**_template_summary(t), "file": fn})
            except (json.JSONDecodeError, OSError):
                continue
    return result


def save_user_template(template: dict, template_name: str) -> str:
    """保存用户导入的模板，返回文件路径。"""
    _ensure_user_dir()
    path = os.path.join(USER_DIR, f"{_safe_template_name(template_name)}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(template, f, ensure_ascii=False, indent=2)
    return path


def make_template_meta(name: str, doc_type: str, source: str, description: str = "") -> dict:
    """构造模板 meta 段。"""
    return {
        "name": name,
        "doc_type": doc_type,
        "source": source,
        "description": description,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }


_WIN_RESERVED = {"CON", "PRN", "AUX", "NUL"} | {f"COM{n}" for n in range(1, 10)} | {f"LPT{n}" for n in range(1, 10)}


def _safe_template_name(name: str) -> str:
    """模板名消毒：去非法字符，避开 Windows 保留名。"""
    safe = "".join(c for c in name if c.isalnum() or c in "-_") or "template"
    if safe.upper() in _WIN_RESERVED:
        safe = f"tpl_{safe}"
    return safe


def _user_path(name: str) -> str | None:
    """按模板名定位用户模板文件路径。"""
    _ensure_user_dir()
    path = os.path.join(USER_DIR, f"{_safe_template_name(name)}.json")
    return path if os.path.exists(path) else None


def rename_user_template(old_name: str, new_name: str) -> str | None:
    """重命名用户模板，返回新路径；模板不存在返回 None。"""
    src_path = _user_path(old_name)
    if src_path is None:
        return None
    dst_path = os.path.join(USER_DIR, f"{_safe_template_name(new_name)}.json")
    if src_path == dst_path:
        return dst_path
    if os.path.exists(dst_path):
        return "EXISTS"  # 目标已存在，拒绝覆盖
    with open(src_path, encoding="utf-8") as f:
        tpl = json.load(f)
    tpl["meta"]["name"] = new_name
    with open(dst_path, "w", encoding="utf-8") as f:
        json.dump(tpl, f, ensure_ascii=False, indent=2)
    os.remove(src_path)
    return dst_path


def delete_user_template(name: str) -> bool:
    """删除用户模板，返回是否删除成功。"""
    path = _user_path(name)
    if path is None:
        return False
    os.remove(path)
    return True


def export_template(name: str, output_path: str) -> str | None:
    """导出模板（用户模板或内置模板）为 JSON 文件，返回导出路径。"""
    tpl = None
    path = _user_path(name)
    if path is not None:
        with open(path, encoding="utf-8") as f:
            tpl = json.load(f)
    else:
        for dt in sorted(_DOC_TYPES):
            t = get_builtin(dt)
            if t and t.get("meta", {}).get("name") == name:
                tpl = t
                break
    if tpl is None:
        return None
    if os.path.exists(output_path):
        return "EXISTS"  # 拒绝覆盖已存在文件
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(tpl, f, ensure_ascii=False, indent=2)
    return output_path


def load_template_by_name(name: str) -> dict | None:
    """按名称加载模板（用户模板优先，其次内置），供对比/复用。"""
    path = _user_path(name)
    if path is not None:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    for dt in sorted(_DOC_TYPES):
        t = get_builtin(dt)
        if t and t.get("meta", {}).get("name") == name:
            return t
    return None


def compare_templates(name_a: str, name_b: str) -> dict:
    """对比两个模板：页面/样式/骨架差异。"""
    a = load_template_by_name(name_a)
    b = load_template_by_name(name_b)
    if a is None or b is None:
        missing = name_a if a is None else name_b
        return {"ok": False, "error": f"模板不存在: {missing}"}

    diff = {"page": {}, "styles": {}, "skeleton": {}}

    # 页面差异
    pa, pb = a.get("page", {}), b.get("page", {})
    for k in sorted(set(pa) | set(pb)):
        if pa.get(k) != pb.get(k):
            diff["page"][k] = {"a": pa.get(k), "b": pb.get(k)}

    # 样式差异（按 role）
    sa, sb = a.get("styles", {}), b.get("styles", {})
    if isinstance(sa, list):  # 内置模板是列表
        sa = {s["role"]: s for s in sa if "role" in s}
    if isinstance(sb, list):
        sb = {s["role"]: s for s in sb if "role" in s}
    for role in sorted(set(sa) | set(sb)):
        if sa.get(role) != sb.get(role):
            diff["styles"][role] = {"a": sa.get(role), "b": sb.get(role)}

    # 骨架差异（标题行数/条目）
    ka, kb = a.get("skeleton", []), b.get("skeleton", [])
    diff["skeleton"] = {
        "a_count": len(ka),
        "b_count": len(kb),
        "a_only": [x["text"] for x in ka if x not in kb][:20],
        "b_only": [x["text"] for x in kb if x not in ka][:20],
    }
    return {"ok": True, "name_a": name_a, "name_b": name_b, "diff": diff}
