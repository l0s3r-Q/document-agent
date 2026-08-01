"""document-toolkit MCP Server —— 文档解析/生成/模板导入工具箱。

工具（17）：
  - parse_docx(path)            解析 docx 页面/样式/结构
  - extract_structure(path)     仅提取标题结构树（不改结构场景）
  - build_docx(spec_json, out)  按 DocumentSpec 生成 docx
  - import_template(path, name) 导入范文/规范文档为模板
  - get_template(doc_type)      读取预置模板（10 类）
  - list_templates()            列出全部模板（含排版摘要）
  - suggest_restructure(src, t) 结构改造建议（keep/add/remove）
  - batch_build(tpl, rows, dir) 批量生成文档
  - rename_template(o, n)       重命名用户模板
  - delete_template(name)       删除用户模板
  - export_template(n, out)     导出模板 JSON
  - compare_templates(a, b)     对比两模板差异
  - build_excel(spec, out)      按 ExcelSpec 生成 xlsx（含美化）
  - parse_excel(path, sheet?)   解析 xlsx
  - excel_to_data(path, sheet?) Excel 数据表 → JSON 行数组（批量数据源）
  - convert_to_pdf(docx, out?)  docx → PDF（Word/WPS/LibreOffice 降级链）
  - pdf_info(path)              PDF 元信息（页数/大小）

启动：python server.py  （stdio 传输，FastMCP）
"""

from __future__ import annotations

import json
import os
import sys

# 确保可 import 包内模块（兼容从任意 cwd 启动）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import importlib
import threading
import time

# 模块级导入（热重载依赖：工具函数通过模块属性访问最新实现）
import docx_toolkit.batch
import docx_toolkit.builder
import docx_toolkit.parser
import docx_toolkit.restructure
import docx_toolkit.styles
import docx_toolkit.templates_store
import excel_toolkit.builder
import excel_toolkit.parser
import pdf_toolkit.converter

# ── 热重载机制：监听源码 mtime，变化时按依赖顺序 reload 模块 ──
_HOT_RELOAD_ORDER = [
    "docx_toolkit.styles",
    "docx_toolkit.templates_store",
    "docx_toolkit.parser",
    "docx_toolkit.builder",
    "docx_toolkit.batch",
    "docx_toolkit.restructure",
    "excel_toolkit.builder",
    "excel_toolkit.parser",
    "pdf_toolkit.converter",
]
_SRC_DIR = os.path.dirname(os.path.abspath(__file__))
_hot_last_check = 0.0
_hot_mtimes = {}


def _hot_reload():
    """工具调用前检查：源码文件有变化则 reload 相关模块（无需重启进程）。"""
    global _hot_last_check
    now = time.time()
    if now - _hot_last_check < 1.0:
        return
    _hot_last_check = now
    changed = False
    for root, _dirs, files in os.walk(_SRC_DIR):
        if "__pycache__" in root or "user_templates" in root:
            continue
        for f in files:
            if not f.endswith(".py"):
                continue
            p = os.path.join(root, f)
            try:
                mt = os.path.getmtime(p)
            except OSError:
                continue
            if p not in _hot_mtimes:
                _hot_mtimes[p] = mt
            elif mt != _hot_mtimes[p]:
                _hot_mtimes[p] = mt
                changed = True
    if changed:
        for name in _HOT_RELOAD_ORDER:
            mod = sys.modules.get(name)
            if mod is not None:
                try:
                    importlib.reload(mod)
                except Exception as e:  # noqa: BLE001
                    print(f"[hot-reload] {name} reload 失败: {e}", file=sys.stderr)

try:
    from mcp.server.fastmcp import FastMCP
except ImportError:
    # 提供友好错误提示
    raise SystemExit("缺少依赖 mcp：请先执行 pip install \"mcp>=1.26,<2\" python-docx")

# ── 兼容补丁 ──────────────────────────────────────────────────────────
# mcp 1.26.x 内部用 create_model(name, result=annotation) 包装返回值，
# 在 pydantic>=2.10 中该写法抛 PydanticUserError，需改为 (annotation, ...)。
# 补丁作用于本进程，不影响全局环境。
try:
    from mcp.server.fastmcp.utilities import func_metadata
    from pydantic import create_model as _create_model

    def _patched_wrapped_model(func_name: str, annotation):
        return _create_model(func_name, result=(annotation, ...))

    func_metadata._create_wrapped_model = _patched_wrapped_model
except Exception:  # noqa: BLE001  补丁失败不影响启动（旧 pydantic 无需补丁）
    pass

mcp = FastMCP("document-toolkit")


def _ok(data) -> str:
    return json.dumps({"ok": True, **data}, ensure_ascii=False)


def _err(message: str) -> str:
    return json.dumps({"ok": False, "error": message}, ensure_ascii=False)


def _err_sanitized(e: Exception) -> str:
    """错误脱敏：仅返回异常类型与简短信息，避免泄露本地绝对路径。"""
    return _err(f"{type(e).__name__}: {str(e)[:200]}")


@mcp.tool()
def parse_docx(path: str) -> str:
    """解析 .docx 文件，返回页面设置、样式体系与完整结构树（JSON 字符串）。"""
    _hot_reload()
    try:
        return _ok(docx_toolkit.parser.parse(path))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def extract_structure(path: str) -> str:
    """仅提取文档标题结构树（大纲），用于'不改结构'场景的结构锁定。"""
    _hot_reload()
    try:
        return _ok(docx_toolkit.parser.extract_structure(path))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def build_docx(spec_json: str, output_path: str) -> str:
    """按 DocumentSpec JSON 生成 .docx 文件。spec 含 doc_type/title/page/default_font/sections。"""
    _hot_reload()
    try:
        spec = json.loads(spec_json)
        if not isinstance(spec, dict):
            return _err("spec_json 必须是 JSON 对象")
        return _ok(docx_toolkit.builder.build(spec, output_path))
    except json.JSONDecodeError as e:
        return _err(f"spec_json 不是合法 JSON: {e}")
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def import_template(docx_path: str, template_name: str) -> str:
    """导入用户范文/规范说明文档：解析页面/样式/骨架，保存为用户模板，供后续生成使用。"""
    _hot_reload()
    try:
        data = docx_toolkit.parser.parse(docx_path)
        skeleton = [
            {"level": int(s["type"][-1]) if s["type"].startswith("heading") else 0, "text": s["text"]}
            for s in data["structure"]
            if s["type"] in ("heading1", "heading2", "heading3", "paragraph")
        ][:40]
        # 样式按 role 归类（heading1-3/body/table），使导入模板可直接用于 build_docx 的 styles 字段
        role_styles = {}
        for s in data["styles"]:
            if not s.get("font_name"):
                continue
            stype = s.get("style_name", "")
            if stype.startswith("Heading 1") or stype == "标题 1":
                role = "heading1"
            elif stype.startswith("Heading 2") or stype == "标题 2":
                role = "heading2"
            elif stype.startswith("Heading 3") or stype == "标题 3":
                role = "heading3"
            elif "Table" in stype or "表格" in stype:
                role = "table"
            else:
                role = "body"
            role_styles.setdefault(role, s)
        template = {
            "meta": docx_toolkit.templates_store.make_template_meta(template_name, "user", "user", f"从 {docx_path} 导入"),
            "page": data["page"],
            "styles": role_styles,
            "skeleton": skeleton,
        }
        path = docx_toolkit.templates_store.save_user_template(template, template_name)
        return _ok({"template": template, "path": path})
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def get_template(doc_type: str) -> str:
    """读取预置模板（doc_type ∈ general|thesis|official|contract|bidding|legal|government_report|techdoc|resume|notice），返回 JSON。"""
    _hot_reload()
    t = docx_toolkit.templates_store.get_builtin(doc_type)
    if t is None:
        return _err(f"未知类型 {doc_type}，可用: {', '.join(sorted(docx_toolkit.templates_store._DOC_TYPES))}")
    return _ok({"template": t})


@mcp.tool()
def list_templates() -> str:
    """列出全部可用模板（内置 10 类 + 用户导入），含排版摘要。"""
    return _ok({"templates": docx_toolkit.templates_store.list_templates()})


@mcp.tool()
def suggest_restructure(source_path: str, target_doc_type: str) -> str:
    """结构决策增强：对比源文档与目标类型模板，生成改造建议（保留/新增/移除清单）。"""
    _hot_reload()
    try:
        return _ok(docx_toolkit.restructure.suggest_restructure(source_path, target_doc_type))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def batch_build(spec_template_json: str, data_rows_json: str, output_dir: str,
                filename_field: str = "") -> str:
    """批量生成：spec 模板（字符串可含 {变量} 占位）+ 多组数据 → 批量产出到 output_dir。"""
    _hot_reload()
    try:
        spec = json.loads(spec_template_json)
        rows = json.loads(data_rows_json)
        return _ok(docx_toolkit.batch.batch_build(spec, rows, output_dir, filename_field or None))
    except json.JSONDecodeError as e:
        return _err(f"JSON 解析失败: {e}")
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def rename_template(old_name: str, new_name: str) -> str:
    """重命名用户导入的模板。"""
    _hot_reload()
    try:
        path = docx_toolkit.templates_store.rename_user_template(old_name, new_name)
        if path == "EXISTS":
            return _err(f"目标模板名已存在: {new_name}")
        if path is None:
            return _err(f"模板不存在: {old_name}")
        return _ok({"old_name": old_name, "new_name": new_name, "path": path})
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def delete_template(name: str) -> str:
    """删除用户导入的模板（内置模板不可删）。"""
    _hot_reload()
    try:
        if docx_toolkit.templates_store.delete_user_template(name):
            return _ok({"deleted": name})
        return _err(f"模板不存在或为内置模板: {name}")
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def export_template(name: str, output_path: str) -> str:
    """导出模板（用户或内置）为 JSON 文件。"""
    _hot_reload()
    try:
        path = docx_toolkit.templates_store.export_template(name, output_path)
        if path == "EXISTS":
            return _err(f"目标文件已存在: {output_path}")
        if path is None:
            return _err(f"模板不存在: {name}")
        return _ok({"name": name, "path": path})
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def compare_templates(name_a: str, name_b: str) -> str:
    """对比两个模板的页面/样式/骨架差异。"""
    return _ok(docx_toolkit.templates_store.compare_templates(name_a, name_b))




# ══════════════════════════ Excel 工具 ══════════════════════════

@mcp.tool()
def build_excel(spec_json: str, output_path: str) -> str:
    """按 ExcelSpec JSON 生成 .xlsx 报表：sheets 含 rows/styles/merges/col_widths/freeze/filter；默认表头美化（加粗+浅蓝填充+边框+自适应列宽）。"""
    _hot_reload()
    try:
        spec = json.loads(spec_json)
        if not isinstance(spec, dict):
            return _err("spec_json 必须是 JSON 对象")
        return _ok(excel_toolkit.builder.build(spec, output_path))
    except json.JSONDecodeError as e:
        return _err(f"spec_json 不是合法 JSON: {e}")
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def parse_excel(path: str, sheet_name: str = "") -> str:
    """解析 .xlsx：返回各 sheet 的行数据/合并单元格/列宽。sheet_name 为空解析全部。"""
    _hot_reload()
    try:
        return _ok(excel_toolkit.parser.parse(path, sheet_name or None))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def excel_to_data(path: str, sheet_name: str = "") -> str:
    """把 Excel 数据表转为 JSON 行数组（首行为字段名），可直接作为 batch_build 的 data_rows 数据源。"""
    _hot_reload()
    try:
        return _ok(excel_toolkit.parser.to_data(path, sheet_name or None))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


# ══════════════════════════ PDF 工具 ══════════════════════════

@mcp.tool()
def convert_to_pdf(docx_path: str, output_path: str = "") -> str:
    """docx → PDF 转换。自动引擎降级：Word COM → WPS COM → LibreOffice headless。"""
    _hot_reload()
    try:
        return _ok(pdf_toolkit.converter.convert(docx_path, output_path or None))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def pdf_info(path: str) -> str:
    """读取 PDF 元信息（页数/文件大小）。"""
    _hot_reload()
    try:
        if not os.path.exists(path):
            return _err(f"文件不存在: {path}")
        try:
            from pypdf import PdfReader
        except ImportError:
            return _err("缺少依赖 pypdf：请执行 pip install pypdf")
        with open(path, "rb") as f:
            reader = PdfReader(f)
            pages = len(reader.pages)
        return _ok({"path": path, "pages": pages,
                    "size_bytes": os.path.getsize(path)})
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


if __name__ == "__main__":
    mcp.run()