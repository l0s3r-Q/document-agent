"""document-toolkit MCP Server —— 文档解析/生成/模板导入工具箱。

工具（24）：
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
  - build_pptx(spec, out)       按 PptxSpec 生成 pptx（8 版式/4 主题）
  - parse_pptx(path)            解析 pptx（页/形状/表格）
  - docx_tables_to_excel(d,o)  docx 表格 → xlsx
  - excel_to_docx(e,o)         xlsx → docx 表格文档
  - docx_to_markdown(d,o)      docx → Markdown
  - merge_pdfs(paths, out)     合并多个 PDF
  - quality_check(path)        交付质量体检（AIGC 痕迹/占位符/排版）

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
import docx_toolkit.quality
import docx_toolkit.restructure
import docx_toolkit.styles
import docx_toolkit.templates_store
import docx_toolkit.ai_generator
import excel_toolkit.builder
import excel_toolkit.parser
import pdf_toolkit.converter
import pptx_toolkit.builder
import pptx_toolkit.parser

# ── 热重载机制：监听源码 mtime，变化时按依赖顺序 reload 模块 ──
_HOT_RELOAD_ORDER = [
    "docx_toolkit.styles",
    "docx_toolkit.templates_store",
    "docx_toolkit.parser",
    "docx_toolkit.builder",
    "docx_toolkit.batch",
    "docx_toolkit.markdown",
    "docx_toolkit.restructure",
    "docx_toolkit.quality",
    "docx_toolkit.ai_generator",
    "excel_toolkit.builder",
    "excel_toolkit.parser",
    "pdf_toolkit.converter",
    "pptx_toolkit.themes",
    "pptx_toolkit.builder",
    "pptx_toolkit.parser",
]
_SRC_DIR = os.path.dirname(os.path.abspath(__file__))
_hot_last_check = 0.0
_hot_mtimes = {}
_hot_lock = threading.Lock()


def _hot_reload():
    """工具调用前检查：源码文件有变化则 reload 相关模块（无需重启进程）。

    并发安全：全程加锁；reload 前 compile 预检语法，失败则跳过并记录。
    """
    global _hot_last_check
    now = time.time()
    if now - _hot_last_check < 1.0:
        return
    with _hot_lock:
        now2 = time.time()
        if now2 - _hot_last_check < 1.0:
            return
        _hot_last_check = now2
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
                if mod is None:
                    continue
                try:
                    compile(open(mod.__file__, encoding="utf-8").read(), mod.__file__, "exec")
                    importlib.reload(mod)
                except Exception as e:  # noqa: BLE001
                    print(f"[hot-reload] {name} reload 失败（保留旧模块）: {e}", file=sys.stderr)

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
    """错误脱敏：异常类型 + 去除本地绝对路径的简短信息。"""
    import re as _re
    msg = str(e)[:300]
    msg = _re.sub(r'[A-Za-z]:[\\/][^ \t\r\n]+', '<path>', msg)
    msg = _re.sub(r'[\\/](?:home|Users|users)[\\/][^ \t\r\n]+', '<path>', msg)
    return _err(f"{type(e).__name__}: {msg}")


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
            "meta": docx_toolkit.templates_store.make_template_meta(template_name, "user", "user",
                                              f"从 {os.path.basename(docx_path)} 导入"),
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
    """读取预置模板（doc_type ∈ general|thesis|official|contract|bidding|legal|government_report|techdoc|resume|notice|meeting_minutes|speech|proposal|invitation），返回 JSON。"""
    _hot_reload()
    t = docx_toolkit.templates_store.get_builtin(doc_type)
    if t is None:
        return _err(f"未知类型 {doc_type}，可用: general, thesis, official, contract, bidding, legal, government_report, techdoc, resume, notice")
    return _ok({"template": t})


@mcp.tool()
def list_templates() -> str:
    """列出全部可用模板（内置 10 类 + 用户导入），含排版摘要。"""
    _hot_reload()
    try:
        return _ok({"templates": docx_toolkit.templates_store.list_templates()})
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


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
    _hot_reload()
    try:
        return _ok(docx_toolkit.templates_store.compare_templates(name_a, name_b))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)




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


# ══════════════════════════ PPT 工具 ══════════════════════════

@mcp.tool()
def build_pptx(spec_json: str, output_path: str) -> str:
    """按 PptxSpec JSON 生成 .pptx：slides 支持 cover/agenda/section/content/two_column/table/image/closing 8 版式，theme 支持 corporate/academic/launch/minimal。"""
    _hot_reload()
    try:
        spec = json.loads(spec_json)
        if not isinstance(spec, dict):
            return _err("spec_json 必须是 JSON 对象")
        return _ok(pptx_toolkit.builder.build(spec, output_path))
    except json.JSONDecodeError as e:
        return _err(f"spec_json 不是合法 JSON: {e}")
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def parse_pptx(path: str) -> str:
    """解析 .pptx：返回各页形状/文本/表格结构。"""
    _hot_reload()
    try:
        return _ok(pptx_toolkit.parser.parse(path))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


# ══════════════════════════ 互转工具 ══════════════════════════

@mcp.tool()
def docx_tables_to_excel(docx_path: str, output_path: str) -> str:
    """提取 docx 中全部表格 → 生成 xlsx（每表一个 sheet，自动命名）。"""
    _hot_reload()
    try:
        data = docx_toolkit.parser.parse(docx_path)
        tables = [s for s in data["structure"] if s["type"] == "table"]
        if not tables:
            return _err(f"文档中未找到表格: {docx_path}")
        sheets = []
        for i, t in enumerate(tables, 1):
            rows = t.get("rows", [])
            if not rows:
                continue
            sheets.append({"name": f"表{i}", "rows": rows, "fill": "blue"})
        if not sheets:
            return _err("表格为空")
        return _ok(excel_toolkit.builder.build({"sheets": sheets}, output_path))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def excel_to_docx(excel_path: str, output_path: str, with_sheet_titles: bool = True) -> str:
    """xlsx → docx：每 sheet 生成一个表格（可选带 sheet 名标题）。"""
    _hot_reload()
    try:
        data = excel_toolkit.parser.parse(excel_path)
        sheets = data.get("sheets", [])
        if not sheets:
            return _err("未找到 sheet")
        sections = []
        for sh in sheets:
            if with_sheet_titles and sh.get("name"):
                sections.append({"type": "heading2", "text": sh["name"]})
            rows = sh.get("rows", [])
            if rows:
                sections.append({"type": "table", "rows": rows})
        spec = {"doc_type": "general", "title": os.path.splitext(os.path.basename(excel_path))[0],
                "sections": sections}
        return _ok(docx_toolkit.builder.build(spec, output_path))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def docx_to_markdown(docx_path: str, output_path: str = "") -> str:
    """docx → Markdown：标题/正文/表格/列表完整转换。"""
    _hot_reload()
    try:
        return _ok(docx_toolkit.markdown.to_markdown(docx_path, output_path or None))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def quality_check(path: str) -> str:
    """交付质量体检：AIGC 痕迹词/占位符残留/emoji/表格参差/标题跳级/空页（docx/pptx/xlsx）。"""
    _hot_reload()
    try:
        return _ok(docx_toolkit.quality.quality_check(path))
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def merge_pdfs(pdf_paths_json: str, output_path: str) -> str:
    """合并多个 PDF 为一个文件（pypdf）；缺失文件自动跳过并警告。"""
    _hot_reload()
    try:
        paths = json.loads(pdf_paths_json)
        if not isinstance(paths, list):
            return _err("pdf_paths_json 必须是数组")
        return _ok(pdf_toolkit.converter.merge_pdfs(paths, output_path))
    except json.JSONDecodeError as e:
        return _err(f"JSON 解析失败: {e}")
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


@mcp.tool()
def generate_docx(doc_type: str, topic: str, output_path: str = "", extra: str = "") -> str:
    """AI 生成文档：输入类型+主题 → LLM 生成内容 → 构建 docx（防 AI 味自检）。

    参数：
      doc_type:    文档类型（general/thesis/official/contract/bidding/legal/government_report/techdoc/resume/notice/meeting_minutes/speech/proposal/invitation）
      topic:       主题（如"关于开展安全生产检查的通知"）
      output_path: 输出路径（留空则自动生成）
      extra:       额外要求/要点（可空，如"重点强调消防与用电安全，附检查表"）

    说明：
      - 依赖 AI_GEN_PROVIDER（默认 deepseek）对应的 API key 环境变量
      - 生成内容经过防 AI 味自检（AIGC 痕迹/占位符），不合格自动重生成
      - 未配置 API key 时返回明确提示
    """
    _hot_reload()
    try:
        if not docx_toolkit.ai_generator.is_configured():
            return _err("未配置 AI 生成 API key。请设置环境变量（默认 DEEPSEEK_API_KEY，或 AI_GEN_PROVIDER=mimo 用 MIMO_API_KEY）")
        spec = docx_toolkit.ai_generator.generate_spec(doc_type, topic, extra)
        if not output_path:
            import re as _re
            safe = _re.sub(r"[\\/:*?\"<>|]", "_", topic)[:40]
            output_path = os.path.join(os.path.expanduser("~"), "Desktop", "文档", f"{safe}.docx")
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
        result = docx_toolkit.builder.build(spec, output_path)
        # 附质量体检摘要
        try:
            qc = docx_toolkit.quality.quality_check(output_path)
            return _ok({"path": output_path, "spec": spec, "quality": qc})
        except Exception:
            return _ok({"path": output_path, "spec": spec})
    except Exception as e:  # noqa: BLE001
        return _err_sanitized(e)


if __name__ == "__main__":
    mcp.run()