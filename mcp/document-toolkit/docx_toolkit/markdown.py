"""markdown.py —— docx 结构 → Markdown 导出。"""

from __future__ import annotations

import os

from .parser import parse


def _escape_cell(v) -> str:
    """表格单元格转义（| 和换行）。"""
    s = str(v).replace("|", "\|").replace("\n", "<br>")
    return s


def to_markdown(docx_path: str, output_path: str | None = None) -> dict:
    """docx → Markdown：标题/正文/表格/列表。"""
    data = parse(docx_path)
    lines = []
    for item in data["structure"]:
        stype = item.get("type", "paragraph")
        text = item.get("text", "")
        if stype == "title":
            lines.append(f"# {text}\n")
        elif stype == "heading1":
            lines.append(f"\n## {text}\n")
        elif stype == "heading2":
            lines.append(f"\n### {text}\n")
        elif stype == "heading3":
            lines.append(f"\n#### {text}\n")
        elif stype == "separator":
            lines.append("\n---\n")
        elif stype == "page_break":
            lines.append("\n<!-- page-break -->\n")
        elif stype == "table":
            rows = item.get("rows", [])
            if rows:
                ncols = max(len(r) for r in rows)
                header = rows[0]
                lines.append("")
                lines.append("| " + " | ".join(_escape_cell(c) for c in header) + " |")
                lines.append("| " + " | ".join(["---"] * ncols) + " |")
                for r in rows[1:]:
                    cells = [_escape_cell(c) for c in r]
                    cells += [""] * (ncols - len(cells))
                    lines.append("| " + " | ".join(cells) + " |")
                lines.append("")
        elif stype == "list":
            for it in item.get("items", []):
                lines.append(f"- {it}")
            lines.append("")
        else:  # paragraph
            if text.strip():
                lines.append(f"{text}\n")

    md = "\n".join(lines)
    if output_path:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(md)
    return {"ok": True, "content": md if not output_path else None,
            "path": output_path, "chars": len(md)}
