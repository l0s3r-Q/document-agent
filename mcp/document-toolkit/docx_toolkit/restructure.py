"""restructure.py —— 结构决策增强：对比源文档与目标模板，生成改造建议。"""

from __future__ import annotations

from .parser import parse
from .templates_store import get_builtin


def suggest_restructure(source_path: str, target_doc_type: str) -> dict:
    """对比源文档标题大纲与目标类型模板骨架，输出改造建议。

    返回:
      {ok, source, target_doc_type, summary{keep,add,remove},
       items:[{action: keep|add|remove, level, text, reason}]}
    """
    tpl = get_builtin(target_doc_type)
    if tpl is None:
        return {"ok": False, "error": f"未知目标类型 {target_doc_type}"}

    # 源文档标题大纲（标题+正文段落首句作为内容线索）
    data = parse(source_path)
    source_items = []
    for s in data["structure"]:
        if s["type"].startswith("heading"):
            source_items.append({"level": int(s["type"][-1]), "text": s["text"]})
        elif s["type"] == "paragraph" and s.get("text") and len(s["text"]) > 4:
            source_items.append({"level": 0, "text": s["text"]})

    # 目标模板骨架
    target_items = [{"level": x.get("level", 0), "text": x["text"]} for x in tpl.get("skeleton", [])]

    import re as _re

    _SEQ_RE = _re.compile(
        r"^(?:第[一二三四五六七八九十百0-9]+[章节条款部分]|"
        r"[一二三四五六七八九十]+、|[0-9]+(?:\.[0-9]+)*[.、．]?|"
        r"[（(][一二三四五六七八九十0-9]+[)）])\s*"
    )

    def norm(t: str) -> str:
        """归一化：去首部序号、空白与标点，用于模糊匹配。"""
        if not t:
            return ""
        t = t.strip()
        t = _SEQ_RE.sub("", t)
        t = t.replace(" ", "").replace("　", "").replace("：", "").replace(":", "")
        return t

    # 匹配：目标骨架条目是否能在源中找到语义相近项
    items = []
    keep, add, remove = 0, 0, 0
    matched_ids = set()

    # 源标题与正文分开：短目标词（<4 字）只与源标题匹配，避免误伤正文
    src_titles = [s for s in source_items if s["level"] > 0]
    src_bodies = [s for s in source_items if s["level"] == 0]

    def find_hit(ttext: str) -> dict | None:
        nt = norm(ttext)
        if not nt:
            return None
        short = len(nt) < 4
        pool = src_titles if short else source_items
        for si in pool:
            st = si["text"]
            if not st:
                continue
            ns = norm(st)
            if not ns:
                continue
            if nt == ns or (len(nt) >= 4 and (nt in ns or ns in nt)):
                return si
        return None

    for ti in target_items:
        ttext = ti["text"]
        hit = find_hit(ttext)
        if hit:
            items.append({"action": "keep", "level": ti["level"], "text": ttext,
                          "reason": f"源文档已有对应内容（{hit['text'][:30]}）"})
            matched_ids.add(id(hit))
            keep += 1
        else:
            items.append({"action": "add", "level": ti["level"], "text": ttext,
                          "reason": "目标模板要求，源文档缺失，需新增"})
            add += 1

    # 源中有但目标骨架未覆盖的章节 → 提示移除或合并（level 不要求严格相等）
    for si in src_titles:
        if id(si) in matched_ids:
            continue
        if any(norm(ti["text"]) and (norm(ti["text"]) in norm(si["text"]) or norm(si["text"]) in norm(ti["text"]))
               for ti in target_items):
            continue
        items.append({"action": "remove", "level": si["level"], "text": si["text"],
                      "reason": "目标模板未包含此章节，建议删除或并入相关章节"})
        remove += 1

    return {
        "ok": True,
        "source": source_path,
        "target_doc_type": target_doc_type,
        "target_template": tpl.get("meta", {}).get("name", target_doc_type),
        "summary": {"keep": keep, "add": add, "remove": remove},
        "items": items,
    }
