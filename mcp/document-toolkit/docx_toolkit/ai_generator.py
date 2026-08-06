"""ai_generator.py —— AI 内容生成引擎（防 AI 味）。

输入：主题 + 文档类型 + 要点 → 调用 LLM 生成完整 DocumentSpec JSON（与 build_docx 兼容）。

设计要点：
1. 多 provider 支持：OpenAI 兼容 API（deepseek / mimo / grok / 自定义），通过环境变量读取 key 与 base_url
2. 防 AI 味：prompt 内置严格禁词清单（与 quality.py AIGC_PATTERNS 对齐），生成后用 quality.py 自检，不合格自动重生成
3. 零依赖：仅用标准库 urllib 发 HTTP（不新增 requirements）
4. 离线降级：无 API key 时返回结构化提示，不阻塞 build_docx 主流程
"""

from __future__ import annotations

import json
import os
import re
import urllib.request

# ── provider 配置（OpenAI 兼容 /chat/completions）──
# 环境变量约定：AI_GEN_PROVIDER 选择 provider，默认 deepseek
# 每个 provider 用 {NAME}_BASE_URL / {NAME}_API_KEY / {NAME}_MODEL
_PROVIDERS = {
    "deepseek": {
        "base_url_env": "DEEPSEEK_BASE_URL",
        "api_key_env": "DEEPSEEK_API_KEY",
        "model_env": "DEEPSEEK_MODEL",
        "default_base": "https://api.deepseek.com",
        "default_model": "deepseek-v4-flash",
    },
    "mimo": {
        "base_url_env": "MIMO_BASE_URL",
        "api_key_env": "MIMO_API_KEY",
        "model_env": "MIMO_MODEL",
        "default_base": "https://token-plan-cn.xiaomimimo.com/v1",
        "default_model": "mimo-v2.5",
    },
    "grok": {
        "base_url_env": "GROK_BASE_URL",
        "api_key_env": "GROK_API_KEY",
        "model_env": "GROK_MODEL",
        "default_base": "http://83.229.124.183:8888/v1",
        "default_model": "grok-4.3",
    },
}

# ── 防 AI 味：生成 prompt 中的硬约束（与 quality.py 对齐）──
_AI_FLAVOR_BAN = """绝对禁止以下 AI 腔/空洞表达（出现即视为不合格）：
1. 总结套话：总而言之 / 综上所述 / 总的来说 / 综上 / 总之
2. 提醒套话：值得注意的是 / 需要注意的是 / 需要指出的是
3. 空洞论断：毋庸置疑 / 众所周知 / 不言而喻 / 毫无疑问
4. 身份自述：作为人工智能 / 作为AI / 作为大模型 / 我是一个AI
5. 模板化排比：首先...其次...最后...（机械三段式）
6. 递进套话：不仅...而且...（滥用）
7. 开篇套话：在当今社会 / 随着时代的发展 / 随着社会的进步
8. 空洞修饰：意义重大 / 影响深远 / 十分重要的 / 具有深远意义
9. 占位符：待补充 / 待完善 / XXX / {变量}
10. 空洞结尾：让我们共同努力 / 相信在大家的共同努力下（除非是正式公文惯用收尾）"""

_SYSTEM_PROMPT = """你是一位资深中文公文/文档写作者，擅长写出自然、具体、有信息量的正式文档。

写作铁律（必须严格遵守）：
- 用具体事实、数据、细节支撑，不用空话套话
- 句子长短结合，避免机械排比和模板腔
- 逻辑自然衔接，不堆砌连接词
- 每段有实质内容，删除任何可删的废话
- 语言朴素准确，像有经验的办公室文书人员写的
- 严格避免以下 AI 腔（出现即不合格）：
{ai_ban}

输出要求：
- 只输出一个 JSON 对象（DocumentSpec），不要任何解释文字、markdown 代码块标记
- JSON 结构必须完全符合 DocumentSpec 契约"""

_SPEC_TEMPLATE = """{system}

任务：生成一份「{doc_type_name}」文档。
主题：{topic}
{extra}

请生成完整 DocumentSpec JSON，规则：
- doc_type: "{doc_type}"
- title: 文档标题（具体、准确）
- sections: 按该类型规范组织（标题/正文/列表/表格）
- 内容要充实具体：正文段落 3-6 句，列表项 3-8 条，全部实写不虚
- 表格数据要真实合理（材料未给的信息用占位说明，但不留 {{变量}} 占位符）
- JSON 字段见契约：{{"doc_type", "title", "author"(可省), "date"(可省), "sections":[{{"type": "heading1|heading2|paragraph|list|table", ...}}]}}
- 严禁输出 JSON 之外的任何文本"""


def _provider_cfg():
    """返回当前生效的 provider 配置。"""
    name = os.environ.get("AI_GEN_PROVIDER", "deepseek").lower()
    return _PROVIDERS.get(name, _PROVIDERS["deepseek"])


def _resolve(cfg, key: str, env: str, default: str) -> str:
    v = os.environ.get(env) or ""
    if v:
        return v
    return default


def _chat_completion(prompt: str, temperature: float = 0.7, max_tokens: int = 8192) -> str:
    """调用 OpenAI 兼容 chat completions API。返回 assistant 文本。

    兼容思考模型（deepseek-v4-flash 等）：content 为空时回退 reasoning_content。
    """
    cfg = _provider_cfg()
    base = _resolve(cfg, "base_url", cfg["base_url_env"], cfg["default_base"])
    api_key = os.environ.get(cfg["api_key_env"], "")
    model = _resolve(cfg, "model", cfg["model_env"], cfg["default_model"])

    if not api_key:
        raise RuntimeError(
            f"未配置 {cfg['api_key_env']} 环境变量（AI 生成需要 API key）。"
            "可设置 AI_GEN_PROVIDER 选择 deepseek/mimo/grok。"
        )

    base = base.rstrip("/")
    url = f"{base}/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": _SYSTEM_PROMPT.format(ai_ban=_AI_FLAVOR_BAN)},
            {"role": "user", "content": prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=180) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    msg = data["choices"][0]["message"]
    content = msg.get("content") or ""
    # 思考模型（deepseek-reasoner 类）：content 可能在 reasoning_content 中
    if not content.strip():
        content = msg.get("reasoning_content") or ""
    return content


def _extract_json(text: str) -> dict:
    """从 LLM 输出中提取 JSON 对象（容错：去掉 markdown 代码块/前后噪声）。"""
    text = text.strip()
    # 去掉 ```json ... ``` 包裹
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if m:
        text = m.group(1).strip()
    # 直接尝试解析
    try:
        obj = json.loads(text)
        if isinstance(obj, dict):
            return _normalize_spec(obj)
    except json.JSONDecodeError:
        pass
    # 截取第一个 { 到最后一个 }
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end != -1 and end > start:
        try:
            obj = json.loads(text[start : end + 1])
            if isinstance(obj, dict):
                return _normalize_spec(obj)
        except json.JSONDecodeError:
            pass
    raise ValueError("LLM 输出不是合法 JSON，无法解析")


def _normalize_spec(spec: dict) -> dict:
    """LLM 生成的 spec 字段归一化：兼容 content/text/title 等字段差异。

    契约字段：section 用 type + text（paragraph/list）或 type + text/title（heading*）。
    LLM 常见偏差：用 content 代替 text/title。这里统一映射。
    """
    if not isinstance(spec, dict):
        return spec
    sections = spec.get("sections")
    if isinstance(sections, list):
        for sec in sections:
            if not isinstance(sec, dict):
                continue
            # content → text（paragraph/list 主体）或 content → title（heading）
            if "content" in sec and "text" not in sec and "title" not in sec:
                stype = str(sec.get("type", ""))
                if stype.startswith("heading"):
                    sec["title"] = sec.pop("content")
                else:
                    sec["text"] = sec.pop("content")
            # items 兼容：list 类型可能有 items 或 content 数组
            if sec.get("type") == "list" and isinstance(sec.get("content"), list) and "items" not in sec:
                sec["items"] = sec.pop("content")
    return spec


def _has_ai_flavor(spec: dict) -> list[str]:
    """用 quality.py 的 AIGC 模式自检生成的 spec 文本。返回命中清单（空 = 干净）。"""
    hits = []
    try:
        from .quality import AIGC_PATTERNS, PLACEHOLDER_PATTERNS
    except ImportError:
        return hits
    text = json.dumps(spec, ensure_ascii=False)
    for pat, label in AIGC_PATTERNS + PLACEHOLDER_PATTERNS:
        if re.search(pat, text):
            hits.append(f"{label}：{pat}")
    return hits


def generate_spec(doc_type: str, topic: str, extra: str = "", retries: int = 2) -> dict:
    """生成 DocumentSpec JSON。

    doc_type: 文档类型（general/thesis/official/contract/.../meeting_minutes/speech/proposal/invitation）
    topic:    主题（如"关于开展安全生产检查的通知"）
    extra:    额外要求（要点/数据/约束，可空）
    retries:  防 AI 味自检不合格时重试次数
    """
    # 文档类型名称映射（用于 prompt）
    _NAMES = {
        "general": "通用文档", "thesis": "学位论文", "official": "党政公文",
        "contract": "合同协议", "bidding": "招投标文书", "legal": "法律文书",
        "government_report": "政府工作报告", "techdoc": "技术文档", "resume": "简历",
        "notice": "通知公告", "meeting_minutes": "会议纪要", "speech": "演讲稿",
        "proposal": "方案建议书", "invitation": "邀请函",
        "news_release": "新闻稿", "work_summary": "工作总结",
        "product_manual": "产品说明书", "acceptance_report": "验收报告",
    }
    doc_name = _NAMES.get(doc_type, doc_type)
    extra_section = f"额外要求：{extra}" if extra else ""

    prompt = _SPEC_TEMPLATE.format(
        system="", doc_type_name=doc_name, topic=topic, extra=extra_section, doc_type=doc_type
    )

    last_err = ""
    for attempt in range(retries + 1):
        try:
            raw = _chat_completion(prompt, temperature=0.7 if attempt == 0 else 0.9)
            spec = _extract_json(raw)
            # 防 AI 味自检
            hits = _has_ai_flavor(spec)
            if hits:
                last_err = f"含 AI 腔：{hits[0]}"
                # 追加纠正指令重试
                prompt += f"\n\n上一次生成不合格：{hits[0]}。请改写为自然、具体的表达，严禁该词。"
                continue
            return spec
        except Exception as e:  # noqa: BLE001
            last_err = str(e)
    raise RuntimeError(f"AI 生成失败（重试 {retries} 次）：{last_err}")


def is_configured() -> bool:
    """是否有可用的 provider 配置。"""
    cfg = _provider_cfg()
    return bool(os.environ.get(cfg["api_key_env"]))
