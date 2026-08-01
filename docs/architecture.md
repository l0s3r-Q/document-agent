# 架构设计

## 设计原则

1. **与工具无关**：Skill 遵循 Anthropic Agent Skills 规范（SKILL.md），MCP 遵循标准 MCP 协议（stdio）。任何支持这两种标准的 AI 工具都能消费本项目，不绑定 Reasonix / Claude Code / Cursor 等任何一家。
2. **大脑与手脚分离**：决策逻辑放 Skill（模型可读、可解释、可修改），文档读写放 MCP（确定性、可测试、可复用）。
3. **预置 + 导入双轨**：内置 10 类文档规范模板；用户可导入范文/规范说明文档覆盖预置排版。

## 分层

```
┌─ 意图层（Skill）─────────────────────────────────────────┐
│ document-intelligence（SKILL.md）                        │
│   1. 意图识别：关键词表 → 文档类型                        │
│   2. 类型判定：10 类（official/thesis/contract/bidding/general/legal/│
│      government_report/techdoc/resume/notice）               │
│   3. 结构决策树：无文件→从零；有文件+改结构→重组；         │
│      有文件+不改结构→锁定                                │
│   4. 计划执行：骨架 → 内容 → 排版 → 生成 → 回读校验       │
└──────────────┬──────────────────────────────────────────┘
               │ DocumentSpec JSON（结构化契约）
┌──────────────▼─ 执行层（MCP）────────────────────────────┐
│ document-toolkit（FastMCP / python-docx）                    │
│   parse_docx / extract_structure / build_docx            │
│   import_template / get_template / list_templates        │
│   排版参数源：docx_toolkit/templates/*.json（10 类预置）   │
└──────────────┬──────────────────────────────────────────┘
               │
┌──────────────▼─ 文件层 ──────────────────────────────────┐
│ .docx（Word 2007+）                                      │
└──────────────────────────────────────────────────────────┘
```

## 接口契约

### DocumentSpec（build_docx 输入）

```json
{
  "doc_type": "official|thesis|contract|bidding|general|legal|government_report|techdoc|resume|notice",
  "title": "文档标题",
  "author": "可选",
  "date": "可选",
  "page": {"top_cm": 3.7, "bottom_cm": 3.5, "left_cm": 2.8, "right_cm": 2.6},
  "sections": [
    {"type": "heading1|heading2|heading3", "text": "..."},
    {"type": "paragraph", "text": "..."},
    {"type": "list", "items": ["..."]},
    {"type": "table", "rows": [["...", "..."]]},
    {"type": "page_break"}
  ]
}
```

- `page`/`styles` 缺省时自动使用 `doc_type` 的预置排版
- 段落级 `font` 可覆盖（如 `{"font": {"size_pt": 14}}`）

### 模板 JSON（docx_toolkit/templates/*.json）

```json
{
  "meta": {"name": "...", "doc_type": "...", "source": "builtin|user"},
  "page": {"top_cm": ..., ...},
  "styles": [{"role": "title|heading1|heading2|heading3|body|table", "font_name": "...", "size_pt": ..., "bold": ..., "align": "...", "line_spacing_rule": "EXACTLY|MULTIPLE", "line_spacing_pt": ..., "first_line_indent_chars": 2}],
  "skeleton": [{"level": 0|1|2|3, "text": "章节占位"}]
}
```

## 关键实现点

- **中文字体**：python-docx 设置 `run.font.name` 只影响西文；中文必须同时写 `w:eastAsia`（`rFonts.set(qn('w:eastAsia'), name)`），否则 Word 显示默认字体
- **固定行距**：公文要求 28 磅固定值 → `line_spacing_rule = EXACTLY` + `line_spacing = Pt(28)`
- **首行缩进**：2 字符 ≈ 2 × 字号（用 Pt(size * 2) 近似）
- **标题识别**：生成时设置 Word 内置 `Heading 1-3` 样式，保证回读解析与大纲视图一致
- **结构顺序**：解析时按 body 子元素顺序遍历（段落与表格交错），保证结构树与排版顺序一致

## 部署拓扑

```
GitHub 仓库（本项目）         安装位置（用户机器）
document-agent/          →  各工具共享 skills 目录
├── skills/...           →  C:\Users\<user>\skills\skills\
└── mcp/document-toolkit/    →  C:\Users\<user>\skills\mcp\document-toolkit\
                              （MCP 配置指向该路径）
```

开发在仓库进行，发布时复制（或符号链接）到各工具的 skills/MCP 配置位置。
