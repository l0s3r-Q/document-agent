# 使用手册

## 安装

### 依赖

```bash
pip install -r mcp/document-toolkit/requirements.txt
```

### 配置到各工具

**Reasonix**（`C:\Users\<user>\AppData\Roaming\reasonix\config.toml`）：

```toml
[skills]
paths = ["C:\Users\<user>\skills\skills"]

[[plugins]]
name = "document-toolkit"
command = "python"
args = ["C:\Users\<user>\skills\mcp\document-toolkit\server.py"]
```

**Claude Code**：`~/.claude/skills/` 下为 `skills/document-intelligence` 建符号链接；`~/.claude.json` 的 `mcpServers` 添加 document-toolkit。

**Cursor / VS Code / Cline / Roo**：各自的 MCP 设置中添加：

```json
{
  "mcpServers": {
    "document-toolkit": {
      "command": "python",
      "args": ["<绝对路径>/mcp/document-toolkit/server.py"]
    }
  }
}
```

## 使用流程

### 场景 A：从零生成

1. 用户说："帮我写一份关于开展安全生产检查的通知"
2. Skill 判定 official → 取预置模板骨架
3. 编排章节内容 → `build_docx(spec, "C:/out/通知.docx")`
4. `parse_docx` 回读校验 → 报告完成

### 场景 B：基于文件改结构

1. 用户说："把 D:/docs/方案.docx 改成招标公告"
2. `parse_docx(D:/docs/方案.docx)` 解析原内容
3. 按 bidding 模板重新设计框架，复用内容块
4. `build_docx` 生成 → 校验

### 场景 C：基于文件不改结构

1. 用户说："在这个合同模板基础上，拟一份服务器采购合同"
2. `extract_structure(模板路径)` 锁定结构
3. 按原结构填入新内容 → 生成

### 场景 D：导入范文

1. 用户说："导入我的公文模板 D:/docs/我的模板.docx"
2. `import_template(D:/docs/我的模板.docx, "我的公文模板")`
3. 之后生成时沿用导入模板的排版（优先于预置）

## MCP 工具参考

| 工具 | 参数 | 返回 |
|------|------|------|
| parse_docx | path | {ok, file, page, styles, structure} |
| extract_structure | path | {ok, file, outline, paragraph_count} |
| build_docx | spec_json, output_path | {ok, path} |
| import_template | docx_path, template_name | {ok, template, path} |
| get_template | doc_type | {ok, template} |
| list_templates | - | {ok, templates}（含排版摘要） |
| suggest_restructure | source_path, target_doc_type | {ok, summary, items}（改造建议） |
| batch_build | spec_template_json, data_rows_json, output_dir, filename_field? | {ok, total, succeeded, failed} |
| rename_template | old_name, new_name | {ok, path} |
| delete_template | name | {ok, deleted} |
| export_template | name, output_path | {ok, path} |
| compare_templates | name_a, name_b | {ok, diff} |
| build_excel | spec_json, output_path | {ok, path, sheets} |
| parse_excel | path, sheet_name? | {ok, sheets} |
| excel_to_data | path, sheet_name? | {ok, rows} |
| convert_to_pdf | docx_path, output_path? | {ok, path, engine} |
| pdf_info | path | {ok, pages, size_bytes} |

## 常见问题

**Q: 生成后字体不对？**
A: 确认系统安装了目标字体（方正小标宋/仿宋_GB2312 等）。未安装时 Word 会自动回退，可在 spec 的 `font` 中改用已安装字体（如"仿宋"代替"仿宋_GB2312"）。

**Q: 中文路径报错？**
A: MCP 工具接收绝对路径即可，python-docx 支持中文路径；但建议配置文件路径避免空格与特殊字符。

**Q: 生成慢？**
A: 首次启动 MCP server 有 Python 加载开销（约 1-2 秒），后续调用很快。

**Q: 能在没有 Word 的机器上生成吗？**
A: 可以。python-docx 直接写 OOXML，不依赖 Word 安装；打开文档时才需要 Office/WPS。
