# docx-toolkit MCP Server

文档读写工具箱：解析 `.docx` 结构/样式、按规范生成文档、导入范文为模板。

## 安装

```bash
pip install -r requirements.txt
```

## MCP 配置（.mcp.json）

```json
{
  "mcpServers": {
    "docx-toolkit": {
      "command": "python",
      "args": ["C:/Users/<your-username>/skills/mcp/docx-toolkit/server.py"]
    }
  }
}
```

> 将 `<your-username>` 替换为实际部署位置（Windows 用户目录）。若 `python` 不在 PATH，请用完整路径（如 `C:/Users/<user>/AppData/Local/Programs/Python/Python311/python.exe`）。

## 版本要求

- Python 3.10+
- `mcp>=1.26,<2`：mcp 2.x 移除了 FastMCP API，本 server 基于 FastMCP 实现
- `python-docx>=1.1`：1.1+ 才支持固定行距（EXACTLY）Length 赋值

## 安全说明

本 server 提供任意路径的 docx 读写能力（parse/build/import 均接受绝对路径），**仅供受信任的本地客户端调用**。请勿将其暴露到不可信网络或远程 MCP 网关。

## 工具一览

| 工具 | 参数 | 说明 |
|------|------|------|
| `parse_docx` | path | 解析页面/样式/结构树 |
| `extract_structure` | path | 提取标题大纲（"不改结构"场景） |
| `extract_structure` | path | 提取标题大纲（"不改结构"场景） |
| `build_docx` | spec_json, output_path | 按 DocumentSpec 生成 docx |
| `import_template` | docx_path, template_name | 导入范文为模板 |
| `get_template` | doc_type | 读取预置模板（10 类） |
| `list_templates` | — | 列出全部模板（含排版摘要） |
| `suggest_restructure` | source_path, target_doc_type | 结构改造建议（keep/add/remove） |
| `batch_build` | spec_template_json, data_rows_json, output_dir, filename_field? | 批量生成 |
| `rename_template` | old_name, new_name | 重命名用户模板 |
| `delete_template` | name | 删除用户模板 |
| `export_template` | name, output_path | 导出模板 JSON |
| `compare_templates` | name_a, name_b | 对比两模板差异 |

## DocumentSpec 示例

```json
{
  "doc_type": "official",
  "title": "关于开展安全生产检查的通知",
  "page": {"top_cm": 3.7, "bottom_cm": 3.5, "left_cm": 2.8, "right_cm": 2.6},
  "sections": [
    {"type": "heading1", "text": "一、检查范围"},
    {"type": "paragraph", "text": "本次检查覆盖所有在建项目。"}
  ]
}
```

预置排版（10 类）：党政公文 GB/T 9704-2012、学位论文、合同协议、招投标文书、通用文档、法律文书、政府工作报告、技术文档、简历、通知公告。
