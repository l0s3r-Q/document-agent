# 模板导入功能设计

## 目标

用户提供**范文**或**规范说明文档**（.docx），系统解析其排版与结构，保存为用户模板，后续生成文档时沿用该排版——实现"按单位的格式标准出文档"。

## 流程

```
用户范文/规范文档 (.docx)
        │
        ▼
┌──────────────────────────┐
│ import_template(path,    │
│   template_name)         │
│  ├─ parse 页面设置        │  页边距
│  ├─ parse 段落样式        │  字体/字号/对齐/行距/缩进（按 role 归类，每类 1 条）
│  └─ 提取骨架              │  标题层级 + 段落大纲（前40项）
└──────────┬───────────────┘
           ▼
   用户模板 JSON 保存至
   mcp/docx-toolkit/user_templates/<name>.json
   （meta.source = "user"）
           │
           ▼
┌──────────────────────────┐
│ list_templates() 确认     │
│ 生成时：用户模板优先于     │
│ 预置模板（source=builtin）│
└──────────────────────────┘
```

## 模板 JSON 结构

```json
{
  "meta": {"name": "...", "doc_type": "user", "source": "user", "description": "从 <路径> 导入", "created_at": "..."},
  "page": {"top_cm": ..., "bottom_cm": ..., "left_cm": ..., "right_cm": ...},
  "styles": {"body": {"font_name": "...", "size_pt": ..., "bold": ..., "align": "..."}, "heading1": {...}},   # role → 样式
  "skeleton": [{"level": 0|1|2|3, "text": "章节文本"}]
}
```

## 决策优先级

1. 用户导入模板（source=user）> 预置模板（builtin）> 默认 general
2. `get_template` 仅服务预置类型；用户模板通过 `list_templates` 发现

## 验收标准

- [ ] 导入一份含多级标题+正文+表格的范文，模板可被 list_templates 列出
- [ ] 用导入模板生成的文档，页边距/字体/行距与范文一致（parse 回读对比）
- [ ] 导入含中文与西文字体混排的文档，eastAsia 字体正确保留
- [ ] 重复导入同名模板覆盖旧文件
- [ ] 损坏/非 docx 文件返回 {ok: false} 且不产生残留

## 后续增强

- 模板差异化对比（范文 vs 预置，输出差异报告）
- 模板目录管理（重命名/删除/导出）
- 多模板选择 UI（在 Agent 侧列出并确认）
