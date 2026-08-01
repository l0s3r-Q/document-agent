# document-intelligence

文档智能写作代理 Skill：理解自然语言 → 判定文档类型 → 决策结构策略 → 按规范生成 `.docx`。

## 结构

```
document-intelligence/
├── SKILL.md                  # 主 playbook（四层工作流 + 决策树 + 规范速查 + MCP 契约）
├── templates/                # 10 类文档的排版规范说明（人读）
│   ├── official.md           # 党政公文 GB/T 9704-2012
│   ├── legal.md              # 法律文书
│   ├── government_report.md  # 政府工作报告
│   ├── techdoc.md            # 技术文档
│   ├── resume.md             # 简历
│   ├── notice.md             # 通知公告
│   ├── thesis.md             # 学位论文通用规范
│   ├── contract.md           # 合同协议
│   ├── bidding.md            # 招投标文书
│   └── general.md            # 通用文档
└── examples/
    └── prompt-examples.md    # 典型用户输入与判定示例
```

## 依赖

- 需要 `docx-toolkit` MCP server（本仓库 `mcp/docx-toolkit/`）提供解析/生成/导入能力
- 机器可读的预置模板 JSON 在 MCP server 的 `docx_toolkit/templates/` 下，通过 `get_template` 工具读取

## 安装

1. 将本目录放入工具的 skills 目录（或配置 skills 路径指向其父目录）
2. 配置 docx-toolkit MCP（见 `mcp/docx-toolkit/README.md`）
