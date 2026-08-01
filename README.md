# document-agent

> 一个与工具无关的文档智能写作代理：理解自然语言需求，自动识别文档类型，按规范排版生成 `.docx`。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![CI](https://github.com/l0s3r-Q/document-agent/actions/workflows/ci.yml/badge.svg)](https://github.com/l0s3r-Q/document-agent/actions/workflows/ci.yml)

## 它能做什么

你只需要说一句话：

- "帮我写一份关于开展消防安全检查的通知"
- "写一份毕业论文，题目是《基于深度学习的入侵检测研究》"
- "把这份技术方案改成招标文件"
- "导入我的公文模板，以后都按这个格式写"

系统会自动：**判定文档类型**（党政公文 / 学位论文 / 合同协议 / 招投标文书 / 通用文档 / 法律文书 / 政府工作报告 / 技术文档 / 简历 / 通知公告）→ **决策结构策略**（从零生成 / 改结构 / 不改结构）→ **按规范排版**（字体、字号、行距、页边距）→ **输出成品 .docx**。

## 架构

```
┌─────────────────────────────┐
│  document-intelligence skill │  大脑：意图识别 + 类型判定
│  (SKILL.md + 排版规范库)      │  + 结构决策 + 写作计划
└──────────────┬──────────────┘
               │ DocumentSpec JSON
┌──────────────▼──────────────┐
│  document-toolkit MCP server    │  手脚：解析 / 生成 / 模板导入
│  (python-docx, FastMCP)      │  （GB/T 9704 等预置排版）
└──────────────┬──────────────┘
               │
               ▼
          成品 .docx
```

- **Skill**（`skills/`）：遵循 [Anthropic Agent Skills](https://docs.anthropic.com/en/docs/agents-and-tools/agent-skills) 规范，任何支持 SKILL.md 的 Agent 均可加载
- **MCP**（`mcp/`）：标准 MCP 协议（stdio），任何支持 MCP 的 IDE / Agent 均可接入（Claude Code、Cursor、Cline、VS Code Copilot、Reasonix……）

## 快速开始

> 需要 **Python 3.10+**

```bash
# 1. 安装依赖
pip install -r mcp/document-toolkit/requirements.txt
```

```json
// 2. 配置 MCP（写入你的工具配置或项目 .mcp.json）
{
  "mcpServers": {
    "document-toolkit": {
      "command": "python",
      "args": ["<绝对路径>/mcp/document-toolkit/server.py"]
    }
  }
}
```

```bash
# 3. 安装 skill：将 skills/document-intelligence 放入工具的 skills 目录
#    （或配置 skills 路径指向 skills/ 目录）
```

## 用法示例

| 输入 | 结果 |
|------|------|
| "写一份关于开展安全生产检查的通知" | 党政公文 GB/T 9704 排版（标题 2 号小标宋、正文 3 号仿宋、行距 28 磅固定） |
| "拟一份软件开发委托合同" | 合同规范排版 + 必备条款骨架 |
| "把 D:/docs/方案.docx 改成招标公告" | 改造建议 → 按招投标规范重组 |
| "批量生成 10 份部门会议通知" | 模板 + 数据批量产出 |
| "导入 D:/docs/范文.docx，按它格式写报告" | 模板导入 → 沿用范文排版生成 |

## 目录结构

```
document-agent/
├── skills/
│   └── document-intelligence/     # 文档智能写作 Skill
│       ├── SKILL.md               # 主 playbook
│       ├── templates/             # 10 类文档排版规范（人读）
│       └── examples/              # 典型输入示例
├── mcp/
│   └── document-toolkit/              # 文档读写 MCP server
│       ├── server.py              # FastMCP 入口（17 个工具）
│       ├── docx_toolkit/          # Word 模块（解析/生成/模板）
│       ├── excel_toolkit/          # Excel 模块（生成/解析/数据源）
│       └── pdf_toolkit/            # PDF 模块（三引擎转换）
├── docs/                          # 架构与使用文档
├── examples/                      # 示例
├── LICENSE                        # MIT
└── README.md
```

## 支持的文档类型

| 类型 | doc_type | 规范依据 |
|------|----------|---------|
| 党政公文 | official | GB/T 9704-2012（3号仿宋、28磅固定行距） |
| 学位论文 | thesis | 通用高校规范（宋体小四、1.5倍行距、GB/T 7714 参考文献） |
| 合同协议 | contract | 法定必备条款 + 通用排版 |
| 招投标文书 | bidding | 招标/投标文件标准结构 |
| 通用文档 | general | 报告/方案/纪要等默认规范 |
| 法律文书 | legal | 起诉状/律师函/答辩状 |
| 政府工作报告 | government_report | 政务版式（GB/T 9704） |
| 技术文档 | techdoc | 需求/设计/操作手册 |
| 简历 | resume | 求职简历 |
| 通知公告 | notice | 非公文类公告通知 |

## 路线图

- [x] 10 类文档预置规范与模板
- [x] 范文/规范文档导入功能（`import_template`）
- [x] 更多文档类型（法律文书、政府报告、技术文档、简历、通知公告）
- [ ] 结构决策增强（自动对比原文件与目标模板的差异）
- [x] PDF 导出（convert_to_pdf 三引擎降级）
- [ ] Markdown 导出
- [ ] 在线 API 服务

## License

[MIT](LICENSE)
