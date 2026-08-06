# document-agent

> 一个**与工具无关**的文档智能写作代理：理解自然语言需求，自动识别文档类型，按规范排版生成 **Word / Excel / PDF / PPT** 成品文档，支持批量生成、模板管理、格式互转与 Markdown 导出。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![CI](https://github.com/l0s3r-Q/document-agent/actions/workflows/ci.yml/badge.svg)](https://github.com/l0s3r-Q/document-agent/actions/workflows/ci.yml)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Tests](https://img.shields.io/badge/tests-46%20passed-brightgreen.svg)](tests/)

## ✨ 特色亮点

| 能力 | 说明 |
|------|------|
| 🧠 **自然语言驱动** | 说一句话就出文档：自动识别类型 → 决策结构 → 套用规范排版 |
| 📄 **四格式全覆盖** | `.docx` / `.xlsx` / `.pdf` / `.pptx` 统一入口，24 个工具 |
| 📐 **国标级排版** | 党政公文 GB/T 9704-2012、学位论文、合同等 10 类规范预置 |
| 🔄 **结构决策** | 从零生成 / 改结构（改造建议清单）/ 不改结构 / 模板导入 四种模式 |
| 📊 **批量生成** | 一个模板 + N 组数据（Excel 花名册 → 批量欢迎信） |
| 🔁 **格式互转** | docx ↔ xlsx 表格互转、docx → Markdown、docx → PDF |
| 📈 **数据可视化** | PPT 原生图表（柱/条/线/饼）、Excel 美化报表、大数字卡片 |
| 🛠 **模板管理** | 导入范文、重命名、删除、导出、双模板对比 |
| 🔥 **热重载** | 修改 MCP server 源码无需重启，下次调用自动生效 |
| 🔌 **与工具无关** | Skill 遵循 Agent Skills 规范，MCP 遵循标准协议，任何 AI 工具可接入 |

## 能力矩阵

| 功能 | Word (.docx) | Excel (.xlsx) | PDF (.pdf) | PPT (.pptx) |
|------|:---:|:---:|:---:|:---:|
| 生成/导出 | ✅ build_docx | ✅ build_excel | ✅ convert_to_pdf | ✅ build_pptx |
| 解析读取 | ✅ parse_docx | ✅ parse_excel | ✅ pdf_info | ✅ parse_pptx |
| 批量生成 | ✅ batch_build | ✅ 数据源 excel_to_data | — | — |
| 模板导入/复用 | ✅ import_template | — | — | — |
| 结构改造建议 | ✅ suggest_restructure | — | — | — |
| 格式互转 | ✅ ↔ xlsx / → Markdown | ✅ → docx 表格 | — | — |
| 图表 | — | 美化报表（表头/填充/筛选） | — | ✅ 原生图表（4 类型） |
| 页眉页脚/页码 | ✅ header/footer/PAGE | — | — | — |
| 质量门禁 | ✅ quality_check（AIGC 痕迹/占位符/排版） | ✅ 表头检查 | — | ✅ 空页/溢出 |
| 样式美化 | 字体/字号/行距/缩进 | 表头/填充/边框/筛选 | — | 10 版式/4 主题 |

## 快速上手

> 需要 **Python 3.10+**

### 1. 安装依赖

```bash
pip install -r mcp/document-toolkit/requirements.txt
```

### 2. 配置 MCP（写入你的工具配置或项目 .mcp.json）

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

### 3. 安装 Skill

将 `skills/document-intelligence` 放入工具的 skills 目录（或配置 skills 路径指向 `skills/` 目录）。

### 4. 验证

重启工具后，在对话中尝试：

```
帮我写一份关于开展安全生产检查的通知
```

若工具支持 MCP 工具列表查看，应看到 24 个 `document-toolkit` 工具。

## 工具总表（24 个）

### Word 工具（12 个）

| 工具 | 参数 | 功能 |
|------|------|------|
| `parse_docx` | path | 解析 docx：页面/样式/结构树 |
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

### Excel 工具（3 个）

| 工具 | 参数 | 功能 |
|------|------|------|
| `build_excel` | spec_json, output_path | 生成 xlsx（默认表头美化：加粗+浅蓝填充+边框+自适应列宽+冻结+筛选） |
| `parse_excel` | path, sheet_name? | 解析 xlsx：行数据/合并/列宽（公式缓存值） |
| `excel_to_data` | path, sheet_name? | Excel 数据表 → JSON 行数组（batch_build 数据源） |

### PDF 工具（2 个）

| 工具 | 参数 | 功能 |
|------|------|------|
| `convert_to_pdf` | docx_path, output_path? | docx → PDF（Word COM → WPS COM → LibreOffice 自动降级） |
| `pdf_info` | path | PDF 页数/大小 |

### PPT 工具（2 个）

| 工具 | 参数 | 功能 |
|------|------|------|
| `build_pptx` | spec_json, output_path | 生成 pptx（10 版式/4 主题，可编辑原生形状） |
| `parse_pptx` | path | 解析 pptx：页/形状/表格 |

### 互转工具（3 个）

| 工具 | 参数 | 功能 |
|------|------|------|
| `docx_tables_to_excel` | docx_path, output_path | docx 全部表格 → xlsx（每表一个 sheet） |
| `excel_to_docx` | excel_path, output_path, with_sheet_titles? | xlsx → docx 表格文档 |
| `docx_to_markdown` | docx_path, output_path? | docx → Markdown（标题/正文/表格） |
| `merge_pdfs` | pdf_paths_json, output_path | 合并多个 PDF（缺失自动跳过） |

### 质量检查工具（1 个）

| 工具 | 参数 | 功能 |
|------|------|------|
| `quality_check` | path | **交付质量体检**：AIGC 痕迹词/占位符残留/emoji/重复标点/句末标点/表格参差/表头空/字体缺失/标题跳级（docx）；空页/文本溢出（pptx）；表头重复/空表头（xlsx）。error 必须修复，warning 建议修复 |

## 支持的文档类型与格式

### 10 类 Word 文档规范

| 类型 | doc_type | 规范依据 |
|------|----------|---------|
| 党政公文 | official | GB/T 9704-2012（标题 2 号小标宋、正文 3 号仿宋、行距固定 28 磅、页边距 3.7/3.5/2.8/2.6cm） |
| 学位论文 | thesis | 通用高校规范（宋体小四、1.5 倍行距、GB/T 7714 参考文献） |
| 合同协议 | contract | 法定必备条款（当事人/标的/价款/履行/违约/争议）+ 通用排版 |
| 招投标文书 | bidding | 招标公告/投标文件标准结构 |
| 通用文档 | general | 报告/方案/纪要等默认规范（宋体小四、1.5 倍行距） |
| 法律文书 | legal | 起诉状/律师函/答辩状（仿宋四号） |
| 政府工作报告 | government_report | 政务版式（GB/T 9704，四段式结构） |
| 技术文档 | techdoc | 需求/设计/操作手册（微软雅黑标题、"1/1.1"编号） |
| 简历 | resume | 求职简历（24pt 姓名、紧凑行距） |
| 通知公告 | notice | 非公文类公告通知（黑体 18pt 标题） |

### PPT 十版式

| 版式 | 说明 |
|------|------|
| cover | 封面（标题/副标题/作者，黑色标题默认） |
| agenda | 目录页 |
| section | 章节页（编号圆点 `index` + 色带） |
| content | 内容页（要点列表 **或 cards 卡片化**） |
| stats | 大数字数据卡片（最多 4 卡并排） |
| chart | 原生图表（column/bar/line/pie） |
| two_column | 双栏对比 |
| table | 表格页（表头主色填充） |
| image | 图片页（缺失自动警告） |
| closing | 结尾页 |

主题：`corporate` 商务 / `academic` 学术 / `launch` 发布会（黑底白字）/ `minimal` 极简

## 用法示例

### 从零生成

| 输入 | 结果 |
|------|------|
| "写一份关于开展安全生产检查的通知" | 党政公文 GB/T 9704 排版（2 号小标宋、3 号仿宋、28 磅固定行距） |
| "拟一份软件开发委托合同" | 合同规范排版 + 必备条款骨架 |
| "写一份民事起诉状" | 法律文书排版（诉讼请求/事实与理由结构） |
| "做一份产品发布会的 PPT" | launch 主题（黑底金字） |
| "把这几个人的信息做成花名册 Excel" | 美化报表（表头填充/冻结/筛选） |

### 基于已有文件

| 输入 | 结果 |
|------|------|
| "把 D:/docs/方案.docx 改成招标公告" | suggest_restructure 改造建议 → 按招投标规范重组 |
| "在这个合同模板基础上，拟一份服务器采购合同" | 解析原结构 → 锁定结构 → 填新内容 |
| "润色一下这份总结（D:/docs/总结.docx）" | 保持结构，优化文字 |
| "导入 D:/docs/范文.docx，按它格式写报告" | 模板导入 → 沿用范文排版生成 |

### 批量与数据

| 输入 | 结果 |
|------|------|
| "批量生成 10 份部门会议通知" | 模板 + 数据行批量产出 |
| "用花名册（D:/data/花名册.xlsx）批量生成入职欢迎信" | excel_to_data → batch_build 联动 |

### 格式转换

| 输入 | 结果 |
|------|------|
| "把这份通知转成 PDF" | convert_to_pdf（Word/WPS/LibreOffice 自动选引擎） |
| "把合同里的表格导出成 Excel" | docx_tables_to_excel |
| "把这份报告转成 Markdown" | docx_to_markdown |
| "看看这个 PDF 有几页" | pdf_info |

## 工作原理（四层工作流）

1. **意图识别**：从自然语言提取文档类型、主题、要求（关键词表 + 规则）
2. **类型判定**：映射到 10 类文档规范或 3 类附加格式（最长关键词优先）
3. **结构决策**：未提供文件→从零生成；提供文件+改结构→改造建议；提供文件+不改结构→锁定结构；范文→导入模板
4. **计划与执行**：章节规划 → 内容要点 → 排版参数 → 生成 → **回读校验**（parse 验证结构/字体/行距）

## 工程化质量

| 维度 | 措施 |
|------|------|
| 🧪 测试 | 56 项自动化测试（生成/解析往返、健壮性、错误路径、互转、图表、页眉页脚、质量门禁） |
| 🔬 CI | GitHub Actions：Python 3.10/3.11/3.12 三版本矩阵 + LibreOffice PDF 测试 |
| 🔒 数据安全 | 原子写入（临时文件+replace 防损坏）、异常路径脱敏（`<path>`）、模板并发锁 |
| 🔥 热重载 | 源码 mtime 监听 + 依赖序 reload + 语法预检（改代码免重启） |
| 🛡 健壮性 | 错误路径全部返回 `{ok:false}` JSON、边界防护（空输入/超长/非法值回退） |

## 目录结构

```
document-agent/
├── skills/
│   └── document-intelligence/     # 文档智能写作 Skill
│       ├── SKILL.md               # 主 playbook（四层工作流 + 决策树 + 契约）
│       ├── templates/             # 10 类文档排版规范（人读）
│       └── examples/              # 典型输入示例（30+ 条）
├── mcp/
│   └── document-toolkit/          # 文档读写 MCP server（24 个工具）
│       ├── server.py              # FastMCP 入口 + 热重载机制
│       ├── docx_toolkit/          # Word 模块（解析/生成/模板/批量/Markdown）
│       ├── excel_toolkit/         # Excel 模块（生成/解析/数据源）
│       ├── pdf_toolkit/           # PDF 模块（三引擎转换）
│       └── pptx_toolkit/          # PPT 模块（生成/解析/图表）
├── docs/                          # architecture / usage / import-flow
├── examples/                      # 一键生成 10 类示例文档
├── tests/                         # 56 个自动化测试
├── LICENSE                        # MIT
└── README.md
```

## 常见问题

**Q: 生成 docx 后 Word 打开字体不对？**
A: 确认系统安装了目标字体（方正小标宋/仿宋_GB2312 等）。未安装时 Word 自动回退，可在 spec 的 `font` 中改用已安装字体（如"仿宋"代替"仿宋_GB2312"）。

**Q: PDF 转换失败或报"未检测到引擎"？**
A: 需要本机安装 Word / WPS / LibreOffice 三者之一。转换引擎自动降级：Word → WPS → LibreOffice。

**Q: PPT 标题为什么是黑色？**
A: 设计约定：标题与正文文字默认黑色，主题色仅用于装饰（标题条/表头/强调线）。需要彩色标题时用 slide 级 `title_color`（如 `"#1F4E79"`）指定。

**Q: 修改了 MCP server 代码，需要重启吗？**
A: 修改实现代码无需重启——server 内置热重载，下次工具调用自动加载新代码。仅新增/删除工具（工具列表变化）需要重启。

**Q: 能在没有 Office 的机器上生成文档吗？**
A: 可以。docx/xlsx/pptx 由 python-docx/openpyxl/python-pptx 直接写 OOXML，不依赖 Office；仅 PDF 转换需要 Office 套件之一。

**Q: Excel 数据怎么用来批量生成 Word？**
A: `excel_to_data` 把 Excel 首行作为字段名转 JSON 行数组 → 作为 `batch_build` 的 data_rows，模板中用 `{字段名}` 占位即可。

**Q: 怎么保证交付质量（没有 AI 腔、没有占位符）？**
A: 生成后调用 `quality_check` 做交付体检——检测 AIGC 痕迹词（综上所述/值得注意的是/作为AI 等）、占位符残留（{变量}/待补充）、emoji、表格参差、标题跳级等，error 级必须修复到 `pass=true`。SKILL 也内置了写作质量标准（禁 AI 腔/数据真实/交付前自检）。

**Q: 输出文件会覆盖已有文件吗？**
A: 所有生成采用原子写入（临时文件 + replace），生成过程异常不会损坏原文件；PDF 转换会先删除旧文件但仅在所有引擎均失败时才可能丢失（建议输出到新路径）。

## 路线图

- [x] 10 类文档预置规范与模板
- [x] 范文/规范文档导入功能（`import_template`）
- [x] 结构决策增强（`suggest_restructure` 改造建议）
- [x] Excel 生成/解析/数据源（`build_excel` / `parse_excel` / `excel_to_data`）
- [x] PDF 导出（`convert_to_pdf` 三引擎降级）
- [x] PPT 生成/解析（`build_pptx` / `parse_pptx`，10 版式/4 主题）
- [x] MCP server 热重载（改代码免重启）
- [x] 格式互转（docx ↔ xlsx 表格、docx → Markdown）
- [x] PPT 原生图表 + docx 页眉页脚页码
- [x] 测试驱动加固（56 项健壮性测试 + 多 Agent 审查）
- [x] 交付质量门禁（`quality_check` + SKILL 质量标准）
- [x] wps-office 在线编辑协同（`docs/wps-office-integration.md` + SKILL 协作章节 + 示例）
- [ ] 在线 API 服务（规划中）

## License

[MIT](LICENSE)

## 与 wps-office 的协同（在线编辑）

document-agent 负责**离线生成/解析/转换**规范文档；[wps-office](https://github.com/lc2panda/wps-skills) 负责**在线操作真实 WPS Office**（Excel 82 / Word 32 / PPT 115 工具）。两者组合形成**生成 → 打开 → 编辑 → 导出**闭环：

| 场景 | document-agent | wps-office |
|------|---------------|-----------|
| 生成即预览 | `build_docx/build_excel/build_pptx` 离线生成 | `wps_word_open_document` 打开预览 |
| 生成后精修 | 按规范排版初稿 | `wps_word_set_font`、`wps_excel_set_cell_format`、`wps_ppt_beautify` |
| 存量文档重排 | `parse_docx`/`build_docx` 重排 | `wps_word_get_document_text` 取内容 |
| 转 PDF/导出 | `convert_to_pdf`（三引擎降级） | `wps_common_save_as` / `wps_convert_to_pdf` |
| 批量复核 | `batch_build` 批量生成 | 打开抽查 |

**部署**：两个 MCP 可共存注册（`document-toolkit` + `wps-office`），见 `docs/wps-office-integration.md`。
**示例**：`examples/wps_workflow.py`（生成 → 质量检查 → 转 PDF → WPS 打开全链路）。

---

## 相关项目


- [ppt-master](https://github.com/hugohe3/ppt-master)（MIT © Hugo He）：高级 PPT 设计工作流（SVG 精细排版、AI 配图、模板填充增强），本项目的基础 PPT 能力与其互补
