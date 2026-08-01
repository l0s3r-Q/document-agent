---
name: document-intelligence
description: 文档智能写作代理。理解用户自然语言需求，自动识别文档类型（通用文档/学位论文/党政公文/合同协议/招投标文书），决策结构策略（从零生成/改结构/不改结构），按规范排版生成 .docx；支持导入用户范文/规范说明文档制作成品。触发词：文档、公文、通知、报告、方案、论文、合同、协议、招标、投标、纪要、润色、改写、排版、格式、范文、模板导入。
---

# document-intelligence —— 文档智能写作代理

## 能力概述

输入用户自然语言 → 判定文档类型 → 决策结构策略 → 制定写作计划 → 调用 docx-toolkit MCP 解析/生成 → 输出成品 `.docx`。

四层工作流：

1. **意图识别**：从用户描述中提取文档类型、主题、要求
2. **类型判定**：映射到 5 类规范之一（official/thesis/contract/bidding/general）
3. **结构决策**：根据"是否有输入文件"和"是否要求改结构"选择策略
4. **计划与执行**：章节规划 → 内容要点 → 排版参数 → 生成 → 回读校验

## 一、意图识别规则

| 关键词 | 文档类型 | doc_type |
|--------|---------|----------|
| 通知、请示、报告、函、决定、意见、批复、公告、公文、红头、发文 | 党政公文 | official |
| 论文、学位论文、毕业论文、开题报告、学术论文、毕业设计 | 学位论文 | thesis |
| 合同、协议、合约、条款、甲乙双方、违约责任 | 合同协议 | contract |
| 招标、投标、招标公告、招标文件、投标文件、评标、中标 | 招投标文书 | bidding |
| 报告、方案、说明、总结、计划、会议纪要、周会纪要、邮件、简历、说明书、通用文档 | 通用文档 | general |
| 起诉状、答辩状、律师函、上诉状、法律文书、诉讼、仲裁申请 | 法律文书 | legal |
| 政府工作报告、政府汇报、人大报告、政府述职 | 政府工作报告 | government_report |
| 需求说明书、设计文档、操作手册、技术方案、接口文档、白皮书 | 技术文档 | techdoc |
| 简历、个人简介、求职、应聘 | 简历 | resume |
| 公告、通知公告、活动通知、停水停电通知、公司公告 | 通知公告 | notice |

判定规则：命中多个类型时，取**最先命中且最具体**的类型；无法判定时默认 general。
特别注意：**"报告/纪要"默认归 general**（如工作总结报告、会议纪要），除非明确出现"红头""发文""批复"等公文体裁词才判 official。

## 二、结构决策树（必须执行）

```
用户是否提供了参考/原文文件（.docx 路径）？
├── 未提供
│   └── 从零生成：get_template(doc_type) 取预置模板骨架 → 按用户主题填充章节 → build_docx
├── 提供了文件 + 明确要求"改结构/重组/重构/调整章节"
│   ├── suggest_restructure(路径, 目标doc_type) 生成改造建议（保留/新增/移除清单）
│   ├── parse_docx(路径) 解析内容与样式（含正文文本）
│   ├── 按改造建议 + 目标类型规范重新设计框架（保留可复用内容块）
│   └── build_docx(新框架, 输出路径)
├── 提供了文件 + "不改结构/在此基础上改/润色/按这个格式排"
│   ├── parse_docx(路径) 提取原结构+正文+样式（**必须用 parse_docx，它返回完整正文文本与样式；extract_structure 只返回标题大纲，无法支撑润色/替换内容**）
│   ├── 锁定原结构：只替换/优化各节内容，保持原结构
│   ├── 沿袭样式：把原文件每段样式映射为 section 级 font/align/bold 覆盖（见"样式映射表"）
│   └── build_docx(同结构, 输出路径)
└── 提供了文件 + 意图不明（只说"看看这个文件"）
    ├── 先询问用户要"改结构"还是"不改结构"
    ├── 用户不回答时默认：parse_docx 全文 + 不改结构，仅补全/润色
    └── 若用户明确说"导入这个作为模板"→ 走第三节模板导入流程
```

> `extract_structure` 仅适用于**完全重写内容、只借用标题骨架**的场景（如"按这个大纲重新写"）。

### 模糊表达参考表

| 用户说法 | 判定 |
|---------|------|
| "参考这个写" | 借鉴内容 + 按目标类型的新模板结构 |
| "在这个基础上改" | 保持原结构，改内容（parse_docx 取正文） |
| "按这个格式排" | 沿袭原文件样式（parse_docx 提取 → 逐段 font/align/bold 覆盖） |
| "帮我润色一下" | 保持结构，优化文字（parse_docx 取正文） |
| "把这份报告改成公文" | 改结构：按 official 模板重组 |
| "导入我的范文，按它的格式生成" | 模板导入流程（第三节） |

## 三、模板导入流程（用户范文/规范说明文档 → 成品）

1. `import_template(docx_path, template_name)`：解析范文的页面/样式（按 role 归类）/骨架，保存为用户模板
2. `list_templates()`：确认导入成功（返回 page/style_roles/skeleton_count 摘要）
3. 按用户新内容构建 DocumentSpec：
   - `page`：直接用导入模板的 page
   - `styles`：直接用导入模板的 styles（**build_docx 支持 spec.styles 按 role 覆盖预置排版**：`{"styles": {"body": {"font_name": "宋体", "size_pt": 12, ...}}}`）
   - `skeleton`：作为章节骨架，按"骨架→spec 转换规则"生成 sections
4. `build_docx(spec, 输出路径)` 生成
5. `parse_docx(输出路径)` 回读校验

> 用户导入模板的排版**优先于**预置模板：通过 spec.styles 覆盖实现。

## 四、5 类文档规范速查表

### 党政公文（GB/T 9704-2012）—— official
- 页边距：上 3.7cm、下 3.5cm、左 2.8cm、右 2.6cm
- 标题：2 号方正小标宋（22pt），居中
- 一级标题：3 号黑体（16pt）；二级标题：3 号楷体_GB2312；三级标题：3 号仿宋_GB2312 加粗
- 正文：3 号仿宋_GB2312（16pt），行距**固定值 28 磅**，首行缩进 2 字符
- 要素顺序：发文机关标志（红头）→ 发文字号 → 标题 → 主送机关 → 正文 → 附件说明 → 发文机关署名+成文日期（右对齐）→ 附注 → 抄送

### 学位论文 —— thesis
- 封面（校名/题目/作者/导师/日期）→ 独创性声明 → 中文摘要+关键词 → Abstract+Keywords → 目录 → 正文各章 → 结论 → 参考文献（GB/T 7714）→ 致谢 → 附录
- 正文：宋体小四（12pt），1.5 倍行距，首行缩进 2 字符
- 一级标题：黑体三号（16pt）；二级：黑体四号（14pt）；三级：黑体小四（12pt）

### 合同协议 —— contract
- 标题：二号宋体加粗居中；正文：宋体四号（14pt），1.5 倍行距
- 必备条款：当事人信息 → 鉴于条款 → 标的与数量 → 质量要求 → 价款与支付 → 履行期限/地点/方式 → 违约责任 → 保密 → 争议解决 → 签署栏（双方法定代表人/日期/盖章）

### 招投标文书 —— bidding
- 标题：二号黑体居中；正文：仿宋_GB2312 三号，固定 28 磅
- 招标公告：项目概况 → 资格要求 → 文件获取 → 投标截止 → 开标时间地点 → 联系方式
- 投标文件：投标函 → 报价表 → 技术方案 → 资质证明 → 授权委托书

### 通用文档 —— general
- 宋体小四（12pt）、1.5 倍行距、首行缩进 2 字符、标题黑体
- 结构：标题 → 导言 → 主体章节 → 结论 → 落款

### 法律文书 —— legal
- 标题二号宋体加粗居中；正文仿宋_GB2312 四号（14pt）1.5 倍行距
- 结构：标题 → 当事人信息 → 案由 → 诉讼请求/事实与理由 → 此致法院 → 落款（详见 templates/legal.md）

### 政府工作报告 —— government_report
- 沿用 GB/T 9704 版式：正文仿宋三号（16pt）固定 28 磅；标题二号小标宋
- 结构：回顾 → 总体要求与目标 → 重点工作 → 政府自身建设 → 结语（详见 templates/government_report.md）

### 技术文档 —— techdoc
- 微软雅黑标题 + 宋体小四正文（12pt）1.5 倍行距；编号用"1 / 1.1 / 1.1.1"式
- 结构：引言 → 需求/设计 → 实施 → 附录（详见 templates/techdoc.md）

### 简历 —— resume
- 姓名微软雅黑 24pt 居中；板块标题 14pt；正文宋体 11pt 1.2 倍行距
- 结构：个人信息 → 教育背景 → 工作经历 → 技能 → 自我评价（详见 templates/resume.md）

### 通知公告 —— notice（非党政公文）
- 标题黑体 18pt 加粗居中；正文宋体小四 1.5 倍行距
- 结构：标题 → 对象 → 事项/时间/要求 → 落款（详见 templates/notice.md）

## 五、与 docx-toolkit MCP 的协作契约

| 工具 | 参数 | 用途 |
|------|------|------|
| `parse_docx(path)` | path: str | 解析页面/样式/结构（改结构场景） |
| `extract_structure(path)` | path: str | 提取标题大纲（不改结构场景） |
| `build_docx(spec_json, output_path)` | spec_json: str, output_path: str | 生成 docx |
| `import_template(docx_path, template_name)` | docx_path: str, template_name: str | 导入范文模板 |
| `get_template(doc_type)` | doc_type: str | 读取预置模板（10 类） |
| `list_templates()` | — | 列出模板（含排版摘要） |
| `suggest_restructure(source_path, target_doc_type)` | source_path, target_doc_type: str | 结构决策增强：对比源文档与目标模板，输出改造建议（keep/add/remove） |
| `batch_build(spec_template_json, data_rows_json, output_dir, filename_field?)` | 见"批量生成"节 | 批量生成文档 |
| `rename_template(old_name, new_name)` | str, str | 重命名用户模板 |
| `delete_template(name)` | str | 删除用户模板 |
| `export_template(name, output_path)` | str, str | 导出模板 JSON |
| `compare_templates(name_a, name_b)` | str, str | 对比两模板页面/样式/骨架差异 |

### DocumentSpec JSON 契约（build_docx 输入）

```json
{
  "doc_type": "official|thesis|contract|bidding|general",
  "title": "文档标题",
  "author": "可选：作者/单位",
  "date": "可选：日期",
  "page": {"top_cm": 3.7, "bottom_cm": 3.5, "left_cm": 2.8, "right_cm": 2.6},
  "sections": [
    {"type": "heading1|heading2|heading3", "text": "章节标题"},
    {"type": "paragraph", "text": "正文段落"},
    {"type": "list", "items": ["条目1", "条目2"]},
    {"type": "table", "rows": [["列1", "列2"], ["值1", "值2"]]},
    {"type": "page_break"}
  ]
}
```

**关键点**：
- 不传 page/styles 时自动使用 doc_type 的预置排版
- `styles`（可选）：按 role 覆盖预置排版，形如 `{"styles": {"body": {"font_name": "宋体", "size_pt": 12, "line_spacing_rule": "MULTIPLE", "line_spacing_multiple": 1.5, "first_line_indent_chars": 2}}}`，role ∈ title/heading1/heading2/heading3/body/table；导入模板的 styles 可直接粘贴
- 每个 section 可带 `font`/`align`/`bold` 覆盖
- `align` 可选 CENTER/LEFT/RIGHT/JUSTIFY

### 骨架 → spec 转换规则

| 骨架元素 | 映射为 |
|---------|--------|
| level 1/2/3 | heading1/heading2/heading3（text 替换为真实章节标题） |
| level 0 且为"标题/封面/落款"类 | paragraph + 适当 align（标题 CENTER、落款 RIGHT） |
| level 0 且为"主送机关/签署栏"类 | paragraph（正文样式） |
| 骨架中带"表格"字样 | table（按内容构造 rows） |

### 样式映射表（沿袭原文件样式时用）

| parse_docx 输出字段 | DocumentSpec section 字段 |
|--------------------|--------------------------|
| font_name / east_asia | font.name（中文用 east_asia） |
| size_pt | font.size_pt |
| bold | bold |
| align | align |
| line_spacing_rule + line_spacing_pt/multiple | font.line_spacing_rule + font.line_spacing_pt / font.line_spacing_multiple |
| first_line_indent_pt / first_line_indent_chars | font.first_line_indent_pt / font.first_line_indent_chars |

### official 完整示例

```json
{
  "doc_type": "official",
  "title": "关于开展安全生产大检查的通知",
  "sections": [
    {"type": "heading1", "text": "一、检查范围"},
    {"type": "paragraph", "text": "本次检查覆盖全市所有在建工程项目。"},
    {"type": "heading1", "text": "二、时间安排"},
    {"type": "paragraph", "text": "自本通知印发之日起至2026年9月30日止。"}
  ]
}
```

### 公文成文日期与署名

**不要**用 spec 的 `date` 字段放公文的成文日期（builder 会将其居中）。正确做法：在 sections 中用右对齐段落表达：

```json
{"type": "paragraph", "text": "XX市人民政府办公室", "align": "RIGHT"},
{"type": "paragraph", "text": "2026年8月2日", "align": "RIGHT"}
```

## 五·五、批量生成（batch_build）

一个 spec 模板 + 多组数据批量产出文档：

```json
// spec_template：字符串中可用 {变量} 占位
{"doc_type": "notice", "title": "关于召开{meeting}会议的通知",
 "sections": [{"type": "paragraph", "text": "各{dept}：定于{date}召开{meeting}会议，请准时参加。"}]}
```

```json
// data_rows：每行一组变量
[{"meeting": "安全生产", "dept": "生产部门", "date": "8月10日"},
 {"meeting": "质量评审", "dept": "质检部门", "date": "8月15日"}]
```

- `filename_field` 指定用哪个字段作文件名（缺省用 title+序号）
- 返回 total/succeeded/failed 统计；未提供的变量替换为空并记入 warnings
- 适用：批量会议通知、批量证书、批量合同（一份模板 N 组当事人）

## 五·六、模板管理

| 工具 | 场景 |
|------|------|
| `rename_template` | 导入的模板命名不规范时重命名 |
| `delete_template` | 清理不再使用的用户模板（内置模板不可删） |
| `export_template` | 把用户模板导出为 JSON 分享/备份 |
| `compare_templates` | 比较两模板差异（选模板、排查排版漂移时用） |

## 六、执行约束

1. 输出格式仅 `.docx`（Word 2007+），路径用**绝对路径**
2. 调用 MCP 工具后**必须检查返回的 `ok` 字段**：`ok=false` 时向用户复述 `error` 并终止，不继续
3. 生成后**必须**调用 `parse_docx` 回读校验，检查点：标题数量与层级和 spec 一致、字号/行距/缩进数值符合预期、eastAsia 中文字体非空、无空正文段、无乱码
4. 中文字体由 build_docx 自动写入 w:eastAsia，**无需**在 spec 中额外设置
5. 用户未指定输出路径时，默认输出到桌面 `C:/Users/<user>/Desktop/文档/` 子目录（不存在则创建），文件名用 `类型_主题.docx`；中文路径原样传入
6. `spec_json` 参数需**序列化为 JSON 字符串**传入 build_docx（不是对象）
7. 注意 import_template 的截断：骨架取前 40 项、样式按 role 归类后每类取 1 条
8. 党政公文若涉及红头（发文机关标志/发文字号），按 templates/official.md 要素顺序补全，并将红头作为正文首段（居中）处理
9. 内容为空、路径不存在等情况先向用户确认，不臆造文件
