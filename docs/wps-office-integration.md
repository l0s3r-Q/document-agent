# wps-office 集成：与 WPS 在线编辑能力协同

> document-agent 负责**离线生成/解析/转换**规范文档；[wps-office](https://github.com/lc2panda/wps-skills) 负责**在线操作真实 WPS Office**。
> 两者组合形成完整的"生成 → 打开 → 编辑 → 导出"文档工作流闭环。

## 为什么融合

| 能力维度 | document-agent（本仓库） | wps-office（外部 MCP） |
|---------|------------------------|----------------------|
| 文档生成 | ✅ `build_docx/build_excel/build_pptx` 离线生成（无需 WPS） | ❌ 无（只能操作已打开文档） |
| 规范排版 | ✅ 10 类模板（GB/T 公文/论文/合同等） | ⚠️ 有限（依赖文档已有样式） |
| 文档解析 | ✅ `parse_docx/parse_excel/parse_pptx` | ⚠️ 读取当前文档结构 |
| 在线编辑 | ❌ 无 | ✅ Excel 82 / Word 32 / PPT 115 工具 |
| 格式转换 | ✅ `convert_to_pdf/docx_to_markdown/互转` | ⚠️ 保存为其他格式 |
| 质量检查 | ✅ `quality_check`（AIGC 痕迹/占位符/排版） | ❌ 无 |
| 批量生成 | ✅ `batch_build`（模板+数据） | ❌ 无 |
| 图表/透视 | ✅ PPT 原生图表、Excel 美化报表 | ✅ Excel 图表/透视表、PPT 图表 |
| 模板管理 | ✅ 导入/导出/对比/重命名 | ❌ 无 |
| 运行前提 | Python 3.10+（纯离线） | WPS Office 运行中 + Node.js |

**结论：完全互补，无功能冲突。** 融合收益：

1. **生成即预览**：`build_docx` 生成 → `wps_word_open_document` 打开，用户立即看到成品
2. **生成后精修**：AI 按规范生成初稿 → WPS 中人工/在线微调（字体、间距、插图）
3. **存量文档重排**：`wps_word_get_document_text` 读取 WPS 已开文档 → `parse_docx`/`build_docx` 按规范重排
4. **一键转 PDF**：`convert_to_pdf`（本仓库，Word/WPS COM 降级链）或 wps-office 的转换工具

## 部署集成

### 1. 两个 MCP 共存注册（Reasonix config.toml）

```toml
# document-agent 的 document-toolkit（离线生成）
[[plugins]]
name    = "document-toolkit"
command = "python"
args    = ["C:\\Users\\36078\\skills\\mcp\\document-toolkit\\server.py"]

# wps-office（在线编辑，Windows 轮询模式）
[[plugins]]
name    = "wps-office"
command = "node"
args    = ["C:\\Users\\36078\\skills\\mcp\\wps-office\\dist\\index.js"]
env     = { WPS_USE_POLL = "1" }
```

两个 server 端口/进程完全独立，可同时运行。

### 2. Skills 共存

- `document-intelligence`（本仓库 skills/）——离线生成入口
- `wps-excel` / `wps-word` / `wps-ppt` / `wps-office`（wps-skills 仓库）——在线编辑入口

## 协作工作流

### 场景 A：生成 → 打开 → 转 PDF（最常用）

```
用户："写一份安全生产检查通知，生成 docx 并转 PDF 给我看"
1. document-toolkit.build_docx(spec, 输出.docx)     # 离线生成规范文档
2. document-toolkit.quality_check(输出.docx)          # 质量门禁（可选）
3. document-toolkit.convert_to_pdf(输出.docx)         # 转 PDF（Word/WPS COM 降级链）
4. wps-office.wps_word_open_document(输出.docx)       # 在 WPS 中打开供预览/编辑
```

### 场景 B：在 WPS 中编辑已打开文档并导出

```
用户："把我打开的这份通知美化一下，转成 PDF"
1. wps-office.wps_word_get_active_document           # 确认当前文档
2. wps-office.wps_word_get_document_text             # 读取内容
3. wps-office.wps_word_set_font / set_paragraph ...  # 在线调整格式
4. wps-office.wps_convert_to_pdf                     # 导出 PDF
```

### 场景 C：存量文档按规范重排（文档治理）

```
用户："把这份旧版合同按标准模板重排"
1. wps-office.wps_word_get_document_text             # 提取 WPS 中内容
2. document-toolkit.build_docx(contract spec, 新.docx) # 按合同模板重排
3. document-toolkit.quality_check                     # 质量检查
4. wps-office.wps_word_open_document(新.docx)         # 打开对比
```

### 场景 D：Excel 数据 → 批量文档 → WPS 复核

```
用户："根据花名册给每个人生成欢迎信，然后让我在 WPS 里看"
1. document-toolkit.excel_to_data(花名册.xlsx)        # 读数据
2. document-toolkit.batch_build(spec模板, 数据, 目录)  # 批量生成
3. wps-office.wps_excel_open_workbook(花名册.xlsx)    # 复核数据
4. wps-office.wps_word_open_document(欢迎信1.docx)    # 抽查成品
```

## wps-office 工具速查（融合常用）

### Word（wps_word_*）
- `open_document(filePath)` / `get_open_documents()` / `get_active_document()`
- `get_document_text(start?, end?)` / `insert_text(text, position)`
- `find_replace(find_text, replace_text)` / `find_in_document(find_text)`
- `set_font(font_name, font_size, bold, ...)` / `set_font_style(...)` / `set_paragraph(alignment, lineSpacing)`
- `insert_table(rows, cols)` / `insert_image(imagePath)` / `insert_header(text)` / `insert_footer(text)`
- `generate_toc(levels?)` / `enable_track_changes(enable)` / `set_page_setup(...)`
- `proofread_basic(text)` / `get_paragraphs(start?, end?)`
- `wps_common_save_as(filePath, format)` / `wps_convert_to_pdf(outputPath?)`

### Excel（wps_excel_*）
- `open_workbook(filePath)` / `get_open_workbooks()` / `create_workbook()`
- `set_cell_value(sheet, row, col, value)` / `get_cell_value(sheet, row, col)`
- `set_formula(sheet, range, formula)` / `get_cell_info(sheet, cell)`
- `write_range(sheet, range, data)` / `read_range(sheet, range)`
- `create_sheet(name)` / `rename_sheet(oldName, newName)` / `switch_sheet(name)`
- `create_chart(...)` / `create_pivot_table(...)` / `set_cell_format(...)`
- `wps_common_save_as(filePath, format)` / `wps_convert_to_pdf(...)`

### PPT（wps_ppt_*）
- `open_presentation(filePath)` / `get_open_presentations()` / `create_presentation()`
- `get_slide_count()` / `get_slide_info(slideIndex)` / `add_slide(title, layout)`
- `set_slide_title(slideIndex, title)` / `set_slide_content(...)` / `add_textbox(...)`
- `add_shape(type, ...)` / `insert_image(path, slideIndex)` / `add_table(...)`
- `beautify(color_scheme, font)` / `apply_transition_to_all(effect)`
- `wps_common_save_as(filePath, format)` / `export_slide_as_image(...)`

## 参数坑（实测记录）

- `wps_excel_open_workbook` 参数名是 **`filePath`**（不是 path）
- `wps_excel_set_formula` 参数名是 **`range`**（不是 cell）
- 通用保存是 **`wps_common_save_as`**（参数 `filePath` + `appType` + `format`）
- `get_active_workbook` 工具不存在，用 `get_open_workbooks` + `get_cell_info` 组合
- Word/PPT 命令在对应组件未打开时会**快速失败（约 12 秒）**，非 30 秒超时（v1.1+）

## 局限与注意事项

1. **wps-office 依赖 WPS 运行中**：先启动 WPS 组件（Word/Excel/PPT）→ 再调用对应工具；未打开的组件命令会超时
2. **Windows JS 宏桥**：wps-office 走轮询模式（`WPS_USE_POLL=1`），MCP 需常驻监听 127.0.0.1:58891；先启动 Reasonix 再打开 WPS
3. **文件锁**：WPS 打开的文件会被锁定，document-agent 的 `parse_docx/build_docx` 需在**关闭 WPS 文档**或**另存副本**后操作同一文件
4. **转换引擎**：document-agent 的 `convert_to_pdf` 有 Word COM → WPS COM → LibreOffice 降级链；若 WPS 占用 docx，建议先 `wps_common_save_as` 副本再转
5. **跨平台**：document-agent 全平台可用；wps-office 的轮询模式目前验证于 Windows
