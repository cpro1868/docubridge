# 文档双向转换工具 PRD（完整版）

> 版本：v1.1 
> 最后更新：2026-04-03 
> 修订说明：根据用户反馈，明确样式定义采用声明式配置文件，工具动态创建 Word 样式；确认了全部待定问题。

---

## 0. 相关设计文档

- 设计细化文档：`docs/superpowers/specs/2026-04-07-docubridge-render-design.md`

---

## 1. 项目背景

本项目旨在解决 AI 大模型处理文档时的格式壁垒问题。当前，大模型对 Markdown 格式的理解和生成能力远强于 Word、PDF、Excel 等二进制或复杂排版格式，但日常办公、学术写作、企业知识库中大量文档仍以这些格式存在。同时，大模型生成的 Markdown 结果在导出为 Word 等正式文档时，往往需要大量手动调整格式，效率低下。

因此，我们需要开发一个**双向文档转换工具**：

- **正向转换**：将 Word、Excel、PPT、PDF 等常见文件格式**无损转换为 Markdown**，供大模型读取、RAG 检索或进一步分析。
- **反向转换**：将 Markdown（尤其是大模型生成的 Markdown）按照用户指定的样式模板，**渲染为符合专业排版要求的 Word 文档**，支持多级编号、自定义样式集等高级功能。

工具定位为面向个人开发者、知识工作者及小型团队的**本地优先、跨平台**解决方案，提供 CLI 和可选 GUI 两种交互形式。

---

## 2. 产品目标

### 2.0 v1 发布聚焦（2026-04-10 更新）

为尽快进入首个可发布版本，`v1.0` 的发布判断只围绕以下三项能力：

- `Word (.docx) -> Markdown` 稳定可用，作为正向转换主线
- `Markdown -> Word (.docx)` 稳定可用，作为反向转换主线
- `Markdown -> Word` 的模板定制与 YAML 样式配置能力达到可交付水平

以下能力在代码库中可以继续演进，但**不作为 `v1.0` 是否发布的阻塞条件**：

- `.xlsx -> Markdown`
- `.pptx -> Markdown`
- `.pdf -> Markdown`
- GUI
- 批量处理

### 2.1 核心目标

- 支持常见办公/文档格式（.docx, .xlsx, .pptx, .pdf）到 Markdown 的高保真转换，保留标题层级、列表、表格、图片引用等核心结构。
- 支持 Markdown 到 Word 的转换，允许用户通过**样式配置文件（YAML）** 控制输出文档的字体、段落、多级编号等格式，工具**动态创建** Word 样式。
- 提供命令行接口（CLI），便于集成到自动化脚本、AI Agent 工作流中。
- 提供可选的桌面图形界面（GUI），降低非技术用户的使用门槛。
- 所有转换在本地完成，不上传用户数据，保障隐私与安全。

### 2.2 非目标

- 不做 PDF 的复杂版式还原（如多栏混排、图文环绕），仅保证文本流和表格的正确提取。
- 不做 PPT 动画、备注等非核心内容的转换。
- 不直接支持 .odt、.rtf 等小众格式（可通过 Pandoc 间接支持）。
- 不提供云端转换服务（除非后续作为独立商业版本）。

---

## 3. 用户角色

### 3.1 普通用户（知识工作者）

- 需要将客户发来的 Word/PDF 转成 Markdown，以便放入个人知识库或喂给 AI 分析。
- 需要将大模型生成的 Markdown 报告转为符合公司规范的 Word 文档，避免手动调格式。

### 3.2 开发者/高级用户

- 通过 CLI 将工具集成到 CI/CD 流水线、RAG 预处理脚本中。
- 编写自定义样式配置文件（YAML），为不同文档类型（论文、标书、技术手册）保存样式集。
- 扩展解析器以支持新的站点或特殊格式（如 ePub）。

### 3.3 管理员（可选，如果提供 Web 管理端）

- 管理预置的样式模板库。
- 查看转换日志和错误报告。

---

## 4. 产品范围

### 4.1 输入格式范围（第一阶段 / 产品全景）

| 格式  | 支持程度   | 说明                                                         |
| ----- | ---------- | ------------------------------------------------------------ |
| .docx | 完整支持   | 标题、段落、列表、表格、图片、基本样式                       |
| .xlsx | 部分支持   | 每个 Sheet 转为 Markdown 表格；图表转为数据表；忽略宏、公式  |
| .pptx | 部分支持   | 每页标题+正文列表；忽略动画、备注、任意形状                  |
| .pdf  | 有条件支持 | 文字版 PDF 可提取文本流和表格；扫描版需外接 OCR（Tesseract） |
| .md   | 完整支持   | 用于反向转换流程                                             |

### 4.1.1 v1.0 发布范围（严格口径）

`v1.0` 的发布范围只包含以下主路径：

| 路径 | 发布级别 | 发布说明 |
| ---- | -------- | -------- |
| `.docx -> .md` | 核心发布能力 | 必须覆盖标题、段落、列表、表格、图片引用、基础行内样式保留 |
| `.md -> .docx` | 核心发布能力 | 必须覆盖主要 Markdown 结构渲染与样式控制 |
| YAML 样式配置 | 核心发布能力 | 必须支持默认值、元素级样式、覆盖与校验 |
| `--template` 模板协同 | 核心发布能力 | 必须支持模板读取、样式映射、模板与 YAML 合并规则 |

以下格式或能力在 `v1.0` 中的定位为“实验性/非发布主卖点”：

- `.xlsx -> .md`
- `.pptx -> .md`
- `.pdf -> .md`
- `batch`
- GUI

### 4.2 输出格式范围

- **正向输出**：标准 Markdown（GFM 风格），文件扩展名 .md。
- **反向输出**：Microsoft Word .docx（兼容 Office 2007+）。

### 4.3 功能范围

- **正向转换**：从上述输入格式提取元数据并生成 Markdown。
- **反向转换**：将 Markdown 按**用户提供的样式配置文件**生成 Word 文档，工具**动态创建**所需的样式（字体、段落、编号、表格等），不再依赖用户预先制作 Word 模板文件。但为兼容旧习惯，仍可支持可选的外部模板文件（若提供，则优先使用模板中的样式定义）。
- **样式配置**：用户通过 YAML 定义样式集，包括：
  - 全局默认字体、字号、行距
  - 各 Markdown 元素（标题、正文、列表、代码块、表格等）的样式属性（字体、颜色、间距、边框、背景等）
  - 多级列表的编号格式、缩进、链接样式
- **批量处理**：支持递归转换文件夹，可设置遇错继续或停止。
- **日志与错误处理**：记录转换失败的文件及原因。

---

## 5. 关键业务定义

### 5.1 正向转换的“保真度”层级

- **一级保真（必须）**：文本内容、标题层级、段落顺序、列表结构、表格数据。
- **二级保真（尽量）**：图片引用路径、表格合并单元格（转为拆分+注释）、代码块语言标识。
- **三级保真（可选）**：字体、颜色、背景等样式信息（默认忽略，因为 Markdown 不承载这些）。

### 5.2 样式声明式配置（核心）

用户不再需要手动制作 Word 模板（`.dotx`），而是编写一个**样式配置文件**（例如 `style.yaml`），该文件包含所有样式定义。工具在生成 Word 文档时，会根据配置动态创建相应的样式对象（`doc.styles.add_style(...)`），并将 Markdown 元素映射到这些样式。

**配置文件结构示例**：

```yaml
name: "academic_thesis"
description: "学术论文样式（宋体、章节目录编号）"

defaults:
  font_name: "Times New Roman"
  font_size: 12
  line_spacing: 1.5

elements:
  heading1:
    based_on: "Normal"
    font_name: "黑体"
    font_size: 18
    bold: true
    color: "#2E75B6"
    space_before: 24
    space_after: 6
    outline_level: 0   # 大纲级别，用于目录
  heading2:
    based_on: "heading1"
    font_size: 16
    space_before: 18
    space_after: 6
    outline_level: 1
  heading3:
    based_on: "heading2"
    font_size: 14
    outline_level: 2
  paragraph:
    based_on: "Normal"
    alignment: left
    first_line_indent: 0.75 cm
  code_block:
    font_name: "Consolas"
    font_size: 10
    background_color: "#F5F5F5"
    border: none
    spacing_before: 6
    spacing_after: 6
  table:
    style_name: "TableGrid"   # 可使用内置样式，也可自定义
    header_repeat: true
    border: "single"
  unordered_list:
    based_on: "Normal"
    bullet_type: "•"
    left_indent: 1.27 cm
    hanging_indent: 0.63 cm
  ordered_list:
    based_on: "Normal"
    number_format: "%1."
    left_indent: 1.27 cm
  quote:
    font_italic: true
    left_indent: 1 cm
    color: "#666666"

multilevel_list:
  enabled: true
  levels:
    - level: 1
      element: heading1
      number_format: "%1."
      linked_style: "heading1"
      start_at: 1
    - level: 2
      element: heading2
      number_format: "%1.%2."
      linked_style: "heading2"
    - level: 3
      element: heading3
      number_format: "%1.%2.%3."
      linked_style: "heading3"
  indentation:
    level1: 0 cm
    level2: 0.75 cm
    level3: 1.27 cm
```

**解析规则**：
- `defaults` 段定义全局默认属性，可被具体元素继承。
- `elements` 段定义每个 Markdown 元素对应的样式属性，支持继承（`based_on` 指向另一个元素名或内置样式如 `Normal`）。
- `multilevel_list` 定义标题的多级编号，`element` 字段引用 `elements` 中定义的样式名。
- 工具会解析该配置，动态创建 Word 样式，并将 Markdown 节点映射到对应样式。

**兼容旧方式**：如果用户同时提供了 `--template` 参数（Word 模板文件），则优先使用模板中已存在的样式；若模板中缺失某样式，则根据配置文件动态创建。若同时存在模板和配置文件中的样式定义，以配置文件中的 `defaults` 和 `elements` 为准（用户显式配置覆盖模板）。

### 5.3 多级列表生成规则

- 仅当 `multilevel_list.enabled = true` 时启用。
- 工具根据 Markdown 标题层级（`#` 对应 level 1）自动应用对应的编号格式，**无需用户在 Markdown 中手动编写编号**。
- 普通有序列表（`1. item`）与多级列表无关，使用 `ordered_list` 样式。

### 5.4 样式集（Profile）管理

- 一个样式配置文件即为一个样式集（Profile）。
- 用户可通过 `--profile <name>` 或 `--style <file>` 指定。
- 工具内置 2-3 个示例样式集（如 `academic`、`business`、`code`）。

### 5.5 去重与覆盖规则

- 正向转换时，若输出 Markdown 文件已存在，默认覆盖（可添加 `--no-clobber` 选项跳过）。
- 反向转换时，若输出 Word 文件已存在，默认提示用户确认覆盖（GUI 弹窗，CLI 使用 `-f` 强制覆盖）。

---

## 6. 功能需求

### 6.1 正向转换模块

#### 6.1.1 Word (.docx) 转换

- 读取文档中的所有段落，识别 `Heading 1`~`Heading 6` 样式，转为 `#`~`######`。
- 列表项（`List Paragraph` 样式）转为 Markdown 无序列表 `-` 或有序列表 `1.`，保留嵌套层级。
- 表格转为 Markdown 管道表格，合并单元格拆分为重复内容并添加注释。
- 图片提取为 `![alt](image_path)`，图片文件默认保存在与 Markdown 同级的 `assets/` 文件夹。
- 页眉、页脚、文本框、脚注内容忽略（可选项：转为 Markdown 注释）。

#### 6.1.2 Excel (.xlsx) 转换

- 每个工作表（Sheet）生成一个二级标题 `## SheetName`，下方输出 Markdown 表格。
- 表格第一行作为表头。
- 长文本单元格（>100 字符）截断并添加 `...(truncated)`，完整内容作为脚注。
- 日期单元格转为 ISO 8601 格式（如 `2026-04-03`），避免歧义。
- 图表对象忽略，但可在注释中注明“此处有图表，数据源自区域 X:Y”。

#### 6.1.3 PowerPoint (.pptx) 转换

- 每张幻灯片生成三级标题 `### Slide N`。
- 提取标题占位符作为四级标题 `####`。
- 提取正文占位符作为无序列表。
- 备注内容转为引用块 `> 备注：...`。
- 忽略艺术字、任意形状、动画。

#### 6.1.4 PDF 转换

- **文字版 PDF**：
  - 使用 `pdfplumber` 提取每页文本，按 Y 坐标排序段落。
  - 检测多栏布局（通过 X 坐标聚类），按栏交替排列恢复阅读顺序。
  - 表格提取：识别有线表格转为 Markdown 表格；无线表格通过文本对齐启发式转换。
- **扫描版 PDF**：
  - 用户通过 `--ocr` 开启 OCR（需安装 Tesseract）。
  - 同时保留原图占位符 `![扫描页 N](page_N.png)` 和识别出的文本（两者都输出），由用户选择后续使用。
- 忽略页眉、页脚、页码（可选项：保留为注释）。

#### 6.1.5 通用要求

- 所有转换结果必须保持原文的**时间顺序**（对于文档）或**逻辑顺序**（对于 PPT/Excel）。
- 转换失败时，返回非零退出码，并在 stderr 输出错误信息。

### 6.2 反向转换模块（Markdown → Word）

#### 6.2.1 核心流程

1. 解析 Markdown 为抽象语法树（AST）。
2. 加载用户指定的样式配置文件（YAML）。
3. **动态创建 Word 样式**：遍历配置中的 `elements`，调用 `python-docx` 的 `styles.add_style()` 创建或修改样式，设置字体、段落格式、编号等。
4. 如果用户提供了外部模板文件（`--template`），则先加载该模板，对于模板中已存在的样式名不再覆盖；缺失的样式仍动态创建。
5. 遍历 AST 节点，根据 `elements` 映射创建对应段落/表格/图片，并应用已创建的样式。
6. 若启用了多级列表，创建 `ListTemplate` 并关联到对应的标题段落。
7. 保存 Word 文档。

#### 6.2.2 样式配置属性支持

| 类别      | 属性                | 示例值                       | 说明                 |
| --------- | ------------------- | ---------------------------- | -------------------- |
| 字体      | `font_name`         | "宋体", "Arial"              |                      |
|           | `font_size`         | 12 (pt)                      |                      |
|           | `bold`              | true/false                   |                      |
|           | `italic`            | true/false                   |                      |
|           | `underline`         | true/false                   |                      |
|           | `color`             | "#333333", "auto"            | RGB 十六进制         |
| 段落      | `alignment`         | left, center, right, justify |                      |
|           | `line_spacing`      | 1.5, 2, 等                   | 倍数或磅值           |
|           | `space_before`      | 12 (pt)                      | 段前间距             |
|           | `space_after`       | 6 (pt)                       | 段后间距             |
|           | `first_line_indent` | 0.75 cm                      | 首行缩进             |
|           | `left_indent`       | 1.27 cm                      | 左缩进               |
|           | `hanging_indent`    | 0.63 cm                      | 悬挂缩进             |
| 背景/边框 | `background_color`  | "#F5F5F5"                    | 段落底纹             |
|           | `border`            | single, double, none         | 段落边框（简化）     |
| 列表      | `bullet_type`       | "•", "-", "✓"                | 无序列表符号         |
|           | `number_format`     | "%1.", "(%1)"                | 有序列表编号格式     |
| 表格      | `style_name`        | "TableGrid", "Light Shading" | 使用 Word 内置表样式 |
|           | `header_repeat`     | true/false                   | 标题行重复           |
| 其他      | `outline_level`     | 0-8                          | 大纲级别，用于目录   |

**继承机制**：元素可以通过 `based_on` 继承另一个元素的所有属性，并覆盖部分属性。例如 `heading2` 继承 `heading1` 并减小字号。

#### 6.2.3 多级列表动态创建

由于 `python-docx` 对多级列表的支持较底层，实现辅助函数根据 `multilevel_list` 配置创建 `ListTemplate`：
- 为每个级别设置编号格式（`%1`, `%1.%2` 等）。
- 设置每个级别链接的样式（即 `elements` 中的样式名）。
- 设置缩进和起始编号。
- 将该 `ListTemplate` 注册到文档中，并用于对应的标题段落。

若用户提供了外部模板且模板中已包含所需的多级列表样式（按名称匹配），则直接使用模板中的列表定义，不动态创建。

#### 6.2.4 代码块处理

- 代码块在 Word 中通过插入一个 1 行 1 列的表格来模拟背景色，表格样式为无边框，单元格背景色为浅灰色，字体为 Consolas（或配置中指定的等宽字体）。

#### 6.2.5 图片处理

- Markdown 中的 `![alt](path)` 会被插入到 Word 文档中，图片宽度默认 500px，居中对齐。
- 若图片路径是相对路径，则相对于 Markdown 文件所在目录解析。

#### 6.2.6 脚注处理（第一阶段）

- 忽略原生脚注，将 `[^1]` 转为普通文本 `[^1]`，不生成 Word 脚注。后续版本支持。

### 6.3 CLI 命令设计

提供以下子命令：

```bash
# 正向转换
docubridge parse <input> [-o <output>] [--format <format>] [--ocr] [--verbose]

# 反向转换（核心）
docubridge render <input.md> -o <output.docx> \
    --style <style.yaml> \
    [--template <optional_template.dotx>] \
    [--stop-on-error]

# 样式集管理
docubridge style init                    # 在当前目录生成示例样式配置文件
docubridge style validate <file.yaml>    # 验证样式配置文件语法
docubridge style list                    # 列出内置样式集
docubridge style show <name>             # 显示样式集内容

# 批量处理
docubridge batch <input_dir> -o <output_dir> --parse|--render --style <style.yaml> [--stop-on-error] [--jobs 4]

# 辅助命令
docubridge version
```

### 6.4 GUI 桌面应用（第二阶段）

- 基于 Tauri 或 Electron 开发，提供文件拖拽、格式选择、样式集选择、进度条、日志显示。
- 调用后台 CLI 执行转换，不重复实现核心逻辑。
- 打包为 Windows `.exe`、macOS `.dmg`、Linux `.AppImage`。优先 Windows，保持跨平台能力。

---

## 7. 非功能需求

### 7.1 性能

- 单文件转换时间：10MB 以内的 .docx 转 Markdown 不超过 2 秒；100 页文字 PDF 转 Markdown 不超过 10 秒（不含 OCR）。
- 批量转换时，支持多核并行处理（通过 `--jobs` 参数指定并发数）。

### 7.2 兼容性

- 操作系统：Windows 10+、macOS 11+、Linux（Ubuntu 20.04+）。
- 输出 Word 文档兼容 Office 2007 及以上。
- 样式配置文件编码：UTF-8。

### 7.3 可扩展性

- 解析器采用插件化架构：新增格式只需实现 `Parser` 接口并注册。
- 样式配置属性可扩展，未来支持更多 Word 特性（如制表位、分页控制）。

### 7.4 错误处理与日志

- 默认遇错继续处理其余文件，最后汇总失败列表。
- 提供 `--stop-on-error` 选项，遇到第一个错误即停止。
- 日志级别支持 DEBUG、INFO、WARNING、ERROR。
- 默认输出到 stderr，可通过 `--log-file` 输出到文件。

### 7.5 安全性

- 所有文件处理在本地完成，不上传任何数据。
- 解析 Office 文件时，禁用外部内容（如 OLE 对象、宏），防止恶意代码。

---

## 8. 验收标准

系统达到 v1.0 可用状态，必须满足以下条件：

### 8.0 v1.0 发布判定原则

`v1.0` 是否可发布，不以“所有规划格式都已支持”为判断标准，而以“核心双向主线是否稳定、模板定制是否可交付”为判断标准。

满足以下三条即可进入 `v1.0`：

- `Word -> Markdown` 主场景稳定
- `Markdown -> Word` 主场景稳定
- 模板定制与 YAML 样式配置能力完整、可解释、可验证

`.xlsx/.pptx/.pdf` 相关能力若未完全收口，不阻塞 `v1.0`，但必须在文档中明确支持边界。

### 8.1 正向转换

以下为 `v1.0` 发布必须满足的正向转换验收条件：

- 能正确将 `.docx`（含标题、列表、表格、图片）转换为 Markdown，无内容丢失。
- 能保留 `.docx` 标题、段落、列表、表格、图片引用的基本结构顺序。
- 能在标题、段落、表格单元格中保留基础行内样式：粗体、斜体、删除线、链接、基础行内代码。

以下为 `v1.x` 后续增强项，不作为 `v1.0` 发布阻塞条件：

- `.xlsx` 的多 Sheet 转换、长文本提示、日期标准化
- `.pptx` 的页面结构恢复与富文本增强
- 文字版 PDF 提取、多栏阅读顺序恢复、OCR 接入

### 8.2 反向转换

以下为 `v1.0` 发布必须满足的反向转换验收条件：

- 用户仅提供一个 YAML 样式配置文件（无任何 Word 模板），工具能生成符合配置中字体、段落、编号要求的 Word 文档。
- 多级列表（如 `1. 标题1` → `1.1 标题2`）正确生成，编号连续。
- 代码块背景色正确呈现（浅灰底纹、等宽字体）。
- 表格边框、表头重复符合配置。
- 若同时提供外部模板和样式配置，优先使用模板中已有的样式，缺失的由配置动态创建；若冲突以配置文件为准。
- 图片正确插入并保持比例。

模板定制功能在 `v1.0` 中必须额外满足：

- 用户能通过 YAML 明确控制标题、正文、引用、列表、表格、代码块等核心元素样式
- 样式优先级规则清楚且稳定：`CLI overrides > YAML element settings > YAML defaults > template > system defaults`
- 样式或模板错误有明确诊断，不允许静默降级为不可解释结果
- 至少提供 2 到 3 套可直接复用的样式示例
- 提供模板协同说明，明确模板负责宿主样式资源，YAML 负责显式样式意图

### 8.3 CLI

以下为 `v1.0` 发布必须满足的 CLI 验收条件：

- 所有子命令按设计工作，帮助信息清晰。
- 错误时返回非零退出码，并输出可读错误信息。

`batch` 相关能力可以在 `v1.0` 后续版本交付，不阻塞首次发布。

### 8.4 样式集管理

- 能列出、创建、切换样式集。
- 内置至少 2 个示例样式集（如学术论文、商务报告）。

### 8.5 批量处理

- 不作为 `v1.0` 发布阻塞项，转入 `v1.x` 路线图。

### 8.6 文档

- 提供 README（安装、快速开始）。
- 提供完整的命令参考和样式配置文件编写指南。

---

## 9. 当前确认的产品决策（汇总）

| 决策项         | 结论                                                         |
| -------------- | ------------------------------------------------------------ |
| 样式定义方式   | **声明式配置文件（YAML）**，工具动态创建 Word 样式，不强制用户提供模板文件 |
| PDF 扫描件处理 | 同时保留原图占位符和 OCR 文本，由用户选择                    |
| Excel 日期格式 | 转为 ISO 8601 标准格式（如 `2026-04-03`）                    |
| 多级列表       | 工具根据标题层级自动生成，用户不应在 Markdown 中手动写编号   |
| 脚注           | 第一阶段转为普通文本 `[^1]`，后续版本支持原生脚注            |
| 平台支持       | 优先 Windows，保持跨平台（Linux/macOS）能力                  |
| 样式冲突       | 若同时提供模板和配置文件，以配置文件中的 `defaults` 和 `elements` 为准（用户显式配置覆盖模板） |
| 批量错误处理   | 默认继续，提供 `--stop-on-error` 选项                        |
| 容器化         | 暂不考虑 Docker，面向客户端                                  |

---

## 10. 附录

### 10.1 术语表

- **IR（中间表示）**：抽象语法树，统一表示文档结构。
- **正向转换**：从 Word/Excel/PPT/PDF 到 Markdown。
- **反向转换**：从 Markdown 到 Word。
- **样式集（Profile）**：一组完整的样式配置，对应一种文档类型（如学术论文）。
- **多级列表**：Word 中的章节自动编号功能（1, 1.1, 1.1.1…）。

### 10.2 参考文档

- [Python-docx 文档](https://python-docx.readthedocs.io/)
- [Markdown 语法（GFM）](https://github.github.com/gfm/)
- [YAML 1.2 规范](https://yaml.org/spec/1.2/spec.html)

---

**文档结束**
