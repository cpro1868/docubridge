# AI时代的文档翻译官（docubridge）

**文档格式的翻译官 —— 让 Markdown 和 Word 说同一种语言，顺便讨好 AI。**

> *因为你不应该在手动调格式这件事上浪费时间，尤其是在 AI 已经帮你写了初稿的情况下。*

|📄 [English](README.md) | [中文](README_CN.md) |

---

## 这是什么？

你有没有遇到过这种情况：

- 🤖 让 AI 帮你写了一份报告，AI 输出的是 Markdown
- 👔 领导说："挺好的，发我一份 Word 吧"
- 😤 你开始一个个手动调整标题、字体、缩进……

**Docubridge 就是来解决这个问题的。**

它只做两件事，但做得很好：

| 方向 | 场景 |
|------|------|
| **Word/Excel/PPT → Markdown** | 把办公文档转成 AI 能读懂的 Markdown，喂给 AI 分析、做 RAG、或者只是想让 AI 帮忙改改 |
| **Markdown → Word** | 把 AI 生成的 Markdown 渲染成有模有样的 Word 文档，自带样式、自定义字体、自动编号，看起来像是你亲手排的版 |

---

## 谁会用这个？

### 📝 AI 辅助写作人

你用 AI 写报告、写方案、写文档。AI 输出 Markdown，但你公司要 Word。Docubridge 可以把 Markdown 按你指定的样式渲染成 `.docx`，不用动手。

### 🔬 研究人员

有一堆 Word 文档要喂给 AI 研究助手？先 parse 成 Markdown。保留标题、列表、表格、图片，丢掉二进制格式那些糟心事。

### 🏢 办公室打工人

需要把同事发来的 Word 文档转成 AI 能读的格式？或者把 AI 写的草稿变成看起来正式的文件？Docubridge 帮你搞定。

### 👨‍💻 开发者 / 脚本狂人

把文档转换集成到脚本、CI/CD 流水线、RAG 预处理流程里。CLI 优先、脚本友好、不上传数据。

---

## 安装

**要求：Python 3.12 或更高版本**

---

### 方式一：下载安装包安装（推荐用于正式使用）⭐

从 release 页面下载 wheel 文件，然后安装：

```bash
pip install ./docubridge-0.2.0-py3-none-any.whl
```

或者指定完整路径：

```powershell
# Windows
pip install .\docubridge-0.2.0-py3-none-any.whl

# macOS / Linux
pip install ./docubridge-0.2.0-py3-none-any.whl
```

安装后验证：

```bash
docubridge --help
```

---

### 方式二：从源码安装（用于开发和测试）🔧

如果你想修改代码或运行测试，可以从源码安装：

#### Windows

```powershell
# 1. 创建虚拟环境
python -m venv .venv

# 2. 激活它
.\.venv\Scripts\Activate.ps1

# 3. 安装
python -m pip install --upgrade pip
pip install -e .
```

#### macOS / Linux

```bash
# 1. 创建虚拟环境
python3 -m venv .venv

# 2. 激活它
source .venv/bin/activate

# 3. 安装
python -m pip install --upgrade pip
pip install -e .
```

#### 安装测试依赖（可选）

```bash
pip install -e .[dev]
```

---

### 安装后验证

```bash
docubridge doctor
```

看到 "Environment OK" 就说明没问题。

---

## 快速上手

### 场景："同事发我一个 Word 文档"

```bash
# 把 Word 转成 Markdown（AI 的食物！）
docubridge parse 同事的报告.docx -o 同事的报告.md

# 现在可以喂给任何 AI 工具了
```

### 场景："AI 写完草稿了，我需要一份正经的 Word"

```bash
# 方式一：使用内置样式
docubridge render AI草稿.md -o 最终报告.docx --style academic

# 方式二：从公司模板提取样式（推荐！）
docubridge extract-styles 公司模板.docx -o 公司样式.yaml
docubridge render AI草稿.md -o 最终报告.docx --style 公司样式.yaml --template 公司模板.docx
```

### 场景："先看看这工具有啥功能"

```bash
# 看看有哪些内置样式
docubridge style list

# 看看学术样式长啥样
docubridge style show academic

# 渲染一个测试文档
docubridge render tests/fixtures/sample.md -o 测试输出.docx --style tests/fixtures/style.yaml
```

---

## 核心功能

### 🎯 `parse` — 把 Office 文档转成 Markdown

把 Word、Excel、PPT 文件转成干净的、AI 能读懂的 Markdown。

```bash
docubridge parse 文档.docx -o 输出.md
docubridge parse 表格.xlsx -o 输出.md
docubridge parse 幻灯片.pptx -o 输出.md
```

**它会提取这些内容：**

- 📑 **标题** — `Heading 1` 变成 `#`，`Heading 2` 变成 `##`，以此类推
- 📝 **段落** — 包括基本格式（粗体、斜体、链接、行内代码）
- 📋 **列表** — 有序和无序，支持简单嵌套
- 📊 **表格** — 转成 Markdown 管道表格
- 🖼️ **图片** — 提取到 `assets/` 文件夹，在 Markdown 里引用
- 📑 **Excel 工作表** — 每个 Sheet 生成一个带表格的章节

**如果需要 JSON 格式的输出：**

```bash
docubridge parse 文档.docx -o 输出.md --json
```

---

### ✨ `render` — 把 Markdown 渲染成 Word（完整指南）

这是最核心的功能。把你写的 Markdown 变成有模有样的 Word 文档，支持真实样式。

#### 基础用法

```bash
# 使用内置样式
docubridge render 草稿.md -o 最终版.docx --style academic

# 使用自定义样式文件
docubridge render 草稿.md -o 最终版.docx --style 我的样式.yaml

# 同时使用样式和模板（效果最好）
docubridge render 草稿.md -o 最终版.docx --style 我的样式.yaml --template 公司模板.docx
```

---

#### 搞懂 `--style` 和 `--template` 的区别

这是最重要的概念。这两个参数配合使用，但职责不同：

| 参数 | 作用 | 可以把它理解为 |
|------|------|---------------|
| `--style` | YAML 样式配置文件 | **规则手册** — 告诉 Docubridge 如何把 Markdown 映射到 Word 样式 |
| `--template` | Word 文档（.docx） | **样式库** — 提供可复用的 Word 样式资源 |

**简单理解：**

```
Markdown = 内容结构（写什么）
YAML     = 映射规则（什么结构用什么样式）
Template = Word 样式（那些样式到底是什么样子）
```

**举例说明：**

当你的 Markdown 里有：
```markdown
# 第一章 概述
```

YAML 告诉你：
```yaml
heading1:
  template_style: Heading 1
```

Template 提供实际的 `Heading 1` 样式（你公司的字体、颜色、间距）。

结果：一个完美应用了你公司品牌的一级标题。✓

---

#### 三种使用方式

**方式一：内置样式（最简单）**

```bash
docubridge render 草稿.md -o 输出.docx --style academic
```

内置样式：`academic`（学术）、`business`（商务）、`default`（默认）

适用场景：快速出结果、测试、简单文档。

---

**方式二：只使用自定义 YAML（不需要模板）**

```bash
docubridge render 草稿.md -o 输出.docx --style 我的样式.yaml
```

你自己写 YAML 文件定义样式，Docubridge 会动态创建 Word 样式。

适用场景：便携配置、多台机器共享样式。

**最简 YAML 示例：**

```yaml
defaults:
  font_name: Times New Roman
  font_size: 12

elements:
  heading1:
    template_style: Heading 1
    font_size: 18
    bold: true
  paragraph:
    template_style: Normal
    font_size: 12
```

---

**方式三：YAML + 模板（正式使用推荐）⭐**

```bash
docubridge render 草稿.md -o 输出.docx --style 我的样式.yaml --template 公司模板.docx
```

把 YAML 映射规则和你公司现有的 Word 模板结合起来。

适用场景：公司文档、学术论文、需要统一品牌形象的正式文件。

**为什么这是最佳方式：**

1. **模板** 提供可复用的 Word 样式（字体、颜色、编号、间距）
2. **YAML** 指定哪些 Markdown 元素使用哪些样式
3. **结合使用**：产出的文档符合你组织的规范标准

---

#### 优先级规则

当 YAML 和模板都定义了同一个属性时：

```
YAML 显式值  >  模板值  >  默认值
```

大白话：你在 YAML 里明确写的值，优先级最高。否则用模板的。都没有才用默认值。

**举例：**
```yaml
# YAML 说字号是 14
heading1:
  font_size: 14
```

即使模板里的 `Heading 1` 设置了字号 16，最终输出也会用 14，因为 YAML 优先级更高。

---

#### 优先级详解：三个参数各管什么

| 参数 | 职责 | 包含内容 |
|------|------|----------|
| **Markdown** | 内容本身 | 标题、段落、列表、表格、代码块等结构 |
| **YAML (--style)** | 映射规则 + 显式样式 | 哪些结构 → 哪些 Word 样式 + 字体/大小等属性 |
| **Template (--template)** | Word 样式资源 | 实际的 Word 样式定义（字体、颜色、段落格式、编号等） |

**三者配合的典型场景：**

1. Markdown 里有 `# 标题`
2. YAML 说 `heading1` 要绑定到 `Heading 1` 样式，并设置 `bold: true`
3. Template 里的 `Heading 1` 有宋体字体、18pt 字号、绿色
4. 最终输出：一级标题使用模板的字体、字号、颜色 + YAML 的加粗设置

---

#### 推荐的渲染工作流

按照这个顺序执行，效果最好：

**第一步：第一次使用时，先从模板提取样式**

```bash
docubridge extract-styles 公司模板.docx -o 公司样式.yaml
```

**第二步：运行 doctor 检查一切是否就绪**

```bash
docubridge doctor 草稿.md --style 公司样式.yaml --template 公司模板.docx
```

**第三步：按需检查具体元素**

```bash
# 检查标题样式
docubridge style explain 公司样式.yaml heading1 --template 公司模板.docx --pretty

# 检查正文样式
docubridge style explain 公司样式.yaml paragraph --template 公司模板.docx --pretty

# 检查有序列表
docubridge style explain 公司样式.yaml ordered_list --template 公司模板.docx --pretty
```

**第四步：执行正式渲染**

```bash
docubridge render 草稿.md -o 最终版.docx --style 公司样式.yaml --template 公司模板.docx
```

---

#### YAML 中可以控制的元素

| 元素 | 对应 Markdown | 常用属性 |
|------|--------------|----------|
| `heading1` ~ `heading6` | `#` ~ `######` | `font_size`, `bold`, `template_style` |
| `paragraph` | 正文段落 | `font_size`, `first_line_indent_pt`, `space_after_pt` |
| `ordered_list` | `1. 列表项` | `template_style`, `numbering_style` |
| `unordered_list` | `- 列表项` | `template_style` |
| `quote` | `> 引用` | `template_style` |
| `table` | Markdown 表格 | `template_style` |
| `code_block` | 代码块 | `template_style` |

---

#### 🎨 `extract-styles` — 你的秘密武器 ⭐

这是新加入的王炸功能。**从任意 Word 文档提取样式并复用。**

假设你有一份公司模板，里面有你公司标志性的完美格式：

```bash
# 从公司模板提取样式
docubridge extract-styles 公司模板.docx -o 公司样式.yaml

# 现在可以用公司风格渲染任何 Markdown
docubridge render AI草稿.md -o 公司正式文件.docx --style 公司样式.yaml --template 公司模板.docx
```

**它是怎么工作的：**

- 扫描 Word 样式（Heading 1、Normal、List Number 等）
- 自动映射到 Markdown 元素
- 支持中文 Office 样式（`标题 1` → `heading1`，`正文` → `paragraph`）
- 无法识别的样式保存到 `compat.extracted_styles`，供人工检查
- 生成一个 YAML 文件，你可以继续调整和复用

**常用参数：**

| 参数 | 作用 |
|------|------|
| `--pretty` | 输出人类可读的 YAML（有缩进的那种） |
| `--strict` | 如果有任何 Word 样式无法映射，直接报错（给完美主义者用） |
| `--json` | 返回结构化的 JSON 而不是纯文本 |

---

#### 模板问题排查指南

**"我的模板明明有样式，但没生效"**

运行 `style explain` 来调试：
```bash
docubridge style explain 我的样式.yaml heading1 --template 模板.docx --pretty
```

重点看这几个字段：
- `word_style_name`：最终绑定到了哪个样式？
- `source_map`：这个值来自哪里？（`yaml` 优先级高于 `template`）

**"输出结果和模板不完全一致"**

记住 Docubridge 的优先级：
1. YAML 显式值
2. 模板值
3. 默认值

检查是不是 YAML 覆盖了模板的值。

**"我应该先检查哪些元素？"**

按这个顺序来：
1. `heading1` — 最显眼的部分
2. `paragraph` — 正文内容
3. `ordered_list` 或 `unordered_list` — 列表
4. `table` — 表格

---

#### 模板使用常见问题

**Q: YAML 和 Template 必须同时使用吗？**

A: 不是必须的。但同时使用效果最好：
- 不用模板：Docubridge 会动态创建样式，结果可能缺少你想要的字体、编号等细节
- 同时使用：复用你现有的 Word 模板，产出的文档更专业

**Q: 如果模板里没有某个样式会怎样？**

A: 如果 YAML 指定了 `template_style` 但模板里没有，Docubridge 会尝试 fallback 到默认值。结果可能不是你预期的。建议先用 `doctor` 检查。

**Q: 我可以只修改 YAML 不动模板吗？**

A: 可以！这是推荐的工作流：
1. 用 `extract-styles` 从模板提取样式
2. 手动修改生成的 YAML
3. 保留原始模板不动

这样你可以反复调整样式效果，而不需要重新编辑 Word 模板。

---

### 📋 `style` — 样式管理

内置了三个样式：

| 样式名 | 适用场景 |
|--------|----------|
| `academic` | 学术论文、毕业论文 |
| `business` | 报告、备忘录、方案书 |
| `default` | 通用（中文友好默认设置） |

**常用命令：**

```bash
# 列出所有内置样式
docubridge style list

# 查看某个样式的 YAML
docubridge style show academic

# 验证你的自定义样式文件
docubridge style validate 我的样式.yaml

# 解释某个元素会怎么渲染
docubridge style explain 我的样式.yaml heading1 --template 模板.docx

# 合并覆盖值到样式（适合脚本调用）
docubridge style merge 我的样式.yaml --set document.toc.depth=4
```

---

### 🔍 `doctor` — 渲染前的体检

在渲染重要文档之前，先跑一遍诊断：

```bash
docubridge doctor 草稿.md --style academic --template 模板.docx
```

它会检查：

- ✅ 环境是否就绪
- ✅ Markdown 是否可读
- ✅ 样式文件是否有效
- ✅ 样式解析是否正确
- ✅ 模板编号资源是否可用

---

## 真实场景

### 场景一：AI 写的研究摘要

```
你：        "ChatGPT，帮我写一份量子计算的研究摘要。"

ChatGPT：   *输出了漂亮的 Markdown*

你：        *保存到文件：量子摘要.md*

你：        docubridge render 量子摘要.md -o 客户版.docx --style academic

结果：      一份格式规范的 Word 文档，客户可以直接用。✓
```

### 场景二：分析竞争对手的投标文件

```
你：        "把这个 50 页的投标文件转成 Markdown。"

Docubridge：*把 Word 文档解析成干净的 Markdown*

你：        *丢给 AI 分析*

结果：      AI 友好的内容秒到手，不用一个个复制粘贴。✓
```

### 场景三：公司模板的魔法

```
IT部门：     "这是我们的新 Word 模板，所有正式文档必须用这个。"

你：        docubridge extract-styles 公司模板.docx -o corp-style.yaml

你：        docubridge render markdown草稿.md -o 最终稿.docx --style corp-style.yaml --template 公司模板.docx

结果：      AI 生成的文档瞬间变成公司批准格式。IT 刮目相看。✓
```

---

## 下一步做什么

- 🔜 **PDF 解析** — 文字提取、扫描件 OCR 支持、多栏布局恢复
- 🔜 **批量处理** — 一次转换文件夹里的所有文件
- 🔜 **图形界面** — 面向非技术用户的桌面应用
- 🔜 **实时协作** — 云同步和多用户编辑

核心的 `.docx ↔ Markdown` 双向转换已经稳定可用。

---

## 技术细节

**用到的技术：**

- Python 3.12+
- `python-docx` — Word 操作
- `openpyxl` — Excel 解析
- `python-pptx` — PowerPoint 解析
- `markdown-it-py` — Markdown 解析
- `ruamel.yaml` — YAML 处理
- `typer` — CLI 框架

**退出码：**

| 码 | 含义 |
|----|------|
| 0 | 成功 |
| 4 | 样式或模板校验错误 |
| 5 | 执行失败（文件不存在等） |

---

## 更多资源

- 模板使用手册：`docs/superpowers/specs/2026-04-14-docubridge-template-guide-cn.md`
- 样式配置参考：`docs/superpowers/specs/2026-04-14-docubridge-template-reference-cn.md`
- 5 分钟快速上手：`docs/superpowers/specs/2026-04-12-docubridge-5min-quickstart-cn.md`
- 常见问题 FAQ：`docs/superpowers/specs/2026-04-12-docubridge-faq-cn.md`

---

**为那些用 AI 写东西、但最后还是得交 Word 文档的人而造。**

*生活太短，不值得手动调脚注格式。*
