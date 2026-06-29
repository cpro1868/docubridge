# DocuBridge 非技术用户操作手册

> 日期：2026-04-12  
> 适用对象：不想先理解技术细节、只想把工具用起来的用户

## 1. 这份手册适合谁

如果你符合下面任意一种情况，这份手册适合你：

- 你经常收到 Word 文档，想把它转成 Markdown 给 AI 使用
- 你已经有 Markdown 内容，想快速生成 Word 文档
- 你有公司模板、学校模板或固定 Word 模板，希望直接套进去
- 你不想先看设计文档和实现细节

## 2. 你只需要知道 3 件事

### 2.1 这个工具能做什么

当前最重要的两项能力：

- `Word (.docx) -> Markdown`
- `Markdown -> Word (.docx)`

### 2.2 你最常用的 3 个命令

- `parse`：把 Word 转成 Markdown
- `render`：把 Markdown 转成 Word
- `doctor`：先检查输入、样式和模板有没有问题

### 2.3 如果你要套模板

记住：

- `--style`：告诉工具“我想要什么样式”
- `--template`：告诉工具“Word 模板里有哪些可用样式”

## 3. 场景一：把 Word 转成 Markdown

### 你要做什么

把 `.docx` 文档转成 Markdown，后续喂给 AI、RAG 或知识库。

### 直接这样做

```bash
docubridge parse input.docx -o out.md
```

### 当前版本能保留什么

- 标题
- 段落
- 列表
- 表格
- 图片引用
- 基础粗体、斜体、删除线、链接、行内代码

### 你要有的预期

它的目标是“结构正确、便于后续处理”，不是“完全保留 Word 所有版式细节”。

## 4. 场景二：把 Markdown 转成 Word

### 你要做什么

把 Markdown 内容转回 `.docx`，方便交付、打印、归档。

### 最简单的命令

```bash
docubridge render input.md -o out.docx --style style.yaml
```

### 适合什么时候用

- 你没有现成 Word 模板
- 你只需要基础样式控制
- 你希望先快速出一个结果

## 5. 场景三：把 Markdown 套进现有 Word 模板

### 你要做什么

你已经有企业模板、学校模板或固定格式模板，希望最终输出尽量贴近它。

### 不要直接 `render`

建议按这个顺序：

```bash
docubridge doctor input.md --style style.yaml --template template.docx
docubridge style explain style.yaml heading1 --template template.docx --pretty
docubridge render input.md -o out.docx --style style.yaml --template template.docx
```

### 为什么这样做

- `doctor` 先检查模板文件、样式文件和输入文件
- `style explain` 让你知道某个元素最终绑定到了哪个 Word 样式
- `render` 最后再真正生成输出文件

## 6. 如果你完全不知道从哪开始

先看并照着用这两个示例文件：

- `tests/fixtures/template-sample.md`
- `tests/fixtures/template-style.yaml`

这两个文件已经涵盖最常见的内容：

- 标题
- 正文
- 有序列表
- 无序列表
- 引用块
- 表格
- 代码块

## 7. 最常见的两个问题

### 7.1 模板文件找不到

先跑：

```bash
docubridge doctor input.md --style style.yaml --template template.docx
```

如果这里失败，不要继续跑 `render`。

### 7.2 生成出来的样式不对

跑：

```bash
docubridge style explain style.yaml paragraph --template template.docx --pretty
```

重点看：

- `word_style_name`
- `source_map`

如果 `source_map` 里显示属性来自 `template`，说明样式主要来自模板。  
如果显示来自 `yaml`，说明是 YAML 显式覆盖了模板。

## 8. 当前版本不要误解成什么

当前版本不要按这些目标来期待：

- 完全复刻复杂 Word 模板中的高级编号、分页、目录、节控制
- 图形界面产品
- 把 `.xlsx/.pptx/.pdf` 当成首发主能力

当前版本更适合：

- 本地把 Word 结构化成 Markdown
- 把 Markdown 套进已有 Word 样式体系
- 做 beta 试用和内部使用

## 9. 推荐阅读顺序

如果你是普通用户，建议按这个顺序：

1. `README_CN.md`
2. `docs/superpowers/specs/2026-04-12-docubridge-5min-quickstart-cn.md`
3. 本文档
4. `docs/superpowers/specs/2026-04-11-docubridge-v1-known-issues.md`

如果你看完还想更系统地理解，再看：

- `docs/superpowers/specs/2026-04-12-docubridge-user-guide-cn.md`
