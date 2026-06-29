# DocuBridge 5 分钟快速上手

> 日期：2026-04-12  
> 适合对象：第一次接触 DocuBridge 的中文用户

## 1. 先记住它能做什么

当前最重要的两件事：

- 把 `Word (.docx)` 转成 `Markdown`
- 把 `Markdown` 转成 `Word (.docx)`

如果你只想先把这两个功能用起来，这份文档就够了。

## 2. 先记住 4 个命令

- `parse`：文档转 Markdown
- `render`：Markdown 转 Word
- `doctor`：先检查输入和样式是否可用
- `style explain`：看样式最终怎么解析

## 3. 最快开始的方法

### 3.1 Word 转 Markdown

```bash
docubridge parse input.docx -o out.md
```

### 3.2 Markdown 转 Word

```bash
docubridge render input.md -o out.docx --style style.yaml
```

### 3.3 使用 Word 模板

```bash
docubridge render input.md -o out.docx --style style.yaml --template template.docx
```

## 4. 如果你要用模板，先这样做

不要一上来就直接 `render`，先按这个顺序：

```bash
docubridge doctor input.md --style style.yaml --template template.docx
docubridge style explain style.yaml heading1 --template template.docx --pretty
docubridge render input.md -o out.docx --style style.yaml --template template.docx
```

这样做的好处：

- `doctor` 先帮你发现文件或配置问题
- `style explain` 帮你确认样式绑定是否正确
- `render` 最后再正式输出

## 5. 最值得先看的示例

直接看这两个文件：

- `tests/fixtures/template-sample.md`
- `tests/fixtures/template-style.yaml`

它们展示了最常见的结构：

- 标题
- 正文
- 有序列表
- 无序列表
- 引用
- 表格
- 代码块

## 6. 你只需要理解这一句话

- Markdown 决定内容结构
- YAML 决定样式意图
- Template 提供 Word 里的宿主样式

## 7. 出问题先怎么查

### 模板找不到

先跑：

```bash
docubridge doctor input.md --style style.yaml --template template.docx
```

### 不知道样式为什么没生效

跑：

```bash
docubridge style explain style.yaml paragraph --template template.docx --pretty
```

重点看：

- `word_style_name`
- `source_map`

## 8. 现在不要过度期待什么

当前版本先别按这些目标理解：

- 完全复刻复杂 Word 模板的高级编号和分页
- GUI 图形界面
- 把 `.xlsx/.pptx/.pdf` 当成首发主能力

## 9. 下一步看什么

如果你已经能跟着这份文档跑起来，再继续看：

1. `README_CN.md`
2. `docs/superpowers/specs/2026-04-12-docubridge-user-guide-cn.md`
3. `docs/superpowers/specs/2026-04-11-docubridge-v1-known-issues.md`
