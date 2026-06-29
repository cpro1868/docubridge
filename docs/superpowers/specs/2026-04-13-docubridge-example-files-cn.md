# DocuBridge 示例文件说明

> 日期：2026-04-13  
> 适用对象：第一次使用仓库内示例文件的中文用户

## 1. 这份文档解决什么问题

仓库里已经有几份可以直接拿来试跑的示例文件，但如果只看文件名，不容易知道它们分别适合干什么。

这份文档的目标就是说明：

- 每个示例文件是什么
- 它适合演示哪条能力
- 你应该先看哪个

## 2. 示例文件总览

当前最核心的 4 个示例文件是：

- `tests/fixtures/sample.md`
- `tests/fixtures/style.yaml`
- `tests/fixtures/template-sample.md`
- `tests/fixtures/template-style.yaml`

## 3. `sample.md`

路径：

- `tests/fixtures/sample.md`

适合用途：

- 最小 `render` 演示
- 不带模板的基础 Markdown -> Word 流程

内容特点：

- 标题
- 普通段落
- 行内斜体
- 超链接
- 行内代码
- 图片引用
- 任务列表

适合命令：

```bash
docubridge render tests/fixtures/sample.md -o build/out.docx --style tests/fixtures/style.yaml
```

如果你只是想确认“这个工具能不能把 Markdown 转成 Word”，优先用这个文件。

## 4. `style.yaml`

路径：

- `tests/fixtures/style.yaml`

适合用途：

- 最小样式配置演示
- 不带模板的基础样式解析

内容特点：

- 基础默认字体
- `heading1`
- `paragraph`
- 一个简单的 `document.toc` 配置

适合命令：

```bash
docubridge style explain tests/fixtures/style.yaml heading1 --pretty
docubridge render tests/fixtures/sample.md -o build/out.docx --style tests/fixtures/style.yaml
```

如果你想先理解“最小 YAML 样式配置长什么样”，优先看这个文件。

## 5. `template-sample.md`

路径：

- `tests/fixtures/template-sample.md`

适合用途：

- 带模板的完整结构演示
- 模板场景下的样式绑定验证

内容特点：

- 标题
- 正文
- 有序列表
- 无序列表
- 引用块
- 表格
- 代码块

适合命令：

```bash
docubridge doctor tests/fixtures/template-sample.md --style tests/fixtures/template-style.yaml --template template.docx
docubridge render tests/fixtures/template-sample.md -o build/template-demo.docx --style tests/fixtures/template-style.yaml --template template.docx
```

如果你要学习模板工作流，优先看这个文件。

## 6. `template-style.yaml`

路径：

- `tests/fixtures/template-style.yaml`

适合用途：

- 模板驱动的样式绑定演示
- 学习 Markdown 元素如何映射到 Word 样式

内容特点：

- `heading1 -> Heading 1`
- `paragraph -> Normal`
- `ordered_list -> List Number`
- `unordered_list -> List Bullet`
- `quote -> Quote`
- `table -> Table Grid`
- `code_block -> Quote`

适合命令：

```bash
docubridge style explain tests/fixtures/template-style.yaml ordered_list --template template.docx --pretty
docubridge style explain tests/fixtures/template-style.yaml table --template template.docx --pretty
```

如果你想学“模板 + YAML”到底怎么协同，优先看这个文件。

## 7. 推荐学习顺序

如果你是第一次接触这个项目，建议按下面顺序使用示例：

1. `tests/fixtures/sample.md`
2. `tests/fixtures/style.yaml`
3. `tests/fixtures/template-sample.md`
4. `tests/fixtures/template-style.yaml`

原因是：

- 先从不带模板的最小路径开始
- 再理解基础样式配置
- 最后再进入模板协同场景

## 8. 推荐搭配阅读

建议和这些文档配合看：

- `README_CN.md`
- `docs/superpowers/specs/2026-04-12-docubridge-5min-quickstart-cn.md`
- `docs/superpowers/specs/2026-04-12-docubridge-first-run-cn.md`
- `docs/superpowers/specs/2026-04-12-docubridge-faq-cn.md`
