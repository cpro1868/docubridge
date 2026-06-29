# DocuBridge 中文安装与使用手册

> 日期：2026-04-12  
> 适用版本：`v1.0 beta`

## 1. 这是什么

DocuBridge 是一个本地运行的文档处理工具，当前最核心的两条能力是：

- `Word (.docx) -> Markdown`
- `Markdown -> Word (.docx)`

如果你主要想做这两件事，这个版本已经可以进入 beta 使用。

## 2. 你应该先知道什么

先记住这四个命令：

- `parse`：把文档转成 Markdown
- `render`：把 Markdown 转成 Word
- `doctor`：先检查输入、样式和模板是否可用
- `style explain`：查看某个样式元素最终是怎么解析出来的

再记住这两个参数：

- `--style`：YAML 样式配置，负责“我想要什么样式”
- `--template`：Word 模板，负责“Word 里有哪些宿主样式可用”

## 3. 安装前提

建议环境：

- Windows、macOS 或 Linux
- Python 3.12+

如果你只是使用当前仓库中的代码，至少需要能在终端里运行：

```bash
python --version
```

## 4. 最小使用路径

### 4.1 把 Word 转成 Markdown

```bash
docubridge parse input.docx -o out.md
```

适合场景：

- 把 Word 文档转成 Markdown，交给 AI / RAG / 知识库
- 做结构化内容提取

### 4.2 把 Markdown 转成 Word

```bash
docubridge render input.md -o out.docx --style style.yaml
```

适合场景：

- 把 AI 生成的 Markdown 报告转回 Word
- 根据样式配置生成规范文档

### 4.3 使用模板渲染

```bash
docubridge render input.md -o out.docx --style style.yaml --template corp-template.docx
```

适合场景：

- 你已经有企业模板、学校模板或固定 Word 模板
- 你希望输出文档尽量贴合现有 Word 样式体系

## 5. 推荐工作流

最稳妥的顺序不是直接 `render`，而是：

```bash
docubridge doctor input.md --style style.yaml --template template.docx
docubridge style explain style.yaml heading1 --template template.docx --pretty
docubridge render input.md -o out.docx --style style.yaml --template template.docx
```

为什么这样做：

- `doctor` 先帮你发现输入文件、样式文件、模板文件的问题
- `style explain` 帮你看清某个元素最后绑定到了哪个 Word 样式
- `render` 最后再真正输出

## 6. 建议先看的示例文件

如果你完全是第一次使用，建议先看仓库里的这两个文件：

- `tests/fixtures/template-sample.md`
- `tests/fixtures/template-style.yaml`

它们展示的是最小模板场景：

- 标题
- 正文
- 有序列表
- 无序列表
- 引用块
- 表格
- 代码块

## 7. 你最需要理解的映射关系

Markdown 结构决定“内容是什么”，YAML 决定“这些内容用什么样式输出”。

常见映射如下：

- `heading1`：一级标题
- `paragraph`：正文段落
- `ordered_list`：有序列表
- `unordered_list`：无序列表
- `quote`：引用块
- `table`：表格
- `code_block`：代码块

如果你在模板里已经有这些 Word 样式名，例如：

- `Heading 1`
- `Normal`
- `List Number`
- `List Bullet`
- `Quote`
- `Table Grid`

那么 YAML 可以把这些 Markdown 元素绑定到对应 Word 样式。

## 8. 常见排错方法

### 8.1 模板文件找不到

先跑：

```bash
docubridge doctor input.md --style style.yaml --template template.docx
```

如果这里失败，不要先跑 `render`。

### 8.2 不知道某个样式为什么没生效

跑：

```bash
docubridge style explain style.yaml paragraph --template template.docx --pretty
```

重点看：

- `word_style_name`
- `resolved_properties`
- `source_map`

`source_map` 会告诉你，这个属性来自：

- `yaml`
- `template`
- `defaults`

### 8.3 输出结果跟复杂企业模板不完全一致

先看已知边界文档：

- `docs/superpowers/specs/2026-04-11-docubridge-v1-known-issues.md`

当前版本的重点是“结构正确 + 样式可控”，不是“100% 复刻复杂 Word 模板的全部高级版式能力”。

## 9. 当前版本最适合谁

当前版本适合：

- 需要把常见 `.docx` 转成 Markdown 的用户
- 需要把 Markdown 套进现有 Word 样式体系的用户
- 需要在本地完成文档转换，不想上传文档的用户

当前版本不适合高承诺场景：

- 复杂页眉页脚、目录、节、分页高度依赖的 Word 模板复刻
- 把 `.xlsx/.pptx/.pdf` 当成首发主能力
- 希望现在就有 GUI 的用户

## 10. 你下一步应该看什么

如果你是普通使用者，按这个顺序：

1. `README_CN.md`
2. 本文档
3. `tests/fixtures/template-sample.md`
4. `tests/fixtures/template-style.yaml`
5. `docs/superpowers/specs/2026-04-11-docubridge-v1-known-issues.md`

如果你是负责人或产品同学，再补看：

- `docs/superpowers/specs/2026-04-11-docubridge-v1-summary.md`
- `docs/superpowers/specs/2026-04-11-docubridge-v1-release-checklist.md`

如果你是开发者，再补看：

- `docs/requirements.md`
- `docs/superpowers/specs/2026-04-07-docubridge-render-design.md`
