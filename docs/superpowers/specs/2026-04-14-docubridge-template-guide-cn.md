# DocuBridge 模板使用手册

> 日期：2026-04-14  
> 适用对象：普通使用者  
> 适用范围：`Markdown -> Word (.docx)` 模板定制

## 1. 这份手册解决什么问题

如果你已经有公司模板、学校模板、项目模板，或者你希望把 Markdown 输出成更像“正式 Word 文档”的结果，这份手册就是给你用的。

它重点回答 4 个问题：

- 模板文件到底是什么
- `--template` 和 `--style` 分别负责什么
- 普通用户应该怎么准备模板
- 输出结果不符合预期时怎么排查

这不是开发设计文档，也不是字段字典。  
这是一份面向普通使用者的操作手册。

## 2. 先记住两个概念

### 2.1 `--template`

`--template` 负责提供一个宿主 `.docx` 文档。

你可以把它理解成：

- 这个 Word 模板里已经有一套现成的 Word 样式
- DocuBridge 渲染时会尽量复用这些样式
- 输出结果会更接近你现有文档体系

模板文件通常是：

- 公司公文模板
- 学校论文模板
- 项目报告模板
- 你自己整理好的 Word 样式模板

### 2.2 `--style`

`--style` 负责提供 YAML 样式配置。

你可以把它理解成：

- 模板负责“Word 里有什么样式可用”
- YAML 负责“Markdown 的哪些结构要绑定到哪些样式”

简单说：

- `template` 提供样式资源
- `style` 指定绑定规则

## 3. `--template` 和 `--style` 的关系

普通使用者最容易混淆这一点。

建议这样理解：

1. Markdown 决定内容结构
2. YAML 决定结构映射关系
3. Template 决定最终 Word 中可复用的样式宿主

例如：

- Markdown 里的 `# 标题` 是“一级标题结构”
- YAML 里 `heading1.template_style: Heading 1` 表示它要绑定到 Word 的 `Heading 1`
- 模板文件里如果真的存在 `Heading 1`，最终输出就会复用模板中的一级标题样式

## 4. 优先级怎么理解

普通使用中，按下面顺序理解就够了：

1. 先看 YAML 有没有显式写这个属性
2. 如果 YAML 没写，再看模板里对应 Word 样式是否提供了这个属性
3. 如果两边都没写，再回退到默认值

结论：

- YAML 的显式配置优先级最高
- 模板提供宿主样式
- 默认值最后兜底

所以如果你发现“模板里明明有这个设置，但输出没按模板来”，先检查 YAML 里是不是把它覆盖掉了。

## 5. 你需要准备什么

要使用模板功能，通常需要准备 3 个文件：

1. 一个 Markdown 文件
2. 一个 YAML 样式文件
3. 一个 `.docx` 模板文件

最常见的命令是：

```bash
docubridge render input.md -o out.docx --style style.yaml --template corp-template.docx
```

建议第一次使用时，不要直接拿复杂模板上手。  
先用一个你能控制的简单模板验证流程，再换成正式模板。

## 6. 普通用户的推荐操作流程

### 6.1 第一步：先确认模板文件是正常的 Word 文档

不要直接拿损坏文件、只读锁定文件、导出异常文件测试。

你应该先确认：

- 模板可以在 Word 里正常打开
- 模板里的样式已经整理过
- 你确实希望复用这些 Word 样式

### 6.2 第二步：先运行 `doctor`

```bash
docubridge doctor input.md --style style.yaml --template corp-template.docx
```

这一步的作用是：

- 检查 Markdown 是否可读
- 检查 YAML 是否可读
- 检查模板文件是否存在
- 检查样式绑定是否能解析

普通用户不要跳过这一步。  
直接 `render` 出错时，通常不如先 `doctor` 更容易定位问题。

### 6.3 第三步：用 `style explain` 看单个元素

```bash
docubridge style explain style.yaml heading1 --template corp-template.docx --pretty
```

这一步适合确认：

- 某个元素最终绑定到了哪个 Word 样式
- 某个属性来自 YAML、模板还是默认值

如果你不确定“为什么标题没按模板走”，先查 `heading1`。  
如果你不确定“为什么列表不对”，先查 `ordered_list` 或 `unordered_list`。

### 6.4 第四步：再运行 `render`

```bash
docubridge render input.md -o out.docx --style style.yaml --template corp-template.docx
```

这一步才是真正输出 Word 文件。

## 7. 普通用户怎么准备模板文件

模板文件不需要写任何程序代码。  
它本质上就是一个整理过样式的 `.docx` 文档。

建议按这个思路准备：

- 在 Word 中创建一个空白文档
- 把你常用的标题、正文、引用、列表、表格样式整理好
- 保存为专用模板文件

你至少应该关注这些 Word 样式：

- `Heading 1`
- `Heading 2`
- `Heading 3`
- `Normal`
- `Quote`
- `List Number`
- `List Bullet`
- `Table Grid` 或你自己的表格样式

如果模板里没有这些样式，DocuBridge 仍然可能输出成功，但结果会更接近默认样式，而不是你想要的模板效果。

## 8. 最常见的模板定制场景

### 8.1 只想控制标题和正文

这是最简单、最稳的场景。

你只需要确保：

- 模板里有 `Heading 1`、`Heading 2`、`Normal`
- YAML 中把 `heading1`、`heading2`、`paragraph` 绑定到这些样式

适合：

- 报告
- 说明文档
- 方案文档

### 8.2 想控制列表样式

你可以让 Markdown 列表绑定到 Word 的：

- `List Number`
- `List Bullet`

但当前版本要注意：

- 现在列表更偏“段落样式 + 文本前缀”
- 不是完整的 Word 原生多级编号恢复系统

所以如果你期待的是非常复杂的企业多级编号模板，当前版本还不算完全覆盖。

### 8.3 想控制表格样式

你可以在 YAML 里把 `table` 绑定到某个 Word 表格样式，例如：

- `Table Grid`
- 你自定义的表格样式

适合：

- 数据表
- 对照表
- 项目清单

### 8.4 想控制代码块和引用块

当前版本里：

- `quote` 可以绑定引用类段落样式
- `code_block` 也可以绑定一类适合代码展示的样式

但要注意：

- 代码块目前渲染为单列表格
- 所以它的最终视觉效果不只取决于 YAML，还取决于单元格内部段落样式

## 9. 最小模板工作流示例

仓库里已经有最小示例文件：

- `tests/fixtures/template-sample.md`
- `tests/fixtures/template-style.yaml`

推荐命令：

```bash
docubridge doctor tests/fixtures/template-sample.md --style tests/fixtures/template-style.yaml --template corp-template.docx
docubridge style explain tests/fixtures/template-style.yaml ordered_list --template corp-template.docx --pretty
docubridge render tests/fixtures/template-sample.md -o build/template-demo.docx --style tests/fixtures/template-style.yaml --template corp-template.docx
```

这个流程适合第一次上手模板功能。

## 10. 结果不符合预期时怎么排查

### 10.1 模板文件路径是不是对的

先检查：

- 文件是否真的存在
- 扩展名是不是 `.docx`
- 文件是否能正常打开

### 10.2 先查 `doctor`

如果 `doctor` 已经报错，不要直接继续 `render`。

### 10.3 再查 `style explain`

这是最有用的排查入口之一。

看这几个点：

- `word_style_name`
- `resolved_properties`
- `source_map`

如果 `source_map` 显示某个字段来自 `yaml`，那就说明它不是模板问题，而是 YAML 覆盖了模板。

### 10.4 不要一次排查所有元素

普通用户排查模板时，建议一次只看一个元素：

- 先看 `heading1`
- 再看 `paragraph`
- 再看 `ordered_list`
- 再看 `table`

这样最容易定位。

## 11. 当前版本哪些需求不要期待过高

当前模板功能已经能支持很多常见输出场景，但还不适合期待这些效果：

- 复杂节设置自动恢复
- 页眉页脚完整模板还原
- 目录自动高级恢复
- 复杂分页控制
- 高级 Word 多级编号定义完全复用
- 非常复杂的企业模板布局完全一比一还原

如果你的目标是“正式报告基本能用”，当前版本是可行的。  
如果你的目标是“高度复杂企业模板的精确版式恢复”，当前版本还需要继续开发。

## 12. 普通用户应该先看什么

建议阅读顺序：

1. `README_CN.md`
2. 本文档
3. `tests/fixtures/template-sample.md`
4. `tests/fixtures/template-style.yaml`
5. `style explain` 的实际输出

## 13. 下一步看哪里

如果你想继续深入：

- 看模板字段和规则细节：`2026-04-14-docubridge-template-reference-cn.md`
- 看完整中文使用手册：`2026-04-12-docubridge-user-guide-cn.md`
- 看示例文件说明：`2026-04-13-docubridge-example-files-cn.md`
