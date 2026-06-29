# DocuBridge 模板参考附录

> 日期：2026-04-14  
> 适用对象：需要查规则细节的普通使用者、模板实施人员

## 1. 这份附录是做什么的

这份文档不是入门手册。  
它用于回答“具体规则是什么”。

如果你还没看过主手册，建议先看：

- `2026-04-14-docubridge-template-guide-cn.md`

## 2. `--template` 和 `--style` 的职责边界

### 2.1 `--template`

作用：

- 提供宿主 `.docx`
- 提供可复用 Word 样式
- 提供模板中的已有样式名称和基础外观

不应期待：

- 自动恢复复杂节
- 自动恢复完整分页逻辑
- 自动恢复高级页眉页脚体系
- 自动恢复复杂多级编号定义

### 2.2 `--style`

作用：

- 指定 Markdown 元素如何映射到 Word 样式
- 指定显式样式属性
- 指定文档输出意图

典型内容：

- `heading1`
- `paragraph`
- `ordered_list`
- `unordered_list`
- `quote`
- `table`
- `code_block`

## 3. 优先级规则

普通情况下可按下列顺序理解：

1. YAML 显式字段
2. 模板中的对应 Word 样式属性
3. 默认值

简化结论：

- YAML 明确写了，就优先按 YAML
- YAML 没写，才更多依赖模板
- 两边都没有，再走默认值

## 4. 常见元素映射

当前常见元素包括：

- `heading1`
- `heading2`
- `heading3`
- `paragraph`
- `ordered_list`
- `unordered_list`
- `quote`
- `table`
- `code_block`

常见映射示例：

- `heading1 -> Heading 1`
- `heading2 -> Heading 2`
- `paragraph -> Normal`
- `ordered_list -> List Number`
- `unordered_list -> List Bullet`
- `quote -> Quote`
- `table -> Table Grid`

## 5. 一个最小 YAML 模板配置示例

```yaml
meta:
  name: template-demo
  version: 1

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
  ordered_list:
    template_style: List Number
  unordered_list:
    template_style: List Bullet
  quote:
    template_style: Quote
  table:
    template_style: Table Grid
  code_block:
    template_style: Quote
```

## 6. 你最常看到的字段

### 6.1 `template_style`

作用：

- 指定这个 Markdown 元素最终绑定到哪个 Word 样式

示例：

```yaml
heading1:
  template_style: Heading 1
```

### 6.2 `font_name`

作用：

- 指定字体名称

### 6.3 `font_size`

作用：

- 指定字号

### 6.4 `bold`

作用：

- 指定是否加粗

### 6.5 `defaults`

作用：

- 为未显式配置的元素提供默认值

## 7. 模板样式命名建议

如果你自己制作模板，建议优先复用 Word 的常见样式名：

- `Heading 1`
- `Heading 2`
- `Heading 3`
- `Normal`
- `Quote`
- `List Number`
- `List Bullet`
- `Table Grid`

原因：

- 更容易对照
- 更方便排查
- 对普通用户最友好

如果必须用企业自定义样式名，也可以，但建议命名清晰，不要让人猜。

## 8. 当前模板功能对哪些结构有效

### 8.1 有效的部分

- 标题
- 正文
- 引用块
- 列表
- 表格
- 代码块

### 8.2 部分有效但不要过度期待的部分

- 列表的复杂多级编号
- 代码块的高级视觉样式
- 非常复杂的 Word 表格外观

## 9. 当前版本的已知边界

当前版本模板协同仍然主要聚焦：

- Word 样式复用
- 宿主文档默认样式
- YAML 样式解析

尚不属于当前强项的部分：

- 高级编号定义恢复
- 节布局
- 页眉页脚
- 目录系统
- 复杂分页
- 高度复杂企业模板一比一版式还原

## 10. `style explain` 怎么看

建议优先看下面几个字段：

- `element_name`
- `word_style_name`
- `resolved_properties`
- `source_map`

### 10.1 `word_style_name`

表示最终绑定到哪个 Word 样式。

### 10.2 `resolved_properties`

表示最终生效的属性值。

### 10.3 `source_map`

表示这些属性来自哪里，例如：

- `yaml`
- `template`
- `defaults`

这个字段对排查最重要。

## 11. 常见问题定位表

### 11.1 “模板明明有样式，但结果没生效”

优先检查：

- YAML 是否显式覆盖了模板
- `template_style` 是否写对
- 模板里是否真的存在这个 Word 样式

### 11.2 “标题看起来不对”

优先检查：

- `heading1` / `heading2` 是否绑定对了
- 模板中的 `Heading 1` / `Heading 2` 是否整理过

### 11.3 “列表不像我预期的 Word 编号”

优先检查：

- 当前版本是否只是套了列表段落样式
- 是否误以为已经支持完整原生多级编号恢复

### 11.4 “代码块样式不理想”

优先检查：

- 当前代码块是否作为单列表格渲染
- 单元格内部段落样式是否合适

## 12. 推荐命令

### 12.1 预检查

```bash
docubridge doctor input.md --style style.yaml --template template.docx
```

### 12.2 查看单个元素

```bash
docubridge style explain style.yaml heading1 --template template.docx --pretty
```

### 12.3 正式渲染

```bash
docubridge render input.md -o out.docx --style style.yaml --template template.docx
```

## 13. 相关文档

- 模板主手册：`2026-04-14-docubridge-template-guide-cn.md`
- 中文 README：`README_CN.md`
- 中文安装与使用手册：`2026-04-12-docubridge-user-guide-cn.md`
- 示例文件说明：`2026-04-13-docubridge-example-files-cn.md`
