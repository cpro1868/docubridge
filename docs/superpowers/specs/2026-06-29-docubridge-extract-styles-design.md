# DocuBridge Word 样式提取设计

> 日期：2026-06-29
> 状态：设计确认稿
> 关联文档：`docs/requirements.md`、`docs/superpowers/specs/2026-04-07-docubridge-render-design.md`

## 1. 背景与目标

### 1.1 需求

新增 CLI 能力：将 Word 文档（`.docx`）中的样式提取出来，生成一份 `style.yaml` 配置文件。用户后续可用该 YAML 文件配合 `--template` 或独立地规范 Markdown → Word 的渲染样式。

### 1.2 目标

- 降低用户手写 `style.yaml` 的门槛
- 复用现有模板读取能力，不引入新的 Word 解析依赖
- 生成的 YAML 必须能被 `docubridge style validate` 校验通过
- 生成的 YAML 应能直接用于 `docubridge render ... --style output.yaml --template input.docx`

## 2. CLI 设计

### 2.1 命令

作为顶层子命令，便于与 `style` 子命令区分：

```bash
docubridge extract-styles <input.docx> -o <output.yaml>
```

### 2.2 参数

| 参数 | 必填 | 说明 |
|---|---|---|
| `input.docx` | 是 | 源 Word 文档 |
| `-o, --output` | 是 | 输出 YAML 路径 |
| `--pretty` | 否 | 输出带缩进的可读 YAML |
| `--strict` | 否 | 存在无法映射的样式时直接报错退出 |
| `--json` | 否 | 输出结构化 JSON 结果 |

### 2.3 使用示例

```bash
# 基础提取
docubridge extract-styles contract-template.docx -o contract-style.yaml

# 可读格式
docubridge extract-styles contract-template.docx -o contract-style.yaml --pretty

# 严格模式
docubridge extract-styles contract-template.docx -o contract-style.yaml --strict

# 后续用于渲染（--template 会自动清空原文正文，只保留样式、编号和页面设置）
docubridge render report.md -o report.docx --style contract-style.yaml --template contract-template.docx
```

## 3. 数据流

```
input.docx
  → template_bridge.load_template_view()
      读取 Word 样式、文档默认属性、编号绑定
  → style_extractor.extract_style_profile()
      映射样式名、过滤属性、构建 StyleProfile
  → yaml_adapter.write_yaml()
      写出 output.yaml
```

## 4. 样式映射规则

### 4.1 精确约定映射

Word 样式名与 docubridge 元素名的精确对应，匹配时忽略大小写与首尾空格：

| Word 样式名 | docubridge 元素 |
|---|---|
| `Heading 1` ~ `Heading 6`（中文 Office 常见别名：`标题 1` ~ `标题 6`） | `heading1` ~ `heading6` |
| `Normal`（中文 Office 常见别名：`正文`） | `paragraph` |
| `List Number` | `ordered_list` |
| `List Bullet` | `unordered_list` |
| `Quote` / `Block Quote` / `Blockquote` | `quote` |
| `Table Grid` / `Light Grid` / `Medium Grid` / `Dark Grid` | `table` |
| `Code` / `Preformatted` | `code_block` |

### 4.2 自定义样式相似度映射

对未命中精确映射的样式名，使用 `difflib.SequenceMatcher` 与候选集比较：

- 元素名候选：`heading1` ~ `heading6`、`paragraph`、`ordered_list`、`unordered_list`、`quote`、`table`、`code_block`
- 常见别名候选：`Title`、`Subtitle`、`Body Text`、`Bullet`、`Numbering`、`Caption`

相似度阈值：**0.6**。取最高分的候选；同分时优先精确别名，其次元素名。

示例：

- `My Heading 1` → `heading1`
- `Item Bullet` → `unordered_list`
- `CodeText` → `code_block`

### 4.3 未映射样式

无法映射的样式保留在 `compat.extracted_styles` 段：

```yaml
compat:
  extracted_styles:
    CustomTitle:
      font_name: 楷体
      font_size: 22
```

并附带 YAML 注释说明该样式未自动映射，需人工确认。

### 4.4 映射范围

v1 仅按样式名做静态映射，不分析文档中样式的实际使用场景。

## 5. 提取的属性范围

### 5.1 文档默认属性

从 Word 文档默认设置读取，写入 YAML 的 `defaults` 段：

- `font_name` / `font_ascii` / `font_hansi` / `font_east_asia` / `font_cs`
- `font_size`
- `first_line_indent_pt`
- `left_indent_pt` / `right_indent_pt`
- `space_before_pt` / `space_after_pt`

### 5.2 元素级属性

对每个映射到的元素，输出 `template_style` 和从 Word 样式读取的显式属性：

```yaml
elements:
  heading1:
    template_style: Heading 1
    font_name: 黑体
    font_size: 18
    bold: true
```

### 5.3 属性过滤

- 如果元素属性与 `defaults` 中的对应属性相同，则不重复输出
- 不输出 `false` 或空值属性
- `alignment` 整数值转换为可读字符串：`left`、`center`、`right`、`justify`
- 编号绑定自动推导为 `numbering_style`

## 6. 输出 YAML 结构示例

```yaml
meta:
  name: extracted-from-contract-template
  version: 1
  source: contract-template.docx

defaults:
  font_name: 宋体
  font_size: 12

elements:
  heading1:
    template_style: Heading 1
    font_name: 黑体
    font_size: 18
    bold: true
  heading2:
    template_style: Heading 2
    font_name: 黑体
    font_size: 16
    bold: true
  paragraph:
    template_style: 正文
    font_name: 宋体
  ordered_list:
    template_style: List Number
  unordered_list:
    template_style: List Bullet
  quote:
    template_style: Quote
  table:
    template_style: Table Grid

compat:
  extracted_styles:
    CustomTitle:
      font_name: 楷体
      font_size: 22
      # 未自动映射到已知 Markdown 元素，请人工确认
```

## 7. 错误处理与退出码

沿用项目现有约定：

| 场景 | 退出码 | 诊断码 |
|---|---|---|
| 成功 | `0` | - |
| 输入文件不存在 | `5` | `INPUT_FILE_NOT_FOUND` |
| 输入不是有效 `.docx` | `5` | `INPUT_IO_ERROR` |
| 输出路径不可写 | `5` | `OUTPUT_IO_ERROR` |
| `--strict` 模式下存在未映射样式 | `4` | `STYLE_EXTRACTION_ERROR` |

### 7.1 默认行为下的未映射样式

非严格模式下，未映射样式写入 `compat.extracted_styles`，命令仍成功退出，并在 stderr 打印提示：

```
extracted 12 styles, 2 unmapped styles saved to compat.extracted_styles
```

### 7.2 JSON 输出

支持 `--json`：

成功：

```json
{
  "success": true,
  "output_path": "output.yaml",
  "diagnostics": []
}
```

失败：

```json
{
  "success": false,
  "output_path": "output.yaml",
  "diagnostics": [
    {"code": "STYLE_EXTRACTION_ERROR", "message": "Unmapped styles: CustomTitle, AnotherStyle"}
  ]
}
```

## 8. 模块设计

### 8.1 新增文件

- `src/docubridge/core/style_extractor.py`
  - `extract_style_profile(input_path: Path) -> StyleProfile`
  - `_map_word_styles(styles, document_defaults) -> tuple[dict, dict]`
  - `_normalize_alignment(value)`
  - `_deduplicate_properties(element_props, defaults)`
  - `_compute_similarity(name, candidates)`

- `src/docubridge/adapters/yaml_adapter.py` 增加写入能力
  - `write_yaml(path: Path, data: dict, *, pretty: bool = False)`

- `tests/test_style_extractor.py`
  - 映射逻辑单元测试
  - CLI 命令集成测试

### 8.2 修改文件

- `src/docubridge/cli.py`
  - 新增 `extract_styles` 顶层命令
  - 复用现有 `_emit_json`、`_display_path` 等辅助函数

### 8.3 复用能力

- `template_bridge.load_template_view()`：读取 Word 样式和文档默认属性
- `style_schema.StyleProfile`：构建 YAML 内存对象
- `yaml_adapter`：YAML 读写

## 9. 测试策略

### 9.1 单元测试

- 精确映射：`Heading 1` → `heading1`，`Normal` → `paragraph`
- 相似度映射：`MyBullet` → `unordered_list`
- 未映射样式进入 `compat.extracted_styles`
- `--strict` 下未映射样式报错
- 属性过滤：与 defaults 相同的属性不重复输出
- `alignment` 整数 → 字符串转换

### 9.2 CLI 测试

- 命令能生成 YAML 文件
- `--pretty` 输出带缩进
- `--json` 成功/失败输出正确
- 无效输入返回退出码 5
- 生成的 YAML 可被 `docubridge style validate` 通过

### 9.3 集成测试

- 用生成的 YAML + 原 Word 文件作为 `--template` 渲染 Markdown，验证样式绑定成功

## 10. 边界与限制

- v1 仅按样式名做静态映射，不分析文档中样式的实际使用位置
- 字符样式、表格样式的高级属性不在 v1 提取范围内
- 多语言字体槽位依赖 `template_bridge` 已有的读取能力
- 未映射样式需用户人工复核后决定是否移动到 `elements`
- v1 明确支持 `.docx`；`.dotx` 模板文件可能在后续版本专门支持
- `--template` 加载模板时会清空 body 正文（段落、表格、图片），仅保留样式、编号、文档默认属性、节属性和页眉页脚

## 11. 后续可扩展

- 支持 `--profile` 指定输出基准模板，在基准上增量提取
- 分析文档实际使用场景，做更智能的语义映射
- 提取编号定义并生成 `multilevel_list` 配置
- 支持 `.dotx` 模板文件的提取
