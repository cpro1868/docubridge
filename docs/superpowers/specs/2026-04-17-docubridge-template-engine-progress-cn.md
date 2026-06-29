# DocuBridge 模板排版引擎阶段进展记录

日期：2026-04-17

## 记录目的

本文用于记录 `Markdown -> Word (.docx)` 模板排版引擎升级的当前进展，方便后续继续开发、验收和发布收口。

当前主线目标不是继续做简单的 Markdown 渲染，而是把 `Markdown -> Word` 升级为混合模式模板排版引擎：

- Word 模板提供默认样式、段落样式、字符样式、字体槽位和编号资源
- YAML 提供 Markdown 元素到 Word 样式的映射，以及局部显式覆盖
- 渲染器最终要真实应用 Word 样式、段落属性、字体槽位和原生编号

## 当前阶段结论

第一阶段骨架已经贯通，已从“轻量样式映射器”推进到“可解释的模板排版引擎雏形”。

已经具备的能力：

- 模板元数据读取：可读取 Word 样式、文档默认值、段落属性、字体槽位和样式关联编号资源
- 样式合并：支持 `defaults -> template_document_default -> template_style -> yaml` 的优先级
- 中间模型：已引入 `layout_intent`，渲染前会先生成段落排版意图
- 字体槽位：支持区分 `font_ascii`、`font_hansi`、`font_east_asia`、`font_cs`
- 段落属性：已支持首行缩进、左右缩进、段前段后、对齐等第一阶段字段
- 原生编号：有序列表和普通无序列表在模板提供对应编号资源时，会绑定 Word 原生编号；嵌套有序列表和普通无序列表会把 Markdown 层级写入 Word 编号层级
- 有序列表续号/重启：已支持根据 Markdown 列表起始号（`start`）在跨列表块场景创建新的 Word 编号实例，并写入 `startOverride` 保留起始号
- 诊断解释：`style explain` 和 `doctor` 已能解释编号来源与 fallback 状态

## 已落地的主要代码模块

核心模块：

- `src/docubridge/core/template_bridge.py`
- `src/docubridge/core/style_resolver.py`
- `src/docubridge/core/layout_intent.py`
- `src/docubridge/core/word_renderer.py`
- `src/docubridge/application/render_service.py`
- `src/docubridge/cli.py`

测试模块：

- `tests/test_template_bridge.py`
- `tests/test_style_resolver.py`
- `tests/test_layout_intent.py`
- `tests/test_word_renderer.py`
- `tests/test_cli_render.py`

文档模块：

- `README.md`
- `README_CN.md`
- `docs/superpowers/specs/2026-04-14-docubridge-template-engine-design-cn.md`
- `docs/superpowers/plans/2026-04-14-docubridge-template-engine-implementation-plan.md`

## 当前用户可见行为

### 模板与 YAML 合并

当前合并顺序为：

```text
defaults -> template_document_default -> template_style -> yaml
```

含义：

- `defaults` 是 YAML 默认值
- `template_document_default` 是 Word 模板文档默认值
- `template_style` 是 Word 模板中具体样式的属性
- `yaml` 是用户在 YAML 中对某个元素显式写出的覆盖

用户可以用 `style explain` 查看字段来源：

```bash
docubridge style explain tests/fixtures/template-style.yaml ordered_list --template corp-template.docx --pretty
```

### 字体槽位

第一阶段已开始支持 Word 的多字体槽位：

- `font_ascii`
- `font_hansi`
- `font_east_asia`
- `font_cs`

这能满足中文字体和英文字体分开配置的基础要求。后续仍需要继续补复杂 run、字符样式继承和更多 Word 主题字体场景。

### 段落属性

第一阶段已支持的段落属性包括：

- `first_line_indent_pt`
- `left_indent_pt`
- `right_indent_pt`
- `space_before_pt`
- `space_after_pt`
- `alignment`

这些属性可以来自模板，也可以由 YAML 显式覆盖。

### 原生编号

当前列表编号支持：

- 从模板读取样式关联的编号资源
- 当 `ordered_list` 或普通 `unordered_list` 映射到带编号资源的 Word 样式时，绑定原生 Word 编号
- 嵌套有序列表和普通无序列表会保留模板的 `numId`，并使用 Markdown 列表层级写入 `w:ilvl`
- 当模板缺少显式需要的编号资源时，在严格模板校验路径中失败
- `style explain` 输出 `numbering` 块，说明当前是 `native` 还是 `text-prefix` fallback
- 任务列表继续保留 checkbox 文本前缀，不参与普通项目符号编号绑定

示例诊断字段：

```json
{
  "numbering": {
    "requested_style": "List Number",
    "source": "template_style",
    "available_in_template": true,
    "fallback_mode": "native"
  }
}
```

## 已知边界

当前仍不是完整 Word 排版引擎，以下能力还需要后续继续实现或验证：

- 无序列表已经支持第一阶段原生项目符号绑定与嵌套层级 `w:ilvl` 写入，但多级项目符号、跨列表续号、复杂重启规则仍需继续增强
- 嵌套有序列表和普通无序列表已支持把 Markdown 层级写入 Word 编号层级，但复杂编号格式、续号和重启规则仍需继续增强
- 标题编号与正文列表编号的统一复用策略还需要更完整的设计和测试
- 字符样式继承、主题字体、复杂中英混排细节仍需补充
- 表格样式目前主要绑定表格样式名和单元格段落属性，复杂表格边框、宽度、合并单元格策略还未完整产品化
- 节、页眉页脚、目录、分页、页边距等页面级能力仍不在第一阶段范围
- 代码块目前仍以单列表格承载，视觉效果依赖代码块内部段落样式

## 验证记录

最近一次完整验证：

```text
聚焦模板引擎测试：140 passed in 10.90s
全量测试：194 passed in 12.70s
```

测试命令：

```bash
rtk powershell -Command "pytest tests\test_template_bridge.py tests\test_layout_intent.py tests\test_style_resolver.py tests\test_word_renderer.py tests\test_cli_render.py -q --basetemp build\pytest-temp-focused"
rtk powershell -Command "pytest -q --basetemp build\pytest-temp-full"
```

说明：

- Windows 当前环境中，pytest 默认系统临时目录 `C:\Users\Administrator\AppData\Local\Temp\pytest-of-Administrator` 可能出现权限拒绝
- 后续建议继续使用仓库内 `build\pytest-temp-*` 作为 pytest 临时目录

## 打包记录

最新 wheel：

```text
dist/docubridge-0.1.0-py3-none-any.whl
时间：2026/4/17 15:59:42
```

当前环境下，标准命令曾在 `Processing .\.` 阶段超时：

```bash
rtk powershell -Command "python -m pip wheel . -w dist --no-deps"
```

已验证可用的构建方式：

```bash
rtk cmd /c python -c "from setuptools.build_meta import build_wheel; print(build_wheel('dist'))"
```

后续如果 `pip wheel` 继续卡住，可以优先使用上述 setuptools 后端直连构建命令。

## 后续建议任务

优先级一：

- 继续补多级编号的真实企业模板样本测试
- 补标题编号与普通列表编号共存时的续号/重启规则测试

优先级二：

- 扩展 `style explain` 输出，让段落属性、字体槽位和编号来源更容易被普通用户理解
- 把模板说明书中的“第一阶段已支持/未支持”同步到最新状态
- 准备一份真实企业模板样本进行人工验收

优先级三：

- 继续增强表格样式、代码块样式、图片段落样式的模板落地
- 研究页面级能力，如页眉页脚、页边距、节、目录，但不建议立即并入第一阶段收口

## 接手提示

继续开发前建议先看：

- `docs/superpowers/plans/2026-04-14-docubridge-template-engine-implementation-plan.md`
- `docs/superpowers/specs/2026-04-14-docubridge-template-engine-design-cn.md`
- `README_CN.md` 中的模板定制模型和模板排错建议

继续验证时建议先跑：

```bash
rtk powershell -Command "pytest tests\test_template_bridge.py tests\test_layout_intent.py tests\test_style_resolver.py tests\test_word_renderer.py tests\test_cli_render.py -q --basetemp build\pytest-temp-focused"
```

如果该命令通过，再跑全量测试并重建 wheel。
