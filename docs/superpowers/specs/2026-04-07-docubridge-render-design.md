# DocuBridge v1 设计文档

> 日期：2026-04-07  
> 状态：设计确认稿  
> 关联文档：`requirements.md`、`docs/requirements.md`

## 1. 概述

本文档用于补充和细化 DocuBridge 的实现设计，承接现有 PRD 中已经确认的产品方向，并将本轮设计讨论形成可实现、可验证、可演进的技术方案。

本设计遵循以下原则：

- v1 明确以 `Markdown -> Word` 为核心主线
- 产品形态采用 `Python core + CLI first + GUI 预留边界`
- 设计细到足以指导实现、测试和拆分计划
- 对高风险和高变动区域，明确稳定边界与可替换实现，避免把 v1 写死

### 1.1 v1 目标

- 支持“接近完整 GFM”的 Markdown 输入，并生成 `.docx`
- 支持稳定可用的 `.docx -> markdown` 主场景
- 支持高级可控的 YAML 样式配置
- 支持模板文件 `--template` 作为一级输入能力
- 提供正式可用的 CLI，而不是临时脚本入口

### 1.1.1 v1 发布口径

`v1.0` 的发布门槛只围绕三件事：

- `Word (.docx) -> Markdown` 稳定可用
- `Markdown -> Word (.docx)` 稳定可用
- 模板定制与 YAML 样式配置达到可交付水平

这意味着：

- `.xlsx/.pptx/.pdf -> markdown` 可以继续实现和保留，但不作为 `v1.0` 是否发布的阻塞条件
- GUI、`batch`、更高保真的版面恢复都属于后续迭代，不进入首次发布门槛
- 模板定制不是附属能力，而是 `Markdown -> Word` 主链路的一部分

### 1.2 v1 非目标

- 不在 v1 内完整实现 GUI
- 不在 v1 内将 `docx/xlsx/pptx/pdf -> Markdown` 作为同等优先级主线
- 不承诺完整支持所有 Word 底层高级特性
- 不承诺完整支持 Markdown 中所有 HTML 片段的等价 Word 渲染
- 不要求在 `v1.0` 发布前完成 `xlsx/pptx/pdf` 全量收口

### 1.3 设计结论摘要

本轮已确认的核心设计结论如下：

- v1 的核心价值是“YAML 驱动的 Markdown -> Word 高可控渲染”
- 渲染内核采用“块级轻量节点树 + 行内片段渲染”的混合模式
- 模板与 YAML 同为一级输入，但模板提供资源与宿主能力，YAML 决定显式样式意图
- CLI 采用“子命令清晰 + 高级参数展开”的折中命令模型
- 样式系统、模板系统、渲染系统、诊断系统必须解耦

### 1.4 当前实现状态（2026-04-09）

当前代码库已经落地的范围主要集中在 `Markdown -> Word` 的最小可用链路、`docx/xlsx/pptx -> markdown` 的最小解析切片，以及 CLI 输出层：

- 已实现轻量 Markdown 节点解析，覆盖标题、段落、列表、任务列表，以及简单嵌套有序/无序列表层级恢复
- 已实现 YAML 样式加载、override 校验、样式解析与基础模板视图
- 已实现基础 `.docx` 渲染链路，覆盖标题、段落、基础行内样式与外链、引用块及其行内样式保留、水平分隔线、列表、列表项与任务列表项内行内样式保留、简单嵌套列表缩进、Markdown 表格及单元格内行内样式保留、fenced code block 和独立图片块，其中基础行内样式已覆盖粗体、斜体、删除线和行内代码
- 已实现基础 `.docx -> markdown` 解析链路，覆盖标题、段落、项目列表、简单连续编号列表、简单嵌套列表、表格、标题/段落/表格单元格内基础行内样式、基础行内代码与链接保留、图片引用导出
- 已实现基础 `.xlsx -> markdown` 解析链路，覆盖多 sheet、基础表格输出、日期 ISO 格式化与长文本截断
- 已实现基础 `.pptx -> markdown` 解析链路，覆盖 slide 标题、标题占位符、正文文本要点、普通文本框与分层列表区分输出、文本框与表格单元格内基础粗体/斜体/链接/行内代码保留、表格提取、图片导出、speaker notes 提取，以及按版面位置输出基础块顺序
- 已实现 `parse`、`render`、`doctor`、`style list/show/validate/explain/merge`
- 已实现 `render --json`、`doctor --json`、`style show --json`、`style validate --json`
- 已实现 `parse --json`
- 已实现 `style explain --pretty`、`style merge --pretty`

当前仍未落地或只停留在设计层的能力包括：

- `parse` 的多格式正向转换主线（`pdf`）
- `batch` 批量处理
- 外部模板文件真正接入 `render`
- 目录、分页、编号管理器、表格/图片/代码块的完整 Word 渲染
- 严格/宽松模式的完整用户入口
- `parse` 中更高保真的图片定位、复杂表格、Excel 脚注/图表/公式策略、复杂嵌套列表恢复，以及 `.pptx` 的动画/复杂形状提取与更精细的版面恢复

### 1.5 当前发布判断（2026-04-10）

从发布视角看，当前项目不再追求“所有格式同步完成”，而是进入首个可交付版本的收口阶段。

发布阻塞项应只保留：

- `.docx -> markdown` 主场景稳定性和回归覆盖
- `markdown -> .docx` 主场景稳定性和回归覆盖
- 模板接入、样式合并、样式诊断、样式样例与说明文档

非阻塞项包括：

- `.xlsx/.pptx/.pdf` 能力继续增强
- `batch`
- GUI
- 更高保真的版面细节恢复

## 2. 架构总览

总体架构采用分层设计：

- `core`
- `application`
- `interfaces`
- `adapters`

### 2.1 分层结构

#### core

核心业务层，负责：

- Markdown 解析为内部节点模型
- YAML 样式配置解析、归一化、继承和校验
- 模板资源抽象
- 样式合并与编号解析
- Word 渲染
- 诊断模型

该层不直接感知 CLI 或 GUI。

#### application

应用编排层，负责：

- 接收标准化任务请求
- 组织一次完整渲染流程
- 控制严格/宽松模式
- 收集日志、警告、统计
- 产出统一任务结果

#### interfaces

接口层包含：

- `interfaces.cli`
- `interfaces.gui`

其中 v1 只实现 CLI，GUI 仅预留调用边界，不参与主链路设计。

#### adapters

适配层负责封装第三方库和外部资源：

- Markdown 解析器
- YAML 解析器
- `python-docx`
- 文件系统
- 图片资源读取
- 模板加载

### 2.2 分层设计原因

采用此结构的原因如下：

- CLI 参数未来可变化，不应打穿渲染内核
- GUI 未来接入时不应重写核心逻辑
- 模板与 Word 底层库较易受实现细节影响，应局限在 adapter 和 bridge 层
- 样式系统和渲染系统都较复杂，必须通过稳定内部对象协作

## 3. 核心数据模型

### 3.1 RenderRequest

统一表示一次渲染请求，建议字段包括：

- `input_path`
- `output_path`
- `style_path`
- `template_path`
- `profile_name`
- `mode`：`strict | lenient`
- `overwrite`
- `resource_dir`
- `log_level`
- `dump_ast`
- `output_mode`：`human | quiet | json`
- `features`
- `overrides`：来自 `--set`

其作用是统一 CLI 和未来 GUI 的输入模型。

### 3.2 RenderContext

表示一次任务执行期间的运行时上下文，建议包括：

- 已解析 Markdown 节点树
- `StyleProfile`
- `TemplateView`
- `ResolvedStyle` 集合
- `NumberingContext`
- 资源解析器
- 诊断收集器
- 文档功能开关

### 3.3 RenderResult

表示一次渲染结果，建议至少包括：

- `success`
- `output_path`
- `fatal_errors`
- `warnings`
- `stats`
- `diagnostics`

其中 `stats` 建议包含：

- 标题数
- 段落数
- 列表数
- 表格数
- 图片数
- 降级节点数
- 总耗时

### 3.4 Diagnostic

统一表示错误、警告和提示信息，建议字段：

- `severity`
- `code`
- `message`
- `location`
- `hint`

建议的严重级别：

- `info`
- `warning`
- `error`
- `fatal`

### 3.5 DocumentNode

块级节点统一基类，建议具备：

- `type`
- `source_span`
- `attributes`
- `children`

主要子类型包括：

- `HeadingNode`
- `ParagraphNode`
- `ListNode`
- `ListItemNode`
- `QuoteNode`
- `CodeBlockNode`
- `TableNode`
- `ImageBlockNode`
- `HorizontalRuleNode`
- `HtmlBlockNode`
- `RawBlockNode`

### 3.6 InlineSpan

行内片段模型建议包括：

- `TextSpan`
- `StrongSpan`
- `EmphasisSpan`
- `StrikeSpan`
- `CodeSpan`
- `LinkSpan`
- `ImageSpan`
- `LineBreakSpan`
- `RawInlineSpan`

### 3.7 StyleProfile 与 ResolvedStyle

`StyleProfile` 表示 YAML 解析、继承、归一化后的内部样式对象。  
`ResolvedStyle` 表示样式与模板合并后的最终生效样式。

`ResolvedStyle` 建议包含：

- `element_name`
- `word_style_name`
- `resolved_properties`
- `source_map`

`source_map` 用于记录每个字段来自：

- `set`
- `yaml`
- `template`
- `defaults`
- `system`

### 3.8 TemplateView

模板资源抽象视图，负责向上层暴露：

- 可用样式
- 可用编号定义
- 文档默认设置
- 页眉页脚能力
- 主题信息
- 节级能力

## 4. Markdown 解析设计

### 4.1 解析策略

采用“块级解析 + 行内解析”的混合模式：

- 块级先构建轻量节点树
- 行内对富文本区域做片段解析

该策略兼顾：

- 对复杂块级结构的稳定表达
- 对 Word 段落与 run 层渲染的自然映射
- 对未来语法扩展的可维护性

### 4.2 块级节点范围

建议 v1 支持以下块级节点：

- 标题
- 段落
- 引用
- 列表
- 列表项
- 代码块
- 表格
- 图片块
- 分隔线
- HTML 块
- 原始块兜底

块级节点应保留 `source_span`，至少可回溯源行号。

### 4.3 行内节点范围

建议 v1 支持以下行内节点：

- 纯文本
- 粗体
- 斜体
- 删除线
- 行内代码
- 链接
- 图片
- 自动链接
- 硬换行
- 原始行内兜底

### 4.4 GFM 支持分层

v1 采用支持分层策略：

#### 强支持

- 标题
- 段落
- 强调
- 列表
- 引用
- 代码块
- 表格
- 图片
- 链接

#### 兼容支持

- 任务列表
- 删除线
- 自动链接
- 行内代码

#### 降级支持

- HTML 片段
- 脚注
- 非常规嵌套结构

### 4.5 任务列表策略

任务列表内部仍可建模为列表项，但附加：

- `task = true`
- `checked = true | false`

渲染层可将其输出为：

- `☐ / ☑`
- 或映射为特定列表样式

### 4.6 HTML 与脚注策略

#### HTML

不尝试完整做 `HTML -> Word` 等价转换。建议策略：

- 解析器识别为 `HtmlBlockNode` 或 `RawInlineSpan`
- 宽松模式下保留文本并告警
- 严格模式下可提升为错误

#### 脚注

一期不生成 Word 原生脚注。解析器只负责识别与保留引用，渲染器输出普通文本或保留标记。

## 5. 样式系统设计

这是 v1 的核心章节。

### 5.1 顶层结构

建议样式配置采用以下顶层结构：

- `meta`
- `defaults`
- `elements`
- `multilevel_list`
- `document`
- `assets`
- `compat`

推荐示意：

```yaml
meta:
  name: academic
  version: 1

defaults:
  font_name: Times New Roman
  font_size: 12
  line_spacing: 1.5

elements:
  heading1:
    based_on: Normal
    font_name: 黑体
    font_size: 18
    bold: true
    outline_level: 0
  paragraph:
    based_on: Normal
    first_line_indent: 0.75cm

multilevel_list:
  enabled: true
  levels:
    - level: 1
      element: heading1
      number_format: "%1."

document:
  toc:
    enabled: true
    depth: 3

assets:
  default_image_width: 500px

compat:
  unknown_field: error
```

### 5.2 样式字段范围

v1 支持三层样式能力：

- 基础排版属性
- 结构样式属性
- 高级控制属性

#### 基础排版属性

- 字体
- 字号
- 颜色
- 加粗
- 斜体
- 下划线
- 对齐
- 行距
- 缩进
- 段前段后

#### 结构样式属性

- 大纲级别
- 列表编号
- 表格样式
- 代码块样式
- 图片宽度与对齐

#### 高级控制属性

- 继承链
- 模板映射
- 样式覆盖
- 编号规则
- 回退策略

### 5.3 样式校验

样式校验分为：

- 结构校验
- 类型校验
- 引用校验
- 约束校验
- 模式校验

重点包括：

- 未知字段
- 类型不匹配
- 非法单位
- `based_on` 引用缺失
- 继承循环
- 编号层级冲突
- 模板样式名映射问题

### 5.4 继承与归一化

YAML 解析后不能直接进入渲染，必须执行：

- 单位标准化
- 枚举值标准化
- 布尔值归一化
- 继承展开
- 默认值补齐
- 字段来源追踪

需要特别区分：

- 配置继承
- Word 样式继承

二者不应机械绑定。

### 5.5 `--set` 覆盖

CLI 支持任意 YAML 路径覆盖：

```bash
--set document.toc.depth=3
--set elements.heading1.font_size=20
```

规则如下：

- `--set` 在 YAML 解析后、归一化前应用
- `--set` 仍必须通过完整校验
- `--set` 优先级高于 YAML 显式字段

最终优先级为：

1. `--set`
2. YAML 显式字段
3. 模板对应属性
4. `defaults`
5. 系统默认值

### 5.6 模板与样式合并

合并采用“属性级合并”，而不是整样式覆盖。

任一属性按以下顺序解析：

1. `--set`
2. YAML 显式值
3. 模板同名或映射样式中的对应属性
4. `defaults`
5. 系统默认

### 5.7 严格模式与宽松模式

#### 严格模式

- 未知字段直接报错
- 非法值报错
- 继承循环报错
- 引用缺失报错
- 关键结构冲突报错

#### 宽松模式

- 未知字段告警后忽略
- 非法值告警并回退默认
- 高级能力缺失可降级
- 局部失败可回退默认样式

### 5.8 样式诊断命令的设计依据

由于样式系统高级可控，必须提供配套诊断：

- `style validate`
- `style explain`
- `style merge`

其中：

- `style explain` 输出最终结果和属性来源
- `style merge` 输出 resolved profile 和冲突报告

当前实现中，`style explain` 对缺失的 `headingN` 也会按运行时 fallback 规则给出有效结果，即复用 `heading1` 的解析后属性并显示对应的 `Heading N` 样式名。

## 6. Word 渲染内核设计

### 6.1 总体结构

渲染内核建议分为三层：

- `document composer`
- `block renderer`
- `inline renderer`

并辅以：

- `numbering_manager`
- `asset_resolver`
- `fallback_policy`

### 6.2 渲染总流程

建议固定顺序：

1. 构建 `RenderContext`
2. 读取并解析 Markdown
3. 加载 YAML 和模板
4. 解析 `ResolvedStyle`
5. 初始化 `Document`
6. 注册样式和编号
7. 渲染块级节点
8. 处理目录、分页、页码等文档能力
9. 输出诊断并保存文件

### 6.3 标题渲染

标题需要同时处理：

- 可视样式
- 大纲级别
- 多级编号
- 目录联动
- 分页控制

建议流程：

1. 创建段落
2. 应用标题样式
3. 渲染行内内容
4. 绑定大纲级别
5. 根据配置绑定标题编号
6. 应用分页等段落级控制

标题编号必须通过 Word 编号体系实现，不应手工拼接文本。

### 6.4 段落与引用

普通段落直接使用对应样式并渲染行内片段。  
引用块 v1 优先使用：

- 引用样式
- 缩进
- 字体或颜色差异

不强求复杂视觉装饰。

### 6.5 列表与编号

列表分两类处理：

- 正文列表
- 标题多级编号

#### 正文列表

- 无序列表
- 有序列表
- 任务列表

#### 标题编号

与正文列表完全分离，由 `multilevel_list` 驱动，并与目录联动。

`numbering_manager` 负责：

- 模板编号定义查询
- YAML 编号定义创建
- 段落编号绑定
- 起始值与缩进控制

### 6.6 代码块渲染

代码块在 Word 中通过 `1x1` 表格模拟。建议流程：

1. 创建单列表格
2. 设定边框与背景
3. 在单元格内创建段落
4. 应用等宽字体与代码样式
5. 原样输出代码文本

代码块不再走普通 Markdown 行内渲染。

### 6.7 表格渲染

表格由专门的 `table_renderer` 处理，至少支持：

- 行列创建
- 表头与正文区分
- 单元格文本渲染
- 对齐
- 表头重复
- 表格样式应用

Markdown 规范表格为强支持对象，异常表格可在宽松模式下降级。

### 6.8 图片渲染

图片路径解析与文档插入解耦：

- `asset_resolver` 负责资源解析
- `image_renderer` 负责插图

需要支持：

- 相对路径解析
- 默认资源目录
- 图片存在性检查
- 保持比例
- 默认宽度
- 对齐控制

宽松模式下，缺图可降级为文本占位；严格模式下可报错。

### 6.9 HTML 与降级策略

所有不支持结构统一走 `fallback_policy`，支持三类行为：

- `preserve_text`
- `annotate`
- `fail`

不要让每个 renderer 各自发明降级逻辑。

### 6.10 文档级能力

#### v1 原生支持

- 标题大纲级别
- 目录占位
- 手动分页
- 基础 section
- 页码启用开关

#### 模板优先承载

- 页眉页脚
- 企业封面
- 复杂页面布局
- 奇偶页差异
- 复杂节切换

## 7. 模板协同设计

### 7.1 模板定位

模板是一级输入能力，但其职责是：

- 提供现有 Word 样式和编号定义
- 提供页面、节、主题、页眉页脚等宿主资源

模板不负责主导：

- Markdown 语义映射
- YAML 显式样式意图
- 错误语义
- 降级行为

### 7.2 接入方式

模板应通过 `template_bridge` 读取后转为 `TemplateView`，而不是直接把底层文档对象传遍系统。

### 7.3 样式映射

模板样式绑定建议支持两种方式：

- 约定绑定
- YAML 显式声明 `template_style`

当前实现中，若 `headingN` 未显式声明 `template_style`，则默认约定映射到 Word 的 `Heading N`；普通段落元素默认映射到 `Normal`。若样式配置中缺失更高层级的 `headingN` 元素，但存在 `heading1`，则当前实现会复用 `heading1` 的解析后属性，同时仍绑定到对应的 `Heading N` Word 样式名。

### 7.4 编号复用

当启用 `multilevel_list` 时：

- 先判断模板中是否存在符合要求的编号定义
- 满足明确匹配条件时复用
- 否则动态创建

匹配判断至少包括：

- 层级数
- 编号格式
- 样式绑定
- 起始值
- 缩进

### 7.5 模板失败语义

如果用户显式传入模板，则模板失败不能被静默忽略。

#### 致命情况

- 模板文件不存在
- 模板无法打开

#### 可恢复情况

- 某样式缺失
- 某属性无法读取
- 某编号定义不兼容

严格模式和宽松模式按规则决定报错还是回退。

## 8. CLI 设计

### 8.1 总体命令模型

CLI 采用“子命令清晰 + 高级参数展开”的折中设计。

建议主命令包括：

- `docubridge render`
- `docubridge parse`
- `docubridge style`
- `docubridge batch`
- `docubridge doctor`
- `docubridge version`

其中 v1 主命令是 `render`。

当前已落地的命令子集为：

- `docubridge parse`
- `docubridge render`
- `docubridge doctor`
- `docubridge style list`
- `docubridge style show`
- `docubridge style validate`
- `docubridge style explain`
- `docubridge style merge`

### 8.2 render 命令

采用分层参数模型。

#### 基础层

- `input`
- `-o, --output`
- `--style`
- `--template`
- `--strict`
- `--lenient`

#### 常用增强层

- `--overwrite`
- `--resource-dir`
- `--profile`
- `--quiet`
- `--json`

#### 高级层

- `--feature`
- `--set`
- `--dump-ast`
- `--log-file`
- `--log-level`

当前已实现参数：

- `input_path`
- `-o, --output`
- `--style`
- `--json`

### 8.3 `--feature`

v1 只保留少量稳定 feature：

- `toc`
- `title-numbering`
- `page-number`

更多细粒度控制应进入 YAML 或 `--set`。

### 8.4 输出模式

支持三种输出模式：

- 默认人类可读输出
- `--quiet`
- `--json`

默认模式应提供：

- 阶段信息
- 关键 warning 摘要
- 输出路径
- 结果摘要

当前实现情况：

- `parse` 支持默认文本输出与 `--json`
- `render` 支持默认文本输出与 `--json`
- `doctor` 支持默认文本输出与 `--json`
- `style show`、`style validate` 支持默认文本输出与 `--json`
- `style explain`、`style merge` 默认输出 JSON，并支持 `--pretty`

### 8.5 退出码

v1 建议采用少量高价值退出码：

- `0` 成功
- `2` 参数或用户输入错误
- `3` 输入文件或资源错误
- `4` 样式配置或模板错误
- `5` 渲染执行错误
- `6` 环境依赖错误

CLI 对样式或模板校验失败返回 `4`，对应的诊断码建议以 `STYLE_` 或 `TEMPLATE_` 开头，便于前端统一映射退出行为。

### 8.6 style 子命令

建议支持：

- `style init`
- `style validate`
- `style list`
- `style show`
- `style explain`
- `style merge`

其中：

- `style explain` 输出最终样式与来源
- `style merge` 输出合并结果与冲突报告

当前已实现的 `style` 子命令：

- `list`
- `show`
- `validate`
- `explain`
- `merge`

当前尚未实现：

- `init`

### 8.7 doctor 子命令

`doctor` 采用双入口：

#### 通用检查

```bash
docubridge doctor
```

检查环境和安装。

#### 任务检查

```bash
docubridge doctor input.md --style style.yaml --template corp.dotx
```

对具体任务执行预诊断。

诊断范围包括：

- 环境层
- 资源层
- 任务层

当前实现的 `doctor` 结果模型包括：

- 分项检查输出：`environment`、`markdown`、`style`、`style-resolution`
- 非致命 warning 扫描与摘要
- `--json` 结构化结果：`success`、`summary`、`checks`、`warnings`、`error`

### 8.8 batch 子命令

v1 中 `batch` 仍按统一任务模型组织，不应单独复制渲染逻辑。它负责：

- 目录扫描
- 任务构造
- 并发调度
- 错误汇总

## 9. GUI 预留设计

v1 不实现 GUI 主链路，但必须预留未来接入边界。

GUI 应只消费稳定接口：

- `RenderRequest`
- `RenderResult`
- `Diagnostic`
- `ProgressEvent`

不应直接拼接或依赖内部模块。

未来 GUI 建议形态是“文档渲染工作台”，但不反向影响 v1 CLI 架构。

## 10. 错误模型与诊断设计

### 10.1 错误分级

- `fatal`
- `error`
- `warning`
- `info`

### 10.2 失败语义

建议采用“分级失败”：

- 致命错误直接中止
- 结构错误在宽松模式下可降级
- 样式应用错误按策略回退或报错

### 10.3 诊断码

建议为常见问题提供稳定诊断码，例如：

- `STYLE_UNKNOWN_FIELD`
- `STYLE_INHERITANCE_CYCLE`
- `TEMPLATE_STYLE_MISSING`
- `NUMBERING_TEMPLATE_MISMATCH`
- `IMAGE_NOT_FOUND`
- `HTML_BLOCK_DEGRADED`

## 11. 测试与验证策略

建议至少覆盖以下测试层：

- YAML schema 测试
- 继承与合并测试
- 编号与标题测试
- 模板协同测试
- 表格、图片、代码块渲染测试
- HTML/脚注降级测试
- CLI 集成测试
- 诊断输出测试

## 12. 演进与替换点

### 12.1 稳定边界

- `RenderRequest`
- `RenderResult`
- `Diagnostic`
- `DocumentNode`
- `StyleProfile -> ResolvedStyle`
- `TemplateView`
- `fallback_policy`

### 12.2 可替换实现

- Markdown 解析库
- YAML 解析器实现细节
- `python-docx` 封装方式
- 目录插入策略
- 编号匹配算法
- GUI 接入协议

### 12.3 后续扩展方向

- 将 `parse` 主线补齐为等优先级能力
- 增加更多样式字段
- 增加原生脚注支持
- 增加 GUI 配置界面
- 增加模板与样式的可视化调试工具

## 13. 未决与说明

当前设计已经足以进入实现计划阶段，但以下事项仍可在后续计划或实现阶段细化：

- 最终选定的 Markdown 解析库
- 最终选定的 YAML schema 校验库
- `python-docx` 在目录、编号、节处理上的具体实现封装
- GUI 与 Python core 的最终通信方式

这些属于实现替换点，不影响本文档的总体结构与边界定义。
