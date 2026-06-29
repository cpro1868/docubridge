# DocuBridge 模板排版引擎进展续更（2026-04-21）

## 本次目标

在已支持“跨列表块按 Markdown `start` 重启有序列表”的基础上，继续修复多级有序列表场景下的续号/重启稳定性，避免不同层级编号状态互相污染。

## 已完成实现

- 解析层：
  - `ListItemNode` 新增 `sequence_start` 字段。
  - Markdown 解析时，会把每个有序列表块（包含嵌套列表）的首项起始号写入该字段。
- 布局意图层：
  - `build_layout_intents` 生成有序列表 `NumberingIntent` 时，优先使用 `item.sequence_start` 判断是否重启以及 `start_at` 值。
  - 仍保留对旧节点结构的兼容回退（首项使用 `ListNode.start`）。
- Word 渲染层：
  - 编号续号状态键从 `(numbering_role, style)` 调整为 `(numbering_role, style, ilvl)`。
  - 结果是编号序列按层级隔离：嵌套有序列表可独立重启，不会打断上层列表续号。
  - 非 layout-intent 路径与 layout-intent 路径已同步该策略。

## 新增测试

- `tests/test_markdown_ingest.py`
  - `test_parse_markdown_file_preserves_nested_ordered_list_start_values`
- `tests/test_layout_intent.py`
  - `test_build_layout_intents_uses_per_level_sequence_start_for_nested_ordered_lists`
- `tests/test_cli_render.py`
  - `test_render_command_tracks_ordered_sequences_independently_per_level`

## 验证结果

- 增量用例：`3 passed`
- 模板引擎相关套件：`169 passed`
- 全量测试：`207 passed`

## 下一步 TODO（新增）

- 新增“`Word -> 模板文件` 提炼功能”：允许用户从现有 Word 文档抽取模板结构与样式配置，用于复制其他文档的结构与排版规则。

## 影响文件

- `src/docubridge/core/nodes.py`
- `src/docubridge/core/markdown_ingest.py`
- `src/docubridge/application/render_service.py`
- `src/docubridge/core/word_renderer.py`
- `tests/test_markdown_ingest.py`
- `tests/test_layout_intent.py`
- `tests/test_cli_render.py`
- `README.md`
- `README_CN.md`
