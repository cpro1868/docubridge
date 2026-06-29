# DocuBridge v1 Beta 发布说明

> 日期：2026-04-11  
> 版本阶段：`v1.0 beta`

## 概要

本次 `v1.0 beta` 聚焦两条主线能力：

- `Word (.docx) -> Markdown`
- `Markdown -> Word (.docx)`

同时，本版本将模板定制作为主链路能力交付，而不是附属实验能力。

## Beta 重点能力

### 1. `Word -> Markdown`

当前已支持：

- 标题、段落、列表、表格、图片引用
- 简单嵌套列表
- 简单连续编号恢复
- 标题、段落、表格单元格中的基础粗体、斜体、删除线、链接、行内代码

### 2. `Markdown -> Word`

当前已支持：

- 标题、段落、引用块、水平分隔线
- 有序列表、无序列表、任务列表
- 简单嵌套列表缩进
- 表格、代码块、图片
- 行内粗体、斜体、删除线、链接、代码

### 3. 模板定制与 YAML 样式

当前已支持：

- `--style` 驱动的显式样式控制
- `--template` 驱动的宿主 `.docx` 模板接入
- `doctor --template`
- `style explain --template`
- 标题、正文、列表、引用、表格、代码块的独立样式绑定

## 推荐使用流程

```bash
docubridge doctor input.md --style style.yaml --template corp-template.docx
docubridge style explain style.yaml heading1 --template corp-template.docx --pretty
docubridge render input.md -o out.docx --style style.yaml --template corp-template.docx
```

## 已验证状态

- 最近一次全量自动化验证：`166 passed in 5.65s`
- 当前代码状态已经达到 beta 试用门槛

## Beta 目标

本阶段目标不是继续扩格式范围，而是完成：

- 真实模板人工验证
- 真实 `.docx` 样本人工验证
- 首发已知边界管理
- 首发安装与分发准备
