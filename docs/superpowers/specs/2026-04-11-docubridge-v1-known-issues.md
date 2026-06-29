# DocuBridge v1 已知问题与边界

> 日期：2026-04-11  
> 适用版本：`v1.0 beta`

## 1. 当前明确边界

- 列表当前通过段落样式加文本前缀方式渲染，不是原生 Word 多级编号系统
- 代码块当前渲染为单列表格，其视觉效果依赖单元格内段落样式
- 模板协同当前聚焦样式与宿主文档，不覆盖复杂节、页眉页脚、目录、分页和高级编号定义
- 图片路径有效时会插入图片，失效时会降级为占位段落

## 2. 不作为 `v1.0` 阻塞项的能力

- `.xlsx -> markdown`
- `.pptx -> markdown`
- `.pdf -> markdown`
- `batch`
- GUI

## 3. 用户使用时应有的预期

- 如果目标是“快速把 Markdown 套进现有 Word 样式体系”，当前版本已经可用
- 如果目标是“完全复刻复杂企业模板中的高级编号、分页、目录和节控制”，当前版本还不适合承诺
- 如果目标是“把常见 `.docx` 送入 AI / RAG / 知识库前做结构化 Markdown 提取”，当前版本已经适合 beta 使用

## 4. 推荐排错路径

```bash
docubridge doctor input.md --style style.yaml --template template.docx
docubridge style explain style.yaml paragraph --template template.docx --pretty
docubridge render input.md -o out.docx --style style.yaml --template template.docx
```

优先看：

- 模板文件是否存在
- `style explain` 的 `word_style_name`
- `style explain` 的 `source_map`
- 目标 Word 模板内是否真的存在对应样式
