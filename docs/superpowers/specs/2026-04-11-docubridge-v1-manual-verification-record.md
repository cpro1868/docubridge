# DocuBridge v1 人工验证记录

> 日期：2026-04-11  
> 状态：待执行  
> 适用版本：`v1.0 beta`

## 1. 当前自动化前提

在进入人工验证前，当前自动化状态为：

- 命令：`pytest -q`
- 最近结果：`166 passed in 5.78s`
- 当前判断：已具备进入人工验证的基础

## 2. 本轮人工验证目标

本轮人工验证只关注首发阻塞项：

- 真实模板场景下的 `doctor -> style explain -> render`
- 真实 `.docx -> markdown` 样本输出是否达到可接受水平
- README 中关键命令是否可按文档执行

## 3. 建议执行顺序

### 3.1 模板链路

1. 选择一份真实企业模板或论文模板
2. 准备一份对应的 Markdown 输入
3. 依次执行：

```bash
docubridge doctor input.md --style style.yaml --template template.docx
docubridge style explain style.yaml heading1 --template template.docx --pretty
docubridge render input.md -o out.docx --style style.yaml --template template.docx
```

4. 记录：

- 标题样式是否正确
- 列表样式是否正确
- 引用、表格、代码块样式是否正确
- 是否出现不可解释的降级

### 3.2 `Word -> Markdown`

建议至少准备 2 到 3 份真实 `.docx`：

- 一份偏企业报告
- 一份偏学术文档
- 一份带表格和图片的说明文档

重点检查：

- 标题层级
- 段落顺序
- 列表结构
- 表格提取
- 图片引用
- 粗体/斜体/删除线/链接/行内代码

### 3.3 README 命令回放

建议至少人工执行：

- `doctor`
- `style explain`
- `render`

并确认 README 中的模板示例命令与实际行为一致。

## 4. 建议记录格式

每次人工验证建议补充：

- 样本名称
- 模板名称
- 结果：通过 / 可接受 / 不可接受
- 发现的问题
- 是否阻塞 `v1.0`

## 5. 当前待完成项

- [ ] 真实模板样本验证
- [ ] 真实 `.docx` 样本验证
- [ ] README 命令人工回放
- [ ] 发布方式确认
- [ ] 版本号确认

## 6. 当前建议

当前最合理的动作不是继续扩展功能，而是：

1. 完成人工验证
2. 根据结果决定是否还有阻塞问题
3. 若无阻塞问题，按 beta 说明进入试用窗口
