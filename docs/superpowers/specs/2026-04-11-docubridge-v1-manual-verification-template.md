# DocuBridge v1 人工验证模板

> 用途：记录发布前人工验收结果  
> 建议每次正式验收复制一份再填写

## 1. 基本信息

- 验证日期：
- 验证人：
- 目标版本：
- 运行环境：
- Python 版本：
- 安装方式：

## 2. 验证对象

### 2.1 模板样本

- 模板名称：
- 模板类型：企业模板 / 论文模板 / 其他
- 模板来源：
- 模板内关键样式：

### 2.2 `.docx` 输入样本

- 样本 1：
- 样本 2：
- 样本 3：

## 3. 模板渲染链路验证

### 3.1 `doctor`

命令：

```bash
docubridge doctor input.md --style style.yaml --template template.docx
```

结果：

- [ ] 通过
- [ ] 失败

备注：

### 3.2 `style explain`

命令：

```bash
docubridge style explain style.yaml heading1 --template template.docx --pretty
```

结果：

- [ ] 通过
- [ ] 失败

备注：

### 3.3 `render`

命令：

```bash
docubridge render input.md -o out.docx --style style.yaml --template template.docx
```

结果：

- [ ] 通过
- [ ] 失败

备注：

## 4. 输出效果检查

### 4.1 `Markdown -> Word`

- [ ] 标题样式符合预期
- [ ] 正文样式符合预期
- [ ] 有序列表样式符合预期
- [ ] 无序列表样式符合预期
- [ ] 引用块样式符合预期
- [ ] 表格样式符合预期
- [ ] 代码块样式符合预期
- [ ] 图片行为符合预期

备注：

### 4.2 `Word -> Markdown`

- [ ] 标题提取可接受
- [ ] 段落顺序可接受
- [ ] 列表提取可接受
- [ ] 表格提取可接受
- [ ] 图片引用导出可接受
- [ ] 基础行内样式提取可接受

备注：

## 5. 已发现问题

- 问题 1：
- 严重度：
- 是否阻塞发布：

- 问题 2：
- 严重度：
- 是否阻塞发布：

## 6. 最终结论

- [ ] 可进入 beta 试用
- [ ] 可进入公开发布
- [ ] 不建议发布，需继续修复

结论说明：
