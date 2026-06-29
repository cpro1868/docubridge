# DocuBridge 安装后的第一次使用教程

> 日期：2026-04-12  
> 适用对象：已经完成安装，准备第一次实际运行的中文用户

## 1. 目标

这份教程只做一件事：  
让你在第一次安装完成后，按顺序跑通一次最小可用流程。

跑通后，你应该能确认：

- 命令入口正常
- 样式系统正常
- Word 输出正常
- 模板工作流基本可用

## 2. 第一步：确认命令是否可用

运行：

```bash
docubridge --help
```

如果这里失败，说明安装本身还没完成，不要继续后面的步骤。

## 3. 第二步：确认环境检查命令可用

运行：

```bash
docubridge doctor
```

预期结果：

- 命令正常执行
- 输出环境检查结果

如果这一步失败，先检查：

- 当前 Python 环境是否正确
- 是否安装到了当前虚拟环境

## 4. 第三步：确认内置样式列表可用

运行：

```bash
docubridge style list
```

预期结果：

- 能列出内置样式集

这一步主要确认样式相关命令入口是否正常。

## 5. 第四步：跑通最小渲染流程

运行：

```bash
docubridge render tests/fixtures/sample.md -o build/out.docx --style tests/fixtures/style.yaml
```

预期结果：

- 生成 `build/out.docx`

这一步说明：

- Markdown 读取正常
- 样式配置读取正常
- Word 输出链路正常

## 6. 第五步：跑通最小模板工作流

如果你现在手上已经有一个可用的 `.docx` 模板，可以继续执行：

```bash
docubridge doctor tests/fixtures/template-sample.md --style tests/fixtures/template-style.yaml --template template.docx
docubridge style explain tests/fixtures/template-style.yaml heading1 --template template.docx --pretty
docubridge render tests/fixtures/template-sample.md -o build/template-demo.docx --style tests/fixtures/template-style.yaml --template template.docx
```

你在这一步要重点确认：

- 模板文件是否能被识别
- `heading1` 最终绑定到了哪个 Word 样式
- 模板输出是否成功生成

## 7. 如果你没有模板怎么办

如果你现在没有现成模板，也没关系。

你可以先只完成前四步，先确认：

- 当前安装可用
- 基础渲染可用

模板工作流可以等你拿到企业模板、学校模板或固定 Word 模板后再继续。

## 8. 第一次使用时最常见的问题

### 8.1 `docubridge` 命令找不到

通常说明：

- 没装到当前环境
- 虚拟环境没激活

### 8.2 `render` 执行了，但没有生成想要的样式

先跑：

```bash
docubridge style explain tests/fixtures/style.yaml heading1 --pretty
```

如果涉及模板，再加：

```bash
--template template.docx
```

### 8.3 模板场景失败

先跑：

```bash
docubridge doctor tests/fixtures/template-sample.md --style tests/fixtures/template-style.yaml --template template.docx
```

不要先直接 `render`。

## 9. 跑通之后下一步看什么

如果你已经完成上面的步骤，建议下一步按这个顺序看：

1. `README_CN.md`
2. `docs/superpowers/specs/2026-04-12-docubridge-5min-quickstart-cn.md`
3. `docs/superpowers/specs/2026-04-12-docubridge-user-guide-cn.md`
4. `docs/superpowers/specs/2026-04-12-docubridge-faq-cn.md`

如果你主要是普通用户，再补看：

- `docs/superpowers/specs/2026-04-12-docubridge-nontechnical-guide-cn.md`
