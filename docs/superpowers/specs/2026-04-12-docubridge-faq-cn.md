# DocuBridge 中文 FAQ

> 日期：2026-04-12  
> 适用版本：`v1.0 beta`

## 1. 这个工具现在最适合做什么？

最适合做两件事：

- 把常见 `.docx` 转成 Markdown
- 把 Markdown 套进现有 Word 样式体系，生成 `.docx`

如果你的目标正好是这两件事，当前版本已经适合 beta 使用。

## 2. 为什么我会觉得 `parse` 和 `render` 的“保真度”不一样？

这是正常现象。

- `parse` 的目标是把文档内容结构化成 Markdown，强调“内容和结构正确”
- `render` 的目标是把 Markdown 输出成 Word，强调“结构可控、样式可控”

Markdown 本身不能承载 Word 的全部版式细节，所以这两个方向天然不是完全对称的。

## 3. 为什么我已经传了模板，结果看起来还是不对？

最常见的原因有 4 个：

- 模板里没有你以为存在的 Word 样式
- YAML 里把模板属性显式覆盖掉了
- 你期待的是复杂模板能力，但当前版本只覆盖基础样式协同
- 列表、代码块这类结构在当前版本有明确边界

建议先跑：

```bash
docubridge doctor input.md --style style.yaml --template template.docx
docubridge style explain style.yaml paragraph --template template.docx --pretty
```

## 4. 为什么样式没生效？

先检查这三个东西：

1. `word_style_name` 是不是你预期的样式名
2. `source_map` 里这个属性到底来自 `yaml` 还是 `template`
3. 模板里是否真的有对应的样式

最稳妥的命令：

```bash
docubridge style explain style.yaml heading1 --template template.docx --pretty
```

## 5. 为什么模板里有样式，最终还是用了 YAML 的值？

因为当前规则就是：

`YAML 显式配置 > 模板属性 > 默认值`

也就是说，只要你在 YAML 里明确写了某个字段，这个字段就会优先于模板。

## 6. 为什么列表看起来不像 Word 原生编号？

因为当前版本的列表重点是：

- 结构正确
- 样式可控
- 输出稳定

当前实现方式是“段落样式 + 文本前缀”，不是原生 Word 多级编号系统。这是已知边界，不是偶发 bug。

## 7. 为什么代码块看起来像表格？

因为当前版本里，代码块就是按单列表格渲染的。

这样做的好处是：

- 结构稳定
- 容易控制代码块区域

代价是：

- 最终视觉效果会受单元格内部段落样式影响

## 8. 为什么有图片时，有时插入成功，有时变成占位文字？

规则很直接：

- 图片路径有效：插入图片
- 图片路径无效：输出占位段落

占位文字不是异常崩溃，而是当前版本的可解释降级方式。

## 9. 当前版本适不适合复杂企业模板？

要分情况看。

适合：

- 套用已有标题、正文、列表、引用、表格等基础样式
- 把 Markdown 快速落到现有模板风格里

不适合高承诺场景：

- 复杂节控制
- 页眉页脚
- 目录
- 高级分页
- 原生多级编号精细恢复

## 10. 当前版本适不适合非技术用户？

如果你愿意照着文档运行几条命令，可以用。  
如果你期待的是纯图形界面产品，目前还不适合。

建议阅读顺序：

1. `README_CN.md`
2. `docs/superpowers/specs/2026-04-12-docubridge-5min-quickstart-cn.md`
3. `docs/superpowers/specs/2026-04-12-docubridge-nontechnical-guide-cn.md`

## 11. 我应该先看哪份中文文档？

按你的目标选：

- 想 5 分钟跑起来：看 `2026-04-12-docubridge-5min-quickstart-cn.md`
- 想按场景操作：看 `2026-04-12-docubridge-nontechnical-guide-cn.md`
- 想系统学会：看 `2026-04-12-docubridge-user-guide-cn.md`
- 想知道边界：看 `2026-04-11-docubridge-v1-known-issues.md`

## 12. 现在最值得先做的事是什么？

如果你是使用者：

- 先用示例文件跑通一次 `doctor -> style explain -> render`

如果你是负责人：

- 先用真实模板和真实 `.docx` 做一轮人工验收

如果你是开发者：

- 先看 `README.md`、`README_CN.md`、`docs/requirements.md`
