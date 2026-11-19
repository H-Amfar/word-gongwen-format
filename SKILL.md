---
name: gongwen-word-format
version: "1.1.0"
display_name: "公文 Word 格式化"
display_name_en: "Government Document Word Formatter"
description: "公文 Word 文档格式化技能。当用户需要按公文规范排版 .docx 文件（设置标题与正文字体、字号、段落格式）时使用。规范：标题一为方正大标宋简体（字号按字数自适应）、标题二为方正楷体GBK三号、标题三为方正仿宋GB2312四号（均加粗）、表格内正文为方正仿宋GB2312四号、正文为方正仿宋GB2312三号；正文段落首行缩进2字符、段前段后各8磅、行距1.5倍、两端对齐、字体黑色。触发词：公文格式、公文字体、公文排版、方正大标宋、方正仿宋、方正楷体。This skill should be used when formatting Word documents according to Chinese official document (gongwen) typography standards."
description_zh: "公文 Word 文档格式化技能。当用户需要按公文规范排版 .docx 文件（设置标题与正文字体、字号、段落格式）时使用。规范：标题一为方正大标宋简体（字号按字数自适应）、标题二为方正楷体GBK三号、标题三为方正仿宋GB2312四号（均加粗）、表格内正文为方正仿宋GB2312四号、正文为方正仿宋GB2312三号；正文段落首行缩进2字符、段前段后各8磅、行距1.5倍、两端对齐、字体黑色。触发词：公文格式、公文字体、公文排版、方正大标宋、方正仿宋、方正楷体。"
description_en: "Formats Word documents (.docx) according to Chinese official document (gongwen) typography standards. Use when the user needs to apply official font, size, and paragraph rules to a .docx file. Rules - Title 1 in FZ Biaosong (auto-sized by length, bold), Title 2 in FZ Kaiti GBK 16pt bold, Title 3 and table body in FZ Fangsong GB2312 14pt bold, body text in FZ Fangsong GB2312 16pt; body paragraphs use 2-character first-line indent, 8pt spacing before and after, 1.5 line spacing, justified alignment, and black font."
agent_created: true
---

# Word 公文格式技能

本技能按固定公文排版规范处理 .docx 文件。核心规范如下：

## 字体规范（强制）

| 元素 | 字体 | 字号 | 加粗 |
|------|------|------|------|
| 方案标题（标题一） | 方正大标宋简体 | 根据字数调整（见下表） | 是 |
| 标题二 | 方正楷体GBK | 三号（16pt） | 是 |
| 标题三 | 方正仿宋GB2312 | 四号（14pt） | 是 |
| 正文 | 方正仿宋GB2312 | 三号（16pt） | 否 |
| 表格内正文 | 方正仿宋GB2312 | 四号（14pt） | 否 |

### 标题一字号自适应规则

根据标题字数选择字号（注意：中文字号从小到大为 三号(16pt)、小二(18pt)、二号(22pt)、小一(24pt)、一号(26pt)、小初(36pt)、初号(42pt)）：

- ≤15 字：二号（22pt）
- 16–20 字：小二（18pt）
- 21–30 字：三号（16pt）
- >30 字：小二（18pt）并允许换行折行处理

## 正文段落格式（强制）

- 首行缩进 2 字符
- 段前 8 磅、段后 8 磅
- 行距 1.5 倍
- 对齐方式：两端对齐（justify），方向从左到右
- 字体颜色：黑色（RGB 000000）

## 使用方法

### 方式一：对已有 .docx 直接套用格式（推荐）

运行捆绑脚本，自动识别 Word 内置样式（Heading 1/2/3、Normal）并全量套用上述规范：

```bash
python scripts/apply_gongwen_format.py <input.docx> [--output <output.docx>]
```

- 未指定 `--output` 时在原文件名后追加 `_公文格式` 保存，不覆盖原文件。
- 脚本依赖 `python-docx`（pip install python-docx）。
- 标题一字号按上表根据实际标题文字长度自适应。
- 脚本会同时把表格内所有单元格文本设为方正仿宋GB2312四号。

### 方式二：从零创建公文文档

新建 Word 文档（可用 workbuddy-docx 相关流程），然后运行上述脚本对成品套用格式；或在编写 HTML/生成 docx 时直接按本规范的字体、字号、段落格式写入样式。

## 注意事项

- 方正大标宋简体、方正楷体GBK、方正仿宋GB2312 为方正商用字体，需本机安装才能正确显示；脚本仍会写入字体名，文档在装有这些字体的机器上可正常显示。
- 若文档使用自定义样式名而非内置样式，按大纲级别（outline level）判断标题层级。
- 页眉页脚、页码等默认不处理，除非用户明确要求。
