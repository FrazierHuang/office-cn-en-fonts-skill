# office-cn-en-fonts-skill

用于 Codex 的 Office 输出格式控制 skill。它要求生成或后处理 Word、Excel、PowerPoint 文件时统一中英文字体，并约束 Word/Excel 的基础表格样式。

## 功能

- Word、Excel、PPT 中文使用宋体/SimSun。
- Word、Excel、PPT 英文、数字和西文标点使用 Times New Roman。
- 中文和英文/数字直接相连时不保留空格，例如 `河蚬As`、`As暴露`。
- Word 和 Excel 全部使用黑色字体。
- Word 和 Excel 表格不使用任何填充色或底纹。
- 提供 `scripts/enforce_office_fonts.py`，可对 `.docx`、`.xlsx`、`.pptx` 做 OOXML 后处理。

## 安装

将整个目录复制到 Codex skills 目录：

```bash
cp -R office-cn-en-fonts-skill ~/.codex/skills/
```

复制后重新打开 Codex 会话，让新 skill 被加载。

## 使用场景

当你要求 Codex 生成或整理 Office 文件时，可以直接提出类似需求：

```text
帮我生成一个 Word 文档，使用 office-cn-en-fonts-skill 的格式要求。
```

或：

```text
帮我整理这个 Excel，中文宋体、英文 Times New Roman，黑色字体，表格不要填充色，中英文之间不要空格。
```

## 后处理脚本

对已经生成的 Office 文件，可运行：

```bash
python3 scripts/enforce_office_fonts.py --in-place file.docx file.xlsx file.pptx
```

不加 `--in-place` 时，脚本会生成带 `_fontfixed` 后缀的新文件。

脚本会执行：

- DOCX：设置 run/style 字体，强制黑字，移除表格底纹，清理中英文边界空格。
- XLSX：设置字体颜色为黑色，重置填充样式，将共享字符串/内联字符串拆成中文和非中文富文本 run，并清理中英文边界空格。
- PPTX：设置 Latin 字体为 Times New Roman，East Asian 字体为宋体，并清理中英文边界空格。

## 文件结构

```text
office-cn-en-fonts-skill/
├── SKILL.md
├── agents/
│   └── openai.yaml
├── scripts/
│   └── enforce_office_fonts.py
└── read.md
```

## 注意

- Excel 中混合中英文的单元格会被脚本改写为富文本 run，以便分别指定中文和英文字体。
- PPT 不强制黑色字体；只有 Word 和 Excel 默认强制黑色字体。
- 如果用户明确要求彩色字体或表格填充色，应以用户的明确要求为准。
