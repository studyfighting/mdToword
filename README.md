# mdToword

**把 Markdown 一键转成自带标题层级、表格、图片、代码块的 Word 文档。** 纯 Python 标准库实现，零第三方依赖（无需 pandoc / python-docx）。

Turn Markdown into a well-structured Word document (headings, tables, embedded images, code blocks) — **pure stdlib, no pandoc, no python-docx needed.**

---

## ✨ 最佳使用方式（强烈推荐） / Best usage

**下载到本地后，使用 agent（AI 编程助手 / 智能体）把它指定为工作区，然后让 AI 执行转换。**

具体做法：
1. 把本仓库 clone / 下载到本地；用支持 agent 的工具（如 Claude Code、Copilot、Cursor 等）打开这个文件夹作为工作区。
2. 把你的内容写到 `input/` 目录下的 `.md` 文件（`input/我的文档.md` 是个现成的模板）。
3. 直接对 agent 说一句：**"用 md2docx 把 input/ 下的 md 转成 Word"**。
4. agent 会调用 `python scripts/md2docx.py input/xxx.md`，生成的 Word 自动放进 `output/` 目录。

这样你只需要写 Markdown、口头下指令，转换、排版、图片内嵌都由 agent + 脚本完成。**这是这个工具设计时最推荐的使用方式。**

> The recommended usage: clone this repo, open it as the workspace in an AI agent (Claude Code / Copilot / Cursor), write your content into `input/`, and tell the agent to convert it — the generated Word lands in `output/`.

---

## 为什么做这个 / Why

在嵌入式学习与工作中，我习惯用 Markdown 记录（改起来快、能保持结构清晰、方便版本管理），但**最终分享/阅读常用 Word**。这个工具在两格式之间搭一座桥：Markdown 写内容，Word 交付成品。

- 标题 → Word 标题样式（层级清晰、可折叠、可跳转）
- 表格 → Word 表格（带边框、表头底纹）
- 图片 → 自动内嵌、按页面宽度缩放、居中
- 代码块 → 等宽字体、灰底
- 列表 → Word 带缩进的列表
- 不自动生成目录（标题样式已建好，可在 Word 里「引用 → 目录」自建）

> I prefer writing in Markdown but sharing/reading in Word. This bridges the two.

---

## 快速开始 / Quick Start

### 环境要求 / Requirements
- Python 3.8+
- （可选）[PowerShell 5.1+](https://learn.microsoft.com/powershell/) 用于生成图片的 `gen_*.ps1` 脚本（Windows 自带）

### 最简单用法 / Simplest usage

```bash
# ① 把 input/ 下的 md 转成 Word（自动输出到同级的 output/）
python scripts/md2docx.py input/我的文档.md

# ② 指定输出
python scripts/md2docx.py input.md output.docx

# ③ 转 examples/ 下的演示文档（输出与输入同目录）
python scripts/md2docx.py examples/demo.md
```

**约定**：如果输入 md 在名为 `input/` 的目录里，结果会自动存到同级的 `output/` 目录（同名 .docx）。`input/我的文档.md` 是一个可直接编辑的模板。

生成的 `demo.docx`（在 examples/ 下）包含：多级标题、正文、表格、内嵌图片、代码块、列表、分页符——正好演示全部特性。

### 各能力用法 / Features

| Markdown 语法 | 效果 |
| --- | --- |
| `#` `##` `###` … | Word 标题1~6（样式建好，进导航/可跳转） |
| `**加粗**` | 加粗 |
| `\`code\`` | 等宽行内代码 |
| `| a | b |` 表格 | Word 表格（带表头底纹） |
| `![](img/x.png)` | 图片内嵌，自动缩放居中（相对 .md 所在目录解析） |
| ```` ``` ```` 代码块 | 等宽字体 + 灰底 |
| `<!-- pagebreak -->` | 分页符 |
| `- ` / `1. ` 列表 | Word 带缩进列表 |
| 标题里的编号（如 `## 1. 串口基础`） | 原样显示，构成层级 |

---

## 脚本清单 / Scripts

| 脚本 | 作用 |
| --- | --- |
| `scripts/md2docx.py` | **核心**：Markdown → Word（标题/表格/图片/代码/列表） |
| `scripts/docx2md.py` | 逆向：Word → Markdown（把已有 docx 变成可维护的 md 源） |
| `scripts/docx_inspect.py` | 检查单个 docx：`--outline` / `--stats` / `--audit` |
| `scripts/docx_diff.py` | 比对两个 docx（`--strict` 只比正文） |
| `scripts/gen_image.ps1` | 示例：GDI+ 生成示意图 |
| `scripts/gen_uart_demo.ps1` | 生成示例图 `examples/img/uart_demo.png` |
| `scripts/gen_cpu.ps1` | 示例：CPU 内部结构示意图 |
| `scripts/gen_pyramid.ps1` | 示例：存储金字塔示意图 |

> 注：`gen_*.ps1` 用 Windows GDI+ 画图，仅适合 Windows。输出默认在脚本所在目录，可按需改。

---

## 目录结构 / Layout

```
mdToword/
├── README.md
├── input/             # 【输入】把要转成 Word 的 .md 放这里
│   └── 我的文档.md      # 可直接编辑的模板
├── output/            # 【输出】生成的 Word 自动放这里
├── scripts/           # 核心脚本
│   ├── md2docx.py
│   ├── docx2md.py
│   ├── docx_inspect.py
│   ├── docx_diff.py
│   └── gen_*.ps1
└── examples/          # 演示
    ├── demo.md         # 示例源(演示全部特性)
    ├── demo.docx       # 用 md2docx 生成
    └── img/
        └── uart_demo.png
```

> **用法简记**：内容写进 `input/`，跑 `python scripts/md2docx.py input/xxx.md`，Word 出现在 `output/`。

---

## 许可 / License

MIT

---

*我的使用场景：嵌入式学习笔记。如果你有类似需求，欢迎使用与改进。*
