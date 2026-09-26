# Markdown2Word

Markdown 一键转 Word（.docx）工具，**LaTeX 公式转成 Word 原生可编辑公式**，不是图片。

## 功能

- 📝 Markdown → Word (.docx)
- 🧮 LaTeX 公式（`$...$` 行内、`$$...$$` 独立）→ Word 原生 OMML 公式，可在 Word/WPS 中直接点击编辑
- 🖱️ 三种用法：双击 bat、拖拽文件到 bat、命令行
- 📦 支持批量转换
- 📂 输出到原文件同目录，不修改原 .md

## 前置依赖

- [Pandoc](https://pandoc.org/installing.html)（公式转换引擎，必须安装）
- Python 3.8+

## 使用

### 方式一：拖拽（最方便）

把一个或多个 `.md` 文件直接拖到 `一键转换.bat` 图标上即可。

### 方式二：双击交互

双击 `一键转换.bat`，按提示输入文件路径或序号。

### 方式三：命令行

```bash
python md2word.py 文件1.md 文件2.md
```

## 公式支持

基于 pandoc 的 texmath 引擎，支持：

- 上标/下标：`$x^2$`、`$a_n$`
- 分式、根式：`$\frac{a}{b}$`、`$\sqrt{x}$`
- 求和/积分/连乘：`$\sum$`、`$\int$`、`$\prod$`
- 希腊字母与符号：`$\alpha$`、`$\beta$`、`$\times$`、`$\leq$`
- 重音符号：`$\bar{x}$`、`$\hat{\beta}$`
- 矩阵、分段函数：`$\begin{pmatrix}...\end{pmatrix}$`、`$\begin{cases}...\end{cases}$`
- 公式内中文：`\text{中文说明}`

### 定界符写法（自动识别）

- `$...$` 行内公式，`$$...$$` 独立公式（标准写法）
- `\[...\]` 或 `[...]` 独立公式
- `\(...\)` 行内公式

> 注意：`$$...$$` 起止符之间不要空行，否则会被当作普通文本。

## 文件说明

```
md转Word一键工具/
├── md2word.py       # 核心转换脚本
├── 一键转换.bat      # 双击/拖拽入口
├── 示例.md          # 含公式的演示文件
├── 示例.docx        # 转换结果示例
└── 使用说明.txt
```

## License

MIT
