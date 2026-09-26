#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2word.py —— Markdown 转 Word 一键工具
=========================================
公式（LaTeX 语法 $...$ 行内 / $$...$$ 独立）经 pandoc 精准转换为
Word 原生公式（OMML），可在 Word / WPS 中直接点击编辑，绝非图片。

定界符兼容（自动识别）：
  $...$ / $$...$$        标准 Markdown 公式写法
  \\[ ... \\] / [ ... ]    显示公式（内容含 LaTeX 命令时自动改写为 $$...$$）
  \\( ... \\)              行内公式（内容含 LaTeX 命令时自动改写为 $...$）
用法（三种方式任选）：
  1. 双击“一键转换.bat”，进入交互模式；
  2. 把一个或多个 .md 文件拖到“一键转换.bat”图标上，自动批量转换；
  3. 命令行：python md2word.py <md 文件...>

转换结果：与源 md 同目录、同文件名的 .docx。
"""

import os
import re
import sys
import glob
import shutil
import tempfile
import subprocess
import zipfile

APP_NAME = "md转Word一键工具"
VERSION = "1.1"

# pandoc 常见安装位置（找不到时按顺序尝试）
PANDOC_CANDIDATES = [
    r"%LOCALAPPDATA%\Pandoc\pandoc.exe",
    r"%USERPROFILE%\Anaconda3\Scripts\pandoc.exe",
    r"%USERPROFILE%\Miniconda3\Scripts\pandoc.exe",
    r"%USERPROFILE%\anaconda3\Scripts\pandoc.exe",
    r"C:\Program Files\Pandoc\pandoc.exe",
    r"C:\Program Files (x86)\Pandoc\pandoc.exe",
    r"C:\Python314\Scripts\pandoc.exe",
]

MD_EXT = (".md", ".markdown", ".mdown", ".mkd")

# ============================================================
# [ ... ] / \[ ... \] / \( ... \) 公式定界符预处理
# 旧版 pandoc 不识别这些 LaTeX 风格的公式定界符（会当作普通文本），
# 这里在转换前把“内容含 LaTeX 命令”的方括号公式改写为标准 $$...$$ / $...$。
# 注意避免误伤 Markdown 链接 [text](url)、行内代码与围栏代码块。
# ============================================================

# 公式内容判定的 LaTeX 命令特征（命中即视为公式）
_LATEX_CMD_RE = re.compile(
    r"\\(?:begin|end|frac|sqrt|sum|int|prod|lim|left|right|qquad|quad|"
    r"infty|pm|times|cdot|div|bar|hat|overline|underline|text|mathbf|"
    r"mathrm|operatorname|neq|le|ge|approx|equiv|log|ln|sin|cos|tan|"
    r"partial|nabla|limits)\b"
)


def _collect_across_lines(lines, idx, seg, start, term):
    """从 seg[start:] 开始收集，必要时跨后续行，直到找到 term。
    返回 (内容, 额外消费的行数, 是否闭合, 闭合符之后的尾部文本)。"""
    joined = seg[start:]
    li = idx
    while True:
        j = joined.find(term)
        if j != -1:
            return joined[:j], li - idx, True, joined[j + len(term):]
        li += 1
        if li >= len(lines):
            return joined, li - idx - 1, False, ""
        joined += "\n" + lines[li]


def _convert_segment(seg, lines, idx):
    """扫描一个“反引号外”的文本段，改写 [..] / \\[..\\] / \\(..\\)。
    返回 (改写后的段, 额外消费的行数)。"""
    res = []
    i, n = 0, len(seg)
    extra = 0
    while i < n:
        c = seg[i]
        # \[ ... \] 显示公式
        if c == "\\" and i + 1 < n and seg[i + 1] == "[":
            j = seg.find("\\]", i + 2)
            if j == -1:
                inner, extra, closed, tail = _collect_across_lines(
                    lines, idx, seg, i + 2, "\\]")
                if closed and _LATEX_CMD_RE.search(inner):
                    res.append("$$" + inner + "$$")
                    res.append(tail)
                    return "".join(res), extra
                res.append(seg[i:])
                return "".join(res), extra
            if _LATEX_CMD_RE.search(seg[i + 2:j]):
                res.append("$$" + seg[i + 2:j] + "$$")
                i = j + 2
                continue
            res.append(c)
            i += 1
            continue
        # \( ... \) 行内公式
        if c == "\\" and i + 1 < n and seg[i + 1] == "(":
            j = seg.find("\\)", i + 2)
            if j != -1 and _LATEX_CMD_RE.search(seg[i + 2:j]):
                res.append("$" + seg[i + 2:j] + "$")
                i = j + 2
                continue
            res.append(c)
            i += 1
            continue
        # [ ... ] 显示公式（先做括号配对，兼容内容中的 [ ]）
        if c == "[":
            depth, j = 1, i + 1
            while j < n and depth:
                if seg[j] == "[":
                    depth += 1
                elif seg[j] == "]":
                    depth -= 1
                j += 1
            if depth == 0:
                inner = seg[i + 1:j - 1]
                # ] 后紧跟 ( 视为 Markdown 链接，跳过
                is_link = seg[j:].lstrip().startswith("(")
                if _LATEX_CMD_RE.search(inner) and not is_link:
                    res.append("$$" + inner + "$$")
                    i = j
                    continue
                res.append(c)
                i += 1
                continue
            # 本行未闭合 → 尝试跨行
            inner, extra, closed, tail = _collect_across_lines(
                lines, idx, seg, i + 1, "]")
            if closed and _LATEX_CMD_RE.search(inner):
                res.append("$$" + inner + "$$")
                res.append(tail)
                return "".join(res), extra
            res.append(c)
            i += 1
            continue
        res.append(c)
        i += 1
    return "".join(res), extra


def _convert_line(lines, idx):
    """转换一行（跳过行内代码），返回 (新行, 额外消费的行数)。"""
    line = lines[idx]
    parts = line.split("`")  # 奇数下标为行内代码段，不改写
    extra = 0
    for k in range(0, len(parts), 2):
        new_seg, used = _convert_segment(parts[k], lines, idx)
        parts[k] = new_seg
        extra = max(extra, used)
    return "`".join(parts), extra


def convert_bracket_math(text):
    """把 [ ... ] / \\[ ... \\] / \\( ... \\) 形式的 LaTeX 公式改写为标准定界符。

    仅当内容包含 LaTeX 命令时改写；自动跳过围栏代码块（``` / ~~~）。
    返回改写后的全文；未命中任何改写时原样返回。
    """
    lines = text.split("\n")
    out_lines = []
    fence = None
    idx = 0
    while idx < len(lines):
        line = lines[idx]
        s = line.strip()
        if s.startswith("```") or s.startswith("~~~"):
            fc = s[:3]
            if fence == fc:
                fence = None
            elif fence is None:
                fence = fc
            out_lines.append(line)
            idx += 1
            continue
        if fence:
            out_lines.append(line)
            idx += 1
            continue
        new_line, used = _convert_line(lines, idx)
        out_lines.append(new_line)
        idx += 1 + used
    return "\n".join(out_lines)


def safe_input(prompt=""):
    """带 EOF 保护的 input，管道/异常场景不会崩溃。"""
    try:
        return input(prompt)
    except EOFError:
        return ""


def find_pandoc():
    """优先 PATH，其次常见安装位置。"""
    found = shutil.which("pandoc")
    if found:
        return found
    for cand in PANDOC_CANDIDATES:
        p = os.path.expandvars(cand)
        if os.path.isfile(p):
            return p
    return None


def check_omml(docx_path):
    """统计 docx 中的原生公式数量：行内 oMath 与独立 oMathPara。"""
    inline = display = 0
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/document.xml").decode("utf-8", "replace")
        display = xml.count("<m:oMathPara>")
        # 独立公式（oMathPara）内部也含 <m:oMath>，故行内数 = 总数 - 独立数
        inline = max(0, xml.count("<m:oMath>") - display)
    except Exception:
        pass
    return inline, display


def convert(md_path, pandoc):
    """转换单个 md 文件，返回 (是否成功, 提示信息)。"""
    md_path = os.path.abspath(md_path)
    if not os.path.isfile(md_path):
        return False, "找不到文件：%s" % md_path
    if not md_path.lower().endswith(MD_EXT):
        return False, "不是 Markdown 文件：%s" % md_path

    out_path = os.path.splitext(md_path)[0] + ".docx"

    # 预处理 [ ... ] / \[ ... \] 等非标准公式定界符（不改动源文件）
    src_path = md_path
    tmp_fd = None
    try:
        with open(md_path, "r", encoding="utf-8") as f:
            content = f.read()
        processed = convert_bracket_math(content)
        if processed != content:
            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".md")
            with os.fdopen(tmp_fd, "w", encoding="utf-8") as f:
                f.write(processed)
            tmp_fd = None
            src_path = tmp_path
    except UnicodeDecodeError:
        # 非 UTF-8 编码的 md 文件直接交给 pandoc 处理
        pass

    cmd = [pandoc, src_path, "-o", out_path]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
    except Exception as e:
        return False, "运行 pandoc 失败：%s" % e
    finally:
        if tmp_fd is not None:
            try:
                os.close(tmp_fd)
            except Exception:
                pass
        if src_path != md_path:
            try:
                os.remove(src_path)
            except Exception:
                pass

    if r.returncode != 0:
        detail = r.stderr.strip() or r.stdout.strip() or "未知错误"
        hint = ""
        if "permission denied" in detail.lower():
            hint = "\n（目标 .docx 可能正被 Word/WPS 打开，请先关闭再重试）"
        return False, "转换失败：\n%s%s" % (detail, hint)

    inline, display = check_omml(out_path)
    size_kb = os.path.getsize(out_path) / 1024
    return True, "成功 → %s（%.1f KB，行内公式 %d 个，独立公式 %d 个）" % (
        out_path, size_kb, inline, display)


def collect_local_md():
    """收集当前目录与桌面的 md 文件供交互选择。"""
    here = os.path.abspath(".")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    files, seen = [], set()
    for folder in (here, desktop):
        for pat in ("*.md", "*.markdown", "*.mdown", "*.mkd"):
            for f in sorted(glob.glob(os.path.join(folder, pat))):
                key = os.path.normcase(f)
                if key not in seen:
                    seen.add(key)
                    files.append(f)
    return files


def interactive(pandoc):
    """无参数时的交互模式。"""
    print("=" * 60)
    print("  %s  v%s" % (APP_NAME, VERSION))
    print("  Markdown → Word，公式转为 Word 原生公式")
    print("=" * 60)
    print()
    print("[1] 把 .md 文件拖到“一键转换.bat”上即可直接转换；")
    print("    也可在下方输入文件完整路径（支持一次输入多个，用空格分隔）。")
    print()

    local = collect_local_md()
    if local:
        print("检测到以下 md 文件，输入序号即可选择：")
        for i, f in enumerate(local, 1):
            print("  [%2d] %s" % (i, f))
        print()

    while True:
        raw = safe_input("请输入文件路径 / 序号（直接回车退出）：").strip()
        if not raw:
            break
        targets = []
        # 先尝试把每个空白分隔段当作路径，再尝试整体当作序号
        parts = [p.strip('"') for p in raw.split()]
        if len(parts) == 1 and parts[0].isdigit():
            idx = int(parts[0])
            if 1 <= idx <= len(local):
                targets = [local[idx - 1]]
            else:
                print("序号超出范围。")
                continue
        else:
            targets = parts
        if not targets:
            continue
        print()
        run_batch(targets, pandoc)
        print()


def run_batch(targets, pandoc):
    """批量转换并汇总。"""
    ok_cnt = 0
    for t in targets:
        ok, msg = convert(t, pandoc)
        if ok:
            ok_cnt += 1
            print("  ✔ " + msg)
        else:
            print("  ✘ " + msg)
    print()
    print("完成：成功 %d / 共 %d 个文件。" % (ok_cnt, len(targets)))
    print("-" * 60)


def main():
    args = sys.argv[1:]
    pandoc = find_pandoc()
    if not pandoc:
        print("未找到 pandoc，无法转换。")
        print("请先安装 pandoc：https://pandoc.org/installing.html")
        print("（安装后需重启本工具）")
        safe_input("\n按回车键退出...")
        return 1

    print("已找到 pandoc：%s" % pandoc)
    print()

    if args:
        targets = [a.strip('"') for a in args if a.strip()]
        run_batch(targets, pandoc)
    else:
        interactive(pandoc)

    safe_input("按回车键退出...")
    return 0


if __name__ == "__main__":
    sys.exit(main())
