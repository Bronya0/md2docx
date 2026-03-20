"""
通用 Markdown → Word (.docx) 转换工具

用法:
    python md2docx.py input.md                  # 输出为 input.docx
    python md2docx.py input.md -o output.docx   # 指定输出文件名

支持的 Markdown 特性:
    - 标题 (h1-h6)
    - 粗体 / 斜体 / 粗斜体 / 删除线 / 行内代码
    - 链接 / 图片（本地图片插入，远程图片显示占位文本）
    - 有序列表 / 无序列表 / 多级嵌套列表
    - 表格（带蓝色表头）
    - 代码块（带语言标签，灰底等宽字体）
    - 引用块 (blockquote)
    - 水平分隔线
    - 脚注（转为尾注文本）
    - HTML 注释自动忽略
"""

import argparse
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

# ── 样式常量 ──
FONT_NAME_CN = "宋体"
FONT_NAME_EN = "Times New Roman"
FONT_CODE = "Consolas"
FONT_SIZE_BODY = Pt(12)  # 小四
FONT_SIZE_TABLE = Pt(10.5)  # 五号
FONT_SIZE_CODE = Pt(9)
FONT_SIZE_FOOTNOTE = Pt(9)
HEADING_SIZES = {
    1: Pt(22),  # 二号
    2: Pt(18),  # 小二
    3: Pt(15),  # 小三
    4: Pt(13),  # 四号
    5: Pt(12),  # 小四
    6: Pt(10.5),  # 五号
}
TABLE_HEADER_BG = "4472C4"
BLOCKQUOTE_BG = "F0F4F8"
BLOCKQUOTE_BORDER = "4472C4"
CODE_BLOCK_BG = "F5F5F5"

# ── 内联 Markdown 正则 ──
# 顺序很重要：先匹配更长/更特殊的模式
INLINE_PATTERNS = [
    # 图片 ![alt](url)
    ("image", re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")),
    # 链接 [text](url)
    ("link", re.compile(r"\[([^\]]+)\]\(([^)]+)\)")),
    # 粗斜体 ***text*** 或 ___text___
    ("bold_italic", re.compile(r"\*\*\*(.+?)\*\*\*|___(.+?)___")),
    # 粗体 **text** 或 __text__
    ("bold", re.compile(r"\*\*(.+?)\*\*|__(.+?)__")),
    # 斜体 *text* 或 _text_（不贪婪，避免与粗体冲突）
    (
        "italic",
        re.compile(
            r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)"
        ),
    ),
    # 删除线 ~~text~~
    ("strikethrough", re.compile(r"~~(.+?)~~")),
    # 行内代码 `code`
    ("code", re.compile(r"`([^`]+)`")),
    # 脚注引用 [^id]
    ("footnote_ref", re.compile(r"\[\^(\w+)\]")),
]


# ═══════════════════════════════════════════════════════
#  XML / OPC 辅助
# ═══════════════════════════════════════════════════════


def set_cell_shading(cell, color_hex):
    tc_pr = cell._element.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color_hex)
    shading.set(qn("w:val"), "clear")
    tc_pr.append(shading)


def set_paragraph_shading(paragraph, color_hex):
    p_pr = paragraph._element.get_or_add_pPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color_hex)
    shading.set(qn("w:val"), "clear")
    p_pr.append(shading)


def set_paragraph_left_border(paragraph, color_hex, width_pt=3):
    """给段落加左边框（用于 blockquote）"""
    p_pr = paragraph._element.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(width_pt * 8))  # 1/8 pt 单位
    left.set(qn("w:space"), "4")
    left.set(qn("w:color"), color_hex)
    borders.append(left)
    p_pr.append(borders)


def add_horizontal_line(doc):
    """添加水平分隔线"""
    p = doc.add_paragraph()
    p_pr = p._element.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "999999")
    borders.append(bottom)
    p_pr.append(borders)


# ═══════════════════════════════════════════════════════
#  字体 / Run 辅助
# ═══════════════════════════════════════════════════════


def set_run_font(
    run, bold=False, italic=False, strike=False, size=None, color=None, font_name=None
):
    run.font.name = font_name or FONT_NAME_EN
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:eastAsia"), FONT_NAME_CN)
    if font_name:
        r_fonts.set(qn("w:ascii"), font_name)
        r_fonts.set(qn("w:hAnsi"), font_name)
    if size:
        run.font.size = size
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    if strike:
        run.font.strike = True
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_inline_runs(
    paragraph, text, base_size=None, base_bold=False, base_italic=False, footnotes=None
):
    """
    将 Markdown 内联标记解析为 Word Run 对象。
    支持: **bold**, *italic*, ***bold_italic***, ~~strike~~,
           `code`, [link](url), ![img](path), [^fn]
    """
    if base_size is None:
        base_size = FONT_SIZE_BODY
    if footnotes is None:
        footnotes = {}

    # 统一正则：把所有内联标记提取出来
    combined = re.compile(
        r"(!\[[^\]]*\]\([^)]+\))"  # image
        r"|(\[[^\]]+\]\([^)]+\))"  # link
        r"|(\*\*\*.+?\*\*\*)"  # bold_italic
        r"|(\*\*.+?\*\*)"  # bold
        r"|(\*(?!\*).+?(?<!\*)\*)"  # italic (single *)
        r"|(~~.+?~~)"  # strikethrough
        r"|(`[^`]+`)"  # inline code
        r"|(\[\^\w+\])"  # footnote ref
    )

    pos = 0
    for m in combined.finditer(text):
        # 纯文本部分
        if m.start() > pos:
            run = paragraph.add_run(text[pos : m.start()])
            set_run_font(run, bold=base_bold, italic=base_italic, size=base_size)

        matched = m.group(0)

        if matched.startswith("!["):
            # 图片
            img_m = re.match(r"!\[([^\]]*)\]\(([^)]+)\)", matched)
            alt, src = img_m.group(1), img_m.group(2)  # type: ignore[union-attr]
            _try_insert_image(paragraph, src, alt)

        elif matched.startswith("[^"):
            # 脚注引用
            fn_id = matched[2:-1]
            run = paragraph.add_run(f"[{fn_id}]")
            set_run_font(run, size=FONT_SIZE_FOOTNOTE, color="666666")

        elif matched.startswith("["):
            # 链接
            link_m = re.match(r"\[([^\]]+)\]\(([^)]+)\)", matched)
            link_text = link_m.group(1)  # type: ignore[union-attr]
            run = paragraph.add_run(link_text)
            set_run_font(run, size=base_size, color="2E75B6")
            run.underline = True

        elif matched.startswith("***"):
            inner = matched[3:-3]
            run = paragraph.add_run(inner)
            set_run_font(run, bold=True, italic=True, size=base_size)

        elif matched.startswith("**"):
            inner = matched[2:-2]
            run = paragraph.add_run(inner)
            set_run_font(run, bold=True, italic=base_italic, size=base_size)

        elif matched.startswith("~~"):
            inner = matched[2:-2]
            run = paragraph.add_run(inner)
            set_run_font(run, strike=True, bold=base_bold, size=base_size)

        elif matched.startswith("`"):
            inner = matched[1:-1]
            run = paragraph.add_run(inner)
            set_run_font(run, size=base_size, font_name=FONT_CODE)
            # 行内代码灰底
            rpr = run._element.get_or_add_rPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:fill"), "F0F0F0")
            shd.set(qn("w:val"), "clear")
            rpr.append(shd)

        elif matched.startswith("*"):
            inner = matched[1:-1]
            run = paragraph.add_run(inner)
            set_run_font(run, italic=True, bold=base_bold, size=base_size)

        pos = m.end()

    # 剩余纯文本
    if pos < len(text):
        run = paragraph.add_run(text[pos:])
        set_run_font(run, bold=base_bold, italic=base_italic, size=base_size)


def _try_insert_image(paragraph, src, alt):
    """尝试插入本地图片，失败则显示占位文本"""
    src_path = Path(src)
    if src_path.is_file():
        try:
            run = paragraph.add_run()
            run.add_picture(str(src_path), width=Inches(5))
            return
        except Exception:
            pass
    # 占位文本
    run = paragraph.add_run(f"[图片: {alt or src}]")
    set_run_font(run, italic=True, size=FONT_SIZE_BODY, color="999999")


# ═══════════════════════════════════════════════════════
#  元素构建
# ═══════════════════════════════════════════════════════


def add_heading(doc, text, level, footnotes=None):
    h = doc.add_heading(level=min(level, 4))
    # 清除默认 run
    for r in h.runs:
        r.text = ""
    add_inline_runs(
        h,
        text,
        base_size=HEADING_SIZES.get(level, Pt(12)),
        base_bold=True,
        footnotes=footnotes,
    )


def add_body_paragraph(doc, text, footnotes=None):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    add_inline_runs(p, text, footnotes=footnotes)
    return p


def add_table(doc, headers, rows):
    col_count = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=col_count)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    # 自动列宽
    table.autofit = True

    # 表头
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        set_run_font(run, bold=True, size=FONT_SIZE_TABLE, color="FFFFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, TABLE_HEADER_BG)

    # 数据行
    for r_idx, row in enumerate(rows):
        for c_idx in range(col_count):
            cell = table.rows[r_idx + 1].cells[c_idx]
            val = row[c_idx] if c_idx < len(row) else ""
            cell.text = ""
            # 表格单元格也支持内联标记
            add_inline_runs(cell.paragraphs[0], val, base_size=FONT_SIZE_TABLE)

    doc.add_paragraph()  # 表后空行


def add_code_block(doc, code_text, language=""):
    """添加代码块：带灰底的等宽字体段落"""
    # 语言标签
    if language:
        lang_p = doc.add_paragraph()
        run = lang_p.add_run(f"  {language}")
        set_run_font(run, size=Pt(8), color="888888", font_name=FONT_CODE)

    for line in code_text.split("\n"):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        # 段落左缩进
        p.paragraph_format.left_indent = Cm(0.5)
        set_paragraph_shading(p, CODE_BLOCK_BG)
        run = p.add_run(line if line else " ")  # 空行保留
        set_run_font(run, size=FONT_SIZE_CODE, font_name=FONT_CODE)

    doc.add_paragraph()  # 代码块后空行


def add_blockquote(doc, lines, footnotes=None):
    """添加引用块"""
    for line in lines:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.left_indent = Cm(1)
        set_paragraph_shading(p, BLOCKQUOTE_BG)
        set_paragraph_left_border(p, BLOCKQUOTE_BORDER)
        text = line.lstrip("> ").strip()
        add_inline_runs(p, text, base_italic=True, footnotes=footnotes)


def add_list_item(doc, text, level=0, ordered=False, number=1, footnotes=None):
    """添加列表项（支持嵌套）"""
    if ordered:
        # 有序列表：手动编号
        p = doc.add_paragraph()
        indent = Cm(1.27 * (level + 1))
        p.paragraph_format.left_indent = indent
        p.paragraph_format.first_line_indent = Cm(-0.63)
        p.paragraph_format.space_after = Pt(2)
        # 编号
        run = p.add_run(f"{number}. ")
        set_run_font(run, size=FONT_SIZE_BODY)
        add_inline_runs(p, text, footnotes=footnotes)
    else:
        # 无序列表
        style_name = "List Bullet"
        if level == 1:
            style_name = "List Bullet 2"
        elif level >= 2:
            style_name = "List Bullet 3"
        try:
            p = doc.add_paragraph(style=style_name)
        except KeyError:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.left_indent = Cm(1.27 * (level + 1))
        p.paragraph_format.space_after = Pt(2)
        add_inline_runs(p, text, footnotes=footnotes)


# ═══════════════════════════════════════════════════════
#  Markdown 解析器
# ═══════════════════════════════════════════════════════


def _detect_list_indent(line):
    """检测列表缩进级别（每2或4空格为一级）"""
    stripped = line.rstrip()
    spaces = len(stripped) - len(stripped.lstrip())
    level = spaces // 2  # 2 空格 = 1 级
    return level


def parse_md(md_text):
    """
    结构化 Markdown 解析器。
    返回 (blocks, footnotes) 元组。
    blocks 为块元素列表，footnotes 为脚注字典。
    """
    blocks = []
    footnotes = {}
    lines = md_text.split("\n")
    i = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # ── HTML 注释 ──
        if stripped.startswith("<!--"):
            while i < len(lines) and "-->" not in lines[i]:
                i += 1
            i += 1
            continue

        # ── 脚注定义 [^id]: text ──
        fn_m = re.match(r"^\[\^(\w+)\]:\s*(.*)", stripped)
        if fn_m:
            fn_id, fn_text = fn_m.group(1), fn_m.group(2)
            # 可能多行
            i += 1
            while i < len(lines) and lines[i].startswith("  "):
                fn_text += " " + lines[i].strip()
                i += 1
            footnotes[fn_id] = fn_text
            continue

        # ── 代码块 ── (匹配 ```language)
        if stripped.startswith("```"):
            language = stripped[3:].strip()
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            blocks.append(("code", "\n".join(code_lines), language))
            i += 1  # 跳过结束 ```
            continue

        # ── 标题 ──
        h_m = re.match(r"^(#{1,6})\s+(.*)", line)
        if h_m:
            level = len(h_m.group(1))
            blocks.append(("heading", level, h_m.group(2).strip()))
            i += 1
            continue

        # ── Setext 标题 (下划线样式) ──
        if (
            i + 1 < len(lines)
            and stripped
            and re.match(r"^[=]{3,}$", lines[i + 1].strip())
        ):
            blocks.append(("heading", 1, stripped))
            i += 2
            continue
        if (
            i + 1 < len(lines)
            and stripped
            and re.match(r"^[-]{3,}$", lines[i + 1].strip())
            and not re.match(r"^[-*]\s", stripped)  # 排除列表紧跟分隔线
        ):
            blocks.append(("heading", 2, stripped))
            i += 2
            continue

        # ── 表格 ──
        if (
            "|" in line
            and i + 1 < len(lines)
            and re.match(r"^\|[\s\-:|]+\|", lines[i + 1])
        ):
            headers = [c.strip() for c in stripped.strip("|").split("|")]
            i += 2  # 跳过分隔行
            rows = []
            while (
                i < len(lines) and "|" in lines[i] and lines[i].strip().startswith("|")
            ):
                row = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                rows.append(row)
                i += 1
            blocks.append(("table", headers, rows))
            continue

        # ── 水平线 ──
        if re.match(r"^(\*{3,}|-{3,}|_{3,})$", stripped):
            blocks.append(("hr",))
            i += 1
            continue

        # ── 引用块 ──
        if stripped.startswith(">"):
            quote_lines = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                quote_lines.append(lines[i].strip())
                i += 1
            blocks.append(("blockquote", quote_lines))
            continue

        # ── 有序列表 ──
        ol_m = re.match(r"^(\s*)(\d+)[.)]\s+(.*)", line)
        if ol_m:
            list_items = []
            while i < len(lines):
                ol_line = re.match(r"^(\s*)(\d+)[.)]\s+(.*)", lines[i])
                if ol_line:
                    level = _detect_list_indent(lines[i])
                    num = int(ol_line.group(2))
                    list_items.append((level, num, ol_line.group(3).strip()))
                    i += 1
                elif i < len(lines) and lines[i].startswith("  ") and list_items:
                    # 续行（缩进的后续内容）
                    prev = list_items[-1]
                    list_items[-1] = (
                        prev[0],
                        prev[1],
                        prev[2] + " " + lines[i].strip(),
                    )
                    i += 1
                else:
                    break
            blocks.append(("ordered_list", list_items))
            continue

        # ── 无序列表 ──
        ul_m = re.match(r"^(\s*)[-*+]\s+(.*)", line)
        if ul_m:
            list_items = []
            while i < len(lines):
                ul_line = re.match(r"^(\s*)[-*+]\s+(.*)", lines[i])
                if ul_line:
                    level = _detect_list_indent(lines[i])
                    list_items.append((level, ul_line.group(2).strip()))
                    i += 1
                elif i < len(lines) and lines[i].startswith("  ") and list_items:
                    prev = list_items[-1]
                    list_items[-1] = (prev[0], prev[1] + " " + lines[i].strip())
                    i += 1
                else:
                    break
            blocks.append(("unordered_list", list_items))
            continue

        # ── 普通文本（合并相邻的非空行为一个段落）──
        if stripped:
            para_lines = [stripped]
            i += 1
            while (
                i < len(lines)
                and lines[i].strip()
                and not lines[i].strip().startswith("#")
                and not lines[i].strip().startswith(">")
                and not lines[i].strip().startswith("```")
                and not lines[i].strip().startswith("|")
                and not re.match(r"^(\*{3,}|-{3,}|_{3,})$", lines[i].strip())
                and not re.match(r"^[-*+]\s", lines[i].strip())
                and not re.match(r"^\d+[.)]\s", lines[i].strip())
                and not re.match(r"^\[\^", lines[i].strip())
            ):
                para_lines.append(lines[i].strip())
                i += 1
            blocks.append(("text", " ".join(para_lines)))
            continue

        i += 1  # 空行

    return blocks, footnotes


# ═══════════════════════════════════════════════════════
#  文档构建
# ═══════════════════════════════════════════════════════


def build_docx(blocks, footnotes):
    doc = Document()

    # 页面设置 A4
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

    for block in blocks:
        btype = block[0]

        if btype == "heading":
            _, level, text = block
            add_heading(doc, text, level, footnotes)

        elif btype == "text":
            add_body_paragraph(doc, block[1], footnotes)

        elif btype == "unordered_list":
            for level, text in block[1]:
                add_list_item(
                    doc, text, level=level, ordered=False, footnotes=footnotes
                )

        elif btype == "ordered_list":
            for level, num, text in block[1]:
                add_list_item(
                    doc,
                    text,
                    level=level,
                    ordered=True,
                    number=num,
                    footnotes=footnotes,
                )

        elif btype == "table":
            _, headers, rows = block
            add_table(doc, headers, rows)

        elif btype == "code":
            _, code_text, language = block
            add_code_block(doc, code_text, language)

        elif btype == "blockquote":
            add_blockquote(doc, block[1], footnotes)

        elif btype == "hr":
            add_horizontal_line(doc)

    # ── 脚注附录 ──
    if footnotes:
        doc.add_paragraph()
        add_horizontal_line(doc)
        h = doc.add_heading(level=3)
        run = h.add_run("注释")
        set_run_font(run, bold=True, size=Pt(13))
        for fn_id, fn_text in footnotes.items():
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(f"[{fn_id}] ")
            set_run_font(run, bold=True, size=FONT_SIZE_FOOTNOTE, color="666666")
            add_inline_runs(p, fn_text, base_size=FONT_SIZE_FOOTNOTE)

    return doc


# ═══════════════════════════════════════════════════════
#  CLI 入口
# ═══════════════════════════════════════════════════════


def main():
    parser = argparse.ArgumentParser(
        description="Markdown → Word (.docx) 转换工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument(
        "-o", "--output", help="输出 Word 文件路径（默认与输入同名 .docx）"
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.is_file():
        print(f"错误: 找不到文件 {input_path}")
        return

    output_path = Path(args.output) if args.output else input_path.with_suffix(".docx")

    md_text = input_path.read_text(encoding="utf-8")
    blocks, footnotes = parse_md(md_text)
    doc = build_docx(blocks, footnotes)
    doc.save(str(output_path))
    print(f"已生成: {output_path}")


if __name__ == "__main__":
    main()
