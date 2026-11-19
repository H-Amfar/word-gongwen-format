# -*- coding: utf-8 -*-
"""
公文 Word 格式化脚本
按公文排版规范对 .docx 文件套用字体、字号与段落格式。

规范：
- 标题一（Heading 1）：方正大标宋简体，字号按标题字数自适应
- 标题二（Heading 2）：方正楷体GBK，三号 16pt
- 标题三（Heading 3）：方正仿宋GB2312，四号 14pt
- 正文（Normal）：方正仿宋GB2312，三号 16pt
- 表格内正文：方正仿宋GB2312，四号 14pt
- 正文段落：首行缩进2字符、段前8磅、段后8磅、行距1.5倍、两端对齐、黑色字体

用法：
    python apply_gongwen_format.py <input.docx> [--output <output.docx>]

依赖：pip install python-docx
"""
import argparse
import os
import sys

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("缺少依赖 python-docx，请先执行: pip install python-docx")

FONT_H1 = "方正大标宋简体"
FONT_H2 = "方正楷体GBK"
FONT_BODY = "方正仿宋GB2312"

SIZE_H2 = Pt(16)   # 三号
SIZE_H3 = Pt(14)   # 四号
SIZE_BODY = Pt(16) # 三号
SIZE_TABLE = Pt(14)  # 四号
BLACK = RGBColor(0, 0, 0)


def h1_size_by_length(text):
    """根据标题字数自适应标题一字号。"""
    n = len(text.strip())
    if n <= 15:
        return Pt(22)   # 二号
    elif n <= 20:
        return Pt(18)   # 小二
    elif n <= 30:
        return Pt(16)   # 三号
    else:
        return Pt(18)   # 小二，允许折行


def set_run_font(run, font_name, size, color=BLACK, bold=False):
    """设置 run 的中西文字体、字号、颜色、加粗。"""
    run.font.name = font_name
    # eastAsia 保证中文字体生效
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    if size is not None:
        run.font.size = size
    run.font.color.rgb = color
    run.font.bold = bold


def set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_line_indent_chars=2,
                         space_before=Pt(8), space_after=Pt(8),
                         line_spacing=1.5):
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = space_before
    pf.space_after = space_after
    pf.line_spacing = line_spacing
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if first_line_indent_chars:
        # 用字符单位保证"首行缩进2字符"
        pPr = p._element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = pPr.makeelement(qn("w:ind"), {})
            pPr.append(ind)
        ind.set(qn("w:firstLineChars"), str(first_line_indent_chars * 100))
        ind.set(qn("w:firstLine"), str(int(2 * 16 * 20)))  # 兜底：2字符*16pt*20(twips)


def format_heading(doc, style_names, font_name, size_fn):
    for name in style_names:
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        # 样式级字体
        style.font.name = font_name
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = rPr.makeelement(qn("w:rFonts"), {})
            rPr.append(rFonts)
        rFonts.set(qn("w:eastAsia"), font_name)
        rFonts.set(qn("w:ascii"), font_name)
        rFonts.set(qn("w:hAnsi"), font_name)
        style.font.color.rgb = BLACK
        style.font.bold = True  # 标题加粗
        # 逐段处理（标题一需按字数自适应字号）
        for p in doc.paragraphs:
            if p.style.name == name:
                size = size_fn(p.text) if callable(size_fn) else size_fn
                for run in p.runs:
                    set_run_font(run, font_name, size, bold=True)
                if callable(size_fn):
                    set_paragraph_format(
                        p,
                        align=WD_ALIGN_PARAGRAPH.CENTER,
                        first_line_indent_chars=0,
                        space_before=Pt(8), space_after=Pt(8),
                        line_spacing=1.5)
                else:
                    set_paragraph_format(p)


def format_body(doc):
    try:
        style = doc.styles["Normal"]
    except KeyError:
        return
    style.font.name = FONT_BODY
    style.font.size = SIZE_BODY
    style.font.color.rgb = BLACK
    for p in doc.paragraphs:
        # 只处理正文样式或无样式的普通段落，跳过标题
        if p.style.name.startswith("Heading") or p.style.name in ("Title",):
            continue
        for run in p.runs:
            set_run_font(run, FONT_BODY, SIZE_BODY)
        set_paragraph_format(p)


def format_tables(doc):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        set_run_font(run, FONT_BODY, SIZE_TABLE)
                    set_paragraph_format(p, first_line_indent_chars=0)


def main():
    parser = argparse.ArgumentParser(description="公文 Word 格式化工具")
    parser.add_argument("input", help="输入 .docx 文件路径")
    parser.add_argument("--output", help="输出 .docx 文件路径（默认原文件名后加 _公文格式）")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        sys.exit(f"文件不存在: {args.input}")

    base, ext = os.path.splitext(args.input)
    output = args.output or f"{base}_公文格式{ext}"

    doc = Document(args.input)

    # 标题一：字号按字数自适应
    format_heading(doc, ["Heading 1", "标题 1"], FONT_H1, h1_size_by_length)
    # 标题二
    format_heading(doc, ["Heading 2", "标题 2"], FONT_H2, SIZE_H2)
    # 标题三
    format_heading(doc, ["Heading 3", "标题 3"], FONT_BODY, SIZE_H3)
    # 正文
    format_body(doc)
    # 表格
    format_tables(doc)

    doc.save(output)
    print(f"已完成公文格式化: {output}")


if __name__ == "__main__":
    main()
