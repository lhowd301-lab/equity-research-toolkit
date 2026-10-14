# -*- coding: utf-8 -*-
"""生成《创元科技因子归因分析》Word 报告。"""
import os

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DOCX = os.path.join(BASE_DIR, "output", "创元科技_因子归因分析报告_20260806.docx")
IMG_DIR = os.path.join(BASE_DIR, "output", "图片")

FONT = "Microsoft YaHei"
INK = "0B2545"
BLUE = "2E74B5"
DARK_BLUE = "1F4D78"
MUTED = "595959"
HEADER_FILL = "F2F4F7"
CALLOUT_FILL = "F4F6F9"
BORDER = "D8DEE6"


def set_run_font(run, size=11, bold=False, italic=False, color=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rpr.rFonts.set(qn("w:ascii"), FONT)
    rpr.rFonts.set(qn("w:hAnsi"), FONT)
    rpr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)


def set_style_font(style, size=None, color=None, bold=None):
    style.font.name = FONT
    rpr = style.element.get_or_add_rPr()
    rpr.rFonts.set(qn("w:ascii"), FONT)
    rpr.rFonts.set(qn("w:hAnsi"), FONT)
    rpr.rFonts.set(qn("w:eastAsia"), FONT)
    if size is not None:
        style.font.size = Pt(size)
    if color is not None:
        style.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        style.font.bold = bold


def add_para(doc, text="", size=11, bold=False, italic=False, color=None,
             align=None, before=None, after=None, line=None, keep_next=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    if line is not None:
        p.paragraph_format.line_spacing = line
    p.paragraph_format.keep_with_next = keep_next
    if text:
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold, italic=italic, color=color)
    return p


def add_rich_para(doc, parts, align=None, before=None, after=None, line=1.15):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = line
    for text, kwargs in parts:
        run = p.add_run(text)
        set_run_font(run, **kwargs)
    return p


def set_cell_bg(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_table_borders(table, color=BORDER, sz=4):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn("w:" + edge))
        if el is None:
            el = OxmlElement("w:" + edge)
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def set_table_cell_margins(table, top=80, bottom=80, start=120, end=120):
    tbl_pr = table._tbl.tblPr
    mar = tbl_pr.find(qn("w:tblCellMar"))
    if mar is None:
        mar = OxmlElement("w:tblCellMar")
        tbl_pr.append(mar)
    for edge, val in (("top", top), ("bottom", bottom), ("start", start), ("end", end)):
        el = mar.find(qn("w:" + edge))
        if el is None:
            el = OxmlElement("w:" + edge)
            mar.append(el)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")


def fix_table_geometry(table, widths_inches):
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    tbl = table._tbl
    tbl_pr = tbl.tblPr

    total_dxa = int(sum(widths_inches) * 1440)
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total_dxa))
    tbl_w.set(qn("w:type"), "dxa")

    ind = tbl_pr.find(qn("w:tblInd"))
    if ind is None:
        ind = OxmlElement("w:tblInd")
        tbl_pr.append(ind)
    ind.set(qn("w:w"), "120")
    ind.set(qn("w:type"), "dxa")

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for col in list(grid):
            grid.remove(col)
    else:
        grid = OxmlElement("w:tblGrid")
        tbl.insert(1, grid)
    for w in widths_inches:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(int(w * 1440)))
        grid.append(col)

    for row in table.rows:
        for cell, w in zip(row.cells, widths_inches):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(w * 1440)))
            tc_w.set(qn("w:type"), "dxa")


def mark_header_row(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_table(doc, headers, rows, widths, aligns=None, font_size=9.5,
              header_fill=HEADER_FILL):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    set_table_borders(table)
    set_table_cell_margins(table)
    fix_table_geometry(table, widths)
    mark_header_row(table.rows[0])

    col_aligns = aligns or ["left"] * len(headers)
    for j, text in enumerate(headers):
        cell = table.rows[0].cells[j]
        set_cell_bg(cell, header_fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.1
        run = p.add_run(text)
        set_run_font(run, size=font_size, bold=True, color=DARK_BLUE)

    for i, row in enumerate(rows):
        for j, text in enumerate(row):
            cell = table.rows[i + 1].cells[j]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.1
            align = col_aligns[j]
            if align == "center":
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif align == "right":
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(text))
            set_run_font(run, size=font_size)
    return table


def add_caption(doc, text, before=4, after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    run = p.add_run(text)
    set_run_font(run, size=9, bold=False, color=MUTED)
    return p


def add_figure_caption(doc, text, before=4, after=10):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    run = p.add_run(text)
    set_run_font(run, size=9.5, bold=True, color=MUTED)
    return p


def set_picture_alt(run, desc):
    for drawing in run._r.findall(qn("w:drawing")):
        for inline in drawing.findall(qn("wp:inline")) + drawing.findall(qn("wp:anchor")):
            doc_pr = inline.find(qn("wp:docPr"))
            if doc_pr is not None:
                doc_pr.set("descr", desc)
                doc_pr.set("title", desc)


def add_picture(doc, filename, caption, width_in=6.15):
    path = os.path.join(IMG_DIR, filename)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(path, width=Inches(width_in))
    set_picture_alt(run, caption)
    add_figure_caption(doc, caption)
    return p


def add_rule(doc, color=BLUE, sz=12, space=6, before=2, after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p_pr = p._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)
    return p


def add_callout(doc, label, text):
    table = doc.add_table(rows=1, cols=1)
    set_table_borders(table, color="C7D3E0", sz=6)
    set_table_cell_margins(table, top=120, bottom=120, start=160, end=160)
    fix_table_geometry(table, [6.5])
    cell = table.rows[0].cells[0]
    set_cell_bg(cell, CALLOUT_FILL)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p1 = cell.paragraphs[0]
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(4)
    p1.paragraph_format.line_spacing = 1.1
    run = p1.add_run(label)
    set_run_font(run, size=11, bold=True, color=BLUE)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.25
    run2 = p2.add_run(text)
    set_run_font(run2, size=10.5, color="333333")
    add_para(doc, "", after=4)
    return table


def add_page_field(paragraph):
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)
    set_run_font(run, size=9, color=MUTED)


def setup_styles(doc):
    normal = doc.styles["Normal"]
    set_style_font(normal, size=11)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.1

    h1 = doc.styles["Heading 1"]
    set_style_font(h1, size=16, color=BLUE, bold=True)
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(8)
    h1.paragraph_format.keep_with_next = True

    h2 = doc.styles["Heading 2"]
    set_style_font(h2, size=13, color=BLUE, bold=True)
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.keep_with_next = True

    h3 = doc.styles["Heading 3"]
    set_style_font(h3, size=12, color=DARK_BLUE, bold=True)
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    h3.paragraph_format.keep_with_next = True


def setup_section(doc):
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    header = section.header
    hp = header.paragraphs[0]
    hp.text = ""
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = hp.add_run("创元科技（000551.SZ）· 因子归因分析")
    set_run_font(run, size=9, color=MUTED)

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.text = ""
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run1 = fp.add_run("第 ")
    set_run_font(run1, size=9, color=MUTED)
    add_page_field(fp)
    run2 = fp.add_run(" 页")
    set_run_font(run2, size=9, color=MUTED)


def add_title_block(doc):
    add_para(doc, "创元科技（000551.SZ）· 长期持有体检", size=11, bold=True,
             color=BLUE, before=0, after=4)
    add_para(doc, "因子归因分析：收益到底从哪来", size=22, bold=True,
             color=INK, before=0, after=2)
    add_para(doc, "用数据拆开 2014 年 10 月至 2026 年 8 月的每一份超额收益",
             size=13, color=MUTED, before=0, after=8)

    meta = [
        ("分析对象", "创元科技（000551.SZ），后复权日度收益"),
        ("数据区间", "2014-10-20 至 2026-08-05，共 2846 个交易日"),
        ("分析方法", "CAPM / 三因子 / 四因子 OLS 回归，Newey-West 稳健标准误，加法收益归因"),
        ("报告日期", "2026-08-06"),
    ]
    for label, value in meta:
        add_rich_para(doc, [
            (label + "：", {"size": 10.5, "bold": True, "color": INK}),
            (value, {"size": 10.5, "color": "333333"}),
        ], before=0, after=2, line=1.2)
    add_rule(doc, color=BLUE, sz=12, space=6, before=4, after=10)


def add_work_summary(doc):
    doc.add_heading("这次分析做了什么", level=1)
    add_para(doc,
             "整个分析基于日线行情、指数和利率数据：创元科技用后复权日收益（把分红和送转折算进价格，"
             "避免除权跳空影响）；无风险利率用 10 年期国债收益率按 252 个交易日折算成日度；四个因子"
             "分别是市场、规模、价值、红利，全部用真实指数差值构造。多因子模型的共同窗口从 2014-10-20 "
             "开始，因为规模因子依赖的中证 1000 指数从 2014 年 10 月才有数据；市场因子本身可以从 "
             "2010-01 开始算。",
             after=8)
    add_para(doc, "表 1  因子构造一览", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(
        doc,
        ["变量", "怎么构造", "人话解释"],
        [
            ["市场因子 MKT", "沪深300 日收益 - 无风险利率", "大盘涨跌给全市场带来的系统性机会"],
            ["规模因子 SMB", "中证1000 日收益 - 沪深300 日收益", "小盘股相对大盘股的表现"],
            ["价值因子 HML", "国证价值日收益 - 国证成长日收益", "便宜的价值股相对成长股的表现"],
            ["红利因子", "国证红利日收益 - 沪深300 日收益", "高股息股票相对市场的表现"],
            ["个股超额收益", "创元科技后复权日收益 - 无风险利率", "我们真正想解释的对象"],
        ],
        widths=[1.6, 2.7, 2.2],
        aligns=["left", "left", "left"],
    )
    add_caption(doc, "数据来源：data/创元科技_000551_体检数据_20260806.xlsx。", after=8)
    add_para(doc,
             "模型上，先跑单因子 CAPM，再跑三因子（市场、规模、价值），最后加红利形成四因子。日收益存在"
             "异方差和自相关，所以系数标准误用 Newey-West 方法修正，t 值不会虚高。收益归因采用加法分解："
             "把累计超额收益拆成 Alpha + 各因子载荷乘因子累计收益 + 残差，回答“每一块钱超额收益是从哪来的”。"
             "另外还做了 36 个月滚动回归和年度归因，观察风格与解释力随时间的变化。",
             after=8)


def add_core_results(doc):
    doc.add_heading("核心结果：收益从哪来", level=1)

    doc.add_heading("全样本回归：这只股票是什么风格", level=2)
    add_para(doc, "表 2  全样本回归结果", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(
        doc,
        ["指标", "CAPM\n(最长窗口)", "CAPM\n(共同窗口)", "三因子代理", "四因子\n(+红利)"],
        [
            ["样本区间", "2010-01-04\n至 2026-08-05", "2014-10-20\n至 2026-08-05", "2014-10-20\n至 2026-08-05", "2014-10-20\n至 2026-08-05"],
            ["样本数", "3993", "2846", "2846", "2846"],
            ["年化 Alpha", "12.9%", "7.6%", "10.6%", "10.8%"],
            ["t(Alpha)", "1.25", "0.62", "1.10", "1.12"],
            ["Beta 市场", "1.02", "1.01", "1.10", "1.09"],
            ["t(市场)", "25.07", "20.49", "26.98", "25.88"],
            ["Beta 规模", "-", "-", "1.32", "1.30"],
            ["t(规模)", "-", "-", "29.66", "22.61"],
            ["Beta 价值", "-", "-", "0.32", "0.36"],
            ["t(价值)", "-", "-", "5.53", "3.65"],
            ["Beta 红利", "-", "-", "-", "-0.07"],
            ["t(红利)", "-", "-", "-", "-0.40"],
            ["R²", "0.232", "0.232", "0.457", "0.457"],
        ],
        widths=[1.45, 1.30, 1.25, 1.25, 1.25],
        aligns=["left", "center", "center", "center", "center"],
        font_size=9,
    )
    add_caption(doc, "注：系数标准误为 Newey-West (HAC) 稳健标准误；共同窗口便于模型之间横向比较。", after=8)
    add_para(doc,
             "先看风格画像：市场 Beta 约 1.09，说明大盘涨 1% 时，这只股票平均要跟着动 1.09%，比大盘更敏感；"
             "规模载荷 1.30 说明它明显偏小盘；价值载荷 0.36 说明略偏价值；红利因子不显著，加了它解释力也没有"
             "提升。R² 为 0.457，意思是四因子加起来只能解释约 46% 的日收益波动，剩下 54% 是个股自身的原因。"
             "年化 Alpha 10.8% 看着很诱人，但 t 值只有 1.12，远达不到统计显著；换句话说，“真有稳定选股能力”"
             "这个说法暂时站不住脚。",
             after=8)

    doc.add_heading("累计归因：129.8% 的超额收益拆给谁", level=2)
    add_para(doc, "表 3  四因子全样本收益归因（2014-10-20 至 2026-08-05）", size=10,
             bold=True, color=DARK_BLUE, before=2, after=3, keep_next=True)
    add_table(
        doc,
        ["收益来源", "累计贡献", "占比", "人话解释"],
        [
            ["个股自身（Alpha）", "+121.6%", "93.7%", "市场、规模、价值、红利都解释完之后剩下的部分，是最主要的收益来源，但也可能是运气或波动"],
            ["市场 Beta（MKT）", "+47.7%", "36.7%", "跟着大盘上涨拿到的钱；Beta 高，所以牛市里贡献大"],
            ["规模（SMB）", "-36.6%", "-28.2%", "股票偏小盘，而这段时间小盘整体跑输，风格暴露拖了后腿"],
            ["价值（HML）", "-1.2%", "-0.9%", "价值风格贡献几乎可以忽略"],
            ["红利", "-1.7%", "-1.3%", "红利因子贡献很小"],
            ["残差", "≈0", "≈0", "加法分解的对账误差，可忽略"],
        ],
        widths=[1.55, 1.05, 0.95, 2.95],
        aligns=["left", "center", "center", "left"],
    )
    add_caption(doc, "注：占比 = 该来源贡献 / 累计超额收益；各来源贡献相加等于累计超额收益 129.8%。", after=8)
    add_para(doc,
             "这张表是最核心的一张。市场贡献 +47.7%，是因为 Beta 高于 1，市场涨的时候跟着多吃了一点；"
             "规模贡献 -36.6% 是最反直觉的部分：股票自己是小盘风格，但同期小盘因子累计跑输大盘约 36.7%，"
             "所以“持有小盘风格”本身在亏钱，亏掉的钱正好被个股自身表现赚了回来。Alpha 贡献 +121.6% 是模型"
             "解释完后剩下的部分，它既是超额收益的主要来源，也是最需要谨慎对待的部分，因为统计上不显著，"
             "可能只是高波动和阶段性运气。",
             after=8)

    add_picture(doc, "08_因子累计收益与个股超额.png",
                "图 1  因子累计收益与个股超额：小盘因子长期下行，个股超额主要靠自身段")

    doc.add_heading("分时间段：近期发生了什么", level=2)
    add_para(doc, "表 4  四因子模型子样本归因", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(
        doc,
        ["样本", "样本区间", "累计超额", "Alpha 占比", "MKT 占比", "SMB 占比", "R²"],
        [
            ["全样本", "2014-10-20 至 2026-08-05", "+129.8%", "93.7%", "36.7%", "-28.2%", "0.457"],
            ["近10年", "2016-08-05 至 2026-08-05", "+111.5%", "123.9%", "23.9%", "-45.8%", "0.361"],
            ["近5年", "2021-08-05 至 2026-08-05", "+71.9%", "78.9%", "-12.1%", "25.3%", "0.339"],
            ["近3年", "2023-08-07 至 2026-08-05", "+67.6%", "77.0%", "14.3%", "9.3%", "0.378"],
            ["近1年", "2025-08-05 至 2026-08-05", "+31.8%", "75.4%", "35.0%", "-6.8%", "0.326"],
        ],
        widths=[0.7, 1.75, 0.9, 0.85, 0.85, 0.85, 0.6],
        aligns=["left", "left", "center", "center", "center", "center", "center"],
        font_size=9,
    )
    add_caption(doc, "注：占比超过 100% 说明其他来源为负贡献；近 5/3/1 年样本数分别为 1211 / 726 / 243。", after=8)
    add_para(doc,
             "近 5 年累计超额 71.9%，Alpha 贡献占 78.9%，SMB 从全样本的负贡献变成 +25.3%，说明近五年小盘"
             "风格不再拖后腿；但 R² 只有 0.339，比全样本更低，意味着因子更解释不了近期收益。近 3 年、近 1 年"
             "的 R² 在 0.33-0.38，Alpha 占比都在 75% 以上，个股自身因素的重要性还在上升。",
             after=8)


def add_annual_view(doc):
    doc.add_heading("逐年拆解：哪些年份靠什么赚钱", level=1)
    add_para(doc,
             "把四因子系数固定在全样本估计值，再按年把因子收益累计起来，就能看到每一年超额收益的来源。"
             "表格里“价值+红利”是价值因子和红利因子贡献之和，便于阅读。",
             after=8)
    add_para(doc, "表 5  年度收益归因（2015-2026）", size=10, bold=True,
             color=DARK_BLUE, before=2, after=3, keep_next=True)
    annual_rows = [
        ["2015", "+49.2%", "+10.3%", "-4.9%", "+75.1%", "-5.3%", "-25.9%"],
        ["2016", "-16.2%", "+9.6%", "-14.6%", "-16.3%", "+2.4%", "+2.7%"],
        ["2017", "-23.1%", "+10.4%", "+18.3%", "-49.5%", "+3.4%", "-5.7%"],
        ["2018", "-31.2%", "+10.4%", "-33.3%", "-20.8%", "+5.3%", "+7.3%"],
        ["2019", "+17.6%", "+10.4%", "+32.4%", "-8.9%", "-6.1%", "-10.2%"],
        ["2020", "+31.2%", "+10.4%", "+26.0%", "-6.8%", "-12.0%", "+13.6%"],
        ["2021", "+36.2%", "+10.4%", "-7.2%", "+31.4%", "-2.6%", "+4.2%"],
        ["2022", "-13.5%", "+10.3%", "-27.4%", "+1.5%", "+5.1%", "-3.1%"],
        ["2023", "+20.5%", "+10.3%", "-15.1%", "+7.6%", "+5.9%", "+11.8%"],
        ["2024", "+21.0%", "+10.3%", "+15.0%", "-12.1%", "+3.7%", "+4.0%"],
        ["2025", "+44.6%", "+10.4%", "+17.2%", "+12.1%", "-5.4%", "+10.3%"],
        ["2026", "-0.02%", "+6.1%", "+0.8%", "-0.8%", "-3.1%", "-2.9%"],
    ]
    add_table(
        doc,
        ["年份", "个股超额", "Alpha", "市场 MKT", "规模 SMB", "价值+红利", "残差"],
        annual_rows,
        widths=[0.7, 0.95, 0.9, 0.95, 0.95, 1.0, 1.05],
        aligns=["center", "center", "center", "center", "center", "center", "center"],
        font_size=9,
    )
    add_caption(doc, "注：2026 年为截至 2026-08-05 的部分年度；各行贡献相加等于个股超额收益。", after=8)
    add_para(doc,
             "2015 年超额 +49.2%，规模贡献 +75.1% 是绝对主力，那年是小盘行情，但残差 -25.9% 说明还有一段"
             "涨跌是因子没解释的；2021 年超额 +36.2%，主要靠规模 +31.4%；2025 年超额 +44.6%，市场、规模、"
             "Alpha 和残差都有正贡献，是“哪都出一点力”的年份；2026 年到 8 月 5 日基本打平（-0.02%），各来源"
             "互相抵消。整体看，Alpha 每年都在 10% 上下，但它只是点估计，单年波动很大，不能当成每年的固定收益。",
             after=8)
    add_picture(doc, "11_年度收益归因.png", "图 4  年度收益归因：每年超额收益的来源拆解")


def add_rolling_view(doc):
    doc.add_heading("滚动视角：风格会变，解释力会降", level=1)
    add_para(doc,
             "36 个月滚动回归显示，2017-2019 年 R² 一度高达 0.60-0.70，因子几乎能讲清楚大部分波动；"
             "2021 年以后降到 0.28-0.42，个股自身因素占比明显上升。市场 Beta 从早期的 1.3 左右逐步降到 "
             "2026 年 8 月的 0.71，说明近三年它对大盘的敏感度其实低于 1；规模载荷长期稳定在 1.2-1.3；"
             "价值载荷波动较大，2022-2024 年一度到 0.35-0.45，近一年回落。滚动年化 Alpha 大多数时候在 "
             "0-20% 之间摆动，但 t 值绝大多数小于 2，始终没有形成稳健证据。",
             after=8)
    add_picture(doc, "09_滚动36个月因子载荷.png", "图 2  滚动 36 个月因子载荷：市场 Beta 下行，小盘载荷稳定")
    add_picture(doc, "10_滚动Alpha与解释度.png", "图 3  滚动 Alpha 与解释度：2021 年后 R² 明显下降")


def add_conclusion(doc):
    doc.add_heading("结论与注意", level=1)
    add_callout(doc, "结论",
                "长期超额收益的大头是 Alpha（占比约 94%），不是市场也不是风格；但这个 Alpha 年化点估计约 "
                "10.8%、t 值仅 1.12，统计上不显著。风格画像稳定为“高 Beta + 强小盘 + 温和价值”，而同期小盘"
                "因子累计跑输，导致小盘暴露对收益是负贡献（约 -28%）。近五年 Alpha 仍是主角，2021 年后因子"
                "解释力进一步降到 R²≈0.33-0.38。更稳妥的表述是：创元科技是一只个性强、波动大的股票，超额"
                "收益更像自身故事或运气，而不是能稳定复制的选股能力。")
    add_para(doc,
             "超额收益大头是 Alpha。全样本 129.8% 的超额里，Alpha 贡献约 121.6%（94%），但 t 值仅 1.12，"
             "统计不显著，建议理解成“高波动 + 阶段性运气”，而不是已验证的稳定能力。",
             after=6)
    add_para(doc,
             "风格画像稳定：高 Beta + 强小盘 + 温和价值。Beta≈1.09、SMB≈1.30、HML≈0.36，红利不显著；"
             "因为同期小盘因子累计跑输，小盘暴露对收益是负贡献（约 -28%）。",
             after=6)
    add_para(doc,
             "近五年 Alpha 仍是主角（占比约 79%），SMB 转为正贡献（25%）；2021 年后因子解释力降到 "
             "R²≈0.33-0.38，个股自身因素占比上升。",
             after=6)
    add_para(doc,
             "模型解释力有限，别过度归因。全样本 R² 只有 0.46，超过一半的收益波动来自因子模型之外。",
             after=8)

    doc.add_heading("使用时的注意事项", level=2)
    add_para(doc, "共同窗口限制：多因子模型从 2014-10-20 开始，受中证 1000 指数起点限制；更早年份只有 CAPM 结果，不能直接和后面的模型比较。", after=6)
    add_para(doc, "因子相关性：红利与价值因子相关性很高（相关系数约 0.84，红利 VIF≈5.1），系数拆分会互相挤占，红利贡献的小数点级数字不必过度解读。", after=6)
    add_para(doc, "数据时效：数据截至 2026-08-05，2026 年是部分年度；最新滚动 Alpha 依然不显著。历史归因只能说明过去收益的来源，不能直接外推未来。", after=6)
    add_para(doc,
             "下一步建议：把 Alpha 继续拆到行业、事件和持仓集中度，补充交易行为指标，并对因子构造做稳健性"
             "检验（换指数、换窗口、加风格约束），再判断它是否值得写进实盘策略。",
             after=6)


def main():
    doc = Document()
    setup_section(doc)
    setup_styles(doc)

    doc.core_properties.title = "创元科技因子归因分析：收益到底从哪来"
    doc.core_properties.author = "梁厚德"
    doc.core_properties.subject = "创元科技（000551.SZ）因子归因分析"

    add_title_block(doc)
    add_callout(doc, "一句话结论",
                "长期持有创元科技的 129.8% 超额收益，绝大多数来自因子解释不了的个股自身部分（Alpha 占比约 "
                "94%），而不是市场或风格。这个 Alpha 年化点估计约 10.8%，但 t 值只有 1.12，统计上不显著；"
                "同时公司呈现“高 Beta + 强小盘”画像，而小盘风格在这段时间整体跑输，反而拖累了约 28% 的收益。"
                "更谨慎的说法是：创元科技是一只个性强、波动大的股票，超额收益更像自身故事或运气，而不是能"
                "稳定复制的选股能力。")

    add_work_summary(doc)
    add_core_results(doc)
    add_annual_view(doc)
    add_rolling_view(doc)
    add_conclusion(doc)

    add_rule(doc, color="C7D3E0", sz=6, space=4, before=6, after=2)
    add_para(doc, "数据来源：data/创元科技_000551_体检数据_20260806.xlsx；"
             "回归与归因由 code/analysis_factor_attribution.py 完成，本报告由 code/build_factor_report_docx.py 生成。",
             size=9, color=MUTED, before=2, after=0)

    os.makedirs(os.path.dirname(OUT_DOCX), exist_ok=True)
    doc.save(OUT_DOCX)
    print(OUT_DOCX)


if __name__ == "__main__":
    main()
