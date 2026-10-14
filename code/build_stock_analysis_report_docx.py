# -*- coding: utf-8 -*-
"""生成《创元科技股票综合分析》Word 报告。"""

from pathlib import Path as _PublicPath
_PUBLIC_ROOT = _PublicPath(__file__).resolve().parents[1]
import os

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

BASE_DIR = str(_PUBLIC_ROOT)
OUT_DOCX = os.path.join(BASE_DIR, "output", "创元科技_股票综合分析_20260806.docx")
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

    for level, (size, color_, space_b, space_a) in [
        ("Heading 1", (16, BLUE, 16, 8)),
        ("Heading 2", (13, BLUE, 12, 6)),
        ("Heading 3", (12, DARK_BLUE, 8, 4)),
    ]:
        h = doc.styles[level]
        set_style_font(h, size=size, color=color_, bold=True)
        h.paragraph_format.space_before = Pt(space_b)
        h.paragraph_format.space_after = Pt(space_a)
        h.paragraph_format.keep_with_next = True


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
    run = hp.add_run("创元科技（000551.SZ） 股票综合分析")
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
    add_para(doc, "创元科技（000551.SZ） 股票综合分析", size=11, bold=True,
             color=BLUE, before=0, after=4)
    add_para(doc, "股票内涵、收益归因与未来走势情景", size=22, bold=True,
             color=INK, before=0, after=2)
    add_para(doc, "收益数据以 1999 年以来为准（剔除 1996 年极端行情），结合因子归因与基本面数据",
             size=13, color=MUTED, before=0, after=8)

    meta = [
        ("分析对象", "创元科技（000551.SZ），A 股深圳主板，申万环保行业"),
        ("数据口径", "收益：1999-01-04 至 2026-08-05；因子：2014-10-20 至 2026-08-05；"
                     "财务：1994-2025 年报（结论以近年为主）；估值：2018-01 至 2026-08-06"),
        ("分析方法", "收益统计 / CAPM 与多因子归因 / 杜邦分解 / 估值分位 / 筹码与资金面 / 情景分析"),
        ("报告日期", "2026-08-06"),
    ]
    for label, value in meta:
        add_rich_para(doc, [
            (label + "：", {"size": 10.5, "bold": True, "color": INK}),
            (value, {"size": 10.5, "color": "333333"}),
        ], before=0, after=2, line=1.2)
    add_rule(doc, color=BLUE, sz=12, space=6, before=4, after=10)


def section_executive_summary(doc):
    doc.add_heading("一、核心结论", level=1)
    add_callout(doc, "一句话总结",
        "创元科技是一只盈利质量正在改善、估值不算贵、但波动极大的中型高端制造股。"
        "1999 年以来年化收益约 7.1%，超额收益大部分来自因子解释不了的个股自身部分，"
        "统计上不显著，不能当作稳定的选股能力。当前 PE 处于历史低位、PB 偏高，"
        "长期逻辑取决于盈利改善能否持续；短期则面临筹码分散与前期急涨急跌后的方向选择。")

    add_para(doc, "表 1  关键指标一览", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["维度", "指标", "数值", "评价"],
        [
            ["收益", "1999 年以来 CAGR", "7.1%", "累计涨幅约 5.6 倍，长期为正"],
            ["收益", "年化波动率 / 最大回撤", "45.9% / -81.4%", "波动极大，风险水平偏高"],
            ["收益", "月度胜率 / 年度胜率", "50.9% / 63.0%", "月度几乎一半时间亏损"],
            ["来源", "2014 年以来超额收益", "+129.8%", "Alpha 占约 94%，但 t 值仅 1.12"],
            ["盈利", "2025 年 ROE", "10.2%", "较 2021 年 5.5% 明显改善"],
            ["现金流", "经营现金流 / 净利润", "2.09x（近5年均值）", "利润含金量高"],
            ["估值", "PE(TTM) / 历史分位", "19.3 倍 / 6%", "自身历史低位"],
            ["估值", "PB / 历史分位", "2.15 倍 / 88%", "偏高，靠 ROE 改善支撑"],
            ["筹码", "股东户数一年变化", "+41%", "筹码近期分散"],
        ],
        widths=[0.7, 2.2, 1.7, 1.9],
        aligns=["left", "left", "left", "left"],
    )
    add_caption(doc, "数据来源：output/创元科技_收益表现分析_20260806.xlsx、"
                     "创元科技_因子归因分析_20260806.xlsx、创元科技_基本面分析数据_20260806.xlsx。", after=6)


def section_return_profile(doc):
    doc.add_heading("二、收益画像：高波动下的慢牛", level=1)
    add_para(doc,
        "剔除 1996 年 +716% 的极端行情后，1999-01-04 至 2026-08-05 共 6584 个交易日，"
        "创元科技后复权累计涨幅约 5.6 倍，年化收益率约 7.1%。收益分布极不均匀："
        "最好年度 2009 年 +144.5%，最差年度 2008 年 -56.1%；月度胜率仅 50.9%，"
        "最好月份 +53.9%，最差月份 -41.7%。相对沪深300 累计超额约 1.6 倍、相对申万环保行业约 0.8 倍，"
        "但相对中证500 累计跑输约 0.4 倍，说明它显著强于环保板块、强于大盘，弱于中小盘指数。",
        after=6)

    add_para(doc, "表 2  任意时点买入持有 N 年的年化收益分布（1999 年以来）", size=10,
             bold=True, color=DARK_BLUE, before=2, after=3, keep_next=True)
    add_table(doc,
        ["持有年限", "样本数", "胜率", "平均年化", "中位年化", "最差年化", "最好年化"],
        [
            ["1 年", "320", "55.0%", "14.6%", "5.4%", "-59.9%", "259.1%"],
            ["3 年", "296", "62.2%", "6.0%", "5.7%", "-26.4%", "56.9%"],
            ["5 年", "272", "66.9%", "6.0%", "5.7%", "-22.9%", "47.3%"],
            ["10 年", "212", "88.7%", "6.3%", "5.1%", "-4.4%", "24.1%"],
        ],
        widths=[0.7, 0.7, 0.7, 0.8, 0.8, 0.9, 0.9],
        aligns=["center", "center", "center", "center", "center", "center", "center"],
    )
    add_caption(doc, "注：按月度收益滚动计算，持有 10 年最差年化仍有 -4.4%，即长期持有并非无风险。", after=6)

    add_picture(doc, "01_累计净值与回撤.png",
                "图 1  1999 年以来累计净值与历史回撤（后复权，对数轴）")
    add_picture(doc, "05_滚动收益.png",
                "图 2  滚动 1/3/5/10 年年化收益：当前处于历史中高分位")


def section_factor_attribution(doc):
    doc.add_heading("三、收益来源：超额主要靠个股自身", level=1)
    add_para(doc,
        "2014-10 至 2026-08，四因子模型（市场、规模、价值、红利）只能解释约 46% 的日收益波动。"
        "风格画像为高 Beta（市场载荷约 1.09）+ 强小盘（规模载荷约 1.30）+ 温和价值（价值载荷约 0.36），"
        "红利因子不显著。累计超额收益 129.8% 中，Alpha 贡献约 121.6%（占 94%），市场贡献约 47.7%，"
        "规模风格贡献约 -36.6%，即长期持有小盘风格本身在拖后腿。",
        after=6)
    add_para(doc,
        "Alpha 年化点估计约 10.8%，但 t 值仅 1.12，统计上不显著；更稳妥的理解是，"
        "超额收益主要来自个股自身的高波动和阶段性行情，而不是已验证的稳定选股能力。"
        "近 5 年累计超额 71.9%，Alpha 占比约 79%，但模型 R² 降到 0.34-0.38，"
        "市场 Beta 降至 0.7-0.8，个股自身因素的重要性进一步上升。",
        after=6)

    add_para(doc, "表 3  全样本与子样本因子归因", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["样本", "累计超额", "Alpha 占比", "市场占比", "规模占比", "R²"],
        [
            ["全样本（2014-10 起）", "+129.8%", "93.7%", "+36.7%", "-28.2%", "0.457"],
            ["近 10 年", "+111.5%", "123.9%", "+23.9%", "-45.8%", "0.361"],
            ["近 5 年", "+71.9%", "78.9%", "-12.1%", "+25.3%", "0.339"],
            ["近 3 年", "+67.6%", "77.0%", "+14.3%", "+9.3%", "0.378"],
            ["近 1 年", "+31.8%", "75.4%", "+35.0%", "-6.8%", "0.326"],
        ],
        widths=[1.6, 1.0, 1.0, 1.0, 1.0, 0.7],
        aligns=["left", "center", "center", "center", "center", "center"],
    )
    add_caption(doc, "注：占比超过 100% 说明其他来源为负贡献；模型共同窗口受中证1000指数起点限制。", after=6)

    add_picture(doc, "08_因子累计收益与个股超额.png",
                "图 3  因子累计收益与个股超额：小盘因子长期下行，个股超额主要靠自身段")


def section_fundamental(doc):
    doc.add_heading("四、基本面与估值：盈利改善撑起估值", level=1)
    add_para(doc,
        "公司营收增长平缓（2025 年 42.21 亿元，同比约 +0.7%），但盈利质量明显提升："
        "2025 年归母净利润 2.95 亿元，EPS 0.6105 元；ROE 从 2021 年 5.5% 升至 2025 年 10.2%，"
        "归母净利率从 3% 左右升至 7.0%，毛利率升至 26.5%。杜邦分解显示，ROE 提升主要靠净利率修复，"
        "资产周转率（0.66 次）和权益乘数（2.21）相对稳定，杠杆处于合理偏低水平。",
        after=6)
    add_para(doc,
        "现金流质量是基本面中最突出的部分：近 5 年经营现金流/净利润均值约 2.09 倍，"
        "2025 年经营现金流 6.22 亿元，自由现金流 4.29 亿元，连续两年转正；"
        "资产负债率降至 37.3%，财务结构稳健。利润不是纸面富贵，而是有真实现金支撑。",
        after=6)

    add_para(doc, "表 4  盈利与估值核心数据", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["指标", "数值", "含义"],
        [
            ["2025 年 ROE（杜邦）", "10.2%", "较近 5 年均值 7.8% 明显上台阶"],
            ["2025 年毛利率 / 归母净利率", "26.5% / 7.0%", "毛利率稳定，净利率持续修复"],
            ["2025 年经营现金流 / 自由现金流", "6.22 亿 / 4.29 亿", "利润含金量高，资本开支回报期"],
            ["2025 年资产负债率", "37.3%", "制造业偏低，偿债稳健"],
            ["近 5 年净利润 CAGR", "23.5%", "高增速有低基数成分"],
            ["PE(TTM) / 历史分位", "19.3 倍 / 6%", "2018 年以来自身历史低位"],
            ["PB / 历史分位", "2.15 倍 / 88%", "偏高，由 ROE 改善支撑"],
            ["PEG（近5年净利增速）", "约 0.82", "相对增速不算贵"],
            ["2025 年股息率", "0.30%", "非高股息标的，回报靠股价"],
        ],
        widths=[2.1, 1.6, 2.8],
        aligns=["left", "center", "left"],
    )
    add_caption(doc, "注：PE 低、PB 高的组合本质是 ROE 上移（PB=PE×ROE），"
                     "需同时跟踪盈利趋势与净资产变化。", after=6)

    add_picture(doc, "06_基本面质量分析.png",
                "图 4  创元科技基本面质量全景（1999-2025）：营收&利润、利润率、ROE杜邦、现金流质量")


def section_chips(doc):
    doc.add_heading("五、筹码与资金：长周期集中、近一年分散", level=1)
    add_para(doc,
        "从长周期看，股东户数自 2019 年约 5.5 万户持续下降，筹码一度趋于集中；"
        "但最新披露的股东户数约 3.93 万户，较一年前 2.79 万户增加约 41%，"
        "说明近一年筹码明显分散，散户参与度上升，机构覆盖仍不足。"
        "北向持股约占流通市值 1.45%，外资关注度一般；融资余额从 2025 年中的 3.6 亿元"
        "升至 2026 年中的 5.37 亿元，杠杆资金参与度在提高。",
        after=6)

    add_para(doc, "表 5  筹码与资金指标", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["指标", "最新值", "变化"],
        [
            ["股东户数", "39,273 户", "一年 +41%，近期筹码分散"],
            ["北向持股占比", "1.45%（2026-06-30）", "季度披露，关注度一般"],
            ["融资余额", "5.37 亿元（2026-06-30）", "较 2025 年中 3.6 亿元明显上升"],
            ["十大流通股东", "以自然人、一般法人为主", "机构覆盖不足"],
        ],
        widths=[1.5, 2.6, 2.4],
        aligns=["left", "left", "left"],
    )
    add_caption(doc, "数据来源：data/创元科技_000551_体检数据_20260806.xlsx。", after=6)


def section_current(doc):
    doc.add_heading("六、当前状态：急涨急跌后的缩量修复", level=1)
    add_para(doc,
        "2026 年截至 8 月 5 日收益约 -7.7%，相对沪深300 跑输约 8.3%。"
        "6 月底至 7 月初股价快速冲高（后复权从 166 附近升至 213 附近，约两周 +33%），"
        "随后一个月回落约 45% 至 117 附近，8 月初连续三日反弹至 131.49。"
        "成交量从 6 月末高峰的 15 亿元萎缩至 8 月初的 2-4 亿元，反弹为明显缩量。",
        after=6)
    add_para(doc,
        "滚动收益显示，当前 1/3/5/10 年年化分别为 22.8%/15.3%/15.4%/6.7%，"
        "处于历史 67.8%/68.9%/79.7%/59.0% 分位，中短期位置并不低；"
        "当前实际股价约 13.5 元（2026-08-06），PE(TTM) 19.3 倍、PB 2.15 倍。",
        after=6)

    add_para(doc, "表 6  近期行情关键节点（后复权收盘）", size=10, bold=True,
             color=DARK_BLUE, before=2, after=3, keep_next=True)
    add_table(doc,
        ["日期", "后复权收盘", "说明"],
        [
            ["2026-06-25", "166.46", "启动前平台"],
            ["2026-07-01", "203.47", "阶段高点，成交额约 12.3 亿元"],
            ["2026-07-30", "117.11", "阶段低点，较 7 月高点回落约 42%"],
            ["2026-08-05", "131.49", "缩量反弹，成交额约 4.1 亿元"],
        ],
        widths=[1.1, 1.4, 4.0],
        aligns=["center", "center", "left"],
    )
    add_caption(doc, "注：后复权价用于展示真实收益，当前实际股价约 13.5 元（2026-08-06）。", after=6)


def section_scenarios(doc):
    doc.add_heading("七、未来可能走势：三种情景", level=1)
    add_para(doc,
        "未来走势不是单点预测，而是盈利兑现程度与市场风格共同作用的结果。"
        "以下三种情景基于历史数据与当前状态推演，仅供分析参考。",
        after=6)

    add_para(doc, "表 7  情景推演", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["情景", "触发条件", "可能的路径", "关键信号"],
        [
            ["基准",
             "中报/年报验证 ROE 守 10%、毛利率与现金流稳定",
             "股价随盈利温和修复，PE 在 19-24 倍区间波动，波动仍大",
             "ROE、毛利率、经营现金流/净利润"],
            ["乐观",
             "盈利超预期，或小盘风格回归、精密轴承/机器人等题材发酵",
             "高 Beta+强小盘画像提供弹性，PE 向 24-30 倍回归，涨幅可能快速放大",
             "净利润增速、小盘因子、成交额放大"],
            ["谨慎",
             "营收继续停滞、毛利率连续下滑、现金流恶化或业绩不及预期",
             "高 PB 先承压，PE 低分位不构成支撑，回撤可能达 30-50%",
             "股东户数继续回升、融资余额高位回落、放量下跌"],
        ],
        widths=[0.7, 2.1, 2.1, 1.6],
        aligns=["center", "left", "left", "left"],
        font_size=9,
    )
    add_caption(doc, "注：历史数据显示该股年化波动率约 46%，月度胜率约 51%，任何单一路径都不应被过度外推。", after=6)

    add_para(doc, "表 8  建议跟踪的核心指标", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["指标", "阈值 / 方向", "含义"],
        [
            ["ROE", "是否守住 10%", "盈利改善逻辑的根基"],
            ["毛利率", "连续两季下滑需警惕", "竞争恶化信号"],
            ["经营现金流 / 净利润", "低于 1 需深挖", "利润含金量下降"],
            ["PE(TTM)", "升破 24 倍", "估值安全边际减弱"],
            ["股东户数", "继续回升", "筹码进一步分散"],
            ["大小盘风格", "小盘因子趋势", "高 SMB 载荷的放大器"],
        ],
        widths=[1.6, 2.2, 2.7],
        aligns=["left", "left", "left"],
    )
    add_caption(doc, "建议频率：季度跟踪财报与股东户数，半年全面复检，重大公告即时检查。", after=6)


def section_risks(doc):
    doc.add_heading("八、风险提示", level=1)
    add_para(doc,
        "1. 高波动风险：年化波动率约 46%，历史最大回撤约 81%，持有体验较差，"
        "风险承受能力低的投资者不适合重仓。", after=3)
    add_para(doc,
        "2. 盈利可持续性风险：ROE 从 5.5% 快速升至 10.2%，净利高增速含低基数成分；"
        "若 2026 年盈利回落，高 PB 估值将明显承压。", after=3)
    add_para(doc,
        "3. 估值矛盾风险：PE 处历史低位而 PB 处历史高位，若 ROE 回落，PB 回落速度可能快于 PE。", after=3)
    add_para(doc,
        "4. 流动性风险：日均成交额较小、机构覆盖不足，大资金进出冲击成本高，"
        "极端行情下流动性可能快速收缩。", after=3)
    add_para(doc,
        "5. 筹码风险：近一年股东户数上升 41%、融资余额上升，短期筹码结构偏散，"
        "急涨急跌后容易出现放量方向选择。", after=3)
    add_para(doc,
        "6. 行业与政策风险：环保与高端制造竞争加剧，政策执行节奏可能影响订单；"
        "宏观 PMI 荣枯线附近、PPI 负值，制造业景气度一般。", after=3)

    add_callout(doc, "总体判断",
        "创元科技的“内涵”是盈利质量改善 + 现金流扎实 + 估值不贵，"
        "但“脾气”是高波动、低月度胜率、Alpha 统计不显著。"
        "适合能承受 45% 级别波动、认可盈利改善逻辑、持有周期 3 年以上的卫星配置；"
        "不适合追求稳定现金流、短期博弈或大资金频繁进出的投资者。")


def section_data_sources(doc):
    add_rule(doc, color="C7D3E0", sz=6, space=4, before=6, after=2)
    add_para(doc,
        "数据来源：data/创元科技_000551_体检数据_20260806.xlsx；\n"
        "收益表现分析：output/创元科技_收益表现分析_20260806.xlsx（1999 年起）；\n"
        "因子归因分析：output/创元科技_因子归因分析_20260806.xlsx（2014 年起）；\n"
        "基本面分析：output/创元科技_基本面分析数据_20260806.xlsx 与基本面数据分析报告；\n"
        "图表来源：output/图片/01、05、06、08；\n"
        "本报告由 code/build_stock_analysis_report_docx.py 生成。\n"
        "报告日期：2026-08-06。所含数据与分析均为历史表现与情景推演，不构成投资建议。",
        size=9, color=MUTED, before=2, after=0)


def main():
    doc = Document()
    setup_section(doc)
    setup_styles(doc)

    doc.core_properties.title = "创元科技股票综合分析"
    doc.core_properties.author = "梁厚德"
    doc.core_properties.subject = "创元科技（000551.SZ）股票内涵与未来走势分析"

    add_title_block(doc)
    section_executive_summary(doc)
    section_return_profile(doc)
    section_factor_attribution(doc)
    section_fundamental(doc)
    section_chips(doc)
    section_current(doc)
    section_scenarios(doc)
    section_risks(doc)
    section_data_sources(doc)

    os.makedirs(os.path.dirname(OUT_DOCX), exist_ok=True)
    doc.save(OUT_DOCX)
    print(f"Report saved: {OUT_DOCX}")


if __name__ == "__main__":
    main()
