# -*- coding: utf-8 -*-
"""生成《创元科技基本面数据分析》Word 报告。"""

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
OUT_DOCX = os.path.join(BASE_DIR, "output", "创元科技_基本面数据分析报告_20260806.docx")
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
    run = hp.add_run("创元科技（000551.SZ） 基本面数据分析")
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


# ═══════════════════════════════════════════════════════════════
# 数据 (从 extract_fundamental_data.py 输出提取)
# ═══════════════════════════════════════════════════════════════

# 收益数据
RTN_CAGR = 7.1
RTN_VOL = 45.85
RTN_MAXDD = -81.41
RTN_BEST_YR = "2009 (+144.5%)"
RTN_WORST_YR = "2008 (-56.1%)"
RTN_MONTHLY_WIN = 0.509
RTN_ANNUAL_WIN = 0.6296
RTN_VS_HS300 = 1.6343
RTN_VS_ZZ500 = -0.4203
RTN_VS_HB = 0.7617

# 估值
PE_CURRENT = 19.26
PE_PCT = 6.0
PE_5Y_P20 = 20.7
PE_5Y_P50 = 24.0
PE_5Y_P80 = 30.4
PB_CURRENT = 2.149
PB_PCT = 88.0
PB_5Y_P20 = 1.657
PB_5Y_P50 = 1.844
PB_5Y_P80 = 2.107

# 盈利
ROE_5Y = 7.79
ROE_2025 = 10.15
NET_MARGIN_2025 = 7.0  # 归母口径
ASSET_TURN_2025 = 0.66
EQUITY_MULT_2025 = 2.21
NM_5Y = 4.69
GM_5Y = 22.76
RG_5Y = 5.84
NIG_5Y = 23.50
ROA_5Y = 3.67
AT_5Y = 0.67
EM_5Y = 2.22

# 现金流
OCF_NI_5Y = 2.09
FCF_5Y_SUM = -0.28
DEBT_2025 = 37.34

# 其他
REV_2025 = 42.21
NI_2025 = 2.95
EPS_2025 = 0.6105
DIV_2025 = 0.06
DIV_YIELD_2025 = 0.3036
SH_CNT = 39273


def add_title_block(doc):
    add_para(doc, "创元科技（000551.SZ） 长期持有体检", size=11, bold=True,
             color=BLUE, before=0, after=4)
    add_para(doc, "基本面数据分析：盈利质量与估值定位", size=22, bold=True,
             color=INK, before=0, after=2)
    add_para(doc, "基于 1994-2025 年完整财务数据与 2018-2026 年日频估值数据",
             size=13, color=MUTED, before=0, after=8)

    meta = [
        ("分析对象", "创元科技（000551.SZ），A 股深圳主板，申万环保行业"),
        ("数据区间", "财务：1994-2025 年报；估值：2018-01-02 至 2026-08-06"),
        ("分析方法", "杜邦分解 / 现金流质量 / 估值分位 / 盈利趋势 / 分红与筹码分析"),
        ("报告日期", "2026-08-06"),
    ]
    for label, value in meta:
        add_rich_para(doc, [
            (label + "：", {"size": 10.5, "bold": True, "color": INK}),
            (value, {"size": 10.5, "color": "333333"}),
        ], before=0, after=2, line=1.2)
    add_rule(doc, color=BLUE, sz=12, space=6, before=4, after=10)


def section_executive_summary(doc):
    doc.add_heading("核心结论", level=1)
    add_callout(doc, "一句话总结",
        f"创元科技是一家中型环保/高端制造企业，近几年盈利质量正在明显改善（ROE 从 5.5% 爬升至 "
        f"10.2%，经营现金流充沛），估值处于自身历史低位（PE 仅 6% 分位），但 PB 偏高（88% 分位），"
        f"说明低 PE 主要是盈利改善驱动的而非股价低迷。过去收益波动极大（年化波动率近 46%），"
        f"长期持有的核心逻辑在于盈利改善能否持续、估值能否维持在合理区间。")

    add_para(doc, "表 1  基本面核心指标一览", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["维度", "指标", "数值", "评价"],
        [
            ["盈利质量", f"近5年ROE均值", f"{ROE_5Y:.1f}%", "中等偏下，但2025年回升至10.2%"],
            ["盈利质量", f"近5年净利率均值", f"{NM_5Y:.1f}%", "偏低，制造业属性"],
            ["盈利质量", f"近5年毛利率均值", f"{GM_5Y:.1f}%", "稳定，竞争格局尚可"],
            ["现金流", f"经营现金流/净利润均值", f"{OCF_NI_5Y:.2f}x", "优秀，利润含金量高"],
            ["现金流", f"近5年自由现金流总和", f"{FCF_5Y_SUM:.1f}亿", "基本打平，扩张期特征"],
            ["成长性", f"近5年营收增速均值", f"{RG_5Y:.1f}%", "温和增长"],
            ["成长性", f"近5年净利润增速均值", f"{NIG_5Y:.1f}%", "快速增长，盈利改善驱动"],
            ["偿债", f"资产负债率(2025)", f"{DEBT_2025:.1f}%", "偏低，财务稳健"],
            ["估值", f"PE(TTM)={PE_CURRENT:.1f}", f"历史分位 {PE_PCT:.0f}%", "估值低位（自2018年）"],
            ["估值", f"PB={PB_CURRENT:.3f}", f"历史分位 {PB_PCT:.0f}%", "估值偏高（市净率口径）"],
            ["分红", f"2025年股息率", f"{DIV_YIELD_2025:.2f}%", "偏低，非高股息标的"],
            ["筹码", f"股东户数", f"{SH_CNT:,}", "近年持续下降，筹码趋于集中"],
        ],
        widths=[0.8, 1.8, 1.5, 2.4],
        aligns=["left", "left", "center", "left"],
    )
    add_caption(doc, "数据来源：创元科技_基本面分析数据_20260806.xlsx。", after=6)

    add_para(doc,
        f"从收益端看，剔除 1996 年极端行情后（仅取 1999 年以来数据），创元科技年化收益率约 {RTN_CAGR}%，"
        f"累计涨幅约 5.6 倍，但收益波动仍大：年化波动率约 {RTN_VOL}%，最大回撤 {RTN_MAXDD}%。"
        f"相对沪深300有超额（累计超额 {RTN_VS_HS300:.1f}倍），但跑输中证500（-{abs(RTN_VS_ZZ500):.1f}倍），"
        f"跑赢申万环保行业指数（超额 {RTN_VS_HB:.1f}倍）。"
        f"从基本面看，公司近年盈利质量明显提升，ROE 从 2021 年的 5.5% 上升至 2025 年的 {ROE_2025}%，"
        f"利润增速远超营收增速，说明盈利能力在增强。经营现金流长期为净利润的 2 倍以上，利润含金量极高。",
        after=8)


def section_profit_quality(doc):
    doc.add_heading("一、盈利质量分析", level=1)

    doc.add_heading("1.1 收入与利润趋势", level=2)
    add_para(doc,
        f"2025 年公司实现营收 {REV_2025} 亿元，归母净利润 {NI_2025} 亿元，EPS {EPS_2025} 元。"
        f"从长期趋势看，公司营收从 2015 年的 21.5 亿增长至 2025 年的 {REV_2025} 亿，10 年间 CAGR 约 7.0%。"
        f"但利润端增速更快：同期归母净利润从 0.24 亿增长至 {NI_2025} 亿，CAGR 约 28.5%，"
        f"这主要得益于 2020 年以来盈利能力的持续改善。",
        after=6)

    add_para(doc, "表 2  营收与利润趋势（2015-2025）", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["年份", "营收(亿)", "营收增速", "归母净利(亿)", "净利增速", "EPS"],
        [
            ["2015", "21.50", "-5.4%", "0.24", "-20.9%", "0.06"],
            ["2016", "22.82", "+6.1%", "0.50", "+72.4%", "0.12"],
            ["2017", "25.91", "+13.5%", "0.66", "+26.0%", "0.17"],
            ["2018", "29.84", "+15.2%", "0.86", "+15.6%", "0.22"],
            ["2019", "32.37", "+8.5%", "1.13", "+25.9%", "0.28"],
            ["2020", "32.13", "-0.7%", "1.17", "+3.6%", "0.29"],
            ["2021", "36.96", "+15.1%", "1.09", "-3.6%", "0.27"],
            ["2022", "41.80", "+13.1%", "1.64", "+38.8%", "0.41"],
            ["2023", "43.23", "+3.4%", "1.55", "-0.8%", "0.38"],
            ["2024", "41.92", "-3.0%", "2.51", "+40.7%", "0.52"],
            [f"2025", f"{REV_2025:.1f}", "+0.7%", f"{NI_2025:.2f}", "+13.3%", f"{EPS_2025:.2f}"],
        ],
        widths=[0.6, 1.0, 1.0, 1.2, 1.0, 0.8],
        aligns=["center", "center", "center", "center", "center", "center"],
        font_size=9,
    )
    add_caption(doc, "注：增速为年报同比数据。2022年营收增长含并购因素。", after=6)

    add_para(doc,
        f"收入端看，公司在 2021-2023 年经历了较快增长，2024 年小幅回落，2025 年基本持平。"
        f"利润端则保持强劲增长态势，2024 年归母净利润大增 40.7%，2025 年继续增长 13.3%。"
        f"这说明公司近年来在成本控制、产品结构优化或非经常性收益方面有较大改善。",
        after=6)

    doc.add_heading("1.2 利润率分析", level=2)
    add_para(doc,
        f"毛利率长期维持在 20-27% 之间，近 5 年均值约 {GM_5Y:.1f}%，说明公司在产业链中有一定的定价权，"
        f"但并非高毛利赛道。2024 年毛利率显著提升至 26.5%，2025 年在 26% 左右。"
        f"归母口径的净利率从 2015 年的 1.1% 提升至 2025 年的 {NET_MARGIN_2025:.1f}%，"
        f"改善幅度超过 6 个百分点，反映了规模效应和经营效率的提升。",
        after=6)

    add_para(doc, "表 3  利润率与ROE趋势", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["年份", "毛利率(%)", "归母净利率(%)", "ROE(%)", "ROA(%)"],
        [
            ["2015", "18.7", "1.1", "1.8", "0.7"],
            ["2016", "18.6", "2.2", "3.1", "1.2"],
            ["2017", "17.9", "2.6", "4.1", "1.6"],
            ["2018", "18.1", "2.9", "5.3", "2.0"],
            ["2019", "19.9", "3.5", "6.3", "2.4"],
            ["2020", "20.4", "3.6", "6.1", "2.3"],
            ["2021", "19.7", "3.0", "5.5", "2.0"],
            ["2022", "22.7", "3.9", "7.1", "2.6"],
            ["2023", "22.9", "3.6", "6.6", "2.5"],
            ["2024", "26.5", "6.0", "9.6", "4.0"],
            ["2025", "26.5", "7.0", "10.2", "4.6"],
        ],
        widths=[0.8, 1.3, 1.3, 1.2, 1.2],
        aligns=["center", "center", "center", "center", "center"],
        font_size=9,
    )
    add_caption(doc, "注：归母净利率=归母净利润/营业总收入；ROE=归母净利润/归属母公司权益。毛利率为(营收-营业成本)/营收的近似计算。", after=6)

    doc.add_heading("1.3 杜邦分解：ROE从哪来", level=2)
    add_para(doc,
        f"2025 年 ROE 为 {ROE_2025}%，杜邦分解为：\n"
        f"    ROE = 净利率({NET_MARGIN_2025:.1f}%) 资产周转率({ASSET_TURN_2025:.2f}次) 权益乘数({EQUITY_MULT_2025:.2f}) = {ROE_2025:.1f}%\n\n"
        f"这反映了公司的盈利模式：",
        after=6)

    add_para(doc,
        f"净利率驱动力在增强：2015-2020 年归母净利率仅 1-4%，2024 年突破 6%，2025 年达到 7.0%。"
        f"此前 ROE 偏低的主要原因就是净利率拖后腿，而近年这一短板正在快速补上。",
        after=4)
    add_para(doc,
        f"资产周转率维持在中等水平（近 5 年均值 {AT_5Y:.2f} 次），制造业特征明显，"
        f"重资产属性决定了周转率难以大幅提升。",
        after=4)
    add_para(doc,
        f"权益乘数约 {EM_5Y:.1f} 倍（资产负债率 {DEBT_2025:.1f}%），杠杆处于合理偏低水平，"
        f"财务风险可控，但也意味着 ROE 的提升主要依靠盈利能力的增强而非加杠杆。",
        after=6)

    add_para(doc,
        f"近 5 年 ROE 均值仅 {ROE_5Y:.1f}%，主要是 2021-2023 年拉低（5.5-7.1%），"
        f"近两年已回升至 10% 附近。如果盈利改善趋势持续，ROE 有望保持在 10% 以上，"
        f"达到 A 股中等偏上水平。",
        after=6)


def section_cashflow_quality(doc):
    doc.add_heading("二、现金流质量与偿债能力", level=1)

    doc.add_heading("2.1 经营现金流：利润的含金量", level=2)
    add_para(doc,
        f"创元科技的现金流质量在一众 A 股中属于优秀水平。近 5 年经营现金流净额与净利润的比值"
        f"均值高达 {OCF_NI_5Y:.2f}x，几乎每 1 元账面利润对应 2 元以上的真金白银流入。"
        f"2025 年经营现金流净额 6.22 亿元，为净利润的 1.44 倍。"
        f"这说明公司的利润不是'纸面富贵'，而是有真实现金支撑的。",
        after=6)

    add_para(doc, "表 4  现金流质量（2015-2025）", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["年份", "经营现金流(亿)", "投资现金流(亿)", "自由现金流(亿)", "经营现金流/净利润"],
        [
            ["2015", "1.48", "-0.61", "0.87", "6.16"],
            ["2016", "2.01", "-0.18", "1.82", "4.03"],
            ["2017", "1.74", "-0.78", "0.96", "2.61"],
            ["2018", "2.50", "-1.34", "1.17", "2.91"],
            ["2019", "2.83", "-1.17", "1.65", "2.50"],
            ["2020", "2.06", "-1.78", "0.29", "1.77"],
            ["2021", "2.69", "-1.56", "1.13", "2.46"],
            ["2022", "4.29", "-1.18", "3.10", "2.61"],
            ["2023", "2.19", "-2.29", "-0.09", "1.41"],
            ["2024", "4.64", "-0.56", "4.09", "1.85"],
            ["2025", "6.22", "-1.93", "4.29", "1.44"],
        ],
        widths=[0.7, 1.3, 1.3, 1.3, 1.5],
        aligns=["center", "center", "center", "center", "center"],
        font_size=9,
    )
    add_caption(doc, "注：自由现金流=经营现金流+投资现金流（均为净额）。", after=6)

    add_para(doc,
        f"自由现金流方面，2024 年转正大幅改善，2025 年继续保持 4.29 亿的正自由现金流。"
        f"此前公司处于扩张期，投资活动持续消耗现金，经营现金流刚好覆盖或略有缺口。"
        f"近两年的改善说明公司可能进入了资本开支回报期。",
        after=6)

    doc.add_heading("2.2 资产负债结构", level=2)
    add_para(doc,
        f"2025 年末公司资产负债率 {DEBT_2025:.1f}%，较 2018 年的 50.8% 显著下降，"
        f"处于制造业中偏低水平。总资产 {64.25} 亿，归属母公司权益 {29.05} 亿，"
        f"有息负债规模有限，流动比率 1.89、速动比率 1.33，短期偿债能力充裕。"
        f"财务结构整体稳健，不存在明显的债务风险。",
        after=6)


def section_valuation(doc):
    doc.add_heading("三、估值分析", level=1)

    doc.add_heading("3.1 PE/PB 历史分位", level=2)
    add_para(doc,
        f"截至 2026 年 8 月 6 日，创元科技 PE(TTM) 为 {PE_CURRENT} 倍，"
        f"处于 2018 年以来历史分位的 {PE_PCT:.0f}%（即仅有 {PE_PCT:.0f}% 的交易日估值低于当前值）。"
        f"PB 为 {PB_CURRENT} 倍，处于历史分位的 {PB_PCT:.0f}%。",
        after=6)

    add_para(doc, "表 5  估值分位表（2018-01 至 2026-08-06）", size=10, bold=True,
             color=DARK_BLUE, before=2, after=3, keep_next=True)
    add_table(doc,
        ["估值指标", "当前值", "历史分位", "近5年20%分位", "近5年中位数", "近5年80%分位", "近5年最低", "近5年最高"],
        [
            ["PE(TTM)", f"{PE_CURRENT}", f"{PE_PCT:.0f}%", f"{PE_5Y_P20}", f"{PE_5Y_P50}", f"{PE_5Y_P80}", "13.8", "65.2"],
            ["PB", f"{PB_CURRENT}", f"{PB_PCT:.0f}%", f"{PB_5Y_P20}", f"{PB_5Y_P50}", f"{PB_5Y_P80}", "1.30", "2.68"],
        ],
        widths=[0.85, 0.85, 0.8, 1.1, 1.1, 1.1, 0.95, 0.95],
        aligns=["left", "center", "center", "center", "center", "center", "center", "center"],
        font_size=9,
    )
    add_caption(doc, "注：PE 近5年 20%/50%/80% 分位是基于近5年约1250个交易日数据计算的。历史分位基于2018年起全部日频数据。", after=6)

    add_para(doc,
        f"PE 估值处于历史低位（{PE_PCT:.0f}%分位），远低于近 5 年中位数 {PE_5Y_P50} 倍，"
        f"也低于 20% 分位（{PE_5Y_P20} 倍）。这意味着当前 PE 估值在历史上看是'便宜'的。"
        f"但 PB 处于 {PB_PCT:.0f}% 分位，远高于中位数 {PB_5Y_P50}，"
        f"甚至在近 5 年 80% 分位之上。",
        after=6)

    doc.add_heading("3.2 PE低、PB高：怎么看", level=2)
    add_para(doc,
        "PE 低而 PB 高的矛盾现象，背后是净资产收益率（ROE）的变化：\n\n"
        f"PB = PE  ROE\n"
        f"当 ROE 上升时，同样的 PB 对应更低的 PE。创元科技近年 ROE 从 5-6% 升至 10.2%，"
        f"每股净资产也在增长（从 2020 年的 4.79 元增至 2025 年的 5.99 元），"
        f"PB 的'贵'是因为净资产本身的价值在提升，而 PE 的'便宜'反映的是盈利改善后的估值消化。\n\n"
        f"综合来看，当前估值水平对长期持有者偏有利：PE 在历史低位提供了安全边际，"
        f"PB 偏高但由 ROE 改善支撑，并不构成泡沫。如果未来盈利继续增长，PE 有进一步扩张空间。",
        after=6)

    doc.add_heading("3.3 PEG 与股息率", level=2)
    add_para(doc,
        f"近 5 年净利润 CAGR 约 {NIG_5Y:.1f}%，当前 PE 约 {PE_CURRENT} 倍，"
        f"PEG 约 {PE_CURRENT/NIG_5Y:.2f}，显著低于 1，说明相对于盈利增速，估值并不贵。"
        f"但需注意近 5 年的高增速部分来自低基数效应（2021 年净利润仅 1.09 亿），"
        f"未来能否维持高增速是关键。\n\n"
        f"2025 年度每股分红 0.06 元，按当前股价计算的股息率约 {DIV_YIELD_2025:.2f}%，"
        f"远低于 10 年期国债收益率（约 1.6-1.8%）。公司并非高股息标的，"
        f"长期持有的回报主要依靠股价增值而非分红。",
        after=6)


def section_dividend(doc):
    doc.add_heading("四、分红与股东回报", level=1)
    add_para(doc,
        "创元科技近年分红记录较为连续，2017 年以来基本保持每年分红，"
        "但分红金额偏低。2025 年每 10 股派 0.6 元，每股股利 0.06 元，"
        "按 2025 年 EPS 0.61 元计算，分红率约 9.8%。",
        after=6)

    add_para(doc, "表 6  近年分红记录", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["报告期", "每股股利(元)", "股息率(%)", "EPS", "分红率"],
        [
            ["2017", "0.10", "1.53", "0.17", "58.8%"],
            ["2020", "0.10", "1.18", "0.29", "34.5%"],
            ["2022", "0.07", "0.66", "0.41", "17.1%"],
            ["2023", "0.07", "0.79", "0.38", "18.4%"],
            ["2024", "0.10", "0.78", "0.52", "19.2%"],
            ["2025", "0.06", "0.30", "0.61", "9.8%"],
        ],
        widths=[1.0, 1.2, 1.0, 1.0, 1.0],
        aligns=["center", "center", "center", "center", "center"],
        font_size=9,
    )
    add_caption(doc, "注：2024年实施10转2股送转，总股本从4.0亿股增至4.85亿股，对每股股利有摊薄效应。", after=6)

    add_para(doc,
        "分红率近年从 2017 年的 58.8% 下降至 2025 年的 9.8%，大幅降低。"
        "一方面是因为利润高增长下分红金额并未同步提升，另一方面公司可能将更多利润用于再投资。"
        "2024 年实施了每 10 股转增 2 股的送转方案，总股本从 4.0 亿增至 4.85 亿股，"
        "也相当于一种非现金的股东回报。",
        after=6)


def section_shareholders(doc):
    doc.add_heading("五、股东结构与筹码", level=1)
    add_para(doc,
        f"截至最新披露，创元科技股东户数约 {SH_CNT:,} 户。"
        f"从较长周期看，股东户数自 2019 年的约 5.5 万户呈下降趋势，"
        f"表明筹码在持续集中。股东户数下降而股价总体上涨，"
        f"通常意味着机构或大户在逐步吸筹，这是一个偏积极的信号。",
        after=6)

    add_para(doc,
        f"融资余额方面，截至 2025 年末融资余额呈温和上升态势，"
        f"杠杆资金参与度在可控范围内。北向资金持股约占流通市值的 1.45% 左右，"
        f"外资关注度一般。（注意：目前北向持股已改为季度披露，实时数据不可得。）",
        after=6)

    add_para(doc,
        "从机构持仓角度看，十大流通股东中机构占比不高，"
        "以自然人股东和一般法人为主。基金持股比例较低，"
        "说明机构投资者对这只股票的覆盖和研究还不够充分，"
        "可能存在'未被充分定价'的机会，但也意味着流动性主要来自散户。",
        after=6)


def section_industry_and_macro(doc):
    doc.add_heading("六、行业与宏观环境", level=1)
    add_para(doc,
        "创元科技归属申万环保行业，但实际业务横跨环保设备、精密轴承、"
        "磨料磨具等多个领域，更像一家多元化高端制造企业。"
        "环保行业近年来处于政策驱动与竞争加剧并存的阶段，"
        "\"双碳\"目标提供了长期需求支撑，但行业内分化严重。",
        after=6)

    add_para(doc,
        "从相对表现看，创元科技长期跑赢申万环保行业指数"
        f"（累计超额 {RTN_VS_HB:.1f} 倍），说明公司在行业内属于佼佼者。"
        "从中观环境看，当前 PMI 处于荣枯线附近震荡，PPI 仍在负值区间，"
        "制造业整体景气度一般。无风险利率方面，10 年期国债收益率"
        "在 1.6-1.8% 的低位徘徊，低利率环境对权益资产估值形成支撑。",
        after=6)

    add_para(doc, "表 7  宏观经济简况（2026年8月）", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["指标", "当前状态", "对创元科技的影响"],
        [
            ["PMI", "49-50 荣枯线附近", "制造业需求偏弱，营收增长承压"],
            ["PPI", "同比负值", "产品价格可能承压，但成本端同步受益"],
            ["10Y国债", "1.6-1.8%", "低利率支撑估值，股息率对比优势不明显"],
            ["M2增速", "7-8%", "流动性充裕，有利于中小盘估值"],
            ["环保政策", "\"双碳\"持续推进", "长期利好环保设备与减排技术需求"],
        ],
        widths=[1.2, 1.8, 3.5],
        aligns=["left", "left", "left"],
        font_size=9,
    )
    add_caption(doc, "注：宏观数据为截至2026年8月的近似值，不构成预测。", after=8)


def section_risks(doc):
    doc.add_heading("七、主要风险", level=1)

    add_para(doc, "1. 盈利改善的可持续性风险", size=11, bold=True, color=INK, before=4, after=2)
    add_para(doc,
        f"ROE 从 5.5% 升至 10.2% 速度较快，需要验证是否可持续。"
        f"2024 年净利润大增 40.7% 部分可能来自一次性收益或成本控制，"
        f"如果 2026 年出现盈利回落，高 PB 估值将面临压力。",
        after=6)

    add_para(doc, "2. 高波动性风险", size=11, bold=True, color=INK, before=4, after=2)
    add_para(doc,
        f"年化波动率近 46%，最大回撤近 81%（历史极值），"
        f"月度胜率约 {RTN_MONTHLY_WIN:.1%}，意味着持有一半的月份都在亏损。"
        f"不适合风险承受能力低的投资者。",
        after=6)

    add_para(doc, "3. 流动性风险", size=11, bold=True, color=INK, before=4, after=2)
    add_para(doc,
        f"日均成交额相对较小，机构覆盖不足，大资金进出可能产生较大冲击成本。"
        f"极端行情下流动性可能急剧收缩。",
        after=6)

    add_para(doc, "4. 行业竞争与政策风险", size=11, bold=True, color=INK, before=4, after=2)
    add_para(doc,
        "环保与高端制造行业竞争加剧，技术迭代可能影响公司市场份额。"
        "环保政策执行节奏的变化也可能影响订单获取。",
        after=6)

    add_para(doc, "5. 估值矛盾风险", size=11, bold=True, color=INK, before=4, after=2)
    add_para(doc,
        f"PE 处于历史低位（{PE_PCT:.0f}%分位）而 PB 处于高位（{PB_PCT:.0f}%分位），"
        f"这种偏差意味着：如果 ROE 下降，PB 可能快速回落，PE 反而上升。"
        f"投资者需同时关注盈利趋势和净资产变化。",
        after=6)


def section_comprehensive_rating(doc):
    doc.add_heading("八、综合体检评分", level=1)

    add_para(doc,
        "基于以上分析，按照分析大纲的综合评分框架，对创元科技各维度打分如下：",
        after=6)

    add_para(doc, "表 8  综合体检评分卡", size=10, bold=True, color=DARK_BLUE,
             before=2, after=3, keep_next=True)
    add_table(doc,
        ["维度", "评分(1-5)", "核心依据"],
        [
            ["收益能力", "3", f"CAGR {RTN_CAGR}%，相对沪深300有超额，但波动极大且不稳定"],
            ["风险水平", "2", f"年化波动率 {RTN_VOL}%，最大回撤 {RTN_MAXDD}%，VaR 偏高"],
            ["盈利质量", "4", f"ROE 持续改善至 10.2%，经营现金流/净利润>2x，利润含金量高"],
            ["估值位置", "4", f"PE 仅 {PE_PCT:.0f}% 分位，PEG<1，估值有安全边际；PB 偏高是隐忧"],
            ["现金流", "5", f"OCF/NI 均值 {OCF_NI_5Y:.2f}x，自由现金流转正，现金充沛"],
            ["筹码结构", "3", "股东户数下降、筹码集中，但机构覆盖不足"],
            ["流动性", "2", "日均成交额偏小，大资金进出不便"],
            ["分红回报", "2", f"股息率仅 {DIV_YIELD_2025:.2f}%，分红率偏低"],
            ["成长性", "4", f"净利 CAGR {NIG_5Y:.1f}%，盈利改善趋势明确"],
            ["偿债能力", "5", f"资产负债率 {DEBT_2025:.1f}%，财务结构稳健"],
        ],
        widths=[0.9, 0.8, 4.8],
        aligns=["left", "center", "left"],
        font_size=9,
    )
    add_caption(doc, "评分标准：1=差，2=较差，3=中性，4=较好，5=优秀。", after=6)

    add_para(doc, "加权综合评分：3.4 / 5", size=11, bold=True, color=DARK_BLUE, before=4, after=4)

    doc.add_heading("最终结论", level=2)
    add_callout(doc, "总体判断：有条件适合长期持有",
        "创元科技在盈利质量、现金流和成长性方面表现优秀，估值处于历史低位提供了安全边际。"
        "但高波动性、低流动性和偏弱的分红回报是需要正视的短板。\n\n"
        "适合长期持有的条件：\n"
        "1）投资者能承受 45% 级别的波动和阶段性大幅回撤\n"
        "2）认可公司盈利改善的长期逻辑，而非短期博弈\n"
        "3）作为组合中的卫星配置，仓位控制在可承受风险范围内\n\n"
        "不适合的情况：\n"
        "1）追求稳定现金流的保守型投资者\n"
        "2）需要频繁交易或大资金进出的机构\n"
        "3）对短期业绩波动敏感、持有周期短于 3 年的投资者")

    doc.add_heading("建议跟踪频率与关键指标", level=2)
    add_para(doc, "需持续盯住的 5 个核心指标：", size=11, bold=True, after=4)
    add_para(doc,
        "1. ROE 变化：每季度财报发布后，验证 ROE 是否维持在 10% 以上\n"
        "2. 毛利率：若毛利率出现连续 2 个季度以上下滑，需警惕竞争恶化\n"
        "3. 经营现金流/净利润：若该比值降至 1 以下，需深挖原因\n"
        "4. PE(TTM) 分位：若 PE 升破 24 倍（近5年中位数），估值安全边际消失\n"
        "5. 股东户数：若户数开始回升，可能意味着筹码分散、机构减持",
        after=6)

    add_para(doc, "建议跟踪频率：", size=11, bold=True, after=4)
    add_para(doc,
        "季度：阅读季报/年报，更新 ROE、现金流、毛利率、股东户数\n"
        "半年：全面复检杜邦分解、估值分位、行业比较\n"
        "重大事件：公告日即时检查（收购、定增、大股东增减持、业绩预告）",
        after=6)


def section_data_sources(doc):
    add_rule(doc, color="C7D3E0", sz=6, space=4, before=6, after=2)
    add_para(doc,
        "数据来源：data/创元科技_000551_体检数据_20260806.xlsx；\n"
        f"核心分析数据：output/创元科技_基本面分析数据_20260806.xlsx；\n"
        f"收益表现分析：output/创元科技_收益表现分析_20260806.xlsx；\n"
        f"因子归因分析：output/创元科技_因子归因分析_20260806.xlsx；\n"
        f"图表来源：output/图片/06_基本面质量分析.png；\n"
        "本报告由 code/build_fundamental_report_docx.py 生成。\n"
        "报告日期：2026-08-06。所含数据与分析均为历史表现，不构成投资建议。",
        size=9, color=MUTED, before=2, after=0)


def main():
    doc = Document()
    setup_section(doc)
    setup_styles(doc)

    doc.core_properties.title = "创元科技基本面数据分析"
    doc.core_properties.author = "梁厚德"
    doc.core_properties.subject = "创元科技（000551.SZ）基本面数据分析"

    add_title_block(doc)
    section_executive_summary(doc)
    section_profit_quality(doc)

    # Insert the fundamental quality chart
    add_picture(doc, "06_基本面质量分析.png",
                "图 1  创元科技基本面质量全景（1999-2025）：营收&利润、利润率、ROE杜邦、营收增速、现金流质量、EPS")

    section_cashflow_quality(doc)
    section_valuation(doc)
    section_dividend(doc)
    section_shareholders(doc)
    section_industry_and_macro(doc)
    section_risks(doc)
    section_comprehensive_rating(doc)
    section_data_sources(doc)

    os.makedirs(os.path.dirname(OUT_DOCX), exist_ok=True)
    doc.save(OUT_DOCX)
    print(f"Report saved: {OUT_DOCX}")


if __name__ == "__main__":
    main()
