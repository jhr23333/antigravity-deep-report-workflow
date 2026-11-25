# -*- coding: utf-8 -*-
"""
国金证券公司深度研究 Word 模板自动渲染脚本 (render_docx.py)

核心功能：
1. 克隆国金官方模板（国金证券_公司深度研究模板.docx），完整保护：
   - 首页右侧公司基本情况、股价走势图、主要财务指标表（留待国金插件导入）
   - 自动目录域（在 Word 中按 Ctrl+A -> F9 即可刷新页码）
   - 文末“附录：三张报表预测摘要”、“市场中相关报告评级比率分析”、“投资评级说明”及“特别声明”
2. 将 Markdown 初稿（05_完整初稿.md）精准渲染为国金原生样式：
   - 一级标题 -> ContentFirTitle（加粗、不另起页、不空行）
   - 二级标题 -> ContentSecTitle（加粗、不空行）
   - 正文段落 -> ContentContentNoFirLinIndent（支持有限加粗，禁用斜体/下划线/标红）
   - 单图占位 -> 默认 `OneFullText`（宽 8192 dxa ≈ 14.45 cm，左缩进 2551 dxa 与正文对齐）；超宽图加 `width: full` 走 `OneFullPage`（18.95 cm），均自动编号“图表 N：”
   - 双图并排 -> 3x3 容器表（TwoHalfPage，左 9.25 + 间距 0.45 + 右 9.25 cm）
   - 数据表格 -> 标题（ContentGraphTitle/Suo 原生加粗斜体）+ 国金标准表格（楷体 9pt 全单元格居中）+ 来源（ContentResource）
3. 内置质控部红线自动体检：检查主标题是否 >20 字、全文是否含禁用极端词或禁用排版格式。

命令行用法：
    python 脚本与规范/render_docx.py --md 05_完整初稿.md --meta report_meta.json --out 公司名_深度报告.docx
"""

import argparse
import os
import re
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor


# 质控部与分析师正文细则明确禁止的极端词汇清单
FORBIDDEN_EXTREME_WORDS = [
    "绝对领先",
    "全面领先",
    "遥遥领先",
    "确定性极强",
    "爆发式增长",
    "颠覆性",
    "唯一",
    "不可替代",
]

# 严禁出现在图表/表格来源中的空泛占位词（必须写具体一手机构或文档名）
FORBIDDEN_VAGUE_SOURCES = [
    "公开行业资料",
    "公开行业研究资料",
    "行业公开资料",
    "产业链公开拆解",
    "公开行业资讯",
    "公开供应链BOM拆解资料",
]

# 偏宣传化/过度确定性措辞告警清单（命中后提醒收敛为审慎卖方表述）
PROMOTIONAL_WORDS_WARN = [
    "爆发",
    "飙升",
    "重金卡位",
    "重金布局",
    "收获前夜",
    "胜负手",
    "确立拐点",
    "拐点特征明确",
    "精准锁定",
    "精准承接",
    "核心增长引擎",
    "核心增量引擎",
    "顺价传导顺畅",
]

# 国金官方模板中单图/双图容器表格的精确列宽与左缩进（单位：dxa，实测自北方华创/天通股份样本）
ONE_FULL_TEXT_WIDTH_DXA = 8192      # 常规单图（OneFullText，与正文左边界对齐，约 14.45 cm）
ONE_FULL_TEXT_INDENT_DXA = 2551     # 正文左侧留白缩进（匹配 ContentContentNoFirLinIndent 的 left=2552）
ONE_FULL_PAGE_WIDTH_DXA = 10743     # 宽幅通栏单图（OneFullPage，仅用于超宽时间轴/股权图/客户矩阵，约 18.95 cm）
DOUBLE_CHART_COL_DXA = 5244         # 双图并排单列宽（TwoHalfPage，约 9.25 cm）
DOUBLE_CHART_GAP_DXA = 255          # 双图并排中间间距（约 0.45 cm）


def set_cell_width_dxa(cell, width_dxa: int) -> None:
    """精确设置表格单元格宽度（单位：dxa）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag == qn("w:tcW"):
            tcPr.remove(child)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(width_dxa))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


def set_table_indent_dxa(table, indent_dxa: int) -> None:
    """为表格设置精确左缩进（如 2551 dxa，使 OneFullText 单图与正文左侧严丝合缝对齐）。"""
    tblPr = table._tbl.tblPr
    for child in list(tblPr):
        if child.tag in (qn("w:tblInd"), qn("w:jc")):
            tblPr.remove(child)
    tblInd = OxmlElement("w:tblInd")
    tblInd.set(qn("w:w"), str(indent_dxa))
    tblInd.set(qn("w:type"), "dxa")
    tblPr.append(tblInd)


def set_run_font_explicit(run, font_name: str, size_pt: float, bold: bool = False, color_hex: Optional[str] = None) -> None:
    """显式设置中英文字体、字号、加粗与颜色（用于精确还原首页 18pt 思源黑体 CN Heavy 主标题）。"""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color_hex:
        run.font.color.rgb = RGBColor.from_string(color_hex)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)


def add_page_break_before(doc: Document, anchor_el) -> None:
    """在指定 XML 节点之前插入独立整页换页符（用于目录后起第一章、第五章后起附录三表）。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)
    if anchor_el is not None:
        anchor_el.addprevious(p._element)


def set_cell_shading(cell, hex_color: str) -> None:
    """设置单元格背景底色（如国金深蓝表头 #004870 或浅灰隔行底色 #F5F7FA）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def set_table_borders(table, border_color: str = "BFBFBF") -> None:
    """为数据表格设置简洁的国金三线/细灰线边框。"""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="8" w:space="0" w:color="004870"/>'
        f'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="004870"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideV w:val="none"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def set_cell_border_bottom(cell, color_hex: str = "004870", sz: str = "8") -> None:
    """为单个单元格设置底部国金深蓝横线（模拟国金 Word 图表工具样式）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag == qn("w:tcBorders"):
            tcPr.remove(child)
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '  <w:top w:val="none"/>'
        '  <w:left w:val="none"/>'
        f'  <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color_hex}"/>'
        '  <w:right w:val="none"/>'
        "</w:tcBorders>"
    )
    tcPr.append(tcBorders)


def apply_guojin_chart_box_borders(table, is_pair: bool = False) -> None:
    """
    套用国金 Word“图表工具”容器表边框规范：
    外框与竖线全部隐藏，仅在 Row 0（图表标题下方）与 Row 1（图片下方/来源上方）绘制国金深蓝横线。
    双图并排时中间间距列（col=1）不画横线，左右两图各自独立成框。
    """
    remove_table_borders(table)
    cols = [0, 2] if is_pair else [0]
    for col in cols:
        set_cell_border_bottom(table.cell(0, col), color_hex="004870", sz="8")
        set_cell_border_bottom(table.cell(1, col), color_hex="004870", sz="6")


def set_run_kaiti(run, size_pt: float = 9.0) -> None:
    """将表格文字统一设置为楷体（比正文小一号），兼容中文东亚字体渲染。"""
    run.font.name = "楷体"
    run.font.size = Pt(size_pt)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "楷体")
    rFonts.set(qn("w:hAnsi"), "楷体")
    rFonts.set(qn("w:eastAsia"), "楷体")


def remove_table_borders(table) -> None:
    """移除图表容器表的全部边框，确保页面干净无杂线。"""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '  <w:top w:val="none"/>'
        '  <w:left w:val="none"/>'
        '  <w:bottom w:val="none"/>'
        '  <w:right w:val="none"/>'
        '  <w:insideH w:val="none"/>'
        '  <w:insideV w:val="none"/>'
        "</w:tblBorders>"
    )
    tblPr.append(borders)


def parse_kv_block(block_text: str) -> Dict[str, str]:
    """解析 <!-- CHART ... --> 或 <!-- TABLE_META ... --> 注释块内的键值对。"""
    result: Dict[str, str] = {}
    for line in block_text.splitlines():
        line = line.strip()
        if not line or ":" not in line:
            continue
        k, v = line.split(":", 1)
        result[k.strip().lower()] = v.strip()
    return result


def parse_markdown(md_text: str) -> List[Dict[str, Any]]:
    """
    将 05_完整初稿.md 解析为结构化区块列表。
    支持元素：
    - h1         : # 一、...
    - h2         : ## 1.1 ...
    - chart      : <!-- CHART ... -->
    - chart_pair : <!-- CHART_PAIR ... -->
    - table      : 可选 <!-- TABLE_META ... --> + Markdown 管道表格
    - paragraph  : 普通正文段落
    """
    blocks: List[Dict[str, Any]] = []
    lines = md_text.splitlines()
    i = 0
    n = len(lines)
    pending_table_meta: Optional[Dict[str, str]] = None

    while i < n:
        line = lines[i].strip()

        # 跳过空行（遵循国金质控红线：正文段落间不输出空行）
        if not line:
            i += 1
            continue

        # 1. 解析 HTML 注释块（CHART / CHART_PAIR / TABLE_META）
        if line.startswith("<!--"):
            comment_lines = [line]
            while "-->" not in comment_lines[-1] and i + 1 < n:
                i += 1
                comment_lines.append(lines[i])
            full_comment = "\n".join(comment_lines)
            inner = re.sub(r"^<!--\s*|\s*-->$", "", full_comment, flags=re.S).strip()
            first_line = inner.splitlines()[0].strip().upper() if inner else ""
            body_text = "\n".join(inner.splitlines()[1:])

            if first_line == "CHART":
                blocks.append({"type": "chart", "meta": parse_kv_block(body_text)})
            elif first_line == "CHART_PAIR":
                blocks.append({"type": "chart_pair", "meta": parse_kv_block(body_text)})
            elif first_line == "TABLE_META":
                pending_table_meta = parse_kv_block(body_text)
            i += 1
            continue

        # 2. 一级标题 (# )
        if line.startswith("# ") and not line.startswith("## "):
            blocks.append({"type": "h1", "text": line[2:].strip()})
            i += 1
            continue

        # 3. 二级标题 (## )
        if line.startswith("## ") and not line.startswith("### "):
            blocks.append({"type": "h2", "text": line[3:].strip()})
            i += 1
            continue

        # 4. Markdown 表格（以 | 开头）
        if line.startswith("|"):
            table_lines = []
            while i < n and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            rows = []
            for tl in table_lines:
                cells = [c.strip() for c in tl.strip("|").split("|")]
                # 跳过表头分隔线（如 |---|---|）
                if all(re.match(r"^:?-{2,}:?$", c) for c in cells if c):
                    continue
                rows.append(cells)
            if rows:
                blocks.append(
                    {
                        "type": "table",
                        "rows": rows,
                        "meta": pending_table_meta or {},
                    }
                )
            pending_table_meta = None
            continue

        # 5. 普通正文段落（合并连续非空普通行）
        para_lines = [line]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (
                not nxt
                or nxt.startswith("# ")
                or nxt.startswith("## ")
                or nxt.startswith("<!--")
                or nxt.startswith("|")
            ):
                break
            para_lines.append(nxt)
            i += 1
        blocks.append({"type": "paragraph", "text": "".join(para_lines)})

    return blocks


def add_formatted_runs(paragraph, text: str) -> int:
    """
    向段落写入文本，支持 **加粗** 语法（其余一律使用正文默认黑色非斜体非下划线）。
    返回该段落中包含的加粗片段数量，用于质控提示。
    """
    bold_count = 0
    # 清除可能误入的 Markdown 斜体/下划线标记，严格遵守质控红线
    clean_text = re.sub(r"</?u>", "", text)
    parts = re.split(r"(\*\*.*?\*\*)", clean_text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
            bold_count += 1
        else:
            paragraph.add_run(part)
    return bold_count


def run_qc_checks(md_text: str, meta: Optional[Dict[str, Any]]) -> List[str]:
    """执行国金质控部核心红线自动检查，返回告警列表。"""
    warnings: List[str] = []

    if meta:
        title = meta.get("report_title", "")
        if len(title) > 20:
            warnings.append(
                f"🔴 [标题超长] 报告主标题共 {len(title)} 字（含标点），超过质控红线 20 字上限：{title}"
            )
        if meta.get("rating_type") == "首次覆盖" and not meta.get("target_price"):
            warnings.append("🔴 [缺少目标价] 首次覆盖公司研究报告必须提供目标价。")

    for word in FORBIDDEN_EXTREME_WORDS:
        if word in md_text:
            warnings.append(f"🔴 [禁用极端词] 正文包含质控/文风禁词「{word}」，请替换为审慎表述。")

    for vague in FORBIDDEN_VAGUE_SOURCES:
        if vague in md_text:
            warnings.append(f"🔴 [空泛来源标注] 检测到「{vague}」，请替换为具体一手机构或公开文档名称。")

    meta_blob = ""
    if meta:
        for sec in meta.get("summary_sections", []):
            meta_blob += " ".join(sec.get("paragraphs", []))
    full_blob = md_text + "\n" + meta_blob
    for pw in PROMOTIONAL_WORDS_WARN:
        if pw in full_blob:
            warnings.append(f"🟡 [偏宣传化措辞] 检测到「{pw}」，建议收敛为客观审慎的卖方研究表述。")

    if re.search(r"<u>|</u>", md_text, flags=re.I):
        warnings.append("🟡 [禁用下划线] 检测到下划线标签，脚本已自动清除。")

    return warnings


def update_cover_summary(doc: Document, meta: Dict[str, Any]) -> None:
    """
    更新模板首页左侧（table[0] cell(0, 0)）的报告主标题与核心摘要段落，
    严格还原国金官方模板原生字体与样式，不触碰首页右侧 cell(0, 1) 的股价走势图与主要财务指标表：
    - 报告主标题：思源黑体 CN Heavy、18pt、加粗、深蓝 #004A70（左对齐/原样）
    - 摘要小标题（投资逻辑/盈利预测、估值和评级/风险提示）：原生 SummaryTitle（思源黑体 CN Regular 12pt 加粗深蓝 #004A70）
    - 摘要正文：原生 SummaryContentNoFirLinIndent（楷体五号、黑字、无首行缩进）
    """
    if not doc.tables:
        return
    cover_tbl = doc.tables[0]
    if len(cover_tbl.rows) < 1 or len(cover_tbl.columns) < 2:
        return

    style_by_id = {s.style_id: s for s in doc.styles}
    left_cell = cover_tbl.cell(0, 0)
    tc = left_cell._tc

    # 仅保留单元格属性 <w:tcPr>，移除原有的 <w:p> 和 <w:sdt> 占位控件
    for child in list(tc):
        if child.tag != qn("w:tcPr"):
            tc.remove(child)

    # 1. 写入主标题（精确还原模板：思源黑体 CN Heavy 18pt 加粗深蓝 #004A70）
    p0 = left_cell.add_paragraph()
    p0.paragraph_format.space_before = Pt(2)
    p0.paragraph_format.space_after = Pt(6)
    title_run = p0.add_run(meta.get("report_title", ""))
    set_run_font_explicit(title_run, "思源黑体 CN Heavy", 18.0, bold=True, color_hex="004A70")

    # 2. 依次写入投资逻辑、盈利预测与估值、风险提示（使用官方原生 SummaryTitle + SummaryContentNoFirLinIndent）
    for sec in meta.get("summary_sections", []):
        p_head = left_cell.add_paragraph()
        if "SummaryTitle" in style_by_id:
            p_head.style = style_by_id["SummaryTitle"]
        p_head.add_run(sec.get("heading", ""))

        for para_txt in sec.get("paragraphs", []):
            p_body = left_cell.add_paragraph()
            if "SummaryContentNoFirLinIndent" in style_by_id:
                p_body.style = style_by_id["SummaryContentNoFirLinIndent"]
            elif "SummaryContent" in style_by_id:
                p_body.style = style_by_id["SummaryContent"]
                p_body.paragraph_format.first_line_indent = Pt(0)
            add_formatted_runs(p_body, para_txt)


def find_appendix_anchor(doc: Document):
    """
    定位文末“附录：三张报表预测摘要”表格的 XML 元素，作为正文插入锚点。
    同时清理模板原有的 3 个空白/注意事项占位段落。
    """
    body = doc.element.body
    anchor_el = None
    for child in list(body):
        tag = child.tag.split("}")[-1]
        if tag == "tbl":
            txt = "".join(child.itertext())
            if "三张报表预测摘要" in txt or "损益表" in txt:
                anchor_el = child
                break

    # 删除锚点前方的模板占位说明（“注意事项：该部分内容要比摘要丰富，适当展开”）
    if anchor_el is not None:
        children = list(body)
        anchor_idx = children.index(anchor_el)
        # 向前检查最多 5 个段落，仅删除正文占位段落，保留目录 <sdt>
        for idx in range(anchor_idx - 1, max(-1, anchor_idx - 6), -1):
            el = children[idx]
            tag = el.tag.split("}")[-1]
            if tag == "sdt":
                break
            if tag == "p":
                txt = "".join(el.itertext()).strip()
                if (
                    not txt
                    or "注意事项" in txt
                ):
                    body.remove(el)
    return anchor_el


def insert_element_before(anchor_el, new_el) -> None:
    """将新建的段落或表格 XML 节点移动到锚点元素之前（保持正文在附录之前）。"""
    if anchor_el is not None:
        anchor_el.addprevious(new_el)


def resolve_image_path(img_path: str, base_dir: str) -> str:
    """支持绝对路径或相对于 Markdown 文件所在目录的相对路径。"""
    if os.path.isabs(img_path):
        return img_path
    return os.path.normpath(os.path.join(base_dir, img_path))


def render_report(
    template_path: str,
    md_path: str,
    out_path: str,
    meta_path: Optional[str] = None,
) -> None:
    """主渲染函数：将 Markdown + 元数据渲染进国金官方 Word 模板。"""
    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    meta: Optional[Dict[str, Any]] = None
    if meta_path and os.path.exists(meta_path):
        import json

        with open(meta_path, "r", encoding="utf-8") as f:
            meta = json.load(f)

    # 1. 运行质控红线自动检查
    qc_warnings = run_qc_checks(md_text, meta)
    for w in qc_warnings:
        print(w)

    # 2. 打开国金官方模板并定位文末附录锚点
    doc = Document(template_path)
    if meta:
        update_cover_summary(doc, meta)

    anchor_el = find_appendix_anchor(doc)
    base_dir = os.path.dirname(os.path.abspath(md_path))
    blocks = parse_markdown(md_text)

    # 辅助函数：按 style_id 查找样式对象（python-docx 默认按 name 查，这里兼容 style_id）
    style_by_id = {s.style_id: s for s in doc.styles}

    def get_style(style_id: str, fallback: str = "Normal"):
        return style_by_id.get(style_id, doc.styles[fallback])

    chart_counter = 0

    # 硬换页符 1：目录/图表目录结束后，强制另起新页开始正文第一章
    add_page_break_before(doc, anchor_el)

    for blk in blocks:
        btype = blk["type"]

        # ---------- 一级标题 ----------
        if btype == "h1":
            p = doc.add_paragraph(style=get_style("ContentFirTitle"))
            p.add_run(blk["text"])
            insert_element_before(anchor_el, p._element)

        # ---------- 二级标题 ----------
        elif btype == "h2":
            p = doc.add_paragraph(style=get_style("ContentSecTitle"))
            p.add_run(blk["text"])
            insert_element_before(anchor_el, p._element)

        # ---------- 正文段落 ----------
        elif btype == "paragraph":
            p = doc.add_paragraph(style=get_style("ContentContentNoFirLinIndent"))
            add_formatted_runs(p, blk["text"])
            insert_element_before(anchor_el, p._element)

        # ---------- 单图容器（默认 OneFullText 缩进对齐正文；width: full 时走 OneFullPage 通栏） ----------
        elif btype == "chart":
            chart_counter += 1
            m = blk["meta"]
            # 国金 ContentGraphTitle 样式自带 Word 自动编号“图表N：”与楷体五号加粗斜体，严禁覆盖 bold
            title_txt = m.get("title", "未命名图表")
            src_txt = f"来源：{m.get('source', '公司公告、国金证券研究所')}"
            img_abs = resolve_image_path(m.get("image", ""), base_dir)
            is_full_page = m.get("width", "text").strip().lower() == "full"

            tbl = doc.add_table(rows=3, cols=1)
            apply_guojin_chart_box_borders(tbl, is_pair=False)
            if is_full_page:
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                col_w_dxa = ONE_FULL_PAGE_WIDTH_DXA
                img_width_in = Inches(6.5)
            else:
                set_table_indent_dxa(tbl, ONE_FULL_TEXT_INDENT_DXA)
                col_w_dxa = ONE_FULL_TEXT_WIDTH_DXA
                img_width_in = Inches(5.35)

            for r in range(3):
                set_cell_width_dxa(tbl.cell(r, 0), col_w_dxa)
                # 禁止单图容器任意行被跨页腰斩
                trPr = tbl.rows[r]._tr.get_or_add_trPr()
                trPr.append(parse_xml(f"<w:cantSplit {nsdecls('w')}/>"))

            # 第 1 行：图表标题（保留 ContentGraphTitle 原生样式，并加 keep_with_next 防图题与图分离）
            p_title = tbl.cell(0, 0).paragraphs[0]
            p_title.style = get_style("ContentGraphTitle")
            p_title.paragraph_format.keep_with_next = True
            p_title.add_run(title_txt)

            # 第 2 行：图表图片（加 keep_with_next 防图片与来源行分离）
            p_img = tbl.cell(1, 0).paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.keep_with_next = True
            if img_abs and os.path.exists(img_abs):
                p_img.add_run().add_picture(img_abs, width=img_width_in)
            else:
                p_img.add_run(f"[待补充图片: {m.get('image', '')}]")

            # 第 3 行：来源标注
            p_src = tbl.cell(2, 0).paragraphs[0]
            p_src.style = get_style("ContentResource")
            p_src.add_run(src_txt)

            insert_element_before(anchor_el, tbl._element)

        # ---------- 双图并排（3x3 容器表，TwoHalfPage，左右各自带国金深蓝横线） ----------
        elif btype == "chart_pair":
            m = blk["meta"]
            chart_counter += 2

            tbl = doc.add_table(rows=3, cols=3)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            apply_guojin_chart_box_borders(tbl, is_pair=True)
            col_widths = [DOUBLE_CHART_COL_DXA, DOUBLE_CHART_GAP_DXA, DOUBLE_CHART_COL_DXA]
            for r in range(3):
                trPr = tbl.rows[r]._tr.get_or_add_trPr()
                trPr.append(parse_xml(f"<w:cantSplit {nsdecls('w')}/>"))
                for c in range(3):
                    set_cell_width_dxa(tbl.cell(r, c), col_widths[c])

            # 分别填充左列（col=0）与右列（col=2），完整保留 ContentGraphTitle 原生样式并防跨页分离
            for col_idx, prefix in [(0, "left"), (2, "right")]:
                p_t = tbl.cell(0, col_idx).paragraphs[0]
                p_t.style = get_style("ContentGraphTitle")
                p_t.paragraph_format.keep_with_next = True
                p_t.add_run(m.get(f"{prefix}_title", ""))

                p_i = tbl.cell(1, col_idx).paragraphs[0]
                p_i.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_i.paragraph_format.keep_with_next = True
                img_abs = resolve_image_path(m.get(f"{prefix}_image", ""), base_dir)
                if img_abs and os.path.exists(img_abs):
                    p_i.add_run().add_picture(img_abs, width=Inches(3.3))
                else:
                    p_i.add_run(f"[待补充图片: {m.get(f'{prefix}_image', '')}]")

                p_s = tbl.cell(2, col_idx).paragraphs[0]
                p_s.style = get_style("ContentResource")
                p_s.add_run(f"来源：{m.get(f'{prefix}_source', '公司公告、国金证券研究所')}")

            insert_element_before(anchor_el, tbl._element)

        # ---------- 数据表格（统一楷体 9pt；默认 OneFullText 缩进宽，width: full 时通栏） ----------
        elif btype == "table":
            chart_counter += 1
            rows_data: List[List[str]] = blk["rows"]
            m = blk.get("meta", {})
            n_rows = len(rows_data)
            n_cols = max(len(r) for r in rows_data)
            is_full_table = m.get("width", "auto").strip().lower() == "full" or (
                m.get("width", "auto").strip().lower() == "auto" and n_cols >= 5
            )

            # 表格标题（加 keep_with_next 防表题留在上一页、表格掉到下一页）
            title_style = "ContentGraphTitle" if is_full_table else "ContentGraphTitleSuo"
            res_style = "ContentResource" if is_full_table else "ContentResourceSuo"
            p_t = doc.add_paragraph(style=get_style(title_style, "ContentGraphTitle"))
            p_t.paragraph_format.keep_with_next = True
            p_t.add_run(m.get("title", "核心数据对比"))
            insert_element_before(anchor_el, p_t._element)

            tbl = doc.add_table(rows=n_rows, cols=n_cols)
            if "GJStylesTableDefalutStyle" in style_by_id:
                tbl.style = style_by_id["GJStylesTableDefalutStyle"]
            set_table_borders(tbl)
            if is_full_table:
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                total_w = ONE_FULL_PAGE_WIDTH_DXA
            else:
                set_table_indent_dxa(tbl, ONE_FULL_TEXT_INDENT_DXA)
                total_w = ONE_FULL_TEXT_WIDTH_DXA

            col_w = total_w // max(n_cols, 1)
            for r_idx, row_vals in enumerate(rows_data):
                # 禁止单行跨页截断；首行表头开启 tblHeader 跨页自动重复
                trPr = tbl.rows[r_idx]._tr.get_or_add_trPr()
                trPr.append(parse_xml(f"<w:cantSplit {nsdecls('w')}/>"))
                if r_idx == 0:
                    trPr.append(parse_xml(f"<w:tblHeader {nsdecls('w')}/>"))
                for c_idx in range(n_cols):
                    cell = tbl.cell(r_idx, c_idx)
                    set_cell_width_dxa(cell, col_w)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    val = row_vals[c_idx] if c_idx < len(row_vals) else ""
                    p_cell = cell.paragraphs[0]
                    # 表格全部内容（含首列）统一水平与垂直居中
                    p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_cell.paragraph_format.first_line_indent = Pt(0)
                    p_cell.paragraph_format.space_before = Pt(2)
                    p_cell.paragraph_format.space_after = Pt(2)
                    run = p_cell.add_run(val)
                    # 统一使用楷体、比正文小一号（9pt）
                    set_run_kaiti(run, size_pt=9.0)
                    if r_idx == 0:
                        set_cell_shading(cell, "004870")
                        run.bold = True
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    elif r_idx % 2 == 0:
                        set_cell_shading(cell, "F5F7FA")

            insert_element_before(anchor_el, tbl._element)

            # 表格来源标注
            p_s = doc.add_paragraph(style=get_style(res_style, "ContentResource"))
            p_s.add_run(f"来源：{m.get('source', '公司公告、国金证券研究所')}")
            insert_element_before(anchor_el, p_s._element)

    # 硬换页符 2：第五章风险提示结束后，强制另起新页开始“附录：三张报表预测摘要”
    add_page_break_before(doc, anchor_el)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    print(f"[OK] 已生成国金标准深度报告 Word：{out_path}（共渲染 {chart_counter} 张图表/表格）")
    print("[提示] 请在 Word 中打开后按 Ctrl+A -> F9 刷新目录页码，并用国金插件导入首页财务表与文末三表。")


def main() -> None:
    """命令行入口。"""
    default_tpl = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "模板资产",
        "国金证券_公司深度研究模板.docx",
    )
    parser = argparse.ArgumentParser(description="国金证券深度报告 Word 模板渲染脚本")
    parser.add_argument("--md", required=True, help="完整初稿 Markdown 路径")
    parser.add_argument("--out", required=True, help="输出 Word (.docx) 路径")
    parser.add_argument("--meta", default=None, help="封面元数据 JSON 路径（可选）")
    parser.add_argument("--template", default=default_tpl, help="国金官方 Word 模板路径")
    args = parser.parse_args()

    render_report(args.template, args.md, args.out, args.meta)


if __name__ == "__main__":
    main()
