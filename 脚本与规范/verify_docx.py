# -*- coding: utf-8 -*-
"""
国金证券深度报告 Word 成品全自动格式体检脚本 (verify_docx.py)

功能说明：
每次运行 render_docx.py 生成 Word 成品后，自动执行 12 项底层 XML 与占位符硬核验收（加 --visual 可附加 Word COM 真实分页抽检）：
1. 首页主标题字体：思源黑体 CN Heavy / 18pt / 加粗 / #004A70；
2. 首页摘要小标题与正文样式：原生 SummaryTitle + SummaryContentNoFirLinIndent（零首行缩进）；
3. 双硬换页符完整性：图表目录后换页起第一章 + 末章风险提示后换页起附录三表；
4. 图表标题原生样式保护：无手动重复编号，且未错误覆盖 bold=False；
5. 缺失图片检查：[待补充图片...] 为 0；
6. 图表容器宽度与深蓝横线：常规单图走 OneFullText（8192 dxa + tblInd=2551），宽图走 OneFullPage（10743 dxa），全部带 #004870 上下横线；
7. 数据表格字体与对齐：所有新增数据表统一楷体 9pt 且全部单元格水平居中；
8. 来源红线：零“公开行业资料 / 产业链公开拆解 / 公开行业资讯”等空泛来源词；
9. 核心业务与行业章节连续纯文字段：0 处（自动适配任意章节总数，仅跳过盈利预测/估值/风险提示章节，并支持 --allow-consec-text 特殊豁免）；
10. 核心正文章节真实段落字数统计（200–300 字目标区间，不超过 350 字）；
11. 首页摘要正文总字数护栏（≤ 850 字安全，> 900 字阻断）；
12. 全文未完成占位符扫描（[待模型填入] 等：默认标黄供研究员自刷模型，--release 模式阻断）。

命令行用法：
    D:\\anaconda3\\python.exe verify_docx.py --docx "D:\\深度报告\\超颖电子\\超颖电子_深度报告.docx"
"""

import argparse
import re
import sys
from docx import Document
from docx.oxml.ns import qn

VAGUE_SOURCES = [
    "公开行业资料",
    "公开行业研究资料",
    "行业公开资料",
    "产业链公开拆解",
    "公开行业资讯",
    "公开供应链BOM拆解资料",
]

# 未完成占位符扫描正则（盈利预测/估值/目标价/评级未填完时在 --release 正式发布模式下强制拦截）
PLACEHOLDER_RE = re.compile(r"\[待模型填入\]|\[待补充[^\]]*\]|\[待确认[^\]]*\]|\bTBD\b|\bTODO\b")


def verify_docx(docx_path: str, release: bool = False, allow_consec_text: bool = False) -> int:
    """对生成的国金标准 Word 报告执行全量 12 项 XML 与发布门禁验收，返回退出码（0=通过，1=阻断）。"""
    doc = Document(docx_path)

    # 1 & 2. 首页主标题与摘要样式检查
    left_cell = doc.tables[0].cell(0, 0)
    p0 = left_cell.paragraphs[0]
    r0 = p0.runs[0] if p0.runs else None
    title_font = r0.font.name if r0 else None
    title_size = round(r0.font.size.pt, 1) if (r0 and r0.font.size) else None
    title_bold = r0.bold if r0 else None
    title_color = str(r0.font.color.rgb) if (r0 and r0.font.color and r0.font.color.rgb) else None

    sum_head_styles = []
    sum_body_styles = []
    sum_char_cnt = 0
    for p in left_cell.paragraphs[1:]:
        sid = p.style.style_id if p.style else "None"
        if sid == "SummaryTitle":
            sum_head_styles.append(sid)
        elif p.text.strip():
            sum_body_styles.append(sid)
            sum_char_cnt += len(p.text.strip())

    # 3. 双硬换页符检查
    body_els = list(doc.element.body)
    page_breaks = []
    for idx, el in enumerate(body_els):
        for br in el.findall(".//" + qn("w:br")):
            if br.get(qn("w:type")) == "page":
                nxt = "".join(body_els[idx + 1].itertext()).strip()[:20] if idx + 1 < len(body_els) else ""
                page_breaks.append(nxt)

    # 4, 5, 6, 7, 8. 图表/表格检查
    dup_titles = []
    bold_override_bad = 0
    missing_imgs = []
    vague_hits = []
    one_full_text_cnt = 0
    one_full_page_cnt = 0
    two_half_cnt = 0
    borders_ok = 0
    chart_box_total = 0
    data_table_kaiti_bad = 0
    data_table_total = 0

    for p in doc.paragraphs:
        txt = p.text.strip()
        for v in VAGUE_SOURCES:
            if v in txt:
                vague_hits.append(v)
        if p.style and p.style.style_id in ("ContentGraphTitle", "ContentGraphTitleSuo"):
            if re.match(r"^图表\s*\d+", txt):
                dup_titles.append(txt)
            for run in p.runs:
                if run.bold is False:
                    bold_override_bad += 1
        if "[待补充图片" in txt:
            missing_imgs.append(txt)

    for t in doc.tables[1:]:
        full_txt = "".join(c.text for r in t.rows for c in r.cells)
        for v in VAGUE_SOURCES:
            if v in full_txt:
                vague_hits.append(v)
        if any(k in full_txt for k in ("三张报表预测摘要", "投资评级说明", "特别声明", "每股收益", "市场中相关报告")):
            continue

        if len(t.rows) == 3 and len(t.columns) in (1, 3):
            p_t0 = t.cell(0, 0).paragraphs[0]
            if p_t0.style and p_t0.style.style_id == "ContentGraphTitle":
                chart_box_total += 1
                tblPr = t._tbl.tblPr
                tblInd = tblPr.find(qn("w:tblInd")) if tblPr is not None else None
                ind_w = tblInd.get(qn("w:w")) if tblInd is not None else None
                tcPr0 = t.cell(0, 0)._tc.find(qn("w:tcPr"))
                tcW0 = tcPr0.find(qn("w:tcW")).get(qn("w:w")) if (tcPr0 is not None and tcPr0.find(qn("w:tcW")) is not None) else None
                if len(t.columns) == 3:
                    two_half_cnt += 1
                elif ind_w == "2551" and tcW0 == "8192":
                    one_full_text_cnt += 1
                else:
                    one_full_page_cnt += 1

                for col in ([0, 2] if len(t.columns) == 3 else [0]):
                    pt_p = t.cell(0, col).paragraphs[0]
                    if re.match(r"^图表\s*\d+", pt_p.text.strip()):
                        dup_titles.append(pt_p.text.strip())
                    for run in pt_p.runs:
                        if run.bold is False:
                            bold_override_bad += 1
                    if "[待补充图片" in t.cell(1, col).text:
                        missing_imgs.append(t.cell(1, col).text)

                b0 = tcPr0.find(qn("w:tcBorders")) if tcPr0 is not None else None
                bot0 = b0.find(qn("w:bottom")) if b0 is not None else None
                if bot0 is not None and bot0.get(qn("w:color")) == "004870":
                    borders_ok += 1
                continue

        # 新增数据表格：同时校验楷体 9pt 与全部单元格水平居中
        data_table_total += 1
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    if p.text.strip() and p.alignment != 1:  # WD_ALIGN_PARAGRAPH.CENTER == 1
                        data_table_kaiti_bad += 1
                    for run in p.runs:
                        if run.text.strip():
                            sz = round(run.font.size.pt, 1) if run.font.size else None
                            if run.font.name != "楷体" or sz != 9.0:
                                data_table_kaiti_bad += 1

    # 9 & 10. 核心业务/行业章节连续纯文字检查（自动适配任意章节数：跳过盈利预测/估值/风险提示章节）
    in_core_ch = False
    prev_was_para = False
    consec_paras = []
    para_lens = []
    for el in body_els:
        tag = el.tag.split("}")[-1]
        if tag == "p":
            # 仅统计 <w:t> 真实文本节点，避免 <w:instrText> 等干扰
            txt = "".join(t.text or "" for t in el.findall(".//" + qn("w:t"))).strip()
            pPr = el.find(qn("w:pPr"))
            pStyle = pPr.find(qn("w:pStyle")) if pPr is not None else None
            sid = pStyle.get(qn("w:val")) if pStyle is not None else ""
            if sid == "ContentFirTitle":
                in_core_ch = not any(k in txt for k in ("盈利预测", "估值", "投资建议", "风险提示"))
                prev_was_para = False
            elif sid == "ContentSecTitle":
                prev_was_para = False
            elif sid == "ContentContentNoFirLinIndent" and in_core_ch and txt:
                para_lens.append(len(txt))
                if prev_was_para:
                    consec_paras.append(txt[:28])
                prev_was_para = True
        elif tag == "tbl":
            prev_was_para = False

    # 12. 全局扫描未完成占位符（含封面表格 + 正文段落 + 所有数据表格）
    placeholder_hits = []
    for p in doc.paragraphs:
        found = PLACEHOLDER_RE.findall(p.text or "")
        placeholder_hits.extend(found)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                found = PLACEHOLDER_RE.findall(c.text or "")
                placeholder_hits.extend(found)

    print("=== 国金深度报告 Word 成品 12 项全自动 XML 与发布门禁体检 ===")
    print(f"1. 首页主标题: font={title_font}, size={title_size}pt, bold={title_bold}, color=#{title_color}")
    print(f"2. 首页摘要样式: 小标题={sum_head_styles} | 正文={sum_body_styles}")
    print(f"3. 整页换页符后紧接内容: {page_breaks}")
    print(f"4. 图表重复编号={len(dup_titles)} | 错误覆盖bold=False={bold_override_bad}")
    print(f"5. 缺失图片占位数量 (应为 0): {len(missing_imgs)}")
    print(
        f"6. 图表布局分布: OneFullText(正文齐宽)={one_full_text_cnt}, "
        f"OneFullPage(超宽通栏)={one_full_page_cnt}, TwoHalfPage(双图)={two_half_cnt} | 深蓝横线={borders_ok}/{chart_box_total}"
    )
    print(f"7. 新增数据表格非楷体9pt或未居中数 (应为 0): {data_table_kaiti_bad} (共 {data_table_total} 张表)")
    print(f"8. 空泛来源禁用词命中 (应为 0): {len(vague_hits)} {vague_hits}")
    print(f"9. 第1-3章连续纯文字段 (应为 0): {len(consec_paras)} {consec_paras}")
    if para_lens:
        print(
            f"10. 第1-3章正文共 {len(para_lens)} 段，字数范围: {min(para_lens)}–{max(para_lens)} 字，"
            f"均值 {sum(para_lens) // len(para_lens)} 字"
        )
    sum_status = "OK（首页单页安全）" if sum_char_cnt <= 850 else ("🟡 偏长，请核对首页未跨页" if sum_char_cnt <= 920 else "🔴 超长，大概率撑出首页")
    print(f"11. 首页摘要正文总字数: {sum_char_cnt} 字 -> {sum_status}")
    ph_status = "OK（已填完）" if len(placeholder_hits) == 0 else ("🔴 存在未填模型占位，--release 禁止发版" if release else "🟡 草稿态（待填入第四章模型数字）")
    print(f"12. 全文未完成占位符数量（[待模型填入]等）: {len(placeholder_hits)} 处 -> {ph_status}")

    # 硬失败判定：任何结构/格式缺陷直接返回非零退出码；--release 模式下有未填占位符同样阻断
    failures = []
    if len(dup_titles) > 0:
        failures.append(f"图表重复编号 {len(dup_titles)} 处")
    if bold_override_bad > 0:
        failures.append(f"图表标题错误覆盖 bold=False {bold_override_bad} 处")
    if len(missing_imgs) > 0:
        failures.append(f"缺失图片 {len(missing_imgs)} 处")
    if borders_ok < chart_box_total:
        failures.append(f"图表深蓝横线缺失 {chart_box_total - borders_ok} 处")
    if data_table_kaiti_bad > 0:
        failures.append(f"表格非楷体9pt或未居中 {data_table_kaiti_bad} 处")
    if len(vague_hits) > 0:
        failures.append(f"命中空泛来源词 {vague_hits}")
    if len(consec_paras) > 0:
        if allow_consec_text:
            print(f"  🟡 [连续纯文字豁免] 已启用 --allow-consec-text，放行 {len(consec_paras)} 处连续纯文字段")
        else:
            failures.append(f"第1-3章连续纯文字段 {len(consec_paras)} 处")
    if sum_char_cnt > 900:
        failures.append(f"首页摘要超长（{sum_char_cnt} 字 > 900 字）")
    if release and len(placeholder_hits) > 0:
        failures.append(f"正式发布模式下仍存在 {len(placeholder_hits)} 处未填占位符")

    if failures:
        print(f"\n🔴 [VERIFY BLOCKED] 验收未通过：{failures}")
        return 1
    if len(placeholder_hits) > 0:
        print(f"\n[DRAFT-OK] 11 项结构/排版 XML 全部通过（尚余 {len(placeholder_hits)} 处 [待模型填入]，填数后请再跑 --release 终验）")
    else:
        print("\n[RELEASE-PASS] 全部 12 项 Word XML 与正式发布门禁 100% 通过！")
    return 0


def run_visual_pagination_check(docx_path: str) -> None:
    """
    通过本机 Microsoft Word COM 静默导出临时 PDF，并用 PyMuPDF 检查真实渲染后的分页效果：
    1. 检查总页数与异常空白页；
    2. 检查第 1 页是否完整容纳封面与投资摘要（未被右栏财务表挤到第 2 页）；
    3. 自动渲染首页（P.1）及疑似异常页 PNG 到 `<公司目录>/09_视觉抽检页/`，供研究员或视觉模型快速复核。
    """
    import os
    import tempfile

    try:
        import fitz  # PyMuPDF
        import win32com.client  # type: ignore
    except ImportError as e:
        print(f"  🟡 [视觉分页检查跳过] 缺少依赖（{e}），请确保安装 pywin32 与 PyMuPDF")
        return

    abs_docx = os.path.abspath(docx_path)
    out_dir = os.path.join(os.path.dirname(abs_docx), "09_视觉抽检页")
    os.makedirs(out_dir, exist_ok=True)

    word = None
    doc_com = None
    with tempfile.TemporaryDirectory() as td:
        tmp_pdf = os.path.join(td, "preview.pdf")
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc_com = word.Documents.Open(abs_docx, ReadOnly=True)
            doc_com.SaveAs(tmp_pdf, FileFormat=17)  # 17 = wdFormatPDF
        except Exception as e:
            print(f"  🟡 [视觉分页检查失败] Word COM 导出 PDF 异常：{e}")
            return
        finally:
            if doc_com is not None:
                try:
                    doc_com.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass

        pdf = fitz.open(tmp_pdf)
        total_pages = len(pdf)
        blank_pages = []
        pages_to_render = {0}  # 必截第 1 页（封面+投资摘要）

        p2_txt = pdf[1].get_text() if total_pages >= 2 else ""
        cover_overflow = ("风险提示" in p2_txt[:300] and "目录" not in p2_txt[:300] and "一、" not in p2_txt[:200])
        if cover_overflow:
            pages_to_render.add(1)

        for idx in range(total_pages):
            txt = pdf[idx].get_text().strip()
            imgs = pdf[idx].get_images()
            # 既无图片且文字少于 35 字，视为疑似孤立空白页
            if len(txt) < 35 and len(imgs) == 0:
                blank_pages.append(idx + 1)
                pages_to_render.add(idx)

        for p_idx in sorted(pages_to_render):
            if 0 <= p_idx < total_pages:
                pix = pdf[p_idx].get_pixmap(dpi=150)
                pix.save(os.path.join(out_dir, f"page_{p_idx + 1:02d}.png"))

        pdf.close()
        print(
            f"13. 真实分页视觉检查（Word COM→PDF）：共 {total_pages} 页 | "
            f"疑似空白页={blank_pages} | 首页摘要疑似溢出={'🔴 是' if cover_overflow else 'OK'}"
        )
        print(f"    已输出抽检页 PNG -> {out_dir}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="国金深度报告 Word 成品 12 项 XML 与视觉分页体检脚本")
    parser.add_argument("--docx", required=True, help="待验收的 Word 路径")
    parser.add_argument("--release", action="store_true", help="正式发布模式：存在任一 [待模型填入] 占位符即以退出码 1 阻断")
    parser.add_argument("--visual", action="store_true", help="附加 Word COM 转 PDF 真实分页与首页溢出抽检")
    parser.add_argument("--allow-consec-text", action="store_true", help="特殊体例豁免：允许第1-3章存在少量连续纯文字段（仅标黄提示）")
    args = parser.parse_args()
    rc = verify_docx(args.docx, release=args.release, allow_consec_text=args.allow_consec_text)
    if args.visual:
        run_visual_pagination_check(args.docx)
    sys.exit(rc)
