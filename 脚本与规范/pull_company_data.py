# -*- coding: utf-8 -*-
"""
国金证券电子组深度报告——Step 1 财报全科目快照与行情提取脚本 (pull_company_data.py)

定位说明（2026-11-23 优化）：
- 常规财务指标优先走同花顺 iFinD（更快、底稿免截图）；
- 本脚本用于对 AI 后台按需拉取的最新一期 A 股季报/年报 PDF 执行【按报表分节、按行重构、按列校验】的精确抽取，
  彻底杜绝“把同比增速误当上期金额”“跨行折行漏读归母净利润/经营现金流”“缺关键科目仍静默返回 0”三大隐患。

命令行用法：
    D:\\anaconda3\\python.exe D:\\深度报告\\脚本与规范\\pull_company_data.py `
      --name 超颖电子 --code 603175.SH `
      --pdf D:\\深度报告\\超颖电子\\00_原始资料\\pdfs\\超颖电子_20261031_2026三季报.pdf `
      --out D:\\深度报告\\超颖电子\\00_原始资料\\01_最新财报全科目快照.md
"""

import argparse
import datetime
import os
import re
import sys
from typing import Dict, List, Optional, Tuple

import fitz


# 按三大合并报表严格分区解析，绝不混抓首页“主要会计数据”（该表第2列是同比%，极易误当上期金额）
STATEMENT_ITEMS: Dict[str, List[Tuple[str, List[str], bool]]] = {
    # (科目显示名, 正则匹配列表, 是否为必抽核心科目)
    "合并利润表": [
        ("营业总收入", [r"^(?:一、)?营业总收入$", r"^其中：营业收入$"], True),
        ("营业成本", [r"^其中：营业成本$", r"^营业成本$"], True),
        ("销售费用", [r"^销售费用$"], True),
        ("管理费用", [r"^管理费用$"], True),
        ("研发费用", [r"^研发费用$"], True),
        ("财务费用", [r"^财务费用$"], True),
        ("其中：利息费用", [r"^其中：利息费用$"], False),
        ("归母净利润", [r"归属于母公司股东的净利润"], True),
    ],
    "合并资产负债表": [
        ("货币资金", [r"^货币资金$"], True),
        ("应收账款", [r"^应收账款$"], True),
        ("存货", [r"^存货$"], True),
        ("固定资产", [r"^固定资产$"], True),
        ("在建工程", [r"^在建工程$"], True),
        ("短期借款", [r"^短期借款$"], False),
        ("一年内到期的非流动负债", [r"^一年内到期的非流动负债$"], False),
        ("长期借款", [r"^长期借款$"], False),
        ("资产总计", [r"^资产总计$"], True),
        ("负债合计", [r"^负债合计$"], True),
        ("归母净资产", [r"归属于母公司所有者权益.*合计"], True),
    ],
    "合并现金流量表": [
        ("经营活动现金流净额", [r"^经营活动产生的现金流量净额$"], True),
        ("购建长期资产支付现金", [r"购建固定资产、无形资产和其他长期资产支付的现金"], False),
    ],
}


def _parse_cn_number(tok: str) -> Optional[float]:
    """严格解析带千分位逗号的金额（单位：元）；含 % 或中文字符一律拒绝。"""
    s = (tok or "").strip()
    if "%" in s or not s:
        return None
    s = s.replace(",", "").replace("，", "")
    if not re.match(r"^-?\d+(?:\.\d+)?$", s):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _detect_statement_sections(doc: fitz.Document) -> Dict[str, List[int]]:
    """
    先定位三大合并报表的起止页区间（跳过首页“主要会计数据”和母公司单体报表），
    确保每个科目只在对应合并报表区间内按列提取。
    """
    starts: Dict[str, int] = {}
    for p_idx in range(len(doc)):
        txt = doc[p_idx].get_text()
        norm = "".join(txt.split())
        if "合并资产负债表" in norm and "母公司资产负债表" not in norm[:60] and "合并资产负债表" not in starts:
            starts["合并资产负债表"] = p_idx
        if "合并利润表" in norm and "母公司利润表" not in norm[:60] and "合并利润表" not in starts:
            starts["合并利润表"] = p_idx
        if "合并现金流量表" in norm and "母公司现金流量表" not in norm[:60] and "合并现金流量表" not in starts:
            starts["合并现金流量表"] = p_idx

    ordered = sorted(starts.items(), key=lambda kv: kv[1])
    ranges: Dict[str, List[int]] = {}
    for idx, (name, st_p) in enumerate(ordered):
        # 下一张合并报表或母公司报表开始前均属于本报表（最多跨 4 页）
        next_p = ordered[idx + 1][1] if idx + 1 < len(ordered) else min(len(doc), st_p + 4)
        pages = []
        for p in range(st_p, min(next_p + 1, st_p + 5, len(doc))):
            if p > st_p and any(
                stop in "".join(doc[p].get_text().split())[:120]
                for stop in ("母公司资产负债表", "母公司利润表", "母公司现金流量表")
            ):
                break
            pages.append(p)
        ranges[name] = pages
    return ranges


def _build_logical_rows(page_lines_with_pno: List[Tuple[str, int]]) -> List[Tuple[str, List[float], int]]:
    """
    将 PDF 垂直折行拆碎的科目名与紧随其后的数值列重组为逻辑行：
    (合并后的完整科目名, [本期/期末元, 上期/期初元], 页码)。
    彻底解决“归属于母公司股东的净\\n利润（净亏损以“-”号填列）”跨行漏读问题。
    """
    rows: List[Tuple[str, List[float], int]] = []
    label_buf: List[str] = []
    label_pno: int = 0
    num_buf: List[float] = []

    skip_prefixes = (
        "编制单位：", "单位：", "审计类型：", "项目", "法定代表人", "公司负责人",
        "主管会计", "会计机构", "202", "1 /", "2 /", "3 /", "4 /", "5 /",
        "6 /", "7 /", "8 /", "9 /", "10 /", "11 /", "12 /", "13 /",
    )

    def flush():
        nonlocal label_buf, num_buf
        if label_buf and num_buf:
            full_label = "".join(label_buf)
            # 去掉括号内会计提示，便于正则稳定匹配
            clean_label = re.sub(r"[（(][^）)]*[）)]", "", full_label).strip()
            rows.append((clean_label, list(num_buf[:2]), label_pno))
        label_buf = []
        num_buf = []

    for ln, pno in page_lines_with_pno:
        val = _parse_cn_number(ln)
        if val is not None:
            if label_buf:
                num_buf.append(val)
                if len(num_buf) >= 2:
                    flush()
            continue

        # 遇到非数字行：如果此前已攒到至少 1 个数字，先落盘上一条
        if num_buf:
            flush()

        # 跳过页眉页脚噪声
        if any(ln.startswith(sp) for sp in skip_prefixes) or re.match(r"^\d+\s*/\s*\d+$", ln):
            label_buf = []
            continue
        if "第三季度报告" in ln or "半年度报告" in ln or "年度报告" in ln:
            label_buf = []
            continue

        # 续行判断：若上一行明显被截断（含中文括号未闭合，或以“的净/的净利/流量/权益/亏损以”等结尾，
        # 或下一行以“润/利润/净额/合计/填列）”等续接词开头）则拼接为同一逻辑行
        is_continuation = False
        if label_buf:
            joined_prev = "".join(label_buf)
            if (
                joined_prev.endswith((
                    "的净", "的净利", "净利", "的现金", "流量", "所有者权益", "权益",
                    "亏损以", "净亏损以", "损失以", "（", "(", "的", "支付", "收到",
                ))
                or ln.startswith((
                    "润", "利润", "净额", "额", "合计", "填列）", "列）", "的现金",
                    "资产支付", "收益", "“-", '"-', "号填列", "）", ")",
                ))
                or (joined_prev.count("（") > joined_prev.count("）"))
                or (joined_prev.count("(") > joined_prev.count(")"))
            ):
                is_continuation = True

        if is_continuation:
            label_buf.append(ln)
        else:
            label_buf = [ln]
            label_pno = pno

    flush()
    return rows


def extract_financials_from_pdf(pdf_path: str) -> Tuple[Dict[str, Tuple[Optional[float], Optional[float], int]], List[str]]:
    """
    按三大合并报表分节提取核心科目，执行量级合理性校验（防止把百分比当金额）与必抽科目完整性校验。
    返回 (extracted_dict, missing_required_list)。
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"财报 PDF 不存在：{pdf_path}")

    doc = fitz.open(pdf_path)
    sec_pages = _detect_statement_sections(doc)
    results: Dict[str, Tuple[Optional[float], Optional[float], int]] = {}

    for stmt_name, items in STATEMENT_ITEMS.items():
        pages = sec_pages.get(stmt_name, [])
        if not pages:
            continue
        raw_lines: List[Tuple[str, int]] = []
        for p_idx in pages:
            for ln in doc[p_idx].get_text().splitlines():
                s = ln.strip()
                if s:
                    raw_lines.append((s, p_idx + 1))

        logical_rows = _build_logical_rows(raw_lines)
        for label, patterns, _ in items:
            if label in results:
                continue
            for r_label, nums, pno in logical_rows:
                if any(re.search(pat, r_label) for pat in patterns):
                    v1 = nums[0] if len(nums) >= 1 else None
                    v2 = nums[1] if len(nums) >= 2 else None
                    # 量级 sanity check：报表单位为元，若本期 > 1000万元 而“上期”绝对值 < 1000，判定误读比例列并拒收
                    if v1 is not None and v2 is not None and abs(v1) > 1e7 and abs(v2) < 1000:
                        v2 = None
                    results[label] = (v1, v2, pno)
                    break

    doc.close()

    # 检查必抽核心科目是否齐全
    missing_req: List[str] = []
    for stmt_name, items in STATEMENT_ITEMS.items():
        for label, _, required in items:
            if required and label not in results:
                missing_req.append(f"{stmt_name}·{label}")

    return results, missing_req


def fetch_latest_quote(code: str, date_str: str) -> Optional[Dict]:
    """可选调用本地 ~/.ifind/ifind_helper.py 获取最新单日行情，失败时优雅降级。"""
    if not code or not date_str:
        return None
    helper_dir = os.path.expanduser("~/.ifind")
    if not os.path.exists(os.path.join(helper_dir, "ifind_helper.py")):
        return None
    try:
        sys.path.insert(0, helper_dir)
        from ifind_helper import quotes_as_dict  # type: ignore

        d = quotes_as_dict([code], date_str)
        return d.get(code)
    except Exception as e:
        print(f"[WARN] iFinD 行情拉取跳过（{e}）")
        return None


def fmt_yi(val_yuan: Optional[float]) -> str:
    """将元转换为亿元（保留 2 位小数）。"""
    if val_yuan is None:
        return "—"
    return f"{val_yuan / 1e8:.2f}"


def generate_snapshot_markdown(
    company_name: str,
    stock_code: str,
    pdf_path: str,
    extracted: Dict[str, Tuple[Optional[float], Optional[float], int]],
    missing_req: List[str],
    quote: Optional[Dict],
    quote_date: str,
) -> str:
    """将分节精确提取结果与自动衍生比率渲染为标准化核对表 Markdown。"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = [
        f"# {company_name}（{stock_code}）最新一期合并报表核心科目核对快照",
        "",
        f"> 生成时间：{now}｜一手 PDF 来源：`{os.path.basename(pdf_path)}`（严格按合并三大表分节列解析，跳过首页同比列）",
        "> **用途**：常规财务可优先走 iFinD，本表用于核查四费/财务费用/资产负债表/现金流附注明细并喂给 Step 3.5 Opus 审稿。",
        "",
    ]

    if missing_req:
        lines.append(f"> 🔴 **缺失必抽科目告警**：{missing_req}")
        lines.append("")

    if quote:
        lines.extend([
            f"## 一、最新行情参考（同花顺 iFinD，{quote_date}）",
            f"- 收盘价：`{quote.get('close', '—')}` 元｜涨跌幅：`{quote.get('changeRatio', '—')}%`",
            "",
        ])

    lines.extend([
        "## 二、最新一期合并三大表核心科目（单位：亿元；原表单位：元）",
        "",
        "| 合并报表 | 科目 | 本期/期末（亿元） | 上期/期初（亿元） | 同比/较期初变动 | 原PDF页码 |",
        "|---|---|---|---|---|---|",
    ])

    for stmt_name, items in STATEMENT_ITEMS.items():
        for label, _, _ in items:
            if label not in extracted:
                continue
            v1, v2, pno = extracted[label]
            chg_str = "—"
            if v1 is not None and v2 is not None and abs(v2) > 1e-6:
                pct = (v1 - v2) / abs(v2) * 100.0
                diff_yi = (v1 - v2) / 1e8
                chg_str = f"{diff_yi:+.2f}亿 ({pct:+.2f}%)"
            lines.append(f"| {stmt_name} | {label} | {fmt_yi(v1)} | {fmt_yi(v2)} | {chg_str} | P.{pno} |")

    # 自动计算关键比率，防口算误差
    lines.extend(["", "## 三、自动计算关键财务比率", ""])
    rev1, rev2, _ = extracted.get("营业总收入", (None, None, 0))
    cost1, cost2, _ = extracted.get("营业成本", (None, None, 0))

    if rev1 and cost1:
        gm1 = (rev1 - cost1) / rev1 * 100.0
        s = f"- **本期累计综合毛利率**：`{gm1:.2f}%`"
        if rev2 and cost2:
            gm2 = (rev2 - cost2) / rev2 * 100.0
            s += f"（上年同期 `{gm2:.2f}%`，同比 `{gm1 - gm2:+.2f}pct`）"
        lines.append(s)

    for exp_name in ("销售费用", "管理费用", "研发费用", "财务费用"):
        e1, e2, _ = extracted.get(exp_name, (None, None, 0))
        if rev1 and e1 is not None:
            r1 = e1 / rev1 * 100.0
            s = f"- **{exp_name}率**：`{r1:.2f}%`"
            if rev2 and e2 is not None:
                r2 = e2 / rev2 * 100.0
                s += f"（上年同期 `{r2:.2f}%`，同比 `{r1 - r2:+.2f}pct`）"
            lines.append(s)

    assets1, assets2, _ = extracted.get("资产总计", (None, None, 0))
    liab1, liab2, _ = extracted.get("负债合计", (None, None, 0))
    if assets1 and liab1:
        dr1 = liab1 / assets1 * 100.0
        s = f"- **期末资产负债率**：`{dr1:.2f}%`"
        if assets2 and liab2:
            dr2 = liab2 / assets2 * 100.0
            s += f"（期初 `{dr2:.2f}%`，变动 `{dr1 - dr2:+.2f}pct`）"
        lines.append(s)

    return "\n".join(lines) + "\n"


def main() -> None:
    """命令行入口：按合并报表分节提取核心科目；缺必抽科目时显式非零报错。"""
    parser = argparse.ArgumentParser(description="国金深度报告 Step 1 合并报表核心科目快照提取脚本")
    parser.add_argument("--name", required=True, help="公司简称（如 超颖电子）")
    parser.add_argument("--code", required=True, help="股票代码（如 603175.SH）")
    parser.add_argument("--pdf", required=True, help="最新一期季报或年报 PDF 路径")
    parser.add_argument("--out", required=True, help="输出快照 Markdown 路径")
    parser.add_argument("--date", default="", help="可选：拉取 iFinD 收盘行情的交易日 YYYY-MM-DD")
    parser.add_argument("--allow-partial", action="store_true", help="允许缺失部分必抽科目而不返回非零退出码")
    args = parser.parse_args()

    extracted, missing_req = extract_financials_from_pdf(args.pdf)
    quote = fetch_latest_quote(args.code, args.date) if args.date else None
    md = generate_snapshot_markdown(args.name, args.code, args.pdf, extracted, missing_req, quote, args.date)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(md)

    total_defined = sum(len(v) for v in STATEMENT_ITEMS.values())
    print(f"[OK] 已从合并三大表精确提取 {len(extracted)}/{total_defined} 项科目 -> {args.out}")
    if missing_req:
        print(f"🔴 [必抽核心科目缺失] 未提取到：{missing_req}（请检查报表格式或页码！）")
        if not args.allow_partial:
            sys.exit(1)


if __name__ == "__main__":
    main()
