# -*- coding: utf-8 -*-
"""
免账号跨平台公开数据兜底抓取脚本 (fetch_public_data.py)

运行原理（大白话说明）：
本脚本专为**本机没有 iFinD / Choice / Wind 付费终端**的外部用户与跨平台环境（Windows / macOS / Linux / 云沙盒）设计：
1. **A 股（--market a，纯标准库零依赖）**：
   - 直连巨潮资讯网（cninfo.com.cn）官方公开接口，无过滤拉取公司**近 N 天（默认 90 天）全部临时公告标题与 PDF 直链**；
   - 自动检索近 14 个月最新定期报告（年报 / 半年报 / 季报）PDF 直链，支持 `--download-reports` 一键落盘供 `pull_company_data.py` 抽数；
   - 若本机可选安装了 `akshare`，自动补充近 5 年财务指标速览；未安装或云沙盒限流时优雅跳过不报错。
2. **美股（--market us，纯标准库零依赖）**：
   - 直连美国 SEC EDGAR 官方公开 API（无需 API Key），自动将股票代码转为 CIK；
   - 无过滤拉取**近 N 天（默认 90 天）全部申报文件（20-F / 10-K / 10-Q / 6-K / 8-K / Form 4 等）及官方直链**；
   - 同步解析 SEC XBRL 官方财报事实（Revenue / NetIncome），若安装了 `yfinance` 则补充最新股价与市值。
3. **港股（--market hk）**：
   - 提供港交所披露易（hkexnews.hk）检索入口，并在安装了 `yfinance` 时补充最新股价、市值与年度财报速览。

质控合规声明：
- 巨潮资讯网 PDF 与 SEC EDGAR 申报文件为**官方一手公开出处**，可直接作为工作底稿红框截图原件；
- `akshare` / `yfinance` 等免费二手接口仅用于 Step 1 快速摸底，**正式工作底稿严禁免截图引用**，须回落至官方 PDF 原件或付费金融终端（iFinD / Choice / Wind / Bloomberg）复核。

命令行用法（在仓库根目录运行）：
    python 脚本与规范/fetch_public_data.py --code 600584.SH --out 示例公司/00_原始资料/00_公开公告与财务速览.md
    python 脚本与规范/fetch_public_data.py --code 600584.SH --out 示例公司/00_原始资料/00_公开公告与财务速览.md --download-reports 示例公司/00_原始资料/pdfs
    python 脚本与规范/fetch_public_data.py --code ATAT --market us --out 亚朵集团/00_原始资料/00_公开公告与财务速览.md
"""

import argparse
import datetime as dt
import json
import os
import re
import ssl
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, List, Optional, Tuple


# 巨潮资讯网与 SEC EDGAR 合规请求头（SEC 强制要求 User-Agent 含应用名与联系邮箱格式）
CNINFO_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": "http://www.cninfo.com.cn/",
    "Accept": "application/json, text/plain, */*",
}
SEC_HEADERS = {
    "User-Agent": "DeepReportWorkflow ResearchScript/1.0 (research-workflow@example.com)",
    "Accept-Encoding": "identity",
    "Accept": "application/json",
}


def detect_market(code: str, market_arg: str = "auto") -> str:
    """根据股票代码后缀与形态自动识别市场（a / us / hk）。"""
    m = (market_arg or "auto").strip().lower()
    if m in ("a", "us", "hk"):
        return m
    c = code.strip().upper()
    if c.endswith((".SH", ".SZ", ".BJ")) or re.fullmatch(r"\d{6}", c):
        return "a"
    if c.endswith(".HK") or re.fullmatch(r"\d{4,5}", c):
        return "hk"
    return "us"


def normalize_a_code(code: str) -> str:
    """将 600584.SH / 002475.SZ / 600584 统一清洗为 6 位数字代码。"""
    m = re.search(r"(\d{6})", code.strip())
    return m.group(1) if m else code.strip()


def http_get_json(
    url: str,
    headers: Dict[str, str],
    data: Optional[bytes] = None,
    method: str = "GET",
    timeout: int = 20,
) -> Any:
    """通用标准库 HTTP JSON 请求封装，兼容 GET/POST 并统一 UTF-8 解码。"""
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
    return json.loads(raw)


def clean_html_tags(text: str) -> str:
    """清除巨潮高亮标题中的 <em> 等 HTML 标签，保留干净中文标题。"""
    return re.sub(r"<[^>]+>", "", text or "").strip()


def ms_to_date_str(ms_val: Any) -> str:
    """将毫秒时间戳转换为 YYYY-MM-DD 字符串。"""
    try:
        ts = float(ms_val) / 1000.0
        return dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d")
    except Exception:
        return ""


# ==============================================================================
# 一、A 股模块：巨潮资讯网官方公告 + 定期报告 PDF + 可选 AkShare
# ==============================================================================


def cninfo_resolve_org(code6: str) -> Tuple[Optional[str], str]:
    """通过巨潮搜索接口将 6 位 A 股代码解析为 (orgId, 公司简称)。"""
    q = urllib.parse.urlencode({"keyWord": code6, "maxNum": 10})
    url = f"http://www.cninfo.com.cn/new/information/topSearch/query?{q}"
    items = http_get_json(url, headers=CNINFO_HEADERS, method="POST")
    if not isinstance(items, list) or not items:
        return None, code6
    # 优先匹配精确代码且为 A 股的主体（避免误命中可转债/债券）
    for it in items:
        if str(it.get("code", "")).strip() == code6 and "A股" in str(it.get("category", "")):
            return str(it.get("orgId", "")), str(it.get("zwjc", code6))
    for it in items:
        if str(it.get("code", "")).strip() == code6:
            return str(it.get("orgId", "")), str(it.get("zwjc", code6))
    return str(items[0].get("orgId", "")), str(items[0].get("zwjc", code6))


def cninfo_column_for_code(code6: str) -> str:
    """根据 6 位代码首字符返回巨潮交易所栏目名（sse / szse / third）。"""
    if code6.startswith(("6", "9")):
        return "sse"
    if code6.startswith(("0", "2", "3")):
        return "szse"
    return "third"


def cninfo_fetch_announcements(
    code6: str,
    org_id: str,
    start_date: str,
    end_date: str,
    category: str = "",
    max_items: int = 120,
) -> List[Dict[str, str]]:
    """分页拉取巨潮官方公告列表（不加任何关键词过滤，确保近 90 天全量覆盖）。"""
    col = cninfo_column_for_code(code6)
    results: List[Dict[str, str]] = []
    page_num = 1
    page_size = 30

    while len(results) < max_items:
        payload = urllib.parse.urlencode(
            {
                "stock": f"{code6},{org_id}",
                "tabName": "fulltext",
                "pageSize": page_size,
                "pageNum": page_num,
                "column": col,
                "category": category,
                "plate": "",
                "seDate": f"{start_date}~{end_date}",
                "searchkey": "",
                "secid": "",
                "sortName": "",
                "sortType": "",
                "isHLtitle": "false",
            }
        ).encode("utf-8")
        headers = {**CNINFO_HEADERS, "Content-Type": "application/x-www-form-urlencoded"}
        data = http_get_json(
            "http://www.cninfo.com.cn/new/hisAnnouncement/query",
            headers=headers,
            data=payload,
            method="POST",
        )
        anns = data.get("announcements") or []
        if not anns:
            break
        for a in anns:
            adj = (a.get("adjunctUrl") or "").lstrip("/")
            pdf_url = f"http://static.cninfo.com.cn/{adj}" if adj else ""
            results.append(
                {
                    "date": ms_to_date_str(a.get("announcementTime")),
                    "title": clean_html_tags(a.get("announcementTitle", "")),
                    "url": pdf_url,
                }
            )
            if len(results) >= max_items:
                break
        if not data.get("hasMore"):
            break
        page_num += 1
    return results


def download_cninfo_core_pdfs(periodic_list: List[Dict[str, str]], target_dir: str, company_name: str) -> List[str]:
    """按需下载最新 1 份年报全文 + 最新 1 份季报/半年报全文 PDF 到本地目录。"""
    os.makedirs(target_dir, exist_ok=True)
    picked: List[Dict[str, str]] = []
    got_annual = False
    got_interim = False

    for item in periodic_list:
        title = item["title"]
        url = item["url"]
        if not url.lower().endswith(".pdf"):
            continue
        # 过滤掉摘要、英文版、更正公告、取消公告，只下正式正文 PDF
        if any(skip in title for skip in ("摘要", "英文", "更正", "提示性", "关于")):
            continue
        if "年度报告" in title and not got_annual:
            picked.append(item)
            got_annual = True
        elif any(k in title for k in ("季度报告", "半年度报告")) and not got_interim:
            picked.append(item)
            got_interim = True
        if got_annual and got_interim:
            break

    saved: List[str] = []
    for item in picked:
        safe_title = re.sub(r'[\\/:*?"<>|\s]+', "_", item["title"])[:50]
        fname = f"{company_name}_{item['date'].replace('-', '')}_{safe_title}.pdf"
        out_fp = os.path.join(target_dir, fname)
        try:
            req = urllib.request.Request(item["url"], headers=CNINFO_HEADERS)
            with urllib.request.urlopen(req, timeout=45) as resp:
                with open(out_fp, "wb") as wf:
                    wf.write(resp.read())
            saved.append(out_fp)
        except Exception as e:
            print(f"[WARN] 下载定期报告 PDF 失败（{item['title']}）：{e}")
    return saved


def try_akshare_a_summary(code6: str) -> str:
    """可选调用 akshare 拉取 A 股近 5 期核心财务摘要，未安装或网络受限时优雅跳过。"""
    try:
        import akshare as ak  # type: ignore
    except ImportError:
        return "（未安装可选依赖 `akshare`，已跳过免费二手财务拉取；如需启用可运行 `pip install akshare`，或直接使用官方定期报告 PDF + `pull_company_data.py`。）"

    try:
        df = ak.stock_financial_abstract_ths(symbol=code6, indicator="按报告期")
        if df is None or df.empty:
            return "（`akshare` 返回空财务表，请改用官方定期报告 PDF 或付费金融终端。）"
        cols = [c for c in ["报告期", "营业总收入", "净利润", "扣非净利润", "销售毛利率", "销售净利率", "净资产收益率"] if c in df.columns]
        if not cols:
            cols = list(df.columns[:7])
        sub = df[cols].head(8)
        lines = ["| " + " | ".join(str(c) for c in cols) + " |", "| " + " | ".join("---" for _ in cols) + " |"]
        for _, row in sub.iterrows():
            lines.append("| " + " | ".join(str(row[c]) for c in cols) + " |")
        return "\n".join(lines)
    except Exception as e:
        return f"（`akshare` 拉取跳过：{e}；请改用巨潮官方 PDF + `pull_company_data.py` 或付费金融终端。）"


# ==============================================================================
# 二、美股模块：SEC EDGAR 官方申报清单 + XBRL 核心财报事实 + 可选 yfinance
# ==============================================================================


def sec_resolve_cik(ticker: str) -> Tuple[Optional[str], str]:
    """通过 SEC 官方 company_tickers.json 将美股 ticker 解析为 10 位补零 CIK 与公司全称。"""
    t_up = ticker.strip().upper().split(".")[0]
    data = http_get_json("https://www.sec.gov/files/company_tickers.json", headers=SEC_HEADERS)
    if isinstance(data, dict):
        for entry in data.values():
            if str(entry.get("ticker", "")).upper() == t_up:
                cik_str = f"{int(entry['cik_str']):010d}"
                return cik_str, str(entry.get("title", t_up))
    return None, t_up


def parse_sec_submissions(sub_json: Dict[str, Any], cik_padded: str, days: int = 90) -> Tuple[List[Dict[str, str]], List[Dict[str, str]]]:
    """
    解析 SEC submissions JSON，返回：
    1) 近 days 天全部申报记录（无过滤全量清单）；
    2) 最近 6 份核心定期/业绩申报（10-K / 20-F / 10-Q / 6-K / 8-K）。
    """
    recent = (sub_json.get("filings") or {}).get("recent") or {}
    forms = recent.get("form") or []
    dates = recent.get("filingDate") or []
    accs = recent.get("accessionNumber") or []
    docs = recent.get("primaryDocument") or []
    descs = recent.get("primaryDocDescription") or []

    cutoff = (dt.date.today() - dt.timedelta(days=days)).isoformat()
    cik_int = str(int(cik_padded))
    recent_all: List[Dict[str, str]] = []
    major_periodic: List[Dict[str, str]] = []

    for i in range(min(len(forms), len(dates), len(accs), len(docs))):
        form = str(forms[i])
        fdate = str(dates[i])
        acc_nodash = str(accs[i]).replace("-", "")
        doc = str(docs[i])
        desc = str(descs[i]) if i < len(descs) else ""
        url = f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc_nodash}/{doc}" if doc else ""
        rec = {"date": fdate, "form": form, "desc": desc or form, "url": url}
        if fdate >= cutoff:
            recent_all.append(rec)
        if form in ("10-K", "20-F", "10-Q", "6-K", "8-K") and len(major_periodic) < 8:
            major_periodic.append(rec)

    return recent_all, major_periodic


def extract_sec_xbrl_highlights(facts_json: Dict[str, Any]) -> str:
    """从 SEC EDGAR XBRL companyfacts JSON 中提取近 5 年官方年度营收与净利润（兼容 US-GAAP 与 IFRS）。"""
    facts = facts_json.get("facts") or {}
    rev_tags = [
        ("us-gaap", "Revenues"),
        ("us-gaap", "RevenueFromContractWithCustomerExcludingAssessedTax"),
        ("us-gaap", "SalesRevenueNet"),
        ("ifrs-full", "Revenue"),
    ]
    ni_tags = [
        ("us-gaap", "NetIncomeLoss"),
        ("us-gaap", "ProfitLoss"),
        ("ifrs-full", "ProfitLoss"),
    ]

    def _pick_annual(tag_candidates: List[Tuple[str, str]]) -> Tuple[str, Dict[str, float]]:
        for ns, tag in tag_candidates:
            node = (facts.get(ns) or {}).get(tag)
            if not node:
                continue
            units = node.get("units") or {}
            # 优先取 USD 或 CNY 等货币单位
            for unit_name, entries in units.items():
                if unit_name.upper() not in ("USD", "CNY", "RMB", "HKD", "EUR"):
                    continue
                by_fy: Dict[str, float] = {}
                for e in entries:
                    form = str(e.get("form", ""))
                    fy = e.get("fy")
                    fp = str(e.get("fp", ""))
                    val = e.get("val")
                    if form in ("10-K", "20-F") and fp == "FY" and fy and isinstance(val, (int, float)):
                        by_fy[str(fy)] = float(val)
                if by_fy:
                    return unit_name, by_fy
        return "", {}

    rev_unit, rev_map = _pick_annual(rev_tags)
    ni_unit, ni_map = _pick_annual(ni_tags)
    years = sorted(set(rev_map.keys()) | set(ni_map.keys()), reverse=True)[:5]
    if not years:
        return "（SEC XBRL 未命中标准年度营收/净利润标签，请直接下载上方 20-F/10-K/6-K 官方文件或使用 Choice/Bloomberg。）"

    lines = [
        f"| 财年 | 营业收入（百万 {rev_unit or '原币'}） | 净利润（百万 {ni_unit or '原币'}） |",
        "| --- | --- | --- |",
    ]
    for y in years:
        rv = f"{rev_map[y] / 1e6:,.2f}" if y in rev_map else "—"
        nv = f"{ni_map[y] / 1e6:,.2f}" if y in ni_map else "—"
        lines.append(f"| FY{y} | {rv} | {nv} |")
    return "\n".join(lines)


def try_yfinance_summary(ticker: str) -> str:
    """可选调用 yfinance 拉取美股/港股最新收盘价、市值与估值速览，未安装时优雅跳过。"""
    try:
        import yfinance as yf  # type: ignore
    except ImportError:
        return "（未安装可选依赖 `yfinance`，已跳过免费行情拉取；如需启用可运行 `pip install yfinance`。）"

    try:
        tk = yf.Ticker(ticker)
        info = tk.info or {}
        price = info.get("regularMarketPrice") or info.get("currentPrice")
        mcap = info.get("marketCap")
        ccy = info.get("currency", "")
        pe = info.get("trailingPE")
        if price is None and mcap is None:
            return "（`yfinance` 未返回有效行情，请改用 Choice / Bloomberg / 官方财报。）"
        mcap_str = f"{mcap / 1e8:,.2f} 亿 {ccy}" if isinstance(mcap, (int, float)) else "—"
        pe_str = f"{pe:.2f}x" if isinstance(pe, (int, float)) else "—"
        return f"- 最新价格：{price} {ccy} | 总市值：{mcap_str} | Trailing PE：{pe_str}（来源：Yahoo Finance，仅供参考，不可免截图入底稿）"
    except Exception as e:
        return f"（`yfinance` 拉取跳过：{e}）"


# ==============================================================================
# 三、Markdown 报告组装与主入口
# ==============================================================================


def build_output_markdown(
    company_label: str,
    code: str,
    market: str,
    days: int,
    sections: List[Tuple[str, str]],
) -> str:
    """组装结构化公开数据速览 Markdown，顶部强制附合规边界声明。"""
    today_s = dt.date.today().isoformat()
    header = (
        f"# {company_label}（{code}）公开公告与财务速览（免账号跨平台兜底版）\n\n"
        f"- **拉取日期**：{today_s} | **市场**：{market.upper()} | **公告回溯窗口**：近 {days} 天（无关键词过滤）\n"
        "- **质控合规边界**：\n"
        "  1. 巨潮资讯网 PDF 直链与 SEC EDGAR 官方申报文件属于**官方一手公开出处**，可直接下载用于 `pull_company_data.py` 抽数与底稿红框截图；\n"
        "  2. `akshare` / `yfinance` 等免费二手数据仅用于 Step 1 快速摸底与初稿作图，**严禁在工作底稿中免截图使用**，正式底稿须回落到官方 PDF 原件或付费金融终端（iFinD / Choice / Wind / Bloomberg）复核。\n"
    )
    body_parts = [header]
    for title, content in sections:
        body_parts.append(f"\n## {title}\n\n{content.strip()}\n")
    return "\n".join(body_parts)


def run_fetch(
    code: str,
    market_arg: str,
    out_path: str,
    days: int = 90,
    download_dir: str = "",
) -> str:
    """执行跨市场公开数据抓取并输出 Markdown 文件。"""
    market = detect_market(code, market_arg)
    end_d = dt.date.today()
    start_d = end_d - dt.timedelta(days=days)
    sections: List[Tuple[str, str]] = []
    company_label = code

    if market == "a":
        code6 = normalize_a_code(code)
        try:
            org_id, cname = cninfo_resolve_org(code6)
            company_label = cname
            if not org_id:
                sections.append(("一、近 90 天全量临时公告（巨潮资讯网）", f"未在巨潮检索到代码 `{code6}` 的主体 `orgId`，请核对股票代码。"))
            else:
                anns = cninfo_fetch_announcements(code6, org_id, start_d.isoformat(), end_d.isoformat(), max_items=120)
                if anns:
                    rows = ["| 披露日期 | 公告标题 | 官方 PDF 直链 |", "| --- | --- | --- |"]
                    for a in anns:
                        rows.append(f"| {a['date']} | {a['title']} | {a['url']} |")
                    sections.append((f"一、近 {days} 天全量临时公告清单（共 {len(anns)} 条，无过滤）", "\n".join(rows)))
                else:
                    sections.append((f"一、近 {days} 天全量临时公告清单", "（该窗口期内无新增公告或接口返回为空）"))

                # 另拉近 400 天定期报告（年报/半年报/季报），便于直接下载 PDF
                p_start = (end_d - dt.timedelta(days=400)).isoformat()
                periodic = cninfo_fetch_announcements(
                    code6,
                    org_id,
                    p_start,
                    end_d.isoformat(),
                    category="category_ndbg_szsh;category_bndbg_szsh;category_yjdbg_szsh;category_sjdbg_szsh",
                    max_items=20,
                )
                if periodic:
                    prows = ["| 披露日期 | 定期报告标题 | 官方 PDF 直链 |", "| --- | --- | --- |"]
                    for p in periodic[:10]:
                        prows.append(f"| {p['date']} | {p['title']} | {p['url']} |")
                    extra = ""
                    if download_dir:
                        saved = download_cninfo_core_pdfs(periodic, download_dir, cname)
                        if saved:
                            extra = "\n\n**已自动下载到本地的核心定期报告 PDF**：\n" + "\n".join(f"- `{fp}`" for fp in saved)
                    sections.append(("二、最新定期报告 PDF 直链（近 14 个月）", "\n".join(prows) + extra))
        except Exception as e:
            sections.append(
                (
                    f"一、近 {days} 天全量临时公告（巨潮资讯网）",
                    f"（直连巨潮受限或超时：`{e}`。如处于受限云沙盒，请改用 Choice/Wind 连接器，或手工从巨潮资讯网下载最新年报/季报 PDF 放入 `00_原始资料/pdfs/`。）",
                )
            )

        sections.append(("三、可选免费财务指标速览（AkShare，非底稿豁免源）", try_akshare_a_summary(code6)))

    elif market == "us":
        ticker = code.strip().upper().split(".")[0]
        try:
            cik, cname = sec_resolve_cik(ticker)
            company_label = cname
            if not cik:
                sections.append(("一、SEC EDGAR 官方申报清单", f"未在 SEC `company_tickers.json` 中找到 `{ticker}`，请核对美股代码。"))
            else:
                sub = http_get_json(f"https://data.sec.gov/submissions/CIK{cik}.json", headers=SEC_HEADERS)
                recent_all, major_periodic = parse_sec_submissions(sub, cik, days=days)
                if recent_all:
                    rows = ["| 申报日期 | 表格类型 | 说明 | SEC 官方直链 |", "| --- | --- | --- | --- |"]
                    for r in recent_all:
                        rows.append(f"| {r['date']} | {r['form']} | {r['desc']} | {r['url']} |")
                    sections.append((f"一、近 {days} 天 SEC EDGAR 全量申报清单（CIK {cik}，共 {len(recent_all)} 条）", "\n".join(rows)))
                else:
                    sections.append((f"一、近 {days} 天 SEC EDGAR 全量申报清单（CIK {cik}）", "（窗口期内无新申报）"))

                if major_periodic:
                    mrows = ["| 申报日期 | 表格类型 | SEC 官方直链 |", "| --- | --- | --- |"]
                    for m in major_periodic:
                        mrows.append(f"| {m['date']} | {m['form']} | {m['url']} |")
                    sections.append(("二、最近核心定期/业绩申报（20-F / 10-K / 10-Q / 6-K / 8-K）", "\n".join(mrows)))

                try:
                    facts = http_get_json(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json", headers=SEC_HEADERS)
                    sections.append(("三、SEC XBRL 官方年度核心财报事实", extract_sec_xbrl_highlights(facts)))
                except Exception as fe:
                    sections.append(("三、SEC XBRL 官方年度核心财报事实", f"（XBRL 提取跳过：`{fe}`，请直接查阅上方 20-F/10-K 原文。）"))
        except Exception as e:
            sections.append(
                (
                    "一、SEC EDGAR 官方申报清单",
                    f"（访问 SEC EDGAR 受限或超时：`{e}`。请改用 Choice / Bloomberg 连接器，或下载 20-F/10-K PDF 至 `00_原始资料/pdfs/`。）",
                )
            )
        sections.append(("四、可选行情与估值速览（yfinance，非底稿豁免源）", try_yfinance_summary(ticker)))

    else:
        # 港股：引导披露易官方一手公告 + 可选 yfinance 行情
        sections.append(
            (
                "一、港股官方一手公告入口（香港交易所披露易）",
                f"- 请在港交所披露易（https://www.hkexnews.hk/）检索 `{code}` 近 {days} 天全部公告及最新年报/中报 PDF，落盘至 `00_原始资料/pdfs/`。",
            )
        )
        sections.append(("二、可选行情与估值速览（yfinance，非底稿豁免源）", try_yfinance_summary(code)))

    md = build_output_markdown(company_label, code, market, days, sections)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as wf:
        wf.write(md)
    print(f"[OK] 已生成公开公告与财务速览 -> {out_path}")
    return out_path


def main() -> None:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="免账号跨平台公开数据兜底抓取脚本（巨潮 A 股 / SEC EDGAR 美股 / 可选 AkShare & yfinance）")
    parser.add_argument("--code", required=True, help="股票代码（如 600584.SH / 002475.SZ / ATAT / 0700.HK）")
    parser.add_argument("--market", default="auto", choices=["auto", "a", "us", "hk"], help="市场类型（默认 auto 自动识别）")
    parser.add_argument("--days", type=int, default=90, help="全量临时公告回溯天数（默认 90 天）")
    parser.add_argument("--out", required=True, help="输出 Markdown 路径（如 示例公司/00_原始资料/00_公开公告与财务速览.md）")
    parser.add_argument("--download-reports", default="", help="可选（仅 A 股）：自动下载最新 1 份年报 + 1 份季报/半年报 PDF 的目标目录")
    args = parser.parse_args()

    run_fetch(
        code=args.code,
        market_arg=args.market,
        out_path=args.out,
        days=args.days,
        download_dir=args.download_reports,
    )


if __name__ == "__main__":
    main()
