# -*- coding: utf-8 -*-
"""
国金证券深度报告——三端【期间×数值】有序对账、审稿版本绑定与发布门禁脚本 (sync_check.py)

核心升级（彻底堵死“年份互换不报错”“正文表格与底稿金额脱节”“审稿意见未绑定终稿版本”漏洞）：
1. **正文图表 vs charts.json vs 07_图表/ 核对**：检查图片引用完整性并支持 `--clean-orphans` 一键清孤儿旧图；
2. **【期间/分类 × 数值】有序对账（charts.json ↔ 06_证据清单.json）**：
   不再用无序扁平集合！而是把 `charts.json` 的 `(category, value)` 绑定到底稿 `raw_data` 对应表头列（或首列行）上，
   **哪怕只把 2024 与 2025 两年数值互换，也会立刻亮 🔴 报出精确错配期间**；
3. **正文 Markdown 表格 ↔ 底稿 table 逐行数值对账**：
   解析 `05_完整初稿.md` 中每张 `<!-- TABLE_META -->` 下方的完整 Markdown 表格，逐行提取关键数字与底稿 `type: "table"` 的 `raw_data` 交叉比对，杜绝正文表格和底稿两张皮；
4. **红队审稿版本哈希绑定校验（`draft_sha256` + `verdict`）**：
   自动计算当前 `05_完整初稿.md + report_meta.json` 的 SHA-256 指纹，比对 `05_红队审稿意见_Claude.md` 的 `REVIEW_META` 头；
   若修订后尚未复审或结论为 `REJECT`，草稿模式打印 🟡 提醒，`--release` 正式发布模式直接 🔴 阻断；
5. **首页摘要字数护栏（≤ 850 字）与 `[待模型填入]` 占位符检查**（`--release` 模式零容忍阻断）。

命令行用法：
    python 脚本与规范/sync_check.py --dir 示例公司 --clean-orphans --strict
    python 脚本与规范/sync_check.py --dir 示例公司 --release
"""

import argparse
import datetime
import hashlib
import json
import os
import re
import sys
from typing import Any, Dict, List, Optional, Set, Tuple


PLACEHOLDER_RE = re.compile(r"\[待模型填入\]|\[待补充[^\]]*\]|\[待确认[^\]]*\]|\bTBD\b|\bTODO\b")


def _norm_key(s: Any) -> str:
    """归一化期间/分类/公司名，用于跨表精确对齐（通用正则将 20XX 归一为 XX，不限年份区间）。"""
    t = re.sub(r"[\s（）()：:【】\[\]—–\-]", "", str(s or "")).lower()
    t = re.sub(r"20(\d{2})", r"\1", t)
    return t


def _extract_cell_nums(val: Any) -> List[float]:
    """提取单元格中的所有数值（保留 2 位小数；年份过滤由调用方按业务语义按需处理）。"""
    if isinstance(val, (int, float)):
        return [round(float(val), 2)]
    out: List[float] = []
    # 去掉千分位逗号后抽取数字
    cleaned = str(val or "").replace(",", "")
    for tok in re.findall(r"-?\d+(?:\.\d+)?", cleaned):
        try:
            f = round(float(tok), 2)
            out.append(f)
        except ValueError:
            pass
    return out


def compute_draft_sha256(md_text: str, meta_text: str) -> str:
    """与 redteam_review.py 保持完全一致的稿件 SHA-256 指纹（前 16 位）。"""
    h = hashlib.sha256()
    h.update(md_text.encode("utf-8"))
    h.update(b"\n---META---\n")
    h.update(meta_text.encode("utf-8"))
    return h.hexdigest()[:16]


def parse_md_chart_titles(md_text: str) -> List[Tuple[str, str]]:
    """提取正文中每个 CHART / CHART_PAIR 注释块对应的 (图片文件名, 图表标题)，供年份区间一致性核查。"""
    pairs: List[Tuple[str, str]] = []
    # 解析单图 CHART 块
    for blk in re.finditer(r"<!--\s*CHART\b(.*?)-->", md_text, flags=re.S):
        body = blk.group(1)
        m_img = re.search(r"^\s*image\s*:\s*(.+)$", body, flags=re.M)
        m_tit = re.search(r"^\s*title\s*:\s*(.+)$", body, flags=re.M)
        if m_img and m_tit:
            pairs.append((os.path.basename(m_img.group(1).strip()), m_tit.group(1).strip()))
    # 解析双图 CHART_PAIR 块
    for blk in re.finditer(r"<!--\s*CHART_PAIR\b(.*?)-->", md_text, flags=re.S):
        body = blk.group(1)
        for side in ("left", "right"):
            m_img = re.search(rf"^\s*{side}_image\s*:\s*(.+)$", body, flags=re.M)
            m_tit = re.search(rf"^\s*{side}_title\s*:\s*(.+)$", body, flags=re.M)
            if m_img and m_tit:
                pairs.append((os.path.basename(m_img.group(1).strip()), m_tit.group(1).strip()))
    return pairs


def audit_chart_title_year_range(
    chart_title_pairs: List[Tuple[str, str]],
    cfg_by_file: Dict[str, Dict[str, Any]],
) -> List[str]:
    """
    核查正文图表标题中标注的起止年份区间（如“2023-2026E”“2021至2026Q3”）
    是否与 charts.json 对应图表的实际横轴年份范围严格一致；不一致直接亮 🔴 阻断。
    """
    errs: List[str] = []
    # 匹配标题中的起止年份区间：支持 2023-2026E、2021–2026Q3、2022至2025年、2025年4月–2026年6月、FY2024-FY2026E
    range_re = re.compile(
        r"(?:FY)?(20\d{2})(?:年(?:\d{1,2}月)?|Q[1-4]|H[12]|E)?\s*(?:[-–—~]|至)\s*(?:FY)?(20\d{2})(?:年(?:\d{1,2}月)?|Q[1-4]|H[12]|E)?"
    )

    for fn, title in chart_title_pairs:
        cfg = cfg_by_file.get(fn)
        if not cfg:
            continue
        cats = cfg.get("categories") or []
        if len(cats) < 2:
            continue
        m = range_re.search(title)
        if not m:
            continue
        title_start, title_end = int(m.group(1)), int(m.group(2))
        # 从 categories 中抽取全部横轴年份（兼容 "2024", "2026Q1-Q3", "25Q1", "FY2025"）
        axis_years: List[int] = []
        for c in cats:
            cs = str(c)
            m4 = re.search(r"(20\d{2})", cs)
            if m4:
                axis_years.append(int(m4.group(1)))
            else:
                m2 = re.match(r"^(\d{2})[QH]", cs)
                if m2:
                    axis_years.append(2000 + int(m2.group(1)))
        # 仅当横轴确实是年份/季度时间序列（至少包含 2 个有效年份标签）时执行比对
        if len(axis_years) < 2:
            continue
        min_y, max_y = min(axis_years), max(axis_years)
        if min_y != title_start or max_y != title_end:
            errs.append(
                f"🔴 [图表标题年份与横轴不符] {fn}：标题写「{title_start}–{title_end}」（{title}），"
                f"但 charts.json 实际横轴年份范围为 {min_y}–{max_y}（categories={cats}）！"
            )
    return errs


def audit_style_guardrails(md_text: str) -> List[str]:
    """
    依据 2026-11-24 沉淀的永久写作红线做轻量提示扫描：
    1. 禁止分析财务费用率（命中“财务费用率”即输出 🟡 提醒，附行号）；
    2. 避免单段正文连续堆砌一连串逐年/逐季数字清单（单段独立带单位数值/百分数 ≥ 9 处时输出 🟡 提醒）。
    """
    warns: List[str] = []
    num_unit_re = re.compile(r"\d+(?:\.\d+)?(?:%|pct|个百分点|亿元|万元|亿美元|百万美元|亿日元|万股)")
    in_comment = False
    for idx, raw_line in enumerate(md_text.splitlines(), 1):
        s = raw_line.strip()
        if s.startswith("<!--"):
            in_comment = True
        if in_comment:
            if "-->" in s:
                in_comment = False
            continue
        if not s or s.startswith("#") or s.startswith("|") or s.startswith(">"):
            continue
        if "财务费用率" in s:
            warns.append(f"🟡 [文风红线·财务费用率] 正文第 {idx} 行提及「财务费用率」（请按规范删除财务费用率分析）")
        hits = num_unit_re.findall(s)
        if len(hits) >= 9:
            warns.append(
                f"🟡 [文风红线·疑似连续堆砌数字] 正文第 {idx} 行单段出现 {len(hits)} 个带单位数值，建议精简流水账让图表说话"
            )
    return warns


def parse_md_charts_and_tables(md_text: str) -> Tuple[List[str], List[Tuple[str, List[List[str]]]]]:
    """解析 05_完整初稿.md 中的全部图片文件名，以及每张 TABLE_META 对应的 Markdown 表格矩阵。"""
    img_refs: List[str] = []
    for m in re.finditer(r"^(?:image|left_image|right_image)\s*:\s*(.+)$", md_text, flags=re.M):
        img_refs.append(os.path.basename(m.group(1).strip()))

    lines = md_text.splitlines()
    tables: List[Tuple[str, List[List[str]]]] = []
    i = 0
    while i < len(lines):
        if re.match(r"^\s*<!--\s*TABLE_META", lines[i]):
            title = ""
            i += 1
            while i < len(lines) and "-->" not in lines[i]:
                tm = re.match(r"^\s*title\s*:\s*(.+)$", lines[i])
                if tm:
                    title = tm.group(1).strip()
                i += 1
            i += 1
            while i < len(lines) and not lines[i].strip():
                i += 1
            rows: List[List[str]] = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                raw_ln = lines[i].strip()
                if not re.match(r"^\|[\s:\-\|]+\|$", raw_ln):
                    cells = [c.strip() for c in raw_ln.strip("|").split("|")]
                    rows.append(cells)
                i += 1
            tables.append((title, rows))
            continue
        i += 1

    return img_refs, tables


def _is_compatible_period_substring(k1: str, k2: str) -> bool:
    """
    子串回退期间兼容性保护：
    禁止纯年度键（如 25 / 25年 / 25e）与带季度/半年度/月度限定的键（如 25h1 / 25q3 / 25年4月）互相子串匹配，
    防止年度画图值误撞半年度/季度列导致漏报期间错配。
    """
    sub_period_re = re.compile(r"(?:q[1-4]|h[12]|[1-3]q|月|半|季)")
    year_only_re = re.compile(r"^(?:fy)?\d{2}(?:年|e|f|a)?$")
    k1_sub = bool(sub_period_re.search(k1))
    k2_sub = bool(sub_period_re.search(k2))
    if (k1_sub and year_only_re.match(k2)) or (k2_sub and year_only_re.match(k1)):
        return False
    return True


def audit_chart_ordered_values(
    fn: str,
    cfg: Dict[str, Any],
    m_item: Dict[str, Any],
) -> List[str]:
    """
    对单张图表执行【期间/分类 × 数值】有序对齐核查：
    - 同时兼容“横向表头（期间在首行第 1..N 列）”和“纵向表头（分类在首列第 1..M 行）”；
    - 若某期间在 charts.json 中的值未出现在底稿 raw_data 的对应期间列/行中，立即报 🔴！
    """
    errs: List[str] = []
    cats = cfg.get("categories") or []
    series = cfg.get("series") or []
    raw_data = m_item.get("raw_data") or []
    if not cats or not series or len(raw_data) < 2:
        return errs

    header = raw_data[0]
    header_norm = [_norm_key(c) for c in header]
    col0_norm = [_norm_key(r[0]) if r else "" for r in raw_data]

    # 构建底稿按“列期间”（跳过第 0 列行标签列）和按“行分类”的数值映射
    by_col_key: Dict[str, Set[float]] = {}
    for c_idx, hk in enumerate(header_norm[1:], 1):
        vals_set: Set[float] = set()
        for r in raw_data[1:]:
            if c_idx < len(r):
                vals_set.update(_extract_cell_nums(r[c_idx]))
        by_col_key[hk] = vals_set

    by_row_key: Dict[str, Set[float]] = {}
    for r_idx, rk in enumerate(col0_norm[1:], 1):
        vals_set = set()
        for cell in raw_data[r_idx][1:]:
            vals_set.update(_extract_cell_nums(cell))
        by_row_key[rk] = vals_set

    # 逐期间/分类核验每一个数值是否落在正确位置（年份互换必抓！）
    for s in series:
        sname = s.get("name", "")
        svals = s.get("values") or []
        for idx, val in enumerate(svals):
            if idx >= len(cats) or not isinstance(val, (int, float)):
                continue
            target_v = round(float(val), 2)
            ck = _norm_key(cats[idx])

            # 优先精确匹配（列→行），再回退子串匹配（列→行），避免表头短词误吞分类名
            matched_bucket: Optional[Set[float]] = None
            if ck in by_col_key:
                matched_bucket = by_col_key[ck]
            elif ck in by_row_key:
                matched_bucket = by_row_key[ck]
            if matched_bucket is None:
                for hk, bset in by_col_key.items():
                    if hk and (ck in hk or hk in ck) and _is_compatible_period_substring(ck, hk):
                        matched_bucket = bset
                        break
            if matched_bucket is None:
                for rk, bset in by_row_key.items():
                    if rk and (ck in rk or rk in ck) and _is_compatible_period_substring(ck, rk):
                        matched_bucket = bset
                        break

            if matched_bucket is None:
                errs.append(
                    f"🔴 [期间列缺失] {fn}（{m_item.get('label', '')}）：图表期间/分类「{cats[idx]}」在底稿 raw_data 中未找到对应列/行！"
                )
            elif target_v not in matched_bucket:
                errs.append(
                    f"🔴 [期间-数值错配] {fn}（{m_item.get('label', '')}）：「{cats[idx]} / {sname}」画图值为 {target_v}，"
                    f"但底稿该期间列/行实际数值为 {sorted(matched_bucket)}（疑似年份互换或改图忘改底稿）！"
                )
    return errs


def _is_pure_metric_cell(cell_text: str) -> bool:
    """
    判断表格单元格是否为“纯数值指标格”（如 '1,200'、'45%'、'8家'、'120亿元'、'+15pct'），
    以区分需要精确比对的整数业务指标与长句备注中的附带整数。
    """
    s = str(cell_text or "").strip()
    if not s or not re.search(r"\d", s):
        return False
    # 剥离常见数值符号与金融/业务单位后，若不再残留中文或英文字母，视为纯数值格
    stripped = re.sub(
        r"(?:万亿美元|亿美元|百万美元|人民币|亿元|万元|万㎡|万平方米|平方米|亿片|万片|万件|万台|万吨|万套|"
        r"个百分点|pct|bps|bp|usd|rmb|hkd|cny|%|％|倍|x|家|座|条|个|天|次|年|月|元|港元|美元|约|超|近|约合|以上|以下|左右|"
        r"[\s,+\-–—~≈<>≤≥/()（）\[\]·])",
        "",
        s,
        flags=re.I,
    )
    return bool(re.fullmatch(r"\d+(?:\.\d+)?", stripped))


def _match_md_and_manifest_tables(
    md_tables: List[Tuple[str, List[List[str]]]],
    manifest_tables: List[Dict[str, Any]],
) -> List[Tuple[int, Tuple[str, List[List[str]]], Dict[str, Any]]]:
    """
    按标题词重合度 + 表头/首列实体重合度为正文表格匹配最相近的底稿 table，
    解决两处建表顺序微调时位置 zip 错配导致的误报或漏报（无特征重合时回退原顺序）。
    """
    n = len(md_tables)
    if n <= 1:
        return [(1, md_tables[0], manifest_tables[0])] if n == 1 else []

    def _bigrams(text: str) -> Set[str]:
        t = _norm_key(text)
        if len(t) < 2:
            return {t} if t else set()
        return {t[i : i + 2] for i in range(len(t) - 1)}

    def _row_keys(rows: List[List[Any]]) -> Set[str]:
        keys: Set[str] = set()
        if not rows:
            return keys
        for h in rows[0]:
            nk = _norm_key(h)
            if nk:
                keys.add(nk)
        for r in rows[1:]:
            if r:
                k0 = _norm_key(r[0])
                if re.match(r"^\d+$", k0) and len(r) > 1:
                    k0 = _norm_key(r[1])
                if k0:
                    keys.add(k0)
        return keys

    # 计算两两相似度并贪心匹配（相似度相同优先保持原始顺序）
    scored_pairs: List[Tuple[float, int, int]] = []
    for mi, (md_title, md_rows) in enumerate(md_tables):
        md_bg = _bigrams(md_title)
        md_keys = _row_keys(md_rows)
        for ji, m_item in enumerate(manifest_tables):
            m_title = f"{m_item.get('label', '')} {m_item.get('title', '')}"
            m_bg = _bigrams(m_title)
            m_keys = _row_keys(m_item.get("raw_data") or [])
            t_score = len(md_bg & m_bg) / max(len(md_bg | m_bg), 1)
            k_score = len(md_keys & m_keys) / max(len(md_keys | m_keys), 1)
            pos_bonus = 0.05 if mi == ji else 0.0
            scored_pairs.append((t_score * 0.45 + k_score * 0.55 + pos_bonus, mi, ji))

    scored_pairs.sort(key=lambda x: (-x[0], x[1], x[2]))
    used_m: Set[int] = set()
    used_j: Set[int] = set()
    matched: Dict[int, int] = {}
    for score, mi, ji in scored_pairs:
        if mi not in used_m and ji not in used_j:
            matched[mi] = ji
            used_m.add(mi)
            used_j.add(ji)

    return [(mi + 1, md_tables[mi], manifest_tables[matched[mi]]) for mi in range(n)]


def audit_md_tables_vs_manifest(
    md_tables: List[Tuple[str, List[List[str]]]],
    manifest_tables: List[Dict[str, Any]],
) -> List[str]:
    """逐表逐行核对正文 Markdown 表格中的关键数字（含小数与纯数值列整数）是否与底稿 type='table' 一致。"""
    errs: List[str] = []
    if len(md_tables) != len(manifest_tables):
        errs.append(
            f"🔴 [正文vs底稿表格数不一致] 正文 TABLE_META 共 {len(md_tables)} 张表，但底稿登记了 {len(manifest_tables)} 张 table！"
        )
        return errs

    for idx, (md_title, md_rows), m_item in _match_md_and_manifest_tables(md_tables, manifest_tables):
        m_rows = m_item.get("raw_data") or []
        if len(md_rows) < 2 or len(m_rows) < 2:
            continue
        # 跳过包含 [待模型填入] 的第四章预测/估值空表（另由占位符门禁拦截）
        if "[待模型填入]" in json.dumps(md_rows, ensure_ascii=False):
            continue

        # 按行首实体名（或排名后第 2 列公司名）映射底稿各行数字
        m_row_nums_by_key: Dict[str, Set[float]] = {}
        for r in m_rows[1:]:
            r_nums: Set[float] = set()
            for cell in r:
                r_nums.update(_extract_cell_nums(cell))
            if r:
                k0 = _norm_key(r[0])
                if re.match(r"^\d+$", k0) and len(r) > 1:
                    k0 = _norm_key(r[1])
                m_row_nums_by_key[k0] = r_nums

        # 核对正文表格每一行的关键数值：
        # 1) 纯数值指标格：拦截所有非年份数字冲突（包括核心整数如营收 120 vs 240、工厂数 8 vs 5）；
        # 2) 长文本备注格：拦截所有小数指标冲突，避免备注句中的描述性小整数引发误报。
        for r_i, mr in enumerate(md_rows[1:], 1):
            if not mr:
                continue
            k0 = _norm_key(mr[0])
            start_col = 1
            if re.match(r"^\d+$", k0) and len(mr) > 1:
                k0 = _norm_key(mr[1])
                start_col = 2

            # 找到底稿中匹配的行
            target_set: Optional[Set[float]] = None
            for mk, mset in m_row_nums_by_key.items():
                if mk and (k0 == mk or k0 in mk or mk in k0):
                    target_set = mset
                    break
            if target_set is None:
                continue

            strict_diffs: List[float] = []
            for c_i, cell in enumerate(mr):
                # 跳过首列主键/排名序号格
                if c_i < start_col:
                    continue
                is_metric = _is_pure_metric_cell(cell)
                for n in _extract_cell_nums(cell):
                    is_year = (abs(n - round(n)) < 1e-4) and (2015 <= n <= 2035)
                    if is_year:
                        continue
                    is_decimal = abs(n - round(n)) > 1e-4
                    if (is_decimal or is_metric) and (n not in target_set):
                        strict_diffs.append(n)

            if strict_diffs:
                errs.append(
                    f"🔴 [正文表格vs底稿数值错位] 表{idx}「{md_title[:24]}」第 {r_i} 行（{mr[0]}）："
                    f"正文数值 {strict_diffs} 未出现在底稿对应行 {sorted(target_set)}！"
                )
    return errs


def audit_prose_numbers(
    md_text: str,
    manifest_data: Dict[str, Any],
    snapshot_text: str,
) -> List[Dict[str, str]]:
    """
    正文纯文字段落数字核查（高精度防误报版）：
    提取正文中带明确财务/业务单位的小数指标（亿元/万元/亿美元/万㎡/%/pct），
    对照【底稿清单全量数字 + 财报快照全量数字 + 正文表格数字】核验；
    凡未落在已核验数字池（含 1 位小数四舍五入与正负绝对值）中的小数，记为 🟡 待人工核对条目，
    并附精确行号、所在章节与原文片段，写入 09_对账与质检报告.md 供研究员直接定位。
    """
    verified_pool: Set[float] = set()

    def _add_nums(text: str) -> None:
        cleaned = str(text or "").replace(",", "")
        for tok in re.findall(r"-?\d+(?:\.\d+)?", cleaned):
            try:
                v = abs(float(tok))
                verified_pool.add(round(v, 2))
                verified_pool.add(round(v, 1))
            except ValueError:
                pass

    _add_nums(json.dumps(manifest_data, ensure_ascii=False))
    _add_nums(snapshot_text)

    # 同步纳入正文 Markdown 表格中的数字（表格本身已与底稿逐行核对）
    for ln in md_text.splitlines():
        if ln.strip().startswith("|"):
            _add_nums(ln)

    unit_pat = re.compile(
        r"([+-]?\d+\.\d+)\s*(亿元|万元|亿美元|百万美元|万亿美元|万㎡|万平方米|%|pct|个百分点)"
    )

    findings: List[Dict[str, str]] = []
    cur_sec = "前言/未分节"
    in_comment = False

    for lineno, raw_ln in enumerate(md_text.splitlines(), 1):
        s = raw_ln.strip()
        if "<!--" in s:
            in_comment = True
        if in_comment:
            if "-->" in s:
                in_comment = False
            continue
        if not s or s.startswith("|") or s.startswith("![") or s.startswith(">"):
            continue
        if s.startswith("#"):
            cur_sec = s.lstrip("#").strip()[:36]
            continue

        # 先去掉数字内部的千分位逗号（如 4,250.48万元 -> 4250.48万元），避免被误拆成 250.48万元
        s_clean = re.sub(r"(?<=\d),(?=\d)", "", s)
        for m in unit_pat.finditer(s_clean):
            val_str, unit = m.group(1), m.group(2)
            v_abs = abs(float(val_str))
            r2 = round(v_abs, 2)
            r1 = round(v_abs, 1)
            # 容差匹配：允许 0.05 以内舍入或恰好命中 1/2 位小数池
            hit = (r2 in verified_pool) or (r1 in verified_pool) or any(
                abs(v_abs - pv) <= 0.05 for pv in verified_pool
            )
            if not hit:
                start = max(0, m.start() - 18)
                end = min(len(s_clean), m.end() + 18)
                snippet = s_clean[start:end].replace("|", "/")
                findings.append({
                    "line": str(lineno),
                    "section": cur_sec,
                    "metric": f"{val_str}{unit}",
                    "snippet": f"…{snippet}…",
                })
    return findings


def audit_outline_shrinkage(
    company_dir: str,
    actual_visual_cnt: int,
    md_text: str,
) -> Tuple[List[str], List[str]]:
    """
    防单回合偷懒缩水门禁（专为云端/单对话框客户端设计）：
    1. 若存在 03_段落级大纲.md，统计大纲拟定图表编号总数；若拟定 ≥ 6 张但成稿（图+表）低于拟定数的 75%，
       直接亮 🔴 拦截（防止出现“大纲拟 17 张图、成稿偷懒只出 4 图 9 页”的严重缩水）；
    2. 统计正文非表格纯文字总字数，若不足 4,500 字打印 🟡 篇幅偏薄提醒。
    """
    reds: List[str] = []
    warns: List[str] = []
    outline_path = os.path.join(company_dir, "03_段落级大纲.md")
    if os.path.exists(outline_path):
        with open(outline_path, "r", encoding="utf-8") as f:
            outline_txt = f.read()
        planned_ids = set(re.findall(r"图表\s*(\d+)", outline_txt))
        planned_cnt = len(planned_ids)
        if planned_cnt >= 6 and actual_visual_cnt < int(planned_cnt * 0.75):
            reds.append(
                f"🔴 [防缩水拦截·图表严重缺口] 大纲（03_段落级大纲.md）拟定 {planned_cnt} 张图表，"
                f"但成稿仅落实 {actual_visual_cnt} 张（不足 75%），严禁偷懒缩水为简报！"
            )

    # 统计非注释、非表格、非标题的纯正文字数
    prose_chars = 0
    in_comment = False
    for ln in md_text.splitlines():
        s = ln.strip()
        if "<!--" in s:
            in_comment = True
        if in_comment:
            if "-->" in s:
                in_comment = False
            continue
        if not s or s.startswith(("#", "|", ">", "![")):
            continue
        prose_chars += len(s)
    if prose_chars < 4500:
        warns.append(
            f"🟡 [篇幅偏薄提醒] 全文纯正文字数仅 {prose_chars} 字（深度报告通常 ≥ 6,000 字），请核查是否遗漏分节落盘。"
        )
    return reds, warns


def audit_qualitative_claims(
    md_text: str,
    manifest_data: Dict[str, Any],
) -> List[Dict[str, str]]:
    """
    正文高敏定性判断（“狠话”）底稿覆盖扫描：
    扫描正文中包含“第一/唯一/龙头/领先/率先/独家/核心供应链/批量供货/核心供应商”等高敏定性词的句子；
    若该句对应核心关键词未在 06_证据清单.json（含 type='text'、图表标题、raw_data）中建档，
    则列为 🟡 待核查项写入 09_对账与质检报告.md（提醒：无公开出处就删改收敛，有出处就补独立底稿）。
    """
    qual_re = re.compile(r"(全球第一|国内第一|行业第一|稳居第一|唯一|独家|率先量产|率先突破|核心供应链|批量供货|核心供应商|国际领先|国内领先)")
    manifest_blob = json.dumps(manifest_data, ensure_ascii=False)

    findings: List[Dict[str, str]] = []
    cur_sec = "前言/未分节"
    in_comment = False
    for lineno, raw_ln in enumerate(md_text.splitlines(), 1):
        s = raw_ln.strip()
        if "<!--" in s:
            in_comment = True
        if in_comment:
            if "-->" in s:
                in_comment = False
            continue
        if not s or s.startswith("|") or s.startswith("![") or s.startswith(">"):
            continue
        if s.startswith("#"):
            cur_sec = s.lstrip("#").strip()[:36]
            continue

        # 按中文句号/分号拆句，精准定位含高敏词的完整句子
        for sent in re.split(r"[。；;]", s):
            sent = sent.strip()
            if not sent:
                continue
            m = qual_re.search(sent)
            if not m:
                continue
            kw = m.group(1)
            # 若该高敏词已在底稿清单中出现过对应证据描述，视为已建档（仍可在人工抽查时复核）
            if kw in manifest_blob:
                continue
            findings.append({
                "line": str(lineno),
                "section": cur_sec,
                "keyword": kw,
                "sentence": sent[:90].replace("|", "/"),
            })
    return findings


def audit_prospectus_concentration(manifest_data: Dict[str, Any]) -> List[str]:
    """
    行业与竞争章节（第二、三章）防“照搬招股书”来源多元化自动检测：
    - 扫描 06_证据清单.json 中 sheet_name 以 '2.' 或 '3.' 开头的 chart / table 条目；
    - 判定某条目是否仅依赖目标公司自身招股书（含招股章程/上市申请资料，且未叠加独立外部机构/年报/中报/数据库/竞品财报）；
    - 当第二、三章有效条目 ≥ 4 且招股书单一来源占比 > 50% 时，输出 🟡 告警，
      提示按 脚本与规范/industry_sources.md 补充独立三方库、官方宏观/海关数据或可比竞品财报。
    """
    warns: List[str] = []
    prospectus_kw = ("招股书", "招股書", "招股章程", "上市申请资料", "上市申請資料", "聆讯后资料集", "聆訊後資料集")
    # 独立非招股书的外部/定期来源标识（若与招股书并列出现，说明已有跨文档印证，不计为纯招股书单一来源）
    independent_kw = (
        "年报", "年報", "中期报告", "中期報告", "季报", "季報", "ifind", "wind", "choice",
        "euromonitor", "欧睿", "歐睿", "nielsen", "尼尔森", "kantar", "凯度",
        "world bank", "世界银行", "imf", "un ", "联合国", "海关", "海關", "统计局", "統計局",
        "semi", "trendforce", "idc", "gartner", "omdia", "yole", "prismark", "counterpoint",
        "百川", "卓创", "隆众", "usgs", "中汽协", "乘联会", "ggii", "bnef", "iea", "iqvia",
        "10-k", "10-q", "20-f", "reuters", "bloomberg", "财联社", "证券时报", "第一财经",
    )

    total_cnt = 0
    prospectus_only_cnt = 0
    for sec in manifest_data.get("sections", []):
        sname = str(sec.get("sheet_name", "")).strip()
        if not (sname.startswith("2.") or sname.startswith("3.")):
            continue
        for item in sec.get("items", []):
            if item.get("type") not in ("chart", "table"):
                continue
            total_cnt += 1
            src = str(item.get("source_name", ""))
            src_low = src.lower()
            has_pros = any(k in src for k in prospectus_kw)
            if not has_pros:
                continue
            # 剔除“招股书（援引XXX）”括号内附带的招股书转引字样，只判断招股书之外是否还有独立来源文档
            outside_paren = re.sub(r"[（(][^）)]*[）)]", "", src).lower()
            # 检查是否存在除招股书外且未标数据库豁免的独立文档来源
            has_indep = bool(item.get("is_database_exempt")) or any(
                k in outside_paren for k in independent_kw
            )
            if not has_indep:
                prospectus_only_cnt += 1

    if total_cnt >= 4:
        ratio = prospectus_only_cnt / total_cnt
        if ratio > 0.50:
            warns.append(
                f"🟡 [来源单一预警·过度依赖招股书] 第二、三章共 {total_cnt} 张图表/表格中，"
                f"{prospectus_only_cnt} 张（{ratio * 100:.1f}%，超 50% 红线）仅源自目标公司招股书；"
                f"建议按 脚本与规范/industry_sources.md 补充独立三方数据库（如消费查欧睿、半导体查SEMI）、"
                f"官方宏观/海关数据及可比竞品一手财报以提升增量深度。"
            )
    return warns


def write_qc_report(
    company_dir: str,
    reds: List[str],
    yellows: List[str],
    prose_findings: List[Dict[str, str]],
    qual_findings: List[Dict[str, str]],
    summary_line: str,
    release: bool,
) -> str:
    """
    每次对账自动落盘《09_对账与质检报告.md》，把所有 🔴 阻断项、🟡 标黄提示
    以及正文段落未锚定小数（含精确行号与原句片段）集中列成表格，研究员点开即查，无需翻终端日志。
    """
    out_path = os.path.join(company_dir, "09_对账与质检报告.md")
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mode = "正式发布门禁 (--release)" if release else "草稿体检模式 (--strict)"

    lines = [
        "# 深度报告三端对账与质检报告",
        "",
        f"> 生成时间：{now}｜检查模式：{mode}",
        f"> 汇总：`{summary_line}`",
        "",
        "## 一、🔴 必须修复项（阻断发版）",
        "",
    ]
    if not reds:
        lines.append("- 无 🔴 阻断项。")
    else:
        for r in reds:
            lines.append(f"- {r}")

    lines.extend(["", "## 二、🟡 标黄提醒项（草稿允许，发版前需逐项确认）", ""])
    if not yellows:
        lines.append("- 无 🟡 标黄项。")
    else:
        for y in yellows:
            lines.append(f"- {y}")

    lines.extend([
        "",
        "## 三、🟡 正文纯文字段落未在图表/底稿/财报快照中锚定的小数指标（供快速复核）",
        "",
        "> 说明：以下为正文中带明确财务/业务单位的小数，但未直接出现在 `06_证据清单.json` 或 `01_最新财报全科目快照.md` 中（可能是合理文字衍生计算/背景常识，也可能是改图忘改文导致的数字漂移），请按行号在 `05_完整初稿.md` 中快速抽核。",
        "",
    ])
    if not prose_findings:
        lines.append("- 全部正文小数指标均已在底稿/财报快照中锚定（0 处异常）。")
    else:
        lines.extend([
            "| 行号 | 所在章节 | 指标数值 | 原文片段 |",
            "|---:|---|---|---|",
        ])
        for pf in prose_findings:
            lines.append(f"| L{pf['line']} | {pf['section']} | `{pf['metric']}` | {pf['snippet']} |")

    lines.extend([
        "",
        "## 四、🟡 正文高敏定性判断（“狠话”）底稿核查清单",
        "",
        "> 原则：**无公开出处的定性狠话直接删改收敛；确有公开一手出处的硬判断，必须补建 `type: \"text\"` 底稿条目并整句画红框。**",
        "",
    ])
    if not qual_findings:
        lines.append("- 未发现未入底稿的高敏定性狠话（0 处）。")
    else:
        lines.extend([
            "| 行号 | 所在章节 | 命中高敏词 | 原文完整句子 |",
            "|---:|---|---|---|",
        ])
        for qf in qual_findings:
            lines.append(f"| L{qf['line']} | {qf['section']} | `{qf['keyword']}` | {qf['sentence']} |")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return out_path


def check_redteam_review_binding(company_dir: str, current_sha: str) -> Tuple[Optional[str], Optional[str]]:
    """
    核查红队审稿报告（兼容 05_红队审稿意见.md 与 05_红队审稿意见_Claude.md，支持 Opus/Sonnet/ChatGPT/Gemini）
    是否存在、是否绑定到当前最新稿件哈希、以及放行结论。
    返回 (release_blocker_msg, info_or_warn_msg)。
    """
    candidates = [
        os.path.join(company_dir, "05_红队审稿意见.md"),
        os.path.join(company_dir, "05_红队审稿意见_Claude.md"),
    ]
    existing = [p for p in candidates if os.path.exists(p)]
    if not existing:
        return (
            "缺少红队审稿报告（05_红队审稿意见.md 或 05_红队审稿意见_Claude.md）",
            "🟡 [未执行Step 3.5红队审稿] 未找到 05_红队审稿意见.md / 05_红队审稿意见_Claude.md",
        )
    # 若两者并存，优先取最新修改的一份
    rev_path = max(existing, key=os.path.getmtime)
    rev_name = os.path.basename(rev_path)

    with open(rev_path, "r", encoding="utf-8") as f:
        txt = f.read()

    sha_m = re.search(r"draft_sha256\s*:\s*([a-fA-F0-9]+)", txt)
    ver_m = re.search(r"verdict\s*:\s*([A-Z_]+)", txt)
    mod_m = re.search(r"reviewer_model\s*:\s*([^\r\n]+)", txt)
    model_str = mod_m.group(1).strip() if mod_m else "未注明模型"

    if not sha_m or not ver_m:
        return (
            f"审稿报告 {rev_name} 未包含机器可读版本头（当前稿件指纹={current_sha}），无法证明最终稿已通过复审",
            f"🟡 [审稿版本未绑定] {rev_name} 缺少 REVIEW_META（draft_sha256={current_sha}），正式发版前须复审签认",
        )

    rev_sha = sha_m.group(1).strip()
    verdict = ver_m.group(1).strip()
    if rev_sha != current_sha:
        return (
            f"审稿绑定指纹({rev_sha})与当前最新稿件指纹({current_sha})不一致，修订后尚未复审",
            f"🟡 [审稿已过期] 稿件已修订（审稿指纹={rev_sha} vs 当前={current_sha}），正式出稿前须复审",
        )
    if verdict not in ("PASS", "CONDITIONAL_PASS"):
        return (
            f"审稿结论为 {verdict}，未获放行",
            f"🟡 [审稿未放行] 当前版本审稿结论为 {verdict}",
        )
    return None, f"[INFO] 红队审稿版本绑定一致（{rev_name}｜模型={model_str}｜draft_sha256={current_sha}｜verdict={verdict}）"


def run_sync_check(
    company_dir: str,
    clean_orphans: bool = False,
    release: bool = False,
) -> Tuple[List[str], List[str]]:
    """执行六维三端有序对账与发布门禁检查，返回 (🔴阻断列表, 🟡提醒列表)。"""
    md_path = os.path.join(company_dir, "05_完整初稿.md")
    charts_path = os.path.join(company_dir, "charts.json")
    if not os.path.exists(charts_path):
        alt_charts = os.path.join(company_dir, "07_图表", "charts.json")
        if os.path.exists(alt_charts):
            charts_path = alt_charts
    manifest_path = os.path.join(company_dir, "06_证据清单.json")
    meta_path = os.path.join(company_dir, "report_meta.json")
    chart_dir = os.path.join(company_dir, "07_图表")

    reds: List[str] = []
    yellows: List[str] = []

    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()
    with open(charts_path, "r", encoding="utf-8") as f:
        charts_data = json.load(f)
    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest_data = json.load(f)
    meta_text = ""
    if os.path.exists(meta_path):
        with open(meta_path, "r", encoding="utf-8") as f:
            meta_text = f.read()

    # 1. 正文引用图片 vs charts.json vs 磁盘文件（兼容 list 或 {"charts": [...]} 两种格式）
    md_imgs, md_tables = parse_md_charts_and_tables(md_text)
    md_img_set = set(md_imgs)
    charts_list = charts_data if isinstance(charts_data, list) else charts_data.get("charts", [])
    cfg_by_file: Dict[str, Dict[str, Any]] = {
        os.path.basename(c["filename"]): c for c in charts_list
    }
    cfg_img_set = set(cfg_by_file.keys())

    missing_in_cfg = md_img_set - cfg_img_set
    if missing_in_cfg:
        reds.append(f"🔴 [正文vs画图错位] 正文引用但 charts.json 中缺失的图表：{sorted(missing_in_cfg)}")

    unused_in_cfg = cfg_img_set - md_img_set
    if unused_in_cfg:
        yellows.append(f"🟡 [画图配置冗余] charts.json 定义了但正文未引用的图表：{sorted(unused_in_cfg)}")

    for fn in sorted(md_img_set):
        if not os.path.exists(os.path.join(chart_dir, fn)):
            reds.append(f"🔴 [图片文件缺失] 07_图表/{fn} 不存在，请先运行 chart_maker.py")

    # 2. 孤儿旧图检测与清理
    if os.path.isdir(chart_dir):
        disk_pngs = {f for f in os.listdir(chart_dir) if f.lower().endswith(".png")}
        orphans = sorted(disk_pngs - md_img_set)
        if orphans:
            if clean_orphans:
                for ofn in orphans:
                    os.remove(os.path.join(chart_dir, ofn))
                print(f"[CLEAN] 已自动清理 07_图表/ 下 {len(orphans)} 张废弃孤儿旧图")
            else:
                yellows.append(
                    f"🟡 [孤儿旧图残留] 07_图表/ 存在 {len(orphans)} 张未被正文引用的旧图（可加 --clean-orphans 清理）"
                )

    # 3. 收集底稿图表与表格
    manifest_charts: Dict[str, Dict[str, Any]] = {}
    manifest_tables: List[Dict[str, Any]] = []
    for sec in manifest_data.get("sections", []):
        for item in sec.get("items", []):
            if item.get("type") == "chart" and item.get("chart_image"):
                manifest_charts[os.path.basename(item["chart_image"])] = item
            elif item.get("type") == "table":
                manifest_tables.append(item)

    missing_in_manifest = md_img_set - set(manifest_charts.keys())
    if missing_in_manifest:
        reds.append(f"🔴 [正文vs底稿缺图] 正文引用但 06_证据清单.json 漏登的图表：{sorted(missing_in_manifest)}")

    # 4. 【期间 × 数值】有序对账（charts.json ↔ 底稿 raw_data）+ 图表标题年份与横轴一致性校验
    chart_drift_errs: List[str] = []
    for fn, cfg in cfg_by_file.items():
        if fn in manifest_charts:
            chart_drift_errs.extend(audit_chart_ordered_values(fn, cfg, manifest_charts[fn]))
    title_year_errs = audit_chart_title_year_range(parse_md_chart_titles(md_text), cfg_by_file)
    reds.extend(chart_drift_errs)
    reds.extend(title_year_errs)

    # 4.5 永久写作红线轻量扫描（不分析财务费用率 + 防止正文连续堆砌数字）
    yellows.extend(audit_style_guardrails(md_text))

    # 5. 正文 Markdown 表格 ↔ 底稿 table 逐行数值对账
    table_drift_errs = audit_md_tables_vs_manifest(md_tables, manifest_tables)
    reds.extend(table_drift_errs)

    # 5.5 正文纯文字段落小数指标锚定检查（对照底稿清单 + 财报快照 + 表格）
    snapshot_path = os.path.join(company_dir, "00_原始资料", "01_最新财报全科目快照.md")
    snapshot_text = ""
    if os.path.exists(snapshot_path):
        with open(snapshot_path, "r", encoding="utf-8") as f:
            snapshot_text = f.read()
    prose_findings = audit_prose_numbers(md_text, manifest_data, snapshot_text)
    if prose_findings:
        yellows.append(
            f"🟡 [正文段落未锚定小数] 发现 {len(prose_findings)} 处正文带单位小数未出现在底稿/财报快照中（已带行号写入 09_对账与质检报告.md 供抽核）"
        )

    qual_findings = audit_qualitative_claims(md_text, manifest_data)
    if qual_findings:
        yellows.append(
            f"🟡 [定性狠话未建底稿] 发现 {len(qual_findings)} 处高敏定性表述未在底稿清单命中（无出处请删改收敛，有出处请补 type='text' 底稿，详见 09_对账与质检报告.md）"
        )

    # 5.6 防单回合偷懒缩水门禁（大纲拟定图表数 vs 成稿落地图表数 + 篇幅下限检查）
    shrink_reds, shrink_warns = audit_outline_shrinkage(company_dir, len(md_imgs) + len(md_tables), md_text)
    reds.extend(shrink_reds)
    yellows.extend(shrink_warns)

    # 5.7 行业与竞争章节防“照搬招股书”来源多元化检测（第二、三章招股书单一来源占比 ≤ 50%）
    yellows.extend(audit_prospectus_concentration(manifest_data))

    # 6. 首页摘要字数、模型占位符、红队审稿版本哈希核查
    if meta_text:
        meta = json.loads(meta_text)
        sum_chars = sum(
            len(p.replace("**", ""))
            for sec in meta.get("summary_sections", [])
            for p in sec.get("paragraphs", [])
        )
        print(f"[INFO] 首页摘要总字数：{sum_chars} 字（建议上限 ≤ 850 字）")
        if sum_chars > 900:
            reds.append(f"🔴 [首页摘要超长风险] 摘要共 {sum_chars} 字（>900 字），极易撑出首页。")
        elif sum_chars > 850:
            yellows.append(f"🟡 [首页摘要偏长提示] 摘要共 {sum_chars} 字（建议 ≤ 850 字）。")

    ph_cnt = len(PLACEHOLDER_RE.findall(md_text)) + len(PLACEHOLDER_RE.findall(meta_text))
    if ph_cnt > 0:
        msg = f"全文与摘要尚有 {ph_cnt} 处 [待模型填入] 占位符未填写"
        if release:
            reds.append(f"🔴 [正式发布阻断·模型未填] {msg}")
        else:
            yellows.append(f"🟡 [AI交付态·模型占位] {msg}（待研究员在 Word 中自行刷模型填入；若回填进 Markdown 可选跑 --release）")

    cur_sha = compute_draft_sha256(md_text, meta_text)
    rev_blocker, rev_msg = check_redteam_review_binding(company_dir, cur_sha)
    if rev_blocker:
        if release:
            reds.append(f"🔴 [正式发布阻断·审稿未绑定] {rev_blocker}")
        elif rev_msg:
            yellows.append(rev_msg)
    elif rev_msg:
        print(rev_msg)

    summary_line = (
        f"正文/画图/底稿图表 {len(md_imgs)}/{len(cfg_img_set)}/{len(manifest_charts)} 张 | "
        f"正文/底稿表格 {len(md_tables)}/{len(manifest_tables)} 张 | "
        f"期间数值错配 {len(chart_drift_errs)} 处 | 表格数值错位 {len(table_drift_errs)} 处 | "
        f"正文未锚定小数 {len(prose_findings)} 处 | 占位符 {ph_cnt} 处"
    )
    print(f"[SUMMARY] {summary_line}")
    qc_path = write_qc_report(company_dir, reds, yellows, prose_findings, qual_findings, summary_line, release)
    print(f"[REPORT] 已生成可视化质检清单 -> {qc_path}")
    return reds, yellows


def main() -> None:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="国金深度报告三端有序对账、审稿版本绑定与发布门禁脚本")
    parser.add_argument("--dir", required=True, help="公司报告目录（如 示例公司）")
    parser.add_argument("--clean-orphans", action="store_true", help="自动删除 07_图表/ 中未被正文引用的旧 PNG")
    parser.add_argument("--strict", action="store_true", help="草稿严格模式：存在任一 🔴 错位即返回退出码 1")
    parser.add_argument("--release", action="store_true", help="正式发布门禁：同时拦截未填模型占位与未绑定终稿版本的红队审稿")
    args = parser.parse_args()

    reds, yellows = run_sync_check(args.dir, clean_orphans=args.clean_orphans, release=args.release)
    for r in reds:
        print(" ", r)
    for y in yellows:
        print(" ", y)

    if not reds:
        if args.release:
            print("[RELEASE-PASS] 三端有序对账、审稿版本绑定与占位符全部 100% 通过正式发布门禁！")
        else:
            print("[DRAFT-PASS] 三端【期间×数值】与正文表格对账 0 错位（0 🔴）！")
    elif args.strict or args.release:
        sys.exit(1)


if __name__ == "__main__":
    main()
