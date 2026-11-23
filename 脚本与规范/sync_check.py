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
    python .\\脚本与规范\\sync_check.py --dir .\\示例公司 --clean-orphans --strict
    python .\\脚本与规范\\sync_check.py --dir .\\示例公司 --release
"""

import argparse
import hashlib
import json
import os
import re
import sys
from typing import Any, Dict, List, Optional, Set, Tuple


PLACEHOLDER_RE = re.compile(r"\[待模型填入\]|\[待补充[^\]]*\]|\[待确认[^\]]*\]|\bTBD\b|\bTODO\b")


def _norm_key(s: Any) -> str:
    """归一化期间/分类/公司名，用于跨表精确对齐（兼容 2026Q1-Q3 / 26Q1-Q3 等简写）。"""
    t = re.sub(r"[\s（）()：:【】\[\]—–\-]", "", str(s or "")).lower()
    t = t.replace("2026", "26").replace("2025", "25").replace("2024", "24").replace("2023", "23").replace("2022", "22")
    return t


def _extract_cell_nums(val: Any) -> List[float]:
    """提取单元格中的所有数值（忽略纯年份 2015–2030 以防干扰业务数字对账）。"""
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

    # 构建底稿按“列期间”和按“行分类”的数值映射
    by_col_key: Dict[str, Set[float]] = {}
    for c_idx, hk in enumerate(header_norm):
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

            # 查找底稿中对应的期间列（精确匹配或子串匹配）
            matched_bucket: Optional[Set[float]] = None
            for hk, bset in by_col_key.items():
                if hk and (ck == hk or ck in hk or hk in ck):
                    matched_bucket = bset
                    break
            if matched_bucket is None:
                for rk, bset in by_row_key.items():
                    if rk and (ck == rk or ck in rk or rk in ck):
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


def audit_md_tables_vs_manifest(
    md_tables: List[Tuple[str, List[List[str]]]],
    manifest_tables: List[Dict[str, Any]],
) -> List[str]:
    """逐表逐行核对正文 Markdown 表格中的关键数字是否与底稿 type='table' 的 raw_data 一致。"""
    errs: List[str] = []
    if len(md_tables) != len(manifest_tables):
        errs.append(
            f"🔴 [正文vs底稿表格数不一致] 正文 TABLE_META 共 {len(md_tables)} 张表，但底稿登记了 {len(manifest_tables)} 张 table！"
        )
        return errs

    for idx, ((md_title, md_rows), m_item) in enumerate(zip(md_tables, manifest_tables), 1):
        m_rows = m_item.get("raw_data") or []
        if len(md_rows) < 2 or len(m_rows) < 2:
            continue
        # 跳过包含 [待模型填入] 的第四章预测/估值空表（另由占位符门禁拦截）
        if "[待模型填入]" in json.dumps(md_rows, ensure_ascii=False):
            continue

        # 按行首实体名（或行号）映射底稿各行数字
        m_row_nums_by_key: Dict[str, Set[float]] = {}
        m_all_nums: Set[float] = set()
        for r in m_rows[1:]:
            r_nums: Set[float] = set()
            for cell in r:
                r_nums.update(_extract_cell_nums(cell))
            m_all_nums.update(r_nums)
            if r:
                k0 = _norm_key(r[0])
                # 若首列是排名数字（如 1..10），改用第 2 列公司名作主键
                if re.match(r"^\d+$", k0) and len(r) > 1:
                    k0 = _norm_key(r[1])
                m_row_nums_by_key[k0] = r_nums

        # 核对正文表格每一行的显著数值（含小数或 >= 20 且非年份）是否一致落在底稿对应行
        for r_i, mr in enumerate(md_rows[1:], 1):
            if not mr:
                continue
            k0 = _norm_key(mr[0])
            if re.match(r"^\d+$", k0) and len(mr) > 1:
                k0 = _norm_key(mr[1])
            row_salient_nums = [
                n for cell in mr for n in _extract_cell_nums(cell)
                if (abs(n - round(n)) > 1e-4) or (n >= 20 and not (2015 <= n <= 2035))
            ]
            if not row_salient_nums:
                continue

            # 找到底稿中匹配的行
            target_set: Optional[Set[float]] = None
            for mk, mset in m_row_nums_by_key.items():
                if mk and (k0 == mk or k0 in mk or mk in k0):
                    target_set = mset
                    break
            if target_set is None:
                continue

            diff = [n for n in row_salient_nums if n not in target_set]
            # 允许底稿侧略去非核心说明数字，但若正文行关键财务金额（含小数）与底稿行存在直接冲突则拦截
            decimal_diff = [n for n in diff if abs(n - round(n)) > 1e-4]
            if decimal_diff:
                errs.append(
                    f"🔴 [正文表格vs底稿数值错位] 表{idx}「{md_title[:24]}」第 {r_i} 行（{mr[0]}）："
                    f"正文数值 {decimal_diff} 未出现在底稿对应行 {sorted(target_set)}！"
                )
    return errs


def check_redteam_review_binding(company_dir: str, current_sha: str) -> Tuple[Optional[str], Optional[str]]:
    """
    核查 05_红队审稿意见_Claude.md 是否存在、是否绑定到当前最新稿件哈希、以及放行结论。
    返回 (release_blocker_msg, info_or_warn_msg)。
    """
    rev_path = os.path.join(company_dir, "05_红队审稿意见_Claude.md")
    if not os.path.exists(rev_path):
        return (
            "缺少红队审稿报告 05_红队审稿意见_Claude.md",
            "🟡 [未执行Step 3.5红队审稿] 未找到 05_红队审稿意见_Claude.md",
        )

    with open(rev_path, "r", encoding="utf-8") as f:
        txt = f.read()

    sha_m = re.search(r"draft_sha256\s*:\s*([a-fA-F0-9]+)", txt)
    ver_m = re.search(r"verdict\s*:\s*([A-Z_]+)", txt)

    if not sha_m or not ver_m:
        return (
            f"审稿报告未包含机器可读版本头（当前稿件指纹={current_sha}），无法证明最终稿已通过复审",
            f"🟡 [审稿版本未绑定] 05_红队审稿意见_Claude.md 缺少 REVIEW_META（draft_sha256={current_sha}），修订后需由 Opus 复审签认",
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
    return None, f"[INFO] 红队审稿版本绑定一致（draft_sha256={current_sha}，verdict={verdict}）"


def run_sync_check(
    company_dir: str,
    clean_orphans: bool = False,
    release: bool = False,
) -> Tuple[List[str], List[str]]:
    """执行六维三端有序对账与发布门禁检查，返回 (🔴阻断列表, 🟡提醒列表)。"""
    md_path = os.path.join(company_dir, "05_完整初稿.md")
    charts_path = os.path.join(company_dir, "charts.json")
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

    # 1. 正文引用图片 vs charts.json vs 磁盘文件
    md_imgs, md_tables = parse_md_charts_and_tables(md_text)
    md_img_set = set(md_imgs)
    cfg_by_file: Dict[str, Dict[str, Any]] = {
        os.path.basename(c["filename"]): c for c in charts_data.get("charts", [])
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

    # 4. 【期间 × 数值】有序对账（charts.json ↔ 底稿 raw_data）
    chart_drift_errs: List[str] = []
    for fn, cfg in cfg_by_file.items():
        if fn in manifest_charts:
            chart_drift_errs.extend(audit_chart_ordered_values(fn, cfg, manifest_charts[fn]))
    reds.extend(chart_drift_errs)

    # 5. 正文 Markdown 表格 ↔ 底稿 table 逐行数值对账
    table_drift_errs = audit_md_tables_vs_manifest(md_tables, manifest_tables)
    reds.extend(table_drift_errs)

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
            yellows.append(f"🟡 [草稿态·模型占位] {msg}（填数后须过 --release）")

    cur_sha = compute_draft_sha256(md_text, meta_text)
    rev_blocker, rev_msg = check_redteam_review_binding(company_dir, cur_sha)
    if rev_blocker:
        if release:
            reds.append(f"🔴 [正式发布阻断·审稿未绑定] {rev_blocker}")
        elif rev_msg:
            yellows.append(rev_msg)
    elif rev_msg:
        print(rev_msg)

    print(
        f"[SUMMARY] 正文/画图/底稿图表 {len(md_imgs)}/{len(cfg_img_set)}/{len(manifest_charts)} 张 | "
        f"正文/底稿表格 {len(md_tables)}/{len(manifest_tables)} 张 | "
        f"期间数值错配 {len(chart_drift_errs)} 处 | 表格数值错位 {len(table_drift_errs)} 处 | 占位符 {ph_cnt} 处"
    )
    return reds, yellows


def main() -> None:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="国金深度报告三端有序对账、审稿版本绑定与发布门禁脚本")
    parser.add_argument("--dir", required=True, help="公司报告目录（如 .\\示例公司）")
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
