# -*- coding: utf-8 -*-
"""
国金证券研报合规工作底稿 Excel 自动生成与六道质控门禁审计脚本 (build_workpaper.py)

核心功能：
1. 基于官方模板（国金研报工作底稿模板.xlsx），正式输出时自动移除《发布底稿要求》《数字未来实验室-ID》
   以及模板示例页，仅保留《主要数据》+ 报告各次级章节 Sheet（1.1、1.2、2.1、2.2…）。
2. 严格遵循左右分栏结构：
   - 左侧（A–G 列）：报告图表编号 + 报告图 / 基础数据表（自动换行与动态行高）/ 正文关键文字摘录
   - 右侧（I–P 列）：资料来源名称 + 可点击精确网址/页码 + 单图或多图纵向堆叠的高清红框截图（无截图标黄）
3. 内置【六道程序化质控门禁】（彻底杜绝“多源只贴单图”“缺漏客户/同行佐证”“张冠李戴”“空泛网址”）：
   - 门禁 1（黑名单来源拦截）：扫描来源名/网址/截图标题，命中自媒体/论坛/调研纪要/空泛来源直接红灯告警；
   - 门禁 2（缺失/失效截图拦截）：非数据库条目若无截图或截图文件路径不存在，红灯告警并写入待补清单；
   - 门禁 3（多源/多公司数量门禁）：自动识别“多公司对比表/图”（如 6 家同行泰国基地对比）、“多文档组合来源”
     （如招股书+年报+调研表）与“复合多赛道矩阵图”（如客户矩阵），强制核查截图张数是否与覆盖实体/文档数匹配，
     杜绝复合图表只贴 1 张局部截图蒙混过关；
   - 门禁 4（原子事实 required_claims + Sidecar 文本覆盖率门禁）：自动读取每张截图旁的 `<图片名>.meta.json`
     （由 auto_evidence_capture.py 落盘）以及清单中的 covered_claims，逐条核验 required_claims 中的客户名、
     排名、百分比与关键数字是否 100% 被已贴截图并集覆盖，有缺口精确报出缺失词表；
   - 门禁 5（精确网址/页码门禁）：拦截仅填官网根域名（如 https://www.xxx.com）的粗放网址；
   - 门禁 6（质控体检总表 + --strict 严格失败退出）：每次生成必打印完整覆盖率体检卡，传入 `--strict` 时
     只要存在任一 🔴 严重缺口即以非零退出码阻断流程。
4. 高清图片嵌入（已取消 20 MB 体积限制）：
   - 默认以宽上限 1400px、JPEG 质量 90 高清嵌入，确保红框、小字与页码清晰可辨。

命令行用法：
    D:\\anaconda3\\python.exe build_workpaper.py --manifest 06_证据清单.json --out 公司名_工作底稿.xlsx
    D:\\anaconda3\\python.exe build_workpaper.py --manifest 06_证据清单.json --out 公司名_工作底稿.xlsx --strict
    D:\\anaconda3\\python.exe build_workpaper.py --manifest 06_证据清单.json --audit-only
"""

import argparse
import io
import json
import os
import re
import sys
from typing import Any, Dict, List, Optional, Set, Tuple
from urllib.parse import urlparse

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image as PILImage, ImageDraw


# 合规黑名单来源关键词（命中即打印红色告警）
BLACKLIST_SOURCES = [
    "调研纪要",
    "内部交流",
    "私下交流",
    "百度",
    "新浪",
    "sina.cn",
    "sina.com",
    "搜狐",
    "微博",
    "知乎",
    "股吧",
    "雪球",
    "hao.cnyes.com",
    "钜亨号",
    "peoplenews.tw",
    "民报财经",
    "998pcb.com",
    "Morgan Stanley",
    "摩根士丹利",
    "高盛",
    "中信证券",
    "中金公司",
    "华泰证券",
    "招商证券",
    "广发证券",
    "中信建投",
    "公开资料整理",
    "公开行业资料",
    "公开行业研究资料",
    "行业公开资料",
    "产业链公开拆解",
    "公开行业资讯",
]

# 数据库来源关键词（列出导出基础数据+来源名称即可，无需强制截图）
DATABASE_KEYWORDS = ["ifind", "wind", "万得", "同花顺", "choice", "bloomberg"]

# 单一权威榜单关键词（若为单份统一排名表，如 Prismark 全球前十大榜单，1 张榜单原图即可覆盖全部行）
SINGLE_RANKING_KEYWORDS = ["百强榜", "前十大", "排名", "tracker", "榜单", "top10", "top 10"]

# 不同类型公司官方披露文档关键词（用于自动识别“多文档组合来源”）
DOC_TYPE_PATTERNS = [
    ("招股书", ["招股说明书", "上市申请资料", "招股书"]),
    ("年报", ["年报", "年度报告"]),
    ("半年报", ["半年报", "半年度报告", "中报"]),
    ("季报", ["季报", "季度报告"]),
    ("调研表", ["投资者关系活动记录表", "调研公告", "调研记录"]),
    ("临时公告", ["号公告", "公开公告", "公告"]),
]

# 通用单元格样式（与国金深蓝+浅灰风格保持一致）
FILL_NAVY = PatternFill(start_color="004870", end_color="004870", fill_type="solid")
FILL_SUBHEADER = PatternFill(start_color="D9E6F2", end_color="D9E6F2", fill_type="solid")
FILL_WARN = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_TABLE_HEAD = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")

FONT_HEADER = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
FONT_BANNER = Font(name="微软雅黑", size=10, bold=True, color="004870")
FONT_BOLD = Font(name="微软雅黑", size=9.5, bold=True, color="262626")
FONT_NORMAL = Font(name="微软雅黑", size=9.5, color="262626")
FONT_LINK = Font(name="微软雅黑", size=9.5, color="0563C1", underline="single")

THIN_SIDE = Side(style="thin", color="BFBFBF")
BORDER_THIN = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)


def _norm(text: str) -> str:
    """去空白归一化，统一用于文本与原子事实匹配。"""
    return "".join((text or "").split()).lower()


def compress_image_to_bytes(
    img_path: str,
    max_width: int = 1400,
    quality: int = 90,
    highlight_box: Optional[List[float]] = None,
) -> Tuple[io.BytesIO, int, int]:
    """
    将本地 PNG/JPG 图片等比缩放并编码为高清 JPEG 内存流，返回 (BytesIO, 显示宽px, 显示高px)。
    若传入 highlight_box=[x1, y1, x2, y2]（0–1 相对比例），自动在关键文字/排名/市占率处画红框。
    """
    with PILImage.open(img_path) as im:
        if im.mode in ("RGBA", "LA", "P"):
            bg = PILImage.new("RGB", im.size, (255, 255, 255))
            if im.mode == "P":
                im = im.convert("RGBA")
            bg.paste(im, mask=im.split()[-1] if im.mode in ("RGBA", "LA") else None)
            im = bg
        else:
            im = im.convert("RGB")

        w, h = im.size
        if w > max_width:
            ratio = max_width / float(w)
            w, h = max_width, max(1, int(h * ratio))
            im = im.resize((w, h), PILImage.LANCZOS)

        if highlight_box and len(highlight_box) == 4:
            x1, y1, x2, y2 = highlight_box
            draw = ImageDraw.Draw(im)
            draw.rectangle(
                [int(x1 * w), int(y1 * h), int(x2 * w), int(y2 * h)],
                outline="#D92D20",
                width=4,
            )

        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=quality, optimize=True)
        buf.seek(0)
        return buf, w, h


def resolve_path(path_str: str, base_dir: str) -> str:
    """将清单中的相对路径转换为绝对路径。"""
    if not path_str:
        return ""
    if os.path.isabs(path_str):
        return path_str
    return os.path.normpath(os.path.join(base_dir, path_str))


def is_database_source(source_name: str) -> bool:
    """判断是否为合规数据库来源（iFinD / Wind 等）。"""
    low = (source_name or "").lower()
    return any(k in low for k in DATABASE_KEYWORDS)


def extract_shot_entries(item: Dict[str, Any]) -> List[Dict[str, Any]]:
    """统一提取单图（evidence_screenshot）或多图（evidence_screenshots）条目列表。"""
    raw_list = item.get("evidence_screenshots") or []
    if not raw_list and item.get("evidence_screenshot"):
        raw_list = [
            {
                "path": item.get("evidence_screenshot", ""),
                "caption": "",
                "highlight_box": item.get("highlight_box"),
            }
        ]
    normalized: List[Dict[str, Any]] = []
    for se in raw_list:
        if isinstance(se, str):
            normalized.append({"path": se, "caption": ""})
        elif isinstance(se, dict):
            normalized.append(se)
    return normalized


def _is_bare_homepage_url(url: str) -> bool:
    """判断一个 HTTP(S) 链接是否只是官网首页根域名（缺少具体文章/公告路径）。"""
    u = (url or "").strip()
    if not u.startswith(("http://", "https://")):
        return False
    # 若包含多条用分号隔开的 URL，只要第一条是精确路径即放行
    first_u = u.split(";")[0].split()[0].strip()
    parsed = urlparse(first_u)
    path = (parsed.path or "").strip("/")
    return len(path) == 0 and not parsed.query


def _detect_peer_company_count(raw_data: List[List[Any]], source_name: str) -> int:
    """
    自动识别“多公司对比表/对比图”中涉及的独立上市公司数量。
    若属于单份权威机构统一排名表（如 Prismark 全球前十大榜单），返回 0（1 张原榜图即可覆盖）。
    """
    if not raw_data or len(raw_data) < 2:
        return 0
    low_src = (source_name or "").lower()
    if any(k in low_src for k in SINGLE_RANKING_KEYWORDS):
        return 0

    header = [str(x or "").strip() for x in raw_data[0]]
    h0 = header[0] if header else ""
    # 情况 A：纵向对比表（首列为公司名称，且至少有 3 家公司行）
    if any(k in h0 for k in ("公司", "厂商", "企业", "标的")) and len(raw_data) >= 4:
        # 排除“每列一家公司”的横向对比图
        if len(raw_data) >= 4:
            companies = [str(row[0] or "").strip() for row in raw_data[1:] if row and str(row[0] or "").strip()]
            # 过滤纯指标名行（如“泰国基地公开投资额”）
            if len(companies) >= 3 and not any("投资额" in c or "亿元" in c for c in companies):
                return len(set(companies))

    # 情况 B：横向对比图（首行首列是“公司名称”，第 2..N 列分别是各公司名，如 图表 28）
    if any(k in h0 for k in ("公司", "厂商", "企业")) and len(header) >= 4:
        peer_cols = [c for c in header[1:] if c and not re.match(r"^(20\d\d|Q\d|指标|占比)", c)]
        if len(peer_cols) >= 3:
            return len(set(peer_cols))

    return 0


def _count_distinct_doc_types(source_name: str) -> int:
    """统计 source_name 中提及的不同公司官方披露文档种类数量（如招股书+年报+调研表=3）。"""
    src = source_name or ""
    if "各上市公司" in src or "各公司" in src:
        return 2
    hit_types = 0
    for _, kws in DOC_TYPE_PATTERNS:
        if any(kw in src for kw in kws):
            hit_types += 1
    return hit_types


def _split_distinct_sources(source_name: str) -> List[str]:
    """
    按顿号/分号/斜杠拆分 source_name 中的独立机构或文档来源（剔除括号注释、自家研究所与数据库词）。
    用于自动识别“写了 A、B、C 三个来源，却只贴了 1 张图”的多源遗漏漏洞。
    """
    if not source_name:
        return []
    # 去除括号内日期/说明，仅按顿号/分号/斜杠拆分（不按“及”拆分，避免误拆“工信部及发改委联合通知”）
    cleaned = re.sub(r"[（(][^）)]*[）)]", "", source_name)
    parts = re.split(r"[、;；/]", cleaned)
    ignore_kws = ("国金证券", "研究所", "ifind", "wind", "万得", "同花顺", "choice")
    sources: List[str] = []
    for p in parts:
        tok = p.strip()
        if len(tok) < 2:
            continue
        if any(ig in tok.lower() for ig in ignore_kws):
            continue
        sources.append(tok)
    return sources


def _load_screenshot_verified_corpus(shot_abs: str, shot_entry: Dict[str, Any]) -> Tuple[str, bool, int]:
    """
    读取单张截图对应的已验证文本语料（堵死“空白图 + 自填 covered_claims 蒙混过关”漏洞）：
    - 当同目录存在 `<shot_abs>.meta.json`（由 auto_evidence_capture.py 真实截屏落盘）时：
      **仅信任真实抓取的 `page_text_norm` 与 `verified_tokens`**，严禁用手写 caption/covered_claims 绕过核验；
    - 当不存在 `.meta.json` 时：仅在显式人工视觉签认 `visually_verified=true` 时才采信 `covered_claims`，否则返回空语料并触发缺失元数据告警。
    返回 (归一化文本, 是否存在meta文件, 红框数量)。
    """
    chunks: List[str] = []
    has_meta = False
    box_count = -1

    meta_path = f"{shot_abs}.meta.json"
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r", encoding="utf-8") as mf:
                meta = json.load(mf)
            has_meta = True
            box_count = int(meta.get("matched_box_count", -1))
            # 仅当红框数 > 0 时才计入有效覆盖语料，防止无红框白板图通过原子事实核验
            if box_count != 0:
                chunks.append(meta.get("page_text_norm", ""))
                chunks.extend(meta.get("verified_tokens", []) or [])
        except Exception:
            pass
    elif shot_entry.get("visually_verified") is True:
        # 仅允许显式人工视觉复核签认的旧图/手工红框图走 covered_claims 兜底
        for c in (shot_entry.get("covered_claims") or []):
            chunks.append(str(c))
        if shot_entry.get("caption"):
            chunks.append(str(shot_entry.get("caption")))

    return _norm("".join(chunks)), has_meta, box_count


def audit_manifest(manifest: Dict[str, Any], base_dir: str) -> Dict[str, Any]:
    """
    执行【六道质控门禁】全量审计，返回结构化体检报告字典：
    - red_errors: 严重缺陷列表（黑名单、缺截图、多源/多公司仅贴单图、原子事实缺口、无红框图）
    - yellow_warnings: 需复核提示（裸根域名网址、无meta sidecar等）
    - missing_shots: 待补截图清单
    - stats: 量化统计指标
    """
    red_errors: List[str] = []
    yellow_warnings: List[str] = []
    missing_shots: List[str] = []

    stats = {
        "total_items": 0,
        "db_exempt": 0,
        "single_shot_items": 0,
        "multi_shot_items": 0,
        "total_embedded_shots": 0,
        "missing_shot_items": 0,
        "claims_checked_items": 0,
        "claims_100pct_items": 0,
    }

    for sec in manifest.get("sections", []):
        sname = sec.get("sheet_name", "未命名章节")
        for item in sec.get("items", []):
            stats["total_items"] += 1
            src_name = item.get("source_name", "")
            src_url = item.get("source_url", "")
            label = item.get("label", "未命名条目")
            raw_data = item.get("raw_data", []) or []
            shot_entries = extract_shot_entries(item)

            # ---- 门禁 1：黑名单来源硬拦截（忽略英文大小写，防 morgan stanley 漏网） ----
            cap_concat = " ".join(str(se.get("caption", "")) for se in shot_entries)
            combined = f"{src_name} {src_url} {cap_concat}"
            combined_low = combined.lower()
            for bad in BLACKLIST_SOURCES:
                if bad.lower() in combined_low:
                    red_errors.append(
                        f"🔴 [门禁1·黑名单来源] Sheet「{sname}」{label} -> 命中禁用词「{bad}」（来源：{src_name}）"
                    )

            # 数据库来源免截图豁免（支持显式 is_database_exempt=true 或自动识别 iFinD/Wind 等关键词）
            if (item.get("is_database_exempt") or is_database_source(src_name)) and not shot_entries:
                stats["db_exempt"] += 1
                continue

            # ---- 门禁 2：截图存在性检查 ----
            valid_shots: List[Tuple[str, Dict[str, Any]]] = []
            for se in shot_entries:
                p_abs = resolve_path(se.get("path", ""), base_dir)
                if p_abs and os.path.exists(p_abs):
                    valid_shots.append((p_abs, se))
                elif se.get("path"):
                    red_errors.append(
                        f"🔴 [门禁2·截图文件丢失] Sheet「{sname}」{label} -> 路径不存在：{se.get('path')}"
                    )

            if not valid_shots:
                stats["missing_shot_items"] += 1
                missing_shots.append(f"  - [{sname}] {label} | 需补截图来源：{src_name or '未填写'}")
                if _is_bare_homepage_url(src_url):
                    yellow_warnings.append(
                        f"🟡 [门禁5·缺图条目且网址仅为根域名] Sheet「{sname}」{label} -> {src_url}"
                    )
                continue

            stats["total_embedded_shots"] += len(valid_shots)
            if len(valid_shots) >= 2:
                stats["multi_shot_items"] += 1
            else:
                stats["single_shot_items"] += 1

            # ---- 门禁 3：多公司 / 多文档 / 复合图数量门禁（核心防遗漏，支持显式单源豁免） ----
            single_ok = bool(
                item.get("single_source_covers_all", False)
                or item.get("single_source_verified", False)
            )
            peer_cnt = _detect_peer_company_count(raw_data, src_name)
            if peer_cnt >= 3 and len(valid_shots) < peer_cnt and not single_ok:
                red_errors.append(
                    f"🔴 [门禁3·多公司对比表佐证不足] Sheet「{sname}」{label} -> "
                    f"表格对比了 {peer_cnt} 家公司，但仅附了 {len(valid_shots)} 张佐证截图！必须逐家公司贴齐官方出处。"
                )

            doc_types = _count_distinct_doc_types(src_name)
            distinct_srcs = _split_distinct_sources(src_name)
            if doc_types >= 2 and len(valid_shots) < 2 and not single_ok:
                red_errors.append(
                    f"🔴 [门禁3·多文档组合来源仅贴单图] Sheet「{sname}」{label} -> "
                    f"来源涉及 {doc_types} 类独立文档（{src_name}），但仅贴了 {len(valid_shots)} 张截图！请使用 evidence_screenshots 纵向堆叠补齐。"
                )
            elif len(distinct_srcs) >= 2 and len(valid_shots) < 2 and not single_ok:
                red_errors.append(
                    f"🔴 [门禁3·多机构来源仅贴单图] Sheet「{sname}」{label} -> "
                    f"来源列出了 {len(distinct_srcs)} 个独立出处 {distinct_srcs}，但仅贴了 1 张图！"
                    f"必须逐源堆叠补齐截图，或删除未实际引用的冗余出处名（若单图确实全覆盖，须同时填写 single_source_covers_all=true 与完整 required_claims）。"
                )

            # 复合多赛道/多维度表格检查：表有 ≥3 行业务内容且单行信息丰富（含多个分号/顿号/竖线），但只贴 1 张图
            if len(valid_shots) == 1 and not single_ok and len(raw_data) >= 4:
                rich_rows = sum(
                    1
                    for row in raw_data[1:]
                    if sum(len(str(c or "")) for c in row) >= 28
                    and any(sep in " ".join(str(c or "") for c in row) for sep in ("｜", "|", "、", "；", "/"))
                )
                if rich_rows >= 3 and not any(k in src_name.lower() for k in SINGLE_RANKING_KEYWORDS):
                    red_errors.append(
                        f"🔴 [门禁3·复合信息表仅贴单图] Sheet「{sname}」{label} -> "
                        f"含 {rich_rows} 行复合业务事实但仅贴 1 张图！必须逐行补齐对应佐证截图，或显式标注 single_source_covers_all=true 并通过 required_claims 全量核验。"
                    )

            # 严防滥用 single_source_covers_all 豁免：启用单图豁免时强制要求提供非空 required_claims
            req_claims = item.get("required_claims") or []
            if single_ok and not req_claims:
                red_errors.append(
                    f"🔴 [门禁4·单图豁免缺少原子事实清单] Sheet「{sname}」{label} -> "
                    f"启用了 single_source_covers_all=true，但未提供 required_claims 进行逐项核验！"
                )

            # ---- 门禁 4：原子事实 required_claims 100% 覆盖核验 ----
            union_corpus_parts: List[str] = []
            for p_abs, se in valid_shots:
                corpus, has_meta, box_cnt = _load_screenshot_verified_corpus(p_abs, se)
                union_corpus_parts.append(corpus)
                if has_meta and box_cnt == 0:
                    red_errors.append(
                        f"🔴 [门禁4·无红框白板截图] Sheet「{sname}」{label} -> {os.path.basename(p_abs)} 红框数为 0！"
                    )
                elif not has_meta and se.get("visually_verified") is not True:
                    yellow_warnings.append(
                        f"🟡 [门禁4·缺少.meta.json溯源元数据] Sheet「{sname}」{label} -> {os.path.basename(p_abs)} "
                        f"缺少机器溯源元数据（请用 auto_evidence_capture.py 重截，或人工复核红框后显式标注 visually_verified=true）。"
                    )
            union_corpus = "".join(union_corpus_parts)

            if req_claims:
                stats["claims_checked_items"] += 1
                uncovered = [c for c in req_claims if _norm(str(c)) not in union_corpus]
                if uncovered:
                    red_errors.append(
                        f"🔴 [门禁4·原子事实未覆盖] Sheet「{sname}」{label} -> "
                        f"已贴 {len(valid_shots)} 张图仍缺少以下事实：{uncovered}"
                    )
                else:
                    stats["claims_100pct_items"] += 1

            # ---- 门禁 5：精确网址/页码核查 ----
            if _is_bare_homepage_url(src_url):
                yellow_warnings.append(
                    f"🟡 [门禁5·网址未精确到具体文章/公告页] Sheet「{sname}」{label} -> 当前网址仅为首页域名：{src_url}"
                )

    return {
        "red_errors": red_errors,
        "yellow_warnings": yellow_warnings,
        "missing_shots": missing_shots,
        "stats": stats,
    }


def check_compliance(manifest: Dict[str, Any], base_dir: str = ".") -> Tuple[List[str], List[str]]:
    """向下兼容旧接口：返回 (红线告警列表, 待补截图清单)。"""
    rep = audit_manifest(manifest, base_dir)
    return rep["red_errors"], rep["missing_shots"]


def init_section_sheet(ws) -> None:
    """初始化分节底稿 Sheet 的列宽与顶部左右分栏标题行。"""
    col_widths = {
        "A": 18, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14,
        "H": 3,
        "I": 12, "J": 18, "K": 18, "L": 18, "M": 18, "N": 16, "O": 16, "P": 16,
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    ws.merge_cells("A1:G1")
    ws["A1"] = "报告图 / 表 / 正文关键摘录及基础数据（须显示研报中编号）"
    ws["A1"].font = FONT_HEADER
    ws["A1"].fill = FILL_NAVY
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("I1:P1")
    ws["I1"] = "公开数据来源、原始网址与完整截图（需含来源名称+原文+页码/网址）"
    ws["I1"].font = FONT_HEADER
    ws["I1"].fill = FILL_NAVY
    ws["I1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24


def fill_main_data_sheet(wb: openpyxl.Workbook, manifest: Dict[str, Any]) -> None:
    """填充模板中的《主要数据》Sheet（资料收集人分工 + 摘要/正文关键数据）。"""
    ws = wb["主要数据"] if "主要数据" in wb.sheetnames else wb.create_sheet("主要数据")
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 42

    ws["A1"] = "资料收集人"
    ws["B1"] = "负责章节"
    for col in ("A", "B"):
        ws[f"{col}1"].font = FONT_HEADER
        ws[f"{col}1"].fill = FILL_NAVY
        ws[f"{col}1"].alignment = Alignment(horizontal="center", vertical="center")

    r = 2
    for c in manifest.get("collectors", []):
        ws.cell(r, 1, c.get("name", "")).font = FONT_NORMAL
        ws.cell(r, 2, c.get("sections", "")).font = FONT_NORMAL
        ws.cell(r, 1).border = BORDER_THIN
        ws.cell(r, 2).border = BORDER_THIN
        r += 1

    start_row = max(r + 2, 10)
    headers = ["关键数据项", "数值/口径", "报告出现位置", "公开资料来源", "网址/页码"]
    for idx, h in enumerate(headers, 1):
        cell = ws.cell(start_row, idx, h)
        cell.font = FONT_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = Alignment(horizontal="center", vertical="center")

    kr = start_row + 1
    for kd in manifest.get("key_data", []):
        vals = [
            kd.get("item", ""),
            kd.get("value", ""),
            kd.get("location", ""),
            kd.get("source", ""),
            kd.get("url", ""),
        ]
        for idx, v in enumerate(vals, 1):
            c = ws.cell(kr, idx, v)
            c.font = FONT_NORMAL
            c.border = BORDER_THIN
            c.alignment = Alignment(wrap_text=True, vertical="center")
        kr += 1


def build_workbook(
    template_path: str,
    manifest: Dict[str, Any],
    base_dir: str,
    out_path: str,
    img_max_width: int = 1400,
    img_quality: int = 90,
) -> float:
    """核心构建函数：生成正式工作底稿 Excel，返回文件大小（MB）。"""
    wb = openpyxl.load_workbook(template_path)

    # 正式底稿自动移除模板示例页及两页非数据说明页
    for demo_sheet in [
        "发布底稿要求",
        "数字未来实验室-ID",
        "图示例",
        "表示例",
        "文字示例",
        "1.1",
        "1.2",
    ]:
        if demo_sheet in wb.sheetnames:
            del wb[demo_sheet]

    fill_main_data_sheet(wb, manifest)

    for sec in manifest.get("sections", []):
        sheet_name = str(sec.get("sheet_name", "未命名"))[:31]
        ws = wb.create_sheet(title=sheet_name)
        init_section_sheet(ws)

        cur_row = 3
        for item in sec.get("items", []):
            itype = item.get("type", "chart")
            label = item.get("label", "未命名条目")

            # -------- 条目横幅（左侧 A:G）--------
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=7)
            c_banner = ws.cell(cur_row, 1, label)
            c_banner.font = FONT_BANNER
            c_banner.fill = FILL_SUBHEADER
            c_banner.alignment = Alignment(vertical="center")
            ws.row_dimensions[cur_row].height = 20

            # -------- 右侧来源信息（I:P）--------
            src_name = item.get("source_name", "")
            src_url = item.get("source_url", "")
            ws.cell(cur_row, 9, "资料来源：").font = FONT_BOLD
            ws.merge_cells(start_row=cur_row, start_column=10, end_row=cur_row, end_column=16)
            ws.cell(cur_row, 10, src_name).font = FONT_NORMAL

            ws.cell(cur_row + 1, 9, "网址/页码：").font = FONT_BOLD
            ws.merge_cells(start_row=cur_row + 1, start_column=10, end_row=cur_row + 1, end_column=16)
            url_cell = ws.cell(cur_row + 1, 10, src_url)
            if src_url.startswith(("http://", "https://")) and ";" not in src_url:
                url_cell.hyperlink = src_url.split()[0]
                url_cell.font = FONT_LINK
            else:
                url_cell.font = FONT_NORMAL

            ws.cell(cur_row + 2, 9, "截图/佐证：").font = FONT_BOLD

            left_end_row = cur_row + 1
            right_end_row = cur_row + 3

            # -------- 左侧内容：报告图 / 文字摘录 / 基础数据表 --------
            if itype == "chart":
                chart_abs = resolve_path(item.get("chart_image", ""), base_dir)
                if chart_abs and os.path.exists(chart_abs):
                    buf, w, h = compress_image_to_bytes(chart_abs, max_width=img_max_width, quality=img_quality)
                    xl_img = XLImage(buf)
                    scale = min(1.0, 480.0 / max(w, 1))
                    xl_img.width = int(w * scale)
                    xl_img.height = int(h * scale)
                    ws.add_image(xl_img, f"A{left_end_row}")
                    left_end_row += max(10, int(xl_img.height / 18) + 2)

            if item.get("excerpt"):
                ws.merge_cells(start_row=left_end_row, start_column=1, end_row=left_end_row + 3, end_column=7)
                ex_cell = ws.cell(left_end_row, 1, item["excerpt"])
                ex_cell.font = FONT_NORMAL
                ex_cell.alignment = Alignment(wrap_text=True, vertical="top")
                ex_cell.border = BORDER_THIN
                left_end_row += 5

            raw_data = item.get("raw_data", [])
            if raw_data:
                for r_i, row_vals in enumerate(raw_data):
                    max_lines = 1
                    for c_i, val in enumerate(row_vals[:7], 1):
                        cell = ws.cell(left_end_row, c_i, val)
                        cell.border = BORDER_THIN
                        cell.alignment = Alignment(wrap_text=True, vertical="center")
                        if r_i == 0:
                            cell.font = FONT_BOLD
                            cell.fill = FILL_TABLE_HEAD
                        else:
                            cell.font = FONT_NORMAL
                        sval = str(val or "")
                        max_lines = max(max_lines, (len(sval) // 11) + 1)
                    ws.row_dimensions[left_end_row].height = max(18, min(72, max_lines * 15))
                    left_end_row += 1

            # -------- 右侧内容：单图或多图纵向堆叠全覆盖 --------
            shot_entries = extract_shot_entries(item)
            valid_shots = []
            for se in shot_entries:
                s_abs = resolve_path(se.get("path", ""), base_dir)
                if s_abs and os.path.exists(s_abs):
                    valid_shots.append((s_abs, se))

            if valid_shots:
                r_cursor = cur_row + 2
                for idx_s, (s_abs, se) in enumerate(valid_shots):
                    cap = se.get("caption", "")
                    s_url = se.get("url", "")
                    cov = se.get("covered_claims") or []
                    hbox = se.get("highlight_box")
                    if cap or cov:
                        ws.merge_cells(start_row=r_cursor, start_column=10, end_row=r_cursor, end_column=16)
                        cov_suffix = f"（覆盖：{' / '.join(str(x) for x in cov)}）" if cov else ""
                        cap_cell = ws.cell(r_cursor, 10, f"佐证 {idx_s + 1}/{len(valid_shots)}：{cap}{cov_suffix}")
                        cap_cell.font = FONT_BOLD
                        cap_cell.fill = FILL_TABLE_HEAD
                        r_cursor += 1
                    if s_url:
                        ws.merge_cells(start_row=r_cursor, start_column=10, end_row=r_cursor, end_column=16)
                        su_cell = ws.cell(r_cursor, 10, s_url)
                        if s_url.startswith(("http://", "https://")):
                            su_cell.hyperlink = s_url
                            su_cell.font = FONT_LINK
                        else:
                            su_cell.font = FONT_NORMAL
                        r_cursor += 1
                    buf, w, h = compress_image_to_bytes(
                        s_abs,
                        max_width=img_max_width,
                        quality=img_quality,
                        highlight_box=hbox,
                    )
                    xl_shot = XLImage(buf)
                    scale = min(1.0, 560.0 / max(w, 1))
                    xl_shot.width = int(w * scale)
                    xl_shot.height = int(h * scale)
                    ws.add_image(xl_shot, f"J{r_cursor}")
                    r_cursor += max(10, int(xl_shot.height / 18) + 2)
                right_end_row = r_cursor
            elif item.get("is_database_exempt") or is_database_source(src_name):
                ws.merge_cells(start_row=cur_row + 2, start_column=10, end_row=cur_row + 2, end_column=16)
                ws.cell(cur_row + 2, 10, "【数据库导出数据：左侧已列示导出基础数据与来源名称，合规免截图】").font = FONT_NORMAL
            else:
                ws.merge_cells(start_row=cur_row + 2, start_column=10, end_row=cur_row + 2, end_column=16)
                prompt_txt = (
                    "【待补关键文字红框截图：请截取完整来源原文并用红框框出行业排名/市占率关键表述】"
                    if itype == "text"
                    else "【待补充完整来源截图：须含来源名称+原文+网址/页码，关键数据建议画红框】"
                )
                warn_cell = ws.cell(cur_row + 2, 10, prompt_txt)
                warn_cell.font = FONT_BOLD
                warn_cell.fill = FILL_WARN

            cur_row = max(left_end_row, right_end_row) + 2

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)
    return os.path.getsize(out_path) / (1024.0 * 1024.0)


def print_audit_report(report: Dict[str, Any]) -> None:
    """打印结构化的六道质控门禁体检总表。"""
    st = report["stats"]
    reds = report["red_errors"]
    yellows = report["yellow_warnings"]
    miss = report["missing_shots"]

    print("\n" + "=" * 68)
    print("【国金工作底稿·六道质控门禁体检汇总】")
    print(
        f"  总条目：{st['total_items']} | 数据库豁免：{st['db_exempt']} | "
        f"单图佐证：{st['single_shot_items']} | 多图堆叠佐证：{st['multi_shot_items']}（累计嵌入 {st['total_embedded_shots']} 张高清图）"
    )
    print(
        f"  待补截图：{st['missing_shot_items']} | 原子事实核验通过：{st['claims_100pct_items']}/{st['claims_checked_items']} | "
        f"🔴 严重缺陷：{len(reds)} | 🟡 提醒：{len(yellows)}"
    )
    print("=" * 68)

    if reds:
        print("\n【🔴 严重质控缺陷（必须修复，严禁带病交付）】")
        for r in reds:
            print(" ", r)

    if yellows:
        print("\n【🟡 质控提醒（建议复核确认）】")
        for y in yellows:
            print(" ", y)

    if miss:
        print("\n【待补截图清单（无公开原件或付费报告，已标黄留空待人工补图）】")
        for line in miss:
            print(line)

    if not reds and not miss:
        print("\n[PASS] 六道质控门禁全绿（0 🔴、0 待补截图），满足正式发布标准！")
    elif not reds and miss:
        print(f"\n[INCOMPLETE] 0 🔴 规则违规，但仍存在 {len(miss)} 项待补截图（仅供草稿预览；正式交付须补齐并通过 --release 门禁）！")


def main() -> None:
    """命令行入口：执行六道质控门禁校验，并生成高清正式工作底稿 Excel。"""
    default_tpl = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "模板资产",
        "国金研报工作底稿模板.xlsx",
    )
    parser = argparse.ArgumentParser(description="国金证券合规工作底稿 Excel 自动生成与六道质控门禁脚本")
    parser.add_argument("--manifest", required=True, help="06_证据清单.json 文件路径")
    parser.add_argument("--out", default="", help="输出工作底稿 (.xlsx) 路径（--audit-only 时可不传）")
    parser.add_argument("--template", default=default_tpl, help="国金官方底稿模板路径")
    parser.add_argument("--strict", action="store_true", help="严格模式（草稿门禁）：存在任一 🔴 严重缺陷即以退出码 1 阻断")
    parser.add_argument("--release", action="store_true", help="正式发布门禁：存在任一 🔴 缺陷或任一待补截图/裸根域名即以退出码 1 阻断")
    parser.add_argument("--audit-only", action="store_true", help="仅运行六道门禁审计，不写 Excel 文件")
    args = parser.parse_args()

    with open(args.manifest, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    base_dir = os.path.dirname(os.path.abspath(args.manifest))
    report = audit_manifest(manifest, base_dir)
    print_audit_report(report)

    if args.release:
        release_blockers = list(report["red_errors"])
        if report["missing_shots"]:
            release_blockers.append(f"存在 {len(report['missing_shots'])} 项待补截图未闭环")
        bare_url_warns = [w for w in report["yellow_warnings"] if "门禁5" in w]
        if bare_url_warns:
            release_blockers.append(f"存在 {len(bare_url_warns)} 项裸根域名网址未精确到文章/页码")
        if release_blockers:
            print(f"\n[BLOCKED] --release 正式发布门禁阻断：{release_blockers}")
            sys.exit(1)
    elif args.strict and report["red_errors"]:
        print("\n[BLOCKED] --strict 严格门禁已触发：请修复上述 🔴 严重缺陷后再出稿！")
        sys.exit(1)

    if args.audit_only:
        return

    if not args.out:
        print("[ERROR] 非 --audit-only 模式必须传入 --out 输出路径。")
        sys.exit(2)

    size_mb = build_workbook(args.template, manifest, base_dir, args.out, img_max_width=1400, img_quality=90)
    print(f"\n[OK] 已生成国金合规工作底稿：{args.out}（当前体积：{size_mb:.2f} MB，高清红框模式）")


if __name__ == "__main__":
    main()
