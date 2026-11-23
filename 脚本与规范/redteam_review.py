# -*- coding: utf-8 -*-
"""
国金证券电子组深度报告——Step 3.5 红队审稿预扫描、稿件哈希绑定与任务包生成脚本 (redteam_review.py)

核心升级（2026-11-23 全面加固）：
1. **统一使用 Antigravity 原生内置 Claude Opus 5.5**（`claude-opus-5-5-medium` / `claude-opus-5-5-high`），严禁调用外部 Claude Code CLI；
2. **稿件版本哈希绑定（`draft_sha256`）**：对 `05_完整初稿.md + report_meta.json` 计算 SHA-256 指纹并写入任务包，
   要求 Claude 在 `05_红队审稿意见_Claude.md` 顶部回写 `draft_sha256` 与 `verdict`（PASS / CONDITIONAL_PASS / REJECT），
   供 `sync_check.py --release` 自动核对“最终交付稿是否确实通过了对应版本的复审”；
3. **高效审稿策略**：默认直接扫正文、摘要、预抽快照并调用 `search_web` 联网查新；**不全文翻阅大体积 PDF，但遇到明显异常数值或关键分歧时，允许定点查看 PDF 对应页码**；
4. **去公司硬编码**：保持通用的卖方专业护航立场，不将单个样本公司的特殊业务预设进全局模板。

命令行用法：
    python .\\脚本与规范\\redteam_review.py `
      --md .\\示例公司\\05_完整初稿.md `
      --meta .\\示例公司\\report_meta.json `
      --snapshot .\\示例公司\\00_原始资料\\01_最新财报全科目快照.md `
      --manifest .\\示例公司\\06_证据清单.json `
      --out .\\示例公司\\05_红队审稿任务包.md
"""

import argparse
import hashlib
import json
import os
from typing import List


# 偏宣传化/过度确定性措辞本地扫描词库
PROMOTIONAL_PHRASES = [
    "爆发",
    "飙升",
    "重金卡位",
    "重金布局",
    "进入收获前夜",
    "收获前夜",
    "胜负手",
    "确立拐点",
    "拐点特征明确",
    "精准锁定",
    "精准承接",
    "核心增长引擎",
    "核心增量引擎",
    "顺价传导顺畅",
    "已加速转化为订单",
    "迅速转化为",
]


def compute_draft_sha256(md_text: str, meta_text: str) -> str:
    """计算当前初稿正文与封面摘要的联合 SHA-256 指纹（取前 16 位），用于审稿结论版本绑定。"""
    h = hashlib.sha256()
    h.update(md_text.encode("utf-8"))
    h.update(b"\n---META---\n")
    h.update(meta_text.encode("utf-8"))
    return h.hexdigest()[:16]


def scan_local_risks(md_text: str, meta_text: str = "") -> str:
    """本地确定性预扫描：揪出宣传化措辞及疑似累计/环比混写。"""
    hits: List[str] = []

    for idx, line in enumerate(md_text.splitlines(), 1):
        for w in PROMOTIONAL_PHRASES:
            if w in line:
                hits.append(f"- [宣传措辞] 正文第 {idx} 行命中「{w}」：{line.strip()[:100]}")
        if ("前三季度" in line or "累计" in line) and "环比" in line and "较上半年" not in line and "单季" not in line:
            hits.append(f"- [疑似累计/环比混用] 正文第 {idx} 行：{line.strip()[:110]}")

    if meta_text:
        for w in PROMOTIONAL_PHRASES:
            if w in meta_text:
                hits.append(f"- [宣传措辞] 封面摘要（report_meta.json）命中「{w}」")

    if not hits:
        return "（本地预扫描：0 处常见宣传化措辞或明显累计/环比混写）"
    return "\n".join(hits)


def summarize_manifest_sources(manifest_path: str) -> str:
    """将 06_证据清单.json 压缩为极简来源清单（<1.5 KB），供 Opus 一眼扫查非白名单与过时出处。"""
    if not manifest_path or not os.path.exists(manifest_path):
        return "（未提供证据清单）"
    try:
        with open(manifest_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        rows = []
        for sec in data.get("sections", []):
            sname = sec.get("sheet_name", "")
            for it in sec.get("items", []):
                src = it.get("source_name", "")
                if "iFinD" in src and not it.get("evidence_screenshots") and not it.get("evidence_screenshot"):
                    continue
                rows.append(f"- [{sname}] {it.get('label', '')[:42]} -> 来源：{src}")
        return "\n".join(rows[:45]) if rows else "（全部为 iFinD 数据库来源）"
    except Exception:
        return "（证据清单读取失败）"


def build_antigravity_opus_brief(
    md_path: str,
    meta_path: str,
    draft_sha: str,
    snapshot_text: str,
    sources_summary: str,
    pre_scan: str,
) -> str:
    """生成给 Antigravity 内置 Claude Opus 5.5 的精简高效红队审稿任务包。"""
    return f"""# Step 3.5 国金深度报告——Antigravity 原生 Claude Opus 5.5 独立审稿任务包

> **执行模型**：Antigravity 内置 **Claude Opus 5.5**（`claude-opus-5-5-medium` / `claude-opus-5-5-high`，零历史独立上下文，禁用外部 Claude Code CLI）。
> **待审稿件指纹（`draft_sha256`）**：`{draft_sha}`

## 你的身份、方式与边界
1. **身份定位（卖方专业护航）**：你是国金证券质控部资深审核员 + 资深卖方电子研究总监。尊重卖方深度报告**整体偏乐观、突出核心成长主线、不过度堆砌负面风险**的框架，但凡涉及影响盈利质量或合规过审的重大事实不得隐瞒，次要结构因素用客观简短语句点出即可。
2. **高效审稿方式（直接扫正文 + 联网查证 + 异常时定点翻页）**：
   - **主路径**：直接通读待审正文与摘要（`{md_path}`、`{meta_path or "无"}`），对照下方【合并报表核心科目快照】与【非数据库来源清单】，**重点调用 `search_web` 独立联网核查**正文中的关键行业数据、最新季度披露、竞争格局、技术规格与推理合理性；
   - **PDF 使用边界**：不要从头到尾遍历大体积 PDF 原件；**仅当发现明显异常数值、口径矛盾或重大分歧时，允许定点查看对应 PDF 页码核实**。

## 本地确定性预扫描线索
{pre_scan}

## 预抽：最新一期合并报表核心科目快照（由 pull_company_data.py 按三大表分节提取）
{snapshot_text or "（未附本地快照，请结合正文内部一致性与 search_web 核查最新季报核心科目）"}

## 预抽：非数据库图表来源清单（速查友商研报/非白名单/过时来源）
{sources_summary}

## 必查五大红队维度
1. **数据时效性（联网查新）**：行业核心数据（下游资本开支、终端销量、出货预测等）及公司财务数据（在建工程/营收/净利）是否已有更新一期公开数据？
2. **会计与统计口径**：累计 vs 单季毛利率、同比 vs 环比、同榜统一版本求和一致性、产能口径是否前后自洽、利润表是否漏掉实质影响盈利的核心科目？
3. **技术规格与测算边界**：产品/工艺代际、材料体系、标准归属是否准确？行业级/单机柜测算是否与公司实际配套价值区分清楚？
4. **合规白名单红线**：正文与来源中是否混入友商卖方自制测算、个人自媒体、厂商博客或空泛来源？首页摘要是否误列未覆盖可比公司名称？
5. **措辞克制与推理合理性**：针对尚未批量交付的前瞻业务，是否存在过度确定或偏宣传化表述？从行业景气到公司盈利兑现的逻辑是否存在明显跳步？

## 输出与版本绑定硬要求
将结构化审稿意见写入同目录 `05_红队审稿意见_Claude.md`，**文件前 6 行必须包含如下机器可读头信息（供 sync_check.py 自动核对版本一致性）**：
```markdown
<!-- REVIEW_META
draft_sha256: {draft_sha}
reviewer_model: claude-opus-5-5
verdict: PASS | CONDITIONAL_PASS | REJECT
-->
```
正文中依次列出：
1. **一、优先必须修改的硬伤清单（位置 + 原文 + 问题 + 建议改写）**
2. **二、建议微调或收敛措辞的清单**
3. **三、放行结论（与上方 verdict 严格一致）**
"""


def main() -> None:
    """命令行入口：预扫描 + 计算稿件指纹 + 导出 Opus 5.5 审稿任务包。"""
    parser = argparse.ArgumentParser(description="Step 3.5 红队审稿预扫描与 Antigravity Opus 5.5 任务包生成器")
    parser.add_argument("--md", required=True, help="待审 05_完整初稿.md 路径")
    parser.add_argument("--meta", default="", help="可选 report_meta.json 路径")
    parser.add_argument("--snapshot", default="", help="可选 01_最新财报全科目快照.md 路径")
    parser.add_argument("--manifest", default="", help="可选 06_证据清单.json 路径")
    parser.add_argument("--out", required=True, help="输出的红队审稿任务包 (.md) 路径")
    args = parser.parse_args()

    with open(args.md, "r", encoding="utf-8") as f:
        md_text = f.read()

    meta_text = ""
    if args.meta and os.path.exists(args.meta):
        with open(args.meta, "r", encoding="utf-8") as f:
            meta_text = f.read()

    draft_sha = compute_draft_sha256(md_text, meta_text)

    snapshot_text = ""
    if args.snapshot and os.path.exists(args.snapshot):
        with open(args.snapshot, "r", encoding="utf-8") as f:
            snapshot_text = f.read()[:3500]

    pre_scan = scan_local_risks(md_text, meta_text)
    sources_summary = summarize_manifest_sources(args.manifest)
    brief = build_antigravity_opus_brief(
        os.path.abspath(args.md),
        os.path.abspath(args.meta) if args.meta else "",
        draft_sha,
        snapshot_text,
        sources_summary,
        pre_scan,
    )

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(brief)
    print(f"[OK] 已生成 Antigravity Opus 5.5 精简审稿任务包：{args.out}（draft_sha256={draft_sha}）")
    print(pre_scan)


if __name__ == "__main__":
    main()
