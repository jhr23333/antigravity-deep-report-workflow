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
    python 脚本与规范/redteam_review.py \
      --md 示例公司/05_完整初稿.md \
      --meta 示例公司/report_meta.json \
      --snapshot 示例公司/00_原始资料/01_最新财报全科目快照.md \
      --manifest 示例公司/06_证据清单.json \
      --out 示例公司/05_红队审稿任务包.md
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

# 去 AI 味套词清单（与 style_guide.md / render_docx.py 保持同步）
AI_FLUFF_PHRASES = [
    "托底",
    "双击",
    "压舱石",
    "进攻矛",
    "定海神针",
    "攻守兼备",
    "梯度接力",
    "形成合力",
    "提供坚实底座",
    "打开向上弹性",
    "深度赋能",
    "理性看待",
    "双轮驱动成型",
]


def compute_draft_sha256(md_text: str, meta_text: str) -> str:
    """计算当前初稿正文与封面摘要的联合 SHA-256 指纹（取前 16 位），用于审稿结论版本绑定。"""
    h = hashlib.sha256()
    h.update(md_text.encode("utf-8"))
    h.update(b"\n---META---\n")
    h.update(meta_text.encode("utf-8"))
    return h.hexdigest()[:16]


def scan_local_risks(md_text: str, meta_text: str = "") -> str:
    """本地确定性预扫描：揪出宣传化措辞、疑似累计/环比混写、违规分析财务费用率及单段数字过度堆砌。"""
    import re as _re

    hits: List[str] = []
    num_unit_re = _re.compile(r"\d+(?:\.\d+)?(?:%|pct|个百分点|亿元|万元|亿美元|百万美元|亿日元|万股)")
    in_comment = False

    for idx, line in enumerate(md_text.splitlines(), 1):
        s = line.strip()
        if s.startswith("<!--"):
            in_comment = True
        if in_comment:
            if "-->" in s:
                in_comment = False
            continue
        for w in PROMOTIONAL_PHRASES:
            if w in s:
                hits.append(f"- [宣传措辞] 正文第 {idx} 行命中「{w}」：{s[:100]}")
        for aw in AI_FLUFF_PHRASES:
            if aw in s:
                hits.append(f"- [去AI味提醒] 正文第 {idx} 行命中模板化套词「{aw}」：{s[:100]}")
        if ("前三季度" in s or "累计" in s) and "环比" in s and "较上半年" not in s and "单季" not in s:
            hits.append(f"- [疑似累计/环比混用] 正文第 {idx} 行：{s[:110]}")
        if s and not s.startswith(("#", "|", ">")):
            if "财务费用率" in s:
                hits.append(f"- [违反文风红线·财务费用率] 正文第 {idx} 行不应展开分析财务费用率：{s[:90]}")
            n_cnt = len(num_unit_re.findall(s))
            if n_cnt >= 9:
                hits.append(f"- [疑似连续数据堆砌] 正文第 {idx} 行单段含 {n_cnt} 处带单位数值，建议精简并由图表呈现")

    if meta_text:
        for w in PROMOTIONAL_PHRASES:
            if w in meta_text:
                hits.append(f"- [宣传措辞] 封面摘要（report_meta.json）命中「{w}」")
        for aw in AI_FLUFF_PHRASES:
            if aw in meta_text:
                hits.append(f"- [去AI味提醒] 封面摘要（report_meta.json）命中模板化套词「{aw}」")

    # 质控部风险提示强制排查项提醒：前 3 月/后 6 月解禁与大股东高质押（≥50%）
    full_all = md_text + "\n" + meta_text
    if "解禁" not in full_all and "质押" not in full_all:
        hits.append(
            "- [质控风险必查提醒] 全文及摘要均未提及限售股解禁或大股东质押核查结果，请审稿时确认近3月/后6月无解禁计划且大股东质押比例<50%"
        )

    if not hits:
        return "（本地预扫描：0 处宣传化措辞、累计/环比混写或财务费用率/数字堆砌告警）"
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
                src_low = src.lower()
                is_db = bool(it.get("is_database_exempt")) or any(
                    k in src_low for k in ("ifind", "wind", "choice", "万得", "同花顺", "bloomberg")
                )
                if is_db and not it.get("evidence_screenshots") and not it.get("evidence_screenshot"):
                    continue
                rows.append(f"- [{sname}] {it.get('label', '')[:42]} -> 来源：{src}")
        return "\n".join(rows[:45]) if rows else "（全部为合规数据库来源）"
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
    """生成给多客户端（Antigravity / Codex / ChatGPT Work）通用的精简高效红队审稿任务包。"""
    return f"""# Step 3.5 深度报告——独立联网红队审稿任务包

> **执行模型（跨客户端兼容）**：有异构多 Agent 环境（如 Antigravity `swarm`）时首选内置 **Claude Opus 5.5**（可切 Sonnet 5.5 / Gemini 4）；在单对话框客户端（如 **ChatGPT Work / Codex**）中，必须以“独立红队审稿员”视角在本步骤逐项联网核查并落盘 `05_红队审稿意见.md`，**严禁以“没有指定模型/工具”为由跳过审稿**！
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

## 必查六大红队维度
1. **数据时效性与近 90 天公告核查（联网查新）**：核查公司最近 90 天全部临时公告（近期生产经营公告、自建基地开工、H股筹划/回购、最新定增发行结果）及行业机构近 60 天最新预测（防引用数月前旧预测），同时核对**实控人及一致行动人持股比例是否以最新发行/权益变动公告为准**（勿轻信数据库衍生值）。
2. **会计口径与常识勾稽**：累计 vs 单季毛利率、同期间同比 vs 跨期对比（如前三季度 vs 上年全年需明示）；单季利润总额转负但归母净利为正时须核查**所得税费用与少数股东损益**；核查**交易性金融资产（广义现金）**与反推**最新单季经营现金流**；核查公司原表是否存在常识勾稽异常（如在研项目累计投入远超预计总投资）以及**含税重大合同金额是否被误当成不含税收入或尚未确收余额**。
3. **技术规格、协议标准与测算边界**：产品/工艺代际、物理规格（如基板尺寸）、通信/存储接口协议标准速率（如 ONFI 6.0 最高 4,800 MT/s）与测试机台自身速率必须区分；工艺流程表必须按物理先后顺序拆分（如堆叠前 KGD 裸片测试 vs 堆叠后 TSV 互连测试）。
4. **合规白名单红线**：正文与来源中是否混入友商卖方自制测算（含以权威机构名义转引券商国产化率测算）、个人自媒体或空泛来源？首页摘要是否误列未覆盖可比公司名称？
5. **文风纪律与图表标题一致性**：正文严禁连续堆砌一连串逐年/逐季数字清单（让图表说话，正文讲逻辑趋势），严禁分析财务费用率；核对每张图表标题标注的起止年份是否与实际横轴一致。
6. **措辞克制与推理合理性**：针对尚未批量交付的前瞻业务，是否存在过度确定或偏宣传化表述？从终端市场收入到设备需求的传导逻辑是否存在明显跳步？

## 输出与版本绑定硬要求
将结构化审稿意见写入同目录 `05_红队审稿意见.md`（兼容历史文件名 `05_红队审稿意见_Claude.md`），**文件前 6 行必须包含如下机器可读头信息（供 sync_check.py 自动核对版本一致性）**：
```markdown
<!-- REVIEW_META
draft_sha256: {draft_sha}
reviewer_model: claude-opus-5-5 | claude-sonnet-5-5 | gemini-4 | chatgpt
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
    print(f"[OK] 已生成精简红队审稿任务包：{args.out}（draft_sha256={draft_sha}）")
    print(pre_scan)


if __name__ == "__main__":
    main()
