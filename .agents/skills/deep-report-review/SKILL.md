---
name: deep-report-review
description: 深度报告 Step 3.5 六维独立联网红队审稿与版本哈希签认。支持 Antigravity 多模型派发（首选 Opus 5.5，可切 Sonnet/Gemini）以及 ChatGPT Work / Codex 单客户端独立红队自审，严禁跳过审稿。当用户说“红队审稿”“Opus审稿”“复审报告”“审稿签认”或“/deep-report-review”时触发。
---

# 深度报告独立红队审稿（跨客户端通用版）

本 Skill 负责在出图渲染前对 `05_完整初稿.md` 与 `report_meta.json` 执行独立事实核查、口径核对与合规审稿，并写入机器可读的 `REVIEW_META` 版本指纹。

## 核心原则

1. **多客户端自适应，严禁以“缺少特定模型”为由跳过审稿**：
   - **环境 A（支持多模型派发，如 Antigravity `swarm`）**：首选派发内置 `Claude Opus 5.5 (Medium/High)`，配额紧张时可无缝切 `Claude Sonnet 5.5 (High)` 或零历史独立 `Gemini 4 (High)`；
   - **环境 B（单客户端，如 ChatGPT Work / Codex）**：直接以“质控审核员 + 卖方研究总监”的独立红队角色，按任务包六大维度逐项调用联网搜索核查正文并落盘 `05_红队审稿意见.md`；
   - **环境 C（外部手工复审回填）**：将 `05_红队审稿任务包.md` + `05_完整初稿.md` 交外部模型复审后贴回 `05_红队审稿意见.md`。
2. **高效路径 = 直接扫正文 + 联网核查**：默认不全文翻阅大 PDF，仅遇异常数字或分歧时定点查阅对应页码。
3. **版本哈希绑定**：审稿报告顶部必须写入与当前稿件一致的 `draft_sha256` 与 `verdict`。

## 执行步骤（在根目录运行）

### 1. 生成最新精简审稿任务包（含 `draft_sha256` 与本地预扫描线索）
```bash
python 脚本与规范/redteam_review.py \
  --md <公司名>/05_完整初稿.md \
  --meta <公司名>/report_meta.json \
  --snapshot <公司名>/00_原始资料/01_最新财报全科目快照.md \
  --manifest <公司名>/06_证据清单.json \
  --out <公司名>/05_红队审稿任务包.md
```
（若非 A 股未生成财报快照，可省略 `--snapshot` 参数。）

### 2. 执行六维独立联网红队审稿并落盘
对照 `<公司名>/05_红队审稿任务包.md` 与 `<公司名>/05_完整初稿.md`：
- 坚持**建设性/偏乐观卖方深度框架**，重点对照任务包**六大必查维度**抓硬伤：① 近 90 天临时公告漏项与最新股本/实控人持股；② 会计口径、单季/累计/跨期对比、广义现金/所得税/少数股东损益及公告原表常识勾稽；③ 技术协议标准速率、物理规格与工艺先后顺序；④ 友商/非白名单来源与首页摘要未覆盖可比公司禁令；⑤ 图表标题起止年份与横轴一致性、正文违规分析财务费用率或连续堆砌数字；⑥ 无出处定性狠话与推理跳步；
- 调用联网搜索交叉核验关键行业与公司最新公开数据；
- 将审稿结果写入 `<公司名>/05_红队审稿意见.md`（兼容 `_Claude.md`），顶部**必须**包含：
  ```markdown
  <!-- REVIEW_META
  draft_sha256: <任务包中的16位指纹>
  verdict: PASS | CONDITIONAL_PASS | REJECT
  reviewer_model: <实际执行的模型名，如 claude-opus-5-5 / gpt / gemini-4>
  reviewed_at: YYYY-MM-DD HH:MM
  -->
  ```

### 3. 定点修订
- 若 `verdict` 为 `REJECT` 或 `CONDITIONAL_PASS`：按审稿意见定点修订 `05_完整初稿.md`、`report_meta.json`、`charts.json` 与 `06_证据清单.json`（保留未被指出的原有段落，不做无故全文重写）；
- 若后续需要通过 `--release` 终验，在最终定稿后复跑 Step 1 + Step 2 签发匹配最新 `draft_sha256` 的 `PASS`。
