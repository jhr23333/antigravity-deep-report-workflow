---
name: deep-report-review
description: 深度报告 Step 3.5 红队审稿与终稿版本签认。生成带 draft_sha256 指纹的精简任务包，调用 Antigravity 内置 Claude Opus 5.5 直接扫正文并联网核查，输出机器可读的 REVIEW_META 审稿结论。当用户说“红队审稿”“Opus审稿”“复审报告”“审稿签认”或“/deep-report-review”时触发。
---

# 深度报告红队审稿（Antigravity Claude Opus 5.5）

本 Skill 负责在成稿渲染前（及终稿发版前）对深度报告进行独立事实核查、数据口径核对与合规审稿。

## 核心原则

1. **三级审稿模型路由（配额波动无缝切换，严禁调用外部 Claude Code CLI）**：
   - **首选（全自动）**：Antigravity 内置 **`Claude Opus 5.5 (Medium)`**（`claude-opus-5-5-medium`，配额宽裕时可用 `high`），通过 `swarm` 派发；
   - **备选 A（全自动，Opus 配额不足时）**：Antigravity 内置 **`Claude Sonnet 5.5 (High)`** 或零历史独立 **`Gemini 4 (High)`**；
   - **备选 B（外部 ChatGPT 复审）**：将 `05_红队审稿任务包.md` + `05_完整初稿.md` 交给 ChatGPT 网页/桌面端，将审稿结果保存为 `05_红队审稿意见.md`（或 `05_红队审稿意见_Claude.md`），顶部按相同格式写上 `REVIEW_META`（`reviewer_model: chatgpt`），`sync_check.py --release` 同样认账。
2. **默认快路径 = 直接扫正文 + 联网核查**：不让审稿员去翻几百页大 PDF。直接对照 `05_红队审稿任务包.md`（内含最新财报合并三大表全科目快照）+ `05_完整初稿.md` + `search_web` 逐项核查行业/公司最新事实与口径。
3. **异常定点查 PDF**：仅当正文与快照/联网信息存在明显数值冲突或关键分歧时，允许按快照标注的 PDF 页码定点核对对应 1–2 页。
4. **版本哈希绑定**：每次审稿结论必须写入与当前 `05_完整初稿.md + report_meta.json` 一致的 `draft_sha256`，供 `sync_check.py --release` 自动验真。

## 执行步骤

### 1. 生成最新精简审稿任务包（含 `draft_sha256`）
```powershell
D:\anaconda3\python.exe D:\深度报告\脚本与规范\redteam_review.py `
  --md D:\深度报告\{公司名}\05_完整初稿.md `
  --meta D:\深度报告\{公司名}\report_meta.json `
  --snapshot D:\深度报告\{公司名}\00_原始资料\01_最新财报全科目快照.md `
  --manifest D:\深度报告\{公司名}\06_证据清单.json `
  --out D:\深度报告\{公司名}\05_红队审稿任务包.md
```
记录终端打印的 `draft_sha256`（16 位指纹）。

### 2. 派发 Antigravity Opus 5.5 审稿员
使用 `swarm` 工具 `add` 一个独立审稿 worker（`model: "claude-opus-5-5-medium"` 或配额充足时用 `"claude-opus-5-5-high"`），在 brief 中明确要求：
- 读取 `05_红队审稿任务包.md` 与 `05_完整初稿.md`；
- 按**建设性/偏乐观卖方框架**审稿（不强求堆砌风险段落，重点对照任务包六大维度抓：近 90 天临时公告漏项与实控人最新股本核对、公告原表常识勾稽与含税合同执行进度、单季/累计/跨期口径与所得税/广义现金核查、技术协议标准速率与工艺流程顺序、友商/非白名单来源、图表标题年份与横轴一致性、正文违规分析财务费用率或连续堆砌数字、首页摘要超 850 字或含可比公司名）；
- 使用 `search_web` 交叉核验文中关键行业数据、最新政策/协会统计与竞品动态；
- 将审稿结果写入 `D:\深度报告\{公司名}\05_红队审稿意见.md`（兼容历史文件名 `05_红队审稿意见_Claude.md`），文件顶部**必须**包含机器可读版本头：
  ```markdown
  <!-- REVIEW_META
  draft_sha256: <任务包中的16位指纹>
  verdict: PASS | CONDITIONAL_PASS | REJECT
  reviewer_model: claude-opus-5-5-medium
  reviewed_at: YYYY-MM-DD HH:MM
  -->
  ```

### 3. 定点修订与终稿复审
- 若 `verdict` 为 `REJECT` 或 `CONDITIONAL_PASS`：按审稿意见定点修订 `05_完整初稿.md`、`report_meta.json`、`charts.json` 与 `06_证据清单.json`（保留未被指出的原有段落，不做无故全文重写）。
- **草稿期 vs 发版期复审策略**：
  - 修订过程中微调措辞不强制立刻重跑 Opus（`sync_check.py --strict` 仅打印 🟡 提醒版本哈希已变更）；
  - **终稿复审签认**：如需通过 `--release` 终验，重新跑一次 Step 1 + Step 2，让审稿模型（Opus / Sonnet / Gemini / ChatGPT）对最终定稿版本签发 `verdict: PASS`（或已全部闭环的 `CONDITIONAL_PASS`）并匹配最新 `draft_sha256`。
