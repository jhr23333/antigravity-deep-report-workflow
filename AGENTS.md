# AGENTS.md — 国金证券电子组公司深度报告工作台

> 仅维护 Antigravity / Codex 入口：读取本文件与 `.agents/skills/`。完整运行原理与脚本命令见 [README.md](README.md)。

## 工作流技能矩阵（`.agents/skills/`）

| Skill | 触发词 | 负责阶段 |
|---|---|---|
| `deep-report-run` | 写深度报告 / 新建深度报告 / 跑完整深度报告流程 | 端到端总控（Step 1–5 全链串联） |
| `deep-report-review` | 红队审稿 / Opus审稿 / 复审报告 / 审稿签认 | Step 3.5 Antigravity 内置 Claude Opus 5.5 审稿与 `draft_sha256` 签认 |
| `deep-report-charts` | 画报告图表 / 三端对账 / 生成深度报告Word / 渲染docx / 发版验收 | Step 5A 300 DPI 制图、三端【期间×数值】与正文表格对账、Word 渲染与 12 项体检 |
| `deep-report-workpaper` | 做工作底稿 / 生成底稿 / 底稿截图 / 底稿合规检查 | Step 5B 防歧义 PDF/网页红框截图、`.meta.json` 验真与六道门禁 Excel 打包 |

## 不可违背的全局硬规则（单一真源在 `脚本与规范/`）

1. **数据源分工**：常规历史财务与行情优先走同花顺 iFinD（底稿享受数据库免截图豁免）；仅对非标业务明细（产能/实名客户/财务费用拆解）及底稿原页红框截图，后台按需自动拉取 2–3 份核心公开 PDF。
2. **模型分工**：Step 3.5 红队审稿首选 **Antigravity 内置 Claude Opus 5.5**（`claude-opus-5-5-medium` / `high`），配额紧张时可无缝切换内置 `Claude Sonnet 5.5 (High)` / 零历史 `Gemini 4 (High)`，或把任务包交外部 ChatGPT 复审并回填至 `05_红队审稿意见.md`（带 `REVIEW_META`），严禁调用外部 Claude Code CLI；默认直接扫正文 + `search_web` 核查，仅遇异常数字或分歧时定点查阅 PDF 对应页。
3. **卖方文风与质控**：整体采用建设性/偏乐观卖方框架，不堆砌过多风险段落；严禁引用友商卖方研报（高盛/大摩/中信/中金/华泰等）及非白名单论坛/自媒体；首页摘要正文 `≤ 850 字` 且不出现可比公司名；未填模型的第四章数字一律保留 `[待模型填入]`，严禁杜撰 EPS/目标价/评级。
4. **双门禁制度（与“AI 留 `[待模型填入]`、研究员自刷模型”对齐）**：
   - **AI 默认交付门禁（`--strict`）**：拦死一切真实错误（黑名单、无红框/假覆盖、年份互换、表格错位、XML 缺陷），每次对账落盘 `09_对账与质检报告.md`，允许保留 `[待模型填入]` 供研究员拿到 Word 后自行刷模型（暂缺付费报告截图时打印 `[INCOMPLETE]` 标黄）。
   - **可选终验门禁（`--release`）**：仅在研究员将模型数字回填进 Markdown 且补齐全部付费报告截图、需要机器做零占位终验时执行——校验 0 占位符、0 待补截图、0 裸根域名、以及审稿报告（兼容 Opus / Sonnet / Gemini / ChatGPT）的 `draft_sha256` 与终稿一致。
5. **所有规范单一真源（SSOT）**：文风改 `脚本与规范/style_guide.md`，配色改 `脚本与规范/chart_theme.json`，Word 模板规则改 `脚本与规范/report_template_spec.md`，底稿门禁改 `脚本与规范/workpaper_spec.md`，SOP 改 `脚本与规范/workflow_sop.md`。
