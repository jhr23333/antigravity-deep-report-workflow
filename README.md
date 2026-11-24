# 国金证券电子组——公司深度研究报告半自动化工作台

> 用大白话说明这套工作台怎么运转、每一步跑什么脚本，以及想改文风/配色/模板/门禁时直接去改哪个文件（单一真源 SSOT 设计，不重复抄写规则）。

---

## 一、全链路 6 步运行原理（从立项到 Word + 合规底稿交付）

1. **Step 1 立项建档与财务取数**：新建 `D:\深度报告\<公司名>\`，**常规财务与同业对比优先走同花顺 iFinD（免下 PDF、底稿免截图）**；必须无过滤拉取近 90 天全部临时公告清单防漏最新经营/产能/再融资公告，股本与实控人权益以最新发行/股本变动公告 PDF 为唯一真源；后台按需自动拉取核心公开 PDF 并用 `pull_company_data.py` 分合并报表精确抽取快照（自动算三费率与广义现金，不输出财务费用率）。同时用知识星球 MCP 摸底友商逻辑（物理隔离禁入底稿）。
2. **Step 2 段落级大纲 + 两阶段精准联网与一手 URL 锁定**：边搜边出精确到 `(a)/(b)/(c)` 的段落级大纲与配图方案；采用**两阶段强制联网检索**（先宽搜定框架，再对每个核心数字/协议规格/近 60 天行业更新单点深搜核验），并在动笔前把每张核心行业图的**真实文章页 URL 或海外官方 IR PDF 链接 + `must_contain` 核验词**落盘至 `02_行业公开数据包.md`（若海外官网遇 WAF 403 则同步锁定国内白名单财经/电子媒体转引页），你拍板后冻结。
3. **Step 3 Subagent 分节写正文**：按二级章节派发 Subagent 写初稿（只喂规范 + 轻量锚点卡 + 本节资料包；严守**正文不连续堆砌数字、不分析财务费用率**红线），汇总生成 `05_完整初稿.md`、`charts.json`、`06_证据清单.json` 与 `report_meta.json`。
4. **Step 3.5 Antigravity 内置 Claude Opus 5.5 独立联网审稿（必过门禁）**：先跑 `redteam_review.py` 生成带 `draft_sha256` 版本指纹的精简任务包（含宣传词/财务费用率/数字堆砌本地预扫描），再在 Antigravity 内派发零历史 **`Claude Opus 5.5`** worker **直接通读正文与摘要并调用 `search_web` 联网核查**（六大必查维度：近 90 天公告与股本、会计与常识勾稽、协议/工艺规格、白名单、文风与年份对齐、推理克制）；坚持卖方偏乐观护航立场，改完才进下一步。
5. **Step 4 初稿复核（保留 `[待模型填入]`）**：你审定修订后的正文逻辑与事实；**AI 不写、不填任何盈利预测与目标价**，全部保留 `[待模型填入]`，待产出 Word 后由你自行刷模型并配合国金插件填入。
6. **Step 5 一键对账、出图、渲染 Word、生成零遗漏底稿（区分草稿 `--strict` 与正式发布 `--release` 双门禁）**：
   - `chart_maker.py` 生成 300 DPI 国金黑蓝灰图表；
   - `sync_check.py` 执行【期间×数值】有序对账、**图表标题年份与横轴一致性校验**、正文 Markdown 表格逐行对账、文风红线扫描与审稿版本指纹核验（`--release` 拦未填占位与未绑定复审）；
   - `render_docx.py` + `verify_docx.py`（12 项 XML + 摘要字数 + 占位符检验，异常直接非零退出）产出原生国金 Word 报告；
   - `auto_evidence_capture.py`（含通用 `div` DOM 标红、加载轮询、CDP 重试与失败 `attempted_urls` 自动留痕）+ `build_workpaper.py`（`--strict` 含**门禁 2B 防未抓先弃**拦死一切偷懒留空，`--release` 零容忍拦截一切待补截图与裸根域名）产出高清红框工作底稿 Excel。

---

## 二、目录结构一览（单一真源 SSOT）

```text
D:\深度报告\
├─ AGENTS.md                            # Antigravity / Codex 工作区总入口规则（指向 .agents/skills/）
├─ README.md                            # 本总览说明（运行原理 + 命令 + 修改入口）
├─ .agents\skills\                      # 打包好的 4 个模块化工作流 Skill（单步或端到端均可直接触发）
│  ├─ deep-report-run\SKILL.md          # 总控：Step 1–5 端到端完整深度报告工作流
│  ├─ deep-report-review\SKILL.md       # Step 3.5：Antigravity Opus 5.5 红队审稿与终稿哈希签认
│  ├─ deep-report-charts\SKILL.md       # Step 5A：300 DPI 制图、三端【期间×数值】对账、Word 渲染与 12 项体检
│  └─ deep-report-workpaper\SKILL.md    # Step 5B：自动红框截图、.meta.json 验真与六道门禁底稿 Excel 打包
├─ 模板资产\                             # 国金官方只读模板
│  ├─ 国金证券_公司深度研究模板.docx      # Word 官方模板（保留首页右栏财务表与文末三表位）
│  └─ 国金研报工作底稿模板.xlsx          # Excel 官方模板（出正式稿自动去模板两页）
├─ 脚本与规范\                           # 核心底座（5 份规范/档案 + 1 份配色 + 8 个自动化脚本）
│  ├─ style_guide.md                    # 【SSOT】质控红线 + 12 条正文细则 + 1 年内时效 + 卖方语体
│  ├─ report_template_spec.md           # 【SSOT】Markdown -> Word XML 样式映射（OneFullText/双换页/表全居中）
│  ├─ workpaper_spec.md                 # 【SSOT】工作底稿规范 + 四级来源梯队 + 六道程序化门禁
│  ├─ workflow_sop.md                   # 【SSOT】6 步协作流程 + 锚点卡/大纲/Subagent 模板
│  ├─ chart_theme.json                  # 【SSOT】国金黑蓝灰无网格线配色与字号配置
│  ├─ quality_fix_log.md                # 历史 27 项实战排错复盘档案（参考用，不重复喂给 Agent）
│  ├─ pull_company_data.py              # Step 1：一键抽取最新季报/年报三大表全科目快照 + iFinD 行情
│  ├─ redteam_review.py                 # Step 3.5：预扫描 + 导出 Antigravity Opus 5.5 精简审稿任务包
│  ├─ chart_maker.py                    # Step 5：批量画图（9 种图表：柱/线/双轴/饼/时间轴/股权/客户矩阵）
│  ├─ sync_check.py                     # Step 5：正文/画图/底稿三端一致性自动对账 + 清理孤儿旧图
│  ├─ render_docx.py                    # Step 5：Markdown 渲染进国金 Word 模板（含极端/空泛/宣传词检查）
│  ├─ verify_docx.py                    # Step 5：Word 成品 12 项 XML/占位符体检 + Word COM 真实分页视觉抽检（--visual）
│  ├─ auto_evidence_capture.py          # Step 5：防歧义 PDF 找页 + Chrome CDP DOM 红框截图 + .meta.json 落盘
│  └─ build_workpaper.py                # Step 5：正式高清底稿生成 + 六道质控门禁（--strict 零遗漏阻断）
├─ 参考样本\                             # 历年报告样本、质控部注意事项 PPT、团队优秀报告 PDF
└─ <公司名>\                             # 每家新公司独立子文件夹（00_原始资料 / 01–02 共享锚点 / 03_段落级大纲★ / 04_分节草稿 / 05_完整初稿★及审稿 / 06_证据清单 / 07_图表 / 08_底稿证据★ / 09_对账与质检报告★及视觉抽检页 / 最终 Word★与底稿 Excel★）
```

---

## 三、标准一键命令（以 `D:\深度报告\示例公司\` 为例）

```powershell
# Step 1：一键抽取最新一期季报/年报全科目快照（含归母净利/广义现金/所得税/在建工程，自动剔除财务费用率）
D:\anaconda3\python.exe D:\深度报告\脚本与规范\pull_company_data.py `
  --name 示例公司 --code 600000.SH `
  --pdf D:\深度报告\示例公司\00_原始资料\pdfs\最新季报.pdf `
  --out D:\深度报告\示例公司\00_原始资料\01_最新财报全科目快照.md

# Step 3.5：生成 Antigravity Opus 5.5 精简审稿任务包（直接扫正文 + 联网查新，不翻大 PDF）
D:\anaconda3\python.exe D:\深度报告\脚本与规范\redteam_review.py `
  --md D:\深度报告\示例公司\05_完整初稿.md `
  --meta D:\深度报告\示例公司\report_meta.json `
  --snapshot D:\深度报告\示例公司\00_原始资料\01_最新财报全科目快照.md `
  --manifest D:\深度报告\示例公司\06_证据清单.json `
  --out D:\深度报告\示例公司\05_红队审稿任务包.md

# Step 5-1：批量生成 300 DPI 图表 + 三端【期间×数值】与正文表格对账并清废弃旧图
D:\anaconda3\python.exe D:\深度报告\脚本与规范\chart_maker.py --spec D:\深度报告\示例公司\charts.json --out-dir D:\深度报告\示例公司\07_图表
D:\anaconda3\python.exe D:\深度报告\脚本与规范\sync_check.py --dir D:\深度报告\示例公司 --clean-orphans --strict

# Step 5-2：渲染国金官方 Word 模板 + 12 项 XML/摘要字数/占位符自动体检
D:\anaconda3\python.exe D:\深度报告\脚本与规范\render_docx.py --md D:\深度报告\示例公司\05_完整初稿.md --meta D:\深度报告\示例公司\report_meta.json --out D:\深度报告\示例公司\示例公司_深度报告.docx
D:\anaconda3\python.exe D:\深度报告\脚本与规范\verify_docx.py --docx D:\深度报告\示例公司\示例公司_深度报告.docx

# Step 5-3：一键批量截取全部红框原件图 + 生成高清正式工作底稿 Excel（严格六道门禁 + 机器元数据验真）
D:\anaconda3\python.exe D:\深度报告\脚本与规范\auto_evidence_capture.py --manifest D:\深度报告\示例公司\06_证据清单.json --out-dir D:\深度报告\示例公司\08_底稿证据 --write-back
D:\anaconda3\python.exe D:\深度报告\脚本与规范\build_workpaper.py --manifest D:\深度报告\示例公司\06_证据清单.json --out D:\深度报告\示例公司\示例公司_工作底稿.xlsx --strict

# 【终稿发版前必跑】三端 --release 零容忍门禁（补齐模型数字、付费报告截图并完成终稿审稿签认后执行）：
D:\anaconda3\python.exe D:\深度报告\脚本与规范\sync_check.py --dir D:\深度报告\示例公司 --release
D:\anaconda3\python.exe D:\深度报告\脚本与规范\verify_docx.py --docx D:\深度报告\示例公司\示例公司_深度报告.docx --release --visual
D:\anaconda3\python.exe D:\深度报告\脚本与规范\build_workpaper.py --manifest D:\深度报告\示例公司\06_证据清单.json --audit-only --release
```

> **为什么区分 `--strict`（草稿门禁）与 `--release`（发布门禁）？**
> - **草稿阶段（`--strict`）**：拦死一切真实错误（黑名单、门禁 2B 未抓先弃、无红框/假红框、期间-数值互换、图表标题年份与横轴不符、正文表格错位、Word XML 结构缺陷），每次对账自动生成 `09_对账与质检报告.md` 列出全部 🔴/🟡 与行号，同时允许研究员尚未填完第四章 `[待模型填入]` 或确已尝试但仍缺原件的闭源付费报告截图（须有 `attempted_urls` 留痕，打印 `[DRAFT-PASS]` / `[INCOMPLETE]`）；
> - **正式发布（`--release`）**：零容忍模式——只要存在任意 1 处 `[待模型填入]`、任意 1 张待补截图、任意 1 条裸根域名网址，或审稿报告（`05_红队审稿意见.md` / `05_红队审稿意见_Claude.md`）的 `draft_sha256` 未绑定到当前终稿版本，一律非零退出码（exit 1）硬阻断。

---

## 四、修改入口速查（改什么去哪找）

| 想调整的内容 | 唯一修改入口 | 说明 |
|---|---|---|
| 正文文风、质控红线、字数范围、示范段 | `脚本与规范\style_guide.md` | 12 条正文细则与合规红线单一真源 |
| 图表颜色、字体、无网格线、画布尺寸 | `脚本与规范\chart_theme.json` | 纯 JSON 配置，严守黑蓝灰色系 |
| Word 样式映射、单图缩进宽度、表居中 | `脚本与规范\report_template_spec.md` / `render_docx.py` | 顶部极端词/空泛词/宣传词库可直接增删 |
| 流程步骤、大纲格式、Opus 审稿定位 | `脚本与规范\workflow_sop.md` / `redteam_review.py` | 调整任务包模板或宣传词预扫描库 |
| 底稿黑名单（含友商研报）、六道门禁 | `脚本与规范\workpaper_spec.md` / `build_workpaper.py` | 顶部 `BLACKLIST_SOURCES` 与 `audit_manifest()` |
| 截图防歧义词表与红框线宽 | `脚本与规范\auto_evidence_capture.py` | 顶部 `DEFAULT_AMBIGUOUS_CONTEXTS` |
| 替换新版国金官方 Word / Excel 模板 | `模板资产\` 下同名文件覆盖 | 覆盖后脚本自动读取新版模板 |
