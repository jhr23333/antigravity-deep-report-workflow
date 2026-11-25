# 卖方公司深度研究报告——跨平台半自动化工作台

> **开箱即用**：兼容 **Windows / macOS / Linux（含 ChatGPT Work 云端沙盒）**，支持 **Antigravity、Codex、ChatGPT Work**。
> 已内置官方 Word / Excel 模板、8 个自动化 Python 脚本、4 个模块化 Skill、4 份单一真源规范与 1 份跨平台配色配置。

---

## 一、快速上手（30 秒避坑指南）

1. **安装 Python 依赖**（在仓库/工作区根目录运行一次即可）：
   ```bash
   pip install -r requirements.txt
   ```
2. **务必开一个全新对话启动**（避免旧对话的旧版预测、旧图表污染上下文），并把本目录作为当前工作目录。
3. **直接触发总控 Skill**：
   > `按 .agents/skills/deep-report-run/SKILL.md 的完整流程，为 <公司名>（<股票代码>）撰写公司深度研究报告，财务数据源优先用 <iFinD / Choice / Wind / 公开财报>。`
4. **核心避坑**：
   - 大纲阶段（Step 2）会停下来等你确认章节与图表清单，拍板后再动笔；
   - 即使在没有多 Agent 并发的单对话框客户端（如 ChatGPT Work），也必须逐节写入 `04_分节草稿/` 并实跑 `chart_maker.py` / `sync_check.py` / `render_docx.py` / `build_workpaper.py`，**绝不允许单回合压缩成 8–9 页简报**；
   - 盈利预测、PE、目标价与评级默认保留 `[待模型填入]`，拿到最终 Word 后由你自行刷模型填入。

---

## 二、全链路 6 步运行原理

1. **Step 1 立项建档与全量取数**：新建 `<公司名>/` 目录，常规财务优先走你当前可用的金融终端（iFinD / Choice / Wind / SEC 20-F / 官方财报，底稿免截图）；**必须无过滤拉取近 90 天全部临时公告标题清单**防漏重大催化，股本与实控人权益以最新发行/权益变动公告 PDF 为唯一真源；A 股可按需跑 `pull_company_data.py` 抽取三大表快照（含广义现金/所得税/少数股东损益，不输出财务费用率）。如本机配了知识星球 MCP，可同步摸底市场逻辑（物理隔离禁入底稿）。
2. **Step 2 段落级大纲 + 两阶段精准联网与一手 URL 锁定**：采用**两阶段强制联网检索**（先宽搜定框架 → 再对每个核心数字/协议规格/近 60 天行业更新单点深搜核验），产出精确到 `(a)/(b)/(c)` 的 **`03_段落级大纲.md`**（章节数灵活，每个二级节不少于 3–4 段），并在 `02_行业公开数据包.md` 落盘真实文章页 URL 或海外官方 IR PDF 链接 + `must_contain` 核验词（海外官网遇 WAF/403 时同步锁定国内白名单财经/电子媒体转引页），**停下来等你确认后冻结**。
3. **Step 3 逐节撰稿与统稿汇总**：有子 Agent 则并发，无子 Agent 则按二级节逐个落盘 `04_分节草稿/*.md`（每段 200–300 字、不超过 350 字、1 段配 1 图/表；严守**正文不连续堆砌数字、不分析财务费用率**双红线），合并生成 `05_完整初稿.md`、`charts.json`、`06_证据清单.json` 与 `report_meta.json`。
4. **Step 3.5 独立联网红队审稿（不可跳过）**：跑 `redteam_review.py` 生成带 `draft_sha256` 指纹与本地预扫描线索的 `05_红队审稿任务包.md`，再由独立审稿员（Antigravity 首选内置 Opus 5.5，可切 Sonnet/Gemini；ChatGPT Work/Codex 执行独立联网红队自审）对照六大维度核查并落盘 `05_红队审稿意见.md`，定点修订后再渲染。
5. **Step 4 初稿复核（保留 `[待模型填入]`）**：你审定修订后的正文逻辑与事实；盈利预测与估值保持 `[待模型填入]` 留待最终 Word 中自刷模型。
6. **Step 5 一键制图、防缩水对账、官方模板 Word 渲染、整句红框截图与底稿 Excel**：
   - `chart_maker.py`：300 DPI 黑蓝灰图表 + 饼图防重叠 + 像素框重叠告警；
   - `sync_check.py`：【期间×数值】有序对账 + **图表标题年份与横轴一致性硬校验（🔴）** + Markdown 表格逐行核验 + 正文千分位小数/高敏定性句/文风双红线扫描 + **防缩水门禁（成稿图数 < 大纲 75% 直接 🔴 拦截）** + 自动落盘 `09_对账与质检报告.md`；
   - `render_docx.py` + `verify_docx.py`：套用 `模板资产/国金证券_公司深度研究模板.docx`（内置 `keepNext / cantSplit / tblHeader`）+ 12 项 XML 体检 + 跨平台真实分页视觉抽检（`--visual`，Windows 走 Word COM、Linux/Mac 自动走 LibreOffice）；
   - `auto_evidence_capture.py`（跨平台 PDF 整句/整行红框 + Chrome/Chromium CDP 通用 `div` DOM 标红、加载轮询、重试与失败 `attempted_urls` 自动留痕）+ `build_workpaper.py`（`--strict` 含**门禁 2B 防未抓先弃**拦死偷懒留空，`--release` 零容忍）打包 `模板资产/国金研报工作底稿模板.xlsx`。

---

## 三、仓库目录结构（相对路径开箱即用）

```text
<工作区/仓库根目录>/
├─ AGENTS.md                            # 工作区总入口规则（含防缩水、双红线与跨平台铁律）
├─ README.md                            # 本总览说明与标准命令
├─ requirements.txt                     # Python 依赖清单
├─ .agents/skills/                      # 4 个模块化工作流 Skill（单步或端到端均可直接触发）
│  ├─ deep-report-run/SKILL.md          # 总控：Step 1–5 端到端完整深度报告工作流
│  ├─ deep-report-review/SKILL.md       # Step 3.5：多客户端自适应六维红队审稿与哈希签认
│  ├─ deep-report-charts/SKILL.md       # Step 5A：300 DPI 制图、防缩水与三端对账、Word 渲染与体检
│  └─ deep-report-workpaper/SKILL.md    # Step 5B：一键批量整句红框截图、门禁 2B 与六道门禁底稿打包
├─ 模板资产/                             # 官方只读模板（脚本自动按相对路径读取）
│  ├─ 国金证券_公司深度研究模板.docx
│  └─ 国金研报工作底稿模板.xlsx
└─ 脚本与规范/                           # 4 份单一真源规范 + 1 份跨平台配色 + 8 个自动化脚本
   ├─ style_guide.md                    # 【SSOT】质控红线 + 12 条正文细则 + 永久双红线 + 卖方语体
   ├─ report_template_spec.md           # 【SSOT】Markdown -> Word XML 样式映射
   ├─ workpaper_spec.md                 # 【SSOT】工作底稿规范 + 四级梯队 + 门禁 2B + 六道门禁
   ├─ workflow_sop.md                   # 【SSOT】6 步协作流程 + 00–09 目录速查 + 共享锚点模板
   ├─ chart_theme.json                  # 【SSOT】国金黑蓝灰无网格线配色（兼容 Win/Mac/Linux 中文字体）
   ├─ pull_company_data.py              # Step 1（A股可选）：合并三大表精确快照（含广义现金，剔除财务费用率）
   ├─ redteam_review.py                 # Step 3.5：文风/口径预扫描 + 导出带 draft_sha256 的审稿任务包
   ├─ chart_maker.py                    # Step 5：9 种标准图表批量生成 + 饼图防重叠 + 像素重叠检测
   ├─ sync_check.py                     # Step 5：防缩水 + 标题年份 + 有序对账 + 落盘 09_对账与质检报告.md
   ├─ render_docx.py                    # Step 5：渲染进官方 Word 模板（防跨页断裂）
   ├─ verify_docx.py                    # Step 5：12 项 XML 体检 + 跨平台 PDF 真实分页抽检（--visual）
   ├─ auto_evidence_capture.py          # Step 5：一键 --manifest 批量 PDF 整句红框 + 网页 CDP 红框 + 留痕
   └─ build_workpaper.py                # Step 5：正式高清底稿打包 + 门禁 2B 与六道门禁（--strict / --release）
```

---

## 四、标准一键命令（在根目录直接运行，无需改盘符）

```bash
# Step 1（A股可选）：从最新季报/年报 PDF 精确提取合并三大表快照
python 脚本与规范/pull_company_data.py \
  --name 示例公司 --code 600000.SH \
  --pdf 示例公司/00_原始资料/pdfs/最新季报.pdf \
  --out 示例公司/00_原始资料/01_最新财报全科目快照.md

# Step 3.5：生成带稿件指纹与本地预扫描线索的红队审稿任务包
python 脚本与规范/redteam_review.py \
  --md 示例公司/05_完整初稿.md \
  --meta 示例公司/report_meta.json \
  --snapshot 示例公司/00_原始资料/01_最新财报全科目快照.md \
  --manifest 示例公司/06_证据清单.json \
  --out 示例公司/05_红队审稿任务包.md

# Step 5-1：批量画图 + 防缩水/标题年份/三端有序对账（同时输出 示例公司/09_对账与质检报告.md）
python 脚本与规范/chart_maker.py --spec 示例公司/charts.json --out-dir 示例公司/07_图表
python 脚本与规范/sync_check.py --dir 示例公司 --clean-orphans --strict

# Step 5-2：套用官方模板渲染 Word + 12 项 XML 与跨平台真实分页视觉抽检
python 脚本与规范/render_docx.py --md 示例公司/05_完整初稿.md --meta 示例公司/report_meta.json --out 示例公司/示例公司_深度报告.docx
python 脚本与规范/verify_docx.py --docx 示例公司/示例公司_深度报告.docx --visual

# Step 5-3：一键批量整句红框截图（失败自动记录 attempted_urls）+ 打包合规高清工作底稿 Excel
python 脚本与规范/auto_evidence_capture.py --manifest 示例公司/06_证据清单.json --out-dir 示例公司/08_底稿证据 --write-back
python 脚本与规范/build_workpaper.py --manifest 示例公司/06_证据清单.json --out 示例公司/示例公司_工作底稿.xlsx --strict

# 【可选终验】回填模型并补齐付费报告截图后，跑 --release 零容忍门禁：
python 脚本与规范/sync_check.py --dir 示例公司 --release
python 脚本与规范/verify_docx.py --docx 示例公司/示例公司_深度报告.docx --release --visual
python 脚本与规范/build_workpaper.py --manifest 示例公司/06_证据清单.json --audit-only --release
```
