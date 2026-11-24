---
name: deep-report-workpaper
description: 深度报告 Step 5B 合规工作底稿自动红框截图（PDF防歧义定位 + 网页Chrome CDP红框 + .meta.json机器溯源）与六道门禁 Excel 打包。当用户说“做工作底稿”“生成底稿”“底稿截图”“底稿合规检查”或“/deep-report-workpaper”时触发。
---

# 深度报告合规工作底稿截图与打包

本 Skill 负责把 `06_证据清单.json` 中的全部外部/公司数据落地为可过质控的红框高清原件截图，并通过六道合规门禁打包为国金证券官方工作底稿 Excel。

## 规范真源

- 底稿六道合规门禁与黑名单规则：`D:\深度报告\脚本与规范\workpaper_spec.md`
- 官方底稿 Excel 模板：`D:\深度报告\模板资产\国金研报工作底稿模板.xlsx`

## 取证与豁免铁律

1. **常规财务/行情走 iFinD（免截图）**：公司历史营收/净利/毛利率/股价/估值等常规数据库指标，在 `06_证据清单.json` 中标明 `is_database_exempt: true`（来源写 `同花顺iFinD`），免贴原件截图。
2. **非标准业务细节与行业数据必须按四级梯队穷尽取证并贴高清红框原件图**：
   - 严格遵循 `workpaper_spec.md` 四级取证梯队：**① 官方 PDF 原件**（公司年报/季报/招股书/海外上市公司 IR 财报 PDF） → **② 合规金融数据库（iFinD/Wind）** → **③ 权威机构官网原稿**（SEMI/TrendForce/IDC/Omdia/中汽协等） → **④ 白名单财经/专业电子媒体对独立第三方报告的明确转引报道**（财联社、证券时报、上证报、中证报、第一财经、华尔街见闻、界面、澎湃、DIGITIMES、经济日报、集微网、电子发烧友/EEWorld，须通过 `must_contain` 逐数核验）。
   - PDF 原件：必须用 `auto_evidence_capture.py` 的 `find_verified_pdf_page` + `capture_pdf_with_redbox(..., require_redbox=True, expand_to_sentence=True)` 截取，自动将红框扩展至完整句子/表格整行并落盘 `<图片>.meta.json`。
   - 公开网页：必须用 `auto_evidence_capture.py` 的 Chrome CDP 模式在 DOM 标红后整段截屏并写 `.meta.json`（若海外官网遇 WAF/403，必须自动降级至第 ④ 级白名单媒体转引页再试，严禁不抓直接留空）。
   - **严禁**：无红框空白图、裸根域名网址（如 `https://www.prismark.com`）、多文档/多机构只截单张图、引用友商卖方研报（高盛/大摩/中信/中金/华泰等）或非白名单论坛/自媒体。
3. **防“未抓先弃”硬门禁（门禁 2B）**：严禁对可公开获取的官网新闻稿、海外财报 PDF 或白名单媒体转引数据直接写 `missing_reason` 跳过抓取；仅当穷尽四级梯队均真实尝试失败（由脚本自动回写 `attempted_urls` 留痕）或确为闭源付费报告无任何白名单公开转引时，才允许诚实留空 `evidence_screenshot` 标黄（未配置 `capture_spec` 且无 `attempted_urls` 将在 `--strict` 下直接 🔴 阻断）。

## 标准执行命令

### 0. 一键批量截取全部 PDF/网页红框证据图（免写公司临时脚本）
在 `06_证据清单.json` 各非豁免条目中写好 `capture_spec`（或 `capture_specs`）后，直接运行：
```powershell
D:\anaconda3\python.exe D:\深度报告\脚本与规范\auto_evidence_capture.py `
  --manifest D:\深度报告\{公司名}\06_证据清单.json `
  --out-dir D:\深度报告\{公司名}\08_底稿证据 --write-back
```
（自动跳过已存在含 `.meta.json` 的截图；加 `--overwrite` 可强制重截。）

### 1. 仅跑六道合规门禁审计（快速排雷）
```powershell
# 草稿预览模式（0 🔴 违规时若仍有待补截图，打印 [INCOMPLETE] 而非 PASS）：
D:\anaconda3\python.exe D:\深度报告\脚本与规范\build_workpaper.py `
  --manifest D:\深度报告\{公司名}\06_证据清单.json --audit-only --strict

# 正式发布模式（零容忍：存在任意 1 张待补截图或 1 条裸根域名网址即 exit 1 阻断）：
D:\anaconda3\python.exe D:\深度报告\脚本与规范\build_workpaper.py `
  --manifest D:\深度报告\{公司名}\06_证据清单.json --audit-only --release
```

### 2. 正式生成高清工作底稿 Excel
```powershell
D:\anaconda3\python.exe D:\深度报告\脚本与规范\build_workpaper.py `
  --manifest D:\深度报告\{公司名}\06_证据清单.json `
  --out D:\深度报告\{公司名}\{公司名}_工作底稿.xlsx --strict
```
（正式发版前将 `--strict` 替换为 `--release`，确保 0 待补截图、0 裸根域名、0 黑名单、0 无红框图。）

## 门禁 4 防造假说明

- 当截图旁存在 `.meta.json` 时，`build_workpaper.py` **完全忽略**清单中人工自填的 `covered_claims`，只信任机器提取的 `page_text_norm + verified_tokens` 且 `box_count != 0`；
- 历史遗留无 `.meta.json` 的截图，须用 `auto_evidence_capture.py` 重截，或在人工目视核对红框后显式标注 `"visually_verified": true`，否则固定触发 🟡 告警。
