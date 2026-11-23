---
name: deep-report-workpaper
description: 深度报告 Step 5B 合规工作底稿自动红框截图（PDF防歧义定位 + 网页Chrome CDP红框 + .meta.json机器溯源）与六道门禁 Excel 打包。当用户说“做工作底稿”“生成底稿”“底稿截图”“底稿合规检查”或“/deep-report-workpaper”时触发。
---

# 深度报告合规工作底稿截图与打包

本 Skill 负责把 `06_证据清单.json` 中的全部外部/公司数据落地为可过质控的红框高清原件截图，并通过六道合规门禁打包为国金证券官方工作底稿 Excel。

## 规范真源

- 底稿六道合规门禁与黑名单规则：`.\脚本与规范\workpaper_spec.md`
- 官方底稿 Excel 模板：`.\模板资产\国金研报工作底稿模板.xlsx`

## 取证与豁免铁律

1. **常规财务/行情走 iFinD（免截图）**：公司历史营收/净利/毛利率/股价/估值等常规数据库指标，在 `06_证据清单.json` 中标明 `is_database_exempt: true`（来源写 `同花顺iFinD`），免贴原件截图。
2. **非标准业务细节与行业数据必须贴高清红框原件图**：
   - PDF 原件（年报/季报/招股书/行业白皮书）：必须用 `auto_evidence_capture.py` 的 `find_verified_pdf_page` + `capture_pdf_evidence(..., require_redbox=True)` 截取，内置 `DEFAULT_AMBIGUOUS_CONTEXTS` 防错页，自动生成 `<图片>.meta.json` 溯源元数据（记录 `box_count`、`page_text_norm`、`verified_tokens`）。
   - 公开网页（TrendForce/IDC/中汽协/公司官网新闻稿等）：必须用 `auto_evidence_capture.py` 的 Chrome CDP 模式在 DOM 标红后整段截屏并写 `.meta.json`。
   - **严禁**：无红框空白图、裸根域名网址（如 `https://www.prismark.com`）、多文档/多机构只截单张图、引用友商卖方研报（高盛/大摩/中信/中金/华泰等）或非白名单论坛/自媒体。
3. **无公开原件的付费报告（如 Prismark 完整版）**：诚实留空 `evidence_screenshot` 并在 `missing_reason` 注明，严禁拿无关网页滥竽充数。

## 标准执行命令

### 1. 仅跑六道合规门禁审计（快速排雷）
```powershell
# 草稿预览模式（0 🔴 违规时若仍有待补截图，打印 [INCOMPLETE] 而非 PASS）：
python .\脚本与规范\build_workpaper.py `
  --manifest .\{公司名}\06_证据清单.json --audit-only --strict

# 正式发布模式（零容忍：存在任意 1 张待补截图或 1 条裸根域名网址即 exit 1 阻断）：
python .\脚本与规范\build_workpaper.py `
  --manifest .\{公司名}\06_证据清单.json --audit-only --release
```

### 2. 正式生成高清工作底稿 Excel
```powershell
python .\脚本与规范\build_workpaper.py `
  --manifest .\{公司名}\06_证据清单.json `
  --out .\{公司名}\{公司名}_工作底稿.xlsx --strict
```
（正式发版前将 `--strict` 替换为 `--release`，确保 0 待补截图、0 裸根域名、0 黑名单、0 无红框图。）

## 门禁 4 防造假说明

- 当截图旁存在 `.meta.json` 时，`build_workpaper.py` **完全忽略**清单中人工自填的 `covered_claims`，只信任机器提取的 `page_text_norm + verified_tokens` 且 `box_count != 0`；
- 历史遗留无 `.meta.json` 的截图，须用 `auto_evidence_capture.py` 重截，或在人工目视核对红框后显式标注 `"visually_verified": true`，否则固定触发 🟡 告警。
