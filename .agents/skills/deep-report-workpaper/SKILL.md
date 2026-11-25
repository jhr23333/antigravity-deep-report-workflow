---
name: deep-report-workpaper
description: 深度报告 Step 5B 合规工作底稿跨平台自动整句红框截图（PDF防歧义扩框 + Chrome/Chromium CDP网页红框 + .meta.json机器验真）与六道门禁官方 Excel 模板打包。当用户说“做工作底稿”“生成底稿”“底稿截图”“底稿合规检查”或“/deep-report-workpaper”时触发。
---

# 深度报告合规工作底稿截图与打包（跨平台通用版）

本 Skill 负责把 `<公司名>/06_证据清单.json` 中的全部数据落地为整句/整行红框高清原件截图，并通过六道合规门禁打包进官方底稿 Excel 模板。**注意：工作底稿 Excel 是合规证据底稿，严禁与盈利预测计算模型混写成一张普通表。**

## 规范与模板真源（相对路径）

- 底稿六道合规门禁与黑名单规则：`脚本与规范/workpaper_spec.md`
- 内置官方底稿 Excel 模板：`模板资产/国金研报工作底稿模板.xlsx`（脚本自动定位）

## 取证与豁免铁律

1. **合规金融数据库常规财务免截图**：来自 **iFinD / Choice / Wind / Bloomberg** 的常规财务、估值与行情数据，在 `06_证据清单.json` 中标明 `is_database_exempt: true`（或来源含数据库名），列示导出原数即可免贴原件截图。
2. **非标准业务细节与行业数据必须贴高清整句/整行红框原件图**：
   - PDF 原件（年报/季报/招股书/20-F/行业白皮书）：使用 `auto_evidence_capture.py` 的 `find_verified_pdf_page` + `capture_pdf_with_redbox(..., require_redbox=True, expand_to_sentence=True)`，防歧义找页并自动将红框扩展至完整句子或整行表格，落盘 `.meta.json`。
   - 公开网页：使用 Chrome / Chromium / Edge CDP 模式在 DOM 标红后截屏并写 `.meta.json`（无红框自动拒截）。
   - **严禁**：无红框空白图、裸根域名网址、多文档/多机构只截单张图、引用友商卖方研报或非白名单论坛/自媒体。
3. **无公开原件的付费报告**：诚实留空 `evidence_screenshot` 并在 `missing_reason` 注明，严禁凑图。

## 标准执行命令（在仓库根目录运行）

### 1. 一键批量截取全部 PDF/网页红框证据图（跨平台免写临时脚本）
```bash
python 脚本与规范/auto_evidence_capture.py \
  --manifest <公司名>/06_证据清单.json \
  --out-dir <公司名>/08_底稿证据 --write-back
```
（自动跳过已存在含 `.meta.json` 的截图；加 `--overwrite` 强制重截。）

### 2. 跑六道合规门禁审计与打包高清工作底稿 Excel
```bash
# 默认交付打包（0 🔴 违规时若仍有待补付费报告截图，诚实打印 [INCOMPLETE]）：
python 脚本与规范/build_workpaper.py \
  --manifest <公司名>/06_证据清单.json \
  --out <公司名>/<公司名>_工作底稿.xlsx --strict

# 可选终验（零容忍：存在任意 1 张待补截图或 1 条裸根域名即 exit 1 阻断）：
python 脚本与规范/build_workpaper.py \
  --manifest <公司名>/06_证据清单.json --audit-only --release
```
