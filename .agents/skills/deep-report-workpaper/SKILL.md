---
name: deep-report-workpaper
description: 深度报告 Step 5B 合规工作底稿跨平台自动整句红框截图（四级梯队穷尽取证 + PDF防歧义扩框 + Chrome/Chromium CDP网页红框 + .meta.json验真）与门禁 2B/六道门禁官方 Excel 模板打包。当用户说“做工作底稿”“生成底稿”“底稿截图”“底稿合规检查”或“/deep-report-workpaper”时触发。
---

# 深度报告合规工作底稿截图与打包（跨平台通用版）

本 Skill 负责把 `<公司名>/06_证据清单.json` 中的全部数据落地为整句/整行红框高清原件截图，并通过门禁 2B 与六道合规门禁打包进官方底稿 Excel 模板。**注意：工作底稿 Excel 是合规证据底稿，严禁与盈利预测计算模型混写成一张普通表。**

## 规范与模板真源（相对路径）

- 底稿四级梯队、门禁 2B 与六道合规门禁：`脚本与规范/workpaper_spec.md`
- 内置官方底稿 Excel 模板：`模板资产/国金研报工作底稿模板.xlsx`（脚本自动定位）

## 取证与豁免铁律

1. **合规金融数据库常规财务免截图**：来自 **iFinD / Choice / Wind / Bloomberg** 的常规财务、估值与行情数据，在 `06_证据清单.json` 中标明 `is_database_exempt: true`（或来源含数据库名），列示导出原数即可免贴原件截图。
2. **非标准业务细节与行业数据必须按四级梯队穷尽取证并贴高清整句红框原件图**：
   - 严格遵循 `workpaper_spec.md` 四级梯队：**① 官方 PDF 原件**（年报/季报/招股书/20-F/海外 IR 财报 PDF） → **② 合规金融数据库（iFinD/Choice/Wind）** → **③ 权威机构官网原稿**（SEMI/TrendForce/IDC/Omdia/中汽协等） → **④ 白名单财经/专业电子媒体对独立第三方报告的明确转引报道**（财联社、证券时报、上证报、中证报、第一财经、华尔街见闻、界面、澎湃、DIGITIMES、经济日报、集微网、电子发烧友/EEWorld，须过 `must_contain` 逐数核验）。
   - PDF 原件：使用 `auto_evidence_capture.py` 的 `find_verified_pdf_page` + `capture_pdf_with_redbox(..., require_redbox=True, expand_to_sentence=True)`，防歧义找页并自动扩框至完整句子或整行表格，落盘 `.meta.json`。
   - 公开网页：使用 Chrome / Chromium / Edge CDP 模式（内置通用 `div` DOM 标红回退、加载轮询与断连重试）截屏并写 `.meta.json`（无红框自动拒截；若海外官网遇 WAF/403，必须自动降级至第 ④ 级白名单媒体转引页再试）。
   - **严禁**：无红框空白图、裸根域名网址、多文档/多机构只截单张图、引用友商卖方研报或非白名单论坛/自媒体。
3. **防“未抓先弃”硬门禁（门禁 2B）**：严禁对可公开获取的官网新闻稿、海外财报 PDF 或白名单媒体转引数据直接写 `missing_reason` 跳过抓取；仅当穷尽四级梯队均真实尝试失败（由脚本 `--write-back` 自动回写 `attempted_urls` 留痕）或确为闭源付费报告无任何白名单公开转引时，才允许留空标黄（未配 `capture_spec` 且无 `attempted_urls` 将在 `--strict` 下直接 🔴 阻断）。

## 标准执行命令（在根目录运行）

### 1. 一键批量截取全部 PDF/网页红框证据图（失败自动回写 `attempted_urls`）
```bash
python 脚本与规范/auto_evidence_capture.py \
  --manifest <公司名>/06_证据清单.json \
  --out-dir <公司名>/08_底稿证据 --write-back
```
（自动跳过已存在含 `.meta.json` 的截图；加 `--overwrite` 强制重截。）

### 2. 跑六道合规门禁（含门禁 2B）审计与打包高清工作底稿 Excel
```bash
# 默认交付打包（0 🔴 违规时若仍有已尝试留痕的闭源付费报告缺图，诚实打印 [INCOMPLETE]）：
python 脚本与规范/build_workpaper.py \
  --manifest <公司名>/06_证据清单.json \
  --out <公司名>/<公司名>_工作底稿.xlsx --strict

# 可选终验（零容忍：存在任意 1 张待补截图或 1 条裸根域名即 exit 1 阻断）：
python 脚本与规范/build_workpaper.py \
  --manifest <公司名>/06_证据清单.json --audit-only --release
```
