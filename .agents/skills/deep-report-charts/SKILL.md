---
name: deep-report-charts
description: 深度报告 Step 5A 制图、三端【期间×数值】有序对账、Markdown表格核验、Word 渲染与 12 项 XML 体检。支持 --strict 草稿门禁与 --release 正式发布门禁。当用户说“画报告图表”“三端对账”“生成深度报告Word”“渲染docx”“发版验收”或“/deep-report-charts”时触发。
---

# 深度报告制图、三端对账与 Word 渲染验收

本 Skill 负责把结构化数据与初稿渲染为符合国金证券官方视觉规范的 300 DPI 图表与 `.docx` 报告，并用确定性脚本拦死三端数据漂移与排版缺陷。

## 规范真源

- 图表视觉配色（黑蓝灰、无网格线、300 DPI）：`.\脚本与规范\chart_theme.json`
- Word 样式映射与单图缩进规范：`.\脚本与规范\report_template_spec.md`
- 官方 Word 模板：`.\模板资产\国金证券_公司深度研究模板.docx`

## 标准执行流水线

### 1. 批量生成 300 DPI 高清图表
```powershell
python .\脚本与规范\chart_maker.py `
  --spec .\{公司名}\charts.json `
  --out-dir .\{公司名}\07_图表
```

### 2. 三端【期间 × 数值】有序对账 + 清理孤儿旧图
```powershell
# 草稿阶段（允许 [待模型填入] 与未复审黄灯，拦死年份互换/表格错位/缺图）：
python .\脚本与规范\sync_check.py `
  --dir .\{公司名} --clean-orphans --strict

# 正式发布（零容忍：占位符必须为 0、审稿 draft_sha256 必须匹配最新终稿）：
python .\脚本与规范\sync_check.py `
  --dir .\{公司名} --release
```
**对账覆盖范围**：
- 正文 `05_完整初稿.md` ↔ `charts.json` ↔ `07_图表/` 图片完整一一对应；
- `charts.json` ↔ `06_证据清单.json` 按 **`(期间/分类, 数值)` 有序配对**（哪怕仅互换两年数值也立即 🔴 报错）；
- `05_完整初稿.md` 每张 `TABLE_META` 表格 ↔ `06_证据清单.json` `type="table"` 逐行数值核对；
- 首页摘要字数护栏（`≤ 850 字`）与 `05_红队审稿意见_Claude.md` 指纹绑定检查。

### 3. 渲染国金官方 Word 报告
```powershell
python .\脚本与规范\render_docx.py `
  --md .\{公司名}\05_完整初稿.md `
  --meta .\{公司名}\report_meta.json `
  --out .\{公司名}\{公司名}_深度报告.docx
```

### 4. Word 12 项 XML 与占位符体检
```powershell
# 草稿阶段（11 项 XML/表格居中/摘要字数全过即 exit 0，第 12 项仅标黄提示剩余 [待模型填入]）：
python .\脚本与规范\verify_docx.py `
  --docx .\{公司名}\{公司名}_深度报告.docx

# 正式发布（12 项全检，存在任意 1 处 [待模型填入]/TBD/TODO 即 exit 1 阻断）：
python .\脚本与规范\verify_docx.py `
  --docx .\{公司名}\{公司名}_深度报告.docx --release
```

## 异常处置原则

- 任何脚本返回非零退出码时，**先修源数据（`charts.json` / `05_完整初稿.md` / `06_证据清单.json`）再重跑全链**，严禁手动改 Word 后不回写 Markdown/JSON。
- 每次修改图表后，务必带 `--clean-orphans` 重跑 `sync_check.py`，防止废弃旧图残留在 `07_图表\`。
