---
name: deep-report-charts
description: 深度报告 Step 5A 制图、防缩水与三端【期间×数值】有序对账、Markdown表格核验、官方模板 Word 渲染与 12 项 XML + 跨平台真实分页体检。当用户说“画报告图表”“三端对账”“生成深度报告Word”“渲染docx”或“/deep-report-charts”时触发。
---

# 深度报告制图、三端对账与 Word 渲染验收（跨平台通用版）

本 Skill 负责把结构化数据与初稿渲染为 300 DPI 标准图表与国金官方模板 `.docx` 报告，并用确定性脚本拦死图表缩水、三端漂移与排版缺陷。**严禁绕过 `render_docx.py` 手搓无模板 Word。**

## 规范与模板真源（相对路径）

- 跨平台图表配色（兼容 Win/Mac/Linux 中文字体）：`脚本与规范/chart_theme.json`
- Word 样式映射与单图缩进规范：`脚本与规范/report_template_spec.md`
- 内置官方 Word 模板：`模板资产/国金证券_公司深度研究模板.docx`（脚本自动定位，无需传绝对路径）

## 标准执行流水线（在仓库根目录运行）

### 1. 批量生成 300 DPI 高清图表（含饼图自动防重叠与像素框重叠检测）
```bash
python 脚本与规范/chart_maker.py \
  --spec <公司名>/charts.json \
  --out-dir <公司名>/07_图表
```

### 2. 三端【期间 × 数值】有序对账 + 防缩水拦截 + 清理孤儿旧图
```bash
# 默认交付门禁（自动输出 <公司名>/09_对账与质检报告.md）：
python 脚本与规范/sync_check.py --dir <公司名> --clean-orphans --strict

# 可选零占位终验（回填模型且终稿签认完毕后）：
python 脚本与规范/sync_check.py --dir <公司名> --release
```
**对账覆盖范围**：
- 防缩水门禁：成稿图表总数低于 `03_段落级大纲.md` 拟定数 75% 直接 🔴 拦截；
- 正文 `05_完整初稿.md` ↔ `charts.json` ↔ `07_图表/` 一一对应；
- `charts.json` ↔ `06_证据清单.json` 按 `(期间/分类, 数值)` 有序配对（年份互换必抓）；
- `05_完整初稿.md` 每张 `TABLE_META` 表格 ↔ `06_证据清单.json` `type="table"` 逐行核对；
- 正文纯文字带单位小数锚定检查 + 高敏定性狠话底稿覆盖扫描（均带行号写入 `09_对账与质检报告.md`）；
- 首页摘要字数护栏（`≤ 850 字`）与审稿报告 `draft_sha256` 校验。

### 3. 套用官方模板渲染 Word（内置 `keepNext / cantSplit / tblHeader` 防跨页断裂）
```bash
python 脚本与规范/render_docx.py \
  --md <公司名>/05_完整初稿.md \
  --meta <公司名>/report_meta.json \
  --out <公司名>/<公司名>_深度报告.docx
```

### 4. Word 12 项 XML 体检 + 跨平台真实分页视觉抽检
```bash
# 默认交付体检（Windows 走 Word COM、Linux/Mac 自动走 LibreOffice 导出首页抽检 PNG）：
python 脚本与规范/verify_docx.py --docx <公司名>/<公司名>_深度报告.docx --visual

# 可选零占位终验：
python 脚本与规范/verify_docx.py --docx <公司名>/<公司名>_深度报告.docx --release --visual
```

## 异常处置原则

- 任何脚本返回非零退出码时，先修源数据（`charts.json` / `05_完整初稿.md` / `06_证据清单.json`）再重跑全链，严禁手动改 Word 后不回写源文件。
