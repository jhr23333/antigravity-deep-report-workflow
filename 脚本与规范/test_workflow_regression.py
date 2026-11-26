# -*- coding: utf-8 -*-
"""
国金证券公司深度研究报告工作流 —— 端到端离线回归测试脚本 (test_workflow_regression.py)

运行原理（大白话说明）：
本脚本不依赖外部网络或付费终端，在系统临时目录中自动构造“正常样本”与“刻意埋雷的脏样本”，
把本仓库全部核心脚本、内置国金 Word/Excel 官方模板以及关键规范契约一口气端到端跑一遍，
专门用来防止后续改代码或跨会话更新规范时出现“悄悄回退”或“上下游脚本格式打架”。

覆盖的 16 组回归测试：
1. 静态零硬编码：防止本机个人盘符/路径或历史个股名误入通用仓库；
2. 规范防回退契约：校验入口、Skill 与 5 份 SSOT 规范（含 `industry_sources.md`）关键红线未被误删或漂移；
3. 脚本编译与 CLI 入口：校验全部 9 个生产脚本可正常编译且相对路径 `--help` 退出码为 0；
4. `chart_maker.py` 实出图：校验柱线双轴图与小扇区环形图按内置 `chart_theme.json` 正常出图；
5. `sync_check.py` 脏样本全拦截：校验防缩水(75%)、图表标题年份错位、年份数值互换、财务费用率、连续堆砌数字与 list 格式 `07_图表/charts.json` 回退；
6. `sync_check.py` 干净样本与纵向分类防误报：校验表头“产品”与分类“核心产品A”子串不误报，且 `--release` 准确拦截 `[待模型填入]`；
7. `build_workpaper.py` 六道门禁与实建底稿：校验门禁 2B（未抓先弃）、小写黑名单、Choice/`is_database_exempt` 豁免及套内置国金 Excel 模板生成 `.xlsx`；
8. `redteam_review.py` 来源豁免与预扫描：校验 Choice/自定义豁免不漏入非数据库表，且预扫描捕获财务费用率、堆砌数字、宣传词、AI 味词及解禁/质押排查提醒；
9. `render_docx.py` 8 项文风体检与 `*斜体*` 清除：校验纯名词标题、超长标题、口语词、破折号、去 AI 味、加粗频次、段首加粗与斜体剥离（保留 `**加粗**`），且干净样本 0 误报；
10. `verify_docx.py` 动态多章与发布门禁：校验内置国金 Word 模板渲染 6 章报告、动态拦截第四章连续纯文字段及 `--release` 拦截未填占位符；
11. `fetch_public_data.py` 离线解析与合规声明：校验市场识别、SEC submissions/XBRL 离线解析及免截图红线声明；
12. `auto_evidence_capture.py` P0 契约与写回：校验 `mode/pdf/out` 别名兼容、防双前缀目录、`--write-back` 保留手工字段与清陈旧多图、PDF 失败写 `attempted_urls`；
13. `chart_maker.py` P1 修复：校验双轴非百分比右轴不误加 `%`、负值分组柱状图底部留白与显式 `show_labels=True` 生效；
14. `pull_company_data.py` P1 修复：校验真实 PDF 跳过目录页并正确匹配 `一、营业收入`；
15. `sync_check.py` P1 修复与防照搬招股书：校验含月标题年份区间、`2025` 不误配 `2025H1`、乱序表特征对齐、纯数值列整数错位拦截与招股书集中度 >50% 预警；
16. P2 门禁闭环：校验 `build_workpaper.py` 8 列宽表续表/分号多 URL 裸域名/扩充黑名单/Bloomberg 过滤，`render_docx.py --strict` 拦摘要禁词与三级标题拆块，`verify_docx.py` 拦首页标题/换页符破坏。

命令行用法（在仓库根目录运行）：
    python 脚本与规范/test_workflow_regression.py
"""

import json
import os
import py_compile
import re
import subprocess
import sys
import tempfile
import unittest

from docx import Document


# 自动定位仓库根目录（本脚本位于 <根目录>/脚本与规范/）
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)


def run_repo_script(rel_script: str, args: list) -> subprocess.CompletedProcess:
    """以仓库根目录为工作目录，使用当前 Python 解释器执行指定生产脚本。"""
    return subprocess.run(
        [sys.executable, rel_script] + args,
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


class DeepReportWorkflowRegressionTests(unittest.TestCase):
    """深度报告工作流全链路离线回归测试集。"""

    # 纳入回归守卫的核心文本文件清单（不含本测试文件自身，避免正则字面量自触发）
    CORE_TEXT_FILES = [
        "AGENTS.md",
        "README.md",
        ".agents/skills/deep-report-run/SKILL.md",
        ".agents/skills/deep-report-review/SKILL.md",
        ".agents/skills/deep-report-charts/SKILL.md",
        ".agents/skills/deep-report-workpaper/SKILL.md",
        "脚本与规范/style_guide.md",
        "脚本与规范/workflow_sop.md",
        "脚本与规范/workpaper_spec.md",
        "脚本与规范/report_template_spec.md",
        "脚本与规范/industry_sources.md",
        "脚本与规范/chart_theme.json",
        "脚本与规范/fetch_public_data.py",
        "脚本与规范/pull_company_data.py",
        "脚本与规范/auto_evidence_capture.py",
        "脚本与规范/chart_maker.py",
        "脚本与规范/build_workpaper.py",
        "脚本与规范/redteam_review.py",
        "脚本与规范/render_docx.py",
        "脚本与规范/sync_check.py",
        "脚本与规范/verify_docx.py",
    ]

    PROD_SCRIPTS = [f for f in CORE_TEXT_FILES if f.endswith(".py")]

    def test_01_no_hardcoded_local_paths(self) -> None:
        """检查全部 21 份核心文件无本机盘符、个人目录或遗留个股名硬编码。"""
        # 用拼接方式构造禁止模式，防止本文件在全局搜索时产生误命中
        forbidden_tokens = [
            "D:" + "\\",
            "ana" + "conda3",
            "53" + "271",
            "超颖" + "电子",
        ]
        bad_re = re.compile("|".join(re.escape(t) for t in forbidden_tokens), re.I)
        hits = []
        for rel in self.CORE_TEXT_FILES:
            fp = os.path.join(REPO_ROOT, rel)
            self.assertTrue(os.path.exists(fp), f"核心文件缺失：{rel}")
            with open(fp, "r", encoding="utf-8") as f:
                for ln, line in enumerate(f, 1):
                    if bad_re.search(line):
                        hits.append((rel, ln, line.strip()[:100]))
        self.assertEqual(hits, [], f"发现硬编码路径或历史个股名残留：{hits}")

    def test_02_ssot_key_rules_anti_regression(self) -> None:
        """校验入口与 5 份 SSOT 规范的关键红线与字段契约未被意外覆写丢失。"""
        checks = {
            "AGENTS.md": ["近 90 天", "财务费用率", "门禁 2B", "75%", "[待模型填入]", "Choice", "fetch_public_data.py", "industry_sources.md", "50%"],
            ".agents/skills/deep-report-run/SKILL.md": ["近 90 天", "两阶段", "财务费用率", "75%", "[待模型填入]", "fetch_public_data.py", "industry_sources.md", "50%"],
            "脚本与规范/style_guide.md": ["不可逆历史里程碑", "双分句", "压舱石", "财务费用率"],
            "脚本与规范/workflow_sop.md": ["防提示词注入", "Choice", "04_分节草稿", "fetch_public_data.py", "industry_sources.md"],
            "脚本与规范/workpaper_spec.md": ["is_database_exempt", "evidence_screenshots", "capture_spec", "attempted_urls", "required_claims", "pdf_path"],
            "脚本与规范/report_template_spec.md": ["GJStylesTableDefalutStyle", "[待模型填入]"],
            "脚本与规范/industry_sources.md": ["Euromonitor", "SEMI", "50%", "百川"],
        }
        for rel, required_phrases in checks.items():
            with open(os.path.join(REPO_ROOT, rel), "r", encoding="utf-8") as f:
                txt = f.read()
            for phrase in required_phrases:
                self.assertIn(phrase, txt, f"{rel} 缺失关键规范契约词：{phrase}")

    def test_03_all_scripts_compile_and_cli_help(self) -> None:
        """校验全部 9 个生产脚本语法编译通过且 --help 可用。"""
        for rel in self.PROD_SCRIPTS:
            py_compile.compile(os.path.join(REPO_ROOT, rel), doraise=True)
            res = run_repo_script(rel, ["--help"])
            self.assertEqual(res.returncode, 0, f"{rel} --help 失败：{res.stderr}")

    def test_04_chart_maker_bar_line_and_small_slice_donut(self) -> None:
        """校验 chart_maker.py 使用内置色板实出柱线双轴图与小扇区环形图。"""
        with tempfile.TemporaryDirectory() as td:
            charts_json = os.path.join(td, "charts.json")
            spec = [
                {
                    "filename": "fig01_rev.png",
                    "type": "bar_line_dual",
                    "categories": ["2022", "2023", "2024", "2025"],
                    "series": [
                        {"name": "营业收入", "values": [10.0, 12.5, 15.0, 18.2], "kind": "bar"},
                        {"name": "同比增速", "values": [12.0, 25.0, 20.0, 21.3], "kind": "line"},
                    ],
                    "y_label": "亿元",
                    "y2_label": "%",
                },
                {
                    "filename": "fig02_mix.png",
                    "type": "donut",
                    "categories": ["核心产品A", "核心产品B", "其他"],
                    "series": [{"name": "占比", "values": [65.0, 30.5, 4.5]}],
                },
            ]
            with open(charts_json, "w", encoding="utf-8") as f:
                json.dump(spec, f, ensure_ascii=False, indent=2)
            res = run_repo_script("脚本与规范/chart_maker.py", ["--spec", charts_json, "--out-dir", td])
            self.assertEqual(res.returncode, 0, res.stdout + res.stderr)
            self.assertGreater(os.path.getsize(os.path.join(td, "fig01_rev.png")), 5000)
            self.assertGreater(os.path.getsize(os.path.join(td, "fig02_mix.png")), 5000)

    def test_05_and_06_sync_check_dirty_and_clean_paths(self) -> None:
        """校验 sync_check.py 全部红黄门禁、list 格式回退与纵向中文分类防误报修复。"""
        with tempfile.TemporaryDirectory() as td:
            chart_dir = os.path.join(td, "07_图表")
            os.makedirs(chart_dir, exist_ok=True)
            # 故意使用 list 格式并放在 07_图表/ 下，验证回退兼容不崩溃
            charts_json = os.path.join(chart_dir, "charts.json")
            spec = [
                {
                    "filename": "fig01_rev.png",
                    "type": "bar_line_dual",
                    "categories": ["2022", "2023", "2024", "2025"],
                    "series": [
                        {"name": "营业收入", "values": [10.0, 12.5, 15.0, 18.2], "kind": "bar"},
                        {"name": "同比增速", "values": [12.0, 25.0, 20.0, 21.3], "kind": "line"},
                    ],
                    "y_label": "亿元",
                    "y2_label": "%",
                },
                {
                    "filename": "fig02_mix.png",
                    "type": "donut",
                    "categories": ["核心产品A", "核心产品B", "其他"],
                    "series": [{"name": "占比", "values": [65.0, 30.5, 4.5]}],
                },
            ]
            with open(charts_json, "w", encoding="utf-8") as f:
                json.dump(spec, f, ensure_ascii=False, indent=2)
            r_chart = run_repo_script("脚本与规范/chart_maker.py", ["--spec", charts_json, "--out-dir", chart_dir])
            self.assertEqual(r_chart.returncode, 0, r_chart.stdout + r_chart.stderr)

            # 1) 脏样本：同时埋入 5 类问题（防缩水、年份错位、数值互换、财务费用率、连续堆砌数字）
            md_bad = (
                "# 一、公司概览\n## 1.1 历史沿革\n"
                "公司在2022年实现营收10亿元，2023年12.5亿元，2024年15.0亿元，2025年18.2亿元，"
                "增速分别达12%、25%、20%、21.3%，毛利率为28.5%、29.1%，净利率11.2%，财务费用率为-0.8%。\n\n"
                "<!-- CHART\nimage: 07_图表/fig01_rev.png\ntitle: 2021-2025年公司营业收入及增速\n"
                "source: 公司公告、iFinD、国金证券研究所\n-->\n"
            )
            with open(os.path.join(td, "05_完整初稿.md"), "w", encoding="utf-8") as f:
                f.write(md_bad)
            with open(os.path.join(td, "03_段落级大纲.md"), "w", encoding="utf-8") as f:
                f.write("# 大纲\n" + "\n".join(f"- 图表{i}：规划图表{i}" for i in range(1, 18)))
            manifest_bad = {
                "company_name": "测试公司",
                "sections": [{
                    "sheet_name": "1.1",
                    "items": [{
                        "label": "图表1",
                        "type": "chart",
                        "chart_image": "07_图表/fig01_rev.png",
                        "source_name": "iFinD",
                        "is_database_exempt": True,
                        "raw_data": [
                            ["年份", "营业收入(亿元)", "同比增速(%)"],
                            ["2022", 12.5, 12.0],
                            ["2023", 10.0, 25.0],
                            ["2024", 15.0, 20.0],
                            ["2025", 18.2, 21.3],
                        ],
                    }],
                }],
            }
            with open(os.path.join(td, "06_证据清单.json"), "w", encoding="utf-8") as f:
                json.dump(manifest_bad, f, ensure_ascii=False, indent=2)

            r_bad = run_repo_script("脚本与规范/sync_check.py", ["--dir", td, "--strict"])
            out_bad = r_bad.stdout + r_bad.stderr
            self.assertEqual(r_bad.returncode, 1, f"脏样本未被阻断：{out_bad}")
            for kw in ["防缩水拦截", "图表标题年份与横轴不符", "期间-数值错配", "文风红线·财务费用率", "文风红线·疑似连续堆砌数字"]:
                self.assertIn(kw, out_bad, f"脏样本漏报 {kw}：\n{out_bad}")

            # 2) 干净样本：验证表头“产品”与分类“核心产品A”纵向匹配不误报，且 --release 拦截占位符
            with open(os.path.join(td, "03_段落级大纲.md"), "w", encoding="utf-8") as f:
                f.write("# 大纲\n- 图表1：营收\n- 图表2：产品结构\n")
            md_clean = (
                "# 一、营收稳步增长，盈利结构优化\n## 1.1 主营业务放量，产品结构升级\n"
                "公司核心产品需求旺盛，具体年度营收与产品结构参见下图。\n\n"
                "<!-- CHART\nimage: 07_图表/fig01_rev.png\ntitle: 2022-2025年公司营业收入及增速\n"
                "source: Choice、国金证券研究所\n-->\n\n"
                "<!-- CHART\nimage: 07_图表/fig02_mix.png\ntitle: 公司分产品收入结构\n"
                "source: 公司2025年年报、国金证券研究所\n-->\n\n"
                "# 四、盈利预测与投资建议\n预计公司2026年归母净利润为[待模型填入]亿元。\n"
            )
            with open(os.path.join(td, "05_完整初稿.md"), "w", encoding="utf-8") as f:
                f.write(md_clean)
            manifest_clean = {
                "company_name": "测试公司",
                "sections": [{
                    "sheet_name": "1.1",
                    "items": [
                        {
                            "label": "图表1",
                            "type": "chart",
                            "chart_image": "07_图表/fig01_rev.png",
                            "source_name": "Choice",
                            "is_database_exempt": True,
                            "raw_data": [
                                ["年份", "营业收入(亿元)", "同比增速(%)"],
                                ["2022", 10.0, 12.0],
                                ["2023", 12.5, 25.0],
                                ["2024", 15.0, 20.0],
                                ["2025", 18.2, 21.3],
                            ],
                        },
                        {
                            "label": "图表2",
                            "type": "chart",
                            "chart_image": "07_图表/fig02_mix.png",
                            "source_name": "公司2025年年报",
                            "raw_data": [
                                ["产品", "占比(%)"],
                                ["核心产品A", 65.0],
                                ["核心产品B", 30.5],
                                ["其他", 4.5],
                            ],
                        },
                    ],
                }],
            }
            with open(os.path.join(td, "06_证据清单.json"), "w", encoding="utf-8") as f:
                json.dump(manifest_clean, f, ensure_ascii=False, indent=2)

            r_ok = run_repo_script("脚本与规范/sync_check.py", ["--dir", td, "--strict"])
            self.assertEqual(r_ok.returncode, 0, f"干净样本意外失败：{r_ok.stdout + r_ok.stderr}")
            r_rel = run_repo_script("脚本与规范/sync_check.py", ["--dir", td, "--release"])
            self.assertEqual(r_rel.returncode, 1, "正式发布模式未拦截 [待模型填入]")

    def test_07_build_workpaper_gates_and_builtin_template(self) -> None:
        """校验 build_workpaper.py 门禁 2B、小写黑名单、Choice 豁免及内置国金底稿模板生成。"""
        with tempfile.TemporaryDirectory() as td:
            # 1) 门禁 2B：未配置 capture_spec 且无 attempted_urls 直接留空 -> 🔴 阻断
            p_2b = os.path.join(td, "m_2b.json")
            with open(p_2b, "w", encoding="utf-8") as f:
                json.dump(
                    {"company_name": "测试", "sections": [{"sheet_name": "2.1", "items": [
                        {"label": "图表1", "type": "chart", "source_name": "Prismark 2025"}
                    ]}]},
                    f, ensure_ascii=False,
                )
            r_2b = run_repo_script("脚本与规范/build_workpaper.py", ["--manifest", p_2b, "--audit-only", "--strict"])
            self.assertEqual(r_2b.returncode, 1)
            self.assertIn("门禁2B", r_2b.stdout + r_2b.stderr)

            # 2) 小写黑名单拦截（morgan stanley）
            p_bl = os.path.join(td, "m_bl.json")
            with open(p_bl, "w", encoding="utf-8") as f:
                json.dump(
                    {"company_name": "测试", "sections": [{"sheet_name": "2.1", "items": [
                        {"label": "图表1", "type": "chart", "source_name": "morgan stanley 2025", "attempted_urls": ["https://x.com"]}
                    ]}]},
                    f, ensure_ascii=False,
                )
            r_bl = run_repo_script("脚本与规范/build_workpaper.py", ["--manifest", p_bl, "--audit-only", "--strict"])
            self.assertEqual(r_bl.returncode, 1)
            self.assertIn("黑名单", r_bl.stdout + r_bl.stderr)

            # 3) 合规打包：Choice + is_database_exempt + 带 attempted_urls 留痕 + 实图，套内置国金模板输出 xlsx
            spec_p = os.path.join(td, "c.json")
            with open(spec_p, "w", encoding="utf-8") as f:
                json.dump([{"filename": "shot.png", "type": "bar", "categories": ["2025"], "series": [{"name": "A", "values": [10]}]}], f)
            run_repo_script("脚本与规范/chart_maker.py", ["--spec", spec_p, "--out-dir", td])
            shot_p = os.path.join(td, "shot.png")

            p_ok = os.path.join(td, "m_ok.json")
            out_xlsx = os.path.join(td, "测试公司_工作底稿.xlsx")
            with open(p_ok, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "company_name": "测试公司",
                        "sections": [{
                            "sheet_name": "1.1",
                            "items": [
                                {"label": "图表1", "type": "chart", "source_name": "Choice金融终端"},
                                {"label": "图表2", "type": "table", "source_name": "自定义付费库", "is_database_exempt": True},
                                {"label": "图表3", "type": "chart", "source_name": "Prismark 2025", "attempted_urls": ["https://prismark.com/403"], "missing_reason": "付费墙403"},
                                {"label": "图表4", "type": "chart", "source_name": "公司2025年年报", "evidence_screenshots": [{"path": shot_p, "caption": "年报P.12", "visually_verified": True}]},
                            ],
                        }],
                    },
                    f, ensure_ascii=False,
                )
            r_ok = run_repo_script("脚本与规范/build_workpaper.py", ["--manifest", p_ok, "--out", out_xlsx, "--strict"])
            self.assertEqual(r_ok.returncode, 0, r_ok.stdout + r_ok.stderr)
            self.assertIn("[INCOMPLETE]", r_ok.stdout + r_ok.stderr)
            self.assertTrue(os.path.exists(out_xlsx))
            self.assertGreater(os.path.getsize(out_xlsx), 10000)
            # 打开生成的底稿核验 Sheet 结构：保留目标节与主要数据，自动剔除模板说明页
            import openpyxl
            wb = openpyxl.load_workbook(out_xlsx)
            try:
                self.assertIn("1.1", wb.sheetnames)
                self.assertNotIn("发布底稿要求", wb.sheetnames)
                self.assertNotIn("数字未来实验室-ID", wb.sheetnames)
            finally:
                wb.close()

    def test_08_redteam_review_db_exempt_and_prescan(self) -> None:
        """校验 redteam_review.py 数据库豁免不漏出与全部本地预扫描项。"""
        with tempfile.TemporaryDirectory() as td:
            p_snap = os.path.join(td, "snap.md")
            p_meta = os.path.join(td, "meta.json")
            p_man = os.path.join(td, "man.json")
            with open(p_snap, "w", encoding="utf-8") as f:
                f.write("# 快照\n营业收入 18.2亿元")
            with open(p_meta, "w", encoding="utf-8") as f:
                json.dump({"report_title": "测试公司：主标题", "summary_sections": []}, f, ensure_ascii=False)
            with open(p_man, "w", encoding="utf-8") as f:
                json.dump(
                    {"company_name": "测试", "sections": [{"sheet_name": "1.1", "items": [
                        {"label": "图表1", "source_name": "Choice金融终端"},
                        {"label": "图表2", "source_name": "自定义库", "is_database_exempt": True},
                        {"label": "图表3", "source_name": "Prismark 2025"},
                    ]}]},
                    f, ensure_ascii=False,
                )

            p_md1 = os.path.join(td, "md1.md")
            p_out1 = os.path.join(td, "rt1.md")
            with open(p_md1, "w", encoding="utf-8") as f:
                f.write(
                    "# 一、测试标题\n"
                    "公司在2022年10亿元、2023年12.5亿元、2024年15.0亿元、2025年18.2亿元、增速12%、25%、20%、21.3%、毛利率28.5%，"
                    "财务费用率为-0.8%，成为定海神针并重金卡位实现爆发。\n"
                )
            r1 = run_repo_script(
                "脚本与规范/redteam_review.py",
                ["--md", p_md1, "--meta", p_meta, "--snapshot", p_snap, "--manifest", p_man, "--out", p_out1],
            )
            self.assertEqual(r1.returncode, 0, r1.stderr)
            with open(p_out1, "r", encoding="utf-8") as f:
                t1 = f.read()
            self.assertNotIn("Choice金融终端", t1)
            self.assertNotIn("自定义库", t1)
            self.assertIn("Prismark 2025", t1)
            for kw in ["违反文风红线·财务费用率", "疑似连续数据堆砌", "宣传措辞", "去AI味提醒", "质控风险必查提醒"]:
                self.assertIn(kw, t1, f"红队预扫描漏报 {kw}")

            # 正例：包含解禁排查且无违规词时，提示静默
            p_md2 = os.path.join(td, "md2.md")
            p_out2 = os.path.join(td, "rt2.md")
            with open(p_md2, "w", encoding="utf-8") as f:
                f.write("# 五、风险提示\n经核查，公司近3个月及未来6个月无大额限售股解禁计划。\n")
            r2 = run_repo_script(
                "脚本与规范/redteam_review.py",
                ["--md", p_md2, "--meta", p_meta, "--snapshot", p_snap, "--manifest", p_man, "--out", p_out2],
            )
            self.assertEqual(r2.returncode, 0, r2.stderr)
            with open(p_out2, "r", encoding="utf-8") as f:
                t2 = f.read()
            self.assertNotIn("质控风险必查提醒", t2)
            self.assertNotIn("去AI味提醒", t2)

    def test_09_and_10_render_and_verify_docx(self) -> None:
        """校验 render_docx.py 8 项新体检 + 斜体剥离，以及 verify_docx.py 6 章动态检测与发布门禁。"""
        with tempfile.TemporaryDirectory() as td:
            p_meta = os.path.join(td, "meta.json")
            meta = {
                "report_title": "测试公司：核心逻辑清晰",
                "rating_type": "首次覆盖",
                "target_price": "[待模型填入]",
                "summary_sections": [
                    {"heading": "投资逻辑", "paragraphs": ["公司主业稳健增长，产品结构持续升级。"]},
                    {"heading": "盈利预测、估值和评级", "paragraphs": ["预计2026年归母净利为[待模型填入]亿元。"]},
                    {"heading": "风险提示", "paragraphs": ["下游需求不及预期风险。"]},
                ],
            }
            with open(p_meta, "w", encoding="utf-8") as f:
                json.dump(meta, f, ensure_ascii=False)

            # 1) 脏文本触发全部 8 项 render_docx.py 新检查，并核验 *斜体* 剥离且 **加粗** 保留
            p_bad = os.path.join(td, "bad.md")
            out_bad = os.path.join(td, "bad.docx")
            with open(p_bad, "w", encoding="utf-8") as f:
                f.write(
                    "# 一、公司概况\n"
                    "## 1.1 AI把测试推成核心价值环节——这是一段超过二十六个字的二级标题示例测试\n"
                    "**段首误加粗**文本含 *斜体词* 和**第二处加粗**，并含压舱石套词。\n"
                )
            r_bad = run_repo_script("脚本与规范/render_docx.py", ["--md", p_bad, "--meta", p_meta, "--out", out_bad])
            self.assertEqual(r_bad.returncode, 0, r_bad.stderr)
            for kw in ["纯名词标题", "二级标题超长", "标题口语词", "标题标点规范", "去AI味提醒", "正文加粗过多", "正文段首加粗", "禁用斜体"]:
                self.assertIn(kw, r_bad.stdout, f"render_docx 漏报 {kw}")

            doc_bad = Document(out_bad)
            hit_paras = [p for p in doc_bad.paragraphs if "斜体词" in p.text]
            self.assertTrue(hit_paras and "*斜体词*" not in hit_paras[0].text)
            bold_runs = [run.text for run in hit_paras[0].runs if run.bold]
            self.assertIn("段首误加粗", bold_runs)
            self.assertIn("第二处加粗", bold_runs)

            # 2) 6 章干净报告套内置国金模板渲染 + verify_docx 草稿通过 / --release 阻断 / 第4章连续纯文字阻断
            chart_dir = os.path.join(td, "07_图表")
            os.makedirs(chart_dir, exist_ok=True)
            spec_p = os.path.join(chart_dir, "c.json")
            with open(spec_p, "w", encoding="utf-8") as f:
                json.dump([{"filename": "f1.png", "type": "bar", "categories": ["2025"], "series": [{"name": "营收", "values": [18.2]}]}], f)
            run_repo_script("脚本与规范/chart_maker.py", ["--spec", spec_p, "--out-dir", chart_dir])

            md_6ch = (
                "# 一、主业稳健增长，盈利能力提升\n## 1.1 核心产品放量，份额持续提高\n"
                "公司核心产品需求稳健，具体数据见图表。\n\n"
                "<!-- CHART\nimage: 07_图表/f1.png\ntitle: 2025年公司营业收入\nsource: Choice、国金证券研究所\n-->\n\n"
                "# 四、海外产能落地，新业务贡献增量\n## 4.1 泰国基地试产，客户导入加速\n"
                "公司加速海外基地建设，核心进度见下表。\n\n"
                "<!-- TABLE_META\ntitle: 海外基地进度\nsource: 公司公告、国金证券研究所\n-->\n"
                "| 基地 | 进度 |\n|---|---|\n| 泰国一期 | 试产 |\n\n"
                "# 五、盈利预测与投资建议\n## 5.1 盈利预测与估值\n预计2026年归母净利润为[待模型填入]亿元。\n\n"
                "# 六、风险提示\n## 6.1 需求波动风险\n下游需求不及预期风险。\n"
            )
            p_clean = os.path.join(td, "clean.md")
            out_clean = os.path.join(td, "clean.docx")
            with open(p_clean, "w", encoding="utf-8") as f:
                f.write(md_6ch)

            r_ren = run_repo_script("脚本与规范/render_docx.py", ["--md", p_clean, "--meta", p_meta, "--out", out_clean])
            self.assertEqual(r_ren.returncode, 0, r_ren.stderr)
            self.assertNotIn("🟡", r_ren.stdout, f"干净样本出现误报：\n{r_ren.stdout}")
            self.assertNotIn("🔴", r_ren.stdout, f"干净样本出现误报：\n{r_ren.stdout}")

            r_vd = run_repo_script("脚本与规范/verify_docx.py", ["--docx", out_clean])
            self.assertEqual(r_vd.returncode, 0, r_vd.stdout + r_vd.stderr)
            self.assertIn("核心业务与行业章节", r_vd.stdout)

            r_vr = run_repo_script("脚本与规范/verify_docx.py", ["--docx", out_clean, "--release"])
            self.assertEqual(r_vr.returncode, 1, "verify_docx --release 未拦截 [待模型填入]")

            # 在第 4 章注入两段连续纯文字，确认动态核心章节检测生效
            p_c4 = os.path.join(td, "c4.md")
            out_c4 = os.path.join(td, "c4.docx")
            with open(p_c4, "w", encoding="utf-8") as f:
                f.write(md_6ch.replace("公司加速海外基地建设，核心进度见下表。", "第一段纯文字。\n\n第二段连续纯文字。"))
            run_repo_script("脚本与规范/render_docx.py", ["--md", p_c4, "--meta", p_meta, "--out", out_c4])
            r_c4 = run_repo_script("脚本与规范/verify_docx.py", ["--docx", out_c4])
            self.assertEqual(r_c4.returncode, 1)
            self.assertIn("核心业务与行业章节连续纯文字段 1 处", r_c4.stdout)

    def test_11_fetch_public_data_offline_parsers_and_compliance(self) -> None:
        """离线校验 fetch_public_data.py 市场识别、SEC 解析、XBRL 提取与合规声明输出。"""
        if SCRIPT_DIR not in sys.path:
            sys.path.insert(0, SCRIPT_DIR)
        import fetch_public_data as fpd

        self.assertEqual(fpd.detect_market("600584.SH"), "a")
        self.assertEqual(fpd.detect_market("002475"), "a")
        self.assertEqual(fpd.detect_market("0700.HK"), "hk")
        self.assertEqual(fpd.detect_market("ATAT"), "us")
        self.assertEqual(fpd.normalize_a_code("600584.SH"), "600584")
        self.assertEqual(fpd.cninfo_column_for_code("600584"), "sse")
        self.assertEqual(fpd.cninfo_column_for_code("002475"), "szse")
        self.assertEqual(fpd.cninfo_column_for_code("830001"), "third")

        # 验证 SEC submissions 与 XBRL 离线解析
        import datetime as _dt
        today_s = _dt.date.today().isoformat()
        fake_sub = {
            "filings": {
                "recent": {
                    "form": ["6-K", "20-F"],
                    "filingDate": [today_s, "2025-04-20"],
                    "accessionNumber": ["0001193125-26-000001", "0001193125-25-000002"],
                    "primaryDocument": ["d6k.htm", "d20f.htm"],
                    "primaryDocDescription": ["6-K Earnings", "Annual Report"],
                }
            }
        }
        recent_all, major = fpd.parse_sec_submissions(fake_sub, "0001853717", days=90)
        self.assertEqual(len(recent_all), 1)
        self.assertIn("1853717/000119312526000001/d6k.htm", recent_all[0]["url"])
        self.assertEqual(len(major), 2)

        fake_facts = {
            "facts": {
                "us-gaap": {
                    "Revenues": {"units": {"CNY": [{"form": "20-F", "fp": "FY", "fy": 2025, "val": 9790000000}]}},
                    "NetIncomeLoss": {"units": {"CNY": [{"form": "20-F", "fp": "FY", "fy": 2025, "val": 1620000000}]}},
                }
            }
        }
        xbrl_md = fpd.extract_sec_xbrl_highlights(fake_facts)
        self.assertIn("FY2025", xbrl_md)
        self.assertIn("9,790.00", xbrl_md)
        self.assertIn("1,620.00", xbrl_md)

        md = fpd.build_output_markdown("测试公司", "600584.SH", "a", 90, [("一、测试", "内容")])
        self.assertIn("严禁在工作底稿中免截图使用", md)

    def test_12_auto_evidence_capture_p0_contract_and_writeback(self) -> None:
        """实跑校验 auto_evidence_capture.py P0 修复：旧别名兼容、防双前缀、写回保留字段、清旧多图与 PDF 失败留痕。"""
        import fitz

        with tempfile.TemporaryDirectory() as td:
            pdf_dir = os.path.join(td, "00_原始资料", "pdfs")
            out_dir = os.path.join(td, "08_底稿证据")
            os.makedirs(pdf_dir, exist_ok=True)
            pdf_path = os.path.join(pdf_dir, "sample.pdf")

            # 构造 1 页真实 PDF，包含可检索英文指标行
            doc = fitz.open()
            pg = doc.new_page()
            pg.insert_text((72, 100), "Annual Report 2025: Africa Revenue reached 454.4 million USD.", fontsize=12)
            pg.insert_text((72, 130), "Baby diapers market share ranked No.1 at 20.3% in 2024.", fontsize=12)
            doc.save(pdf_path)
            doc.close()

            manifest_path = os.path.join(td, "06_证据清单.json")
            manifest = {
                "company_name": "测试公司",
                "sections": [{
                    "sheet_name": "2.1",
                    "items": [
                        {
                            # 测试 1：旧版别名 mode/pdf/out 且 out 误带 08_底稿证据/ 前缀 + 原先挂有陈旧 evidence_screenshots
                            "label": "图表1：非洲营收",
                            "type": "chart",
                            "source_name": "Annual Report 2025",
                            "evidence_screenshots": [{"path": "08_底稿证据/stale_old.png", "caption": "旧图"}],
                            "capture_spec": {
                                "mode": "pdf",
                                "pdf": "00_原始资料/pdfs/sample.pdf",
                                "page": 1,
                                "keywords": ["454.4", "20.3%"],
                                "banner": "年报P.1 | 营收与市占率",
                                "out": "08_底稿证据/ev01_single.png",
                            },
                        },
                        {
                            # 测试 2：多图模式保留原有 caption / covered_claims / visually_verified
                            "label": "图表2：多图保留",
                            "type": "table",
                            "source_name": "Annual Report 2025",
                            "evidence_screenshots": [
                                {
                                    "path": "08_底稿证据/ev02_a.png",
                                    "caption": "手工精选说明",
                                    "covered_claims": ["454.4", "20.3%"],
                                    "visually_verified": True,
                                }
                            ],
                            "capture_specs": [
                                {
                                    "type": "pdf",
                                    "pdf_path": "00_原始资料/pdfs/sample.pdf",
                                    "page": 1,
                                    "keywords": ["454.4"],
                                    "banner": "默认横幅",
                                    "out_file": "ev02_a.png",
                                },
                                {
                                    "type": "pdf",
                                    "pdf_path": "00_原始资料/pdfs/sample.pdf",
                                    "page": 1,
                                    "keywords": ["20.3%"],
                                    "banner": "第二张默认横幅",
                                    "out_file": "ev02_b.png",
                                },
                            ],
                        },
                        {
                            # 测试 3：PDF 命中失败自动记录 attempted_urls
                            "label": "图表3：失败留痕",
                            "type": "chart",
                            "source_name": "Annual Report 2025",
                            "capture_spec": {
                                "type": "pdf",
                                "pdf_path": "00_原始资料/pdfs/sample.pdf",
                                "page": 99,
                                "keywords": ["NOT_EXIST_TOKEN"],
                                "out_file": "ev03_fail.png",
                            },
                        },
                    ],
                }],
            }
            with open(manifest_path, "w", encoding="utf-8") as f:
                json.dump(manifest, f, ensure_ascii=False, indent=2)

            res = run_repo_script(
                "脚本与规范/auto_evidence_capture.py",
                ["--manifest", manifest_path, "--out-dir", out_dir, "--write-back"],
            )
            # 含 1 项故意触发失败的 PDF 条目，脚本应返回退出码 1 但同时完成其余 3 张截图与失败留痕回写
            self.assertEqual(res.returncode, 1, res.stdout + res.stderr)
            self.assertIn("成功 3 张 | 跳过 0 张 | 失败 1 张", res.stdout)

            # 确认未出现双嵌套 08_底稿证据/08_底稿证据/
            self.assertTrue(os.path.exists(os.path.join(out_dir, "ev01_single.png")))
            self.assertTrue(os.path.exists(os.path.join(out_dir, "ev01_single.png.meta.json")))
            self.assertFalse(os.path.exists(os.path.join(out_dir, "08_底稿证据")))

            with open(manifest_path, "r", encoding="utf-8") as f:
                updated = json.load(f)
            items = updated["sections"][0]["items"]

            # 核验测试 1：单图写回已清除陈旧 evidence_screenshots，路径无重复前缀
            self.assertEqual(items[0]["evidence_screenshot"], "08_底稿证据/ev01_single.png")
            self.assertNotIn("evidence_screenshots", items[0])

            # 核验测试 2：第一张保留手工 caption/covered_claims/visually_verified，第二张自动补全 banner
            shots2 = items[1]["evidence_screenshots"]
            self.assertEqual(len(shots2), 2)
            self.assertEqual(shots2[0]["caption"], "手工精选说明")
            self.assertEqual(shots2[0]["covered_claims"], ["454.4", "20.3%"])
            self.assertTrue(shots2[0]["visually_verified"])
            self.assertEqual(shots2[1]["caption"], "第二张默认横幅")

            # 核验测试 3：PDF 越界失败已自动写入 attempted_urls 留痕
            self.assertTrue(items[2].get("attempted_urls"), f"PDF 失败未写入 attempted_urls：{items[2]}")

    def test_13_chart_maker_non_pct_dual_axis_and_negative_grouped_bar(self) -> None:
        """实跑校验 chart_maker.py P1 修复：非百分比右轴不加 %、负值分组柱底置标签与显式 show_labels 生效。"""
        if SCRIPT_DIR not in sys.path:
            sys.path.insert(0, SCRIPT_DIR)
        import chart_maker as cm
        import matplotlib.pyplot as plt
        import matplotlib.ticker as mtick

        theme = cm.load_theme(os.path.join(SCRIPT_DIR, "chart_theme.json"))
        cm.setup_matplotlib(theme)

        captured = {}
        orig_close = plt.close

        def _spy_close(fig=None):
            if fig is not None:
                captured["axes"] = list(fig.axes)
            return orig_close(fig)

        with tempfile.TemporaryDirectory() as td:
            plt.close = _spy_close
            try:
                # 1) 非百分比右轴（单位“座”）：不得套用 PercentFormatter，数据标签不得含 %
                spec_dual = {
                    "filename": "dual_non_pct.png",
                    "type": "bar_line_dual",
                    "categories": ["2023", "2024", "2025"],
                    "series": [
                        {"name": "资本开支", "values": [10.0, 14.0, 18.0], "kind": "bar"},
                        {"name": "新增工厂数", "values": [4, 6, 8], "kind": "line"},
                    ],
                    "y_label": "亿元",
                    "y2_label": "座",
                    "y2_label_fmt": "{:.0f}",
                }
                cm.render_single_chart(spec_dual, theme, td)
                ax2 = captured["axes"][1]
                self.assertNotIsInstance(ax2.yaxis.get_major_formatter(), mtick.PercentFormatter)
                y2_texts = [t.get_text() for t in ax2.texts]
                self.assertTrue(y2_texts and all("%" not in tx for tx in y2_texts), f"非百分比右轴误含%：{y2_texts}")

                # 2) 4 系列分组柱状图：显式 show_labels=True + 含负值柱子（校验底部留白与 va='top'）
                spec_neg = {
                    "filename": "grouped_neg.png",
                    "type": "grouped_bar",
                    "categories": ["2024", "2025"],
                    "show_labels": True,
                    "series": [
                        {"name": "S1", "values": [-12.0, 15.0]},
                        {"name": "S2", "values": [8.0, -5.0]},
                        {"name": "S3", "values": [6.0, 9.0]},
                        {"name": "S4", "values": [4.0, 7.0]},
                    ],
                    "y_label": "亿元",
                }
                cm.render_single_chart(spec_neg, theme, td)
                ax_g = captured["axes"][0]
                self.assertEqual(len(ax_g.texts), 8, "显式 show_labels=True 未对 4 系列分组柱生效")
                self.assertLess(ax_g.get_ylim()[0], -12.0, "负值分组柱未扩展底部 y 轴留白")
                neg_labels = [t for t in ax_g.texts if t.get_text().startswith("-")]
                self.assertTrue(neg_labels and all(t.get_va() == "top" for t in neg_labels))
            finally:
                plt.close = orig_close

    def test_14_pull_company_data_toc_skip_and_short_revenue_label(self) -> None:
        """实跑校验 pull_company_data.py P1 修复：跳过目录页并准确抽取“一、营业收入”。"""
        if SCRIPT_DIR not in sys.path:
            sys.path.insert(0, SCRIPT_DIR)
        import fitz
        import pull_company_data as pcd

        with tempfile.TemporaryDirectory() as td:
            pdf_path = os.path.join(td, "annual_test.pdf")
            doc = fitz.open()
            # 第 1 页：年报目录页（包含三大合并报表名称与点线页码，应被自动跳过）
            p1 = doc.new_page()
            p1.insert_text(
                (60, 80),
                "目录\n第一节 合并资产负债表 .................. 2\n第二节 合并利润表 ...................... 2\n第三节 合并现金流量表 .................. 2",
                fontname="china-ss",
                fontsize=11,
            )
            # 第 2 页：真实合并利润表（使用“一、营业收入”而非“营业总收入”）
            p2 = doc.new_page()
            lines_p2 = [
                "合并利润表",
                "编制单位：测试制造股份有限公司",
                "单位：元",
                "项目",
                "一、营业收入",
                "1,200,000,000.00",
                "1,000,000,000.00",
                "营业成本",
                "800,000,000.00",
                "700,000,000.00",
                "销售费用",
                "50,000,000.00",
                "45,000,000.00",
                "管理费用",
                "60,000,000.00",
                "55,000,000.00",
                "研发费用",
                "90,000,000.00",
                "80,000,000.00",
                "财务费用",
                "-5,000,000.00",
                "-3,000,000.00",
                "归属于母公司股东的净利润",
                "180,000,000.00",
                "120,000,000.00",
            ]
            y = 50
            for ln in lines_p2:
                p2.insert_text((60, y), ln, fontname="china-ss", fontsize=10)
                y += 20
            doc.save(pdf_path)
            doc.close()

            extracted, _ = pcd.extract_financials_from_pdf(pdf_path)
            self.assertIn("营业总收入", extracted, f"未抽到“一、营业收入”：{extracted}")
            rev1, rev2, pno = extracted["营业总收入"]
            self.assertEqual(pno, 2, "误将第 1 页目录当成合并利润表")
            self.assertAlmostEqual(rev1, 1.2e9, places=1)
            self.assertAlmostEqual(rev2, 1.0e9, places=1)
            self.assertIn("归母净利润", extracted)
            self.assertAlmostEqual(extracted["归母净利润"][0], 1.8e8, places=1)

    def test_15_sync_check_p1_fixes_and_prospectus_concentration(self) -> None:
        """实跑校验 sync_check.py P1 修复与防照搬招股书集中度检测。"""
        if SCRIPT_DIR not in sys.path:
            sys.path.insert(0, SCRIPT_DIR)
        import sync_check as sc

        # 1) 含月标题年份区间校验（2024年4月–2026年6月 vs 横轴 2025年4月–2026年6月）
        errs_yr = sc.audit_chart_title_year_range(
            [("f1.png", "2024年4月–2026年6月区域销量趋势")],
            {"f1.png": {"categories": ["2025年4月", "2025年10月", "2026年6月"]}},
        )
        self.assertEqual(len(errs_yr), 1)
        self.assertIn("图表标题年份与横轴不符", errs_yr[0])

        # 2) 纯年度 2025 不得子串误匹配 2025H1
        errs_pd = sc.audit_chart_ordered_values(
            "f2.png",
            {"categories": ["2025"], "series": [{"name": "营收", "values": [10.0]}]},
            {"label": "图表2", "raw_data": [["期间", "营收"], ["2025H1", 10.0]]},
        )
        self.assertEqual(len(errs_pd), 1)
        self.assertIn("期间列缺失", errs_pd[0])

        # 3) 表格按特征对齐（乱序不误报）+ 纯数值列整数冲突拦截 + 备注小整数不误报
        md_tables = [
            ("产能对比表", [["公司", "工厂数", "备注"], ["甲公司", "8座", "覆盖20余国"]]),
            ("营收对比表", [["公司", "2025营收(亿元)"], ["甲公司", "240"]]),
        ]
        manifest_tables = [
            {"label": "营收对比表", "raw_data": [["公司", "2025营收(亿元)"], ["甲公司", 120]]},
            {"label": "产能对比表", "raw_data": [["公司", "工厂数", "备注"], ["甲公司", 8, "布局广泛"]]},
        ]
        errs_tbl = sc.audit_md_tables_vs_manifest(md_tables, manifest_tables)
        self.assertEqual(len(errs_tbl), 1, f"乱序表对齐或整数冲突校验异常：{errs_tbl}")
        self.assertIn("240.0", errs_tbl[0])

        # 4) 第二/三章招股书集中度 > 50% 预警，补充独立源至 ≤ 50% 后解除
        man_over = {
            "sections": [{
                "sheet_name": "2.1",
                "items": [
                    {"type": "chart", "source_name": "公司招股书（援引弗若斯特沙利文）"},
                    {"type": "chart", "source_name": "公司招股章程"},
                    {"type": "table", "source_name": "公司招股书"},
                    {"type": "chart", "source_name": "Euromonitor 欧睿国际"},
                ],
            }]
        }
        w_over = sc.audit_prospectus_concentration(man_over)
        self.assertEqual(len(w_over), 1)
        self.assertIn("过度依赖招股书", w_over[0])

        man_balanced = {
            "sections": [{
                "sheet_name": "2.1",
                "items": [
                    {"type": "chart", "source_name": "公司招股书"},
                    {"type": "chart", "source_name": "公司招股书"},
                    {"type": "table", "source_name": "World Bank 世界银行数据库"},
                    {"type": "chart", "source_name": "Euromonitor 欧睿国际"},
                ],
            }]
        }
        self.assertEqual(sc.audit_prospectus_concentration(man_balanced), [])

    def test_16_p2_workpaper_wide_table_bare_url_and_docx_strict_gates(self) -> None:
        """实跑校验 P2 修复：8 列宽表续表、分号多 URL 裸域名、扩充黑名单、Bloomberg 过滤、render_docx --strict 与 verify_docx 硬门禁。"""
        if SCRIPT_DIR not in sys.path:
            sys.path.insert(0, SCRIPT_DIR)
        import build_workpaper as bw
        import openpyxl

        # 1) 分号多 URL 第二条裸根域名检出 & Bloomberg 多源过滤
        self.assertTrue(bw._is_bare_homepage_url("https://example.com/report/123; https://bare-home.com"))
        self.assertFalse(bw._is_bare_homepage_url("https://example.com/report/123; https://ok.com/doc/456"))
        self.assertEqual(bw._split_distinct_sources("公司2025年报、Bloomberg、国金证券研究所"), ["公司2025年报"])

        with tempfile.TemporaryDirectory() as td:
            # 2) 扩充黑名单拦截（Goldman Sachs / 微信公众号）
            m_bl = {
                "company_name": "测试",
                "sections": [{
                    "sheet_name": "2.1",
                    "items": [
                        {"label": "图1", "type": "chart", "source_name": "Goldman Sachs 研报", "attempted_urls": ["https://x.com/1"]},
                        {"label": "图2", "type": "chart", "source_name": "某行业微信公众号", "attempted_urls": ["https://x.com/2"]},
                    ],
                }],
            }
            rep_bl = bw.audit_manifest(m_bl, td)
            self.assertGreaterEqual(len(rep_bl["red_errors"]), 2)

            # 3) 8 列宽表自动续表（第 8 列“2026E”完整落盘，不被静默丢弃）
            m_wide_p = os.path.join(td, "wide.json")
            out_xlsx = os.path.join(td, "wide.xlsx")
            m_wide = {
                "company_name": "测试",
                "sections": [{
                    "sheet_name": "1.2",
                    "items": [{
                        "label": "表1：8列历史财务",
                        "type": "table",
                        "source_name": "Choice",
                        "raw_data": [
                            ["科目", "2020", "2021", "2022", "2023", "2024", "2025", "2026E"],
                            ["营收", 10, 12, 14, 16, 18, 20, 999],
                        ],
                    }],
                }],
            }
            with open(m_wide_p, "w", encoding="utf-8") as f:
                json.dump(m_wide, f, ensure_ascii=False)
            r_wide = run_repo_script("脚本与规范/build_workpaper.py", ["--manifest", m_wide_p, "--out", out_xlsx, "--strict"])
            self.assertEqual(r_wide.returncode, 0, r_wide.stdout + r_wide.stderr)
            wb = openpyxl.load_workbook(out_xlsx)
            try:
                ws = wb["1.2"]
                all_vals = [cell.value for row in ws.iter_rows() for cell in row if cell.value is not None]
                self.assertIn("2026E", all_vals, f"宽表第 8 列表头被截断：{all_vals}")
                self.assertIn(999, all_vals, f"宽表第 8 列数据被截断：{all_vals}")
            finally:
                wb.close()

            # 4) render_docx --strict 拦截封面摘要禁词 & 三级标题独立拆块不粘连
            p_md = os.path.join(td, "m.md")
            p_meta = os.path.join(td, "meta.json")
            out_docx = os.path.join(td, "out.docx")
            with open(p_md, "w", encoding="utf-8") as f:
                f.write("# 一、主业稳健，份额领先\n## 1.1 需求增长，盈利改善\n### 独立三级引导\n正文紧随其后。\n")
            with open(p_meta, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "report_title": "测试：逻辑清晰",
                        "summary_sections": [{"heading": "投资逻辑", "paragraphs": ["公司处于遥遥领先地位。"]}],
                    },
                    f,
                    ensure_ascii=False,
                )
            r_strict = run_repo_script("脚本与规范/render_docx.py", ["--md", p_md, "--meta", p_meta, "--out", out_docx, "--strict"])
            self.assertEqual(r_strict.returncode, 1, f"render_docx --strict 未拦截摘要极端禁词：{r_strict.stdout}")
            self.assertIn("禁用极端词", r_strict.stdout)

            # 改为干净摘要后渲染，校验三级标题未与正文粘连，再人为破坏首页标题字体校验 verify_docx 阻断
            with open(p_meta, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "report_title": "测试：逻辑清晰",
                        "summary_sections": [{"heading": "投资逻辑", "paragraphs": ["公司经营稳健，现金流良好。"]}],
                    },
                    f,
                    ensure_ascii=False,
                )
            r_ok = run_repo_script("脚本与规范/render_docx.py", ["--md", p_md, "--meta", p_meta, "--out", out_docx, "--strict"])
            self.assertEqual(r_ok.returncode, 0, r_ok.stdout + r_ok.stderr)
            self.assertIn("三级标题提醒", r_ok.stdout)
            d_chk = Document(out_docx)
            self.assertTrue(any(p.text.strip() == "独立三级引导" for p in d_chk.paragraphs), "三级标题被误粘连进正文")

            # 人为破坏首页主标题字体为 Arial，校验 verify_docx 第 1 项硬门禁阻断退出码=1
            d_chk.tables[0].cell(0, 0).paragraphs[0].runs[0].font.name = "Arial"
            corrupt_docx = os.path.join(td, "corrupt.docx")
            d_chk.save(corrupt_docx)
            r_corrupt = run_repo_script("脚本与规范/verify_docx.py", ["--docx", corrupt_docx])
            self.assertEqual(r_corrupt.returncode, 1, f"verify_docx 未阻断异常首页标题字体：{r_corrupt.stdout}")
            self.assertIn("首页主标题样式不符", r_corrupt.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
