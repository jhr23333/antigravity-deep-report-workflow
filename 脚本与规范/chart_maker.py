# -*- coding: utf-8 -*-
"""
国金证券电子组深度报告 —— 标准图表批量生成脚本 (chart_maker.py)

核心功能：
1. 读取同目录下的 chart_theme.json，严格遵循国金质控部要求：
   仅使用黑、蓝、灰色系，严禁出现红色、桔色、绿色，字号与研报正文协调。
2. 根据图表配置 JSON 文件，批量生成 300 DPI 高清 PNG 图片，支持：
   - bar           : 单系列柱状图
   - grouped_bar   : 多系列分组柱状图
   - stacked_bar   : 多系列堆叠柱状图
   - line          : 单/多系列折线图
   - bar_line_dual : 左轴柱状图 + 右轴折线图（白底胶囊标签防深蓝柱遮挡）
   - pie / donut   : 饼图 / 环形图（如收入结构、市场份额）
   - timeline      : 发展历程横向时间轴（国金深蓝卡片风格）
   - org_chart     : 股权穿透与制造基地架构图
   - customer_matrix: 多赛道核心客户与应用场景矩阵图
3. 支持 single（通栏单图）与 double（双图并排）两种画布尺寸，避免插入 Word 后字号变形。

命令行用法：
    python 脚本与规范/chart_maker.py --spec charts.json --out-dir 07_图表
"""

import argparse
import json
import os
import sys
from typing import Any, Dict, List

import matplotlib
# 使用非交互式 Agg 后端，确保在命令行/后台无图形界面环境下稳定出图
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mtick
import numpy as np


def load_theme(theme_path: str) -> Dict[str, Any]:
    """加载国金标准色板与样式配置 JSON。"""
    with open(theme_path, "r", encoding="utf-8") as f:
        return json.load(f)


def setup_matplotlib(theme: Dict[str, Any]) -> None:
    """配置全局中文字体、负号显示与基础字号。"""
    font_cfg = theme["font"]
    # 按优先级注册中文字体，避免中文乱码或方块
    plt.rcParams["font.sans-serif"] = font_cfg["family_priority"]
    # 解决坐标轴负号 '-' 显示为方块的问题
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["axes.edgecolor"] = theme["colors"]["spine"]
    plt.rcParams["axes.linewidth"] = 0.8
    plt.rcParams["xtick.color"] = font_cfg["text_color"]
    plt.rcParams["ytick.color"] = font_cfg["text_color"]
    plt.rcParams["text.color"] = font_cfg["text_color"]


def format_value(val: float, fmt: str = "") -> str:
    """根据格式化模板输出数据标签，默认自动去掉多余的 .0。"""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    if fmt:
        return fmt.format(val)
    if abs(val - round(val)) < 1e-6:
        return f"{int(round(val))}"
    return f"{val:.1f}"


def style_axes(ax: plt.Axes, theme: Dict[str, Any], show_grid: bool = False) -> None:
    """统一美化坐标轴：隐藏上/右边框；默认不放网格线，保持国金研报极简风格。"""
    layout = theme["layout"]
    colors = theme["colors"]
    font_cfg = theme["font"]

    ax.set_facecolor(colors["background"])
    ax.spines["top"].set_visible(layout.get("show_top_spine", False))
    ax.spines["right"].set_visible(layout.get("show_right_spine", False))
    ax.spines["left"].set_color(colors["spine"])
    ax.spines["bottom"].set_color(colors["spine"])
    ax.tick_params(axis="both", labelsize=font_cfg["tick_size"], length=3)

    if show_grid and layout.get("show_grid", False):
        ax.yaxis.grid(
            True,
            linestyle=layout.get("grid_linestyle", "--"),
            linewidth=layout.get("grid_linewidth", 0.6),
            color=colors["grid"],
            alpha=0.9,
        )
        ax.set_axisbelow(True)
    else:
        ax.grid(False)


def render_single_chart(spec: Dict[str, Any], theme: Dict[str, Any], out_dir: str) -> str:
    """
    根据单张图表配置字典 spec 绘制并保存 PNG 文件。

    spec 核心字段：
    - filename    : 输出文件名（如 "fig01_revenue.png"）
    - type        : bar / grouped_bar / stacked_bar / line / bar_line_dual / pie / donut
    - layout      : "single"（默认）或 "double"
    - categories  : X 轴类目列表（如 ["2022", "2023", "2024", "2025"]）
    - series      : 系列列表 [{"name": "营业收入", "values": [...], "axis": "left"/"right", "kind": "bar"/"line"}]
    - y_label     : 左 Y 轴单位/标题（如 "亿元"）
    - y2_label    : 右 Y 轴单位/标题（仅 bar_line_dual，如 "%"）
    - y2_is_pct   : 布尔值，右轴是否格式化为百分比（默认 True）
    - show_labels : 是否显示数据标签（默认 True）
    - label_fmt   : 数值标签格式（如 "{:.1f}"）
    - y2_label_fmt: 右轴标签格式（如 "{:.1f}%"）
    """
    chart_type = spec.get("type", "bar").lower()
    layout_mode = spec.get("layout", "single").lower()
    figsize = (
        theme["layout"]["double_figsize"]
        if layout_mode == "double"
        else theme["layout"]["single_figsize"]
    )
    dpi = theme["layout"].get("dpi", 300)
    palette = theme["colors"]["series"]
    font_cfg = theme["font"]
    show_labels = spec.get("show_labels", True)
    label_fmt = spec.get("label_fmt", "")

    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    fig.patch.set_facecolor(theme["colors"]["background"])

    categories: List[str] = [str(c) for c in spec.get("categories", [])]
    series_list: List[Dict[str, Any]] = spec.get("series", [])
    x = np.arange(len(categories))

    # ==================== 1. 单系列柱状图 ====================
    if chart_type == "bar":
        style_axes(ax, theme)
        s = series_list[0]
        vals = s["values"]
        bar_w = theme["layout"].get("bar_width", 0.52)
        bars = ax.bar(x, vals, width=bar_w, color=palette[0], label=s.get("name", ""))
        ax.set_xticks(x)
        ax.set_xticklabels(categories)
        if show_labels:
            for bar in bars:
                h = bar.get_height()
                if h is not None and not np.isnan(h):
                    # 负值柱子标签放在柱底外侧，防止黑字压在深蓝柱体上
                    va = "bottom" if h >= 0 else "top"
                    y_off = 3 if h >= 0 else -4
                    ax.annotate(
                        format_value(h, label_fmt),
                        xy=(bar.get_x() + bar.get_width() / 2, h),
                        xytext=(0, y_off),
                        textcoords="offset points",
                        ha="center",
                        va=va,
                        fontsize=font_cfg["data_label_size"],
                        color=font_cfg["text_color"],
                    )
        # 若存在负值柱子，向下留出 25% 标签安全空间并补画 0 刻度细灰基线，防止负值数字撞到底轴
        numeric_vals = [float(v) for v in vals if v is not None and not np.isnan(v)]
        if numeric_vals and min(numeric_vals) < 0:
            ax.axhline(0, color=theme["colors"]["spine"], linewidth=0.8, zorder=1)
            cur_bot, cur_top = ax.get_ylim()
            ax.set_ylim(bottom=min(cur_bot, min(numeric_vals) * 1.28), top=cur_top)
        if s.get("name"):
            ax.legend(
                loc="upper center",
                bbox_to_anchor=(0.5, 1.13),
                frameon=False,
                fontsize=font_cfg["legend_size"],
            )

    # ==================== 2. 分组柱状图 ====================
    elif chart_type == "grouped_bar":
        style_axes(ax, theme)
        n_series = len(series_list)
        total_w = 0.72
        single_w = total_w / max(n_series, 1)
        for i, s in enumerate(series_list):
            offset = (i - (n_series - 1) / 2) * single_w
            color = palette[i % len(palette)]
            bars = ax.bar(
                x + offset,
                s["values"],
                width=single_w * 0.92,
                color=color,
                label=s.get("name", f"系列{i+1}"),
            )
            if show_labels and n_series <= 3:
                for bar in bars:
                    h = bar.get_height()
                    if h is not None and not np.isnan(h):
                        ax.annotate(
                            format_value(h, label_fmt),
                            xy=(bar.get_x() + bar.get_width() / 2, h),
                            xytext=(0, 2),
                            textcoords="offset points",
                            ha="center",
                            va="bottom",
                            fontsize=font_cfg["data_label_size"] - 0.5,
                        )
        ax.set_xticks(x)
        ax.set_xticklabels(categories)
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, 1.15),
            ncol=min(n_series, 4),
            frameon=False,
            fontsize=font_cfg["legend_size"],
        )

    # ==================== 3. 堆叠柱状图 ====================
    elif chart_type == "stacked_bar":
        style_axes(ax, theme)
        bottom = np.zeros(len(categories), dtype=float)
        bar_w = theme["layout"].get("bar_width", 0.52)
        for i, s in enumerate(series_list):
            vals = np.array([0.0 if v is None else float(v) for v in s["values"]])
            color = palette[i % len(palette)]
            ax.bar(
                x,
                vals,
                width=bar_w,
                bottom=bottom,
                color=color,
                label=s.get("name", f"系列{i+1}"),
            )
            bottom += vals
        if show_labels:
            for xi, total_h in enumerate(bottom):
                ax.annotate(
                    format_value(total_h, label_fmt),
                    xy=(xi, total_h),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha="center",
                    va="bottom",
                    fontsize=font_cfg["data_label_size"],
                )
        ax.set_xticks(x)
        ax.set_xticklabels(categories)
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, 1.15),
            ncol=min(len(series_list), 4),
            frameon=False,
            fontsize=font_cfg["legend_size"],
        )

    # ==================== 4. 折线图 ====================
    elif chart_type == "line":
        style_axes(ax, theme)
        lw = theme["layout"].get("line_width", 2.0)
        ms = theme["layout"].get("marker_size", 4.5)
        for i, s in enumerate(series_list):
            color = palette[i % len(palette)]
            ax.plot(
                x,
                s["values"],
                color=color,
                linewidth=lw,
                marker="o",
                markersize=ms,
                label=s.get("name", f"系列{i+1}"),
            )
            if show_labels and len(series_list) <= 2:
                for xi, yv in enumerate(s["values"]):
                    if yv is not None and not np.isnan(yv):
                        ax.annotate(
                            format_value(yv, label_fmt),
                            xy=(xi, yv),
                            xytext=(0, 5),
                            textcoords="offset points",
                            ha="center",
                            va="bottom",
                            fontsize=font_cfg["data_label_size"],
                        )
        ax.set_xticks(x)
        ax.set_xticklabels(categories)
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, 1.15),
            ncol=min(len(series_list), 4),
            frameon=False,
            fontsize=font_cfg["legend_size"],
        )

    # ==================== 5. 柱线双轴图（最常用：营收+同比增速） ====================
    elif chart_type == "bar_line_dual":
        style_axes(ax, theme)
        ax2 = ax.twinx()
        ax2.spines["top"].set_visible(False)
        ax2.spines["left"].set_visible(False)
        ax2.spines["right"].set_color(theme["colors"]["spine"])
        ax2.tick_params(axis="y", labelsize=font_cfg["tick_size"], length=3)

        bar_series = [s for s in series_list if s.get("kind", "bar") == "bar"]
        line_series = [s for s in series_list if s.get("kind") == "line"]

        bar_w = theme["layout"].get("bar_width", 0.50)
        n_bars = max(len(bar_series), 1)
        sub_w = bar_w / n_bars if n_bars > 1 else bar_w

        handles, labels = [], []
        for i, s in enumerate(bar_series):
            offset = (i - (n_bars - 1) / 2) * sub_w if n_bars > 1 else 0
            color = palette[i % len(palette)]
            bars = ax.bar(
                x + offset,
                s["values"],
                width=sub_w * 0.92,
                color=color,
                label=s.get("name", ""),
            )
            handles.append(bars)
            labels.append(s.get("name", ""))
            if show_labels and n_bars == 1:
                for bar in bars:
                    h = bar.get_height()
                    if h is not None and not np.isnan(h):
                        ax.annotate(
                            format_value(h, label_fmt),
                            xy=(bar.get_x() + bar.get_width() / 2, h),
                            xytext=(0, 3),
                            textcoords="offset points",
                            ha="center",
                            va="bottom",
                            fontsize=font_cfg["data_label_size"],
                        )

        lw = theme["layout"].get("line_width", 2.0)
        ms = theme["layout"].get("marker_size", 4.5)
        line_colors = [theme["colors"]["line_secondary"], theme["colors"]["line_tertiary"]]
        y2_fmt = spec.get("y2_label_fmt", "{:.1f}%")
        for j, s in enumerate(line_series):
            lcolor = line_colors[j % len(line_colors)]
            (line_obj,) = ax2.plot(
                x,
                s["values"],
                color=lcolor,
                linewidth=lw,
                marker="o",
                markersize=ms,
                label=s.get("name", ""),
            )
            handles.append(line_obj)
            labels.append(s.get("name", ""))
            if show_labels and len(line_series) == 1:
                for xi, yv in enumerate(s["values"]):
                    if yv is not None and not np.isnan(yv):
                        ax2.annotate(
                            format_value(yv, y2_fmt),
                            xy=(xi, yv),
                            xytext=(0, 6),
                            textcoords="offset points",
                            ha="center",
                            va="bottom",
                            fontsize=font_cfg["data_label_size"],
                            color=lcolor,
                            bbox=dict(boxstyle="round,pad=0.15", facecolor="#FFFFFF", edgecolor="none", alpha=0.88),
                        )

        ax.set_xticks(x)
        ax.set_xticklabels(categories)
        if spec.get("y2_label"):
            ax2.set_ylabel(spec["y2_label"], fontsize=font_cfg["axis_label_size"])
        if spec.get("y2_is_pct", True):
            ax2.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))

        # 合并左右轴图例置于顶部居中
        ax.legend(
            handles,
            labels,
            loc="upper center",
            bbox_to_anchor=(0.5, 1.15),
            ncol=min(len(labels), 4),
            frameon=False,
            fontsize=font_cfg["legend_size"],
        )

    # ==================== 6. 饼图 / 环形图（自动防小扇区文字撞叠） ====================
    elif chart_type in ("pie", "donut"):
        vals = series_list[0]["values"]
        total_v = sum(v for v in vals if v is not None) or 1.0
        pcts = [(float(v) / total_v * 100.0) if v is not None else 0.0 for v in vals]
        # 若存在任意小扇区（< 8%），自动切换为“大扇区内标占比 + 右侧完整图例”模式，彻底杜绝相邻小扇区标签撞叠
        has_small_slice = any(p < 8.0 for p in pcts)

        colors = [palette[i % len(palette)] for i in range(len(categories))]
        wedgeprops = {"edgecolor": "#FFFFFF", "linewidth": 1.2}
        if chart_type == "donut":
            wedgeprops["width"] = 0.45

        def _smart_autopct(pct: float) -> str:
            if has_small_slice and pct < 7.0:
                return ""
            return f"{pct:.1f}%"

        wedges, texts, autotexts = ax.pie(
            vals,
            labels=None if has_small_slice else categories,
            colors=colors,
            autopct=_smart_autopct,
            startangle=90,
            counterclock=False,
            pctdistance=0.75 if chart_type == "donut" else 0.65,
            wedgeprops=wedgeprops,
            textprops={"fontsize": font_cfg["tick_size"], "color": font_cfg["text_color"]},
        )
        for at in autotexts:
            at.set_color("#FFFFFF")
            at.set_fontsize(font_cfg["data_label_size"])
        if has_small_slice:
            legend_labels = [f"{cat}  {p:.1f}%" for cat, p in zip(categories, pcts)]
            ax.legend(
                wedges,
                legend_labels,
                loc="center left",
                bbox_to_anchor=(0.98, 0.5),
                frameon=False,
                fontsize=font_cfg["legend_size"],
            )
        ax.axis("equal")

    # ==================== 7. 公司发展历程时间轴 (timeline) ====================
    elif chart_type == "timeline":
        # 节点格式: spec["milestones"] = [{"year": "2015", "title": "...", "desc": "..."}, ...]
        milestones = spec.get("milestones", [])
        n_ms = max(len(milestones), 1)
        fig.set_size_inches(9.2, 3.8)
        ax.set_xlim(-0.6, n_ms - 0.4)
        ax.set_ylim(-1.55, 1.55)
        ax.axis("off")

        # 绘制国金深蓝水平主轴与箭头
        ax.annotate(
            "",
            xy=(n_ms - 0.45, 0),
            xytext=(-0.55, 0),
            arrowprops=dict(arrowstyle="-|>", color=palette[0], lw=2.4),
        )

        for idx, ms in enumerate(milestones):
            # 上下交替排布卡片，防止文字拥挤
            direction = 1 if idx % 2 == 0 else -1
            node_color = palette[0] if idx % 2 == 0 else palette[1]

            # 轴上圆点 + 竖向连接线
            ax.plot([idx, idx], [0, direction * 0.42], color=node_color, lw=1.4, zorder=2)
            ax.scatter([idx], [0], s=85, color=node_color, edgecolors="#FFFFFF", linewidths=1.5, zorder=4)

            # 年份胶囊标签
            year_y = direction * 0.56
            ax.text(
                idx,
                year_y,
                str(ms.get("year", "")),
                ha="center",
                va="center",
                fontsize=9.0,
                fontweight="bold",
                color="#FFFFFF",
                bbox=dict(boxstyle="round,pad=0.28", facecolor=node_color, edgecolor="none"),
                zorder=5,
            )

            # 事件说明卡片（浅灰蓝底色 + 细边框）
            card_y = direction * 1.08
            card_text = f"{ms.get('title', '')}\n{ms.get('desc', '')}".strip()
            ax.text(
                idx,
                card_y,
                card_text,
                ha="center",
                va="center",
                fontsize=7.8,
                color=font_cfg["text_color"],
                linespacing=1.35,
                bbox=dict(
                    boxstyle="round,pad=0.45",
                    facecolor="#F4F7FA",
                    edgecolor=palette[2],
                    linewidth=0.9,
                ),
                zorder=4,
            )

    # ==================== 8. 股权穿透架构图 (org_chart) ====================
    elif chart_type == "org_chart":
        # 结构: spec["top_nodes"] = [{"name": "...", "pct": "86.10%"}, ...]
        #       spec["center_node"] = "超颖电子电路股份有限公司（603175.SH）"
        #       spec["bottom_nodes"] = [{"name": "...", "pct": "100% / 生产基地"}, ...]
        top_nodes = spec.get("top_nodes", [])
        center_node = spec.get("center_node", "上市公司")
        bottom_nodes = spec.get("bottom_nodes", [])

        fig.set_size_inches(9.2, 4.4)
        ax.set_xlim(0, 10)
        ax.set_ylim(0, 6)
        ax.axis("off")

        def draw_box(cx, cy, w, h, text, bg, fg="#FFFFFF", border="none", fs=8.3, bold=False):
            rect = mpatches.FancyBboxPatch(
                (cx - w / 2, cy - h / 2),
                w,
                h,
                boxstyle="round,pad=0.08,rounding_size=0.12",
                facecolor=bg,
                edgecolor=border,
                linewidth=1.0,
                zorder=3,
            )
            ax.add_patch(rect)
            ax.text(
                cx,
                cy,
                text,
                ha="center",
                va="center",
                color=fg,
                fontsize=fs,
                fontweight="bold" if bold else "normal",
                linespacing=1.25,
                zorder=4,
            )

        # 中央上市公司核心框（国金深蓝）
        draw_box(5.0, 3.0, 4.4, 0.72, center_node, bg=palette[0], fg="#FFFFFF", fs=9.5, bold=True)

        # 上层股东框与持股连线
        if top_nodes:
            n_top = len(top_nodes)
            xs = np.linspace(1.3, 8.7, n_top) if n_top > 1 else [5.0]
            box_w = min(2.0, 7.6 / max(n_top, 1))
            # 水平汇流总线
            if n_top > 1:
                ax.plot([xs[0], xs[-1]], [3.85, 3.85], color=palette[1], lw=1.2, zorder=1)
            ax.plot([5.0, 5.0], [3.85, 3.36], color=palette[1], lw=1.2, zorder=1)
            for idx, nd in enumerate(top_nodes):
                x0 = xs[idx]
                bg = palette[1] if idx == 0 else "#EBF1F6"
                fg = "#FFFFFF" if idx == 0 else font_cfg["text_color"]
                draw_box(x0, 4.95, box_w, 0.95, nd.get("name", ""), bg=bg, fg=fg, border=palette[2], fs=7.6, bold=(idx == 0))
                ax.plot([x0, x0], [4.45, 3.85], color=palette[1], lw=1.2, zorder=1)
                if nd.get("pct"):
                    ax.text(
                        x0,
                        4.14,
                        nd["pct"],
                        ha="center",
                        va="center",
                        fontsize=7.6,
                        fontweight="bold" if idx == 0 else "normal",
                        color=palette[0],
                        bbox=dict(boxstyle="square,pad=0.12", facecolor="#FFFFFF", edgecolor="none"),
                        zorder=5,
                    )

        # 下层子公司/三大基地框与连线
        if bottom_nodes:
            n_bot = len(bottom_nodes)
            xs = np.linspace(1.6, 8.4, n_bot) if n_bot > 1 else [5.0]
            box_w = min(2.4, 7.8 / max(n_bot, 1))
            if n_bot > 1:
                ax.plot([xs[0], xs[-1]], [2.15, 2.15], color=palette[1], lw=1.2, zorder=1)
            ax.plot([5.0, 5.0], [2.64, 2.15], color=palette[1], lw=1.2, zorder=1)
            for idx, nd in enumerate(bottom_nodes):
                x0 = xs[idx]
                draw_box(x0, 1.05, box_w, 0.95, nd.get("name", ""), bg="#F4F7FA", fg=font_cfg["text_color"], border=palette[1], fs=7.8)
                ax.plot([x0, x0], [2.15, 1.55], color=palette[1], lw=1.2, zorder=1)
                if nd.get("pct"):
                    ax.text(
                        x0,
                        1.85,
                        nd["pct"],
                        ha="center",
                        va="center",
                        fontsize=7.4,
                        color=palette[1],
                        bbox=dict(boxstyle="square,pad=0.12", facecolor="#FFFFFF", edgecolor="none"),
                        zorder=5,
                    )

    # ==================== 9. 客户与应用场景矩阵图 (customer_matrix) ====================
    elif chart_type == "customer_matrix":
        # 结构: spec["columns"] = [{"header": "汽车电子", "subtitle": "占比约57%", "items": ["大陆汽车", "法雷奥", ...], "footer": "IATF 16949认证"}, ...]
        columns = spec.get("columns", [])
        n_col = max(len(columns), 1)
        fig.set_size_inches(9.2, 3.8)
        ax.set_xlim(0, 10)
        ax.set_ylim(0, 5.2)
        ax.axis("off")

        gap = 0.25
        col_w = (9.6 - gap * (n_col - 1)) / n_col
        for idx, col in enumerate(columns):
            x_left = 0.2 + idx * (col_w + gap)
            cx = x_left + col_w / 2
            head_bg = palette[0] if idx == 0 else (palette[1] if idx == 1 else palette[2])

            # 顶部深蓝标题头
            header_rect = mpatches.FancyBboxPatch(
                (x_left, 4.15), col_w, 0.85,
                boxstyle="round,pad=0.04,rounding_size=0.10",
                facecolor=head_bg, edgecolor="none", zorder=3
            )
            ax.add_patch(header_rect)
            head_txt = f"{col.get('header', '')}\n{col.get('subtitle', '')}".strip()
            ax.text(cx, 4.58, head_txt, ha="center", va="center", color="#FFFFFF", fontsize=8.8, fontweight="bold", linespacing=1.25, zorder=4)

            # 中间浅灰蓝内容框
            body_rect = mpatches.FancyBboxPatch(
                (x_left, 0.95), col_w, 3.05,
                boxstyle="round,pad=0.04,rounding_size=0.10",
                facecolor="#F4F7FA", edgecolor=palette[2], linewidth=0.9, zorder=2
            )
            ax.add_patch(body_rect)
            items = col.get("items", [])
            if items:
                ys = np.linspace(3.55, 1.35, len(items)) if len(items) > 1 else [2.45]
                for iy, it_txt in zip(ys, items):
                    pill = mpatches.FancyBboxPatch(
                        (x_left + 0.18, iy - 0.22), col_w - 0.36, 0.44,
                        boxstyle="round,pad=0.02,rounding_size=0.08",
                        facecolor="#FFFFFF", edgecolor="#CBD5E1", linewidth=0.8, zorder=3
                    )
                    ax.add_patch(pill)
                    ax.text(cx, iy, it_txt, ha="center", va="center", color=font_cfg["text_color"], fontsize=8.0, zorder=4)

            # 底部壁垒/特征条
            if col.get("footer"):
                foot_rect = mpatches.FancyBboxPatch(
                    (x_left, 0.18), col_w, 0.60,
                    boxstyle="round,pad=0.04,rounding_size=0.08",
                    facecolor="#EBF1F6", edgecolor="none", zorder=3
                )
                ax.add_patch(foot_rect)
                ax.text(cx, 0.48, col["footer"], ha="center", va="center", color=palette[0], fontsize=7.6, fontweight="bold", zorder=4)

    else:
        raise ValueError(f"不支持的图表类型: {chart_type}")

    # 设置左 Y 轴单位与顶部留白（防止数据标签顶到图例）
    if chart_type not in ("pie", "donut", "timeline", "org_chart", "customer_matrix"):
        if spec.get("y_label"):
            ax.set_ylabel(spec["y_label"], fontsize=font_cfg["axis_label_size"])
        ymin, ymax = ax.get_ylim()
        if ymax > 0:
            ax.set_ylim(ymin, ymax * 1.16)
        # 若 X 轴标签较长则适当旋转
        if any(len(c) > 6 for c in categories):
            plt.setp(ax.get_xticklabels(), rotation=20, ha="right")

    # 注意：国金研报的图表标题由 Word 模板的 ContentGraphTitle 段落承载，不在图片内部重复画大标题
    if spec.get("inner_title"):
        ax.set_title(spec["inner_title"], fontsize=font_cfg["title_size"], pad=22)

    plt.tight_layout()

    # 像素级数据标签重叠自动检测（防止双轴/密集柱图出现文字糊在一起）
    overlaps = detect_label_overlaps(fig)
    if overlaps:
        print(f"  🟡 [标签重叠提醒] {spec['filename']} 检测到 {len(overlaps)} 处数据标签像素框重叠：{overlaps[:3]}")

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, spec["filename"])
    fig.savefig(out_path, dpi=dpi, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return out_path


def detect_label_overlaps(fig) -> List[str]:
    """
    渲染画布后提取所有坐标轴内数据标签（Text）的屏幕像素边界框，
    检查两两之间是否在水平与垂直方向均发生 > 2px 重叠，及时暴露标签糊成一团的图表。
    """
    try:
        fig.canvas.draw()
        renderer = fig.canvas.get_renderer()
    except Exception:
        return []

    boxes = []
    for ax in fig.axes:
        for txt in ax.texts:
            content = (txt.get_text() or "").strip()
            if not content or not txt.get_visible():
                continue
            bb = txt.get_window_extent(renderer=renderer)
            if bb.width <= 0 or bb.height <= 0:
                continue
            boxes.append((content.replace("\n", " ")[:14], bb))

    overlaps: List[str] = []
    for i in range(len(boxes)):
        t1, b1 = boxes[i]
        for j in range(i + 1, len(boxes)):
            t2, b2 = boxes[j]
            ox = min(b1.x1, b2.x1) - max(b1.x0, b2.x0)
            oy = min(b1.y1, b2.y1) - max(b1.y0, b2.y0)
            if ox > 2.0 and oy > 2.0:
                overlaps.append(f"'{t1}'↔'{t2}'")
    return overlaps


def main() -> None:
    """命令行入口：解析参数并批量出图。"""
    parser = argparse.ArgumentParser(description="国金证券标准图表批量生成脚本")
    parser.add_argument("--spec", required=True, help="图表配置 JSON 文件路径")
    parser.add_argument("--out-dir", required=True, help="输出 PNG 图片目录")
    parser.add_argument(
        "--theme",
        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "chart_theme.json"),
        help="色板配置 JSON 路径（默认同目录 chart_theme.json）",
    )
    args = parser.parse_args()

    theme = load_theme(args.theme)
    setup_matplotlib(theme)

    with open(args.spec, "r", encoding="utf-8") as f:
        data = json.load(f)

    charts = data if isinstance(data, list) else data.get("charts", [])
    if not charts:
        print("[警告] 配置中未发现任何图表条目。")
        sys.exit(0)

    for idx, chart_spec in enumerate(charts, 1):
        out_path = render_single_chart(chart_spec, theme, args.out_dir)
        print(f"[OK] ({idx}/{len(charts)}) 已生成图表: {out_path}")


if __name__ == "__main__":
    main()
