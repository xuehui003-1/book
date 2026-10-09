#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tools/generate_p13_images.py
为《项目13 财务AI智能体全流程综合实战》全量生成 6 幅专业级教学与实训图表资源。
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.font_manager as fm
import numpy as np

OUTPUT_DIR = "/tmp/p13_images"
os.makedirs(OUTPUT_DIR, exist_ok=True)

FONT_PATH = "/usr/local/lib/python3.11/dist-packages/mplfonts/fonts/NotoSansCJKsc-Regular.otf"
PROP_TITLE = fm.FontProperties(fname=FONT_PATH, size=13, weight="bold")
PROP_HEADING = fm.FontProperties(fname=FONT_PATH, size=10, weight="bold")
PROP_BODY = fm.FontProperties(fname=FONT_PATH, size=8.5)
PROP_SMALL = fm.FontProperties(fname=FONT_PATH, size=7.5)

def draw_rounded_rect(ax, x, y, w, h, bg_color, border_color="#CBD5E1", border_width=1.5, rx=0.02):
    rect = patches.FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={rx}",
                                 facecolor=bg_color, edgecolor=border_color, linewidth=border_width)
    ax.add_patch(rect)

def generate_fig1():
    """图13-1 项目13教学导图"""
    fig, ax = plt.subplots(figsize=(10, 4.5), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 4.5)
    ax.axis("off")
    fig.patch.set_facecolor("#FFFFFF")

    # 标题区
    draw_rounded_rect(ax, 0.5, 3.8, 9.0, 0.55, "#1E3A8A", "#1E3A8A", rx=0.03)
    ax.text(5.0, 4.07, "《项目13 财务AI智能体全流程综合实战》整体教学与技能递进导图",
            fontproperties=PROP_TITLE, color="#FFFFFF", ha="center", va="center")

    steps = [
        ("阶段一：数据底座与清洗", "任务 13.1\n• 多源异构数据湖搭建\n• 飞书多维表格中台\n• 规则+Prompt清洗智能体\n• 实体归一与发票查重", "#EEF2FF", "#4F46E5"),
        ("阶段二：深度透视与预警", "任务 13.2\n• 12家门店杜邦指标拆解\n• 四象限波士顿矩阵落位\n• 3号/8号病灶门店穿透\n• 三级红黄牌风险预警", "#ECFDF5", "#059669"),
        ("阶段三：BI驾驶舱叙事", "任务 13.3\n• 高管全局经营驾驶舱\n• 门店端日常自查看板\n• 成本动因联动下钻分析\n• 金字塔原理诊断白皮书", "#FEF3C7", "#D97706"),
        ("阶段四：自动化与协同", "任务 13.4\n• 定时触发与循环批处理\n• 食材超标条件分支分流\n• 三角色智能体协同会诊\n• Streamlit交互应用生成", "#EFF6FF", "#2563EB"),
        ("阶段五：企业级交付答辩", "项目核心实操\n• 48小时五份专业底稿\n• 反AI代做随机三道抽测\n• 全书统一百分配分制\n• 现场动态参数重跑答辩", "#FDF2F8", "#DB2777")
    ]

    for i, (title, content, bg, border) in enumerate(steps):
        x = 0.5 + i * 1.85
        y = 0.6
        w = 1.6
        h = 2.9
        draw_rounded_rect(ax, x, y, w, h, bg, border, border_width=1.8, rx=0.03)
        # 阶段标题
        draw_rounded_rect(ax, x, y + h - 0.5, w, 0.5, border, border, rx=0.02)
        ax.text(x + w/2, y + h - 0.25, title, fontproperties=PROP_HEADING, color="#FFFFFF", ha="center", va="center")
        # 阶段正文
        ax.text(x + 0.1, y + h - 0.7, content, fontproperties=PROP_BODY, color="#1E293B", ha="left", va="top", linespacing=1.6)

        # 箭头连接
        if i < 4:
            ax.annotate("", xy=(x + w + 0.22, y + h/2), xytext=(x + w + 0.02, y + h/2),
                        arrowprops=dict(arrowstyle="->", color="#94A3B8", lw=2.5, mutation_scale=15))

    plt.tight_layout()
    out_path = os.path.join(OUTPUT_DIR, "image1.png")
    plt.savefig(out_path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close()
    print("Generated:", out_path)

def generate_fig2():
    """图13-2 鲜合餐饮多源财务数据清洗与整合架构图"""
    fig, ax = plt.subplots(figsize=(10, 4.6), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 4.6)
    ax.axis("off")
    fig.patch.set_facecolor("#FFFFFF")

    # 标题区
    draw_rounded_rect(ax, 0.5, 3.9, 9.0, 0.5, "#0F766E", "#0F766E", rx=0.03)
    ax.text(5.0, 4.15, "鲜合餐饮多源异构财务数据清洗与整合技术架构",
            fontproperties=PROP_TITLE, color="#FFFFFF", ha="center", va="center")

    # 左侧：数据源
    draw_rounded_rect(ax, 0.5, 0.5, 2.5, 3.1, "#F0FDFA", "#14B8A6", rx=0.03)
    ax.text(1.75, 3.3, "1. 多源原始异构数据", fontproperties=PROP_HEADING, color="#0F766E", ha="center")
    sources = [
        "• POS销售明细流水（微信/支付宝/现金）",
        "• 外卖平台月度对账单（扣佣/满减补贴）",
        "• 本地生鲜食材采购专票（农产品发票）",
        "• 中央厨房统一配送调味品/半成品专票",
        "• 门店日常费用报销台账与电子发票"
    ]
    ax.text(0.65, 3.0, "\n\n".join(sources), fontproperties=PROP_SMALL, color="#1E293B", ha="left", va="top")

    # 中间：清洗治理智能体
    draw_rounded_rect(ax, 3.6, 0.5, 2.8, 3.1, "#FEF2F2", "#EF4444", rx=0.03)
    ax.text(5.0, 3.3, "2. 智能体清洗与质控规则", fontproperties=PROP_HEADING, color="#B91C1C", ha="center")
    rules = [
        "【实体归一】映射12家标准门店编号",
        "【数值校验】拦截非退单负数与畸大值",
        "【专票验真】校验9%/13%法定税率乘积",
        "【防重过滤】3天窗口同商户同金额查重",
        "【业财对齐】核销外卖配送费与平台佣金"
    ]
    ax.text(3.75, 3.0, "\n\n".join(rules), fontproperties=PROP_SMALL, color="#1E293B", ha="left", va="top")

    # 右侧：多维表格与宽表资产
    draw_rounded_rect(ax, 7.0, 0.5, 2.5, 3.1, "#EFF6FF", "#3B82F6", rx=0.03)
    ax.text(8.25, 3.3, "3. 飞书多维表格数据资产", fontproperties=PROP_HEADING, color="#1D4ED8", ha="center")
    outputs = [
        "• 门店主数据字典表（主键关联）",
        "• 统一销售流水清洗表",
        "• 食材BOM采购与消耗跟踪表",
        "• 鲜合12家门店标准化经营宽表",
        "• 直通下游杜邦分析与BI驾驶舱"
    ]
    ax.text(7.15, 3.0, "\n\n".join(outputs), fontproperties=PROP_SMALL, color="#1E293B", ha="left", va="top")

    # 箭头连接
    ax.annotate("", xy=(3.5, 2.05), xytext=(3.05, 2.05), arrowprops=dict(arrowstyle="->", color="#0F766E", lw=2.5))
    ax.annotate("", xy=(6.9, 2.05), xytext=(6.45, 2.05), arrowprops=dict(arrowstyle="->", color="#EF4444", lw=2.5))

    plt.tight_layout()
    out_path = os.path.join(OUTPUT_DIR, "image2.png")
    plt.savefig(out_path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close()
    print("Generated:", out_path)

def generate_fig3():
    """图13-3 鲜合餐饮12家直营门店经营效益四象限波士顿矩阵"""
    fig, ax = plt.subplots(figsize=(8.5, 5.2), dpi=200)
    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#F8FAFC")

    # 数据：翻台率(X)，食材成本率(Y)
    stores = [
        ("01 北京王府井", 4.2, 36.0, 37.3, "#059669"),
        ("02 北京中关村", 3.5, 37.0, 22.9, "#2563EB"),
        ("03 北京朝阳大悦城", 3.8, 44.0, 23.2, "#DC2626"), # 核心问题店
        ("04 天津和平", 3.1, 36.0, 20.2, "#2563EB"),
        ("05 天津南开", 3.0, 37.0, 15.8, "#2563EB"),
        ("06 上海南京东路", 4.5, 36.0, 51.7, "#059669"), # 标杆店
        ("07 上海陆家嘴", 3.9, 36.0, 33.5, "#2563EB"),
        ("08 上海淮海中路", 1.8, 43.0, -4.1, "#DC2626"), # 亏损店
        ("09 杭州武林广场", 3.4, 36.0, 24.4, "#2563EB"),
        ("10 杭州西湖湖滨", 3.8, 36.0, 30.1, "#059669"),
        ("11 南京新街口", 3.2, 37.0, 21.6, "#2563EB"),
        ("12 苏州金鸡湖", 2.9, 36.0, 15.4, "#2563EB")
    ]

    # 象限基准线：翻台率 3.4次/天，食材成本率 38.0%
    x_mean = 3.4
    y_mean = 38.0

    ax.axvline(x=x_mean, color="#94A3B8", linestyle="--", linewidth=1.5)
    ax.axhline(y=y_mean, color="#94A3B8", linestyle="--", linewidth=1.5)

    # 象限背景色
    ax.axhspan(y_mean, 46.0, xmin=(x_mean-1.5)/(5.0-1.5), xmax=1.0, facecolor="#FEE2E2", alpha=0.35) # 右上：高翻台高成本（问题店）
    ax.axhspan(y_mean, 46.0, xmin=0.0, xmax=(x_mean-1.5)/(5.0-1.5), facecolor="#FEE2E2", alpha=0.7) # 左上：低翻台高成本（瘦狗店）
    ax.axhspan(34.0, y_mean, xmin=(x_mean-1.5)/(5.0-1.5), xmax=1.0, facecolor="#DCFCE7", alpha=0.5) # 右下：高翻台健康成本（明星店）
    ax.axhspan(34.0, y_mean, xmin=0.0, xmax=(x_mean-1.5)/(5.0-1.5), facecolor="#E0F2FE", alpha=0.4) # 左下：稳健金牛店

    # 象限标注
    ax.text(1.7, 45.2, "【瘦狗象限】低翻台 / 高成本\n典型：08淮海中路（亏损）", fontproperties=PROP_HEADING, color="#B91C1C", ha="left")
    ax.text(4.8, 45.2, "【问题象限】高翻台 / 高成本\n典型：03朝阳大悦城（损耗高）", fontproperties=PROP_HEADING, color="#C2410C", ha="right")
    ax.text(4.8, 34.4, "【明星象限】高翻台 / 控本佳\n标杆：06南京东路、01王府井", fontproperties=PROP_HEADING, color="#15803D", ha="right")
    ax.text(1.7, 34.4, "【稳健金牛】翻台适中 / 控本健康\n分布：和平、南开、新街口、金鸡湖", fontproperties=PROP_HEADING, color="#0369A1", ha="left")

    for name, x, y, profit, color in stores:
        size = max(80, profit * 8 + 60) if profit > 0 else 100
        ax.scatter(x, y, s=size, color=color, alpha=0.85, edgecolors="#0F172A", linewidth=1.2, zorder=4)
        offset_y = 0.35 if "大悦城" not in name else 0.45
        offset_x = 0.0
        if "王府井" in name: offset_x = -0.15
        if "金鸡湖" in name: offset_y = -0.4
        ax.text(x + offset_x, y + offset_y, f"{name}\n({profit:.1f}万)", fontproperties=PROP_SMALL,
                color="#0F172A", ha="center", va="bottom", zorder=5)

    ax.set_xlim(1.5, 5.0)
    ax.set_ylim(34.0, 46.0)
    ax.set_xlabel("门店日均翻台率（次/天）——【基准均值：3.4 次/天】", fontproperties=PROP_HEADING, color="#1E293B", labelpad=8)
    ax.set_ylabel("门店食材成本率（%）——【全连锁均值：38.0%】", fontproperties=PROP_HEADING, color="#1E293B", labelpad=8)
    ax.set_title("鲜合餐饮全国12家直营门店经营效益四象限波士顿矩阵（气泡大小代表经营净利）", fontproperties=PROP_TITLE, color="#0F172A", pad=12)

    ax.grid(True, linestyle=":", color="#CBD5E1", alpha=0.6)
    plt.tight_layout()
    out_path = os.path.join(OUTPUT_DIR, "image3.png")
    plt.savefig(out_path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close()
    print("Generated:", out_path)

def generate_fig4():
    """图13-4 鲜合餐饮高管经营分析驾驶舱与门店端双重视角BI交互看板"""
    fig, ax = plt.subplots(figsize=(10, 5.0), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5.0)
    ax.axis("off")
    fig.patch.set_facecolor("#FFFFFF")

    # 顶栏
    draw_rounded_rect(ax, 0.4, 4.35, 9.2, 0.55, "#1E293B", "#1E293B", rx=0.02)
    ax.text(0.7, 4.62, "鲜合餐饮业财融合智能决策支持大屏 (Power BI / FineBI Pro)", fontproperties=PROP_TITLE, color="#FFFFFF", va="center")
    ax.text(9.3, 4.62, "口径：2024半年度全连锁12家门店", fontproperties=PROP_BODY, color="#94A3B8", ha="right", va="center")

    # KPI 卡片
    kpis = [
        ("全连锁营收总规模", "1,694.4 万元", "+14.2% 同比", "#10B981"),
        ("半年度经营净利润", "271.1 万元", "净利率 16.0%", "#3B82F6"),
        ("全连锁食材成本率", "38.0 %", "控本警戒线 40%", "#F59E0B"),
        ("全连锁平均翻台率", "3.4 次/天", "标杆达 4.5次", "#8B5CF6")
    ]
    for i, (k, v, sub, c) in enumerate(kpis):
        x = 0.4 + i * 2.35
        draw_rounded_rect(ax, x, 3.45, 2.15, 0.75, "#F8FAFC", "#E2E8F0", rx=0.02)
        ax.text(x + 0.1, 3.98, k, fontproperties=PROP_SMALL, color="#64748B")
        ax.text(x + 0.1, 3.68, v, fontproperties=PROP_HEADING, color="#0F172A", size=12)
        ax.text(x + 2.05, 3.68, sub, fontproperties=PROP_SMALL, color=c, ha="right")

    # 左侧：高管视角驾驶舱
    draw_rounded_rect(ax, 0.4, 0.4, 4.45, 2.9, "#FFFFFF", "#CBD5E1", rx=0.02)
    draw_rounded_rect(ax, 0.4, 2.95, 4.45, 0.35, "#EEF2FF", "#EEF2FF", rx=0.01)
    ax.text(0.6, 3.12, "【高管全局视角】门店营收/利润排行与成本瀑布图", fontproperties=PROP_HEADING, color="#3730A3", va="center")

    # 示意柱状图
    stores_short = ["06南京东", "03大悦城", "01王府井", "07陆家嘴", "10西湖", "08淮海"]
    revenues = [245.8, 210.5, 185.6, 176.2, 158.6, 105.2]
    profits = [51.7, 23.2, 37.3, 33.5, 30.1, -4.1]
    for idx, (st, rev, pr) in enumerate(zip(stores_short, revenues, profits)):
        bx = 0.6
        by = 2.5 - idx * 0.35
        ax.text(bx, by + 0.05, st, fontproperties=PROP_SMALL, color="#334155")
        # 营收条
        bar_w = (rev / 260.0) * 1.8
        draw_rounded_rect(ax, bx + 0.9, by, bar_w, 0.18, "#93C5FD", "#60A5FA", rx=0.005)
        # 利润条
        p_color = "#10B981" if pr > 0 else "#EF4444"
        p_w = max(0.08, abs(pr) / 60.0 * 0.9)
        draw_rounded_rect(ax, bx + 2.85, by, p_w, 0.18, p_color, p_color, rx=0.005)
        ax.text(bx + 2.85 + p_w + 0.05, by + 0.02, f"{pr:.1f}万", fontproperties=PROP_SMALL, color=p_color)

    # 右侧：门店穿透视角
    draw_rounded_rect(ax, 5.15, 0.4, 4.45, 2.9, "#FFFFFF", "#CBD5E1", rx=0.02)
    draw_rounded_rect(ax, 5.15, 2.95, 4.45, 0.35, "#FEF3C7", "#FEF3C7", rx=0.01)
    ax.text(5.35, 3.12, "【门店下钻视角】异常单店（03朝阳大悦城 / 08淮海店）下钻", fontproperties=PROP_HEADING, color="#92400E", va="center")

    panel_texts = [
        "• 03店病灶穿透：食材成本率44.0%（超标6%），月度食材损耗达12.6万",
        "  - 核心肉类损耗率偏离标准BOM值 +8.5%，疑似供应商批次抽检缺位",
        "  - 会员复购率断崖降至21%（均值32.6%），客诉集中在菜品品质不稳定",
        "• 08店病灶穿透：租金折旧占比高达37.0%，日翻台仅1.8次陷入持续亏损",
        "  - 保本点翻台率需达 2.4 次/天，当前面临严重的经营性现金流断流风险",
        "• 店长行动方案：调整BOM直采、启动老客召回促销、申请业主减免租金"
    ]
    ax.text(5.3, 2.65, "\n\n".join(panel_texts), fontproperties=PROP_SMALL, color="#1E293B", va="top")

    plt.tight_layout()
    out_path = os.path.join(OUTPUT_DIR, "image4.png")
    plt.savefig(out_path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close()
    print("Generated:", out_path)

def generate_fig5():
    """图13-5 鲜合餐饮端到端自动化财务分析工作流拓扑编排图"""
    fig, ax = plt.subplots(figsize=(10, 4.6), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 4.6)
    ax.axis("off")
    fig.patch.set_facecolor("#FFFFFF")

    # 标题区
    draw_rounded_rect(ax, 0.5, 3.9, 9.0, 0.5, "#1E40AF", "#1E40AF", rx=0.03)
    ax.text(5.0, 4.15, "鲜合餐饮月度财务分析端到端自动化工作流编排拓扑",
            fontproperties=PROP_TITLE, color="#FFFFFF", ha="center", va="center")

    # 节点定义
    # 1. 触发器
    draw_rounded_rect(ax, 0.5, 2.3, 1.6, 1.2, "#EFF6FF", "#3B82F6", rx=0.03)
    ax.text(1.3, 3.15, "1. 定时触发", fontproperties=PROP_HEADING, color="#1E40AF", ha="center")
    ax.text(1.3, 2.65, "每月1日 08:00\n监听12家门店\n月度数据包到位", fontproperties=PROP_SMALL, color="#334155", ha="center")

    # 2. 循环批处理
    draw_rounded_rect(ax, 2.5, 2.0, 2.0, 1.7, "#F0FDF4", "#16A34A", rx=0.03)
    ax.text(3.5, 3.35, "2. 循环批处理节点", fontproperties=PROP_HEADING, color="#15803D", ha="center")
    ax.text(3.5, 2.7, "• 遍历门店文件列表\n• 字段对齐与专票查重\n• 自动核算食材BOM\n• 生成各单店指标宽表", fontproperties=PROP_SMALL, color="#334155", ha="center")

    # 3. 条件分支
    draw_rounded_rect(ax, 5.0, 2.1, 1.8, 1.5, "#FEF3C7", "#D97706", rx=0.03)
    ax.text(5.9, 3.25, "3. 规则判断分支", fontproperties=PROP_HEADING, color="#B45309", ha="center")
    ax.text(5.9, 2.65, "食材成本率>40% ?\n或 经营净利<0 ?\n(满足其一即触发)", fontproperties=PROP_SMALL, color="#B45309", ha="center")

    # 分支A：异常门店多智能体深度会诊
    draw_rounded_rect(ax, 7.3, 2.5, 2.2, 1.2, "#FEE2E2", "#DC2626", rx=0.03)
    ax.text(8.4, 3.35, "分支A：多智能体深度会诊", fontproperties=PROP_HEADING, color="#B91C1C", ha="center")
    ax.text(8.4, 2.85, "分析师-风控官-管理顾问\n输出红牌穿透预警与对策", fontproperties=PROP_SMALL, color="#7F1D1D", ha="center")

    # 分支B：正常门店标准月报生成
    draw_rounded_rect(ax, 7.3, 1.0, 2.2, 1.1, "#F1F5F9", "#64748B", rx=0.03)
    ax.text(8.4, 1.75, "分支B：标准月报生成", fontproperties=PROP_HEADING, color="#334155", ha="center")
    ax.text(8.4, 1.35, "汇总生成常态化经营宽表", fontproperties=PROP_SMALL, color="#475569", ha="center")

    # 汇聚与推送
    draw_rounded_rect(ax, 4.0, 0.4, 3.8, 0.9, "#FAF5FF", "#9333EA", rx=0.03)
    ax.text(5.9, 0.95, "4. 成果汇聚与多渠道自动化发布", fontproperties=PROP_HEADING, color="#7E22CE", ha="center")
    ax.text(5.9, 0.65, "Markdown白皮书渲染 + 飞书机器人推送高管预警 + 邮件自动归档", fontproperties=PROP_SMALL, color="#581C87", ha="center")

    # 连线
    ax.annotate("", xy=(2.45, 2.85), xytext=(2.15, 2.85), arrowprops=dict(arrowstyle="->", color="#3B82F6", lw=2))
    ax.annotate("", xy=(4.95, 2.85), xytext=(4.55, 2.85), arrowprops=dict(arrowstyle="->", color="#16A34A", lw=2))
    ax.annotate("", xy=(7.25, 3.1), xytext=(6.85, 3.0), arrowprops=dict(arrowstyle="->", color="#DC2626", lw=2))
    ax.text(7.0, 3.2, "是 (异常)", fontproperties=PROP_SMALL, color="#DC2626")
    ax.annotate("", xy=(7.25, 1.55), xytext=(6.85, 2.3), arrowprops=dict(arrowstyle="->", color="#64748B", lw=2))
    ax.text(7.0, 1.9, "否 (稳健)", fontproperties=PROP_SMALL, color="#64748B")

    ax.annotate("", xy=(7.8, 0.85), xytext=(8.4, 1.0), arrowprops=dict(arrowstyle="->", color="#64748B", lw=2))
    ax.annotate("", xy=(7.8, 0.85), xytext=(8.4, 2.5), arrowprops=dict(arrowstyle="->", color="#DC2626", lw=2))

    plt.tight_layout()
    out_path = os.path.join(OUTPUT_DIR, "image5.png")
    plt.savefig(out_path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close()
    print("Generated:", out_path)

def generate_fig6():
    """图13-6 鲜合多智能体协同会诊交互与轻应用展示界面"""
    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 4.8)
    ax.axis("off")
    fig.patch.set_facecolor("#FFFFFF")

    # 标题区
    draw_rounded_rect(ax, 0.5, 4.1, 9.0, 0.5, "#4338CA", "#4338CA", rx=0.03)
    ax.text(5.0, 4.35, "鲜合餐饮多智能体协同会诊机制与 Streamlit 轻应用交互原型",
            fontproperties=PROP_TITLE, color="#FFFFFF", ha="center", va="center")

    # 上半部：三角色协同
    agents = [
        ("角色1：财务分析师", "【职责：客观陈述数据异动】\n• 识别03店食材成本率攀升至44.0%\n• 指出08店租金占比37.0%持续亏损\n• 严禁主观推测，仅输出结构化事实", "#EFF6FF", "#3B82F6"),
        ("角色2：财务风控官", "【职责：审慎怀疑与合规挑刺】\n• 质询03店肉类采购偏离BOM +8.5%\n• 警示08店经营性现金流断流风险\n• 索取供应商送货单与专票核查证据", "#FEF2F2", "#EF4444"),
        ("角色3：管理咨询顾问", "【职责：矛盾仲裁与刚性闭环】\n• 仲裁分歧：严控BOM直采+激活复购\n• 08店提出降租30%或启动闭店止损\n• 最终收敛为落地执行白皮书", "#F0FDF4", "#22C55E")
    ]
    for i, (name, role, bg, border) in enumerate(agents):
        x = 0.5 + i * 3.1
        draw_rounded_rect(ax, x, 2.4, 2.8, 1.5, bg, border, rx=0.02)
        ax.text(x + 1.4, 3.65, name, fontproperties=PROP_HEADING, color=border, ha="center")
        ax.text(x + 0.15, 3.4, role, fontproperties=PROP_SMALL, color="#1E293B", va="top")
        if i < 2:
            ax.annotate("", xy=(x + 3.05, 3.15), xytext=(x + 2.85, 3.15), arrowprops=dict(arrowstyle="->", color="#94A3B8", lw=2))

    # 下半部：Streamlit 前端轻应用
    draw_rounded_rect(ax, 0.5, 0.4, 9.0, 1.8, "#F8FAFC", "#CBD5E1", rx=0.02)
    draw_rounded_rect(ax, 0.5, 1.85, 9.0, 0.35, "#334155", "#334155", rx=0.01)
    ax.text(0.7, 2.02, "Streamlit 财务智能体轻应用交互端原型 (http://localhost:8501)", fontproperties=PROP_HEADING, color="#FFFFFF", va="center")

    # 侧边栏
    draw_rounded_rect(ax, 0.7, 0.55, 2.2, 1.2, "#F1F5F9", "#E2E8F0", rx=0.01)
    ax.text(0.8, 1.55, "【侧边栏参数控制】\n• 门店选择：03朝阳大悦城店\n• 食材成本阈值：[ 40.0% ]\n• 翻台率警戒线：[ 3.0次 ]\n• [ 点击一键触发三方会诊 ]", fontproperties=PROP_SMALL, color="#334155", va="top")

    # 主对话展示
    draw_rounded_rect(ax, 3.1, 0.55, 6.2, 1.2, "#FFFFFF", "#E2E8F0", rx=0.01)
    ax.text(3.25, 1.55, "【会诊交互输出窗口】\n[风控官] 质询：03店6月份牛肉采购单价较5月份上涨14.2%，且库房报损率达4.8%，远超连锁2.0%容忍度！\n[管理顾问] 仲裁对策：1. 总部供应链立即冻结该本地批次生鲜采购；2. 派驻内审专员现场盘点冷库；3. 优化促销套餐！\n[系统状态] 报告已生成并自动推送至鲜合高管管理群（导出 PDF / 导出 Excel 底稿已就绪）", fontproperties=PROP_SMALL, color="#0F172A", va="top")

    plt.tight_layout()
    out_path = os.path.join(OUTPUT_DIR, "image6.png")
    plt.savefig(out_path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close()
    print("Generated:", out_path)

if __name__ == "__main__":
    generate_fig1()
    generate_fig2()
    generate_fig3()
    generate_fig4()
    generate_fig5()
    generate_fig6()
    print("All 6 figures for Project 13 generated successfully in", OUTPUT_DIR)
