#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tools/generate_p13_images.py
为《项目13 财务AI智能体全流程综合实战》全量生成 6 幅专业级教学与实训图表资源。
严格执行全书统一出版视觉规范：
1. 图13-1：对齐全书前12个项目思维导图体例（中心宝蓝圆球 #0000EC + 四分支黄绿紫粉 + 三级子任务 + 底部淡紫装饰条 + 透明背景 RGBA）。
2. 图13-2~13-6：对齐全书统一的结构图与架构图视觉语言（淡紫底色 #ECECFF + 中紫边框 #9370DB, 2px + 深灰正文 #333333 + 细灰连接箭头 #555555 + 优雅留白排版 + 透明背景 RGBA）。
"""

import os, math
from PIL import Image, ImageDraw, ImageFont

OUTPUT_DIR = "/tmp/p13_images"
os.makedirs(OUTPUT_DIR, exist_ok=True)

FONT_PATH = "/usr/local/lib/python3.11/dist-packages/mplfonts/fonts/NotoSansCJKsc-Regular.otf"

def get_font(size):
    return ImageFont.truetype(FONT_PATH, size)

# 统一出版级视觉调色板
BG_CARD = (236, 236, 255, 255)       # #ECECFF 淡紫底色
BORDER_COLOR = (147, 112, 219, 255)   # #9370DB 中紫边框
LINE_MUTED = (180, 160, 230, 255)     # #B4A0E6 卡片内分割线
TEXT_COLOR = (51, 51, 51, 255)        # #333333 深灰正文
TEXT_MUTED = (85, 85, 85, 255)        # #555555 辅助文字
ARROW_COLOR = (85, 85, 85, 255)       # #555555 箭头连接线
ACCENT_BLUE = (0, 0, 236, 255)        # #0000EC 教学导图中心圆球
ACCENT_BAR = (139, 139, 255, 255)     # #8B8BFF 导图底部装饰线

def draw_arrow(draw, x0, y0, x1, y1, color=ARROW_COLOR, width=2, arrow_size=8):
    """绘制带实心箭头的连接线"""
    draw.line([x0, y0, x1, y1], fill=color, width=width)
    dx = x1 - x0
    dy = y1 - y0
    length = math.hypot(dx, dy)
    if length == 0:
        return
    ux, uy = dx / length, dy / length
    vx, vy = -uy, ux
    ax0 = x1 - arrow_size * ux + (arrow_size * 0.5) * vx
    ay0 = y1 - arrow_size * uy + (arrow_size * 0.5) * vy
    ax1 = x1 - arrow_size * ux - (arrow_size * 0.5) * vx
    ay1 = y1 - arrow_size * uy - (arrow_size * 0.5) * vy
    draw.polygon([(x1, y1), (ax0, ay0), (ax1, ay1)], fill=color)

def draw_styled_box(draw, x0, y0, x1, y1, text, font, fill=BG_CARD, border=BORDER_COLOR, 
                    text_color=TEXT_COLOR, border_width=2, radius=4, align="center", line_spacing=4,
                    draw_divider_after=None):
    """绘制全书统一规范的淡紫底紫色细边框卡片"""
    draw.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=fill, outline=border, width=border_width)
    
    lines = text.split("\n")
    line_bboxes = [font.getbbox(l) if l else (0,0,0,0) for l in lines]
    line_heights = [b[3] - b[1] if l else 14 for l, b in zip(lines, line_bboxes)]
    total_h = sum(line_heights) + (len(lines) - 1) * line_spacing
    
    start_y = y0 + (y1 - y0 - total_h) // 2
    cur_y = start_y
    for i, line in enumerate(lines):
        if not line:
            cur_y += 14 + line_spacing
            continue
        bbox = line_bboxes[i]
        lw = bbox[2] - bbox[0]
        if align == "center":
            lx = x0 + (x1 - x0 - lw) // 2
        elif align == "left":
            lx = x0 + 14
        else:
            lx = x1 - lw - 14
        draw.text((lx, cur_y - bbox[1]), line, font=font, fill=text_color)
        cur_y += line_heights[i] + line_spacing
        
        if draw_divider_after is not None and i == draw_divider_after:
            div_y = cur_y - line_spacing // 2 + 2
            draw.line([(x0 + 12, div_y), (x1 - 12, div_y)], fill=LINE_MUTED, width=1)
            cur_y += 6

def draw_diamond(draw, cx, cy, w, h, text, font, fill=BG_CARD, border=BORDER_COLOR, text_color=TEXT_COLOR, border_width=2):
    """绘制菱形条件判断框"""
    pts = [
        (cx, cy - h // 2),
        (cx + w // 2, cy),
        (cx, cy + h // 2),
        (cx - w // 2, cy)
    ]
    draw.polygon(pts, fill=fill, outline=border)
    if border_width > 1:
        for i in range(4):
            draw.line([pts[i], pts[(i+1)%4]], fill=border, width=border_width)
            
    lines = text.split("\n")
    line_bboxes = [font.getbbox(l) if l else (0,0,0,0) for l in lines]
    line_heights = [b[3] - b[1] if l else 14 for l, b in zip(lines, line_bboxes)]
    total_h = sum(line_heights) + (len(lines) - 1) * 4
    start_y = cy - total_h // 2
    cur_y = start_y
    for i, line in enumerate(lines):
        if not line:
            cur_y += 14 + 4
            continue
        bbox = line_bboxes[i]
        lw = bbox[2] - bbox[0]
        lx = cx - lw // 2
        draw.text((lx, cur_y - bbox[1]), line, font=font, fill=text_color)
        cur_y += line_heights[i] + 4

# ==============================================================================
# 图13-1 项目13教学导图 (100% 对齐全书前12个项目思维导图)
# ==============================================================================
def generate_fig1():
    w, h = 1155, 490
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(im)
    
    font_center = get_font(15)
    font_task = get_font(12.5)
    font_sub = get_font(11)
    
    cx, cy = 577, 245
    cr = 52
    
    branches = [
        {
            "task": "任务13.1 财务多源异构数据清洗\n与多维中台整合",
            "subs": [
                "13.1.2 规则与大模型清洗协同",
                "13.1.1 鲜合多源财务数据特征",
                "13.1.3 搭建多维表格数据中台"
            ],
            "color": (255, 255, 120, 255),  # 黄色
            "accent": ACCENT_BAR,
            "tpos": (577, 130),
            "spos": [(577, 45), (300, 75), (855, 75)],
            "c_anchor": (cx, cy - cr)
        },
        {
            "task": "任务13.2 门店财务指标穿透\n与多维风险智能预警",
            "subs": [
                "13.2.1 12家门店杜邦与四象限",
                "13.2.2 03/08病灶门店业财穿透",
                "13.2.3 动因分析与红黄牌预警"
            ],
            "color": (215, 255, 134, 255),  # 绿色
            "accent": ACCENT_BAR,
            "tpos": (795, 245),
            "spos": [(1025, 175), (1035, 245), (1025, 315)],
            "c_anchor": (cx + cr, cy)
        },
        {
            "task": "任务13.3 交互式BI驾驶舱搭建\n与数据叙事白皮书",
            "subs": [
                "13.3.2 财务指标联动下钻卡片",
                "13.3.1 高管全局与门店自查设计",
                "13.3.3 金字塔原理白皮书生成"
            ],
            "color": (195, 134, 255, 255),  # 紫色
            "accent": ACCENT_BAR,
            "tpos": (577, 360),
            "spos": [(577, 448), (320, 432), (835, 432)],
            "c_anchor": (cx, cy + cr)
        },
        {
            "task": "任务13.4 端到端工作流自动化编排\n与多智能体协同",
            "subs": [
                "13.4.3 自动化推送与Streamlit轻应用",
                "13.4.2 分析师-风控官-顾问协同",
                "13.4.1 月度财务分析工作流拓扑"
            ],
            "color": (255, 134, 255, 255),  # 粉色
            "accent": ACCENT_BAR,
            "tpos": (355, 245),
            "spos": [(125, 175), (115, 245), (125, 315)],
            "c_anchor": (cx - cr, cy)
        }
    ]
    
    def draw_mindmap_box(bx, by, text, font, fill_color, accent_color, pad_x=10, pad_y=5, r=5):
        lines = text.split("\n")
        line_bboxes = [font.getbbox(l) for l in lines]
        max_w = max(b[2] - b[0] for b in line_bboxes)
        total_text_h = sum(b[3] - b[1] for b in line_bboxes) + (len(lines) - 1) * 4
        bw = max_w + pad_x * 2
        bh = total_text_h + pad_y * 2 + 3
        x0 = bx - bw // 2
        y0 = by - bh // 2
        x1 = x0 + bw
        y1 = y0 + bh
        draw.rounded_rectangle([x0, y0, x1, y1], radius=r, fill=fill_color)
        draw.rectangle([x0 + 2, y1 - 3, x1 - 2, y1], fill=accent_color)
        
        cur_y = y0 + pad_y
        for i, line in enumerate(lines):
            bbox = line_bboxes[i]
            lw = bbox[2] - bbox[0]
            lx = bx - lw // 2
            draw.text((lx, cur_y - bbox[1]), line, font=font, fill=(0, 0, 0, 255))
            cur_y += (bbox[3] - bbox[1]) + 4
        return (x0, y0, x1, y1)

    # 1. 绘制放射主干与支线
    for b in branches:
        color = b["color"]
        tx, ty = b["tpos"]
        ax, ay = b["c_anchor"]
        draw.line([ax, ay, tx, ty], fill=color, width=4)
        for sx, sy in b["spos"]:
            draw.line([tx, ty, sx, sy], fill=color, width=2)
            
    # 2. 绘制子任务节点
    for b in branches:
        for i, sub in enumerate(b["subs"]):
            sx, sy = b["spos"][i]
            draw_mindmap_box(sx, sy, sub, font_sub, b["color"], b["accent"], pad_x=8, pad_y=4, r=4)
            
    # 3. 绘制主任务节点
    for b in branches:
        tx, ty = b["tpos"]
        draw_mindmap_box(tx, ty, b["task"], font_task, b["color"], b["accent"], pad_x=12, pad_y=5, r=6)

    # 4. 绘制中心宝蓝色圆球
    draw.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], fill=ACCENT_BLUE)
    lines = ["项目13", "财务AI智能体", "全流程综合实战"]
    total_h = len(lines) * 19
    start_y = cy - total_h // 2
    for i, line in enumerate(lines):
        bbox = font_center.getbbox(line)
        lw = bbox[2] - bbox[0]
        lx = cx - lw // 2
        ly = start_y + i * 19 - bbox[1]
        draw.text((lx, ly), line, font=font_center, fill=(255, 255, 255, 255))
        
    out_path = os.path.join(OUTPUT_DIR, "image1.png")
    im.save(out_path)
    print("Generated:", out_path)

# ==============================================================================
# 图13-2 鲜合餐饮多源财务数据清洗与整合架构图 (全书统一淡紫卡片架构)
# ==============================================================================
def generate_fig2():
    w, h = 1568, 800
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(im)
    
    font_h = get_font(18)
    font_b = get_font(13.5)
    
    col_w = 440
    c1_x, c2_x, c3_x = 60, 564, 1068
    
    draw_styled_box(draw, c1_x, 30, c1_x + col_w, 80, "多源原始异构数据输入层", font_h, radius=6)
    draw_styled_box(draw, c2_x, 30, c2_x + col_w, 80, "规则与大模型清洗治理层", font_h, radius=6)
    draw_styled_box(draw, c3_x, 30, c3_x + col_w, 80, "飞书多维表格中台数据资产", font_h, radius=6)
    
    rows = [
        (
            "POS前台销售明细流水\n• 聚合微信/支付宝/银行卡交易流水\n• 记录每笔点单明细、实收与折扣",
            "【实体归一治理】\n• 模糊门店名称归一至12家标准编号\n• 统一核算代码与业财口径字典",
            "门店主数据字典表\n• 标准12家直营门店编号及区域属性\n• 建立各业务表强关联主键"
        ),
        (
            "外卖平台月度对账单\n• 美团/饿了么满减补贴明细\n• 扣除配送服务费与平台扣佣净额",
            "【数值合规拦截】\n• 校验流水金额正负号，拦截异常非退单负数\n• 识别客单价>5000元等畸大离群值",
            "统一销售流水清洗表\n• 规整线上线下全渠道经营流水\n• 自动核销外卖佣金与配送费成本"
        ),
        (
            "本地生鲜食材采购专票\n• 属地直采农产品生鲜增值税专票\n• 法定农产品进项税率 9%",
            "【专票税率验真】\n• 算法校验金额×法定税率是否等于税额\n• 拦截发票代码失真与税率错配",
            "食材BOM消耗跟踪表\n• 匹配标准菜品BOM与实际消耗量\n• 输出各单店肉类及生鲜损耗差额"
        ),
        (
            "中央厨房统配半成品专票\n• 统配半成品、底料调味品专票\n• 法定增值税率 13%",
            "【防重过滤机制】\n• 设定3天时间滑动窗口\n• 拦截同商户、同金额、同发票号重复报销",
            "12家门店标准化经营宽表\n• 汇聚营收、成本、净利、翻台全维数据\n• 统一支撑下游杜邦分析与风险预警"
        ),
        (
            "门店日常费用报销台账\n• 门店租金、水电、员工薪酬\n• 营销推广与办公杂费发票及凭单",
            "【业财对齐核销】\n• 自动化比对平台对账净额与银行实收\n• 确保业财数据100%账实相符",
            "下游综合分析服务层\n• 直通杜邦指标体系\n• 驱动BI驾驶舱与自动化工作流"
        )
    ]
    
    start_y = 110
    row_h = 115
    gap_y = 20
    
    for idx, (col1_text, col2_text, col3_text) in enumerate(rows):
        y0 = start_y + idx * (row_h + gap_y)
        y1 = y0 + row_h
        
        draw_styled_box(draw, c1_x, y0, c1_x + col_w, y1, col1_text, font_b, align="left")
        draw_styled_box(draw, c2_x, y0, c2_x + col_w, y1, col2_text, font_b, align="left")
        draw_styled_box(draw, c3_x, y0, c3_x + col_w, y1, col3_text, font_b, align="left")
        
        mid_y = (y0 + y1) // 2
        draw_arrow(draw, c1_x + col_w + 5, mid_y, c2_x - 5, mid_y, ARROW_COLOR, width=2)
        draw_arrow(draw, c2_x + col_w + 5, mid_y, c3_x - 5, mid_y, ARROW_COLOR, width=2)

    out_path = os.path.join(OUTPUT_DIR, "image2.png")
    im.save(out_path)
    print("Generated:", out_path)

# ==============================================================================
# 图13-3 鲜合餐饮12家直营门店经营效益四象限波士顿矩阵 (全书统一矩阵架构)
# ==============================================================================
def generate_fig3():
    w, h = 1568, 920
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(im)
    
    font_title = get_font(20)
    font_sub = get_font(14)
    font_h = get_font(16)
    font_b = get_font(13.2)
    
    draw_styled_box(draw, 404, 25, 1164, 75, "鲜合餐饮全国12家直营门店经营效益四象限波士顿矩阵", font_title, radius=6)
    draw_styled_box(draw, 434, 85, 1134, 125, "基准均值：门店日均翻台率 3.4 次/天  |  连锁控本警戒线：食材成本率 38.0%", font_sub, radius=4)
    
    quad_w = 705
    quad_h = 360
    left_x = 60
    right_x = 803
    top_y = 150
    bottom_y = 530
    
    # 1. 左上：【瘦狗象限】
    q_top_left = (
        "【瘦狗象限】低翻台率 (<3.4次)  /  高食材成本率 (>38.0%)\n"
        "核心特征：客流严重低迷且固定租金沉重，陷入持续经营亏损病灶\n"
        "★ 典型问题门店：08号 上海淮海中路店 (全连锁唯一亏损门店)\n"
        "• 核心经营指标：日均翻台 1.8 次/天 | 食材成本率 43.0% (超标5.0%) | 半年净利 -4.1 万元\n"
        "• 核心病灶诊断：\n"
        "  - 固定租金及折旧占总营收比重高达 37.0%，远超保本点翻台率要求 (保本需 2.4次/天)；\n"
        "  - 传统核心商圈客流外溢分散，低翻台导致食材周转放缓，进一步加剧食材损耗报损；\n"
        "• 刚性应对举措：限期45天向业主申请降租30%，若谈判无果则启动闭店止损流程。"
    )
    draw_styled_box(draw, left_x, top_y, left_x + quad_w, top_y + quad_h, q_top_left, font_b, align="left", draw_divider_after=1)
    
    # 2. 右上：【问题象限】
    q_top_right = (
        "【问题象限】高翻台率 (≥3.4次)  /  高食材成本率 (>38.0%)\n"
        "核心特征：客流高度旺盛但利润被严重蚕食，供应链与损耗管理失控\n"
        "★ 典型病灶门店：03号 北京朝阳大悦城店 (高损耗吞噬利润)\n"
        "• 核心经营指标：日均翻台 3.8 次/天 (高客流) | 食材成本率 44.0% (超标6.0%) | 半年净利 23.2 万元\n"
        "• 核心病灶诊断：\n"
        "  - 核心肉类食材实际消耗偏离标准菜品BOM达到 +8.5%，月度超额损耗达12.6万元；\n"
        "  - 属地生鲜供应商批次抽检缺失，库房冷链损耗失控，会员复购率断崖下滑至 21.0%；\n"
        "• 刚性应对举措：冻结本地生鲜直采，改由集团央厨集配；现场清点冷库并整改后厨操作。"
    )
    draw_styled_box(draw, right_x, top_y, right_x + quad_w, top_y + quad_h, q_top_right, font_b, align="left", draw_divider_after=1)
    
    # 3. 左下：【稳健金牛象限】
    q_bottom_left = (
        "【稳健金牛象限】翻台适中 (<3.4次)  /  控本健康 (≤38.0%)\n"
        "核心特征：客流与成本控制高度均衡，为全连锁提供坚实基石现金流\n"
        "★ 涵盖门店集群：天津和平、天津南开、南京新街口、苏州金鸡湖等6家门店\n"
        "• 核心经营指标：\n"
        "  - 门店日均翻台：2.9 ~ 3.5 次/天 (平均 3.1 次/天) | 食材成本率：36.0% ~ 37.0% (规范达标)\n"
        "  - 门店半年净利：15.4 ~ 24.4 万元 (平均 20.8 万元，经营净利率约 15.2%)\n"
        "• 战略价值分析：\n"
        "  - 区域次级商圈布局合理，固定租金负担轻，后厨严格执行标准BOM，损耗率可控；\n"
        "  - 经营抗风险能力极强，为总部数字化升级与扩张持续提供稳固充沛的基石资金。"
    )
    draw_styled_box(draw, left_x, bottom_y, left_x + quad_w, bottom_y + quad_h, q_bottom_left, font_b, align="left", draw_divider_after=1)
    
    # 4. 右下：【明星象限】
    q_bottom_right = (
        "【明星象限】高翻台率 (≥3.4次)  /  控本健康 (≤38.0%)\n"
        "核心特征：高客流与精细化供应链双轮驱动，全连锁盈利领跑标杆\n"
        "★ 领跑标杆门店：06号 上海南京东路店 & 01号 北京王府井店\n"
        "• 标杆经营指标：\n"
        "  - 06店：翻台 4.5 次/天 (全连锁第一) | 食材率 36.0% | 半年净利润 51.7 万元 (贡献超标)\n"
        "  - 01店：翻台 4.2 次/天 (华北标杆) | 食材率 36.0% | 半年净利润 37.3 万元\n"
        "  - 梯队门店：07号陆家嘴店 (净利33.5万) | 10号西湖湖滨店 (净利30.1万)\n"
        "• 战略价值分析：\n"
        "  - 规模效应显著，全流程执行数字化生鲜锁价与精细排班，提供标准化运营示范模板。"
    )
    draw_styled_box(draw, right_x, bottom_y, right_x + quad_w, bottom_y + quad_h, q_bottom_right, font_b, align="left", draw_divider_after=1)

    out_path = os.path.join(OUTPUT_DIR, "image3.png")
    im.save(out_path)
    print("Generated:", out_path)

# ==============================================================================
# 图13-4 鲜合餐饮高管经营分析驾驶舱与门店端双重视角BI交互看板 (全书统一架构)
# ==============================================================================
def generate_fig4():
    w, h = 1568, 880
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(im)
    
    font_title = get_font(20)
    font_h = get_font(16)
    font_b = get_font(13.5)
    
    draw_styled_box(draw, 404, 20, 1164, 70, "鲜合餐饮高管全局与门店自查双重视角BI交互看板体系", font_title, radius=6)
    
    kpi_w = 345
    kpi_gap = 24
    start_x = 60
    kpis = [
        ("全连锁营收总规模\n1,694.4 万元 (同比+14.2%)\n华北 765.2万 | 华东 929.2万"),
        ("全连锁经营净利润\n271.1 万元 (净利率 16.0%)\n11家盈利 / 1家单店亏损"),
        ("全连锁食材成本率\n38.0 % (警戒线 40.0%)\n标杆 36.0% / 异常超标 44.0%"),
        ("全连锁平均翻台率\n3.4 次/天 (保本点 2.4次)\n标杆 4.5次 / 最低 1.8次")
    ]
    for i, k_text in enumerate(kpis):
        kx = start_x + i * (kpi_w + kpi_gap)
        draw_styled_box(draw, kx, 85, kx + kpi_w, 175, k_text, font_b, radius=4)
        
    panel_w = 710
    panel_h = 490
    p1_x = 60
    p2_x = 798
    p_y = 195
    
    draw_styled_box(draw, p1_x, p_y, p1_x + panel_w, p_y + 45, "【高管全局经营驾驶舱】(宏观决策与对比视角)", font_h, radius=4)
    p1_items = [
        "12家直营门店营收与净利润全景排行榜\n• 标杆梯队：06南京东路(51.7万)、01王府井(37.3万)、07陆家嘴(33.5万)\n• 稳健梯队：10西湖(30.1万)、09武林(24.4万)、02中关村(22.9万)、11新街口(21.6万)\n• 关注梯队：04和平(20.2万)、05南开(15.8万)、12金鸡湖(15.4万)\n• 预警与亏损：03大悦城(23.2万/损耗高)、08淮海(-4.1万/亏损)",
        "华东 vs 华北两大区域经营效益对比分析\n• 华东区域(6店)：营收 929.2万，净利 150.2万，平均食材率 37.8%\n• 华北区域(6店)：营收 765.2万，净利 120.9万，平均食材率 38.2%\n• 华东高客流带动整体毛利，华北需强化单店供应链精细化控本",
        "全连锁经营成本瀑布图深度拆解\n• 食材成本 38.0% (643.9万) | 员工薪酬 26.0% (440.5万)\n• 租金折旧 18.0% (305.0万) | 营销及其他费用 2.0% (33.9万)\n• 综合经营净利润率保持在 16.0% 稳健区间"
    ]
    cur_y = p_y + 55
    sub_h = [135, 120, 115]
    for i, itm in enumerate(p1_items):
        draw_styled_box(draw, p1_x, cur_y, p1_x + panel_w, cur_y + sub_h[i], itm, font_b, align="left")
        cur_y += sub_h[i] + 12
        
    draw_styled_box(draw, p2_x, p_y, p2_x + panel_w, p_y + 45, "【门店自查与穿透看板】(微观运营与病灶剖析视角)", font_h, radius=4)
    p2_items = [
        "单店经营体检雷达图与BOM偏离度实时监测\n• 监控维度：翻台率、客单价、食材成本率、人效坪效、会员复购率\n• 正常门店浮动范围：实际食材BOM偏差处于 ±1.5% 合理区间\n• 系统自动拦截食材偏离度 >3.0% 的异常单店并点亮预警指示灯",
        "03号朝阳大悦城店病灶穿透下钻卡片\n• 核心肉类消耗偏离BOM +8.5%，月度食材超标12.6万元\n• 会员复购率断崖下滑至 21.0% (全连锁均值 32.6%)\n• 动因归因：本地生鲜供应商批次质检缺位，菜品品质不稳定引发客诉",
        "08号上海淮海中路店保本点与租金盈亏穿透\n• 月租金及折旧占营收比重高达 37.0%，当前翻台仅 1.8次/天\n• 静态测算保本点翻台率需达 2.4次/天，月均净亏损 0.68万元\n• 动因归因：商圈分流与高租金错配，面临严峻现金流断流风险"
    ]
    cur_y = p_y + 55
    for i, itm in enumerate(p2_items):
        draw_styled_box(draw, p2_x, cur_y, p2_x + panel_w, cur_y + sub_h[i], itm, font_b, align="left")
        cur_y += sub_h[i] + 12
        
    bottom_y = 705
    b_text = (
        "【双重视角看板交互联动与数据叙事闭环机制】\n"
        "1. 联动下钻：高管全局看板点击异常门店(03店/08店) -> 仪表盘自动切入门店端微观穿透视图并高亮异常动因卡片；\n"
        "2. 智能叙事：点击【一键诊断】按钮 -> AI智能体遵循“情境-冲突-解答”结构自动撰写月度经营分析白皮书；\n"
        "3. 闭环执行：白皮书刚性输出03店BOM直采核查与08店租金重谈判方案，实现“看盘-诊断-决策”完整闭环。"
    )
    draw_styled_box(draw, 60, bottom_y, 1508, bottom_y + 145, b_text, font_b, align="left")
    
    draw_arrow(draw, 784, 175, 784, 195, ARROW_COLOR, width=2)
    draw_arrow(draw, 784, 685, 784, 705, ARROW_COLOR, width=2)

    out_path = os.path.join(OUTPUT_DIR, "image4.png")
    im.save(out_path)
    print("Generated:", out_path)

# ==============================================================================
# 图13-5 鲜合餐饮端到端自动化财务分析工作流拓扑编排图 (对齐项目12工作流风格)
# ==============================================================================
def generate_fig5():
    w, h = 1568, 920
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(im)
    
    font_h = get_font(16)
    font_b = get_font(13.5)
    
    node_w = 800
    cx = w // 2
    x0 = cx - node_w // 2
    x1 = cx + node_w // 2
    
    # 1. 开始节点
    y_start = 30
    h1 = 70
    draw_styled_box(draw, x0, y_start, x1, y_start + h1, 
                    "开始节点：【定时监听与事件触发器】\n每月1日 08:00 自动触发执行，监听12家直营门店月度财务流水包归档完成", 
                    font_b, radius=6)
    
    draw_arrow(draw, cx, y_start + h1, cx, y_start + h1 + 30)
    
    # 2. 循环批处理代码节点
    y2 = y_start + h1 + 30
    h2 = 75
    draw_styled_box(draw, x0, y2, x1, y2 + h2, 
                    "代码节点：【多源异构数据循环批处理与治理】\n遍历12家门店数据，自动完成实体归一、9%/13%专票乘积验真与发票防重过滤", 
                    font_b, radius=6)
    
    draw_arrow(draw, cx, y2 + h2, cx, y2 + h2 + 30)
    
    # 3. 大模型计算节点
    y3 = y2 + h2 + 30
    h3 = 75
    draw_styled_box(draw, x0, y3, x1, y3 + h3, 
                    "大模型节点：【杜邦指标拆解与标准化宽表构建】\n核算12家单店翻台率、食材成本率、毛利率与经营净利，生成标准化经营宽表资产", 
                    font_b, radius=6)
    
    draw_arrow(draw, cx, y3 + h3, cx, y3 + h3 + 30)
    
    # 4. 条件分支菱形
    y4 = y3 + h3 + 30 + 65
    dia_w = 560
    dia_h = 130
    draw_diamond(draw, cx, y4, dia_w, dia_h, 
                 "条件分支节点：【双重风险规则分流】\n门店食材成本率 > 40.0% ？\n或 门店半年度经营净利润 < 0 万元 ？", 
                 font_b)
    
    # 左右两分支
    y5 = y4 + dia_h // 2 + 50
    branch_w = 640
    branch_h = 135
    b1_x0 = 80
    b1_x1 = b1_x0 + branch_w
    draw_styled_box(draw, b1_x0, y5, b1_x1, y5 + branch_h,
                    "多智能体协同节点：【多角色深度会诊与归因穿透】\n"
                    "• 财务分析师：客观提取03店食材率44%与08店亏损4.1万异动数据\n"
                    "• 财务风控官：审慎质询03店BOM损耗+8.5%与08店现金流断流风险\n"
                    "• 管理顾问：刚性仲裁整改对策，生成“情境-冲突-解答”白皮书",
                    font_b, align="left")
    
    b2_x1 = 1488
    b2_x0 = b2_x1 - branch_w
    draw_styled_box(draw, b2_x0, y5, b2_x1, y5 + branch_h,
                    "大模型节点：【标准月度经营简报自动化生成】\n"
                    "• 汇总10家稳健门店常态化杜邦指标与翻台达标情况\n"
                    "• 自动提炼华东/华北区域标杆门店(06店/01店)经营亮点\n"
                    "• 输出标准版月度门店经营跟踪简报与管理归档底稿",
                    font_b, align="left")
    
    draw.line([(cx - dia_w // 2, y4), (b1_x0 + branch_w // 2, y4)], fill=ARROW_COLOR, width=2)
    draw_arrow(draw, b1_x0 + branch_w // 2, y4, b1_x0 + branch_w // 2, y5, ARROW_COLOR, width=2)
    draw.text((cx - dia_w // 2 - 130, y4 - 24), "是 (异常门店)", font=font_b, fill=TEXT_COLOR)
    
    draw.line([(cx + dia_w // 2, y4), (b2_x0 + branch_w // 2, y4)], fill=ARROW_COLOR, width=2)
    draw_arrow(draw, b2_x0 + branch_w // 2, y4, b2_x0 + branch_w // 2, y5, ARROW_COLOR, width=2)
    draw.text((cx + dia_w // 2 + 25, y4 - 24), "否 (稳健门店)", font=font_b, fill=TEXT_COLOR)
    
    y6 = y5 + branch_h + 45
    h6 = 80
    draw_styled_box(draw, x0, y6, x1, y6 + h6,
                    "输出与推送节点：【自动化多渠道成果发布与通知】\n"
                    "• 将诊断白皮书渲染为标准Markdown与PDF格式，生成Excel穿透底稿\n"
                    "• 通过飞书Webhook机器人向集团高管群推送预警通知，同步抄送审计邮箱",
                    font_b, radius=6)
    
    mid_y6 = y5 + branch_h + 20
    draw.line([(b1_x0 + branch_w // 2, y5 + branch_h), (b1_x0 + branch_w // 2, mid_y6)], fill=ARROW_COLOR, width=2)
    draw.line([(b1_x0 + branch_w // 2, mid_y6), (cx - 100, mid_y6)], fill=ARROW_COLOR, width=2)
    draw_arrow(draw, cx - 100, mid_y6, cx - 100, y6, ARROW_COLOR, width=2)
    
    draw.line([(b2_x0 + branch_w // 2, y5 + branch_h), (b2_x0 + branch_w // 2, mid_y6)], fill=ARROW_COLOR, width=2)
    draw.line([(b2_x0 + branch_w // 2, mid_y6), (cx + 100, mid_y6)], fill=ARROW_COLOR, width=2)
    draw_arrow(draw, cx + 100, mid_y6, cx + 100, y6, ARROW_COLOR, width=2)
    
    y7 = y6 + h6 + 30
    h7 = 45
    draw_arrow(draw, cx, y6 + h6, cx, y7)
    draw_styled_box(draw, cx - 250, y7, cx + 250, y7 + h7,
                    "结束节点：月度自动化财务分析与预警流水线执行完毕",
                    font_b, radius=6)

    out_path = os.path.join(OUTPUT_DIR, "image5.png")
    im.save(out_path)
    print("Generated:", out_path)

# ==============================================================================
# 图13-6 鲜合餐饮多智能体协同会诊交互与轻应用展示界面 (全书统一架构)
# ==============================================================================
def generate_fig6():
    w, h = 1568, 920
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(im)
    
    font_title = get_font(20)
    font_h = get_font(16)
    font_b = get_font(13.2)
    
    draw_styled_box(draw, 404, 25, 1164, 75, "多智能体协同会诊运行机制 (三角色分歧碰撞与仲裁收敛)", font_title, radius=6)
    
    box_w = 440
    box_h = 240
    b1_x = 60
    b2_x = 564
    b3_x = 1068
    top_y = 95
    
    r1_text = (
        "【角色1：财务分析师】(客观数据提取)\n"
        "• 核心职责：客观陈述数据异动，构建事实底座\n"
        "• 03店异动：食材成本率44.0%，超标损耗12.6万\n"
        "• 08店异动：租金折旧占营收37.0%，净亏损4.1万\n"
        "• 行为准则：严禁主观推测原因，仅输出量化指标\n"
        "• 协同传递：将异常宽表与指标事实推送至风控官"
    )
    r2_text = (
        "【角色2：财务风控官】(合规审慎质询)\n"
        "• 核心职责：审慎怀疑与合规挑刺，揭示潜在风险\n"
        "• 质询03店：肉类实际消耗偏离BOM +8.5%，库房失控\n"
        "• 警示08店：1.8次翻台无法覆盖保本点，现金流断流\n"
        "• 索取证据：刚性要求属地门店提供进销存单据\n"
        "• 协同传递：将合规疑点与风险清单呈报咨询顾问"
    )
    r3_text = (
        "【角色3：管理咨询顾问】(矛盾仲裁闭环)\n"
        "• 核心职责：矛盾仲裁与刚性闭环，平衡控本与拓客\n"
        "• 仲裁03店：冻结生鲜直采，央厨集配，激活复购\n"
        "• 仲裁08店：限期45天申请降租30%，若无果闭店\n"
        "• 闭环收敛：最终收敛为金字塔原理经营诊断白皮书\n"
        "• 成果交付：落地整改台账与责任人跟进清单"
    )
    
    draw_styled_box(draw, b1_x, top_y, b1_x + box_w, top_y + box_h, r1_text, font_b, align="left", draw_divider_after=0)
    draw_styled_box(draw, b2_x, top_y, b2_x + box_w, top_y + box_h, r2_text, font_b, align="left", draw_divider_after=0)
    draw_styled_box(draw, b3_x, top_y, b3_x + box_w, top_y + box_h, r3_text, font_b, align="left", draw_divider_after=0)
    
    arrow_y = top_y + box_h // 2
    draw_arrow(draw, b1_x + box_w + 5, arrow_y, b2_x - 5, arrow_y, ARROW_COLOR, width=2)
    draw_arrow(draw, b2_x + box_w + 5, arrow_y, b3_x - 5, arrow_y, ARROW_COLOR, width=2)
    
    draw_arrow(draw, 784, top_y + box_h + 10, 784, top_y + box_h + 40, ARROW_COLOR, width=2)
    
    app_y = top_y + box_h + 45
    draw_styled_box(draw, 404, app_y, 1164, app_y + 45, "Streamlit 财务智能体轻应用原型架构 (http://localhost:8501)", font_h, radius=4)
    
    app_content_y = app_y + 55
    side_w = 440
    main_w = 948
    app_h = 440
    
    side_text = (
        "【侧边栏动态参数控制区】\n"
        "1. 门店选择下拉控件 (st.selectbox)\n"
        "   • 快速切换全连锁12家直营门店\n"
        "   • 默认选中：03号北京朝阳大悦城店\n\n"
        "2. 食材成本率警戒阈值滑块 (st.slider)\n"
        "   • 调节范围：30.0% ~ 50.0%\n"
        "   • 默认警戒阈值：40.0% (超标点亮红牌)\n\n"
        "3. 日均翻台率保本警戒线滑块 (st.slider)\n"
        "   • 调节范围：1.0 ~ 5.0 次/天\n"
        "   • 默认保本线：3.0 次/天 (不足提示客流风险)\n\n"
        "4. 【一键启动多智能体会诊】按钮 (st.button)\n"
        "   • 异步并发调用三角色大模型编排链路\n"
        "   • 实时以流式会话形式呈现各方交锋"
    )
    draw_styled_box(draw, 60, app_content_y, 60 + side_w, app_content_y + app_h, side_text, font_b, align="left", draw_divider_after=0)
    
    main_text = (
        "【主展示窗口交互流与诊断成果看板】\n"
        "1. 关键经营KPI卡片看板 (st.metric)\n"
        "   • 门店总营收：210.5 万元   |   半年度经营净利：23.2 万元 (偏低!)\n"
        "   • 食材成本率：44.0 % (超标6.0%! 红色警戒)   |   日均翻台率：3.8 次/天 (客流充沛)\n\n"
        "2. 智能体会诊交互对话窗口流 (st.chat_message)\n"
        "   • [财务分析师] 03店6月份牛肉采购均价同比上涨14.2%，实际报损率达4.8%，远超连锁2.0%容忍度！\n"
        "   • [财务风控官] 质询：5月至6月生鲜供应商送货单缺失验收签名，疑似存在虚假入库与以次充好！\n"
        "   • [管理咨询顾问] 裁定整改方案：\n"
        "     1. 供应链部立即冻结该属地生鲜供应商合作，切换为中央厨房冷链集配；\n"
        "     2. 财务内审组即日入驻03店现场清点冷库实物资产，追查虚报损耗责任；\n"
        "     3. 运营部优化周末二人惠食促销套餐，改善客诉并激活会员复购率；\n"
        "     4. 设立两周观察期，刚性要求食材成本率回落至 38.0% 健康线以内。\n\n"
        "3. 成果自动化导出与闭环发布 (st.download_button)\n"
        "   • 【导出月度经营诊断白皮书 Markdown / PDF】   |   【下载多源业财清洗底稿 Excel】"
    )
    draw_styled_box(draw, 564, app_content_y, 564 + main_w, app_content_y + app_h, main_text, font_b, align="left", draw_divider_after=0)

    out_path = os.path.join(OUTPUT_DIR, "image6.png")
    im.save(out_path)
    print("Generated:", out_path)

if __name__ == "__main__":
    generate_fig1()
    generate_fig2()
    generate_fig3()
    generate_fig4()
    generate_fig5()
    generate_fig6()
    print("All 6 figures for Project 13 generated successfully in", OUTPUT_DIR)
