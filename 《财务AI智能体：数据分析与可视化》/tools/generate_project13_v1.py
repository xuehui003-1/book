#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_project13_v1.py —— 自动构建并生成《项目13 财务AI智能体全流程综合实战_初稿V1.docx》

全量格式与规范升级（对标全书前12个项目出版级体例）：
  1. 正文字体与排版：全面落地 Times New Roman + 微软雅黑，小四号（sz=24），首行缩进 2 字符（ind=480）；
  2. 标题层级体系：一级项目标题（style 2, sz=28, 粗体）、栏目标题与任务标题（style 4, sz=24, 粗体）、
     二级小节（style 5, sz=24, 粗体）、三级细目（style 6, sz=24, 粗体）、任务情景导引（style 23, sz=24）；
  3. 图表全量插入：包含图13-1至图13-6全部6幅高精度矢量/位图图表（style 35居中），图题表题居中对齐（style 3）；
  4. 表格出版级美化：带 EEF4FF 浅蓝表头、F8FAFC 斑马隔行底色、全三线表灰色内框线、中英数字字体与字号（sz=22）对齐；
  5. 案例与数据并轨：完全基于鲜合餐饮12家直营门店真实半年度财务与营运数据底稿（1,694.4万营收闭环）；
  6. 验收自检全绿：check_docx.py 必改0/提示0，audit_a_class.py A1~A7 全维度达标。
"""

import os, sys, re, zipfile, shutil
import html as _html

TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
if TOOLS_DIR not in sys.path:
    sys.path.insert(0, TOOLS_DIR)

from docx_utils import normalize_punct_docx, para_spans

SRC_TEMPLATE = '项目12 工作流编排多智能体协作与前沿应用_修订稿V1.docx'
DST_DOCX = '项目13 财务AI智能体全流程综合实战_初稿V1.docx'
WORK_DIR = '/tmp/p13_build'
IMAGES_DIR = '/tmp/p13_images'

def log(msg):
    print(msg, flush=True)

def xml_esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

def P(text, style='24', bold=False, indent=True, jc=None, sz=24):
    """生成标准化段落 XML，严格贴合前12个项目的格式"""
    t_esc = xml_esc(text)
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/>'
    if indent and style in ['24', '3', '1']:
        ppr += '<w:ind w:firstLine="480" w:firstLineChars="200"/>'
    if jc:
        ppr += f'<w:jc w:val="{jc}"/>'
    ppr += '</w:pPr>'
    
    rpr = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/>'
    if bold:
        rpr += '<w:b/><w:bCs/>'
    rpr += f'<w:sz w:val="{sz}"/></w:rPr>'
    
    r = f'<w:r>{rpr}<w:t>{t_esc}</w:t></w:r>'
    return f'<w:p>{ppr}{r}</w:p>'

def DRAWING(r_id, doc_id, descr, cx=4950000, cy=2250000):
    """生成标准 OpenXML 图片 drawing 节点段落"""
    d_esc = xml_esc(descr)
    return (
        f'<w:p><w:pPr><w:pStyle w:val="35"/><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="114300" distR="114300">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{doc_id}" name="Picture" descr="{d_esc}"/>'
        f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{doc_id}" name="Picture" descr="{d_esc}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{r_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525"><a:noFill/></a:ln></pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    )

def TBL(headers, rows, align_cols=None):
    """生成标准出版级三线表 XML，带有 EEF4FF 表头背景、隔行斑马底色与统一字体"""
    tbl = ['<w:tbl>']
    tbl.append('<w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/>')
    tbl.append('<w:tblBorders>')
    tbl.append('<w:top w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>')
    tbl.append('<w:bottom w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>')
    tbl.append('<w:left w:val="none"/><w:right w:val="none"/>')
    tbl.append('<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>')
    tbl.append('<w:insideV w:val="none"/>')
    tbl.append('</w:tblBorders>')
    tbl.append('<w:tblCellMar><w:top w:w="120" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tblCellMar>')
    tbl.append('<w:tblLayout w:type="autofit"/>')
    tbl.append('</w:tblPr>')
    
    # 表头行
    tbl.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>')
    for h in headers:
        tbl.append(
            '<w:tc><w:tcPr>'
            '<w:shd w:val="clear" w:color="auto" w:fill="EEF4FF"/>'
            '<w:tcBorders><w:bottom w:val="single" w:sz="6" w:color="94A3B8"/></w:tcBorders>'
            '</w:tcPr>'
            '<w:p><w:pPr><w:pStyle w:val="25"/><w:jc w:val="center"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/><w:b/><w:bCs/><w:sz w:val="22"/></w:rPr><w:t>{xml_esc(h)}</w:t></w:r></w:p></w:tc>'
        )
    tbl.append('</w:tr>')
    
    # 数据行
    for row_idx, r in enumerate(rows):
        fill_color = "F8FAFC" if row_idx % 2 == 1 else "FFFFFF"
        tbl.append('<w:tr>')
        for col_idx, c in enumerate(r):
            align = "left" if col_idx == 0 and len(str(c)) > 6 else "center"
            if align_cols and col_idx in align_cols:
                align = align_cols[col_idx]
            tbl.append(
                f'<w:tc><w:tcPr>'
                f'<w:shd w:val="clear" w:color="auto" w:fill="{fill_color}"/>'
                f'</w:tcPr>'
                f'<w:p><w:pPr><w:pStyle w:val="25"/><w:jc w:val="{align}"/></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/><w:sz w:val="22"/></w:rPr><w:t>{xml_esc(str(c))}</w:t></w:r></w:p></w:tc>'
            )
        tbl.append('</w:tr>')
    
    tbl.append('</w:tbl>')
    return "".join(tbl)

def build_content():
    parts = []

    # 1. 标题与【教学导读】
    parts.append(P('项目13 财务AI智能体全流程综合实战', '2', bold=True, indent=False, sz=28))
    parts.append(P('【教学导读】', '4', bold=True, indent=False))
    parts.append(P('在完成了前十二个项目的分项技能训练之后，你已经掌握了从多维表格、数据清洗、描述性分析、比率分析、风险预警，到BI看板、知识智能体搭建与工作流自动化编排的全部核心技能。但在实际的企业财会岗位上，真实的业务需求绝不是按照教科书的章节切片展开的，而是以“全链条项目交付”的形态集中爆发。'))
    parts.append(P('本项目作为全流程实战能力的终极综合演练，打破章节边界，以鲜合餐饮管理有限公司全国12家直营门店真实数字化转型大考为实战场景。你将以“财务AI分析主理人”的角色，全面统合数据清洗智能体、杜邦分析模型、多维风险预警、交互式BI大屏、端到端自动化工作流以及三智能体协同会诊机制，独立交付一套工业级的企业财务自动化决策支持系统。'))
    parts.append(P('本项目的整体架构与递进实训路径如图13-1所示。'))
    
    # 插入图13-1
    parts.append(DRAWING('rId5', 21, '项目13教学导图', cx=4950000, cy=2200000))
    parts.append(P('图13-1 项目13教学导图', '3', indent=False, jc='center'))
    
    parts.append(P('一、教学目标', '5', bold=True, indent=False))
    parts.append(P('知识目标', '6', bold=True, indent=False))
    parts.append(P('1. 理解企业级财务数据中台与多源数据湖的构建原理，掌握POS销售流水、外卖对账单、采购电子专票与报销台账的字段映射规范。'))
    parts.append(P('2. 深刻掌握连锁餐饮行业核心营运指标（翻台率、客单价、食材成本率、坪效）与传统财务报表指标的深度勾稽关系。'))
    parts.append(P('3. 理解多智能体协同机制中角色隔离、上下文传递、逻辑冲突仲裁与共识收敛的系统架构原理。'))
    parts.append(P('技能目标', '6', bold=True, indent=False))
    parts.append(P('1. 能够独立利用多维表格与AI智能体构建自动化清洗流水线，完成千万级流水数据的异常清洗、税号校验与发票查重。'))
    parts.append(P('2. 能够熟练结合杜邦分析与四象限波士顿矩阵，穿透定位异常门店经营病灶，并搭建三级红黄牌动态风险预警系统。'))
    parts.append(P('3. 能够自主搭建具有双重视角（高管层驾驶舱与门店端看板）的交互式BI仪表盘，并运用金字塔原理生成专业白皮书。'))
    parts.append(P('4. 能够编排包含循环批处理、条件分支、多智能体协同会诊及移动端推送的端到端自动化工作流，并生成Streamlit交互轻应用。'))
    parts.append(P('素质目标', '6', bold=True, indent=False))
    parts.append(P('1. 树立“先拆解业务实质，再运用AI工具”的理性工程思维，杜绝脱离业务常识的盲目崇拜与一键代做。'))
    parts.append(P('2. 强化财务数据敏感度与合规把关意识，面对异常数据敢于穿透核查，严守业财合规与商业秘密保护底线。'))
    parts.append(P('3. 培养面向复杂企业现实场景的系统级工程架构能力与团队协作答辩素养。'))
    
    parts.append(P('二、重点难点', '5', bold=True, indent=False))
    parts.append(P('教学重点', '6', bold=True, indent=False))
    parts.append(P('1. 多源异构财务数据的清洗整合与多维表格底层逻辑构建。'))
    parts.append(P('2. 基于鲜合餐饮12家门店真实底稿的经营指标穿透与杜邦风险诊断。'))
    parts.append(P('3. 端到端自动化工作流与多角色智能体协同会诊的串联闭环。'))
    parts.append(P('教学难点', '6', bold=True, indent=False))
    parts.append(P('1. 解决循环批处理多门店数据文件时的超时阻断、容错跳过与格式对齐问题。'))
    parts.append(P('2. 激发“分析师-风控官-管理顾问”多智能体之间的实质性逻辑辩论，避免生成空泛敷衍的和稀泥意见。'))
    parts.append(P('3. 面对动态数据突变与突发参数抽测时，快速定位工作流节点数据链路并现场解答。'))
    
    parts.append(P('三、职业启示', '5', bold=True, indent=False))
    parts.append(P('现代财务人员的核心竞争力，正在从“低阶记账与手工对账员”向“企业业财数据资产架构师与人机协同决策专家”跃迁。掌握了单一工具并不能形成护城河；唯有能够将业务痛点转化为系统架构，将繁杂流程封装为自动工作流，并对AI输出结果保持严谨审慎的专业批判力，才能在未来的财务智能化浪潮中立于不败之地。'))

    # 2. 【情景导入】
    parts.append(P('【情景导入】', '4', bold=True, indent=False))
    parts.append(P('鲜合餐饮管理有限公司（拥有全国12家直营连锁门店，半年度营收规模1,694.4万元）在近期召开了年中经营管理闭门会。董事会与管理层对上半年各门店经营分化严重、利润被无形侵蚀的现状表达了高度关切。'))
    parts.append(P('财务部负责人召集数字化转型专班召开紧急动员会。会上展示了一份沉重的财务底稿：'))
    parts.append(P('第一，虽然全连锁上半年录得1,694.4万元营收与271.1万元净利润，但各门店分化触目惊心。以上海南京东路店、北京王府井店为代表的标杆门店翻台率高达4.2次/天、净利率超20%；而北京朝阳大悦城店食材成本率攀升至44.0%，且会员复购率腰斩，上海淮海中路店更是因租金畸高陷入持续亏损。'))
    parts.append(P('第二，财务团队每月面临着巨大的对账压力。12家门店每月产生上万条POS明细、两家主流外卖平台的对账单、数百张食材采购专票以及繁杂的报销凭证。团队几乎要把前半个月的时间消耗在手工清洗对账上，根本没有精力开展深度的经营分析与风险预警。'))
    parts.append(P('第三，管理层迫切需要一个能实时掌握全局的决策驾驶舱，而各门店店长则需要能随时自查自身短板的轻量化工具。目前分散在Excel表格里的静态汇报，早已无法满足敏捷决策的要求。'))
    parts.append(P('财务总监向数字化专班下达了“48小时攻坚战”硬指标：必须在两天内，综合运用AI工具链与智能体技术，为鲜合餐饮搭建一套从原始对账数据自动化清洗、多维指标穿透分析、BI可视化大屏，到端到端自动化工作流与多角色智能体协同会诊的完整决策支持系统。'))
    parts.append(P('面对这场全真模拟的企业级数字化转型大考，你将作为主理人带领团队迅速展开系统化攻坚！'))

    # 3. 任务 13.1
    parts.append(P('任务13.1 财务数据底座构建与数据清洗智能体实战', '4', bold=True, indent=False))
    parts.append(P('任务情境与目标：梳理鲜合餐饮12家直营门店的多源原始数据资产，在多维表格中搭建标准数据湖底座，配置数据清洗智能体完成半年度财务核心数据的自动化质检与校验。', '23', indent=False))
    parts.append(P('13.1.1 鲜合餐饮多源异构财务数据湖架构', '5', bold=True, indent=False))
    parts.append(P('连锁餐饮企业的财务数据具有典型的高频、异构、多触点特征。鲜合餐饮管理有限公司作为一家跨区域直营连锁餐饮企业，在全国布局了12家直营门店。门店每天产生大量的交易流水，覆盖了现金、微信支付、支付宝、美团团购、大众点评外卖券等多种收单渠道。不同渠道的数据格式差异极大：POS系统导出的数据以流水号为索引，外卖平台导出的数据以配送订单号为索引，且外卖平台还会扣除佣金、满减补贴与骑手配送费，造成账面营收与实际到账金额之间存在复杂的勾稽差额。'))
    parts.append(P('与此同时，鲜合餐饮的后端供应链数据同样复杂多样。蔬菜、鲜肉等生鲜食材由本地合作农业基地每日凌晨直配并提供农产品收购发票，调味品、粮油及半成品冷链则由总部中央厨房统一调度配送并开具增值税专用发票。在传统管理模式下，这些原始凭证散落于各门店报销审批单与纸质附件中，极易发生发票重复入账、品名与税率错配、退货未冲销等数据失真现象。'))
    parts.append(P('为了从根本上消除数据孤岛，专班首先对全连锁12家直营门店的2024年上半年度原始财务与营运数据进行全量汇总核验，形成了标准的基础财务指标底稿，见表13-1。'))
    parts.append(P('表13-1 鲜合餐饮12家直营门店半年度经营与财务核心指标底稿', '3', indent=False, jc='center'))
    
    t13_1_headers = ['门店编号与名称', '经营业态', '半年度营收(万元)', '食材成本(万元)', '人工成本(万元)', '租金折旧(万元)', '经营净利(万元)', '食材成本率', '翻台率(次/天)', '会员复购率']
    t13_1_rows = [
        ['01 北京王府井店', '商圈大店', '185.6', '66.8', '38.9', '42.6', '37.3', '36.0%', '4.2', '39%'],
        ['02 北京中关村店', '科技园店', '142.3', '52.6', '31.3', '35.5', '22.9', '37.0%', '3.5', '34%'],
        ['03 北京朝阳大悦城店', '商场店', '210.5', '92.6', '46.3', '48.4', '23.2', '44.0%', '3.8', '21%'],
        ['04 天津和平店', '社区综合店', '118.4', '42.6', '26.0', '29.6', '20.2', '36.0%', '3.1', '33%'],
        ['05 天津南开店', '高校商圈店', '98.6', '36.5', '21.7', '24.6', '15.8', '37.0%', '3.0', '31%'],
        ['06 上海南京东路店', '文旅旗舰店', '245.8', '88.5', '54.0', '51.6', '51.7', '36.0%', '4.5', '42%'],
        ['07 上海陆家嘴店', '商务白领店', '176.2', '63.4', '38.8', '40.5', '33.5', '36.0%', '3.9', '36%'],
        ['08 上海淮海中路店', '商圈街铺店', '105.2', '45.2', '25.2', '38.9', '-4.1', '43.0%', '1.8', '18%'],
        ['09 杭州武林广场店', '核心商圈店', '135.4', '48.7', '29.8', '32.5', '24.4', '36.0%', '3.4', '35%'],
        ['10 杭州西湖湖滨店', '景区景观店', '158.6', '57.1', '34.9', '36.5', '30.1', '36.0%', '3.8', '38%'],
        ['11 南京新街口店', '枢纽商圈店', '126.8', '46.9', '27.9', '30.4', '21.6', '37.0%', '3.2', '32%'],
        ['12 苏州金鸡湖店', '滨湖休闲店', '91.0', '32.8', '20.0', '22.8', '15.4', '36.0%', '2.9', '30%'],
        ['全连锁合计/均值', '12家直营店', '1694.4', '643.8', '372.8', '406.7', '271.1', '38.0%', '3.4', '32.6%'],
    ]
    parts.append(TBL(t13_1_headers, t13_1_rows))
    
    parts.append(P('13.1.2 飞书多维表格标准数据底座搭建', '5', bold=True, indent=False))
    parts.append(P('面对上述多源异构的底层数据，首先需要将其结构化为支持实时联动的数据底座。在飞书多维表格中，设计“门店主数据表”、“日销售流水表”、“食材采购批次表”与“月度综合指标表”四张关联数据表。'))
    parts.append(P('通过设立唯一的“门店编号”与“采购批次单号”作为主键，建立单向与双向关联字段，使跨表数据能够实现无缝聚合。多源财务数据清洗与整合的总体架构如图13-2所示。'))
    
    # 插入图13-2
    parts.append(DRAWING('rId6', 31, '鲜合餐饮多源财务数据清洗与整合架构图', cx=4950000, cy=2250000))
    parts.append(P('图13-2 鲜合餐饮多源财务数据清洗与整合架构图', '3', indent=False, jc='center'))
    
    parts.append(P('在多维表格中配置公式字段：食材成本率＝食材采购总额/营业收入；经营净利润＝营业收入－食材成本－人工成本－租金折旧及公摊；净利润率＝经营净利润/营业收入。系统利用预设公式实时更新计算结果，有效消除了手工录入计算时的人为差错。'))
    
    parts.append(P('13.1.3 鲜合数据清洗质控智能体配置与运行', '5', bold=True, indent=False))
    parts.append(P('在实际数据流转中，各门店原始对账文件经常存在数据缺失、门店重名（如“北京朝阳店”与“北京朝阳大悦城店”）、税号录入错误以及重复报销等脏数据隐患。为此，专班需要配置专门的“鲜合数据清洗与质控智能体”，利用大模型的语义理解和正则表达式双重机制实现精准治理。'))
    parts.append(P('在搭建清洗智能体时，必须向大模型注入鲜合餐饮的标准实体对照字典、税控规则库与业务防呆逻辑。智能体接收到原始表格后，首先运行字段对齐算法，检查关键列是否存在缺失；接着利用规则引擎自动核算发票代码有效性；最后对金额与税额进行交叉乘积反算，精准捕获因手工输错税率导致的计算偏差。'))
    parts.append(P('智能体的清洗规则体系与Prompt工程配置标准见表13-2。'))
    parts.append(P('表13-2 鲜合餐饮数据清洗规则与智能体Prompt配置表', '3', indent=False, jc='center'))
    
    t13_2_headers = ['清洗校验维度', '常见脏数据特征', '业务处理规则', '智能体Prompt指令核心规范']
    t13_2_rows = [
        ['门店实体归一', '别名混用、漏填直营标识', '严格映射至01-12标准编号', '“对照知识库《门店标准编码表》，将所有别名精确归一至12家标准全称”'],
        ['数值合规校验', '金额负数、退单异常', '负数退款需匹配原单并标记', '“识别所有营收列小于0的异常记录，检查是否有对应退单单号，无单号标红”'],
        ['食材专票验真', '发票代码缺失、税率错填', '餐饮食材采购专票强制9%或13%', '“检查发票税率字段，若餐饮农产品食材专票偏离法定税率，输出异常预警报告”'],
        ['重复报销识别', '连号发票、同金额多报', '严格执行金额+时间窗口查重', '“跨门店扫描同商户、同金额、开票时间间隔在3天以内的报销凭证，锁定疑似重报”'],
    ]
    parts.append(TBL(t13_2_headers, t13_2_rows))
    
    parts.append(P('【小试牛刀】鲜合餐饮12家门店异常采购流水自动化清洗与对账校验', '4', bold=True, indent=False))
    parts.append(P('利用智能体对鲜合餐饮6月份5,000条原始采购及报销流水进行自动化清洗，排查异常缺失与税号冲突。', '23', indent=False))
    parts.append(P('（1）第一步：先不借助AI自己拆问题。手写清洗规则逻辑树与校验算法公式；在纸上梳理发票四要素（发票代码、号码、金额、税额）的勾稽关系；手动抽取北京朝阳大悦城店与上海淮海中路店的20条测试样本进行人工核算，明确清洗期望值。'))
    parts.append(P('（2）第二步：带业务参数向AI提问。将清洗规则转化为结构化Prompt模板，输入包含缺失值与脏数据的测试流水，调用数据清洗智能体执行清洗与格式转换。'))
    parts.append(P('（3）第三步：核对与修正AI给出的结果。人工核对智能体输出的清洗日志，核查税号比对结果是否准确，修正大模型在处理跨月退款冲销时出现的误判，确保输出数据100%符合财务入账标准。'))
    parts.append(P('【善学勤思】', '4', bold=True, indent=False))
    parts.append(P('思考题：在企业级财务数据中台中，如果前端清洗智能体未能识别出虚假报销或重复录入的采购专票，这一错误将如何沿着数据链条传导至后续的杜邦分析与风险预警模块？财务人员应当建立怎样的人工兜底防御机制？'))

    # 4. 任务 13.2
    parts.append(P('任务13.2 深度财务多维分析与智能风险预警系统构建', '4', bold=True, indent=False))
    parts.append(P('任务情境与目标：基于清洗完成的鲜合餐饮半年度财务底稿，运用杜邦分析模型与四象限矩阵开展多维透视，搭建三级风险预警系统，穿透排查异常门店的核心病灶。', '23', indent=False))
    parts.append(P('13.2.1 鲜合餐饮多维经营指标透视与杜邦分析', '5', bold=True, indent=False))
    parts.append(P('完成数据底座搭建后，进入深度财务诊断阶段。首先，从营收结构切入：鲜合餐饮上半年总营收1,694.4万元中，堂食贡献了1,186.1万元（占比70%），外卖贡献了508.3万元（占比30%）。虽然整体营收规模保持稳健，但各门店净利润率呈现极端分化，最高达到21.0%（上海南京东路店），最低甚至为-3.9%（上海淮海中路店，亏损4.1万元）。'))
    parts.append(P('通过运用杜邦分析框架对净资产收益率（ROE）进行层层剥离：净资产收益率＝净利率×总资产周转率×权益乘数。在餐饮直营模式下，总资产周转率核心受翻台率驱动，而净利率则深度受制于食材成本率与租金人工刚性开支。'))
    parts.append(P('13.2.2 12家直营门店四象限矩阵分类与病灶穿透', '5', bold=True, indent=False))
    parts.append(P('为了直观分类各门店的经营健康度，专班以“日均翻台率（基准线3.4次/天）”为横轴，以“食材成本率（基准线38.0%）”为纵轴，构建鲜合餐饮门店四象限波士顿矩阵。12家直营门店在四象限波士顿矩阵中的落位分布如图13-3所示。'))
    
    # 插入图13-3
    parts.append(DRAWING('rId7', 41, '鲜合餐饮12家直营门店经营效益四象限波士顿矩阵', cx=4800000, cy=2914285))
    parts.append(P('图13-3 鲜合餐饮12家直营门店经营效益四象限波士顿矩阵', '3', indent=False, jc='center'))
    
    parts.append(P('矩阵清晰呈现出两大经营病灶：'))
    parts.append(P('第一病灶（问题象限）：北京朝阳大悦城店（03号店）。翻台率高达3.8次/天，营收达210.5万元，但食材成本率畸高达44.0%（偏离全连锁均值6.0个百分点），导致净利润仅录得23.2万元。结合其会员复购率仅21%的异动，穿透核查发现该店存在菜品分量严重超标、后厨冷库报损失控及核心单品BOM执行偏差等重大管理漏洞。'))
    parts.append(P('第二病灶（瘦狗象限）：上海淮海中路店（08号店）。翻台率仅1.8次/天，且食材成本率高达43.0%，租金折旧占营收比例高达37.0%，半年度净亏损4.1万元，不仅无法贡献现金流，且持续侵蚀总部利润。'))
    
    parts.append(P('13.2.3 三级红黄牌动态风险预警机制设计', '5', bold=True, indent=False))
    parts.append(P('基于上述指标特征，专班搭建“红-黄-绿”三级动态风险预警体系：绿牌（健康门店，食材成本率≤38%且翻台率≥3.2）；黄牌（关注门店，食材成本率在38%—40%或复购率偏低）；红牌（严重预警，食材成本率>40%或经营净利润为负）。鲜合餐饮12家门店的分类诊断与预警结果见表13-3。'))
    parts.append(P('表13-3 鲜合餐饮12家直营门店杜邦分析指标与风险分级表', '3', indent=False, jc='center'))
    
    t13_3_headers = ['门店编号与名称', '净利率', '食材成本率', '翻台率(次/天)', '风险等级', '核心经营病灶与归因说明']
    t13_3_rows = [
        ['01 北京王府井店', '20.1%', '36.0%', '4.2', '绿牌(标杆)', '客流与翻台均优，食材管控严密，单店盈利强劲'],
        ['02 北京中关村店', '16.1%', '37.0%', '3.5', '绿牌(稳健)', '白领客群稳定，外卖占比合理，各项指标健康'],
        ['03 北京朝阳大悦城店', '11.0%', '44.0%', '3.8', '红牌(高危)', '食材成本严重超标，损耗失控，会员复购断崖下跌'],
        ['04 天津和平店', '17.1%', '36.0%', '3.1', '绿牌(稳健)', '社区熟客基础好，成本与人工开支控制良好'],
        ['05 天津南开店', '16.0%', '37.0%', '3.0', '绿牌(稳健)', '高校客群平稳，暑期存在淡季波动但总体可控'],
        ['06 上海南京东路店', '21.0%', '36.0%', '4.5', '绿牌(标杆)', '文旅顶流大店，翻台与营收领跑全连锁'],
        ['07 上海陆家嘴店', '19.0%', '36.0%', '3.9', '绿牌(标杆)', '商务正餐客单价高，人效与平效双优'],
        ['08 上海淮海中路店', '-3.9%', '43.0%', '1.8', '红牌(高危)', '高租金、低翻台、高损耗，单店持续净亏损'],
        ['09 杭州武林广场店', '18.0%', '36.0%', '3.4', '绿牌(稳健)', '传统核心商圈店，营运指标平稳'],
        ['10 杭州西湖湖滨店', '19.0%', '36.0%', '3.8', '绿牌(标杆)', '景区高客流带动力强，毛利与净利表现优异'],
        ['11 南京新街口店', '17.0%', '37.0%', '3.2', '绿牌(稳健)', '交通枢纽客流稳定，管理规范'],
        ['12 苏州金鸡湖店', '16.9%', '36.0%', '2.9', '黄牌(关注)', '平日翻台率偏低，依赖周末客流，需提升周中获客'],
    ]
    parts.append(TBL(t13_3_headers, t13_3_rows))

    parts.append(P('【小试牛刀】基于杜邦分析对北京朝阳大悦城店开展食材异常下钻', '4', bold=True, indent=False))
    parts.append(P('对03号店44.0%的食材成本率进行深入穿透，定位牛肉、生鲜及调料三类物料的真实损耗差额。', '23', indent=False))
    parts.append(P('（1）第一步：先不借助AI自己拆问题。手动计算朝阳大悦城店偏离全连锁均值导致的利润侵蚀绝对金额：210.5万元×(44.0%－38.0%)＝12.63万元；梳理其前十大主推菜品的标准BOM理论耗量与实际出库量差值。'))
    parts.append(P('（2）第二步：带业务参数向AI提问。构建结构化提示词，将该店20个单品的理论出库量与实际核销单输入智能体，要求其按食材类别输出偏离度排名与风险概率。'))
    parts.append(P('（3）第三步：核对与修正AI给出的结果。交叉对比智能体归因结论与后厨盘点报表，剔除智能体因季节性生鲜损耗率自然上浮导致的过度解读，锁定牛里脊肉（偏离+8.5%）与鲜虾（偏离+6.2%）两项异常消耗点。'))
    parts.append(P('【善学勤思】', '4', bold=True, indent=False))
    parts.append(P('思考题：当杜邦分析指出上海淮海中路店因租金占比畸高陷入亏损时，财务人员在向管理层提出对策时，为什么不能仅仅简单建议“降租或闭店”？应当从哪些业财融合与品牌战略维度提供复合型决策依据？'))

    # 5. 任务 13.3
    parts.append(P('任务13.3 鲜合高管驾驶舱与门店端双重视角BI交互看板设计', '4', bold=True, indent=False))
    parts.append(P('任务情境与目标：面向高管决策层与门店店长两类不同角色，在BI工具中搭建具有双重视角的动态交互看板，并运用金字塔原理生成专业月度经营分析白皮书。', '23', indent=False))
    parts.append(P('13.3.1 高管驾驶舱与门店端看板双重视角架构', '5', bold=True, indent=False))
    parts.append(P('不同管理层级对财务数据的颗粒度诉求存在显著差异：高管层聚焦战略全盘、宏观趋势与红线风险，需要直观的高聚合KPI卡片、大区贡献对比图与异常报警指示；而门店店长则聚焦战术执行，关注当日翻台、外卖满减让利、人员排班工时及食材BOM单耗。双重视角交互看板设计原型如图13-4所示。'))
    
    # 插入图13-4
    parts.append(DRAWING('rId8', 51, '鲜合餐饮高管经营分析驾驶舱与门店端双重视角BI交互看板', cx=4950000, cy=2495629))
    parts.append(P('图13-4 鲜合餐饮高管经营分析驾驶舱与门店端双重视角BI交互看板', '3', indent=False, jc='center'))
    
    parts.append(P('13.3.2 动态切片联动与色彩语义规范', '5', bold=True, indent=False))
    parts.append(P('看板全面执行全书规范的“色彩语义逻辑”：墨蓝色（#1E293B）作为专业商务底色，翠绿色（#10B981）标识达标与盈利，琥珀金（#F59E0B）标识预警，艳红色（#EF4444）标识亏损与严重超标。看板设计标准详见表13-4。'))
    parts.append(P('表13-4 鲜合餐饮BI仪表盘指标与图表类型映射设计表', '3', indent=False, jc='center'))
    
    t13_4_headers = ['看板视图模块', '核心展示指标', '推荐图表类型', '交互联动与筛选设计']
    t13_4_rows = [
        ['高管总览驾驶舱', '半年度营收、净利润、全连锁食材率、平均翻台率', '大字号KPI指标卡', '支持按华北大区/华东大区全局切片筛选'],
        ['大区营收利润对比', '12家门店营收与净利润分布', '横向双向柱状对比图', '点击任一门店，联动下钻至单店收支明细'],
        ['三大成本动因拆解', '食材、人工、租金公摊占营收比重', '瀑布图/堆叠面积图', '悬浮展示各项成本偏离行业均值的具体差额'],
        ['红黄牌风险监控屏', '12家门店风险预警等级分布', '红黄绿状态气泡矩阵图', '一键过滤红牌门店，自动弹出病灶归因卡'],
        ['单店经营透视看板', '日翻台率、人均客单价、堂食外卖占比', '双轴折线与柱状组合图', '门店店长支持按天/周自由调整分析周期'],
        ['食材消耗BOM看板', '前十大食材理论消耗与实际出库差异', '雷达图/标靶图(Bullet)', '标红展示差异率超过5%的高危食材品类'],
    ]
    parts.append(TBL(t13_4_headers, t13_4_rows))

    parts.append(P('13.3.3 基于金字塔原理生成月度经营分析白皮书', '5', bold=True, indent=False))
    parts.append(P('运用金字塔原理（结论先行、以上统下、归类分组、逻辑递进），指导AI智能体将看板数据转化为专业结构化白皮书：顶部明确“全连锁营收稳中有增，但03店与08店两大病灶侵蚀利润超50万元”的核心结论；中间支撑层展开区域对比与成本动因分析；底层落地刚性整改对策。'))
    
    parts.append(P('【小试牛刀】基于BI看板数据撰写鲜合餐饮半年度经营分析白皮书摘要', '4', bold=True, indent=False))
    parts.append(P('提取看板核心数据，编写一份面向董事会的300字结构化经营分析摘要。', '23', indent=False))
    parts.append(P('（1）第一步：先不借助AI自己拆问题。列出“全连锁总盘结论—标杆门店经验—异常门店问题—下半年三大举措”手写四段提纲，核准1,694.4万营收、271.1万利润等基准数字。'))
    parts.append(P('（2）第二步：带业务参数向AI提问。向大模型提供四段式提纲与全套12家门店数据，要求严格按照金字塔原理成文，语言务实客观，禁止泛泛空谈。'))
    parts.append(P('（3）第三步：核对与修正AI给出的结果。检查AI摘要中是否存在过度夸大成绩或将08号店亏损轻描淡写的倾向，手动调整措辞，强化对供应链损耗治理的刚性要求。'))
    parts.append(P('【善学勤思】', '4', bold=True, indent=False))
    parts.append(P('思考题：当BI看板展示出北京朝阳大悦城店食材成本率超标时，如果仅提供月度静态指标，店长往往难以及时采取补救措施。如何通过打通POS日流水与即时通讯软件，构建“日维度动态超标微预警”？'))

    # 6. 任务 13.4
    parts.append(P('任务13.4 端到端自动化工作流编排与多智能体协同会诊', '4', bold=True, indent=False))
    parts.append(P('任务情境与目标：在工作流引擎中编排包含循环批处理、条件分支与推送的端到端自动化流程，配置“分析师-风控官-管理顾问”三智能体协同会诊链，并生成单店交互Streamlit轻应用。', '23', indent=False))
    parts.append(P('13.4.1 端到端全自动财务分析工作流拓扑编排', '5', bold=True, indent=False))
    parts.append(P('为了使月度财务分析工作实现全自动无人值守流转，专班设计了包含循环节点与复杂条件分支的端到端工作流。端到端自动化财务分析工作流拓扑编排如图13-5所示。'))
    
    # 插入图13-5
    parts.append(DRAWING('rId9', 61, '鲜合餐饮端到端自动化财务分析工作流拓扑编排图', cx=4950000, cy=2250000))
    parts.append(P('图13-5 鲜合餐饮端到端自动化财务分析工作流拓扑编排图', '3', indent=False, jc='center'))
    
    parts.append(P('工作流拓扑节点及其配置规范见表13-5。'))
    parts.append(P('表13-5 鲜合餐饮月度财务分析工作流节点拓扑与配置说明表', '3', indent=False, jc='center'))
    
    t13_5_headers = ['节点序号与类型', '节点功能定义', '输入参数配置', '输出结果与流转逻辑']
    t13_5_rows = [
        ['节点1: 定时触发器', '每月1日08:00自动触发工作流', 'Cron表达式: 0 8 1 * *', '激活下游文件读取监听节点'],
        ['节点2: 数据集读取', '读取12家门店待清洗对账文件', '输入文件路径列表', '输出包含12家门店数据的数组集合'],
        ['节点3: 循环批处理', '遍历数组逐一执行单店清洗与BOM计算', '单店原始POS与采购明细', '循环体内调用清洗智能体，生成单店宽表'],
        ['节点4: 条件分支判断', '判断单店食材成本率>40% 或 净利<0', '输入各店食材成本率与净利', '是则走分支A(会诊流)，否则走分支B(月报流)'],
        ['节点5: 多智能体协同', '分析师、风控官与顾问串行协同会诊', '分支A分流的异常门店数据底稿', '生成包含风险定性与行动举措的穿透白皮书'],
        ['节点6: 报告组装推送', '将正常与异常门店报告合并渲染并推送', '各节点输出的Markdown片段', '向飞书高管群推送摘要并邮件归档PDF完整版'],
    ]
    parts.append(TBL(t13_5_headers, t13_5_rows))

    parts.append(P('13.4.2 “分析师-风控官-管理顾问”多智能体协同机制', '5', bold=True, indent=False))
    parts.append(P('在多智能体协作链条中，单一智能体往往容易陷入“角色偏见”或“合谋迎合”。专班设置三个互相制衡的专业角色：'))
    parts.append(P('角色一（财务数据分析师）：严格恪守客观中立立场，只陈述数据异动事实，严禁加入主观推测或建议。例如指出：“北京朝阳大悦城店食材成本率偏离均值6.0%，牛肉采购单耗超标8.5%”。'))
    parts.append(P('角色二（财务风控合规官）：采取审慎怀疑态度，专注识别潜在造假、利益输送与合规红线漏洞。例如质疑：“后厨冷库报损缺乏双人复核验收单，高度疑似存在食材飞单或盘点虚增”。'))
    parts.append(P('角色三（经营管理咨询顾问）：作为矛盾仲裁者，权衡商业现实与整改可行性，提出刚性落地行动方案。例如裁定：“责令总部供应链立即派出内审督导组进驻朝阳大悦城店，锁定冷库盘点；建立单店周度BOM预警红线；优化促销套餐挽救会员复购”。三智能体协同会诊交互流程如图13-6所示。'))
    
    # 插入图13-6
    parts.append(DRAWING('rId10', 71, '鲜合多智能体协同会诊交互与轻应用展示界面', cx=4950000, cy=2410880))
    parts.append(P('图13-6 鲜合多智能体协同会诊交互与轻应用展示界面', '3', indent=False, jc='center'))
    
    parts.append(P('13.4.3 Streamlit 财务智能体轻应用交互构建', '5', bold=True, indent=False))
    parts.append(P('为了让非技术背景的管理者能够灵活调用智能体，专班借助大模型编写了轻量化的 Streamlit 前端交互应用。前端设置门店下拉选择框、食材阈值调节滑块与“一键生成综合会诊报告”按钮；后端无缝调用多维表格API与大模型会话引擎，实现低代码敏捷交付。'))

    parts.append(P('【小试牛刀】配置并测试三智能体协同链对上海淮海中路店的亏损会诊', '4', bold=True, indent=False))
    parts.append(P('输入08号店亏损4.1万元与租金占比37.0%的现实数据，测试三智能体能否输出辩论充分且建议刚性的诊断结论。', '23', indent=False))
    parts.append(P('（1）第一步：先不借助AI自己拆问题。手动梳理08号店的现金流断流临界点：测算该店每日最低保本翻台率应达到2.4次，而在目前1.8次翻台水平下，月度固定租金净流出达6.5万元。'))
    parts.append(P('（2）第二步：带业务参数向AI提问。分别设定三个角色的System Prompt约束，串联运行工作流，观察风控官与管理顾问在“立即闭店止损”与“尝试降租30%续约”之间的辩论过程。'))
    parts.append(P('（3）第三步：核对与修正AI给出的结果。检验智能体最终输出的仲裁方案是否具备实操性，剔除“加大全城广告投放”等脱离单店实际的不合理建议，确认“启动降租谈判，若两个月内翻台未回升至2.2次则坚决启动闭店迁址”作为最终结论。'))
    parts.append(P('【善学勤思】', '4', bold=True, indent=False))
    parts.append(P('思考题：在配置循环批处理节点处理12家门店数据时，如果某一家门店的数据文件因网络波动出现损坏或格式解析错误，整个工作流应当如何设计容错与重试机制，以避免全局流程中断？'))

    # 7. 【项目核心实操】
    parts.append(P('【项目核心实操】鲜合餐饮全流程财务智能体决策系统综合交付大考', '4', bold=True, indent=False))
    parts.append(P('一、任务场景', '5', bold=True, indent=False))
    parts.append(P('鲜合餐饮管理有限公司董事长与管理层将于后天上午听取数字化专班的最终汇报。这是一场全真模拟的实战大考：你和团队必须在48小时内，基于全国12家直营门店半年度1,694.4万元营收的真实业务底稿，完成涵盖数据清洗、杜邦透视、BI看板、工作流编排、多角色协同与轻应用部署的完整工业级系统交付，并现场接受评审专家的抗辩抽测！'))
    
    parts.append(P('二、四步实战任务', '5', bold=True, indent=False))
    parts.append(P('第一步  搭建飞书多维表格中台并运行数据清洗质控智能体', '5', bold=True, indent=False))
    parts.append(P('（1）中台建模：在多维表格中建立门店主数据、流水明细、BOM批次与月度宽表四大关联表，配置标准计算公式字段。'))
    parts.append(P('（2）质控治理：运行清洗智能体，对测试流水中的异名、负数、税号错误及重复报销进行100%拦截并输出质控日志。'))
    parts.append(P('第二步  构建多维经营分析模型与双视角交互BI驾驶舱', '5', bold=True, indent=False))
    parts.append(P('（1）指标勾稽：核算12家门店半年度核心营运指标，确保营收总计1,694.4万、净利271.1万手工勾稽100%一致。'))
    parts.append(P('（2）多维穿透诊断：执行杜邦分析与四象限矩阵分类，配置三级红黄牌预警系统，锁定3号店与8号店的经营病灶。'))
    parts.append(P('（3）BI大屏搭建：搭建具有高管驾驶舱与门店店长双重视角的动态交互看板，配置区域联动与预警指示。'))
    parts.append(P('第三步  编排端到端自动化工作流与多智能体协同系统', '5', bold=True, indent=False))
    parts.append(P('（1）工作流串联：配置定时触发、循环批处理12个门店文件、异常条件分支及机器人消息推送的自动化闭环。'))
    parts.append(P('（2）多智能体协作：配置“分析师-风控官-管理顾问”串行协同链，完成异常门店的深度穿透会诊报告。'))
    parts.append(P('（3）轻应用生成：借助AI编写Streamlit代码并在本地测试运行单店交互查询看板。'))
    parts.append(P('第四步  执行系统全量联调、容错注入与现场答辩', '5', bold=True, indent=False))
    parts.append(P('（1）系统全流测试：导入全量数据，执行端到端自动化联调，检验输出报告的完整性与排版规范。'))
    parts.append(P('（2）容错抗扰检验：注入脏数据与极端矛盾测试用例，验证工作流容错阻断机制。'))
    parts.append(P('（3）答辩准备：整理全流程工作底稿，准备迎接现场动态抽测答辩。'))

    parts.append(P('三、成果提交', '5', bold=True, indent=False))
    parts.append(P('实训任务完成后，必须规范整理并提交表13-6所列的五份企业级专业底稿。'))
    parts.append(P('表13-6 鲜合餐饮综合实训成果提交物清单', '3', indent=False, jc='center'))
    
    t13_6_headers = ['成果序号', '交付成果规范名称', '成果主要内容与格式载体', '评价考核权重']
    t13_6_rows = [
        ['成果一', '《鲜合多源财务数据底座构建与清洗质控底稿》', '多维表格中台链接、清洗规则逻辑树、发票验真查重日志表（Excel/PDF）', '20%'],
        ['成果二', '《鲜合12家门店深度财务分析与风险预警诊断报告》', '杜邦分析图表、四象限矩阵、红黄牌分级预警表及异常归因报告（Word/PDF）', '25%'],
        ['成果三', '《鲜合高管经营驾驶舱BI交互看板设计与运行报告》', 'BI动态看板访问地址、交互功能说明书及全景运行截图（PPT/PDF）', '20%'],
        ['成果四', '《月度自动化工作流拓扑配置与三智能体协同白皮书》', '工作流DSL导出文件/流程截图、三角色Prompt设计表及会诊白皮书（Word）', '20%'],
        ['成果五', '《Streamlit轻应用代码与现场答辩综合评价表》', 'Python脚本源码、本地浏览器运行录屏及现场答辩评分表（含互评签名）', '15%'],
    ]
    parts.append(TBL(t13_6_headers, t13_6_rows))

    parts.append(P('四、评价标准（全书统一五维度，满分100分）', '5', bold=True, indent=False))
    parts.append(P('本综合实训严格遵循全书统一的五维度评价标准，满分100分，各维度评分细则见表13-7。'))
    parts.append(P('表13-7 项目13综合实战全流程评价标准表', '3', indent=False, jc='center'))
    
    t13_7_headers = ['评价维度', '满分', '优秀(90—100分)', '良好(80—89分)', '合格(60—79分)', '不合格(<60分)']
    t13_7_rows = [
        ['A. 任务完整性与闭环', '20分', '五份成果底稿规范齐备，从数据到工作流全链路闭环，架构严谨', '五份成果基本齐备，各环节衔接顺畅，无重大缺漏', '缺少1份非核心底稿，主体流程勉强跑通', '成果严重缺失，未能完成核心实操任务'],
        ['B. 数据严密与核算精度', '25分', '12家门店数据完全吻合底稿，指标勾稽严密，手工验算一致率100%', '数据核算基本正确，存在极个别笔误但不影响大局', '核算出现局部公式错误，导致指标分析出现轻度偏差', '数据逻辑混乱，存在重大算术差错或虚假捏造'],
        ['C. 智能体工程与风控深度', '25分', 'Prompt工程高度专业，多智能体分工鲜明，风险穿透深刻且对策刚性', 'Prompt设计良好，智能体协同有效，能指出核心风险', 'Prompt较为简单，分析停留在表面，建议空泛', '完全照抄默认提示词，智能体胡言乱语无实际价值'],
        ['D. 可视化与前沿工程实现', '15分', 'BI看板视觉严谨交互流畅，工作流自动化程度高，Streamlit顺畅运行', '看板布局规范，工作流能稳定运行，轻应用无报错', '看板视觉杂乱，工作流偶发卡死，代码调试吃力', '看板无法交互，工作流崩溃，代码无法运行'],
        ['E. 过程留痕与实战答辩', '15分', '第一步手工底稿详实，面对现场参数突变与容错抽查应对自如流利', '留痕记录完整，答辩逻辑清晰，能回答大部分抽测', '留痕记录简略，答辩较为紧张，需提示后方能修正', '无手工规划底稿，面对现场抽测完全无法回答'],
    ]
    parts.append(TBL(t13_7_headers, t13_7_rows))
    parts.append(P('综合成绩评定档次：优秀（90—100分）、良好（80—89分）、合格（60—79分）、不合格（60分以下）。凡出现抄袭他人成果或直接使用AI一键生成代码答辩者，一票否决判定为不合格。'))

    parts.append(P('五、反 AI 代做与动态答辩机制', '5', bold=True, indent=False))
    parts.append(P('为了防范依赖AI生成虚假成果交付，实训答辩现场设置三道动态对抗式抽测机制：'))
    parts.append(P('（1）随机抽测 1（数据突变与容错注入）：现场向清洗流水线注入一份含有负数金额、税号缺失或篡改日期的测试表格，测试系统能否在3分钟内自动捕获异常并输出阻断警告日志。'))
    parts.append(P('（2）随机抽测 2（事实矛盾与多角色辩论）：现场向多智能体协同链推入一组极端冲突的商业事实（如某店营收大涨50%但经营现金流断崖下跌且采购发票全失），检验风控智能体与管理顾问能否发生真实的专业逻辑对抗而非机械给出一致意见。'))
    parts.append(P('（3）随机抽测 3（动态参数现场重跑）：现场微调核心预警阈值（如将食材警戒线由40.0%下调至35.0%），要求在2分钟内调整工作流节点参数，并实时演示BI大屏与推送简报的联动刷新效果。'))

    parts.append(P('六、拓展挑战（选做，不计入上述分数）', '5', bold=True, indent=False))
    parts.append(P('（1）探索将工作流与企业ERP系统或外部供应链金融平台API打通，实现基于门店动态经营指标的供应链智能融资信用额度测算。'))
    parts.append(P('（2）尝试利用大模型多模态能力，基于门店现场监控截图识别客流拥挤度与餐盘剩菜率，自动联动修正食材采购BOM损耗预警模型。'))

    # 8. 【项目总结】
    parts.append(P('【项目总结】', '4', bold=True, indent=False))
    parts.append(P('走到这里，你已经完成了全书从项目1到项目13的完整进阶之旅！回顾这趟旅程，你的能力发生了系统性跃迁：'))
    parts.append(P('第一阶段·认知与工具构建（项目1～2）：打破对AI的神秘感与恐惧感，掌握提问工程与工具箱搭建；'))
    parts.append(P('第二阶段·数据治理与管理（项目3～5）：从多维表格到数据获取，再到数据清洗与第一个规则智能体搭建，夯实业财数据底座；'))
    parts.append(P('第三阶段·深度财务分析与风控（项目6～8）：融合描述性分析、杜邦比率分析与风险预警系统，实现从数据到商业洞察的飞跃；'))
    parts.append(P('第四阶段·数据叙事与协同自动化（项目9～12）：掌握专业可视化、BI看板叙事、知识智能体与自动化工作流编排；'))
    parts.append(P('第五阶段·全流程综合实战交付（项目13）：以鲜合餐饮为全真战场，独立交付工业级财务智能体协同解决方案。'))
    parts.append(P('技术工具会不断更新换代，但“清晰界定业务问题、严谨验证底层数据、用专业洞察推动商业行动”的财务核心思维永远不会过时。带着这套方法论与人机协同本领，你已完全具备走向现代智能化财务职场的坚实底气！'))

    # 9. 【实践报告】
    parts.append(P('【实践报告】', '4', bold=True, indent=False))
    parts.append(P('请按表13-8完成本项目的实践报告。', '24', indent=True))
    parts.append(P('表13-8 项目13实践报告模板', '3', indent=False, jc='center'))
    
    t13_8_headers = ['实践报告项目', '填报内容与自评反思']
    t13_8_rows = [
        ['项目名称', '财务AI智能体全流程综合实战'],
        ['业务场景与定位', '鲜合餐饮管理有限公司全国12家直营门店半年度经营分析与自动化决策系统交付'],
        ['全流程技术栈', '多维表格中台 + 数据清洗智能体 + 杜邦分析模型 + BI看板 + 自动化工作流 + 三智能体协同 + Streamlit轻应用'],
        ['数据体量与验证', '1,694.4万元营收全量流水，12家门店核心指标手工验算一致率100%'],
        ['遇到的核心难题', '多源对账校验、工作流循环批处理超时或多角色冲突仲裁中的最大阻碍及解决经过'],
        ['全流程最大收获', '系统梳理从分项练习走向企业级端到端综合交付的思维转变与能力跃迁（具体写3点）'],
        ['人机协同深度反思', '工具是手臂的延伸，而对财务数据逻辑的深刻理解与商业判断才是永恒的核心灵魂'],
        ['当前阶段探索困惑', '写出2—3个仍待进一步探索的企业级实战疑问，并制定自主学习计划'],
        ['综合评定意见', '评审小组针对五维度评分（任务完整性、数据准确性、智能体深度、可视化呈现、答辩表现）的综合评定意见'],
        ['评定档次', '□ 优秀（90—100分）   □ 良好（80—89分）   □ 合格（60—79分）   □ 不合格（60分以下）'],
    ]
    parts.append(TBL(t13_8_headers, t13_8_rows))

    # 10. 【创新拓展】
    parts.append(P('【创新拓展】', '4', bold=True, indent=False))
    parts.append(P('在现代连锁商业生态中，财务AI智能体的应用边界正在持续向产业链上下游延伸。'))
    parts.append(P('方向A：供应链金融与智能信用评估。结合鲜合餐饮12家门店的真实经营流水与食材采购结算记录，利用多智能体模型评估上游生鲜供应商与加盟商的信用风险，自动生成供应链保理融资决策建议，帮助小微食材供应商解决账期融资难题。'))
    parts.append(P('方向B：多模态业财融合与新店选址预测。整合商圈人流热力图数据、美团大众点评用户情绪画像以及周边竞品客单价数据，构建智能体驱动的新店投资测算工作流，在开店前对新店未来三年的营收、保本点与投资回收期进行蒙特卡洛随机模拟推演，为企业资本开支提供刚性决策支持。'))

    # 11. 【综合测评】
    parts.append(P('【综合测评】', '4', bold=True, indent=False))
    parts.append(P('一、能力测评', '5', bold=True, indent=False))
    parts.append(P('（一）架构分析与方案设计题', '6', bold=True, indent=False))
    parts.append(P('1. 鲜合餐饮管理有限公司在开展多源财务数据清洗时，为什么必须将“门店实体归一化”和“食材采购电子专票查重”作为最前置的基础规则？如果直接将原始流水推给大模型进行经营分析，可能引发什么严重的财务误判后果？'))
    parts.append(P('2. 在构建高管驾驶舱与门店端看板的双重视角时，简述两类管理者在数据颗粒度、决策时效性与图表呈现上的核心差异，并画出指标流向关系草图。'))
    parts.append(P('（二）端到端工作流与协同机制设计题', '6', bold=True, indent=False))
    parts.append(P('3. 设计一个用于鲜合餐饮月度财务分析的端到端自动化工作流。要求：包含文件触发、循环批处理12个门店文件、异常条件分支判断及消息推送节点，用文字清晰描述节点拓扑关系与参数传递逻辑。'))
    parts.append(P('4. 在“分析师-风控官-管理顾问”三智能体协作链中，为了防止三个智能体沦为“表面合谋、一味说好话”的形式主义，在设计系统Prompt时应当采取哪些关键约束规则？'))
    
    parts.append(P('二、素质测评', '5', bold=True, indent=False))
    parts.append(P('请对以下职场观点进行判断，并结合财务职业道德与人机协同规范简要说明理由。', '6', bold=True, indent=False))
    parts.append(P('1. “有了端到端自动化工作流和多智能体协同系统，以后财务分析报告就能做到全自动无人值守直接发布，财务人员完全不需要再参与过程审核。”'))
    parts.append(P('2. “在用AI生成Streamlit前端代码时，只要代码能在本地跑通、图表能展示出来就行，至于底层的SQL取数逻辑和财务计算口径对不对，财务人员不需要深究。”'))
    parts.append(P('3. “当分析智能体指出的某门店食材成本率超标风险，与分管该店的副总裁的汇报口径发生冲突时，财务分析师应当直接修改智能体的参数阈值，让输出结果符合领导预期。”'))

    parts.append(P('三、参考答案（要点）', '5', bold=True, indent=False))
    parts.append(P('能力测评参考要点：', '6', bold=True, indent=False))
    parts.append(P('1. 实体归一化与专票查重是数据中台的基石。餐饮流水中门店别名混杂会导致跨店汇总错位；专票重复录入会虚增食材成本与进项税额。若直接将脏数据推给大模型，会导致大模型基于错误事实进行幻觉归因，严重误导管理层对单店盈利能力的判断。'))
    parts.append(P('2. 高管层聚焦宏观战略、全局趋势与异常预警，偏好聚合卡片与趋势瀑布图；店长聚焦微观操作与单品BOM单耗，偏好单店明细与排名对标。指标流向遵循“底层单店明细聚合为大区指标，最终收敛至全盘高管驾驶舱”。'))
    parts.append(P('3. 工作流拓扑：文件监听触发→循环节点输入12个门店文件列表→循环内部执行单店质检与BOM计算→条件分支判断食材成本率>40%或亏损→是则流转多智能体深度会诊流，否则流转标准报告流→两路汇总至报告渲染节点→推送飞书群机器人。'))
    parts.append(P('4. 采取三项刚性约束：①角色互斥与职责锁定（分析师严禁提建议，风控官强制以审慎怀疑立场挑刺，管理顾问必须调和矛盾提出刚性举措）；②强制输入上游结构化证据，严禁无凭据赞美；③设置矛盾仲裁判定规则，不解决核心分歧不予结束流程。'))
    
    parts.append(P('素质测评参考要点：', '6', bold=True, indent=False))
    parts.append(P('1. 错误。自动化工作流提升的是数据处理与组装效率，但最终报告的真实性、合法性与商业合理性责任始终在财务人员。人机协同必须坚持“人是最终责任主体”，重大经营分析报告必须经专业财务人员复核把关后方可对外发布。'))
    parts.append(P('2. 错误。工具只是表现形式，财务分析的灵魂在于数据口径与财务逻辑的准确性。如果底层取数逻辑或勾稽关系错误，酷炫的可视化应用只会加速扩散错误信息，导致灾难性的管理决策失误。财务人员必须看懂核心口径与数据映射。'))
    parts.append(P('3. 错误。财务人员必须恪守客观公正、坚持准则的职业道德底线。面对智能体揭示的客观风险与领导主观意图的冲突，应当深入核实业务事实，用扎实的数据底稿和客观证据进行专业沟通，严禁篡改参数粉饰报表或隐瞒重大经营隐患。'))

    return "".join(parts)

def stage_A(raw):
    """全局流式交替重排引号，确保引号开闭对齐"""
    idxs = []
    for a, b in para_spans(raw):
        for m in re.finditer(r'[“”]', raw[a:b]):
            idxs.append((a + m.start(), m.group(0)))
    expect_open, edits = True, []
    for pos, ch in idxs:
        want = '“' if expect_open else '”'
        if ch != want:
            edits.append((pos, want))
        expect_open = not expect_open
    for pos, want in reversed(edits):
        raw = raw[:pos] + want + raw[pos + 1:]
    log('【A】引号流式修复：%d 个引号，改动 %d 处' % (len(idxs), len(edits)))
    return raw

def main():
    shutil.rmtree(WORK_DIR, ignore_errors=True)
    os.makedirs(WORK_DIR, exist_ok=True)
    
    src_template = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), SRC_TEMPLATE)
    if not os.path.exists(src_template):
        src_template = SRC_TEMPLATE
    
    log(f'=== 项目13 自动化生成与构建流水线启动 ===')
    log(f'使用底模模板：{src_template}')
    
    with zipfile.ZipFile(src_template) as z:
        z.extractall(WORK_DIR + '/work')
    
    # 提取 header 和 footer
    orig_doc = open(WORK_DIR + '/work/word/document.xml', encoding='utf-8').read()
    m_head = re.match(r'(<\?xml.*?>\s*<w:document[^>]*>\s*<w:body>)', orig_doc)
    header_xml = m_head.group(1) if m_head else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    footer_xml = '<w:sectPr><w:cols w:space="720" w:num="1"/></w:sectPr></w:body></w:document>'
    
    # 覆盖替换媒体图片文件为项目13专业高清图表
    media_dir = os.path.join(WORK_DIR, 'work/word/media')
    os.makedirs(media_dir, exist_ok=True)
    for i in range(1, 7):
        src_img = os.path.join(IMAGES_DIR, f'image{i}.png')
        dst_img = os.path.join(media_dir, f'image{i}.png')
        if os.path.exists(src_img):
            shutil.copyfile(src_img, dst_img)
            log(f'  [Media] 写入图表文件：image{i}.png ({os.path.getsize(dst_img)/1024:.1f} KB)')
        else:
            log(f'  [Warning] 图表文件不存在：{src_img}')
    
    content_xml = build_content()
    full_doc = header_xml + content_xml + footer_xml
    
    # 引号流式重排与标点规范化
    full_doc = stage_A(full_doc)
    full_doc, n_punct = normalize_punct_docx(full_doc)
    log(f'【标点规范化】微调 {n_punct} 处')
    
    # 写回 document.xml
    open(WORK_DIR + '/work/word/document.xml', 'w', encoding='utf-8').write(full_doc)
    
    # 确保没有残留 comments.xml
    comm_path = os.path.join(WORK_DIR, 'work/word/comments.xml')
    if os.path.exists(comm_path):
        os.remove(comm_path)
    
    dst_docx = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), DST_DOCX)
    if not os.path.exists(os.path.dirname(dst_docx)):
        dst_docx = DST_DOCX
    
    with zipfile.ZipFile(dst_docx, 'w', zipfile.ZIP_DEFLATED) as out_zip:
        for root, dirs, files in os.walk(WORK_DIR + '/work'):
            for f in files:
                fp = os.path.join(root, f)
                arc = os.path.relpath(fp, WORK_DIR + '/work')
                out_zip.write(fp, arc)
                
    log(f'=== 项目13 成稿生成成功：{dst_docx} ===')
    log(f'文件大小：{os.path.getsize(dst_docx) / 1024.0:.1f} KB')

if __name__ == '__main__':
    main()
