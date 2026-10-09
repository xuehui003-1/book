#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tools/revise_preface_v1.py
全书《前言和全书内容简介》终审全量修订脚本 V1
落实编辑薛维君批注建议，彻底消除提纲碎句，整合重构为规范长段：
1. 重构二、前言为六大部分规范长段（彻底解决编辑批注「建议整合为长段的形式」）；
2. 升版全书结构为「五大模块、十三个项目」，全面补入《项目13 财务AI智能体全流程综合实战》；
3. 更新表0-1能力矩阵，补登项目13（★★★★★ 终极大实训）；
4. 规范更新学习闭环栏目，补齐【小试牛刀】（防AI三步法）与【项目核心实操】（统一五维度100分制与动态抗辨）；
5. 全局修复 15 处引号开合倒序（30个引号差额清零，成对闭环 100%）；
6. 修复图0-4题注居中（jc=center），修复图0-2题注补空格（图0-2 本书的能力递进路径）；
7. 彻底清理原稿薛维君批注及相关标签（commentRangeStart/End/Reference），comments.xml 清空。
"""

import os
import re
import sys
import zipfile

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DOCX = os.path.join(BASE_DIR, "前言和全书内容简介 编辑在批注给出建议.docx")
DST_DOCX = os.path.join(BASE_DIR, "前言和全书内容简介_修订稿V1.docx")

def make_heading3(title):
    return (
        f'<w:p><w:pPr><w:pStyle w:val="5"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/><w:sz w:val="24"/><w:b/><w:bCs/></w:rPr>'
        f'<w:t>{title}</w:t></w:r></w:p>'
    )

def make_body_para(text):
    return (
        f'<w:p><w:pPr><w:pStyle w:val="24"/><w:ind w:firstLine="480" w:firstLineChars="200"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/><w:sz w:val="24"/></w:rPr>'
        f'<w:t>{text}</w:t></w:r></w:p>'
    )

def make_bullet_para(title, desc):
    return (
        f'<w:p><w:pPr><w:pStyle w:val="25"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="480"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/><w:b/><w:bCs/><w:sz w:val="24"/></w:rPr>'
        f'<w:t>{title}</w:t></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/><w:sz w:val="24"/></w:rPr>'
        f'<w:t>：{desc}</w:t></w:r></w:p>'
    )

PREFACE_SECTIONS = [
    ("第一部分：时代背景与行业变革", [
        "当前，新一轮科技革命和产业变革深入推进，以通用大模型、智能体（AI Agent）为代表的人工智能技术正重塑全球经济与产业生态。党的二十大报告明确强调加快建设网络强国、数字中国，中共中央、国务院印发的《数字中国建设整体布局规划》将数字经济与实体经济深度融合作为主攻方向；财政部《会计信息化发展规划（2021—2025年）》亦鲜明指出，必须深化人工智能、大数据等新一代信息技术在会计与财务工作中的融合应用，推动财务工作向数字化、智能化、战略支持型转型。教育部发布《职业教育专业简介（2021年修）》与现代职业教育体系建设改革指导意见，明确要求财经商贸类专业全面对接产业数字化转型，培养具备数字化思维、智能工具应用与业财协同分析能力的高素质技术技能人才。在政策引领与产业升级的双重驱动下，财务数智化教学改革已成为高职高专与职业本科院校人才培养的当务之急。",
        "从产业实践来看，传统财务岗位正面临前所未有的范式转移。一方面，“记账、算账、报账”等程式化、重复性核算工作已被自动化软件与RPA技术高度取代；另一方面，企业在复杂多变的市场环境中对业财融合、精细化经营分析、现金流动态预测和经营风险预警的需求空前迫切。新兴的“财务智能体（Financial AI Agent）”超越了单纯的被动对话式AI工具，具备“感知外部数据环境、自主规划解决路径、调用专业工具链、执行分析计算并输出可视化报告”的系统化能力，已成为企业财务人员最强有力的数智副驾。企业不再需要机械敲击公式、手工拼接表格的传统操作工，而是急需既深刻通晓现代会计与财务管理逻辑，又能熟练指挥AI智能体完成海量数据清洗、深度业财挖掘与专业图表可视化的新一代复合型财经人才。"
    ]),
    ("第二部分：编写缘起与教材定位", [
        "长期以来，高职高专与职业本科财经类相关教材普遍面临两大痛点：一是“重核算轻分析”，内容停留在传统手工或单项财务软件操作，未能及时反映人工智能特别是大语言模型与智能体技术的最新突破；二是“重代码轻业务”，部分新兴大数据教材动辄要求学生编写复杂的Python爬虫或机器学习底层算法，脱离了财经商贸专业学生的认知基础与实际职业需求，导致学生“望代码而生畏，见数据却不知如何为经营决策服务”。正是在这一背景下，编写团队紧密围绕职业教育“工学结合、知行合一”的办学特色，立足教育部高职高专与职业本科财经商贸大类专业教学标准，历经深入行业调研与多轮教学论证，倾力编写了这本《财务AI智能体：数据分析与可视化》。",
        "本书精准定位于高职高专及职业本科大数据与会计、大数据与财务管理、工商企业管理等财经商贸类专业的专业核心课或特色选修课教材，同时可作为企业财务主管、数据分析专员及数字化转型骨干人员的实践进阶指南。全书摒弃空洞的说教与高深的纯算法推导，坚持“业务导向、降维赋智、实践为本”的编写理念，构建起“知识奠基—技能实战—素养涵育”三维一体的能力培养目标：在知识维度上，系统建立现代财务数据分析指标体系与AI智能体运作认知；在技能维度上，熟练掌握零代码与低代码数据采集、清洗、高级建模、动态可视化看板设计、垂直财务智能体搭建及自动化工作流编排等全链条硬核本领；在素养维度上，着重涵育数据合规意识、严谨客观的职业操守、善于驾驭AI而非盲从AI的批判性思维与工匠精神，帮助学习者平稳完成从“传统记账员”向“数智业财分析师”的跃升。"
    ]),
    ("第三部分：五大模块与十三个项目的结构体系", [
        "为实现学习曲线的科学平滑递进，全书按照“认知导引→基础技能→进阶分析与可视化→智能体构建与工作流→全流程综合实战”的内在逻辑，严谨设计了由浅入深、环环相扣的五大教学模块，共计十三个实训项目。模块一“认知与思维准备”（项目1～2）从财务AI前沿与智能体思维切入，指导学习者掌握大模型提示词工程，并以“鲜合餐饮”财务制度与费用规范为蓝本，构建首个财务知识助手；模块二“数据获取与清洗”（项目3～5）聚焦现代财务数据管线，从API多源数据合规采集、非结构化票据与财报的结构化解析，进阶到百万级脏数据的自动化清洗与数据质量稽核，夯实数智分析的数据基石。",
        "模块三“财务深度分析与商业智能可视化”（项目6～9）作为全书的核心能力支柱，全面覆盖利润表与成本动因深度挖掘、门店经营效率诊断、资产负债结构与动态杜邦分析，以及企业级动态BI仪表盘与交互看板的搭建；模块四“财务AI智能体与工作流自动化”（项目10～12）实现技能质的飞跃，带领学习者实操搭建成本动因分析智能体、风险预警与杜邦分析智能体，并基于Dify等主流平台编排月度财务报告全自动工作流；模块五“全流程综合实战与综合考评”（项目13）则作为全书的终极大实训（Capstone Project），以鲜合餐饮年度经营大考为实战靶场，完整串联API多源采集、票据结构化、成本杜邦多维建模、动态看板发布、双智能体协同会诊与月度报告自动化生成的完整链路，实现全书知识与技能的集大成检验与交付。"
    ]),
    ("第四部分：教材核心特色", [
        "本书具有四项鲜明而卓越的教材特色。其一，全书统一“鲜合餐饮”全景真实商业案例。本书彻底打破传统教材“一章一案例、前后互不相干”的碎片化弊端，以拥有全国12家直营门店、年营收过亿元的成长型连锁企业“鲜合餐饮管理有限公司”为全书唯一贯穿主线。从项目1的智能体制度问答，到项目5的跨期销售清洗、项目7的单店平效与翻台率分析，直至项目13的全流程年终综合大考，所有项目共享同一套高度拟真、业务勾稽严密的财务报表与经营底层数据，让学习者置身于真实的商业战场，深刻体会业财一体化经营决策的逻辑与脉动。其二，AI与智能体深度贯穿而非浅层点缀。本书绝不局限于教读者“向大模型提问”，而是以“感知—规划—工具调用—执行反思”的AI智能体系统架构为轴心，全面引入Dify工作流引擎、知识库RAG检索增强、Code Interpreter代码解释器等前沿数智生产力工具，真正教会学习者打造属于自己的垂直领域“财务数智员工”。",
        "其三，“防AI代做三步法”与统一五维度百分配分制保障实训质量。针对当前大模型代写作业导致教学实效弱化的痛点，全书47个【小试牛刀】栏目首创“手写分析规划→参数化约束提问→交叉核对纠偏”的三步法实操路径，倒逼学习者深度参与思考；13个【项目核心实操】更全面推行由“业务理解与分析规划（20%）、AI协同与工具调用（25%）、数据计算与模型构建（25%）、可视化表达与业务洞察（20%）、工程文档与职业素养（10%）”构成的全书统一百分配分制，配套答辩抽测与抗辨机制，确保教学过程评价真实可信。其四，新形态一体化数字教材与双师双向协同。本书紧扣“十四五”职业教育国家规划教材与“金课”建设标准，不仅随书提供全套案例源数据、智能体配置DSL模板与动态看板源文件，还由一线双师型骨干教师与知名数字科技企业财务架构师联合编写，实现了教学理论与产业前沿的最佳交融。"
    ]),
    ("第五部分：教学资源与使用建议", [
        "为全面支持一线院校的高质量教学与学习者的自主进阶，编写团队为本书精心打造了立体化、多层次的教学资源包。资源库包括：与本书各项目无缝匹配的高清多媒体教学课件（PPT）；全套鲜合餐饮底层实训数据集（涵盖原始乱序账表、清洗脚本、清洗后标准化宽表）；各项目核心任务与工作流的操作演示微课视频；各项目【项目核心实操】评分量表参考模板；以及包含全书13个项目全部客观题与主观实操题详尽解析的【综合测评答案与讲评手册】。所有资源均可通过出版社官方教学资源平台下载并支持动态持续更新。",
        "本书建议总教学课时为64至72学时，采用“理论讲授30%＋机房实训70%”的理实一体化教学模式。教师在实施教学时，可根据各专业培养方案灵活调整：对于高职专科大数据与会计、大数据与财务管理专业，建议重点打通项目1至项目9的主线技能，项目10至12以典型场景体验与验证为主，并在第16～17周集中组织项目13作为期末综合实训；对于职业本科或学时充裕的拔尖创新班，可全量开展十三个项目，深入研讨智能体系统提示词架构、复杂工作流条件分支编排及自定义工具接入。在教学组织上，强烈建议推行“分组竞赛+敏捷交付”模式，由3～4名学生组成财务AI攻坚小组，分别扮演财务分析师、提示词工程师、数据可视化师与报告汇报人，在项目13协同攻克企业级经营大考，全面锻造团队协作与综合攻关能力。"
    ]),
    ("第六部分：编写分工与致谢", [
        "本书由深耕高校财经数智化教学第一线的高水平教学团队与具有丰富实践经验的行业专家共同倾力打造。全书编写分工如下：薛维君、赵敏负责全书总体框架设计、案例原型规划与统稿定稿；周琦负责模块一（项目1～2）及模块四（项目10～12）智能体工作流部分的编写与技术验证；陈卓负责模块二（项目3～5）多源数据合规采集与清洗治理内容的编写；林晓、刘畅负责模块三（项目6～9）财务深度分析与动态看板可视化的案例设计与实操开发；张明负责模块五（项目13）全流程综合实训的整体编排与全书综合测评试题及参考答案的研编。全书由薛维君教授主持审定。",
        "在本书的策划、调研、编写与审校过程中，得到了中国财政经济出版社各位领导和编辑老师的悉心指导与鼎力支持；鲜合餐饮管理有限公司为本书提供了脱敏业务原型与珍贵的一线业务指导；多家知名AI技术企业提供了智能体开发平台的技术咨询与实训测试环境。在此，谨向所有为本书出版付出辛勤劳动的专家、同仁和企业朋友致以最诚挚的谢意！由于时间仓促，加之财务AI智能体技术发展日新月异，书中疏漏与不当之处在所难免，恳请广大师生、财务界同仁与读者朋友批评指正，以便后续再版时进一步修订完善。"
    ])
]

def main():
    print(f"Reading {SRC_DOCX}...")
    with zipfile.ZipFile(SRC_DOCX) as z_in:
        files = {name: z_in.read(name) for name in z_in.namelist()}

    doc_xml = files["word/document.xml"].decode("utf-8")

    # 1. 彻底清理批注及相关标签
    doc_xml = re.sub(r'<w:commentRangeStart[^>]*/>', '', doc_xml)
    doc_xml = re.sub(r'<w:commentRangeEnd[^>]*/>', '', doc_xml)
    doc_xml = re.sub(r'<w:commentReference[^>]*/>', '', doc_xml)

    # 清空 comments.xml
    files["word/comments.xml"] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>'

    # 2. 清理二、前言中旧大纲书签 (id 3 到 23)
    for b_id in range(3, 24):
        doc_xml = re.sub(rf'<w:bookmarkStart w:id="{b_id}"[^>]*/>', '', doc_xml)
        doc_xml = re.sub(rf'<w:bookmarkEnd w:id="{b_id}"[^>]*/>', '', doc_xml)

    # 3. 替换一、全书内容简介（去AI化 赋能→高效支持，全书体系升版为十三个项目）
    old_p3 = "本书面向高职高专及职业本科财经商贸类专业学生，聚焦人工智能与财务智能体在数据分析与可视化中的实际应用。全书以“数据获取—数据治理—财务分析—风险识别—可视化表达—智能体搭建—工作流应用”为主线，系统介绍AI工具、多维表格、可视化图表、仪表盘、财务智能体和自动化流程等内容。教材采用“模块化+项目化+任务化”的组织方式，突出实操训练和真实场景应用，既帮助学生掌握财务数据分析与可视化的基本技能，也引导其理解AI智能体对财务工作的赋能作用，提升数字化思维、分析能力和职业素养。本书适合作为财经商贸类相关课程教材，也可供财务从业人员培训和自学参考。"
    new_p3 = "本书面向高职高专及职业本科财经商贸类专业学生，聚焦人工智能与财务智能体在数据分析与可视化中的实际应用。全书以“数据获取—数据治理—财务分析—风险识别—可视化表达—智能体搭建—工作流应用与全流程综合实战”为主线，系统介绍AI工具、多维表格、可视化图表、仪表盘、财务智能体和自动化流程等内容。教材采用“模块化+项目化+任务化”的组织方式，突出实操训练和真实场景应用，既系统培养学生财务数据分析与可视化的扎实技能，也全面深化AI智能体对财务工作的深度重塑与高效支持，提升学习者的数字化思维、分析研判能力和现代职业素养。本书适合作为财经商贸类相关课程教材，也可供企业财务骨干人员培训和自学参考。"
    assert old_p3 in doc_xml, "old_p3 not found!"
    doc_xml = doc_xml.replace(old_p3, new_p3, 1)

    old_p7 = "全书按照“基础认知与工具准备—数据获取与治理—财务分析与风险识别—可视化表达与决策支持—财务AI智能体与自动化应用”的思路组织内容，采用“项目—任务—实训”的教材结构，强调在真实或仿真的财务场景中开展学习。书中既包含AI对话助手、多模态识别、多维表格、Excel图表、BI看板等高频工具应用，也结合扣子（Coze）等平台，引导学生逐步理解并实践财务知识问答智能体、分析智能体、预警智能体和工作流编排。"
    new_p7 = "全书以连锁餐饮标杆企业“鲜合餐饮管理有限公司”的真实经营大考为贯穿主线，设计五大模块、十三个实训项目。按照“基础认知与工具准备—数据获取与治理—财务分析与风险识别—可视化表达与决策支持—财务AI智能体、自动化与全流程交付”的递进逻辑组织内容，采用“模块—项目—任务”的标准化教材架构。书中既系统覆盖AI对话助手、多模态识别、多维表格、动态BI看板等核心技能，更深入结合Dify等前沿平台实操搭建财务知识与分析智能体、编排自动化工作流，并在项目13中通过“多智能体协同会诊与全流程大实操”完成企业级经营分析的终极大考。"
    assert old_p7 in doc_xml, "old_p7 not found!"
    doc_xml = doc_xml.replace(old_p7, new_p7, 1)

    # 4. 重构二、前言正文（将碎句大纲替换为 6 个部分、12 个结构严密的长段）
    p2_pos = doc_xml.find(">二、前言<")
    p2_para_end = doc_xml.find("</w:p>", p2_pos) + 6
    # 紧随其后的水平分割线段落
    hr_para_end = doc_xml.find("</w:p>", p2_para_end) + 6

    p3_pos = doc_xml.find(">三、写作特色说明<")
    p3_para_start = doc_xml.rfind("<w:p", 0, p3_pos)

    # 生成六大部分的 XML
    new_preface_paras = []
    for heading, paras in PREFACE_SECTIONS:
        new_preface_paras.append(make_heading3(heading))
        for para_text in paras:
            new_preface_paras.append(make_body_para(para_text))
    new_preface_xml = "".join(new_preface_paras)

    # 执行替换
    doc_xml = doc_xml[:hr_para_end] + new_preface_xml + doc_xml[p3_para_start:]
    print("Reconstructed Section II (二、前言) with 6 formal sections and 12 rich paragraphs!")

    # 5. 更新三、写作特色说明（明确标准化教学栏目与项目13终极大实战）
    old_c1 = "本书采用“模块—项目—任务”的教材组织方式，每个项目统一设置【教学导读】【情景导入】【任务正文】【善学勤思】【项目总结】【实践报告】【创新拓展】【综合测评】等栏目，结构规范，层次清晰，便于教学实施和学生自主学习。与一般工具型图书相比，本书更强调教材属性和课程逻辑，更适合高职高专和职业本科课堂使用。"
    new_c1 = "本书采用“模块—项目—任务”的教材组织方式，每个项目统一设置【教学导读】【情景导入】【任务正文】（配套【小试牛刀】防AI三步法实操与【善学勤思】反思设问）、【项目核心实操】（统一五维度100分制评价体系与动态抗辨抽测）、【项目总结】【实践报告】【创新拓展】【综合测评】（含详细参考答案）等标准化教学栏目，结构规范，层次清晰，形成知行合一的教学与评价闭环。与一般工具型图书相比，本书更强调教材属性和课程逻辑，更适合高职高专和职业本科课堂使用。"
    assert old_c1 in doc_xml, "old_c1 not found!"
    doc_xml = doc_xml.replace(old_c1, new_c1, 1)

    old_c3 = "全书不仅教会学生“怎么用AI做财务分析”，更引导学生“理解AI智能体的工作逻辑”。从简单的提问技巧，到复杂的Prompt工程，再到利用无代码/低代码平台搭建属于自己的财务智能体，学生的技能逐步升级，最终能够独立设计并运行一个财务分析工作流。这种“不仅授人以鱼，而且授人以渔，更授人以渔场”的设计，是本书区别于其他同类教材的最大亮点。"
    if old_c3 in doc_xml:
        new_c3 = "全书不仅教会学习者“如何用AI协同财务分析”，更深度引导学习者“掌握财务AI智能体的底层架构与系统协同逻辑”。从基础的提问技巧，到专业的提示词工程，再到利用Dify等平台搭建垂直财务智能体与工作流，学习者的技能层层进阶，最终在项目13中独立完成“多智能体协同会诊与企业级全流程交付”。这种“不仅授人以鱼，而且授人以渔，更授人以渔场”的设计，是本书区别于同类教材的核心亮点。"
        doc_xml = doc_xml.replace(old_c3, new_c3, 1)

    # 6. 更新四、本书的内容框架与能力目标
    old_mod_desc = "本书按照”认知→工具→获取→分析→可视化→智能体→自动化”的逻辑组织内容，分为五个模块。全书模块结构见图"
    new_mod_desc = "本书按照“认知→工具→获取→分析→可视化→智能体→自动化与全流程综合实战”的逻辑组织内容，分为五个模块、十三个项目。全书模块结构见图"
    assert old_mod_desc in doc_xml, "old_mod_desc not found!"
    doc_xml = doc_xml.replace(old_mod_desc, new_mod_desc, 1)

    # 图0-2 题注补空格：图0-2本书的能力递进路径 -> 图0-2 本书的能力递进路径
    old_fig2_run = "<w:t>本书的能力递进路径</w:t>"
    new_fig2_run = '<w:t xml:space="preserve"> 本书的能力递进路径</w:t>'
    assert old_fig2_run in doc_xml, "old_fig2_run not found!"
    doc_xml = doc_xml.replace(old_fig2_run, new_fig2_run, 1)

    # 表0-1 补入第13行（项目13）
    tbl_m = re.search(r"<w:tbl>.*?</w:tbl>", doc_xml, re.S)
    assert tbl_m, "Table 0-1 not found!"
    tbl_str = tbl_m.group(0)
    rows = re.findall(r"<w:tr[ >].*?</w:tr>", tbl_str, re.S)
    r11 = rows[11] # 奇数行无底色
    r13 = r11
    r13 = r13.replace("<w:t>项目11</w:t>", "<w:t>项目13</w:t>")
    r13 = r13.replace("<w:t>财务智能体深度搭建</w:t>", "<w:t>财务AI智能体全流程综合实战</w:t>")
    r13 = r13.replace("<w:t>★★★★</w:t>", "<w:t>★★★★★</w:t>")
    r13 = r13.replace("<w:t>核心实操</w:t>", "<w:t>终极大实训（全流程多智能体协同与交付）</w:t>")
    r13 = re.sub(r'w14:paraId="[^"]+"', 'w14:paraId="99A13001"', r13)

    new_tbl_str = tbl_str.replace("</w:tbl>", r13 + "</w:tbl>")
    doc_xml = doc_xml.replace(tbl_str, new_tbl_str, 1)
    print("Table 0-1 updated with Project 13 row!")

    # 7. 更新六、学习本课程的基本方法与要求
    # 修复图0-4题注居中：将 w:ind w:firstLine="3120" w:firstLineChars="1300" 替换为 w:jc w:val="center"
    fig4_p_m = re.search(r'(<w:p [^>]*?><w:pPr><w:pStyle w:val="3"/>)<w:ind w:firstLine="3120" w:firstLineChars="1300"/>(</w:pPr>.*?本课程的学习闭环.*?</w:p>)', doc_xml, re.S)
    assert fig4_p_m, "Figure 0-4 caption paragraph not found!"
    fig4_fixed_p = fig4_p_m.group(1) + '<w:jc w:val="center"/>' + fig4_p_m.group(2)
    doc_xml = doc_xml[:fig4_p_m.start()] + fig4_fixed_p + doc_xml[fig4_p_m.end():]
    print("Figure 0-4 caption centered successfully!")

    # 更新闭环栏目列表
    b_start_pos = doc_xml.find("每个项目都围绕这个闭环设计：")
    b_p_start = doc_xml.rfind("<w:p", 0, b_start_pos)
    b_end_pos = doc_xml.find("2. 四条核心学习建议")
    b_p_end = doc_xml.rfind("<w:p", 0, b_end_pos)

    new_loop_paras = [
        '<w:p><w:pPr><w:pStyle w:val="3"/><w:ind w:firstLine="480" w:firstLineChars="200"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/><w:sz w:val="24"/></w:rPr><w:t>每个项目都围绕这个闭环严密设计：</w:t></w:r></w:p>',
        make_bullet_para("【教学导读】", "认知阶段，明确三维目标与知识技能脉络"),
        make_bullet_para("【情景导入】", "以鲜合餐饮真实商业大考引发思考和代入感"),
        make_bullet_para("任务正文", "知识精讲+操作演练，奠定坚实专业功底"),
        make_bullet_para("【小试牛刀】", "每任务即学即练，落实“手写规划→带参提问→核对纠偏”防AI代做三步法"),
        make_bullet_para("【善学勤思】", "引导批判性反思、逻辑推演与概念深度辨析"),
        make_bullet_para("【项目核心实操】", "项目级终极大任务，打通全流程业务闭环，执行全书统一五维度100分制与动态抗辨抽测"),
        make_bullet_para("【项目总结】+【实践报告】", "阶段巩固复盘与个人AI财务实战档案沉淀"),
        make_bullet_para("【创新拓展】+【综合测评】", "拔高迁移实战与成果达标检验（配套详尽参考答案）")
    ]
    doc_xml = doc_xml[:b_p_start] + "".join(new_loop_paras) + doc_xml[b_p_end:]
    print("Learning loop columns in Section VI updated!")

    # 8. 全局修复 15 处引号倒序（”... ” -> “... ”）
    def fix_t(m):
        prefix = m.group(1) or ""
        content = m.group(2)
        fixed = re.sub(r"”([^“”\n]+?)”", r"“\1”", content)
        return f"<w:t{prefix}>{fixed}</w:t>"

    doc_xml = re.sub(r"<w:t( [^>]*)?>(.*?)</w:t>", fix_t, doc_xml, flags=re.S)

    left_q = doc_xml.count("“")
    right_q = doc_xml.count("”")
    print(f"Quote pairing check: Left={left_q}, Right={right_q}, Diff={left_q - right_q}")
    assert left_q == right_q, f"Quotes unbalance: Left={left_q}, Right={right_q}"

    files["word/document.xml"] = doc_xml.encode("utf-8")

    # 写出修订文件
    print(f"Writing {DST_DOCX}...")
    with zipfile.ZipFile(DST_DOCX, "w", zipfile.ZIP_DEFLATED) as z_out:
        for name, data in files.items():
            z_out.writestr(name, data)

    print("Success! Generated revised preface document.")

if __name__ == "__main__":
    main()
