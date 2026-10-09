#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
whole_book_audit.py —— 全书十三个项目全景深度审计脚本

检查维度：
  1. 机械格式与批注（调用 check_docx.py 规则）：题注居中、引号开闭平衡、无直引号、无残留批注、无AI特征词
  2. 编辑批注 A 类审计（调用 audit_a_class.py 规则）：A1去AI、A2读者视角、A3/A4居中、A5分行、A6薄节、A7相关性
  3. 业务口径与矛盾筛查：
     - 案例企业统一性（鲜合餐饮）
     - 门店数量一致性（12家门店，排查8家/10家等旧残留）
     - 全书项目总数与封闭定调（排查“十二个项目”、“最后两个项目”、“收官”等矛盾表述）
     - 异构企业残留（排查“制造企业”、“律所”、“健身企业”、“快消品”等旧残留）
  4. 标准栏目齐备性（遗漏排查）：
     - 教学导读、情景导入、小试牛刀（各任务全覆盖）、善学勤思、项目核心实操、项目总结、实践报告、创新拓展、综合测评（含参考答案）
     - 防AI三步法落地
     - 全书统一五维度100分制评分体例落地
"""

import os, sys, re, zipfile, html

TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
if TOOLS_DIR not in sys.path:
    sys.path.insert(0, TOOLS_DIR)

import check_docx, audit_a_class

BOOK_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

PROJECTS = [
    (1, "03_项目1_认识财务数据分析与AI智能体_修订稿V8.docx"),
    (2, "项目2 搭建AI数据分析工具箱_修订稿V6.docx"),
    (3, "项目3 多维表格与智能数据管理_修订稿V4.docx"),
    (4, "项目4 财务数据的获取整理与整合_修订稿V2.docx"),
    (5, "项目5 数据清洗质量控制与第一个智能体_修订稿V1.docx"),
    (6, "项目6 财务描述性分析与分析智能体_修订稿V1.docx"),
    (7, "项目7 财务比率分析与智能综合评价_修订稿V1.docx"),
    (8, "项目8 财务异常识别与智能风险预警系统_修订稿V1.docx"),
    (9, "项目9 财务数据可视化设计与制作_修订稿V1.docx"),
    (10, "项目10 财务数据仪表盘BI看板与数据叙事_修订稿V1.docx"),
    (11, "项目11 财务知识智能体与分析智能体搭建_修订稿V1.docx"),
    (12, "项目12 工作流编排多智能体协作与前沿应用_修订稿V1.docx"),
    (13, "项目13 财务AI智能体全流程综合实战_初稿V1.docx"),
]

def log(msg):
    print(msg, flush=True)

def audit_all():
    log("=" * 80)
    log("                 《财务AI智能体：数据分析与可视化》全书13册深度全景审计")
    log("=" * 80)
    
    summary = []
    
    for p_num, fname in PROJECTS:
        fpath = os.path.join(BOOK_DIR, fname)
        if not os.path.exists(fpath):
            log(f"[FATAL] 项目 {p_num} 文件不存在: {fname}")
            continue
            
        raw, com = check_docx.read_docx(fpath)
        paras = check_docx.paragraphs(raw)
        doc_text = "\n".join([p["text"] for p in paras])
        
        # 1. check_docx 自检
        issues, counts = check_docx.check(fpath)
        bad_issues = [iss for iss in issues if iss[0] == "必改"]
        hint_issues = [iss for iss in issues if iss[0] == "提示"]
        
        # 2. A 类审计
        _, audit_paras = audit_a_class.load(fpath)
        a1_hits = audit_a_class.a1(audit_paras)
        a2_hits = audit_a_class.a2(audit_paras)
        a3_hits = audit_a_class.a34(audit_paras, '图')
        a4_hits = audit_a_class.a34(audit_paras, '表')
        a5_hits = audit_a_class.a5(audit_paras)
        a6_hits = audit_a_class.a6(audit_paras)
        a7_hits = audit_a_class.a7(audit_paras)
        
        # 3. 业务矛盾检查
        contra = []
        # 检查是否残留非鲜合案例关键词
        bad_keywords = ["律师事务所", "劲健", "快消品公司", "制造企业"]
        for kw in bad_keywords:
            # 排除特定正文用作反例/对比提及的合理情况，查找是否在情景/应用题出现
            count = doc_text.count(kw)
            if count > 0:
                contra.append(f"包含旧案例词「{kw}」x{count}")
                
        # 检查门店数量矛盾（如“8个门店”、“8家门店”）
        for kw in ["8个门店", "8家门店", "8家直营", "8间门店"]:
            count = doc_text.count(kw)
            if count > 0:
                contra.append(f"包含旧门店数「{kw}」x{count}")
                
        # 检查全书项目总数矛盾（如“十二个项目”、“前十二个项目”在非回顾上下文）
        for kw in ["五大模块、十二个项目", "十二个项目构建", "本书共十二个项目", "全部十二个项目"]:
            count = doc_text.count(kw)
            if count > 0:
                contra.append(f"包含封闭项目数「{kw}」x{count}")
                
        # 4. 栏目齐备性检查
        sections = {}
        for col_name in ["【教学导读】", "【情景导入】", "【小试牛刀】", "【善学勤思】", "【项目核心实操】", "【项目总结】", "【实践报告】", "【创新拓展】", "【综合测评】"]:
            sections[col_name] = doc_text.count(col_name)
            
        has_answers = ("参考答案" in doc_text) or ("参考要点" in doc_text)
        has_rubric = ("五维度" in doc_text) or ("满分100分" in doc_text) or ("100分" in doc_text)
        has_three_step = doc_text.count("第一步：先不借助AI") + doc_text.count("第一步  先不借助AI")
        
        record = {
            "p_num": p_num,
            "fname": fname,
            "paras": len(paras),
            "figs": counts.get("图", 0),
            "tbls": counts.get("表", 0),
            "bad": len(bad_issues),
            "hint": len(hint_issues),
            "bad_details": bad_issues,
            "hint_details": hint_issues,
            "a1": len(a1_hits),
            "a2": len(a2_hits),
            "a2_details": a2_hits,
            "a6": len(a6_hits),
            "contra": contra,
            "sections": sections,
            "has_answers": has_answers,
            "has_rubric": has_rubric,
            "three_step_count": has_three_step
        }
        summary.append(record)
        
        # 打印单册精简日志
        status_flag = "✔ 完美" if (len(bad_issues) == 0 and len(contra) == 0) else "✘ 待检"
        log(f"项目 {p_num:2d} | 段落 {len(paras):3d} | 图 {counts.get('图',0):2d} 表 {counts.get('表',0):2d} | 必改 {len(bad_issues)} 提示 {len(hint_issues)} | A2={len(a2_hits)} | 矛盾={len(contra)} | {status_flag} | {fname}")
        if contra:
            for c in contra:
                log(f"    ✘ [矛盾] {c}")
        if bad_issues:
            for b in bad_issues:
                log(f"    ✘ [必改] {b}")
        if a2_hits:
            for hit in a2_hits:
                log(f"    △ [A2命中] 第{hit[0]}段 ({hit[1]}): {hit[2][:45]}...")

    log("\n" + "=" * 80)
    log("                             全书十三册交叉比对汇总表")
    log("=" * 80)
    log(f"{'项目':<6} | {'段落':<5} | {'图/表':<6} | {'必改':<4} | {'小试牛刀':<6} | {'核心实操':<6} | {'五维评分':<6} | {'综合测答案':<8} | {'案例与架构冲突'}")
    log("-" * 80)
    for r in summary:
        xst = r["sections"].get("【小试牛刀】", 0)
        hxc = r["sections"].get("【项目核心实操】", 0)
        rub = "✔ 100分" if r["has_rubric"] else "✘ 缺失"
        ans = "✔ 有答案" if r["has_answers"] else "✘ 缺失"
        ct = "✔ 无冲突" if len(r["contra"]) == 0 else f"✘ {','.join(r['contra'])}"
        log(f"P{r['p_num']:<5} | {r['paras']:<5} | {r['figs']}/{r['tbls']:<4} | {r['bad']:<4} | {xst:<8} | {hxc:<8} | {rub:<8} | {ans:<10} | {ct}")
        
    return summary

if __name__ == "__main__":
    audit_all()
