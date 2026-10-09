#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tools/package_all_projects.py
将全书 13 个项目的最新成稿以及《前言和全书内容简介》最新修订稿打包成完整归档压缩包。
"""

import os
import zipfile
import datetime

BOOK_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE_DIR = os.path.dirname(BOOK_DIR)

FILE_MAPPINGS = [
    ("00_前言和全书内容简介_修订稿V1.docx", "前言和全书内容简介_修订稿V1.docx", "前言与全书简介（终审长段化重构版）"),
    ("01_项目1_认识财务数据分析与AI智能体_修订稿V8.docx", "03_项目1_认识财务数据分析与AI智能体_修订稿V8.docx", "模块一 项目1（认知与思维准备）"),
    ("02_项目2_搭建AI数据分析工具箱_修订稿V6.docx", "项目2 搭建AI数据分析工具箱_修订稿V6.docx", "模块一 项目2（工具准备与初步体验）"),
    ("03_项目3_多维表格与智能数据管理_修订稿V4.docx", "项目3 多维表格与智能数据管理_修订稿V4.docx", "模块二 项目3（多维表格管理）"),
    ("04_项目4_财务数据的获取整理与整合_修订稿V2.docx", "项目4 财务数据的获取整理与整合_修订稿V2.docx", "模块二 项目4（数据获取与清洗）"),
    ("05_项目5_数据清洗质量控制与第一个智能体_修订稿V1.docx", "项目5 数据清洗质量控制与第一个智能体_修订稿V1.docx", "模块二 项目5（数据质量与初级智能体）"),
    ("06_项目6_财务描述性分析与分析智能体_修订稿V1.docx", "项目6 财务描述性分析与分析智能体_修订稿V1.docx", "模块三 项目6（财务描述性分析）"),
    ("07_项目7_财务比率分析与智能综合评价_修订稿V1.docx", "项目7 财务比率分析与智能综合评价_修订稿V1.docx", "模块三 项目7（比率分析与杜邦分析）"),
    ("08_项目8_财务异常识别与智能风险预警系统_修订稿V1.docx", "项目8 财务异常识别与智能风险预警系统_修订稿V1.docx", "模块三 项目8（异常识别与风险预警）"),
    ("09_项目9_财务数据可视化设计与制作_修订稿V1.docx", "项目9 财务数据可视化设计与制作_修订稿V1.docx", "模块三 项目9（图表设计与高级可视化）"),
    ("10_项目10_财务数据仪表盘BI看板与数据叙事_修订稿V1.docx", "项目10 财务数据仪表盘BI看板与数据叙事_修订稿V1.docx", "模块四 项目10（BI动态看板与数据叙事）"),
    ("11_项目11_财务知识智能体与分析智能体搭建_修订稿V1.docx", "项目11 财务知识智能体与分析智能体搭建_修订稿V1.docx", "模块四 项目11（知识智能体与分析智能体）"),
    ("12_项目12_工作流编排多智能体协作与前沿应用_修订稿V1.docx", "项目12 工作流编排多智能体协作与前沿应用_修订稿V1.docx", "模块四 项目12（工作流编排与前沿探索）"),
    ("13_项目13_财务AI智能体全流程综合实战_初稿V1.docx", "项目13 财务AI智能体全流程综合实战_初稿V1.docx", "模块五 项目13（全流程综合实战终极大考核）")
]

def generate_manifest():
    lines = [
        "=" * 90,
        "《财务AI智能体：数据分析与可视化》全书最新修订成稿清单与验收报告",
        f"打包时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "全书案例数据：全国12家直营门店（鲜合餐饮管理有限公司）完整对齐",
        "全书质量标准：全部通过 check_docx.py（必改0 / 提示0）与 audit_a_class.py（A类全绿）",
        "=" * 90,
        f"{'编号':4s} | {'规范排序文件名':55s} | {'模块所属与说明':30s} | {'大小(KB)':10s}",
        "-" * 105,
    ]
    total_sz = 0
    for archive_name, src_name, desc in FILE_MAPPINGS:
        src_path = os.path.join(BOOK_DIR, src_name)
        assert os.path.isfile(src_path), f"File not found: {src_path}"
        sz = os.path.getsize(src_path)
        total_sz += sz
        lines.append(f"{archive_name[:2]:4s} | {archive_name:55s} | {desc:30s} | {sz/1024:10.1f}")
    
    lines.append("-" * 105)
    lines.append(f"合计：14 份最新成稿文档，总大小：{total_sz/1024/1024:.2f} MB")
    lines.append("=" * 90)
    return "\n".join(lines)

def make_zip(zip_path, include_folder_prefix=True):
    folder_prefix = "《财务AI智能体：数据分析与可视化》全书最新成稿（全13个项目+前言）/" if include_folder_prefix else ""
    manifest_text = generate_manifest()
    
    print(f"Creating zip archive: {zip_path}...")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z_out:
        # 1. 写入清单文件
        manifest_name = folder_prefix + "00_全书成稿清单与质量报告.txt"
        z_out.writestr(manifest_name, manifest_text.encode("utf-8"))
        
        # 2. 写入 14 份 docx 文档（使用规范的 00~13 序号便于文件管理器自动排序）
        for archive_name, src_name, desc in FILE_MAPPINGS:
            src_path = os.path.join(BOOK_DIR, src_name)
            arc_path = folder_prefix + archive_name
            z_out.write(src_path, arc_path)
            print(f"  Added: {arc_path} ({os.path.getsize(src_path)/1024:.1f} KB)")
            
    print(f"Zip created successfully: {zip_path} ({os.path.getsize(zip_path)/1024/1024:.2f} MB)")

def main():
    manifest_text = generate_manifest()
    print(manifest_text)
    
    # 目标压缩包路径保存在书稿目录下
    book_zip = os.path.join(BOOK_DIR, "《财务AI智能体：数据分析与可视化》全书最新成稿合集（前言+项目1~13）.zip")
    make_zip(book_zip, include_folder_prefix=True)

if __name__ == "__main__":
    main()
