#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
【小试牛刀】栏目重定位（修正 R8/R9 遗留缺陷）
============================================
缺陷：revise_project2_v1.py / revise_project3_v1.py 用 prepend_paras(anchor='【善学勤思】')
      插入 4 个【小试牛刀】，但未指定 nth，默认全部插到**第 1 个**【善学勤思】之前，
      导致 4 个栏目堆叠在任务2.1 / 任务3.1 末尾，而不是每任务末尾各 1 个。
样板：项目1（V6）——每个任务末尾各 1 个【小试牛刀】，紧接其后的【善学勤思】之前。

本脚本：把堆叠块按原顺序拆成 4 块，依次移到第 1~4 个【善学勤思】之前。
输出：项目2 V3 → V4；项目3 V1 → V2
"""
import os, sys, shutil, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_utils import para_spans, para_text, find_para

JOBS = [
    ('项目2 搭建AI数据分析工具箱_修订稿V3.docx',
     '项目2 搭建AI数据分析工具箱_修订稿V4.docx', '/tmp/p2v4'),
    ('项目3 多维表格与智能数据管理_修订稿V1.docx',
     '项目3 多维表格与智能数据管理_修订稿V2.docx', '/tmp/p3v2'),
]
HEAD = '【小试牛刀】'
ANCHOR = '【善学勤思】'


def relocate(raw, n_task=4):
    spans = para_spans(raw)
    first = None
    for i, (a, b) in enumerate(spans):
        if para_text(raw, a, b).strip().startswith(HEAD):
            first = i
            break
    if first is None:
        raise ValueError('未找到【小试牛刀】')
    # 该栏目后第一个【善学勤思】
    end = None
    for i in range(first, len(spans)):
        if para_text(raw, *spans[i]).strip() == ANCHOR:
            end = i
            break
    # 拆块：每个【小试牛刀】标题 + 其后正文，直到下一个标题 / 善学勤思
    blocks, cur = [], None
    for i in range(first, end):
        a, b = spans[i]
        if para_text(raw, a, b).strip().startswith(HEAD):
            cur = []
            blocks.append(cur)
        cur.append(raw[a:b])
    # 完整性校验：文档内【小试牛刀】总数必须等于块数（即确实全部堆叠在一起）
    total = sum(1 for a, b in spans if para_text(raw, a, b).strip().startswith(HEAD))
    assert len(blocks) == n_task == total, '块数异常：%d 块 / 全文 %d 个' % (len(blocks), total)
    # 摘除堆叠区
    a0 = spans[first][0]
    b0 = spans[end - 1][1]
    raw = raw[:a0] + raw[b0:]
    # 依次插到第 k 个【善学勤思】之前
    for k, blk in enumerate(blocks):
        _, a, b = find_para(raw, ANCHOR, nth=k)
        raw = raw[:a] + ''.join(blk) + raw[a:]
    return raw, len(blocks)


def main():
    for src, dst, work in JOBS:
        shutil.rmtree(work, ignore_errors=True)
        os.makedirs(work + '/out', exist_ok=True)
        Z = zipfile.ZipFile(src)
        names = Z.namelist()
        Z.extractall(work + '/work')
        f = work + '/work/word/document.xml'
        raw = open(f, encoding='utf-8').read()
        raw, n = relocate(raw)
        open(f, 'w', encoding='utf-8').write(raw)
        out = work + '/out/' + dst
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
            for m in names:
                zo.write(os.path.join(work + '/work', m), m)
        # 复检位置
        spans = para_spans(raw)
        pos = [(i, para_text(raw, a, b).strip()[:34]) for i, (a, b) in enumerate(spans)
               if para_text(raw, a, b).strip().startswith((HEAD, ANCHOR, '任务'))]
        print('【%s】重定位 %d 块 → %s (%.1f KB)' % (src, n, out, os.path.getsize(out) / 1024))
        for i, t in pos:
            print('    [%4d] %s' % (i, t))
        print()


if __name__ == '__main__':
    main()
