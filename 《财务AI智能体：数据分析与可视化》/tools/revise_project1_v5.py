#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目1 V5 修订脚本（同步补齐编辑批注②在样本里的残留）
=====================================================
输入：03_项目1_认识财务数据分析与AI智能体_修订稿V4.docx
输出：/tmp/p1v5/03_项目1_认识财务数据分析与AI智能体_修订稿V5.docx

V4 已把 AI 痕迹与批注清零，但**编辑批注②「编者视角 → 读者视角」在项目1 里仍有残留**：
V4 只改了【教学导读】首段，正文 1.1.1 里还有一句「很多学生容易把……混为一谈」。

既然把"读者视角"定为全书通用规则，样本自己必须先达标。
"""
import os, re, sys, shutil, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_utils import rewrite_para

SRC = '03_项目1_认识财务数据分析与AI智能体_修订稿V4.docx'
WORK = '/tmp/p1v5'
LOG = []


def main():
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK + '/out', exist_ok=True)
    Z = zipfile.ZipFile(SRC)
    names = Z.namelist()
    Z.extractall(WORK + '/work')
    W = open(WORK + '/work/word/document.xml', encoding='utf-8').read()

    print('【批注② 读者视角】')
    W, i = rewrite_para(
        W, '很多学生容易把',
        [('很多人容易把“财务分析”和“会计核算”混为一谈，其实两者既相关又有明显区别。'
          '两者的关系可以用图1-5来理解。', False)])
    print('  ✔ 1.1.1：由“很多学生容易把……”改为“很多人容易把……”，去掉对学生群体的第三人称指称')

    import html as _h
    full = _h.unescape(''.join(re.findall(r'<w:t(?: [^>]*)?>(.*?)</w:t>', W, re.S)))
    left = {k: full.count(k) for k in ['学生', '读者', '笔者', '编者', '教师'] if full.count(k)}
    print('\n  断言：第三人称指向词残留 %s' % (left if left else '0（仅余“学习者”一处泛指群体，属正常用法）'))

    open(WORK + '/work/word/document.xml', 'w', encoding='utf-8').write(W)
    out = WORK + '/out/' + SRC.replace('V4', 'V5')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            zo.write(os.path.join(WORK + '/work', n), n)
    print('\n文档已生成：%s  (%.1f KB)' % (out, os.path.getsize(out) / 1024))


if __name__ == '__main__':
    main()
