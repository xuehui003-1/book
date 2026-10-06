#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx 段落级读写工具 —— 供 check_docx.py 与各册修订脚本复用

设计要点：
  · 不依赖 python-docx，直接操作 word/document.xml
  · 段落定位以「纯文本精确/包含匹配」为准，避免被 run 拆分干扰
  · 文字替换优先在单个 <w:t> 内进行；跨 run 的用整段重建
"""
import re, html

NS_P = re.compile(r'<w:p[ >]')
NS_T = re.compile(r'<w:t(?: [^>]*)?>(.*?)</w:t>', re.S)

# 默认正文 run 属性（Times New Roman + 微软雅黑 + 小四）
DEFAULT_RPR = ('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"'
               ' w:eastAsia="微软雅黑"/><w:sz w:val="24"/>')


def para_spans(raw):
    """返回所有段落的 (起, 止) 位置列表"""
    return [(m.start(), m.end()) for m in re.finditer(r'<w:p[ >].*?</w:p>', raw, re.S)]


def para_text(raw, a, b):
    """抽取 [a,b) 区间内所有 <w:t> 的纯文本（已反转义）"""
    return html.unescape(''.join(NS_T.findall(raw[a:b])))


def paras(raw):
    """返回 [(索引, 起, 止, 纯文本)]"""
    return [(i, a, b, para_text(raw, a, b)) for i, (a, b) in enumerate(para_spans(raw))]


def find_para(raw, needle, exact=True, nth=0):
    """按段落纯文本定位；返回 (索引, 起, 止)"""
    hit = 0
    for i, (a, b) in enumerate(para_spans(raw)):
        t = para_text(raw, a, b).strip()
        ok = (t == needle) if exact else (needle in t)
        if ok:
            if hit == nth:
                return i, a, b
            hit += 1
    raise ValueError('未找到段落：' + needle[:60])


def find_all_para(raw, needle, exclude=''):
    """返回所有含 needle 的段落索引（可选排除含 exclude 的段落）"""
    out = []
    for i, (a, b) in enumerate(para_spans(raw)):
        t = para_text(raw, a, b)
        if needle in t and (not exclude or exclude not in t):
            out.append(i)
    return out


def find_in_t(raw, needle, nth=0):
    """在单个 <w:t> 文本节点内定位；返回 (起, 止)"""
    hit = 0
    for m in re.finditer(r'<w:t(?: [^>]*)?>(.*?)</w:t>', raw, re.S):
        if needle in m.group(1):
            if hit == nth:
                off = m.group(1).index(needle)
                return m.start(1) + off, m.start(1) + off + len(needle)
            hit += 1
    raise ValueError('未在任何 <w:t> 内找到：' + needle[:60])


def has_in_t(raw, needle):
    try:
        find_in_t(raw, needle)
        return True
    except ValueError:
        return False


def std_rpr(raw):
    """从文中挑一个标准正文 run 属性作为模板"""
    m = re.search(r'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"'
                  r' w:eastAsia="微软雅黑"/><w:sz w:val="24"/>', raw)
    return m.group(0) if m else DEFAULT_RPR


def mk_run(text, rpr=None):
    return ('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
            % (rpr or DEFAULT_RPR, html.escape(text, quote=False)))


def mk_para(text, style_id, rpr=None, pid=None, jc=None):
    """新建段落；style_id 为 w:pStyle 的值"""
    attrs = ' w14:paraId="%s"' % pid if pid else ''
    ppr = '<w:pPr>'
    if style_id:
        ppr += '<w:pStyle w:val="%s"/>' % style_id
    if jc:
        ppr += '<w:jc w:val="%s"/>' % jc
    ppr += '</w:pPr>'
    return '<w:p%s>%s%s</w:p>' % (attrs, ppr, mk_run(text, rpr))


def rebuild_para(seg, newtext, rpr=None):
    """整段重建，保留原 pPr（样式/缩进/对齐）与 paraId"""
    pid = re.search(r'w14:paraId="[^"]*"', seg)
    pid = ' ' + pid.group(0) if pid else ''
    ppr = re.search(r'<w:pPr>.*?</w:pPr>', seg, re.S)
    ppr = ppr.group(0) if ppr else ''
    return '<w:p%s>%s%s</w:p>' % (pid, ppr, mk_run(newtext, rpr))


def set_center(seg):
    """把段落设为居中，并去掉首行缩进"""
    seg = re.sub(r'<w:ind [^>]*w:firstLine[^>]*/>', '', seg)
    if '<w:jc ' in seg:
        seg = re.sub(r'<w:jc w:val="[^"]*"/>', '<w:jc w:val="center"/>', seg, count=1)
    else:
        m = re.search(r'<w:pPr>', seg)
        if m:
            seg = seg[:m.end()] + '<w:jc w:val="center"/>' + seg[m.end():]
        else:
            seg = seg.replace('<w:p>', '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>', 1)
    return seg


def para_style(raw, a, b):
    m = re.search(r'w:pStyle w:val="([^"]+)"', raw[a:b])
    return m.group(1) if m else None


def delete_paras(raw, idxs):
    """删除指定索引的整段（从后往前删，保持索引有效）"""
    spans = para_spans(raw)
    for i in sorted(idxs, reverse=True):
        a, b = spans[i]
        raw = raw[:a] + raw[b:]
    return raw
