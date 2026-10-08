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
import html as _html

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


# ───────────────────────────────── 段落重建（保留 pPr / paraId / 加粗）
def _rpr_of(run):
    m = re.search(r'<w:rPr>.*?</w:rPr>', run, re.S)
    return m.group(0) if m else None


def _make_bold(rpr):
    if not rpr or '<w:b/>' in rpr:
        return rpr
    if '</w:rFonts>' in rpr:
        return rpr.replace('</w:rFonts>', '</w:rFonts><w:b/><w:bCs/>', 1)
    return rpr.replace('<w:rPr>', '<w:rPr><w:b/><w:bCs/>', 1)


def _make_plain(rpr):
    return re.sub(r'<w:b/>|<w:bCs/>', '', rpr) if rpr else rpr


def build_para(seg, parts):
    """按 [(文本, 是否加粗), ...] 重建整段，保留原 pPr / paraId / 字体属性"""
    pid = re.search(r'w14:paraId="[^"]*"', seg)
    pid = ' ' + pid.group(0) if pid else ''
    ppr = re.search(r'<w:pPr>.*?</w:pPr>', seg, re.S)
    ppr = ppr.group(0) if ppr else ''
    b_rpr = p_rpr = None
    for r in re.findall(r'<w:r>.*?</w:r>', seg, re.S):
        rpr = _rpr_of(r)
        if not rpr:
            continue
        if re.search(r'<w:b/>', rpr):
            b_rpr = b_rpr or rpr
        else:
            p_rpr = p_rpr or rpr
    if p_rpr is None:
        p_rpr = _make_plain(b_rpr) if b_rpr else '<w:rPr>%s</w:rPr>' % DEFAULT_RPR
    if b_rpr is None:
        b_rpr = _make_bold(p_rpr)
    out = ['<w:p%s>%s' % (pid, ppr)]
    for text, bold in parts:
        out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                   % (b_rpr if bold else p_rpr, html.escape(text, quote=False)))
    out.append('</w:p>')
    return ''.join(out)


def rewrite_para(raw, needle, parts):
    """按纯文本定位段落并整体重建；返回 (新 raw, 段索引)"""
    i, a, b = find_para(raw, needle, exact=False)
    return raw[:a] + build_para(raw[a:b], parts) + raw[b:], i


def prepend_paras(raw, anchor, paras_xml, exact=True, nth=0):
    """在 anchor 段落之前插入若干新段落；返回 (新 raw, 段索引)"""
    i, a, b = find_para(raw, anchor, exact=exact, nth=nth)
    return raw[:a] + ''.join(paras_xml) + raw[a:], i


# ───────────────────────── 标点/间距规范化（第 10 轮新增）
_T_RE = re.compile(r'(<w:t(?: [^>]*)?>)(.*?)(</w:t>)', re.S)
_CJK = '\u4e00-\u9fff'


def _fix_text(s):
    s = s.replace('／', '/')          # 全角斜杠 → 半角
    s = s.replace('**', '')           # 去掉 Markdown 粗体标记（Word 里会原样显示）
    s = re.sub(r'(?<=[A-Za-z]) (?=[' + _CJK + r'])', '', s)   # 拉丁 + 空格 + 中文
    s = re.sub(r'(?<=[' + _CJK + r']) (?=[A-Za-z])', '', s)   # 中文 + 空格 + 拉丁
    return s


def normalize_punct_docx(raw, verbose=False):
    """在 <w:t> 文本节点内做标点/间距规范化，保留 run 与格式。
    规则（依据全书 28 册语料统计：中文紧邻拉丁 5721 处 vs 空格 25 处；半角斜杠 651 vs 全角 31）：
      1) 全角斜杠「／」→ 半角「/」
      2) 删除 Markdown 标记「**」
      3) 中英文之间不留空格（数字与中文之间的空格保留：多为表格步骤编号「1 新建…」）
    返回 (新 raw, 改动处数)。
    """
    n = 0

    def repl(m):
        nonlocal n
        old = m.group(2)
        new = _fix_text(old)
        if new != old:
            n += 1
        return m.group(1) + new + m.group(3)

    out = []
    for a, b in para_spans(raw):
        out.append(raw[a:b])
    # 逐段处理（避免误伤表格边框等非文本内容——只替换 <w:t> 内部）
    res, last = [], 0
    for a, b in para_spans(raw):
        res.append(raw[last:a])
        res.append(_T_RE.sub(repl, raw[a:b]))
        last = b
    res.append(raw[last:])
    new_raw = ''.join(res)
    if verbose:
        print('  normalize_punct：改动 %d 个文本节点' % n)
    return new_raw, n


# ───────────────────────── 跨 run 文本节点替换（保留原有格式）
def replace_text_nodes(raw, old, new, count=0):
    """在 <w:t> 文本节点层面替换字符串，支持跨节点命中（保留 run 与格式）。
    count=0 表示全文替换。返回 (新 raw, 实际替换处数)。"""
    out, last, done = [], 0, 0
    for a, b in para_spans(raw):
        seg = raw[a:b]
        out.append(raw[last:a])
        while True:
            if count and done >= count:
                break
            hits = list(_T_RE.finditer(seg))
            chars, owner, offs = [], [], []
            for k, m in enumerate(hits):
                txt = _html.unescape(m.group(2))
                for j, ch in enumerate(txt):
                    chars.append(ch); owner.append(k); offs.append(j)
            text = ''.join(chars)
            i = text.find(old)
            if i < 0:
                break
            k0, k1 = owner[i], owner[i + len(old) - 1]
            o0, o1 = offs[i], offs[i + len(old) - 1]
            esc = lambda x: _html.escape(x, quote=False)
            rebuild = {}
            for k in range(k0, k1 + 1):
                t = _html.unescape(hits[k].group(2))
                if k0 == k1:
                    rebuild[k] = esc(t[:o0] + new + t[o1 + 1:])
                elif k == k0:
                    rebuild[k] = esc(t[:o0] + new)
                elif k == k1:
                    rebuild[k] = esc(t[o1 + 1:])
                else:
                    rebuild[k] = ''
            for k in range(k1, k0 - 1, -1):          # 倒序替换，保持偏移有效
                m = hits[k]
                seg = seg[:m.start(2)] + rebuild[k] + seg[m.end(2):]
            done += 1
        out.append(seg)
        last = b
    out.append(raw[last:])
    return ''.join(out), done

# ───────────────────── 引号无关匹配 / 段内替换（第 12 轮新增）
_QMAP = str.maketrans({'“': '"', '”': '"', '‘': "'", '’': "'"})


def norm_quotes(s):
    """把弯引号统一成直引号，只用于匹配（长度不变，便于按位置回填）"""
    return s.translate(_QMAP)


def find_para_loose(raw, needle, nth=0):
    """按段落纯文本定位，**引号方向无关**；返回 (索引, 起, 止)"""
    hit = 0
    n = norm_quotes(needle)
    for i, (a, b) in enumerate(para_spans(raw)):
        if n in norm_quotes(para_text(raw, a, b).strip()):
            if hit == nth:
                return i, a, b
            hit += 1
    raise ValueError('未找到段落（引号无关）：' + needle[:60])


def replace_in_span(raw, a, b, old, new):
    """在 [a,b) 段内做跨 <w:t> 的替换，**保留 needle 之外的原文与 run 格式**。
    needle 与段内文字按引号方向无关匹配。返回新 raw。"""
    seg = raw[a:b]
    hits = list(_T_RE.finditer(seg))
    chars, owner, offs = [], [], []
    for k, m in enumerate(hits):
        txt = _html.unescape(m.group(2))
        for j, ch in enumerate(txt):
            chars.append(ch); owner.append(k); offs.append(j)
    text = ''.join(chars)
    i = norm_quotes(text).find(norm_quotes(old))
    if i < 0:
        raise ValueError('段内未找到：' + old[:50])
    k0, k1 = owner[i], owner[i + len(old) - 1]
    o0, o1 = offs[i], offs[i + len(old) - 1]
    esc = lambda x: _html.escape(x, quote=False)
    rebuild = {}
    for k in range(k0, k1 + 1):
        t = _html.unescape(hits[k].group(2))
        if k0 == k1:
            rebuild[k] = esc(t[:o0] + new + t[o1 + 1:])
        elif k == k0:
            rebuild[k] = esc(t[:o0] + new)
        elif k == k1:
            rebuild[k] = esc(t[o1 + 1:])
        else:
            rebuild[k] = ''
    for k in range(k1, k0 - 1, -1):          # 倒序替换，保持偏移有效
        m = hits[k]
        seg = seg[:m.start(2)] + rebuild[k] + seg[m.end(2):]
    return raw[:a] + seg + raw[b:]


def patch_para(raw, needle, new):
    """needle 恰好等于整段 → 整段重建；否则**段内替换**（保留其余原文字，不再吞段）。
    返回 (新 raw, 模式)。这就是「needle 取段中短语吞掉段首段尾」的通用修法。"""
    i, a, b = find_para_loose(raw, needle)
    t = para_text(raw, a, b).strip()
    if norm_quotes(t) == norm_quotes(needle):
        return raw[:a] + build_para(raw[a:b], [(new, False)]) + raw[b:], '整段'
    return replace_in_span(raw, a, b, needle, new), '段内'


def require_full_para(raw, needle, note):
    """guard：needle 必须等于整段，否则报错（防止整段重建时静默吞掉段首/段尾）"""
    i, a, b = find_para_loose(raw, needle)
    t = para_text(raw, a, b).strip()
    if norm_quotes(t) != norm_quotes(needle):
        raise ValueError('needle 不是整段（会吞掉段首/段尾）→ %s\n  段原文：%s' % (note, t[:80]))
    return i, a, b
