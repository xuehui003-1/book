#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
A 类审计：把项目1 的 7 条编辑批注，逐条套到**任意一册**上，并给出证据定位
==========================================================================
用法：
    python3 tools/audit_a_class.py 项目4..._V1.docx [更多 docx...]
    python3 tools/audit_a_class.py *.docx --md /tmp/A类审计.md

对应关系（原文见 `11_` 第一节）：
    A1 薛维君｜正文要去AI化              → 检测 AI 特征词 + 排比 + 模板对称
    A2 薛维君｜角度是编者不是读者        → 检测 学生/学习者/读者/笔者/编者/教师 + 自我指称
    A3 hwy   ｜图题居中                  → 检测 图题是否居中
    A4 hwy   ｜表题居中                  → 检测 表题是否居中
    A5 hwy   ｜本段分行不符合语法        → 检测「碎句成段」（连续短段落、句中断行）
    A6 hwy   ｜介绍略简洁                → 检测「薄节」（整节正文过少）
    A7 薛维君｜内容与标题相关性不大      → 检测「标题关键词在正文中零出现」（候选，需人工判）
输出：每条给出 判定（✔/✘/候选）+ 证据（段落号 + 原文片段）
"""
import argparse, html, re, sys, zipfile

CJK = r'\u4e00-\u9fff'
TERM = '。！？；：””）)'          # 句子/短语的正常结尾

AI_PAT = ['不仅仅是', '恰恰相反', '有力地', '宝贵的', '根本性', '深度解读', '敏锐洞察',
          '切实可行', '赋能', '全方位', '一站式', '值得注意的是', '总而言之', '不容忽视',
          '极大地', '显著地提升', '综上所述', '在数字化时代', '在当今', '随着人工智能技术',
          '为……提供了', '开启了新的篇章', '不可忽视', '至关重要', '举足轻重']

# A2：第三人称谈读者 + 编者视角自我指称（07_ §5.4）
READER_WORDS = ['学生', '学习者', '读者', '笔者', '编者', '教师', '同学们', '大家']
SELF_REF = ['本项目是', '本教材', '本书', '教材里', '本册', '本章将', '我们将']


def load(path):
    raw = zipfile.ZipFile(path).read('word/document.xml').decode('utf-8')
    tbls = [(m.start(), m.end()) for m in re.finditer(r'<w:tbl>.*?</w:tbl>', raw, re.S)]
    paras = []
    for m in re.finditer(r'<w:p[ >].*?</w:p>', raw, re.S):
        a, b = m.start(), m.end()
        txt = html.unescape(''.join(re.findall(r'<w:t(?: [^>]*)?>(.*?)</w:t>', m.group(0), re.S)))
        st = re.search(r'w:pStyle w:val="(\d+)"', m.group(0))
        jc = re.search(r'<w:jc w:val="([^"]+)"', m.group(0))
        paras.append({'i': len(paras), 'a': a, 'b': b,
                      't': txt.strip(), 'style': st.group(1) if st else '-',
                      'jc': jc.group(1) if jc else '',
                      'num': 'numPr' in m.group(0),
                      'in_tbl': any(x <= a < y for x, y in tbls)})
    return raw, paras


def exempt(hit, paras):
    """已判定为「合法保留」的命中（不是机械清零，逐条给理由）"""
    i, word, ctx = hit
    t = paras[i]['t'] if i < len(paras) else ''
    # 1) 引号内引用他人观点（如素质测评题干）里的口语词
    if word in ('大家', '同学们'):
        inside = False
        for m in re.finditer(r'“[^”]*”', t):
            if word in m.group(0):
                inside = True
                break
        if inside:
            return '引用观点原文，不算叙述视角'
    # 2) “读者”指分析报告的使用者，不是对本书读者的指称
    if word == '读者' and '报告' in t:
        return '指报告的读者，不是对本书读者的指称'
    # 3) 提示词模板 / 字段枚举被误判为排比
    if word == '排比句式' and (t.lstrip().startswith('“') or t.count('、') >= 3):
        return '提示词模板或字段枚举，非排比'
    return None


def is_heading(p):
    """三级标题（X.X.X）"""
    return re.match(r'^\d+\.\d+\.\d+\s*\S', p['t'])


def sections(paras):
    """返回 [(标题段, 正文段列表)]"""
    out, cur = [], None
    for p in paras:
        if is_heading(p):
            cur = [p, []]
            out.append(cur)
        elif cur is not None and p['t'] and not p['in_tbl']:
            if p['t'].startswith('【'):
                cur = None
            else:
                cur[1].append(p)
    return [(h, body) for h, body in out if h]


# ─────────────────────────── 各项检测
def a1(paras):
    hits = []
    for p in paras:
        for k in AI_PAT:
            if k in p['t']:
                hits.append((p['i'], k, p['t'][:60]))
    # 排比：一段内出现 3 个以上「四字词，」结构
    for p in paras:
        if len(re.findall(r'[，、][' + CJK + r']{4}[，、]', p['t'])) >= 3:
            hits.append((p['i'], '排比句式', p['t'][:60]))
    return hits


def a2(paras):
    hits = []
    for p in paras:
        for w in READER_WORDS:
            for m in re.finditer(w, p['t']):
                hits.append((p['i'], w, p['t'][max(0, m.start() - 18):m.start() + 18]))
        for w in SELF_REF:
            if w in p['t']:
                hits.append((p['i'], w, p['t'][:60]))
    return hits


def a34(paras, kind):
    pat = re.compile(r'^(图|表)\d+-\d+\s')
    bad = []
    for p in paras:
        if p['in_tbl'] or not pat.match(p['t']):
            continue
        if kind == '图' and not p['t'].startswith('图'):
            continue
        if kind == '表' and not p['t'].startswith('表'):
            continue
        if p['jc'] != 'center':
            bad.append((p['i'], p['jc'] or '无对齐', p['t'][:50]))
    return bad


LIST_RE = re.compile(r'^(\d+[.、)]|[（(]\d+[）)]|[一二三四五六七八九十]+[、.]|[-—·•●○])\s*\S')


def a5(paras):
    """碎句成段：连续 ≥3 个短段落，全部不以句末标点结尾，且都不是列表项/题注/对照式排版。
    真正的病灶是「一句话被硬拆成多行」（编辑批注⑤原话：本段分行不符合语法）。"""
    clusters, cur = [], []

    def flush():
        nonlocal cur
        if len(cur) >= 3 and all(x['t'][-1] not in TERM for x in cur) \
                and not any(x['t'].endswith('：') for x in cur):
            clusters.append([(x['i'], x['t'][:34]) for x in cur])
        cur = []

    for p in paras:
        t = p['t']
        skip = (not t or len(t) > 34 or p['in_tbl']
                or p['style'] in ('2', '4', '5', '13', '23', '24', '30')
                or LIST_RE.match(t) or p['num'] or t.startswith(('图', '表', '【')))
        if skip:
            flush()
            continue
        cur.append(p)
        if t[-1] in TERM:
            flush()
    flush()
    return clusters


def a6(paras, min_chars=220):
    thin = []
    for h, body in sections(paras):
        n = sum(len(x['t']) for x in body)
        if n < min_chars:
            thin.append((h['i'], h['t'][:44], n, len(body)))
    return thin


def a7(paras):
    """标题的 2-gram 与本节正文交集为空 → 候选（标题与内容相关性，编辑批注⑦）"""
    cand = []
    for h, body in sections(paras):
        title = re.sub(r'^\d+\.\d+\.\d+\s*', '', h['t'])
        tail = re.sub(r'^[^：:]{0,14}[：:]', '', title)
        core = re.sub(r'[^' + CJK + r']', '', tail)
        if len(core) < 4:
            continue
        grams = {core[i:i + 2] for i in range(len(core) - 1)}
        # 剔除通用词
        grams = {g for g in grams if g not in ('基本', '方法', '认识', '使用', '实现', '数据', '分析', '应用')}
        body_txt = ''.join(x['t'] for x in body)
        if len(body_txt) < 80 or not grams:
            continue
        if not any(g in body_txt for g in grams):
            cand.append((h['i'], h['t'][:44], '/'.join(sorted(grams)[:5])))
    return cand


def audit(path):
    raw, paras = load(path)
    res = {'A1': a1(paras), 'A2': a2(paras), 'A3': a34(paras, '图'), 'A4': a34(paras, '表'),
           'A5': a5(paras), 'A6': a6(paras), 'A7': a7(paras)}
    # 逐条判定「合法保留」
    kept = {}
    for k in ('A1', 'A2'):
        keep, real = [], []
        for h in res[k]:
            why = exempt(h, paras) if (k == 'A2' or h[1] == '排比句式') else None
            (keep if why else real).append(h if not why else (h[0], h[1], why))
        res[k], kept[k] = real, keep
    heads = len([p for p in paras if is_heading(p)])
    return res, {'段落': len(paras), '三级标题': heads}, kept


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('files', nargs='+')
    ap.add_argument('--md')
    a = ap.parse_args()
    lines = ['# A 类批注逐条审计（项目1 的 7 条编辑批注 → 全书各册）', '']
    for f in a.files:
        res, meta, kept = audit(f)
        print('\n' + '=' * 78)
        print('【%s】段落 %d，三级标题 %d 个' % (f, meta['段落'], meta['三级标题']))
        lines += ['## %s' % f, '', '段落 %d，三级标题 %d 个' % (meta['段落'], meta['三级标题']), '',
                  '| 批注 | 检查项 | 判定 | 证据 |', '|:--:|---|:--:|---|']
        names = {'A1': '正文去AI化', 'A2': '读者视角（不是编者）', 'A3': '图题居中',
                 'A4': '表题居中', 'A5': '段落分行（碎句成段）', 'A6': '介绍过简（薄节）',
                 'A7': '内容与标题相关性'}
        for k in ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7']:
            v = res[k]
            if not v and kept.get(k):
                kp = kept[k]
                verdict, ev = '○', '已判定保留 %d 处：%s' % (len(kp), kp[0][2])
                print('  %s %s：○ 无必改（%d 处已判定保留：%s）' % (k, names[k], len(kp), kp[0][2]))
                lines.append('| %s | %s | ○ | %s |' % (k, names[k], ev))
                continue
            if not v:
                verdict, ev = '✔', '—'
                print('  %s %s：✔ 无命中' % (k, names[k]))
            else:
                verdict = {'A6': '候选', 'A7': '候选'}.get(k, '✘')
                ev = '；'.join(str(x) for x in v[:3])
                print('  %s %s：%s %d 处' % (k, names[k], verdict, len(v)))
                for x in v[:5]:
                    print('        %s' % (x,))
            lines.append('| %s | %s | %s | %s |' % (k, names[k], verdict, ev.replace('|', '/')))
        lines.append('')
    if a.md:
        open(a.md, 'w', encoding='utf-8').write('\n'.join(lines))
        print('\n报告已写出：%s' % a.md)


if __name__ == '__main__':
    main()
