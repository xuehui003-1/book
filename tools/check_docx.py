#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
书稿自检脚本 —— 教材 .docx 稿件机械类问题自动检查
用法：  python3 tools/check_docx.py 稿件1.docx [稿件2.docx ...]
        python3 tools/check_docx.py *.docx --md 报告.md

检查项：
  C1 题注居中      图题/表题是否居中
  C2 题注空格      图1-x / 表1-x 与题名之间是否有空格
  C3 图题表题配对  是否存在只有图无题注 / 只有引用无题注
  C4 图表编号连续  图1-1..1-n / 表1-1..1-n 是否连续无缺号
  C5 悬空交叉引用  正文引用了不存在的图号/表号
  C6 引号配对      左引号“与右引号”数量是否相等
  C7 引号误用      右引号”被当作左引号使用（高频错法）
  C8 英文直引号    正文中残留的 " 直引号
  C9 中文弯引号    引号是否统一使用中文弯引号
  C10 残留批注     文档中是否还挂着未处理的批注
  C11 AI行文特征   可疑的AI写作特征（提示，非错误）
"""
import sys, re, os, zipfile, html, argparse

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS_P = re.compile(r'<w:p[ >]')
NS_T = re.compile(r'<w:t(?: [^>]*)?>(.*?)</w:t>', re.S)

CJK = r'\u4e00-\u9fff'
CAP_RE = re.compile(r'^(图|表)(\d+)-(\d+)\s*(.*)$')

def unesc(s):
    return html.unescape(s)

def read_docx(path):
    with zipfile.ZipFile(path) as z:
        doc = z.read('word/document.xml').decode('utf-8')
        try: com = z.read('word/comments.xml').decode('utf-8')
        except KeyError: com = ''
    return doc, com

def paragraphs(raw):
    out = []
    # 表格区间：表格单元格里的段落不能被当作题注（如实践报告模板里的行标题「图2-7对我的启发」）
    tbl = [(m.start(), m.end()) for m in re.finditer(r'<w:tbl>.*?</w:tbl>', raw, re.S)]
    for m in re.finditer(r'<w:p[ >].*?</w:p>', raw, re.S):
        seg = raw[m.start():m.end()]
        txt = unesc(''.join(NS_T.findall(seg)))
        jc = re.search(r'<w:jc w:val="([^"]+)"', seg)
        st = re.search(r'w:pStyle w:val="([^"]+)"', seg)
        out.append(dict(seg=seg, text=txt, jc=jc.group(1) if jc else None,
                        style=st.group(1) if st else None,
                        in_table=any(a <= m.start() < b for a, b in tbl)))
    return out

# ─────────────────────────────────────────────── 检查
def check(path):
    raw, com = read_docx(path)
    paras = paragraphs(raw)
    issues = []      # (级别, 编号, 说明)

    def add(lv, code, msg):
        issues.append((lv, code, msg))

    # ---- 收集题注与引用
    caps = []        # (类型, 项目号, 编号, 段号, 全文, para)
    BADPUNCT = '。？！，；：'
    for i, p in enumerate(paras):
        if p['in_table']:      # 表格内的段落不是题注
            continue
        t = p['text'].strip()
        if len(t) > 45:
            continue
        m = CAP_RE.match(t)
        if not m:
            continue
        rest = m.group(4).strip()
        # 题名不会带句末标点；排除“表6-2列出了……”这类正文引用句
        if not rest or any(c in rest for c in BADPUNCT) or len(rest) > 32:
            continue
        caps.append((m.group(1), int(m.group(2)), int(m.group(3)), i, t, p))

    # C1 / C2
    bad_center, bad_space = [], []
    for kind, proj, num, i, t, p in caps:
        if p['jc'] != 'center':
            bad_center.append(f'{kind}{proj}-{num}')
        if not re.match(r'^(图|表)\d+-\d+ +', t):
            bad_space.append(t)
    if bad_center:
        add('必改', 'C1', f'题注未居中 {len(bad_center)} 处：' + '、'.join(bad_center))
    if bad_space:
        add('必改', 'C2', f'题注缺空格 {len(bad_space)} 处：' + '；'.join(bad_space))

    # C3 / C4 / C5
    have = {}
    for kind, proj, num, i, t, p in caps:
        have.setdefault(kind, set()).add(num)
    refs = {}
    for i, p in enumerate(paras):
        t = p['text']
        is_cap = any(t.strip() == c[4] for c in caps) if t.strip() else False
        for m in re.finditer(r'(图|表)\d+-(\d+)', t):
            if is_cap and m.start() == 0: continue
            refs.setdefault(m.group(1), {}).setdefault(int(m.group(2)), 0)
            refs[m.group(1)][int(m.group(2))] += 1
    for kind in ('图', '表'):
        hs, rs = have.get(kind, set()), set(refs.get(kind, {}))
        missing = sorted(rs - hs)          # 引用了但无题注
        unused = sorted(hs - rs)           # 有题注但正文未引用
        if missing:
            add('必改', 'C5', f'悬空交叉引用：{kind}1-' + f'、{kind}1-'.join(map(str, missing)) + ' 被正文引用但无对应题注')
        if hs:
            mx = max(hs)
            gaps = [n for n in range(1, mx + 1) if n not in hs]
            if gaps:
                add('必改', 'C4', f'{kind}编号不连续，缺 ' + '、'.join(f'{kind}1-{g}' for g in gaps))
        if unused:
            add('提示', 'C3', f'{kind}题注存在但正文未引用：' + '、'.join(f'{kind}1-{u}' for u in sorted(unused)))

    # C6 / C7 / C8
    full = ''.join(p['text'] for p in paras)
    nl, nr = full.count('“'), full.count('”')
    if nl != nr:
        add('必改', 'C6', f'引号不配对：左引号“ {nl} 个、右引号” {nr} 个，相差 {abs(nl - nr)}')
    # 全局流式交替检测：按出现顺序，第 1/3/5… 个应为 “，第 2/4/6… 个应为 ”。
    # 段内「必须自配对」会把跨段引号（如多行提示词模板）误判，故改为全局判定。
    misuse = 0
    first_bad = None
    expect_open = True
    for p in paras:
        for ch in p['text']:
            if ch not in '“”':
                continue
            want = '“' if expect_open else '”'
            if ch != want:
                misuse += 1
                if first_bad is None:
                    first_bad = p['text'].strip()[:34]
            expect_open = not expect_open
    if misuse:
        add('必改', 'C7', f'引号开合次序异常 {misuse} 处，首例：{first_bad}')
    if '"' in full:
        add('必改', 'C8', f'正文残留英文直引号 {full.count(chr(34))} 处，应改为中文弯引号')

    # C10 残留批注
    if com:
        n = len(re.findall(r'<w:comment ', com))
        if n:
            authors = re.findall(r'w:author="([^"]*)"', com)
            add('提示', 'C10', f'文档中仍有 {n} 条未处理批注（批注人：' + '、'.join(sorted(set(authors))) + '）')

    # C11 AI 行文特征
    AI_PAT = ['不仅仅是', '恰恰相反', '有力地', '宝贵的', '根本性', '深度解读', '敏锐洞察',
              '切实可行', '赋能', '全方位', '一站式', '值得注意的是', '总而言之', '不容忽视',
              '在……的背景下', '随着……的不断', '极大地', '显著地提升']
    hits = {k: full.count(k) for k in AI_PAT if full.count(k)}
    # 排比检测：连续出现 3 个以上「修饰词+动词+四字宾语」
    para_hits = []
    for p in paras:
        t = p['text']
        for kw in ['精准判断', '深度解读', '敏锐洞察', '切实可行']:
            if kw in t and t not in para_hits: para_hits.append(t.strip()[:40])
    if hits or len(para_hits) >= 2:
        msg = '、'.join(f'{k}×{v}' for k, v in list(hits.items())[:8])
        add('提示', 'C11', f'疑似AI行文特征词：{msg}' + (f'；密集排比段 {len(para_hits)} 处' if para_hits else ''))

    return issues, dict(图=len(have.get('图', [])), 表=len(have.get('表', [])),
                        段落=len(paras), 批注=len(re.findall(r'<w:comment ', com)) if com else 0)

# ─────────────────────────────────────────────── 输出
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('files', nargs='+')
    ap.add_argument('--md', help='同时输出 Markdown 报告')
    args = ap.parse_args()
    lines, md = [], ['# 书稿自检报告', '']
    total_bad = 0
    for f in args.files:
        if not os.path.exists(f):
            print(f'跳过（不存在）：{f}'); continue
        try:
            issues, stat = check(f)
        except Exception as e:
            print(f'✗ {os.path.basename(f)} 解析失败：{e}'); continue
        must = [x for x in issues if x[0] == '必改']
        total_bad += len(must)
        head = f'\n【{os.path.basename(f)}】 图{stat["图"]} 表{stat["表"]} 段落{stat["段落"]} 批注{stat["批注"]}'
        lines.append(head)
        md.append(f'## {os.path.basename(f)}')
        md.append(f'- 图 {stat["图"]} / 表 {stat["表"]} / 段落 {stat["段落"]} / 残留批注 {stat["批注"]}')
        if not issues:
            lines.append('  ✔ 未发现机械类问题')
            md.append('- ✔ 未发现机械类问题')
        for lv, code, msg in issues:
            mark = '✘' if lv == '必改' else '△'
            lines.append(f'  {mark} [{code}] {msg}')
            md.append(f'- {mark} **[{code}] {msg}**' if lv == '必改' else f'- {mark} [{code}] {msg}')
        md.append('')
        lines.append(f'  → 必改 {len(must)} 项，提示 {len(issues)-len(must)} 项')
    lines.append(f'\n{"="*70}\n合计必改问题：{total_bad} 项')
    out = '\n'.join(lines)
    print(out)
    if args.md:
        open(args.md, 'w', encoding='utf-8').write('\n'.join(md) + f'\n\n**合计必改问题：{total_bad} 项**\n')
        print(f'\nMarkdown 报告已写入：{args.md}')

if __name__ == '__main__':
    main()
