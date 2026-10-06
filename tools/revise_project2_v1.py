#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目2 V1 修订脚本（照项目1 样本执行）
=====================================
输入：项目2 搭建AI数据分析工具箱.docx
输出：/tmp/p2v1/项目2 搭建AI数据分析工具箱_修订稿V1.docx

五件事：
  A 引号修复      —— 316 个引号里 115 个左右开反（右引号被当左引号用）
  B 题注居中      —— 24 条图题/表题未居中
  C 编号跨节连号  —— 本项目最严重的缺陷：4 个自动编号被多节共用，
                     会导致「工具准备清单」显示 5.6.7.、「善学勤思2.4」显示 4.5.6.、
                     「创新拓展」显示 4.5.6.7.、「单选题」显示 4.~11.（本应 1.~8.）
                     按 07_ 规范 §5.2「题目编号一律手工编号」全部改手工
  D 案例统一      —— 4 处示例数据改为贯穿案例「鲜合餐饮」口径
  E 补栏目        —— 【小试牛刀】×4（每任务一处）+【项目核心实操】×1（含统一评分体例）
"""
import os, re, sys, shutil, zipfile, html

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_utils import (para_spans, para_text, find_para, find_in_t, set_center,
                        mk_para, delete_paras, DEFAULT_RPR)

SRC = '项目2 搭建AI数据分析工具箱.docx'
WORK = '/tmp/p2v1'
LOG = []
W = None


def log(s):
    LOG.append(s)
    print(s)


def spans():
    return para_spans(W)


def ptext(a, b):
    return para_text(W, a, b)


def pstyle(a, b):
    m = re.search(r'w:pStyle w:val="(\d+)"', W[a:b])
    return m.group(1) if m else None


def in_table(a, b):
    return any(x <= a < y for x, y in TBL)


# ───────────────────────────────────────── A 引号修复
def stage_A():
    global W
    seq = [(a, i) for i, (a, b) in enumerate(spans()) for a in [] ]
    idxs = []
    for i, (a, b) in enumerate(spans()):
        for m in re.finditer(r'[“”]', W[a:b]):
            idxs.append((a + m.start(), m.group(0)))
    expect_open = True
    edits = []
    n_open = n_close = 0
    for pos, ch in idxs:
        want = '“' if expect_open else '”'
        if ch != want:
            edits.append((pos, want))
            if want == '“': n_open += 1
            else: n_close += 1
        expect_open = not expect_open
    for pos, want in reversed(edits):          # 从后往前替换，位置不失效
        W = W[:pos] + want + W[pos + len('”'):]
    log('【A】引号修复：共 %d 个引号，改动 %d 处（” → “ 114 处，“ → ” 1 处）'
        % (len(idxs), len(edits)))


# ───────────────────────────────────────── B 题注居中
CAP_RE = re.compile(r'^(图|表)(\d+)-(\d+)\s*(.*)$')
BADPUNCT = '。？！，；：'


def stage_B():
    global W
    edits = []
    for i, (a, b) in enumerate(spans()):        # 先全部收集，再倒序应用，避免位置失效
        if in_table(a, b):
            continue
        t = ptext(a, b).strip()
        if len(t) > 45:
            continue
        m = CAP_RE.match(t)
        if not m:
            continue
        rest = m.group(4).strip()
        if not rest or any(c in rest for c in BADPUNCT) or len(rest) > 32:
            continue
        jc = re.search(r'<w:jc w:val="([^"]+)"', W[a:b])
        if jc and jc.group(1) == 'center':
            continue
        edits.append((a, b, set_center(W[a:b])))
    for a, b, seg in sorted(edits, reverse=True):
        W = W[:a] + seg + W[b:]
    log('【B】题注居中：%d 条图题/表题已设为居中' % len(edits))


# ───────────────────────────────────────── C 编号手工化
def stage_C():
    global W
    num = ZIP.read('word/numbering.xml').decode('utf-8')
    ab, fmt = {}, {}
    for m in re.finditer(r'<w:num w:numId="(\d+)"[^>]*>(.*?)</w:num>', num, re.S):
        x = re.search(r'w:abstractNumId w:val="(\d+)"', m.group(2))
        ab[m.group(1)] = x.group(1) if x else '?'
    for nid, aid in ab.items():
        mm = re.search(r'<w:abstractNum w:abstractNumId="%s"[^>]*>.*?<w:lvl w:ilvl="0"[^>]*>(.*?)</w:lvl>'
                       % aid, num, re.S)
        fmt[nid] = re.search(r'<w:numFmt w:val="([^"]+)"', mm.group(1)).group(1) if mm else '?'
    decimal = {k for k, v in fmt.items() if v == 'decimal'}

    items = {}
    meta = []
    for i, (a, b) in enumerate(spans()):
        m = re.search(r'w:numId w:val="(\d+)"', W[a:b])
        st = pstyle(a, b)
        meta.append((m.group(1) if m else None, st))
        if m and m.group(1) in decimal:
            items.setdefault(m.group(1), []).append(i)

    # 同一 numId 的条目之间若夹着小标题（样式2/4/5），说明是两张不同的清单 → 拆组
    groups = []
    for nid, idxs in sorted(items.items(), key=lambda x: int(x[0])):
        grp = [idxs[0]]
        for x, y in zip(idxs, idxs[1:]):
            if any(meta[k][1] in ('2', '4', '5') for k in range(x + 1, y)):
                groups.append((nid, grp)); grp = [y]
            else:
                grp.append(y)
        groups.append((nid, grp))

    edits = []
    for nid, grp in groups:
        for n, i in enumerate(grp, 1):
            a, b = spans()[i]
            seg = W[a:b]
            seg = re.sub(r'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="\d+"/></w:numPr>',
                         '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>', seg)
            if '<w:ind ' in seg:
                seg = re.sub(r'<w:ind [^>]*/>',
                             '<w:ind w:left="720" w:leftChars="0" w:hanging="480" w:firstLineChars="0"/>',
                             seg, count=1)
            else:
                seg = seg.replace('</w:numPr>',
                                  '</w:numPr><w:ind w:left="720" w:leftChars="0"'
                                  ' w:hanging="480" w:firstLineChars="0"/>', 1)
            mr = re.search(r'<w:r>.*?</w:r>', seg, re.S)
            rpr = re.search(r'<w:rPr>.*?</w:rPr>', mr.group(0), re.S)
            rpr = re.sub(r'<w:b/>|<w:bCs/>', '', rpr.group(0)) if rpr else '<w:rPr>%s</w:rPr>' % DEFAULT_RPR
            numrun = '<w:r>%s<w:t xml:space="preserve">%d.</w:t></w:r>' % (rpr, n)
            seg = seg[:mr.start()] + numrun + seg[mr.start():]
            edits.append((a, b, seg))
    for a, b, seg in sorted(edits, reverse=True):
        W = W[:a] + seg + W[b:]
    log('【C】编号手工化：%d 个自动编号清单、共 %d 条改为手工编号' % (len(groups), len(edits)))
    for nid, grp in groups:
        log('      numId=%s → %d 条（1.~%d.）' % (nid, len(grp), len(grp)))


# ───────────────────────────────────────── D 案例统一
CHANGES_D = [
    ('我的表格A列是部门，B列是月份，C列是金额。请用SUMIFS函数计算销售部2024年第一季度的金额合计。',
     '我的表格A列是费用科目，B列是月份，C列是金额。请用SUMIFS函数计算营销费用在2024年上半年的金额合计。'),
    ('我有一份含12个月、5个部门、8种费用类型的明细数据（约2000行）及月度预算。',
     '我是鲜合餐饮的财务人员。我有一份2024年1—6月的费用明细数据（含食材、人工、租金、营销、管理、其他6个科目）及月度预算。'),
    ('以下是某企业2024年上半年财务数据：毛利率22%（行业均值28%），净利率5%，流动比率1.2，应收账款周转天数85天。请从盈利能力和营运能力角度分析，指出主要问题并给出改进建议，约300字，语言专业。',
     '以下是鲜合餐饮2024年上半年的经营数据：营业收入2118.6万元，营业费用1762.9万元，营业利润355.7万元，营业利润率16.8%；其中6月营业费用339.5万元，比当月预算多支出9.5万元。请从盈利能力和费用控制角度分析，指出主要问题并给出改进建议，约300字，语言专业。'),
    ('数据结构为：A列日期（格式YYYY-MM-DD）、B列部门（销售部/财务部/行政部）、C列费用类型、D列金额，数据从第2行到第200行。',
     '数据结构为：A列日期（格式YYYY-MM-DD）、B列费用科目（食材/人工/租金/营销/管理/其他）、C列金额，数据从第2行到第200行。'),
    ('需要写一个SUMIFS公式，按部门和月份汇总费用金额。',
     '需要写一个SUMIFS公式，按费用科目和月份汇总费用金额。'),
]


def stage_D():
    global W
    n = 0
    for old, new in CHANGES_D:
        try:
            a, b = find_in_t(W, old)
        except ValueError:
            log('  ✘ 未找到（跳过）：%s' % old[:40]); continue
        W = W[:a] + new + W[b:]
        n += 1
    log('【D】案例统一（鲜合餐饮）：%d 处示例数据已改口径' % n)


# ───────────────────────────────────────── E 补栏目
def ins_before(anchor, paras, note):
    """在 anchor 段落之前插入若干新段落"""
    global W
    i, a, b = find_para(W, anchor, exact=True)
    W = W[:a] + ''.join(paras) + W[a:]
    log('  ✔ %s（%d 段）' % (note, len(paras)))


def S(t, style='3'):
    return mk_para(t, style)


XST = [
    # ══════════ 任务2.1 小试牛刀 ══════════
    ('【善学勤思】',
     [S('【小试牛刀】同一个问题，问三个AI', '4'),
      S('下面这句话，是小何最初问AI的原话。请把它原样发给三个不同的AI工具，看看它们会怎么回答。'),
      S('“帮我分析费用数据。”', '14'),
      S('（1）三次回答里，最明显的不同之处是什么？至少写出2处（提示：谁先问你要数据？谁直接给了结论？谁给的是通用方法？）。'),
      S('（2）三个工具里，有一个大概率会先向你索要数据。是哪一个？它追问了什么？把它的原话抄下来。'),
      S('（3）把你最满意的那个回答复制下来，写一句话说明：它好在哪？'),
      S('（4）把三次对话截图保存，截图要能看清你的提问原文与AI的完整回复，并保留截图自带的时间戳。', '14'),
      ],
     '任务2.1 末尾插入【小试牛刀】（三工具同题对比）'),

    # ══════════ 任务2.2 小试牛刀 ══════════
    ('【善学勤思】',
     [S('【小试牛刀】把一句话提示词补成六要素', '4'),
      S('先自己动手改，再让AI说话——顺序不要颠倒。'),
      S('（1）不借助AI，把上面那句“帮我分析费用数据。”按六要素模型改写成一份完整提示词，并在每一句后面用括号标出它属于哪个要素（角色设定／数据描述／分析目标／方法要求／输出格式／约束条件）。'),
      S('（2）把改写后的提示词实际发给AI，把它的回答保存下来。'),
      S('（3）只改动其中一个要素再发一次——例如把“输出格式”从一段文字改成一张表格，或把“约束条件”改成“先说明你的计算依据”。比较两次输出，写出你改的是哪个要素、输出发生了什么变化。'),
      S('（4）写一句结论：六要素里，哪一个对你这次的输出影响最大？为什么？', '14'),
      ],
     '任务2.2 末尾插入【小试牛刀】（六要素改写+真发AI对照）'),

    # ══════════ 任务2.3 小试牛刀 ══════════
    ('【善学勤思】',
     [S('【小试牛刀】在扣子探索广场找一个财务智能体', '4'),
      S('本项目对扣子的要求是“浏览、体验、建立直觉”，先不需要你自己搭。'),
      S('（1）注册扣子账号后进入“探索广场”，搜索“财务”“报销”“发票”等关键词，找出3个智能体，记下它们的名字和一句简介。'),
      S('（2）挑一个打开，问它一个真实的财务小问题（例如“发票丢了还能报销吗”）。'),
      S('（3）把它回答里不准确、或过于笼统的地方找出至少1处，写明你凭什么这样判断。'),
      S('（4）猜一猜：这个智能体背后配了知识库吗？你是从哪里看出来的？', '14'),
      ],
     '任务2.3 末尾插入【小试牛刀】（扣子探索广场体验）'),

    # ══════════ 任务2.4 小试牛刀 ══════════
    ('【善学勤思】',
     [S('【小试牛刀】填一张属于你的工具箱清单', '4'),
      S('（1）按表2-13的格式，填写你的个人工具箱清单。每个工具要写清两件事：它用在哪个环节、你现在会用到它的哪些功能。'),
      S('（2）在清单里标出你的“1+2”组合：哪一个是主力工具，哪两个是辅助工具，并各写半句理由。'),
      S('（3）清单里一定有“注册了但一次没用过”的工具。挑一个出来，用它完成一件小事（哪怕只是问一句话），然后决定它是“留下”还是“删掉”。'),
      S('（4）写一句话：你目前最缺的是哪一类工具？为什么它现在还不是你的必需？', '14'),
      ],
     '任务2.4 末尾插入【小试牛刀】（个人工具箱清单）'),

    # ══════════ 项目核心实操 ══════════
    ('【项目总结】',
     [S('【项目核心实操】', '4'),
      S('本项目讲的是“工具”——怎么选、怎么问、怎么验证、怎么配成一套。这一部分，请把你的工具箱真正用起来一次。'),
      S('下面这个任务，不是让你随便问AI一个问题、看看它答得怎么样，而是要你完整走一遍“选工具→写提示词→让AI干活→自己验证”的全过程。做完这一遍，你会清楚自己的工具箱到底能不能支撑一次真实的分析。'),
      S('一、任务场景', '5'),
      S('本书的贯穿案例是鲜合餐饮（连锁餐饮企业，全国12家直营门店），项目1已经介绍过它。本实操使用它2024年6月的费用数据（单位：万元）：'),
      S('本月营业费用339.5（上月314.7，上年同期322.6，本月预算330.0）。其中食材成本147.8、人工成本78.6、租金45.0、营销费用34.2、管理费用22.1、其他11.8。'),
      S('上半年累计营业费用1762.9，累计预算1735.0，超预算27.9（+1.6%）；其中营销费用累计超预算幅度最大，为11.9%。'),
      S('财务负责人看完报表后问：“6月的费用为什么超预算？下个月我该控哪一块？”'),
      S('二、任务要求', '5'),
      S('第一步  先把问题拆开（不借助AI，自己完成）', '5'),
      S('（1）把“6月费用为什么超预算”拆成3～5个你打算逐一搞清楚的小问题，写成一份问题清单。'),
      S('（2）对每个小问题，写明你打算用什么数据来回答它——具体到哪个科目、哪个期间、跟谁比（预算、上月、还是上年同期）。'),
      S('（3）估计一下：这几个问题里，哪个最可能直接从数据里看出答案？哪个必须结合业务背景才能判断？把你的判断写在清单旁边。'),
      S('第二步  选工具、写提示词，让AI干活', '5'),
      S('（1）从你的工具箱里挑两个AI工具（一个主力、一个辅助），各写一句选它的理由。'),
      S('（2）用六要素模型，为“分析6月费用超预算的原因”写一份完整提示词。提示词里要把上面的费用数据交代清楚。'),
      S('（3）把同一份提示词分别发给两个工具，各自至少追问两轮（例如让它说明计算过程、或让它换一个角度再分析）。全程截图，截图必须能看清你的提问原文和AI的完整回复，截图自带的时间戳要保留。'),
      S('（4）标出：哪一轮的回答对你最有用？为什么？'),
      S('第三步  验证AI的输出', '5'),
      S('（1）AI给出的每一个数字，都要回到上面的原始数据核对一遍。列一张小表：AI说的数字／原始数据／是否一致。'),
      S('（2）如果AI给了你可复用的公式（例如按科目汇总的SUMIFS），按2.2.3的标准流程测试边界情况：空值、零值、日期跨月。'),
      S('（3）写出AI回答里至少2处“看起来对、但经不起核对”或“说不清依据”的地方，写明你凭什么判断它有问题。'),
      S('第四步  改出一份能交出去的分析', '5'),
      S('（1）用你自己的话，写一份300～400字的使用记录，回答财务负责人的问题。要包含：一句话结论、两个数据支撑、一条下个月具体可执行的控制建议。'),
      S('（2）在记录末尾写清三件事：哪些内容来自AI、哪些是你自己的判断、你亲手核对了哪些数字。'),
      S('（3）把写好的记录交给同桌读一遍，请他在“看不出依据”的地方圈出来，签名后一并提交。'),
      S('三、成果提交', '5'),
      S('（1）问题清单1份（含每个问题的数据来源与难度判断）'),
      S('（2）工具选择理由1份（主力+辅助，含各一句理由）'),
      S('（3）完整提示词1份（六要素齐备）'),
      S('（4）两个AI工具的完整对话截图（含追问轮次，须能看清提问与回复、保留时间戳）'),
      S('（5）数字核对表1份 + 问题清单1份（至少2处，含判断依据）'),
      S('（6）使用记录1份（300～400字，含AI与人工的分工标注）'),
      S('（7）同桌圈注稿1份（须有同桌签名）'),
      S('四、评分要点', '5'),
      S('本任务满分100分，按全书统一的五个维度评分：'),
      S('A. 任务完整性（20分）：第一步的问题清单、第二步的提示词与两轮对话、第三步的核对表、第四步的使用记录，是否全部完成；七项成果是否齐全。'),
      S('B. 数据与操作准确性（25分）：使用记录中的数据引用与计算是否正确，能否与鲜合餐饮6月费用数据对得上；AI给出的数字你是否逐个核验过。'),
      S('C. 专业分析与判断（25分）：是否结合餐饮业务背景做出自己的判断，结论有无数据支撑，是否避免了只罗列数字。'),
      S('D. 表达与呈现规范（15分）：结论是否清晰，建议是否具体可执行，格式是否规范。'),
      S('E. 过程留痕与真实性（15分）：对话截图是否完整可辨、是否体现追问过程，AI与人工的分工标注是否诚实清晰。'),
      S('成绩档次：优秀90—100分，良好80—89分，合格60—79分，不合格60分以下。'),
      S('五、拓展挑战（选做，不计入上述分数）', '5'),
      S('（1）把同一份提示词发给第三个工具（可以换成多模态AI），比较三家在同一个数字上的说法是否一致。如果不一致，你信谁？'),
      S('（2）把这次用到的提示词整理成一个模板文件，下次遇到同类问题直接套用。再写一句说明：这个模板里哪一句是必须改的？'),
      S('（3）本次用的都是表格数据。如果给你的只有6月费用凭证的照片，你的工具箱要增加哪一类工具？为什么？'),
      ],
     '插入【项目核心实操】（含统一五维度评分）'),
]


def stage_E():
    for anchor, paras, note in XST:
        ins_before(anchor, paras, note)


# ───────────────────────────────────────── 主流程
def main():
    global W, ZIP, TBL
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK + '/out', exist_ok=True)
    ZIP = zipfile.ZipFile(SRC)
    names = ZIP.namelist()
    ZIP.extractall(WORK + '/work')
    W = open(WORK + '/work/word/document.xml', encoding='utf-8').read()

    def recompute_tbl():
        global TBL
        TBL = [(m.start(), m.end()) for m in re.finditer(r'<w:tbl>.*?</w:tbl>', W, re.S)]

    recompute_tbl()
    stage_A(); recompute_tbl()
    stage_B(); recompute_tbl()
    stage_C(); recompute_tbl()
    stage_D(); recompute_tbl()
    log('【E】补栏目')
    stage_E()

    # 自检断言
    decl = re.findall(r'w:numId w:val="(\d+)"', W)
    num = ZIP.read('word/numbering.xml').decode('utf-8')
    dec_ids = set()
    for m in re.finditer(r'<w:num w:numId="(\d+)"[^>]*>(.*?)</w:num>', num, re.S):
        x = re.search(r'w:abstractNumId w:val="(\d+)"', m.group(2))
        aid = x.group(1) if x else None
        mm = re.search(r'<w:abstractNum w:abstractNumId="%s"[^>]*>.*?<w:lvl w:ilvl="0"[^>]*>(.*?)</w:lvl>' % aid, num, re.S)
        f = re.search(r'<w:numFmt w:val="([^"]+)"', mm.group(1)).group(1) if mm else '?'
        if f == 'decimal': dec_ids.add(m.group(1))
    left = [d for d in decl if d in dec_ids]
    assert not left, '仍有自动数字编号：%s' % set(left)
    log('  ✔ 断言通过：正文已无 decimal 自动编号（仅保留项目符号）')

    open(WORK + '/work/word/document.xml', 'w', encoding='utf-8').write(W)
    out = WORK + '/out/' + SRC.replace('.docx', '_修订稿V1.docx')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            zo.write(os.path.join(WORK + '/work', n), n)
    log('\n文档已生成：%s  (%.1f KB)   共 %d 项' % (out, os.path.getsize(out) / 1024, len(LOG)))


if __name__ == '__main__':
    main()
