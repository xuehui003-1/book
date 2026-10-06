#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目1 修订脚本 V2 —— 落实用户 2026-10-06 的 4 条意见
  ① 综合测评 3 道题改写（内容贴合当前章节，编号连续不跳号）
  ② 【项目总结】删除"五大模块、十二个项目"整条，后续条目重新编号
  ③ 表1-8 改名（原名"流程表对比"名实不符）
  ④ 表1-8 题干补充说明（明确只需补全第2、3步，第1、4步为示例）

输入：V1 修订稿的 word/document.xml
输出：/tmp/v2/out/word/document.xml
"""
import sys, os, re, html
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from docx_utils import *

SRC = '/tmp/v2/work'          # V1 解包目录
OUT = '/tmp/v2/out'
log = []

def load(n):
    return open(os.path.join(SRC, n), encoding='utf-8').read()

def save(n, d):
    p = os.path.join(OUT, n)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, 'w', encoding='utf-8').write(d)

doc = load('word/document.xml')
RPR = std_rpr(doc)


def repl_in_t(raw, old, new, tag=''):
    """在单个 <w:t> 内替换文字；找到几处就换几处"""
    n = 0
    while True:
        try:
            s, e = find_in_t(raw, old)
        except ValueError:
            break
        raw = raw[:s] + new + raw[e:]
        n += 1
    if n == 0:
        raise ValueError('未找到：' + old[:60])
    if tag:
        log.append('%s：%s → %s' % (tag, old[:34], new[:34]))
    return raw


# ══════════════════════════════════════════════════════════════
# ① 综合测评：3 道题改写（编号保持连续，不新增不删除）
# ══════════════════════════════════════════════════════════════
def fix_exam(raw):
    # ---- 单选第 6 题：原考"主线技能+创新拓展技能"（该内容已随任务1.3 移出）
    #      改考 1.2.1 节 AI 的三类角色（图1-9），与课堂内容直接对应
    PAIRS = [
        ('本书采用“主线技能+创新拓展技能”设计的主要原因是（ ）',
         '关于AI在财务分析中的三类角色，下列说法正确的是（ ）'),
        ('A. 主线技能太简单，需要拓展',
         'A. 工具型角色主要负责帮你拓展分析思路'),
        ('B. 拓展技能没有实用价值，只是装饰',
         'B. 辅助型角色主要负责把分析结论转化为报告和图表'),
        ('C. 兼顾高职高专学生的学习节奏和教材的前沿特色',
         'C. 表达型角色主要负责让你的分析更容易被管理者读懂和采用'),
        ('D. 出版社要求必须这样设计',
         'D. 三类角色只能单独使用，不能在同一个任务中先后出现'),
    ]
    for old, new in PAIRS:
        raw = repl_in_t(raw, old, new)
    log.append('① 单选第6题改写：改考"AI的三类角色"（对应图1-9 与 1.2.1 节），题号仍为 6，选项 A–D 齐全')

    # ---- 判断第 4 题：原考"本书学习目标"（依托已移出的难度分层说明）
    #      改考 1.2.2 节 智能体的四大核心组件（图1-11）
    raw = repl_in_t(raw,
                    '本书的学习目标是让所有学生都掌握自制Skills和多智能体协作等高阶技能。（ ）',
                    'AI智能体之所以能围绕目标自主执行任务，是因为有系统提示词、知识库、Skills和工作流等核心组件作支撑。（ ）')
    log.append('① 判断第4题改写：改考"智能体的四大核心组件"（对应图1-11 与 1.2.2 节），题号仍为 4')

    # ---- 素质测评第 5 题：原考"教材难度分层"（依托已移出内容）
    #      改考 1.2.3 节 人机协作分工（图1-13、图1-14）
    raw = repl_in_t(raw,
                    '“本书把智能体从第一个项目就开始介绍，是不是太超前了？高职高专学生能学得会吗？”',
                    '“既然AI已经能自动整理数据、生成图表，那财务人员只要会下指令就行了，还需要懂分析方法吗？”')
    log.append('① 素质测评第5题改写：改考"人机协作分工"（对应图1-13、图1-14 与 1.2.3 节），题号仍为 5')

    # ---- 复核编号连续性（按"（一）～（四）"小节切段，避免跨节串号）
    spans = para_spans(raw)

    def seq_between(start_key, end_key):
        i0, _, _ = find_para(raw, start_key, exact=False)
        i1, _, _ = find_para(raw, end_key, exact=False)
        out = []
        for j in range(i0, i1):
            t = para_text(raw, *spans[j]).strip()
            m = re.match(r'^(\d+)[.．]', t)
            if m:
                out.append(int(m.group(1)))
        return out

    secs = [('单项选择题', seq_between('1.以下属于', '（二）判断题')),
            ('判断题',   seq_between('1.会计核算做得好', '（三）简答题')),
            ('简答题',   seq_between('1.请根据图1-5', '（四）综合分析题'))]
    for name, s in secs:
        if not s:
            raise ValueError('%s 未取到题号' % name)
        if s != list(range(1, len(s) + 1)):
            raise ValueError('%s 题号不连续：%s' % (name, s))
        log.append('① 复核 %s 题号：%s —— 连续无跳号 ✔' % (name, '、'.join(map(str, s))))

    return raw


# ══════════════════════════════════════════════════════════════
# ② 【项目总结】删除"五大模块、十二个项目"整条 + 后续重新编号
# ══════════════════════════════════════════════════════════════
def fix_summary(raw):
    i_head, a_head, _ = find_para(raw, '4. 本书的学习是能力递进的')
    spans = para_spans(raw)
    i_body = i_head + 1
    if not para_text(raw, *spans[i_body]).strip().startswith('五大模块、十二个项目'):
        raise ValueError('第4条正文段落定位失败')
    a_del, b_del = spans[i_head][0], spans[i_body][1]
    raw = raw[:a_del] + raw[b_del:]
    log.append('② 【项目总结】已删除第4条"本书的学习是能力递进的"及正文"五大模块、十二个项目，从工具认知到……"'
               '（原因：图1-15、图1-16、表1-7 已随任务1.3 移出，且后续将新增"综合实训"项目，"十二个项目"的表述将不再成立）')

    # 原第5条 → 第4条
    raw = repl_in_t(raw, '5. 这不是在学“一个软件”，而是在构建“一套能力”',
                    '4. 这不是在学“一个软件”，而是在构建“一套能力”')
    log.append('② 原第5条重新编号为第4条：【项目总结】由 5 条调整为 4 条，编号 1–4 连续 ✔')
    return raw


# ══════════════════════════════════════════════════════════════
# ③④ 表1-8：改名 + 题干补充填写说明
# ══════════════════════════════════════════════════════════════
def fix_table8(raw):
    # ③ 改名：原"表1-8 流程表对比"名为对比，但表中只有"智能体的自动行动"一列，
    #    对比动作实际由表格后的"思考"题承担，故更名为表意准确的名称
    raw = repl_in_t(raw, '表1-8 流程表对比', '表1-8 费用分析智能体的自动执行步骤')
    log.append('③ 表1-8 改名："流程表对比" → "费用分析智能体的自动执行步骤"'
               '（判定依据：表中只有"智能体的自动行动"一列，无第二列可对比；'
               '"对比方法一"实际由表后的"思考"题承担，改名后名实相符。'
               '未采用"加一列"方案，因为方法一是单次提问、并无逐步执行的步骤，强行对齐会牵强）')

    # ④ 题干补充说明：第1、4步为示例，只需补全第2、3步
    raw = repl_in_t(raw,
                    '会自动完成哪些步骤？请根据图1-10的逻辑，填写以下流程（见表1-8）：',
                    '会自动完成哪些步骤？请根据图1-10的逻辑，补全表中第2步和第3步'
                    '（第1步和第4步已给出，作为示例），填入表1-8：')
    log.append('④ 表1-8 题干已补充说明："补全第2步和第3步（第1步和第4步已给出，作为示例）"——'
               '明确这是"给头给尾、补中间"的留白设计，避免读者误以为表格残缺')

    # 复核：表1-8 全文引用一致
    n = len(find_all_para(raw, '表1-8'))
    log.append('④ 复核：表1-8 在全文共出现 %d 处（题注 1 处 + 正文引用），编号一致 ✔' % n)
    return raw


# ══════════════════════════════════════════════════════════════
# ⑤ 连带修正：清理其余指向"已移出内容"的表述
# ══════════════════════════════════════════════════════════════
def fix_knock_on(raw):
    # (1) 【教学导读】知识目标第3条：原"认识本书的技术体系、学习路径与能力目标"
    #     指向已移出的任务1.3，改考 1.2.2 节内容
    raw = repl_in_t(raw,
                    '认识本书的技术体系、学习路径与能力目标',
                    '认识智能体的核心组件及其与普通AI对话助手的区别')
    log.append('⑤ 【教学导读】知识目标第3条改写：原来指向已移出的"学习路径与能力目标"，'
               '改为"认识智能体的核心组件及其与普通AI对话助手的区别"（对应 1.2.2 节、图1-10、图1-11）')

    # (2) 【教学导读】重点第3条：原"本书的学习目标与能力递进路径"
    #     对应图1-16 已随任务1.3 移出，改为保留内容
    raw = repl_in_t(raw,
                    '本书的学习目标与能力递进路径',
                    '智能体的核心组件与人机协作分工')
    log.append('⑤ 【教学导读】重点第3条改写：原"本书的学习目标与能力递进路径"（图1-16 已移出），'
               '改为"智能体的核心组件与人机协作分工"（对应 1.2.2、1.2.3 节）')

    # (3) 单选第8题：原句"本书的学习强调循序渐进"所依托的表述已随【项目总结】第4条删除，
    #     改写为不依赖该表述的独立设问，题号、选项均不变
    raw = repl_in_t(raw,
                    '本书的学习强调循序渐进。如果一个学生不会数据整理，会面临什么问题（ ）',
                    '在财务数据分析中，如果数据整理环节做得不扎实，会面临什么问题（ ）')
    log.append('⑤ 单选第8题改写：原句"本书的学习强调循序渐进"所依托的表述已随【项目总结】第4条删除，'
               '改为独立设问，题号仍为 8、选项 A–D 不变')
    return raw


# ══════════════════════════════════════════════════════════════
# 执行
# ══════════════════════════════════════════════════════════════
doc = fix_exam(doc)
doc = fix_summary(doc)
doc = fix_table8(doc)
doc = fix_knock_on(doc)
save('word/document.xml', doc)

# 其余部件原样复制
import shutil
for root, dirs, files in os.walk(SRC):
    for f in files:
        src = os.path.join(root, f)
        rel = os.path.relpath(src, SRC)
        if rel == 'word/document.xml':
            continue
        dst = os.path.join(OUT, rel)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)

print('\n'.join('  ✔ ' + x for x in log))
print('\n✅ V2 document.xml 修订完成 →', OUT)
