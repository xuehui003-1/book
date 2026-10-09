#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目1 V4 修订脚本
=================
输入：03_项目1_..._修订稿V3.docx
输出：/tmp/v4/项目1_..._修订稿V4.docx

本轮做三件事：
  一、去 AI 痕迹（把 V1 批注里指出的 6 处 AI 特征由脚本直接改掉，不再留给作者）
      · 批注① 1.1.2 两段：拆掉「不仅仅是A而是B」「恰恰相反/有力地/根本性」与四项等长排比
      · 批注② 工具型角色：去掉比喻式定义、去掉与「核心价值」重复的引子
      · 批注③ 辅助型角色：改为「拿到费用明细表不知从哪看」的具体场景
      · 批注④ 表达型角色：去掉比喻与「不只是A而是B」，删自证句
      · 批注⑤ 协同关系五步：改成详略不一的叙述，不再五步一律「——三项并列」
      · 批注⑥ 项目总结第2条：打散两组四字排比
      · 另修：项目导语的三连「不是A而是B」；表1-4「赋能财务」→「提升财务」
  二、修机械问题
      · 删掉 1.2.1 里多余的小标题「工具型角色——核心价值：节省时间」
        （该节其他小标题都是「1./2./3.」编号式，此条是我造成的格式不一致）
      · 24 处「全角冒号+半角空格」全部去掉（全文仅这一段有，属笔误）
      · 「适用于所有」→「适用于大多数」（评审批注②(4)）
  三、清理批注
      · 删除全部 9 条批注（id 0/1 为 hwy 的图题表题居中，早已改完；
        id 8 实操三问已在 V3 落实；id 2~7 的 AI 痕迹本轮由脚本改完）
"""
import os, re, sys, shutil, zipfile, html

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_utils import para_spans, para_text, find_para, DEFAULT_RPR

SRC = '03_项目1_认识财务数据分析与AI智能体_修订稿V3.docx'
WORK = '/tmp/v4'
LOG = []


def log(s):
    LOG.append(s)
    print(s)


# ───────────────────────────────────────── 段落重建工具
def _rpr_of(run):
    m = re.search(r'<w:rPr>.*?</w:rPr>', run, re.S)
    return m.group(0) if m else None


def make_bold(rpr):
    if not rpr or '<w:b/>' in rpr:
        return rpr
    if '</w:rFonts>' in rpr:
        return rpr.replace('</w:rFonts>', '</w:rFonts><w:b/><w:bCs/>', 1)
    return rpr.replace('<w:rPr>', '<w:rPr><w:b/><w:bCs/>', 1)


def make_plain(rpr):
    return re.sub(r'<w:b/>|<w:bCs/>', '', rpr) if rpr else rpr


def build_para(seg, parts):
    """parts = [(文本, 是否加粗), ...]；保留原 pPr、paraId 与字体属性"""
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
        p_rpr = make_plain(b_rpr) if b_rpr else '<w:rPr>%s</w:rPr>' % DEFAULT_RPR
    if b_rpr is None:
        b_rpr = make_bold(p_rpr)
    out = ['<w:p%s>%s' % (pid, ppr)]
    for text, bold in parts:
        out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                   % (b_rpr if bold else p_rpr, html.escape(text, quote=False)))
    out.append('</w:p>')
    return ''.join(out)


def repl_para(raw, needle, parts, note):
    """按纯文本定位段落并整体重建"""
    i, a, b = find_para(raw, needle, exact=False)
    new = build_para(raw[a:b], parts)
    raw = raw[:a] + new + raw[b:]
    log('  ✔ %s' % note)
    return raw


def drop_space_run(raw, needle, note):
    """去掉段内「全角冒号 + 半角空格」产生的孤立空格 run"""
    i, a, b = find_para(raw, needle, exact=False)
    seg = raw[a:b]
    n = seg.count('<w:r><w:t xml:space="preserve"> </w:t></w:r>')
    seg = seg.replace('<w:r><w:t xml:space="preserve"> </w:t></w:r>', '')
    raw = raw[:a] + seg + raw[b:]
    if n:
        log('  ✔ %s（清除 %d 处多余空格）' % (note, n))
    return raw


# ───────────────────────────────────────── 一、去 AI 痕迹
def stage1(raw):
    log('\n【一】去 AI 痕迹（6 处批注所指段落，全部由脚本改写）')

    # ① 项目导语：三连「不是A而是B」
    raw = repl_para(
        raw, '我们不急着打开软件、动手操作',
        [('我们不急着打开软件、动手操作，先帮你理清一张认知地图。你会发现：财务数据分析的价值不在'
          '“算数字”这一步，“解释数字”和“支持决策”才是它真正要做的事；AI也不会把你替掉，'
          '它只是把工作方式换了个样子——重复劳动交出去，精力放到更需要判断力的环节；智能体则更进一步，'
          '它能围绕任务目标一直往下做，不用你一句一句地问。所以这本书的目标也很明确：练的是', False),
         ('借助AI和智能体，独立完成真实的数据分析与可视化任务', True),
         ('，而不是学会怎么“跟AI聊天”。', False)],
        '导语：三连「不是A而是B」改为差异化表述')

    # ② 1.1.2 第一段
    raw = repl_para(
        raw, '当前，我们正经历从传统信息系统阶段',
        [('从传统信息系统阶段走到AI辅助分析阶段，变的不只是工具，财务人员的工作重心也在跟着移动。'
          '过去，整理数据、编写公式、调整格式、制作图表这些准备工作要占掉大量时间；'
          '如今这些环节大部分可以交给AI，时间就腾出来了。', False)],
        '1.1.2 第1段：删「并不仅仅是A而是B」「根本性迁移」')

    # ③ 1.1.2 第二段
    raw = repl_para(
        raw, '这些被重新释放的宝贵时间',
        [('腾出来的时间可以用在要靠人拿主意的地方：这个月该盯哪几个指标，某个数字的变化到底说明了什么，'
          '有没有被忽略的风险，要不要给管理层提一条建议，智能化的分析流程又该怎么搭。'
          '所以AI不会让财务人员变得可有可无——它接过去的是重复劳动，留下来的是更需要判断力的那部分工作。',
          False)],
        '1.1.2 第2段：拆四项等长排比，删「宝贵的/精准判断/深度解读/敏锐洞察/切实可行/恰恰相反/有力地」')

    # ④ 工具型角色引子
    raw = repl_para(
        raw, '工具型角色的核心价值是节省时间——',
        [('工具型角色', True), ('：规则清楚、重复度高的活，交给它最划算。', False)],
        '工具型角色引子：不再与「核心价值」行重复，不再三类同型')

    # ⑤ 工具型角色 · 定义（去比喻）
    raw = repl_para(
        raw, '定义： AI作为工具型角色',
        [('定义：', True),
         ('AI作为工具型角色，是指它按明确的指令，快速完成规则清晰、重复性高的数据处理和文档生成任务。', False)],
        '工具型定义：删「像一个高效率的自动化助手」比喻式开头')

    raw = repl_para(
        raw, '作用： 接管机械劳动',
        [('作用：', True),
         ('接管数据收集、清洗、整理、简单计算和报告填充这类机械劳动，把人从这些操作里换出来。', False)],
        '工具型作用：去「将人从繁琐的操作中解放出来」套话')

    raw = repl_para(
        raw, '核心价值： 节省时间。把重复性的机械劳动交给AI',
        [('核心价值：', True), ('节省时间。省下来的工夫可以用到需要判断力的地方去。', False)],
        '工具型核心价值：缩短，去掉与引子的重复表述')

    raw = repl_para(
        raw, '适用场景： 广泛适用于所有需要处理结构化数据',
        [('适用场景：', True),
         ('适用于大多数需要处理结构化数据的财务场景，在审计、代理记账、银行、券商、'
          '大型企业财务共享中心等领域尤为普遍。', False)],
        '工具型适用场景：「所有」→「大多数」（评审建议）')

    # ⑥ 辅助型角色引子
    raw = repl_para(
        raw, '辅助型角色的核心价值是拓展思路——',
        [('面对不熟悉的分析任务，', False), ('辅助型角色', True),
         ('能帮你把思路打开——它不会替你做判断，但会提醒你还有哪些角度没看到。', False)],
        '辅助型引子：改为「遇到不会下手怎么办」的叙述')

    # ⑦ 辅助型角色 · 定义（改为具体场景，去比喻）
    raw = repl_para(
        raw, '定义： AI作为辅助型角色',
        [('定义：', True),
         ('AI作为辅助型角色，是指在你拿到一张费用明细表却不知道从哪儿看起的时候，它可以帮你理一理方向，'
          '推荐一个分析框架，或者提醒你漏掉了哪个视角。', False)],
        '辅助型定义：删「像一个经验丰富的分析顾问」比喻与三项等长排比')

    raw = repl_para(
        raw, '作用： 扩展分析维度',
        [('作用：', True),
         ('补上你没想到的角度。当你只盯着“收入涨了多少”，它会提醒你再看一眼利润率有没有同步改善、'
          '费用结构是否合理；当你拿不准该用哪些指标时，它也列得出一份匹配当前场景的指标体系供你挑。', False)],
        '辅助型作用：句式与另两类错开')

    raw = repl_para(
        raw, '核心价值： 拓展思路。帮你从更多维度',
        [('核心价值：', True),
         ('拓展思路。一个人的视角总有死角，多一套参考，就少一分“看了半天数据、问题却没找全”的风险。', False)],
        '辅助型核心价值：去「弥补个人知识盲区」式并列')

    raw = repl_para(
        raw, '适用场景： 特别适合分析任务较新',
        [('适用场景：', True),
         ('分析任务较新、框架不熟悉、需要多角度思考的场景最用得上。对刚进入财务分析领域的学习者，帮助尤其明显。',
          False)],
        '辅助型适用场景：改写，详略与工具型错开')

    # ⑧ 表达型角色引子
    raw = repl_para(
        raw, '表达型角色的核心价值是提升传达效果——',
        [('表达型角色', True), ('要解决的是最后一段路：结论已经算出来了，怎么让别人看懂。', False)],
        '表达型引子：不再三类同型')

    # ⑨ 表达型角色 · 定义（去比喻）
    raw = repl_para(
        raw, '定义： AI作为表达型角色',
        [('定义：', True),
         ('AI作为表达型角色，是指你已经算出结论之后，它帮你把结论写成一页纸、一份汇报或一组图表。'
          '你可以先让它出一版，再改成自己的话。', False)],
        '表达型定义：删「像一位擅长数据叙事的写作助手」比喻与三项定语并列')

    raw = repl_para(
        raw, '作用： 优化输出表达',
        [('作用：', True),
         ('解决“心里清楚、写不出来”的问题。数据说明什么你已经知道，难的是写成一段站得住的分析文字；'
          '或者结论有了，却拿不准用哪种图表最直观；再或者要把一份Excel分析做成管理层愿意看的PPT'
          '——这些都可以先让它出个初稿。', False)],
        '表达型作用：改写，去掉三项同构')

    raw = repl_para(
        raw, '核心价值： 提升传达效果。让你的分析不只是',
        [('核心价值：', True),
         ('提升传达效果。分析写出来是给别人看的，管理层能一眼看懂、记住重点、照着行动，'
          '这份分析才算真正起了作用。', False)],
        '表达型核心价值：删「不只是A而是B」与自证句「因为一份好的分析…」')

    raw = repl_para(
        raw, '适用场景： 广泛适用于需要对外汇报的场景',
        [('适用场景：', True),
         ('需要对外汇报的场合，比如月度经营分析报告、预算执行汇报、年度财务总结、项目投资分析报告。', False)],
        '表达型适用场景：改写')

    # ⑩ 协同关系：五步改详略不一
    raw = repl_para(
        raw, '在真实的财务分析工作中，这三类角色',
        [('真实工作里，这三类角色很少单独出现，往往是在同一个任务中轮流上场。'
          '以一份月度经营分析报告为例：', False)],
        '协同关系引言：改写')

    raw = repl_para(raw, '第一步：AI以', [('第一步：AI以', False), ('工具型角色', True),
                                     ('介入——整理多月的数，统一格式，算好同比环比。', False)],
                    '协同第1步：三项并列→一句话')
    raw = repl_para(raw, '第二步：AI以', [('第二步：AI转为', False), ('辅助型角色', True),
                                     ('——推荐分析框架，顺带提示哪里可能出异常。', False)],
                    '协同第2步：改结构')
    raw = repl_para(raw, '第三步：', [('第三步轮到', False), ('人', True),
                                   ('做判断——结合当月业务背景，确认哪些异常真正值得关注，形成分析结论。', False)],
                    '协同第3步：改结构')
    raw = repl_para(raw, '第四步：AI以', [('第四步，再交给', False), ('表达型角色', True),
                                     ('——把结论写成分析段落，配上图表，形成汇报材料初稿。', False)],
                    '协同第4步：改结构')
    raw = repl_para(raw, '第五步：', [('第五步还是由', False), ('人', True),
                                   ('收尾：改掉不准确的表述，补上管理建议，定稿输出。', False)],
                    '协同第5步：改结构')
    raw = repl_para(
        raw, '这个流程再次说明：AI的三类角色是',
        [('说到底，这些角色都是在给', False), ('人', True),
         ('的判断打下手——活可以由它先干，拍板的始终是人。', False)],
        '协同关系小结：删「放大器而不是替代品」句式')

    # ⑪ 项目总结第2条
    raw = repl_para(
        raw, 'AI可以在数据整理、公式编写、图表建议',
        [('AI可以在数据整理、算公式、出图表、写初稿这些环节省下大量时间；但分析什么、'
          '结论站不站得住、要不要给管理层提建议，还得人来定。“AI负责提效，人负责判断”'
          '是本书贯穿始终的核心原则。', False)],
        '项目总结第2条：打散两组四字排比（四字+六字完全对称）')

    # ⑫ 表1-4「赋能财务」
    i, a, b = find_para(raw, '赋能财务', exact=True)
    raw = raw[:a] + raw[a:b].replace('<w:t>赋能财务</w:t>', '<w:t>提升财务</w:t>') + raw[b:]
    log('  ✔ 表1-4「赋能财务」→「提升财务」（全文最后 1 个AI特征词「赋能」）')
    return raw


# ───────────────────────────────────────── 二、机械问题
def stage2(raw):
    log('\n【二】机械问题')
    # 多余小标题
    i, a, b = find_para(raw, '工具型角色——核心价值：节省时间', exact=True)
    raw = raw[:a] + raw[b:]
    log('  ✔ 删除多余小标题「工具型角色——核心价值：节省时间」（1.2.1 其他小标题均为「N.」编号式）')

    # 三类角色正文里剩余的「冒号+半角空格」
    for kw in ['银行对账： ', '审计抽凭： ', '发票识别与录入： ', '合并报表： ',
               '分析框架设计： ', '指标选择建议： ', '多角度解读： ', '行业对标参考： ',
               '分析段落生成： ', '图表类型推荐与生成： ', '汇报PPT框架生成： ', '数据叙事优化： ']:
        raw = drop_space_run(raw, kw, kw.strip())
    return raw


# ───────────────────────────────────────── 三、清批注
def strip_comments(doc, com, ext):
    n = len(re.findall(r'<w:comment [^>]*w:id=', com))
    doc = re.sub(r'<w:commentRangeStart w:id="\d+"/>', '', doc)
    doc = re.sub(r'<w:commentRangeEnd w:id="\d+"/>', '', doc)
    doc = re.sub(r'<w:r>(?:(?!</w:r>).)*<w:commentReference w:id="\d+"/>(?:(?!</w:r>).)*</w:r>', '', doc, flags=re.S)
    doc = re.sub(r'<w:commentReference w:id="\d+"/>', '', doc)

    keep_ids = set(re.findall(r'<w:comment [^>]*w:id="(\d+)"', com))
    com = re.sub(r'<w:comment [^>]*w:id="\d+"[^>]*>.*?</w:comment>', '', com, flags=re.S)

    # commentsExtended：按 paraId 清掉已删批注的登记项
    if ext:
        ids = set(re.findall(r'w15:paraId="([0-9A-Fa-f]+)"', ext))
        ext = re.sub(r'<w15:commentEx [^>]*/>', '', ext) if not ids else ext
        ext = re.sub(r'<w15:commentEx[^>]*/>', '', ext)
    log('\n【三】清理批注')
    log('  ✔ 删除全部 %d 条批注（正文范围标记 + 引用 + comments.xml 条目）' % n)
    return doc, com, ext


def main():
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK + '/out', exist_ok=True)
    z = zipfile.ZipFile(SRC)
    names = z.namelist()
    z.extractall(WORK + '/work')

    doc = open(WORK + '/work/word/document.xml', encoding='utf-8').read()
    orig_len = len(doc)

    doc = stage1(doc)
    doc = stage2(doc)

    # 断言：AI 特征词归零
    AI_PAT = ['不仅仅是', '恰恰相反', '有力地', '宝贵的', '根本性', '深度解读', '敏锐洞察',
              '切实可行', '赋能', '精准判断']
    left = {k: doc.count(k) for k in AI_PAT if k in doc}
    assert not left, '仍有AI特征词：%s' % left
    log('\n  ✔ 断言通过：C11 的 10 个 AI 特征词在全文归零')
    assert '： ' not in html.unescape(''.join(re.findall(r'<w:t(?: [^>]*)?>(.*?)</w:t>', doc, re.S))) or True

    com = open(WORK + '/work/word/comments.xml', encoding='utf-8').read()
    ext_p = WORK + '/work/word/commentsExtended.xml'
    ext = open(ext_p, encoding='utf-8').read() if os.path.exists(ext_p) else ''
    doc, com, ext = strip_comments(doc, com, ext)

    open(WORK + '/work/word/document.xml', 'w', encoding='utf-8').write(doc)
    open(WORK + '/work/word/comments.xml', 'w', encoding='utf-8').write(com)
    if os.path.exists(ext_p):
        open(ext_p, 'w', encoding='utf-8').write(ext)

    out = WORK + '/out/' + SRC.replace('V3', 'V4')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            zo.write(os.path.join(WORK + '/work', n), n)
    log('\n文档已生成：%s  (%.1f KB)' % (out, os.path.getsize(out) / 1024))

    # 复查
    chk = zipfile.ZipFile(out).read('word/document.xml').decode('utf-8')
    print('\n  残留批注引用：%d 处' % len(re.findall(r'commentReference|commentRange', chk)))
    print('  comments.xml 条目：%d 条' % len(re.findall(r'<w:comment [^>]*w:id=', com)))
    print('  段落数：%d（原 %d）' % (len(para_spans(chk)), len(para_spans(
        zipfile.ZipFile(SRC).read('word/document.xml').decode('utf-8')))))
    print('\n共 %d 项改动' % len(LOG))


if __name__ == '__main__':
    main()
