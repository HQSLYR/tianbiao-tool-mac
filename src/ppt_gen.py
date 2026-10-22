# -*- coding: utf-8 -*-
"""PPT 演示文稿生成引擎：Word 文档或文字描述 → 可直接放映的 .pptx。

约定（描述文本的轻量标记，兼容 Word 标题样式自动转出来的格式）：
    #  开头一行  = 整套 PPT 的大标题（只取第一次出现的）
    ## 开头一行  = 一页幻灯片的标题
    - / • / * / 1. / 1、 开头 = 要点条目
    备注：xxx     = 给当前页写一句演讲者备注
    配图：xxx     = 指定当前页插画（按关键词从约 38 种矢量图里自动挑选绘制）
    其它非空行   = 普通句子，自动归入当前页要点

自然语言需求单（含“主题是/面向/N页/内容结构建议包括：A、B、C”）会被自动解析，
每个话题按内置内容库扩成一页“标题 + 6~7 条要点 + 小结 + 演讲备注 + 自动矢量配图”。

自动配图全部用 PPT 形状离线绘制（足球、球场、奖杯、战术板、数据柱图、齿轮、
灯泡、靶心、火箭、奖牌、邮件、盾牌…），不依赖网络与图片素材；并且每生成一次
都会按“款式轮变计数器”更换配图款式、配图左右位置与要点符号，让每次产物都不同。
"""
import os
import re
import time
import unicodedata

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

_MAX_BULLETS = 7      # 每页最多要点数，超出自动分页（标题加“·续”）
_MAX_BULLEN = 40      # 单条要点最长字数，长句只留前一半意思
_MIN_PAGES = 1        # 自定义页数下限
_MAX_PAGES = 30       # 自定义页数上限
_TITLE_PT = 26
_BODY_PT = 16         # 正文略增大，配合更多要点填满页面

# 三套配色：晴空蓝 / 星云紫 / 青岚绿（与⑧演示页面同一口径）
SCHEMES = [
    {'name': '晴空蓝', 'pri': '3370FF', 'bg': '1D3A8A', 'soft': 'EAF1FF', 'acc': '7BA0FF'},
    {'name': '星云紫', 'pri': '7C4DDB', 'bg': '3B2A78', 'soft': 'F2ECFF', 'acc': 'B49AF0'},
    {'name': '青岚绿', 'pri': '0FA47F', 'bg': '0B5D47', 'soft': 'E6F7F1', 'acc': '7FD9BC'},
]

_CN = '一二三四五六七八九十'
_GOLD = 'F2B705'
_GOLD_D = 'C88A04'
_GRASS = '2FA84F'
_RED = 'E5484D'
_INK = '1F2329'
_W = 'FFFFFF'
_GLYPHS = ('▪', '◆', '●', '▶', '■')       # 要点符号，随生成轮变


def _rgb(s):
    return RGBColor(int(s[:2], 16), int(s[2:4], 16), int(s[4:], 16))


def _zh(run, name='微软雅黑'):
    """中英文字体名一起设置，防止中文回退成宋体。"""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', name)


def clean_line(s):
    s = re.sub(r'\s+', ' ', unicodedata.normalize('NFKC', s or '')).strip()
    s = re.sub(r'^[-•*·]+', '', s).strip()
    s = re.sub(r'^\d+[\.、)]\s*', '', s).strip()
    s = re.sub(r'^[▲◆■●○◇△▼#]+', '', s).strip()
    return s


def clip(s):
    """长句不上屏：超长的要点截到 40 字以内，优先在标点处断。"""
    if len(s) <= _MAX_BULLEN:
        return s
    cut = s[:_MAX_BULLEN]
    for p in ('，', '。', '；', '、', ','):
        i = cut.rfind(p)
        if i >= 12:
            return cut[:i]
    return cut + '…'


# ---------------- 来源文本提取 ----------------

def extract_source(path):
    """Word/表格/文本 → 带 # ## - 标记的文本（可在界面里继续手工编辑）。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.doc', '.wps'):
        from word_to_table import to_docx
        path = to_docx(path)
        ext = '.docx'
    if ext == '.docx':
        from docx import Document
        doc = Document(path)
        lines = []
        for p in doc.paragraphs:
            t = p.text.strip()
            if not t:
                continue
            style = (p.style.name or '').lower()
            if 'heading 1' in style or style.startswith('标题 1'):
                lines.append('# ' + t)
            elif 'heading' in style or style.startswith('标题'):
                lines.append('## ' + t)
            elif 'list' in style or '列表' in style:
                lines.append('- ' + t)
            else:
                lines.append(t)
        for tb in doc.tables:
            for row in tb.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    lines.append('- ' + '、'.join(cells))
        return '\n'.join(lines)
    if ext == '.xlsx':
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        lines = []
        for ws in wb.worksheets:
            lines.append('## ' + ws.title)
            for row in ws.iter_rows(values_only=True):
                cells = [str(v).strip() for v in row if v is not None and str(v).strip()]
                if cells:
                    lines.append('- ' + '、'.join(cells))
        wb.close()
        return '\n'.join(lines)
    for enc in ('utf-8', 'gbk'):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(path, 'rb') as f:
        return f.read().decode('utf-8', 'ignore')


# ---------------- 自然语言需求单识别 ----------------

_ASK_WORDS = ('帮我', '请做', '请帮', '制作', '生成', '主题是', '面向',
              '演讲备注', '内容结构', '页左右', '要点')

# 多主题运动/项目百科：每个主题给出一组具体字段，分节模板按字段生成专属要点。
# 换不同主题（足球→篮球→…），内容与配图随之变成该项目本身。
_SUBJECTS = {
    '足球': dict(alias=r'足球|蹴鞠|踢球|绿茵',
        origin='中国古代的蹴鞠，现代足球1863年诞生于英国', govern='国际足联FIFA，世界杯四年一届',
        players='每队上场11人', duration='常规90分钟，分上下半场',
        court='长约105米、宽约68米的草坪球场', score='把球踢进对方球门，进球多者胜',
        fouls='越位、犯规与红黄牌', tactics='常见阵型有4-4-2、4-3-3、3-5-2',
        stars='贝利、马拉多纳、梅西、C罗', events='世界杯、欧冠、欧洲杯',
        health='一场跑动常超一万米，练心肺与协作', start='一双球鞋一块场地就能开始，先练颠传停',
        culture='看台口号、球队围巾与德比之夜', biz='门票、转播权、球衣赞助与球员转会',
        ball='ball', court_pic='field'),
    '篮球': dict(alias=r'篮球|扣篮|NBA|nba|三步上篮',
        origin='1891年由美国体育教师奈史密斯发明，最初把桃篮钉在看台', govern='国际篮联FIBA，最高职业联赛是NBA',
        players='每队上场5人', duration='正式比赛分4节，每节10到12分钟',
        court='长28米、宽15米硬地球场，篮筐高3.05米', score='把球投入篮筐，按区域得2分或3分',
        fouls='走步、二次运球、打手与24秒违例', tactics='挡拆、三角进攻、快攻与联防',
        stars='乔丹、科比、魔术师、詹姆斯、库里', events='NBA总决赛、男篮世界杯、奥运篮球',
        health='跑跳结合，利于长高增肌与反应', start='一个篮球一个篮筐即可，先练运球投篮',
        culture='街球、扣篮、球衣文化与主场欢呼', biz='球星高薪、球鞋代言、转播与周边',
        ball='ball_basket', court_pic='court_hard'),
    '排球': dict(alias=r'排球|扣杀|拦网|自由人',
        origin='1895年源于美国，隔网击球的团队项目', govern='国际排联FIVB，有世界杯与世界联赛',
        players='每队上场6人', duration='五局三胜，前四局25分、决胜局15分',
        court='长18米、宽9米，中间架高网', score='球落到对方场地或对方失误即得分',
        fouls='持球、连击、过网击球与触网', tactics='主攻、副攻、二传与自由人配合',
        stars='郎平、朱婷，以及巴西、意大利名将', events='世界杯、世锦赛、奥运会排球',
        health='起跳扣杀练爆发力与协调', start='一面墙或一张网就能垫球入门',
        culture='女排精神与集体荣誉', biz='联赛赞助、转播与国家队商业',
        ball='ball_volley', court_pic='court_hard'),
    '乒乓球': dict(alias=r'乒乓球|乒乓|国球|弧圈',
        origin='19世纪末源于英国餐桌，后风靡中国', govern='国际乒联ITTF，大赛长期由中国队主导',
        players='单打两人、双打四人', duration='一般五局三胜，每局11分',
        court='标准球台约2.74米长、中间球网', score='用球拍击球过台，对方回接失误得分',
        fouls='发球违例、二次击球与触网', tactics='弧圈、快攻、削球与落点变化',
        stars='马龙、张怡宁、瓦尔德内尔', events='世乒赛、世界杯、奥运会乒乓',
        health='练眼手协调与反应，护眼又解压', start='一副球拍一张球台即可在家练',
        culture='国球荣耀与全民乒乓氛围', biz='赛事、俱乐部与器材赞助',
        ball='ball_pong', court_pic='table'),
    '网球': dict(alias=r'网球|大满贯|费德勒|纳达尔|德约|温网',
        origin='源于12世纪法国修道院游戏，19世纪定型于英国', govern='国际网联ITF，四大满贯最负盛名',
        players='单打1对1，双打2对2', duration='多为三盘两胜或五盘三胜',
        court='长23.77米球场，中间球网分隔', score='球过网落在对方场内，回不到者失分',
        fouls='双误、触网与脚误', tactics='底线对拉、发球上网与截击',
        stars='费德勒、纳达尔、德约科维奇、李娜', events='四大满贯、ATP/WTA、戴维斯杯',
        health='全身有氧，练耐力与协调', start='一支球拍一片场地即可开始',
        culture='草地、白色球衣与优雅礼仪', biz='大满贯转播、代言与青训',
        ball='ball_tennis', court_pic='field'),
    '羽毛球': dict(alias=r'羽毛球|林丹|李宗伟|扣杀羽毛球',
        origin='19世纪源于印度普那，后在中国与东南亚盛行', govern='世界羽联BWF，汤尤杯与全英赛著名',
        players='单打2人、双打4人隔网', duration='三局两胜，每局21分',
        court='长13.4米、中间架高网', score='球落在对方场内或对方失误得分',
        fouls='过网击球、发球过腰与触网', tactics='高远球、吊球、杀球与网前',
        stars='林丹、李宗伟、谌龙', events='奥运会、全英公开赛、汤姆斯杯',
        health='练反应、护颈椎，老少皆宜', start='两副球拍一片空地就能打',
        culture='全民运动与亚洲统治力', biz='赛事、器材与俱乐部',
        ball='shuttle', court_pic='court_hard'),
    '游泳': dict(alias=r'游泳|泳池|蛙泳|自由泳|蝶泳|仰泳',
        origin='远古求生技能，如今是热门竞技与健身项目', govern='世界泳联World Aquatics',
        players='个人项目与4人接力', duration='从50米冲刺到1500米长距离',
        court='标准50米泳池，8条泳道', score='竞速项目最先触壁者获胜',
        fouls='抢跳出发、水下越线等', tactics='自由泳、蛙泳、仰泳、蝶泳四种泳姿',
        stars='菲尔普斯、孙杨、莱德基', events='奥运会、世锦赛',
        health='全身低冲击有氧，护心肺又护关节', start='先学漂浮与换气，务必有人看护',
        culture='水上运动与救生文化', biz='赛事、泳池与品牌泳装',
        ball='pool', court_pic='pool'),
    '跑步': dict(alias=r'跑步|马拉松|路跑|长跑|短跑|慢跑',
        origin='人类最原始的运动，现代路跑风靡全球', govern='世界田联World Athletics',
        players='个人参与，一场可上万人', duration='从5公里到42.195公里全程马拉松',
        court='公路或400米标准跑道', score='竞速项目用时最短者获胜',
        fouls='抢跑、抄近道视为违规', tactics='配速、步频与呼吸节奏',
        stars='基普乔格、博尔特', events='奥运会、世界马拉松大满贯',
        health='增强心肺与耐力，燃脂高效', start='一双跑鞋就能开跑，从走跑结合开始',
        culture='城市马拉松与跑团文化', biz='赛事报名、跑鞋与运动装备',
        ball='boots', court_pic='track'),
    '健身': dict(alias=r'健身|增肌|减脂|撸铁|哑铃|力量训练|体能',
        origin='源于古希腊与近代举重训练', govern='各类健美与力量赛事组织',
        players='以个人训练为主', duration='按训练计划周期性推进',
        court='健身房或一块训练垫的空间', score='以增肌、减脂、体能达标为目标',
        fouls='动作变形容易受伤，需规范', tactics='力量、有氧与柔韧相结合',
        stars='健美与体能领域众多教练与运动员', events='世界健美、Cross Games 等',
        health='增肌减脂、改善体态与代谢', start='徒手或一对哑铃即可入门',
        culture='自律、打卡与身材管理', biz='健身房、器械与营养补剂',
        ball='dumbbell', court_pic='gear'),
}

# 分节模板：命中章节概念时用主题字段填出专属要点（{字段}由主题字典填充）。
# pic 用 COURT / BALL 哨兵，绘制时替换成该主题的球场/用球。
_SUBJ_SECTIONS = [
    (r'起源|发展|历史|由来|背景',
     ['源头：{origin}', '组织：{govern}', '传播：走向世界、走进日常生活',
      '里程碑：从地方玩法到{events}', '今天：广受欢迎的大众项目',
      '看点：规则简单、上手快、观赏性强', '记忆点：一句话记住它从哪来'],
     '一条时间线看懂它的起源与发展', 'timeline',
     '开场提问“你们猜它最早出现在什么时候？”，沿时间线讲起源故事。'),
    (r'规则|场地|玩法|判罚|装备|人数|赛制',
     ['场地：{court}', '人数：{players}', '时长：{duration}', '得分：{score}',
      '判罚：{fouls}', '战术入门：{tactics}', '装备：一套基础装备即可上场'],
     '记住 场地·人数·时长·得分 这四件事', 'COURT',
     '别念条文，挑一条最有意思的规则现场比划一遍最好懂。'),
    (r'阵型|战术|打法|技巧|方法|策略|泳姿',
     ['打法：{tactics}', '进攻：如何创造机会', '防守：站位与相互配合',
      '转换：攻防切换最快', '配合：默契来自反复磨合', '心理：领先不松、落后不乱',
      '看门道：先看阵型、再看配合'],
     '阵型决定站位，配合决定机会', 'board',
     '用战术板现场画一次站位，比口头说十句都管用。'),
    (r'球星|人物|明星|名将|代表|经典',
     ['代表人物：{stars}', '封神一战：被反复回放的经典', '职业态度：天赋背后的刻苦',
      '精神传承：一代代人的接力', '背后团队：教练、队友与保障',
      '跨界影响：从赛场到大众文化', '给我们启发：热爱加上坚持'],
     '伟大不只在于成绩，更在于坚持', 'trophy',
     '挑一位大部分人都认识的讲，其余快速带过。'),
    (r'入门|如何|怎么|学习|训练|新手|开始|计划',
     ['起步：{start}', '装备不必贵：先把基础练牢', '基本功：反复练常见动作',
      '节奏：每周两次、每次一小时', '找伙伴：一起练更容易坚持',
      '避坑：先热身，别急着上强度', '安全：循序渐进、注意防护'],
     '零基础也能开始，重在坚持练', 'BALL',
     '强调“零基础也能开始”，结尾给一个本周就能行动的小任务。'),
    (r'文化|球迷|精神|看台|氛围',
     ['文化：{culture}', '宿敌与德比：为什么这么燃', '赛场礼仪：输赢都要体面',
      '全民参与：人人都能玩', '仪式与口号：归属感的来源',
      '精神内核：团队、拼搏与尊重', '传承：把热爱交给下一代'],
     '体育不只比赛，更是一种生活', 'crowd',
     '描述一段看台或现场的热闹，把气氛点燃。'),
    (r'世界杯|大赛|赛事|联赛|杯赛|冠军',
     ['顶级赛事：{events}', '赛制：从小组赛到决赛', '经典时刻：捧起奖杯的瞬间',
      '观赛：跨越时区的狂欢', '数据看点：进球、纪录与排名',
      '前瞻：下一届看点', '参与：约上朋友一起看球'],
     '大赛是这项运动的最高舞台', 'trophy',
     '讲一届你看得最多的大赛，用“如果你也在现场”收尾。'),
    (r'产业|商业|经济|市场|价值|生意',
     ['产业链：{biz}', '俱乐部收入：门票、赞助与转播', '明星经济：身价与代言',
      '周边：装备、游戏与文旅', '媒体：转播权与流量',
      '普通人：参与也是一门消费', '趋势：数字化与全球化并进'],
     '这项运动背后是一门大生意', 'coin',
     '用一件球衣/一双球鞋的价格切入商业话题。'),
    (r'健康|身心|好处|锻炼|收益',
     ['身体：{health}', '心理：专注、解压与自信', '社交：团队运动更易坚持',
      '习惯：规律训练胜过突击', '提醒：充分热身、量力而行',
      '损伤预防：动作规范与休息', '长期：养成运动好习惯'],
     '练的是身体，养的是心态', 'heart',
     '给一个具体数字或例子最有说服力。'),
    (r'总结|回顾|提问|互动|展望|致谢',
     ['三句话回顾今天的重点', '一个行动：这周去试一试', '互动提问：你最想尝试哪项',
      '延伸：从观看到亲自参与', '资源：入门装备与学习渠道',
      '结伴：拉上朋友一起开始', '谢谢大家，欢迎交流'],
     '把今天的兴趣变成明天的行动', 'star',
     '提问环节先安排一个“托”，避免冷场。'),
]
_FALLBACK = (['{t}：是什么', '为什么值得关注', '它和我们的生活有什么关系',
              '一个关键的例子', '一句话记住它'],
             '围绕 {t} 的三个关键词', 'sparkle', '围绕“{t}”讲一个自己的小故事。')

# 通用内容库：任意主题（老师/笔记/学习/编程/生活……）都按“章节骨架”生成，
# 每条要点都织入真实话题名 {t}，不再是空模板；配图按话题关键词自动选。
# 每项 = (正则, [要点×5], 小结, 默认配图key, 备注)
_CONCEPTS = [
    (r'教学|老师|教师|授课|课堂|讲课|师生|备课|教案|板书',
     ['讲好 {t} 的导入：先抓住注意', '{t} 的知识点拆解', '互动设计：让 {t} 活起来',
      '课件与板书怎么呈现 {t}', '{t} 的难点与易错点',
      '{t} 一课的回顾与作业', '给 {t} 留一个动手小任务'],
     '把 {t} 讲清楚、讲生动', 'graduation',
     '讲 {t} 时多提问、多举学生身边的例子，别只念概念。'),
    (r'笔记|记录|整理|归纳|摘要|复盘',
     ['为什么给 {t} 做笔记', '{t} 笔记的常见方法', '如何提炼 {t} 的重点',
      '{t} 笔记的整理与归档', '{t} 笔记的复习与回顾',
      '用 {t} 笔记检验掌握程度', '让 {t} 笔记更高效的技巧'],
     '把 {t} 记下来、用起来', 'doc',
     '讲 {t} 笔记时现场演示一页，比空说方法管用。'),
    (r'学习|备考|考试|复习|记忆|背诵|掌握|提升',
     ['{t} 的目标先定清楚', '{t} 的具体方法与步骤', '执行 {t} 的关键习惯',
      '{t} 常见误区', '如何检验 {t} 学没学会',
      '坚持 {t} 的小技巧', '把 {t} 拆成每天可完成的一小步'],
     '把 {t} 变成每天能做的事', 'book',
     '给 {t} 配一个可量化的计划，鼓励听众今天就行动。'),
    (r'起源|历史|发展|由来|背景|什么|定义|简介|概述|概念',
     ['{t} 是什么：一句话讲清它', '{t} 从哪来：由来与背景', '为什么重要：{t} 的价值',
      '{t} 和我们的关系', '{t} 发展的关键节点',
      '{t} 最核心的一个特征', '用一个例子记住 {t}'],
     '先讲清楚 {t} 是什么、为什么值得了解', 'timeline',
     '开场用一句大白话定义“{t}”，再讲它的来龙去脉，别一上来堆术语。'),
    (r'原理|机制|算法|逻辑|神经|网络|模型|架构|组成|构成|如何工作|怎么工作',
     ['{t} 的核心原理：它怎么运转', '{t} 最关键的一环', '用一个类比理解 {t}',
      '{t} 的输入与输出', '关于 {t} 的常见误解',
      '{t} 的能力边界', '一句话讲清 {t} 的原理'],
     '用类比把 {t} 的原理讲透', 'gear',
     '讲 {t} 的原理时准备一个生活化的类比，比直接说定义好懂十倍。'),
    (r'方法|步骤|流程|技术|系统|基础|分类|做法',
     ['{t} 的基本流程：先做什么后做什么', '{t} 的关键注意点', '做 {t} 的几种常见做法对比',
      '{t} 最容易踩的坑', '做好 {t} 的工具与资源',
      '上手 {t} 的小技巧', '把 {t} 变成可复制的步骤'],
     '抓住 {t} 的几个关键环节', 'board',
     '把 {t} 拆成几步讲，配一张流程/结构示意图最直观。'),
    (r'应用|场景|用途|案例|例子|实践',
     ['{t} 的典型应用场景一', '{t} 的典型应用场景二', '一个关于 {t} 的真实案例',
      '{t} 带来的实际改变', '{t} 落地要注意什么',
      '{t} 适合与不适合的边界', '从 {t} 的一个应用说起'],
     '用例子说明 {t} 能做什么', 'monitor',
     '讲 {t} 的应用时上真实案例与数字，比空谈功能更有说服力。'),
    (r'优势|好处|优点|为什么选|价值|意义',
     ['{t} 的主要优势', '{t} 带来的具体好处', '和其他选择相比', '{t} 适合哪些人',
      '{t} 的局限与前提', '一个能量化的 {t} 效果',
      '一句话总结为什么选 {t}'],
     '把 {t} 的优势讲得能落地', 'check',
     '讲 {t} 优势时用一个对比或数字，别空喊口号。'),
    (r'数据|算力|资源|成本|投入|效率|性能',
     ['做 {t} 需要哪些资源', '{t} 的成本与投入', '如何衡量 {t} 的效果',
      '{t} 的优化方向', '{t} 的关键指标',
      '{t} 的风险与控制', '用数据说明 {t} 的取舍'],
     '量化地看 {t} 的投入产出', 'chart_bar',
     '给一组对比数据或图表，让 {t} 的效果一目了然。'),
    (r'人物|明星|代表|团队|名师|先驱|榜样',
     ['{t} 领域的代表人物', '他们的主要贡献', '一个细节看出专业度',
      '他们的成长路径', '{t} 先行者带来的启发',
      '团队与协作的作用', '值得学习的地方'],
     '通过代表人物理解 {t}', 'crowd',
     '挑一位有故事的代表人物展开，其余快速带过。'),
    (r'产业|商业|市场|经济|前景|趋势|生意|发展机会',
     ['{t} 的现状与规模', '{t} 的商业模式', '{t} 价值链上的机会', '{t} 的竞争格局',
      '{t} 的未来趋势', '普通人参与 {t} 的切入点',
      '一句话看懂 {t} 的生意'],
     '{t} 背后的商业逻辑', 'coin',
     '用一个大家熟悉的产品/价格切入 {t} 的商业话题。'),
    (r'风险|伦理|安全|隐私|问题|挑战|局限|注意',
     ['{t} 的常见问题', '{t} 背后的原因', '应对与改进办法', '使用 {t} 的注意事项',
      '{t} 的边界与底线', '面对 {t} 我们应有的态度',
      '如何安全地享受 {t}'],
     '客观看待 {t} 的风险与边界', 'shield',
     '讲 {t} 的风险要平衡：既说清问题，也给出可行应对。'),
    (r'健康|身心|锻炼|健身|保养|作息',
     ['{t} 在身体层面的好处', '{t} 在心理层面的好处', '{t} 的社交与长期收益',
      '做 {t} 的注意事项', '如何科学开始 {t}',
      '{t} 常见错误姿势或习惯', '把 {t} 变成可持续的日常'],
     '把 {t} 的好处讲具体', 'heart',
     '给一个具体数字或例子最有说服力。'),
    (r'文化|精神|理念|影响|传承',
     ['{t} 代表的文化/精神', '{t} 对人们的影响', '一个关于 {t} 的动人细节',
      '{t} 背后的故事', '我们为何对 {t} 产生共鸣',
      '{t} 值得传承的部分', '一句话讲出 {t} 的内核'],
     '讲出 {t} 背后的精神', 'flag',
     '用一段有画面感的描述，把 {t} 的精神讲进听众心里。'),
    (r'入门|新手|开始|怎么学|如何学|从零|零基础',
     ['入门 {t} 的第一步', '学 {t} 需要准备什么', '{t} 的基础练习方法',
      '学 {t} 的常见路径', '学 {t} 的节奏建议',
      '新手学 {t} 最容易踩的坑', '一周内就能上手的 {t} 小目标'],
     '零基础也能开始学 {t}', 'graduation',
     '强调“零基础也能开始”，结尾给一个本周就能完成的小任务。'),
    (r'总结|回顾|展望|未来|互动|提问|致谢|结束',
     ['三句话回顾 {t} 的重点', '关于 {t} 的一个行动建议', '互动提问：你对 {t} 哪点感兴趣',
      '{t} 的延伸与展望', '{t} 的下一步怎么走',
      '推荐给想深入了解 {t} 的人', '谢谢大家，欢迎交流'],
     '把 {t} 变成下一步行动', 'star',
     '提问环节先安排一个“托”，避免冷场。'),
]


# 插画库：约 38 种，name → 中文名（右下角配图标题用）
_PICS = {
    'ball': '足球', 'field': '球场', 'board': '战术板', 'trophy': '奖杯',
    'heart': '健康', 'coin': '商业', 'crowd': '球迷', 'timeline': '时间轴',
    'star': '要点', 'book': '知识', 'sparkle': '灵感', 'whistle': '裁判',
    'boots': '装备', 'goal': '进球', 'medal': '荣誉', 'chart_bar': '数据',
    'chart_pie': '占比', 'gear': '技术', 'bulb': '创意', 'target': '目标',
    'flag': '精神', 'clock': '进程', 'calendar': '日程', 'envelope': '联系',
    'phone': '沟通', 'monitor': '在线', 'search': '探索', 'shield': '保障',
    'check': '优势', 'question': '答疑', 'doc': '文档', 'graduation': '教学',
    'map': '布局', 'rocket': '加速', 'sun': '活力', 'music': '氛围',
    'lock': '安全', 'cloud': '云端', 'ball_basket': '篮球', 'ball_volley': '排球',
    'ball_pong': '乒乓球', 'ball_tennis': '网球', 'shuttle': '羽毛球',
    'dumbbell': '哑铃', 'pool': '泳池', 'track': '跑道', 'court_hard': '篮球场',
    'table': '球台',
}
# 关键词 → 插画 name（顺序敏感：具体在前、通用在后），适配任意主题
_PIC_PAT = [
    ('timeline', r'起源|历史|发展|由来|演变'),
    ('clock', r'时间|进程|周期|时长|进度'),
    ('calendar', r'日程|排期|规划|计划|步骤'),
    ('goal', r'球门|进球|得分|射门|进攻'),
    ('boots', r'球鞋|装备|护具|穿戴'),
    ('whistle', r'裁判|哨|执法'),
    ('field', r'场地|球场|草坪|场馆'),
    ('board', r'阵型|战术|打法|走位'),
    ('target', r'目标|策略|定位|精准|靶'),
    ('gear', r'技术|系统|设置|工程|机械|配置|流程'),
    ('chart_bar', r'数据|统计|分析|指标|增长|趋势|销量'),
    ('chart_pie', r'占比|构成|分布|份额|比例'),
    ('coin', r'产业|商业|经济|市场|价值|赚钱|收入|价格|成本|费用'),
    ('rocket', r'创新|启动|加速|前景|突破|腾飞|上线'),
    ('bulb', r'创意|想法|灵感|启发|思路|方案|设计|头脑'),
    ('search', r'搜索|探索|研究|发现|调研'),
    ('monitor', r'电脑|在线|直播|网络|平台|网页|远程'),
    ('envelope', r'邮件|联系|消息|通知|反馈'),
    ('phone', r'电话|沟通|通话|客服|移动'),
    ('music', r'音乐|氛围|娱乐|节奏'),
    ('cloud', r'云|存储|环境|天气|服务'),
    ('lock', r'安全|隐私|密码|权限|加密'),
    ('shield', r'保障|防护|防御|风控'),
    ('check', r'优势|完成|正确|达标|验证'),
    ('question', r'问题|疑问|解答|答疑'),
    ('doc', r'文档|报告|资料|说明|手册|笔记|记录|纪要|课件|作业|摘要|清单'),
    ('graduation', r'老师|教师|授课|课堂|讲课|师生|备课|教案|教育|培训|课程|学习|教学|校园'),
    ('book', r'知识|阅读|读书|理论|书籍|复习|备考|考试|记忆|课本|背书'),
    ('map', r'地图|位置|布局|版图|导览|区域'),
    ('crowd', r'文化|球迷|人群|团队|用户|观众|社交'),
    ('heart', r'健康|身心|好处|锻炼|关爱|体能'),
    ('medal', r'荣誉|奖项|成就|奖牌|第一'),
    ('trophy', r'世界杯|欧冠|大赛|赛事|奖杯|经典|比赛|球星|名将'),
    ('sun', r'青春|活力|阳光|朝气|明亮'),
    ('flag', r'精神|里程碑|领先|旗帜'),
    ('star', r'总结|回顾|互动|展望|要点|亮点'),
    ('ball_basket', r'篮球|扣篮|NBA|nba|上篮'),
    ('ball_volley', r'排球|拦网|扣杀'),
    ('ball_pong', r'乒乓球|乒乓|国球'),
    ('ball_tennis', r'网球|大满贯|费德勒|纳达尔'),
    ('shuttle', r'羽毛球|林丹|羽毛球'),
    ('pool', r'游泳|泳池|蛙泳|自由泳|蝶泳'),
    ('track', r'跑步|马拉松|路跑|长跑|短跑|跑道'),
    ('dumbbell', r'健身|哑铃|撸铁|增肌|力量训练'),
    ('ball', r'足球|蹴鞠|踢球|绿茵|运动|体育'),
]
# 同一命中可轮换的近似插画（每次生成换款式，让产物更丰富多样）
_VARIANTS = {
    'timeline': ['timeline', 'book', 'clock', 'calendar'],
    'field': ['field', 'goal', 'map', 'board'],
    'board': ['board', 'target', 'gear', 'field'],
    'trophy': ['trophy', 'medal', 'flag', 'star'],
    'heart': ['heart', 'sun', 'crowd'],
    'coin': ['coin', 'chart_bar', 'chart_pie'],
    'crowd': ['crowd', 'flag', 'music'],
    'ball': ['ball', 'boots', 'whistle', 'goal'],
    'book': ['book', 'graduation', 'doc'],
    'star': ['star', 'flag', 'target'],
    'gear': ['gear', 'monitor', 'cloud'],
    'chart_bar': ['chart_bar', 'chart_pie', 'coin'],
    'sparkle': ['bulb', 'target', 'chart_bar', 'gear', 'doc', 'flag', 'rocket',
                'star', 'search', 'shield', 'chart_pie', 'graduation', 'map',
                'calendar', 'monitor', 'envelope', 'question', 'check', 'book',
                'coin', 'crowd', 'sun', 'music', 'cloud'],
}


def _looks_like_prompt(text):
    """没有结构标记、但命中两个以上“需求单”用词 → 当自然语言需求处理。"""
    t = (text or '').strip()
    if not t or t.lstrip().startswith(('#', '##')):
        return False
    return sum(w in t for w in _ASK_WORDS) >= 2


def _pick_pic(spec):
    """配图描述 → 基础插画 name；已是 name 直接用，否则按关键词匹配。"""
    if spec in _PICS:
        return spec
    for key, pat in _PIC_PAT:
        if spec and re.search(pat, spec):
            return key
    return 'sparkle'


def _icon_for(spec, seed):
    """配图描述 + 轮变种子 → 具体插画 name（同类近似图按种子轮换）。"""
    base = _pick_pic(spec)
    cands = _VARIANTS.get(base, [base])
    return cands[seed % len(cands)]


def _detect_subject(ctx):
    """从大标题+话题里识别运动/项目主题，命中返回主题名，否则 None。"""
    for name, s in _SUBJECTS.items():
        if re.search(s['alias'], ctx or ''):
            return name
    return None


def _bank_for(topic, subj=None, weave=None):
    """话题 → (要点, 小结, 配图, 备注)。命中运动主题时按该项目字段生成专属内容，
    球场/用球配图也换成该项目专属；否则走通用概念库，再不中则退回兜底。

    weave 是整套 PPT 的真实主题名（大标题），通用要点里的 {t} 一律替换成它，
    避免出现“为什么给 为什么记笔记 做笔记”这种用碎页标签自我重复的怪句；
    章节模板先按本页话题命中，命中不了再按大标题命中，减少落入兜底空模板。"""
    if subj:
        s = _SUBJECTS[subj]
        for pat, bs, tk, pic, note in _SUBJ_SECTIONS:
            if re.search(pat, topic):
                bullets = [b.format(**s) for b in bs]
                if pic == 'COURT':
                    pic = s['court_pic']
                elif pic == 'BALL':
                    pic = s['ball']
                return bullets, tk, pic, note

    def _render(bs, tk, pic, note):
        sub = lambda x: x.replace('{t}', weave)            # noqa: E731
        hit = _pick_pic(topic)
        if hit == 'sparkle':
            hit = _pick_pic(weave)                          # 碎标签选不出→按大标题选贴题配图
        return [sub(b) for b in bs], sub(tk), (hit if hit != 'sparkle' else pic), sub(note)

    for pat, bs, tk, pic, note in _CONCEPTS:
        if re.search(pat, topic):
            return _render(bs, tk, pic, note)
    for pat, bs, tk, pic, note in _CONCEPTS:
        if weave and re.search(pat, weave):
            return _render(bs, tk, pic, note)
    bs, tk, pic, note = _FALLBACK
    return _render(bs, tk, pic, note)


def _extract_prompt(text):
    """需求单 → (大标题, 受众, [话题...], 配色序号或None, 页数或None)。"""
    one = re.sub(r'\s+', ' ', text)
    deck = audience = None
    m = re.search(r'主题是[（(【]?(?:例如[:：])?([^）)】，。；\n]+)', one)
    if m:
        deck = m.group(1).strip()
    else:
        m = re.search(r'关于(.{1,24}?)的\s*(?:PPT|ppt|演示文稿|幻灯片)', one)
        if m:
            deck = m.group(1).strip()
    m = re.search(r'面向[（(【]?(?:例如[:：])?([^）)】，。；/、\n]+)', one)
    if m:
        audience = m.group(1).strip()
    topics = []
    m = re.search(r'[^，。]*包括[：:]([^。]+)', one)
    if m:
        for part in re.split(r'[、，,;；/]+', m.group(1)):
            p = re.sub(r'^\d+[\.、]?\s*', '', part.strip())
            if 2 <= len(p) <= 18 and '例如' not in p:
                topics.append(p)
    ps = None
    if re.search(r'绿[色茵]', one):
        ps = 2
    elif '紫' in one:
        ps = 1
    elif '蓝' in one:
        ps = 0
    target = None
    m = re.search(r'(\d+)\s*页', one)
    if m:
        try:
            target = max(_MIN_PAGES, min(_MAX_PAGES, int(m.group(1))))
        except ValueError:
            target = None
    return deck, audience, topics, ps, target


def parse_outline(text):
    """文本 → (大标题, [(页标题, 要点, 备注, 配图, 小结)], 受众, 需求配色, 页数)。

    支持三种输入：带 # ## - 标记的文本；自然语言需求单（自动抽主题/受众/
    配色/内容结构并扩成整片）；无任何标记的流水行（自动成章分页）。
    """
    lines = [l.strip() for l in (text or '').splitlines() if l.strip()]
    if not lines:
        raise ValueError('请先提供 Word 文档或填写描述内容。')
    if _looks_like_prompt(text):
        deck, audience, topics, ps, target = _extract_prompt(text)
        if topics:
            subj = _detect_subject((deck or '') + ''.join(topics))
            weave = deck or ''                            # 整套真实主题，供通用要点织入
            slides = []
            for t in topics:
                bs, tk, pic, note = _bank_for(t, subj, weave or t)
                slides.append((clip(t), bs, note, pic, tk))
            return (deck or '演示文稿'), slides, audience, ps, target
    deck = None
    slides = []          # 每项 [标题, 要点list, 备注, 配图, 小结]
    cur = None
    for l in lines:
        if l.startswith('##'):
            cur = [clip(l.lstrip('#').strip() or '内容'), [], '', '', '']
            slides.append(cur)
        elif l.startswith('#'):
            deck = deck or clip(l.lstrip('#').strip())
        elif l.startswith(('备注：', '备注:')):
            if cur is not None:
                cur[2] = re.sub(r'^备注[:：]\s*', '', l)
        elif l.startswith(('配图：', '配图:', '图：', '图:')):
            if cur is not None:
                cur[3] = re.sub(r'^(配图|图)[:：]\s*', '', l)
        elif re.match(r'^[-•*]\s*\S', l) or re.match(r'^\d+[\.、)]\s*\S', l):
            if cur is None:
                cur = ['内容', [], '', '', '']
                slides.append(cur)
            cur[1].append(clip(clean_line(l)))
        else:
            s = clip(clean_line(l))
            if not s:
                continue
            if cur is None:
                cur = ['内容', [], '', '', '']
                slides.append(cur)
            cur[1].append(s)
    slides = [[t, [b for b in bs if b], n, im, tk] for t, bs, n, im, tk in slides]
    slides = [sl for sl in slides if sl[1] or len(slides) <= 1]
    if not slides:
        raise ValueError('没有解析出可用内容，请检查文档或描述。')
    # 没有任何页标题（纯段落/纯要点）→ 按每页 7 条自动切页、自动起章名
    if all(t == '内容' for t, _bs, _n, _im, _tk in slides):
        bullets = [b for _t, bs, _n, _im, _tk in slides for b in bs]
        slides = []
        for i in range(0, len(bullets), _MAX_BULLETS):
            k = i // _MAX_BULLETS
            if len(bullets) <= _MAX_BULLETS:
                name = '内容要点'
            elif k < len(_CN):
                name = '第%s部分' % _CN[k]
            else:
                name = '第%d部分' % (k + 1)
            slides.append([name, bullets[i:i + _MAX_BULLETS], '', '', ''])
    deck = deck or clip(clean_line(lines[0])) or '演示文稿'
    deck = clip(deck)
    # 要点过多自动分页
    out = []
    for t, bs, n, im, tk in slides:
        if not bs:
            continue
        for i in range(0, len(bs), _MAX_BULLETS):
            out.append((t if i == 0 else '%s·续' % t, bs[i:i + _MAX_BULLETS],
                        n if i == 0 else '', im if i == 0 else '',
                        tk if i == 0 else ''))
    return deck, out, None, None, None


def _fit_pages(slides, target):
    """把页面调整到 target 页（1~30）：先拆内容多的页，再补“深入”页，
    不够补互动页；多了则合并相邻页。返回 [(标题, 要点, 备注, 配图, 小结), ...]。"""
    if not target:
        return slides
    target = max(_MIN_PAGES, min(_MAX_PAGES, int(target)))
    sl = [list(x)[:5] + [''] * (5 - len(x)) for x in slides] or \
        [['内容', ['（本页待补充内容）'], '', '', '']]
    # 1) 拆分：把要点多于4条的页对半拆成两页
    while len(sl) < target:
        i = max(range(len(sl)), key=lambda k: len(sl[k][1]))
        if len(sl[i][1]) > 4:
            x = sl[i]
            half = (len(x[1]) + 1) // 2
            a = [x[0], x[1][:half], x[2], x[3], x[4]]
            b = ['%s·续' % x[0], x[1][half:], '', '', '']
            sl[i:i + 1] = [a, b]
            continue
        # 2) 深入：为尚未拆出的主题页加一页“·深入”聚焦单条要点
        done = False
        for i, x in enumerate(sl):
            if '·深入' in x[0] or '·续' in x[0] or not x[1]:
                continue
            if any(y[0].startswith(x[0] + '·深入') for y in sl):
                continue
            b = x[1][-1]
            sl.insert(i + 1, ['%s·深入' % x[0],
                              [b, '展开讲讲：背后的原因', '举一个身边的例子'],
                              x[2], x[3], ''])
            done = True
            break
        if done:
            continue
        # 3) 互动：补一页自由讨论
        sl.append(['交流与思考', ['大家有什么想问的？', '课后一起试一试',
                                  '欢迎分享你的看法'], '留 2 分钟自由讨论', 'star', ''])
    # 4) 合并：超页时把相邻最少要点的两页合为一页
    while len(sl) > target:
        best = min(range(len(sl) - 1),
                   key=lambda k: len(sl[k][1]) + len(sl[k + 1][1]))
        a, b = sl[best], sl[best + 1]
        a[1] = (a[1] + b[1])[:_MAX_BULLETS]
        a[2] = a[2] or b[2]
        a[3] = a[3] or b[3]
        a[4] = a[4] or b[4]
        del sl[best + 1]
    return [(x[0], x[1][:_MAX_BULLETS] or ['（本页待补充内容）'], x[2], x[3], x[4])
            for x in sl[:target]]


# ---------------- PPT 绘制：基础件 ----------------

def _textbox(slide, l, t, w, h):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    return tf


def _rect(slide, l, t, w, h, color):
    return _shp(slide, 'RECTANGLE', l, t, w, h, fill=color)


def _shp(slide, kind, l, t, w, h, fill=None, line=None, lw=1.0, rot=0):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(getattr(MSO_SHAPE, kind),
                                Inches(l), Inches(t), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = _rgb(fill)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = _rgb(line)
        sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    if rot:
        sh.rotation = rot
    return sh


def _cap(sh, text, size, color, bold=False):
    tf = sh.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = _rgb(color)
    _zh(r)


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


# ---------------- PPT 绘制：约 38 种自动矢量配图 ----------------
# 每个画图函数签名统一 (s, L, T, W, H, sch)，在 (L,T,W,H) 卡片区内作画。

def _ball(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.3
    r = 1.15
    _shp(s, 'OVAL', cx - r, cy - r, 2 * r, 2 * r, fill=_W, line=_INK, lw=1.5)
    _shp(s, 'PENTAGON', cx - 0.34, cy - 0.34, 0.68, 0.68, fill=_INK)
    for dx, dy in [(-0.92, -0.55), (0.92, -0.55), (-0.72, 0.82), (0.72, 0.82)]:
        _shp(s, 'PENTAGON', cx + dx - 0.22, cy + dy - 0.22, 0.44, 0.44, fill=_INK)


def _field(s, L, T, W, H, g, dots=False):
    gx, gy, gw, gh = L + 0.28, T + 0.65, W - 0.56, 2.95
    _shp(s, 'ROUNDED_RECTANGLE', gx, gy, gw, gh, fill=_GRASS, line=_W, lw=2)
    _shp(s, 'RECTANGLE', gx + 0.2, gy + 0.18, gw - 0.4, gh - 0.36, fill=None, line=_W, lw=1.5)
    _shp(s, 'RECTANGLE', gx + gw / 2 - 0.02, gy + 0.18, 0.04, gh - 0.36, fill=_W)
    _shp(s, 'OVAL', gx + gw / 2 - 0.4, gy + gh / 2 - 0.4, 0.8, 0.8, fill=None, line=_W, lw=1.5)
    _shp(s, 'RECTANGLE', gx + 0.2, gy + gh / 2 - 0.55, 0.42, 1.1, fill=None, line=_W, lw=1.2)
    _shp(s, 'RECTANGLE', gx + gw - 0.62, gy + gh / 2 - 0.55, 0.42, 1.1, fill=None, line=_W, lw=1.2)
    if dots:
        cx = gx + gw / 2
        for dx, dy in [(-1.1, -0.9), (0, -1.0), (1.1, -0.9), (-0.55, 0.1), (0.55, 0.1)]:
            _shp(s, 'OVAL', cx + dx - 0.15, gy + gh / 2 + dy - 0.15, 0.3, 0.3, fill=_RED, line=_W, lw=1)
        for dx, dy in [(0, 1.0), (-1.0, 0.7), (1.0, 0.7)]:
            _shp(s, 'OVAL', cx + dx - 0.15, gy + gh / 2 + dy - 0.15, 0.3, 0.3, fill=_W, line=g['pri'], lw=1)


def _board(s, L, T, W, H, g):
    _field(s, L, T, W, H, g, dots=True)


def _goal(s, L, T, W, H, g):
    x, y, w, h = L + 0.55, T + 1.1, W - 1.1, 2.1
    _shp(s, 'RECTANGLE', x, y, w, h, fill=None, line=_W, lw=3)
    for i in range(1, 5):
        _shp(s, 'RECTANGLE', x + i * w / 5, y, 0.02, h, fill=g['acc'])
        _shp(s, 'RECTANGLE', x, y + i * h / 5, w, 0.02, fill=g['acc'])
    bx, by = L + W / 2 - 0.35, T + 3.35
    _shp(s, 'OVAL', bx, by, 0.7, 0.7, fill=_W, line=_INK, lw=1.2)
    _shp(s, 'PENTAGON', bx + 0.16, by + 0.16, 0.38, 0.38, fill=_INK)


def _whistle(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.35
    _shp(s, 'ROUNDED_RECTANGLE', cx - 1.0, cy - 0.45, 1.7, 0.9, fill=g['pri'], line=g['bg'], lw=1)
    _shp(s, 'RECTANGLE', cx + 0.55, cy - 0.25, 0.55, 0.5, fill=g['acc'])
    _shp(s, 'OVAL', cx + 0.85, cy - 0.18, 0.34, 0.34, fill=_W, line=g['bg'], lw=1)
    _shp(s, 'DONUT', cx - 1.35, cy - 0.3, 0.6, 0.6, fill=_GOLD)


def _boots(s, L, T, W, H, g):
    x, y = L + 0.7, T + 2.4
    _shp(s, 'ROUNDED_RECTANGLE', x, y - 0.7, 0.9, 1.1, fill=g['pri'], line=g['bg'], lw=1)
    _shp(s, 'ROUNDED_RECTANGLE', x, y + 0.1, 2.2, 0.55, fill=g['acc'], line=g['bg'], lw=1)
    _shp(s, 'RECTANGLE', x, y + 0.6, 2.2, 0.16, fill=_INK)
    for i in range(3):
        _shp(s, 'RECTANGLE', x + 0.15, y - 0.5 + i * 0.28, 0.6, 0.05, fill=_W)


def _trophy(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'OVAL', cx - 0.62, T + 0.75, 0.5, 0.75, fill=None, line=_GOLD, lw=3)
    _shp(s, 'OVAL', cx + 0.12, T + 0.75, 0.5, 0.75, fill=None, line=_GOLD, lw=3)
    _shp(s, 'ROUNDED_RECTANGLE', cx - 0.7, T + 0.7, 1.4, 1.25, fill=_GOLD, line=_GOLD_D, lw=1)
    _shp(s, 'STAR_5_POINT', cx - 0.3, T + 0.95, 0.6, 0.6, fill=_W)
    _shp(s, 'RECTANGLE', cx - 0.14, T + 1.95, 0.28, 0.5, fill=_GOLD_D)
    _shp(s, 'ROUNDED_RECTANGLE', cx - 0.55, T + 2.45, 1.1, 0.35, fill=_GOLD_D)


def _medal(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.55
    _shp(s, 'ISOSCELES_TRIANGLE', cx - 0.55, T + 0.65, 0.55, 0.9, fill=g['pri'], rot=180)
    _shp(s, 'ISOSCELES_TRIANGLE', cx, T + 0.65, 0.55, 0.9, fill=g['acc'], rot=180)
    _shp(s, 'OVAL', cx - 0.65, cy - 0.65, 1.3, 1.3, fill=_GOLD, line=_GOLD_D, lw=2)
    _shp(s, 'STAR_5_POINT', cx - 0.32, cy - 0.32, 0.64, 0.64, fill=_W)


def _heart(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'HEART', cx - 0.85, T + 1.1, 1.7, 1.6, fill=_RED, line='C0292E', lw=1)
    _shp(s, 'RECTANGLE', cx - 0.9, T + 1.8, 0.5, 0.06, fill=_W)
    _shp(s, 'RECTANGLE', cx + 0.4, T + 1.8, 0.5, 0.06, fill=_W)


def _coin(s, L, T, W, H, g):
    cx = L + W / 2
    _cap(_shp(s, 'OVAL', cx - 0.8, T + 0.85, 1.6, 1.6, fill=_GOLD, line=_GOLD_D, lw=2),
         '￥', 34, '8A5A00', True)
    _shp(s, 'RECTANGLE', L + 0.5, T + 2.9, 0.5, 0.7, fill=g['pri'])
    _shp(s, 'RECTANGLE', L + 1.15, T + 2.6, 0.5, 1.0, fill=g['acc'])
    _shp(s, 'RECTANGLE', L + 1.8, T + 3.05, 0.5, 0.55, fill=_GOLD)


def _chart_bar(s, L, T, W, H, g):
    base = T + 3.4
    cols = [g['pri'], g['acc'], _GOLD, g['pri']]
    for i, c in enumerate(cols):
        h = [1.3, 2.1, 1.0, 2.6][i]
        _shp(s, 'ROUNDED_RECTANGLE', L + 0.5 + i * 0.78, base - h, 0.55, h, fill=c)
    _shp(s, 'RECTANGLE', L + 0.4, base, W - 0.8, 0.05, fill=_INK)


def _chart_pie(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.35
    _shp(s, 'OVAL', cx - 1.1, cy - 1.1, 2.2, 2.2, fill=g['soft'], line=g['pri'], lw=1)
    _shp(s, 'PIE', cx - 1.1, cy - 1.1, 2.2, 2.2, fill=g['pri'])
    _shp(s, 'BLOCK_ARC', cx - 1.1, cy - 1.1, 2.2, 2.2, fill=_GOLD, rot=140)


def _crowd(s, L, T, W, H, g):
    cols = [g['pri'], g['acc'], _GOLD, g['pri'], g['acc']]
    for i, c in enumerate(cols):
        hx = L + 0.45 + i * 0.66
        hy = T + 1.55 - (0.35 if i in (1, 3) else 0)
        _shp(s, 'OVAL', hx, hy, 0.42, 0.42, fill=c)
        _shp(s, 'ROUNDED_RECTANGLE', hx - 0.06, hy + 0.42, 0.54, 0.75, fill=c)


def _timeline(s, L, T, W, H, g):
    y = T + 2.15
    _shp(s, 'RECTANGLE', L + 0.4, y, W - 0.8, 0.08, fill=g['pri'])
    for lx in [L + 0.55, L + W / 2 - 0.22, L + W - 0.95]:
        _shp(s, 'OVAL', lx, y - 0.17, 0.42, 0.42, fill=g['acc'], line=g['pri'], lw=1.5)
    _shp(s, 'PENTAGON', L + 0.5, T + 1.05, 0.5, 0.5, fill=g['pri'])


def _clock(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.35
    _shp(s, 'OVAL', cx - 1.1, cy - 1.1, 2.2, 2.2, fill=_W, line=g['pri'], lw=2.5)
    _shp(s, 'RECTANGLE', cx - 0.03, cy - 0.75, 0.06, 0.8, fill=_INK)
    _shp(s, 'RECTANGLE', cx, cy - 0.03, 0.6, 0.06, fill=g['acc'])
    _shp(s, 'OVAL', cx - 0.09, cy - 0.09, 0.18, 0.18, fill=g['pri'])


def _calendar(s, L, T, W, H, g):
    x, y, w, h = L + 0.65, T + 1.2, W - 1.3, 2.3
    _shp(s, 'ROUNDED_RECTANGLE', x, y, w, h, fill=_W, line=g['pri'], lw=2)
    _shp(s, 'RECTANGLE', x, y, w, 0.5, fill=g['pri'])
    for i in range(7):
        for j in range(3):
            _shp(s, 'OVAL', x + 0.2 + i * (w - 0.4) / 7, y + 0.75 + j * 0.5, 0.12, 0.12, fill=g['soft'])
    _shp(s, 'OVAL', x + 0.2 + 2 * (w - 0.4) / 7, y + 1.25, 0.16, 0.16, fill=_RED)


def _book(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'ROUNDED_RECTANGLE', cx - 1.35, T + 1.1, 1.3, 1.9, fill=g['pri'], rot=-8)
    _shp(s, 'ROUNDED_RECTANGLE', cx + 0.05, T + 1.1, 1.3, 1.9, fill=g['acc'], rot=8)


def _doc(s, L, T, W, H, g):
    x, y, w, h = L + 1.1, T + 1.0, 1.7, 2.3
    _shp(s, 'FOLDED_CORNER', x, y, w, h, fill=_W, line=g['pri'], lw=1.5)
    for i in range(4):
        _shp(s, 'RECTANGLE', x + 0.28, y + 0.55 + i * 0.38, w - 0.6, 0.06, fill=g['acc'])


def _graduation(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'DIAMOND', cx - 1.2, T + 1.4, 2.4, 1.0, fill=g['bg'])
    _shp(s, 'ROUNDED_RECTANGLE', cx - 0.7, T + 1.95, 1.4, 0.7, fill=g['pri'])
    _shp(s, 'RECTANGLE', cx + 1.0, T + 1.7, 0.06, 1.0, fill=_GOLD)
    _shp(s, 'OVAL', cx + 0.94, T + 2.7, 0.2, 0.2, fill=_GOLD)


def _gear(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.35
    _shp(s, 'GEAR_6', cx - 1.15, cy - 1.15, 2.3, 2.3, fill=g['pri'], line=g['bg'], lw=1)
    _shp(s, 'OVAL', cx - 0.4, cy - 0.4, 0.8, 0.8, fill=_W)


def _monitor(s, L, T, W, H, g):
    x, y, w, h = L + 0.7, T + 1.1, W - 1.4, 1.9
    _shp(s, 'ROUNDED_RECTANGLE', x, y, w, h, fill=g['bg'], line=g['pri'], lw=1.5)
    _shp(s, 'RECTANGLE', x + 0.2, y + 0.2, w - 0.4, h - 0.5, fill=g['soft'])
    _shp(s, 'RECTANGLE', L + W / 2 - 0.1, y + h, 0.2, 0.35, fill=g['pri'])
    _shp(s, 'ROUNDED_RECTANGLE', L + W / 2 - 0.6, y + h + 0.32, 1.2, 0.16, fill=g['pri'])


def _bulb(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.1
    _shp(s, 'OVAL', cx - 0.75, cy - 0.75, 1.5, 1.5, fill=_GOLD, line=_GOLD_D, lw=1)
    _shp(s, 'ROUNDED_RECTANGLE', cx - 0.3, cy + 0.7, 0.6, 0.55, fill=g['pri'])
    for dx in (-0.9, 0.9):
        _shp(s, 'RECTANGLE', cx + dx - 0.03, cy - 0.03, 0.5, 0.06, fill=g['acc'], rot=0)


def _target(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.35
    for r, c in [(1.15, g['acc']), (0.75, _W), (0.4, _RED)]:
        _shp(s, 'OVAL', cx - r, cy - r, 2 * r, 2 * r, fill=c, line=g['pri'], lw=1)


def _flag(s, L, T, W, H, g):
    x = L + 1.0
    _shp(s, 'RECTANGLE', x, T + 0.9, 0.1, 3.0, fill=_INK)
    _shp(s, 'RIGHT_TRIANGLE', x + 0.1, T + 1.0, 1.7, 1.0, fill=g['pri'], rot=90)
    _shp(s, 'OVAL', x - 0.05, T + 0.75, 0.22, 0.22, fill=_GOLD)


def _rocket(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'TEAR', cx - 0.55, T + 1.0, 1.1, 1.9, fill=g['pri'], line=g['bg'], lw=1, rot=135)
    _shp(s, 'OVAL', cx - 0.22, T + 1.55, 0.44, 0.44, fill=_W, line=g['acc'], lw=1.5)
    _shp(s, 'ISOSCELES_TRIANGLE', cx - 0.35, T + 3.0, 0.7, 0.6, fill=_GOLD, rot=180)


def _envelope(s, L, T, W, H, g):
    x, y, w, h = L + 0.75, T + 1.5, W - 1.5, 1.8
    _shp(s, 'RECTANGLE', x, y, w, h, fill=_W, line=g['pri'], lw=1.5)
    _shp(s, 'ISOSCELES_TRIANGLE', x, y, w, h * 0.7, fill=None, line=g['pri'], lw=1.5, rot=180)


def _phone(s, L, T, W, H, g):
    x, y = L + W / 2 - 0.6, T + 1.1
    _shp(s, 'ROUNDED_RECTANGLE', x, y, 1.2, 2.3, fill=g['bg'], line=g['pri'], lw=1.5)
    _shp(s, 'RECTANGLE', x + 0.12, y + 0.28, 0.96, 1.6, fill=g['acc'])
    _shp(s, 'OVAL', x + 0.5, y + 1.98, 0.2, 0.2, fill=_W)


def _search(s, L, T, W, H, g):
    cx, cy = L + W / 2 - 0.25, T + 2.1
    _shp(s, 'OVAL', cx - 0.75, cy - 0.75, 1.5, 1.5, fill=None, line=g['pri'], lw=3)
    _shp(s, 'RECTANGLE', cx + 0.55, cy + 0.4, 0.7, 0.18, fill=g['pri'], rot=45)


def _shield(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'ROUNDED_RECTANGLE', cx - 0.85, T + 1.15, 1.7, 1.5, fill=g['pri'])
    _shp(s, 'ISOSCELES_TRIANGLE', cx - 0.85, T + 2.4, 1.7, 1.1, fill=g['pri'], rot=180)
    _cap(_shp(s, 'OVAL', cx - 0.35, T + 1.55, 0.7, 0.7, fill=_W), '✓', 20, g['pri'], True)


def _check(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.35
    _cap(_shp(s, 'OVAL', cx - 1.0, cy - 1.0, 2.0, 2.0, fill=g['pri'], line=g['bg'], lw=1.5),
         '✓', 46, _W, True)


def _question(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.35
    _cap(_shp(s, 'OVAL', cx - 1.0, cy - 1.0, 2.0, 2.0, fill=g['soft'], line=g['pri'], lw=2),
         '?', 46, g['pri'], True)


def _map(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'ROUNDED_RECTANGLE', L + 0.7, T + 1.5, W - 1.4, 1.8, fill=g['soft'], line=g['acc'], lw=1.2)
    _shp(s, 'RECTANGLE', L + 1.3, T + 1.5, 0.05, 1.8, fill=g['acc'])
    _shp(s, 'RECTANGLE', L + 2.4, T + 1.5, 0.05, 1.8, fill=g['acc'])
    _shp(s, 'OVAL', cx - 0.35, T + 0.95, 0.7, 0.7, fill=_RED, line='C0292E', lw=1)
    _shp(s, 'ISOSCELES_TRIANGLE', cx - 0.3, T + 1.5, 0.6, 0.6, fill=_RED, rot=180)


def _lock(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'BLOCK_ARC', cx - 0.6, T + 1.15, 1.2, 1.5, fill=None, line=g['pri'], lw=3, rot=180)
    _shp(s, 'ROUNDED_RECTANGLE', cx - 0.85, T + 1.9, 1.7, 1.4, fill=g['pri'], line=g['bg'], lw=1)
    _cap(_shp(s, 'OVAL', cx - 0.22, T + 2.35, 0.44, 0.44, fill=_GOLD), '', 12, _W)


def _cloud(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'CLOUD', cx - 1.3, T + 1.3, 2.6, 1.7, fill=g['acc'], line=g['pri'], lw=1)
    _shp(s, 'RECTANGLE', cx - 0.9, T + 2.9, 1.8, 0.08, fill=g['pri'])


def _music(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'RECTANGLE', cx + 0.35, T + 1.2, 0.08, 1.7, fill=_INK)
    _shp(s, 'ROUNDED_RECTANGLE', cx + 0.35, T + 1.2, 0.6, 0.3, fill=_INK, rot=20)
    _shp(s, 'OVAL', cx - 0.25, T + 2.6, 0.6, 0.45, fill=g['pri'])
    _shp(s, 'OVAL', cx + 0.75, T + 2.35, 0.6, 0.45, fill=g['acc'])


def _sun(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.35
    _shp(s, 'SUN', cx - 1.2, cy - 1.2, 2.4, 2.4, fill=_GOLD, line=_GOLD_D, lw=1)


def _star(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'STAR_5_POINT', cx - 0.9, T + 0.95, 1.8, 1.8, fill=_GOLD, line=_GOLD_D, lw=1)
    _shp(s, 'STAR_5_POINT', L + 0.55, T + 2.6, 0.5, 0.5, fill=g['acc'])
    _shp(s, 'STAR_5_POINT', L + W - 1.05, T + 0.85, 0.55, 0.55, fill=g['pri'])


def _sparkle(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'STAR_5_POINT', cx - 0.7, T + 1.05, 1.4, 1.4, fill=g['pri'])
    _shp(s, 'OVAL', L + 0.55, T + 2.7, 0.4, 0.4, fill=g['acc'])
    _shp(s, 'OVAL', L + W - 0.95, T + 0.9, 0.45, 0.45, fill=_GOLD)


def _ball_basket(s, L, T, W, H, g):
    cx, cy, r = L + W / 2, T + 2.3, 1.1
    _shp(s, 'OVAL', cx - r, cy - r, 2 * r, 2 * r, fill='E8730C', line='B5560A', lw=1.5)
    _shp(s, 'RECTANGLE', cx - 0.02, cy - r, 0.04, 2 * r, fill=_INK)
    _shp(s, 'RECTANGLE', cx - r, cy - 0.02, 2 * r, 0.04, fill=_INK)
    _shp(s, 'OVAL', cx - 0.6, cy - r, 1.2, 2 * r, fill=None, line=_INK, lw=1.2)


def _ball_volley(s, L, T, W, H, g):
    cx, cy, r = L + W / 2, T + 2.3, 1.1
    _shp(s, 'OVAL', cx - r, cy - r, 2 * r, 2 * r, fill='F5F0E6', line='C9BC9A', lw=1.5)
    for rot in (0, 60, 120):
        _shp(s, 'OVAL', cx - 0.35, cy - r + 0.05, 0.7, 2 * r - 0.1, fill=None,
             line=g['pri'], lw=1.2, rot=rot)


def _ball_pong(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'OVAL', cx - 0.95, T + 1.0, 1.5, 1.5, fill='D0342C', line='9E221C', lw=1)
    _shp(s, 'ROUNDED_RECTANGLE', cx - 0.28, T + 2.35, 0.4, 1.1, fill='B5793B', rot=-20)
    _shp(s, 'OVAL', cx + 0.7, T + 2.6, 0.45, 0.45, fill='F6F6F0', line='BDBDA6', lw=1)


def _ball_tennis(s, L, T, W, H, g):
    cx, cy, r = L + W / 2, T + 2.3, 1.05
    _shp(s, 'OVAL', cx - r, cy - r, 2 * r, 2 * r, fill='D8E84A', line='A7B82E', lw=1.5)
    _shp(s, 'OVAL', cx - 0.95, cy - 0.55, 1.9, 1.1, fill=None, line=_W, lw=1.4, rot=0)


def _shuttle(s, L, T, W, H, g):
    cx = L + W / 2
    _shp(s, 'ISOSCELES_TRIANGLE', cx - 0.75, T + 1.0, 1.5, 1.9, fill=_W, line='AEB4BD', lw=1.2)
    _shp(s, 'OVAL', cx - 0.42, T + 2.75, 0.84, 0.7, fill='E8C39E', line='B98A54', lw=1)


def _dumbbell(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.35
    _shp(s, 'RECTANGLE', cx - 0.9, cy - 0.12, 1.8, 0.24, fill='9AA0A6')
    _shp(s, 'ROUNDED_RECTANGLE', cx - 1.25, cy - 0.6, 0.4, 1.2, fill=g['pri'])
    _shp(s, 'ROUNDED_RECTANGLE', cx + 0.85, cy - 0.6, 0.4, 1.2, fill=g['pri'])
    _shp(s, 'ROUNDED_RECTANGLE', cx - 0.82, cy - 0.4, 0.28, 0.8, fill=g['acc'])
    _shp(s, 'ROUNDED_RECTANGLE', cx + 0.54, cy - 0.4, 0.28, 0.8, fill=g['acc'])


def _pool(s, L, T, W, H, g):
    x, y, w, h = L + 0.5, T + 1.3, W - 1.0, 2.3
    _shp(s, 'ROUNDED_RECTANGLE', x, y, w, h, fill='4FB3E8', line='2E86C1', lw=1.5)
    for i in range(3):
        _shp(s, 'RECTANGLE', x + 0.1, y + 0.5 + i * 0.55, w - 0.2, 0.04, fill=_W)
    for i in range(6):
        _shp(s, 'OVAL', x + 0.2 + i * (w - 0.4) / 5, y + 0.18, 0.12, 0.12,
             fill=_GOLD if i % 2 else _W)


def _track(s, L, T, W, H, g):
    cx, cy = L + W / 2, T + 2.3
    _shp(s, 'DONUT', cx - 1.4, cy - 1.15, 2.8, 2.3, fill='D1603F', line='A8492C', lw=1)
    _shp(s, 'OVAL', cx - 0.75, cy - 0.55, 1.5, 1.1, fill=_GRASS)


def _court_hard(s, L, T, W, H, g):
    x, y, w, h = L + 0.45, T + 1.0, W - 0.9, 2.7
    _shp(s, 'ROUNDED_RECTANGLE', x, y, w, h, fill='C8956B', line=_W, lw=2)
    _shp(s, 'RECTANGLE', x + 0.15, y + 0.15, w - 0.3, h - 0.3, fill=None, line=_W, lw=1.5)
    _shp(s, 'RECTANGLE', x + w / 2 - 0.02, y + 0.15, 0.04, h - 0.3, fill=_W)
    _shp(s, 'OVAL', x + w / 2 - 0.35, y + h / 2 - 0.35, 0.7, 0.7, fill=None, line=_W, lw=1.5)
    _shp(s, 'OVAL', x + w / 2 - 0.2, y + 0.12, 0.4, 0.24, fill='E8730C')
    _shp(s, 'OVAL', x + w / 2 - 0.2, y + h - 0.36, 0.4, 0.24, fill='E8730C')


def _table(s, L, T, W, H, g):
    x, y, w, h = L + 0.4, T + 1.7, W - 0.8, 1.7
    _shp(s, 'ROUNDED_RECTANGLE', x, y, w, h, fill='2F5C8A', line=_W, lw=2)
    _shp(s, 'RECTANGLE', x + w / 2 - 0.015, y, 0.03, h, fill=_W)
    _shp(s, 'RECTANGLE', x, y + h / 2 - 0.06, w, 0.12, fill=None, line=_W, lw=0.75)
    _shp(s, 'OVAL', x + w * 0.22, T + 1.25, 0.35, 0.35, fill='F26B1F', line='C24E12', lw=1)


_ICONS = {
    'ball': _ball, 'field': _field, 'board': _board, 'goal': _goal,
    'whistle': _whistle, 'boots': _boots, 'trophy': _trophy, 'medal': _medal,
    'heart': _heart, 'coin': _coin, 'chart_bar': _chart_bar, 'chart_pie': _chart_pie,
    'crowd': _crowd, 'timeline': _timeline, 'clock': _clock, 'calendar': _calendar,
    'book': _book, 'doc': _doc, 'graduation': _graduation, 'gear': _gear,
    'monitor': _monitor, 'bulb': _bulb, 'target': _target, 'flag': _flag,
    'rocket': _rocket, 'envelope': _envelope, 'phone': _phone, 'search': _search,
    'shield': _shield, 'check': _check, 'question': _question, 'map': _map,
    'lock': _lock, 'cloud': _cloud, 'music': _music, 'sun': _sun,
    'star': _star, 'sparkle': _sparkle,
    'ball_basket': _ball_basket, 'ball_volley': _ball_volley,
    'ball_pong': _ball_pong, 'ball_tennis': _ball_tennis, 'shuttle': _shuttle,
    'dumbbell': _dumbbell, 'pool': _pool, 'track': _track,
    'court_hard': _court_hard, 'table': _table,
}


def _draw_pic(s, spec, scheme, seed=0, left=False):
    """在卡片区按主题画一张矢量插画（离线、无外部素材）；按种子轮换款式。"""
    L, T, W, H = (0.55 if left else 8.95), 1.5, 3.95, 4.75
    name = _icon_for(spec, seed)
    _shp(s, 'ROUNDED_RECTANGLE', L, T, W, H, fill=_W, line=scheme['acc'], lw=1.2)
    try:
        _ICONS.get(name, _sparkle)(s, L, T, W, H, scheme)
    except Exception:                               # 个别形状名不可用时兜底
        _sparkle(s, L, T, W, H, scheme)
    cap = _textbox(s, L, T + H - 0.5, W, 0.4)
    p = cap.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = '自动配图 · %s' % _PICS.get(name, '示意')
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = _rgb(scheme['pri'])
    _zh(r)


# ---------------- PPT 绘制：各类幻灯片 ----------------

def _title_slide(prs, deck, scheme, total, audience=None):
    s = _blank(prs)
    _rect(s, 0, 0, 13.333, 7.5, scheme['bg'])
    _rect(s, 0.9, 4.05, 1.6, 0.07, scheme['acc'])
    tf = _textbox(s, 0.9, 2.3, 11.5, 1.6)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = deck
    r.font.size = Pt(40)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    _zh(r)
    tf2 = _textbox(s, 0.9, 4.35, 11.5, 1.2)
    p2 = tf2.paragraphs[0]
    r2 = p2.add_run()
    head = ('面向%s · ' % audience) if audience else ''
    r2.text = '%s共 %d 页内容 · 自动生成于 %s' % (head, total, time.strftime('%Y-%m-%d'))
    r2.font.size = Pt(16)
    r2.font.color.rgb = _rgb(scheme['acc'])
    _zh(r2)


def _toc_slide(prs, deck, slides, scheme):
    s = _blank(prs)
    _band(s, '目录', 0, scheme, deck)
    tf = _textbox(s, 1.0, 1.6, 11.3, 5.4)
    for i, item in enumerate(slides):
        t = item[0]
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = '%02d' % (i + 1)
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = _rgb(scheme['pri'])
        _zh(r)
        r2 = p.add_run()
        r2.text = '  ' + t
        r2.font.size = Pt(16)
        r2.font.color.rgb = RGBColor(0x1F, 0x23, 0x29)
        _zh(r2)


def _band(s, title, idx, scheme, deck):
    _rect(s, 0, 0, 13.333, 1.05, scheme['pri'])
    _rect(s, 0, 1.05, 13.333, 0.05, scheme['acc'])
    tf = _textbox(s, 0.55, 0.14, 10.6, 0.8)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    r.font.size = Pt(_TITLE_PT)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    _zh(r)
    if idx:
        tf2 = _textbox(s, 11.6, 0.3, 1.5, 0.5)
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.RIGHT
        r2 = p2.add_run()
        r2.text = '%d' % idx
        r2.font.size = Pt(16)
        r2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _zh(r2)


def _content_slide(prs, deck, title, bullets, idx, scheme, note='', img='',
                   takeaway='', seed=0):
    s = _blank(prs)
    _band(s, title, idx, scheme, deck)
    _rect(s, 0, 7.22, 13.333, 0.28, scheme['soft'])
    variant = seed % 3                              # 每套 PPT 的版式轮变
    glyph = _GLYPHS[seed % len(_GLYPHS)]
    left_pic = bool(img) and (variant == 1 or (variant == 2 and idx % 2 == 0))
    if img:
        _draw_pic(s, img, scheme, seed + idx, left=left_pic)
    if note:
        s.notes_slide.notes_text_frame.text = note
    if not img:
        bx, bw = 0.8, 11.7
    elif left_pic:
        bx, bw = 4.85, 7.9
    else:
        bx, bw = 0.8, 7.6
    bh = 4.62 if takeaway else 5.5
    tf = _textbox(s, bx, 1.5, bw, bh)
    # 按要点条数与实际行数动态分配行间距，让要点均匀铺满整页（消除顶部堆叠、下方大片留白）
    n = len(bullets)
    cpl = max(12, int(bw * 72 / (_BODY_PT * 1.05)))          # 每行大约可容纳的字数
    lines = sum(max(1, (len(b) + cpl - 1) // cpl) for b in bullets)
    gap = (bh * 72 - lines * _BODY_PT * 1.35) / max(n - 1, 1) if n > 1 else 12
    gap = max(6, min(30, gap))
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        rb = p.add_run()
        rb.text = glyph + ' '
        rb.font.size = Pt(_BODY_PT)
        rb.font.color.rgb = _rgb(scheme['pri'])
        _zh(rb)
        head, sep, tail = b.partition('：')
        if sep and 0 < len(head) <= 14:             # “术语：说明”式要点把术语加粗
            rh = p.add_run()
            rh.text = head + '：'
            rh.font.size = Pt(_BODY_PT)
            rh.font.bold = True
            rh.font.color.rgb = RGBColor(0x1F, 0x23, 0x29)
            _zh(rh)
            rt = p.add_run()
            rt.text = tail
        else:
            rt = p.add_run()
            rt.text = b
        rt.font.size = Pt(_BODY_PT)
        rt.font.color.rgb = RGBColor(0x1F, 0x23, 0x29)
        _zh(rt)
    if takeaway:
        bar = _shp(s, 'ROUNDED_RECTANGLE', bx, 6.35, bw, 0.55,
                   fill=scheme['soft'], line=scheme['acc'], lw=1.0)
        _cap(bar, '★ ' + takeaway, 13, scheme['pri'], True)


def _end_slide(prs, deck, scheme):
    s = _blank(prs)
    _rect(s, 0, 0, 13.333, 7.5, scheme['bg'])
    tf = _textbox(s, 1.5, 3.1, 10.3, 1.3)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = '感谢聆听'
    r.font.size = Pt(44)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    _zh(r)
    tf2 = _textbox(s, 1.5, 4.5, 10.3, 0.8)
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = deck
    r2.font.size = Pt(16)
    r2.font.color.rgb = _rgb(scheme['acc'])
    _zh(r2)


def build_ppt(deck, slides, out_path, scheme, audience=None, seed=0):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    content = slides
    _title_slide(prs, deck, scheme, len(content), audience)
    if len(content) >= 4:
        _toc_slide(prs, deck, content, scheme)
    for i, item in enumerate(content):
        t, bs = item[0], item[1]
        note = item[2] if len(item) > 2 else ''
        img = item[3] if len(item) > 3 else ''
        tk = item[4] if len(item) > 4 else ''
        _content_slide(prs, deck, t, bs, i + 1, scheme, note, img, tk, seed)
    _end_slide(prs, deck, scheme)
    prs.save(out_path)
    return out_path


# ---------------- 计次轮转（配色 / 版式款式） ----------------

def _dir():
    base = os.environ.get('LOCALAPPDATA') or os.path.join(os.path.expanduser('~'), '.local')
    d = os.path.join(base, '填表工具')
    os.makedirs(d, exist_ok=True)
    return d


def _count_file():
    return os.path.join(_dir(), 'ppt_scheme.count')


def _read_bump(path):
    try:
        with open(path, encoding='utf-8') as fh:
            n = int(fh.read().strip() or 0)
    except (ValueError, OSError):
        n = 0
    try:
        with open(path, 'w') as fh:
            fh.write(str(n + 1))
    except OSError:
        pass
    return n


def next_scheme():
    """三套配色依次轮流出，重启工具仍记次。"""
    return SCHEMES[_read_bump(_count_file()) % len(SCHEMES)]


def _style_seed():
    """版式轮变计数器：每生成一次自增，驱动配图款式/位置/符号变化。"""
    return _read_bump(os.path.join(_dir(), 'ppt_variant.count'))


# ---------------- 对外主入口 ----------------

def generate_ppt(text, out_dir, title=None, scheme_idx=None, pages=None):
    """文本(或 extract_source 的结果) → .pptx；返回结果信息字典。

    pages 不为空时把内容自适应到该页数（1~30）；为空则用需求单里抽到的页数。
    每次生成用递增的版式种子，使配图款式、位置与要点符号在多次间自动变化。
    """
    deck, slides, audience, p_scheme, p_target = parse_outline(text)
    if title and title.strip():
        deck = clip(title.strip())
    target = pages if pages else p_target
    slides = _fit_pages(slides, target)
    if scheme_idx is not None:
        scheme = SCHEMES[scheme_idx]
    elif p_scheme is not None:          # 需求单里写了颜色就照着用
        scheme = SCHEMES[p_scheme]
    else:
        scheme = next_scheme()
    seed = _style_seed()
    safe = re.sub(r'[\\/:*?"<>|]', '', deck)[:24] or '演示文稿'
    out = os.path.join(out_dir, '%s-%s.pptx' % (safe, time.strftime('%H%M%S')))
    build_ppt(deck, slides, out, scheme, audience, seed)
    return {'path': out, 'deck': deck, 'slides': len(slides),
            'bullets': sum(len(it[1]) for it in slides),
            'notes': sum(1 for it in slides if len(it) > 2 and it[2]),
            'pics': sum(1 for it in slides if len(it) > 3 and it[3]),
            'seed': seed, 'scheme': scheme['name'], 'scheme_idx': SCHEMES.index(scheme)}
