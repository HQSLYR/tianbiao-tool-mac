# -*- coding: utf-8 -*-
"""投标演示页面生成引擎 v6：内容驱动导航与三套独立布局。

同一内容前三次依次生成三套方案，第四次起循环复用三套风格继续生成；
SQLite 持久化计次，失败不占次数，旧文件永不覆盖。
业务能力决定导航环节、页面结构和领域内容，长句只用于解析，不原文上屏。
三套分别采用工作台、资源驾驶舱、左右协同布局与蓝、紫、绿渐变背景。
红线：sanitize 替换 参数/需求/条款/示例；无客户标识；单文件双击即跑。
页面使用本地演练数据，不连接实际业务系统或在线 AI。
"""
import re
import hashlib
import os
import sqlite3
import tempfile
import unicodedata
from pathlib import Path
from functools import lru_cache


SCHEMES = ('晴空蓝 · 业务工作台', '星云紫 · 资源驾驶舱', '青岚绿 · 协同中心')


def _page_scheme(variant):
    """页面可见处用的方案名：去掉前面的配色颜色词（晴空蓝/星云紫/青岚绿），
    只留“业务工作台”等功能名，避免页面上出现颜色字样。"""
    return SCHEMES[variant - 1].split('·')[-1].strip()


def content_key(text):
    """忽略全半角及空白差异；不保存输入原文，也不依赖文件名。"""
    normalized = re.sub(r'\s+', '', unicodedata.normalize('NFKC', text or ''))
    return hashlib.sha256(normalized.encode('utf-8')).hexdigest()


def history_file():
    base = os.environ.get('LOCALAPPDATA') or str(Path.home() / '.local' / 'share')
    return Path(base) / '填表工具' / 'demo_history.sqlite3'


def generate_demo(text, out_dir, history_path=None):
    """持久化计次：前三次依次出三套，第四次起循环复用三套风格；保存失败不计次。"""
    if not text or not text.strip():
        raise ValueError('请先填写有效内容。')
    mods = parse_params(text)
    if not mods:
        raise ValueError('未解析到有效业务。')
    key = content_key(text)
    db = Path(history_path) if history_path is not None else history_file()
    db.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db), timeout=15)
    pending = None
    output = None
    committed = False
    try:
        conn.execute('CREATE TABLE IF NOT EXISTS demos ('
                     'content_key TEXT NOT NULL, variant INTEGER CHECK(variant BETWEEN 1 AND 3), '
                     'output TEXT NOT NULL, PRIMARY KEY(content_key, variant))')
        conn.execute('CREATE TABLE IF NOT EXISTS demo_cycles ('
                     'seq INTEGER PRIMARY KEY AUTOINCREMENT, content_key TEXT NOT NULL, '
                     'variant INTEGER NOT NULL, output TEXT NOT NULL)')
        conn.execute('BEGIN IMMEDIATE')
        existing = conn.execute('SELECT variant FROM demos WHERE content_key=? '
                                'ORDER BY variant', (key,)).fetchall()
        cycles = conn.execute('SELECT COUNT(*) FROM demo_cycles WHERE content_key=?',
                              (key,)).fetchone()[0]
        total = len(existing) + cycles
        if total < 3:
            variant = max((row[0] for row in existing), default=0) + 1
        else:
            variant = (total - 3) % 3 + 1
        html = build_html(mods, variant=variant)
        folder = Path(out_dir).resolve()
        folder.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=folder,
                                         prefix='.demo-', suffix='.tmp', delete=False) as f:
            pending = Path(f.name)
            f.write(html)
            f.flush()
            os.fsync(f.fileno())
        output = folder / ('投标演示页面_%s_方案%d_%s.html' %
                           (key[:12], variant, pending.stem[6:]))
        os.replace(pending, output)
        if total < 3:
            conn.execute('INSERT INTO demos VALUES (?, ?, ?)', (key, variant, str(output)))
        else:
            conn.execute('INSERT INTO demo_cycles(content_key, variant, output) '
                         'VALUES (?, ?, ?)', (key, variant, str(output)))
        conn.commit()
        committed = True
        return {'path': str(output), 'variant': variant, 'scheme': SCHEMES[variant - 1],
                'attempt': total + 1,
                'modules': len(mods), 'features': sum(len(fs) for _, fs in mods),
                'pages': html.count('<section class="page'),
                'navigation': html.count('class="kid"') + len(mods) + 3}
    finally:
        if not committed:
            conn.rollback()
        conn.close()
        if pending is not None and pending.exists():
            pending.unlink()
        if not committed and output is not None and output.exists():
            output.unlink()

# ---------------- 红线字样替换 ----------------
_FORBID = (('参数', '数据'), ('需求', '业务'), ('条款', '规则'), ('示例', ''),
           ('lorem', ''), ('Lorem', ''), ('xxx', ''), ('XXX', ''))


def sanitize(s):
    for a, b in _FORBID:
        s = s.replace(a, b)
    return s.strip() or '业务项'


def _esc(s):
    return (s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;') \
                    .replace('"', '&quot;')


# ---------------- 参数解析：模块 → 功能点 ----------------
_RE_MOD = re.compile(r'^(?:[一二三四五六七八九十]+[、\.]|[（(][一二三四五六七八九十]+[)）]|\d+[、)）])\s*(.+)$')
_RE_FP = re.compile(r'^(?:[（(]\d+[)）]|\d+\.\d+|\d+\.|[①-⑳]|[-·•*▲◆■●])\s*(.+)$')


def parse_params(text):
    mods = []
    cur = None
    for raw in (text or '').splitlines():
        line = raw.strip()
        if not line:
            continue
        m1 = _RE_MOD.match(line)
        m2 = _RE_FP.match(line)
        if m1 and len(m1.group(1)) <= 30:
            cur = [m1.group(1), []]
            mods.append(cur)
        elif m2:
            if cur is None:
                cur = ['业务管理', []]
                mods.append(cur)
            cur[1].append(m2.group(1))
        else:
            if cur is None:
                if len(line) <= 30 and not any(a in line for a, _ in _ACTS):
                    cur = [line, []]
                    mods.append(cur)
                    continue
                cur = ['业务管理', []]
                mods.append(cur)
            cur[1].append(line)
    for name, fps in mods:
        if not fps:
            fps.append(name)
    out = []
    for name, fps in mods:
        seen, kept = set(), []
        for f in fps:
            if f not in seen:
                seen.add(f)
                kept.append(f)
        out.append((name, kept))
    # 参数无模块标题（全部落入默认模块）且功能点较多时，自动聚类成多模块组
    if len(out) == 1 and out[0][0] == '业务管理' and len(out[0][1]) >= 6:
        cl = _auto_cluster(out[0][1])
        if cl:
            out = cl
    return out


# ---------------- 长句 → 短功能名（原文不上屏） ----------------
_ACTS = [('生成', '生成'), ('评估', '评估'), ('评价', '评价'), ('分析', '分析'),
         ('查询', '查询'), ('监控', '监控'), ('对比', '对比'), ('上传', '上传'),
         ('导入', '导入'), ('导出', '导出'), ('配置', '配置'), ('编排', '编排'),
         ('规划', '规划'), ('管理', '管理'), ('办理', '办理'), ('跟踪', '跟踪'),
         ('录入', '录入'), ('提醒', '提醒'), ('登记', '登记'), ('审核', '审核'),
         ('发布', '发布'), ('预警', '预警'), ('维护', '维护'), ('检索', '检索')]
_NOUNS = ['教案框架', '知识图谱', '知识点', '课程目标', '毕业要求', '教学环节',
          '考核手段', '评价方法', '时间分配', '教学互动', '课件', '课时', '成绩',
          '学习路径', '培养方案', '岗位能力', '课堂', '课程', '学生', '教师',
          '班级', '作业', '试卷', '测验', '讨论', '资源', '会话', '数据', '报表',
          '日志', '权限', '角色', '订单', '商品', '库存', '客户', '合同', '项目',
          '任务', '审批', '归档', '工单', '宿舍', '课表', '调停课']


@lru_cache(maxsize=512)
def _short_name(s):
    s = sanitize(s)
    s = re.sub(r'^[▲◆■●○◇△▼#*·\-—]+', '', s).strip()
    s = re.sub(r'^(支持|提供|实现|完成|包括|包含|能够|需要|要求|用于|针对|可)+', '', s).strip()
    if len(s) <= 12:
        return s or '业务项'
    # 优先保留分句中明确的“业务对象 + 动作”，不让后文的交付物抢走主名称。
    pairs = []
    for noun in _NOUNS:
        for verb, action in _ACTS:
            patterns = [re.escape(noun) + r'(?:的|自动|智能|批量)*' + re.escape(verb),
                        re.escape(verb) + r'(?:对应的|相应的|指定的|新的|的)?' + re.escape(noun)]
            for pattern in patterns:
                match = re.search(pattern, s)
                if match:
                    clause = len(re.findall(r'[，,；;。]', s[:match.start()]))
                    pairs.append((clause, 0 if len(noun) >= 4 else 1, match.start(), -len(noun), noun + action))
    if pairs:
        return min(pairs)[4]
    head = re.split(r'[，,；;。]', s)[0]
    scope = head if any(n in head for n in _NOUNS) else s
    nouns = sorted((n for n in _NOUNS if n in scope), key=lambda n: (scope.index(n), -len(n)))
    acts = sorted(((scope.index(k), a) for k, a in _ACTS if k in scope))
    if acts:
        name = ''.join(nouns[:1]) + acts[0][1]
    else:
        name = ''.join(nouns[:2])
    if not name:
        name = re.split(r'[，,；;、]', s)[0][:12]
    return name or '业务项'


def _short_names(fps):
    seen, out = {}, []
    for f in fps:
        n = _short_name(f)
        seen[n] = seen.get(n, 0) + 1
        out.append(n + ('·%d' % seen[n] if seen[n] > 1 else ''))
    return out


def _auto_cluster(fps):
    """参数无模块标题时：按短名前两字聚类成模块组，模块名取组内高频业务名词+管理。"""
    names = _short_names(fps)
    order, members = [], {}
    for f, n in zip(fps, names):
        key = n[:2]
        if key not in members:
            members[key] = []
            order.append(key)
        members[key].append((f, n))
    if len(order) < 2:
        return None
    mods = []
    for key in order:
        mem = members[key]
        cnt = {}
        for _f, n in mem:
            for noun in _NOUNS:
                if noun in n:
                    cnt[noun] = cnt.get(noun, 0) + 1
        if cnt:
            base = sorted(cnt.items(), key=lambda kv: (-kv[1], -len(kv[0])))[0][0]
        else:
            base = key
        mods.append((base + '管理', [f for f, _n in mem]))
    return mods


# ---------------- 模拟数据池 ----------------
_OWNERS = ['张伟', '李娜', '王强', '刘洋', '陈静', '杨帆', '赵磊', '孙悦', '周婷', '吴斌']
_DEPTS = ['综合一部', '综合二部', '运营中心', '服务中心', '稽核组', '调度中心']
_TIMES = ['2026-09-14 09:24', '2026-09-18 15:02', '2026-10-06 10:47',
          '2026-10-11 16:33', '2026-10-15 08:56']
_STATUS = [('进行中', 'tag-b'), ('已完成', 'tag-g'), ('待审核', 'tag-o')]
_TILES = ['t-b', 't-p', 't-g', 't-o', 't-c', 't-k']
_GHOST = ['▦', '', '✚', '◆', '❖', '▤']
_WORDS = ['智能', '协同', '闭环', '画像', '中台', '编排', '洞察', '溯源', '预警',
          '看板', '图谱', '调度', '留痕', '复核', '归档']

_ICO = [
    '<svg viewBox="0 0 24 24"><path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V21h14V9.5"/></svg>',
    '<svg viewBox="0 0 24 24"><rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 8h6M9 12h6M9 16h4"/></svg>',
    '<svg viewBox="0 0 24 24"><path d="M4 20V10M10 20V4M16 20v-8M22 20H2"/></svg>',
    '<svg viewBox="0 0 24 24"><path d="M21 12a8 8 0 1 1-3.2-6.4"/><path d="M21 4v5h-5"/></svg>',
    '<svg viewBox="0 0 24 24"><circle cx="6" cy="6" r="3"/><circle cx="18" cy="18" r="3"/><path d="M8.5 7.5 15.5 16.5"/></svg>',
    '<svg viewBox="0 0 24 24"><rect x="3" y="3" width="8" height="8" rx="2"/><rect x="13" y="3" width="8" height="8" rx="2"/><rect x="3" y="13" width="8" height="8" rx="2"/><rect x="13" y="13" width="8" height="8" rx="2"/></svg>',
    '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M19.1 4.9 17 7M7 17l-2.1 2.1"/></svg>',
    '<svg viewBox="0 0 24 24"><path d="M8 6h13M8 12h13M8 18h13"/><circle cx="4" cy="6" r="1"/><circle cx="4" cy="12" r="1"/><circle cx="4" cy="18" r="1"/></svg>',
]


# ---------------- 基础组件 ----------------
def _tile(cls, lb, vl, unit, sb, gh):
    if sb[:1] in ('↑', '↓'):
        sb = '<span class="%s">%s</span>' % ('tr-up' if sb[0] == '↑' else 'tr-dn', _esc(sb))
    else:
        sb = _esc(sb)
    return ('<div class="tile %s"><div class="lb">%s</div>'
            '<div class="vl">%s<i>%s</i></div><div class="sb">%s</div>'
            '<div class="gh">%s</div></div>' % (cls, _esc(lb), vl, _esc(unit), sb, gh))


def _tiles6(items):
    defs = [('在办业务', '12', '项', '本月新增 5 项'), ('待办事项', '5', '件', '今日到期 2 件'),
            ('按期办结率', '96', '%', '环比 +1.2%'), ('服务满意度', '98', '%', '环比 +0.4%'),
            ('累计办理', '1,860', '件', '近 30 天 342 件'), ('智能调用', '1,024', '次', '近 7 天 216 次')]
    items = (items or []) + defs[len(items):]
    return '<div class="tiles">%s</div>' % ''.join(
        _tile(_TILES[i % 6], lb, vl, unit, sb, _GHOST[i % 6])
        for i, (lb, vl, unit, sb) in enumerate(items[:6]))


def _card(title, inner, sub='', extra=''):
    return ('<div class="card"><div class="ch"><h3>%s</h3><div style="display:flex;'
            'align-items:center;gap:10px"><span class="sub">%s</span>%s</div></div>%s</div>'
            % (_esc(title), _esc(sub), extra, inner))


def _more():
    return '<span class="more" data-act="query">查看全部</span>'


def _bars(vals, labels):
    mx = max(vals) or 1
    cols = ''.join('<div class="c"><div class="v" style="height:%d%%"><b>%d</b></div>'
                   '<div class="lb">%s</div></div>'
                   % (int(v * 88 / mx), v, _esc(lb)) for v, lb in zip(vals, labels))
    return '<div class="bars">%s</div>' % cols


def _hbars(rows, unit='次'):
    mx = max(v for _l, v in rows) or 1
    return ''.join('<div class="hbar"><span>%s</span><div class="tr"><div class="fl" '
                   'style="width:%d%%"></div></div><b>%d %s</b></div>'
                   % (_esc(l), int(v * 100 / mx), v, unit) for l, v in rows)


def _ring(pct, label, color='#3370ff'):
    c = 2 * 3.14159 * 44
    return ('<div class="ring"><svg viewBox="0 0 100 100"><circle cx="50" cy="50" r="44" '
            'fill="none" stroke="#eef1f6" stroke-width="8"/><circle cx="50" cy="50" r="44" '
            'fill="none" stroke="%s" stroke-width="8" stroke-linecap="round" '
            'stroke-dasharray="%.1f %.1f" transform="rotate(-90 50 50)"/>'
            '<text x="50" y="47" text-anchor="middle" font-size="19" font-weight="700" '
            'fill="#1f2329">%d</text><text x="50" y="62" text-anchor="middle" font-size="9" '
            'fill="#86909c">%s</text></svg></div>' % (color, c * pct / 100, c, pct, _esc(label)))


def _radar(labels, vals):
    import math
    n = len(labels)
    pts, dots, txts = [], [], []
    for i, (lb, v) in enumerate(zip(labels, vals)):
        a = -math.pi / 2 + i * 2 * math.pi / n
        x, y = 50 + 34 * v / 100 * math.cos(a), 50 + 34 * v / 100 * math.sin(a)
        pts.append('%.1f,%.1f' % (x, y))
        dots.append('<circle cx="%.1f" cy="%.1f" r="2.2" fill="#3370ff"/>' % (x, y))
        lx, ly = 50 + 46 * math.cos(a), 50 + 46 * math.sin(a)
        txts.append('<text x="%.1f" y="%.1f" text-anchor="middle" font-size="6.5" '
                    'fill="#86909c">%s</text>' % (lx, ly + 2, _esc(lb)))
    grid = ''.join('<polygon points="%s" fill="none" stroke="#eef1f6" stroke-width="1"/>'
                   % ' '.join('%.1f,%.1f' % (50 + r * math.cos(-math.pi / 2 + i * 2 * math.pi / n),
                                             50 + r * math.sin(-math.pi / 2 + i * 2 * math.pi / n))
                              for i in range(n))
                   for r in (11, 22, 34))
    return ('<svg viewBox="-14 -8 128 116" style="width:100%%;max-width:360px;display:block;'
            'margin:0 auto">%s<polygon points="%s" fill="rgba(51,112,255,.16)" '
            'stroke="#3370ff" stroke-width="1.4"/>%s%s</svg>'
            % (grid, ' '.join(pts), ''.join(dots), ''.join(txts)))


def _table(headers, rows):
    th = ''.join('<th>%s</th>' % _esc(h) for h in headers)
    return '<div class="table-scroll"><table class="tbl"><tr>%s</tr>%s</table></div>' % (th, rows)


def _tag(st):
    for name, cls in _STATUS:
        if name == st:
            return '<span class="tag %s">%s</span>' % (cls, st)
    return '<span class="tag tag-b">%s</span>' % _esc(st)


def _pager():
    return ('<div class="pager"><button>‹</button><button class="on">1</button>'
            '<button>2</button><button>3</button><button>›</button></div>')


def _toolbar(fp, acts=('new', 'export', 'query')):
    btns = ''
    if 'new' in acts:
        btns += '<button class="btn btn-pri" data-act="new" data-name="%s">新建</button>' % _esc(fp)
    if 'export' in acts:
        btns += '<button class="btn btn-gh" data-act="export">导出</button>'
    if 'query' in acts:
        btns += '<button class="btn btn-gh" data-act="query">查询</button>'
    return ('<div class="tb-row"><div class="search sm"><svg viewBox="0 0 24 24">'
            '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>'
            '<input placeholder="搜索%s记录"></div>'
            '<div class="seg"><button class="sg on">全部</button><button class="sg">进行中</button>'
            '<button class="sg">已完成</button></div><span class="sp"></span>%s</div>'
            % (_esc(fp), btns))


def _biz_rows(fp, n=5):
    trs = ''
    for i in range(n):
        dept, st = _DEPTS[(i + len(fp)) % 6], _STATUS[(i + len(fp)) % 3]
        trs += ('<tr><td>GD2026-%04d</td><td>%s·%s专项</td><td>%s</td><td>%s</td><td>%s</td>'
                '<td>%s</td><td><button class="btn btn-gh btn-sm" data-act="view" data-name="%s" '
                'data-owner="%s" data-time="%s" data-status="%s">查看</button> '
                '<button class="btn btn-gh btn-sm" data-act="edit" data-name="%s">编辑</button>'
                '</td></tr>'
                % (147 + i + len(fp) * 7, _esc(dept), _esc(fp), _OWNERS[(i + len(fp)) % 10],
                   _DEPTS[(i + 2) % 6], _TIMES[(i + len(fp)) % 5], _tag(st[0]), _esc(fp),
                   _OWNERS[(i + len(fp)) % 10], _TIMES[(i + len(fp)) % 5], st[0], _esc(fp)))
    return trs


def _biz_table(fp, n=5):
    return _table(['编号', '业务名称', '责任人', '部门', '最近更新', '状态', '操作'],
                  _biz_rows(fp, n))


def _tl(items=None):
    items = items or [('数据接入完成', '2026-10-15 08:56 · 调度中心', True),
                      ('规则校验通过', '2026-10-15 09:40 · 稽核组', True),
                      ('人工复核中', '进行中 · 综合一部', False),
                      ('结论推送', '待启动 · 服务中心', False)]
    return '<div class="tl">%s</div>' % ''.join(
        '<div class="ti%s"><div class="dt"></div><div><b>%s</b>'
        '<div style="font-size:12px;color:#86909c">%s</div></div></div>'
        % (' done' if d else '', _esc(t), _esc(m)) for t, m, d in items)


def _note(items):
    return '<div class="kv">%s</div>' % ''.join(
        '<div class="kd"><div class="d"></div><div>%s</div></div>' % _esc(t) for t in items)


def _minis(pairs):
    return '<div class="minis">%s</div>' % ''.join(
        '<div><b>%s</b><span>%s</span></div>' % (v, _esc(l)) for l, v in pairs)


def _tail():
    return ('<div class="grid2">' +
            _card('近期动态', _tl(), extra=_more()) +
            _card('服务保障', _note(['双人复核与电子归档，全程留痕可追溯。',
                                    '时限前一日自动提醒，避免逾期风险。',
                                    '工作日 2 小时内响应业务咨询。',
                                    '数据加密存储，按部门权限隔离。'])) + '</div>')


# ---------------- 九种页面形态（高密度版） ----------------
def _pg_manage(fp, k):
    return (_tiles6([(fp + '在办', '12', '项', '本月新增 5 项'), ('待办', '3', '件', '今日到期 1 件'),
                     ('办结率', '96', '%', '环比 +1.2%')]) +
            '<div class="grid2" style="margin-top:16px">' +
            _card(fp + '· 业务办理', _toolbar(fp) + _biz_table(fp) + _pager(),
                  sub='全流程线上流转', extra=_more()) +
            _card('办理量趋势', _minis([('本月办理', '342'), ('环比', '+12%'), ('平均时长', '1.8 天')]) +
                  _bars([18, 26, 22, 31, 27, 35], ['5月', '6月', '7月', '8月', '9月', '10月']),
                  sub='近六个月') + '</div>' + _tail())


def _pg_chat(fp, k, source=None):
    title = fp if fp.endswith('助手') else fp + '助手'
    p = _profile(source or fp)
    outline = '<br>'.join('%d. %s：检查%s，形成%s。' %
                         (i + 1, _esc(stage), _esc(p[3][i]), _esc(p[5][i % 3]))
                         for i, stage in enumerate(p[4]))
    msgs = ('<div class="msg"><div class="av av-ai">AI</div><div class="bub">你好！我是%s。'
            '当前聚焦%s，可梳理执行步骤、质量检查和交付成果。</div></div>'
            '<div class="msg me"><div class="bub">如何推进%s并形成可复核的成果？</div>'
            '<div class="av av-me">我</div></div>'
            '<div class="msg"><div class="av av-ai">AI</div><div class="bub">可以按以下步骤组织：'
            '<br>%s<br>以上为当前业务的本地演练流程，可结合实际对象调整。</div></div>'
            % (_esc(title), _esc(p[1]), _esc(fp), outline))
    qc = ''.join('<button class="qc" data-act="ask" data-prompt="%s"><b>%s</b><span>%s</span></button>' %
                 (_esc(prompt), _esc(t), _esc(d)) for t, d, prompt in [
                     ('梳理执行步骤', p[4][0] + ' → ' + p[4][-1], fp + '有哪些执行步骤？'),
                     ('检查质量维度', ' / '.join(p[3][:2]), fp + '需要重点检查什么？'),
                     ('整理交付清单', ' / '.join(p[5][:2]), fp + '需要交付哪些成果？')])
    qc += '<button class="qc" data-act="exportchat"><b>导出会话记录</b><span>下载当前对话文本</span></button>'
    left = _card('智能对话',
                 '<div class="chat" data-chatbox>%s</div>'
                 '<div class="chatin"><input data-chatin placeholder="请输入你的问题，例如：%s如何办理？">'
                 '<button class="btn btn-pri" data-act="send">发送</button></div>' % (msgs, _esc(fp)))
    right = ('<div class="card"><div class="ch"><h3>快捷指令</h3></div>%s</div>'
             '<div class="card"><div class="ch"><h3>最近咨询</h3></div>%s</div>'
             % (qc, _hbars([(p[3][0], 26), (p[3][1], 18), (p[3][2], 9)])))
    faq = _card('协作检查清单', _note([
        '对象确认：先选择%s，明确本轮工作范围。' % p[2][0],
        '质量确认：逐项核对%s，记录未完成部分。' % '、'.join(p[3]),
        '交付确认：%s应与执行记录一一对应。' % p[5][0]]), sub=p[1] + '工作要点')
    tips = _card('使用技巧', _note([
        '提问时说明业务对象、当前进度与希望获得的成果。',
        '通过快捷指令分别梳理步骤、质量与交付物。',
        '会话可下载为文本，便于后续人工复核。',
        '当前为离线交互演练，不连接在线模型，刷新后会话复原。']))
    return ('<div class="grid2"><div>%s</div><div>%s</div></div>'
            '<div class="grid2">%s%s</div>' % (left, right, faq, tips))


def _pg_upload(fp, k):
    form = ('<div class="frm"><label>业务类别<span class="ctl"><select><option>%s</option>'
            '<option>专项核查</option><option>常规办理</option></select></span></label>'
            '<label>办理批次<span class="ctl"><select><option>2026 年第四季度</option>'
            '<option>2026 年第三季度</option></select></span></label>'
            '<label>上传材料<span class="ctl upzone" data-act="upload"><div class="ar">↑</div>'
            '点击或拖拽文件到此处<br><span style="font-size:12px">支持 .docx / .xlsx / .pdf，'
            '最大 50MB</span></span></label>'
            '<button class="btn btn-pri" style="width:100%%" data-act="gen">智能生成办理方案</button>'
            '</div>' % _esc(fp))
    preview = ('<div class="kv"><div class="r"><span class="k">方案名称：</span>%s·2026Q4 办理方案</div>'
               '<div class="r"><span class="k">适用环节：</span>受理、核验、复核、归档</div>'
               '<div class="r"><span class="k">关联材料：</span>申请表、佐证清单、核验记录</div>'
               '<div class="r"><span class="k">环节设计：</span></div>'
               '<div class="kd"><div class="d"></div><div><b>受理登记（0.5 天）</b><br>'
               '核对申请材料完整性，登记台账。</div></div>'
               '<div class="kd"><div class="d"></div><div><b>智能核验（1 天）</b><br>'
               '自动比对历史数据，标记疑点供人工复核。</div></div>'
               '<div class="kd"><div class="d"></div><div><b>复核归档（0.5 天）</b><br>'
               '双人复核后电子归档，全程留痕。</div></div></div>' % _esc(fp))
    recs = _table(['材料', '批次', '类型', '生成时间', '状态'],
                  '<tr><td>%s申请表.docx</td><td>2026Q4</td><td>办理方案</td><td>2026-10-14 10:23</td>'
                  '<td>%s</td></tr><tr><td>%s佐证清单.xlsx</td><td>2026Q4</td><td>材料识别</td>'
                  '<td>2026-10-14 10:20</td><td>%s</td></tr><tr><td>专项核查表.pdf</td><td>2026Q3</td>'
                  '<td>办理方案</td><td>2026-10-13 16:45</td><td>%s</td></tr>'
                  '<tr><td>核验记录表.docx</td><td>2026Q3</td><td>材料识别</td><td>2026-10-12 14:18</td>'
                  '<td>%s</td></tr>'
                  % (_esc(fp), _tag('已完成'), _esc(fp), _tag('已完成'),
                     _tag('已完成'), _tag('已完成')))
    return ('<div class="grid2">' + _card('选择类别与材料', form) +
            _card('方案预览', preview, extra='<button class="btn btn-pri btn-sm" data-act="copy">'
                  '复制内容</button><button class="btn btn-gh btn-sm" data-act="export">导出 Word</button>') +
            '</div>' + _card('最近生成记录', recs, extra=_more()) +
            '<div class="grid2">' +
            _card('识别质量', '<div class="ringwrap">%s</div>' %
                  ''.join(_ring(p, l) for p, l in [(96, '字段识别'), (92, '要素抽取'), (88, '匹配度')])) +
            _card('材料处理说明', _note([
                '上传后 10 秒内完成格式校验与内容识别。',
                '识别结果支持人工修正并回流优化模型。',
                '方案生成后可一键导出 Word 或复制全文。'])) + '</div>')


def _pg_analysis(fp, k):
    form = ('<div class="frm"><label>评估对象<span class="ctl"><select><option>%s</option>'
            '<option>综合一部专项</option><option>运营中心专项</option></select></span></label>'
            '<label>评估周期<span class="ctl"><select><option>2026 年度</option>'
            '<option>2026 下半年</option></select></span></label>'
            '<button class="btn btn-pri" style="width:100%%" data-act="gen">开始评估</button></div>'
            '<div class="tile t-b" style="margin-top:14px"><div class="lb">综合得分</div>'
            '<div class="vl">87.3<i>分</i></div></div>' % _esc(fp))
    dims = ''.join('<div class="card dim"><div class="ch"><h3>%s</h3>%s</div><p>%s</p></div>'
                   % (t, _tag(s), d) for t, s, d in [
                       ('完整度分析', '已完成', '目标、流程、材料、归档环节完整覆盖。'),
                       ('可测量性', '进行中', '关键指标可量化程度较高，个别目标需量化表述。'),
                       ('支撑均衡性', '待审核', '部分指标支撑集中，建议分散强化。')])
    acts = _table(['行动项', '责任对象', '优先级', '预计完成'],
                  '<tr><td>完善关键指标可测量性</td><td>综合一部</td><td>%s</td><td>2026-11-15</td></tr>'
                  '<tr><td>调整支撑分布</td><td>运营中心</td><td>%s</td><td>2026-12-01</td></tr>'
                  '<tr><td>优化归档时限</td><td>服务中心</td><td>%s</td><td>2026-12-15</td></tr>'
                  % (_tag('待审核'), _tag('待审核'), _tag('进行中')))
    summary = _card('评估报告摘要',
                    '<div class="kv"><div class="r">该方案在目标契合度、流程完整性方面表现良好。'
                    '建议关注：1）强化关键指标的可测量性；2）优化支撑强度分布；'
                    '3）调整后期环节的工作量分布，避免前紧后松。</div></div>',
                    extra='<button class="btn btn-gh btn-sm" data-act="export">导出报告</button>')
    return ('<div class="grid2">' + _card('评估配置', form) +
            _card('维度评分雷达', _radar(['完整度', '契合度', '可测量', '前瞻性', '均衡性',
                                          '支撑强度', '分布', '达成度'],
                                         [92, 85, 82, 76, 74, 88, 80, 90]), sub='八维智能评估') +
            '</div><div class="grid3">%s</div>%s' % (dims, summary) +
            _card('优化行动项', acts, sub='基于评估结果自动推荐') + _tail())


def _pg_path(fp, k):
    steps = ''.join('<div class="st"><div class="n">%d</div><div><div class="tt">%s</div>'
                    '<div class="mt">预计 %d 分钟 · %s</div><div class="ds">%s</div></div></div>'
                    % (i + 1, t, m, lv, d) for i, (t, m, lv, d) in enumerate([
                        ('受理与登记', 30, '基础环节', '核对材料完整性并登记台账。'),
                        ('智能核验', 60, '核心环节', '自动比对历史数据，标记疑点。'),
                        ('人工复核', 45, '核心环节', '双人复核疑点与结论。'),
                        ('结论出具', 30, '综合应用', '生成结论并推送相关方。'),
                        ('电子归档', 20, '收尾环节', '全流程材料归档留痕。')]))
    tpl = ''.join('<div class="qc"><b>%s</b><span>%d 个环节 · 预计 %.1f 小时 · %d 人使用</span></div>'
                  % (t, n, h, u) for t, n, h, u in [('标准办理路径', 5, 3.1, 128),
                                                   ('加急办理路径', 3, 1.5, 96),
                                                   ('专项核查路径', 7, 5.0, 64)])
    nodes = _table(['节点', '时间', '责任', '状态'],
                   '<tr><td>路径评审</td><td>2026-10-20</td><td>综合一部</td><td>%s</td></tr>'
                   '<tr><td>试点运行</td><td>2026-11-01</td><td>运营中心</td><td>%s</td></tr>'
                   '<tr><td>全面推广</td><td>2026-11-15</td><td>服务中心</td><td>%s</td></tr>'
                   % (_tag('进行中'), _tag('待审核'), _tag('待审核')))
    row1 = ('<div class="grid2">' +
            _card('路径编排', '<div class="steps">%s</div>' % steps, sub='拖拽可调整环节顺序',
                  extra='<button class="btn btn-gh btn-sm" data-act="reset">重置</button>'
                        '<button class="btn btn-pri btn-sm" data-act="save">保存路径</button>') +
            '<div>%s%s</div></div>'
            % (_card('目标与智能编排',
                     '<div class="frm"><label>业务方向<span class="ctl"><select><option>%s</option>'
                     '<option>专项核查能力提升</option><option>综合服务能力提升</option>'
                     '</select></span></label><label>目标描述<span class="ctl"><textarea rows="3">'
                     '掌握从受理到归档的完整办理方法，能够独立完成核验与复核。</textarea></span></label>'
                     '<button class="btn btn-pri" style="width:100%%" data-act="gen">AI 生成推荐路径</button>'
                     '</div>' % _esc(fp)),
               _card('路径模板库', tpl, extra='<button class="btn btn-gh btn-sm" data-act="edit" '
                     'data-name="模板">管理模板</button>')))
    return (row1 +
            '<div class="grid2">' +
            _card('环节时长分布', _bars([30, 60, 45, 30, 20], ['受理', '核验', '复核', '出具', '归档']),
                  sub='单位：分钟') +
            _card('建设节点', nodes) + '</div>' + _tail())


def _pg_entry(fp, k):
    rows = ''
    for i in range(5):
        st = '已完成' if i < 3 else '待提交'
        rows += ('<tr><td>2026%04d</td><td>%s</td><td><input class="cell" value="%s"></td>'
                 '<td><input class="cell" value="%s"></td><td>%s</td></tr>'
                 % (101 + i, _OWNERS[i], '' if i >= 3 else str(88 - i * 6),
                    '' if i >= 3 else str(91 - i * 5), _tag(st)))
    remind = ('<div class="chips">%s</div><p style="font-size:12px;color:#86909c;margin-top:10px">'
              '以上人员尚未提交，已发送提醒通知。</p>'
              % ''.join('<span>%s · 2026%04d</span>' % (_OWNERS[i], 103 + i) for i in (3, 4)))
    return ('<div class="grid2">' +
            _card('数据录入', _toolbar(fp, acts=('export',)) +
                  _table(['编号', '名称', '本期数', '核验数', '提交状态'], rows) +
                  '<div style="margin-top:12px;display:flex;gap:10px">'
                  '<button class="btn btn-gh btn-sm" data-act="tpl">下载模板</button>'
                  '<button class="btn btn-gh btn-sm" data-act="upload">导入数据</button>'
                  '<button class="btn btn-pri btn-sm" data-act="save">提交数据</button></div>',
                  sub='系统自动记录每次提交') +
            _card('达成与分布', '<div class="ringwrap">%s</div>' %
                  ''.join(_ring(p, l) for p, l in [(86, '目标 1'), (82, '目标 2'),
                                                   (78, '目标 3'), (88, '目标 4')]) +
                  '<div style="margin-top:14px">' +
                  _bars([6, 14, 22, 18, 9], ['<60', '60-69', '70-79', '80-89', '90+']) + '</div>',
                  sub='自动计算') + '</div>' +
            '<div class="grid2">' +
            _card('提交记录', _table(['提交时间', '记录数', '状态'],
                  '<tr><td>2026-10-14 14:32</td><td>42</td><td>%s</td></tr>'
                  '<tr><td>2026-10-13 16:05</td><td>40</td><td>%s</td></tr>'
                  '<tr><td>2026-10-12 09:18</td><td>38</td><td>%s</td></tr>'
                  % (_tag('已完成'), _tag('已完成'), _tag('进行中'))),
                  extra='<button class="btn btn-gh btn-sm" data-act="calc">重新计算</button>') +
            _card('未提交提醒', remind) + '</div>')


def _pg_compare(fp, k):
    delta = _table(['指标', '上一版', '当前版', '变化'],
                   '<tr><td>办理环节数量</td><td>6</td><td>8</td><td style="color:#00b96b">+2</td></tr>'
                   '<tr><td>材料清单数量</td><td>12</td><td>12</td><td>0</td></tr>'
                   '<tr><td>平均办理时长</td><td>5.2 天</td><td>3.8 天</td><td style="color:#00b96b">-1.4</td></tr>'
                   '<tr><td>线上办理占比</td><td>76%</td><td>92%</td><td style="color:#00b96b">+16</td></tr>'
                   '<tr><td>一次办结率</td><td>81%</td><td>89%</td><td style="color:#00b96b">+8</td></tr>')
    wc = ''.join('<span style="font-size:%dpx;opacity:%.2f">%s</span>'
                 % (s, o, w) for w, s, o in zip(_WORDS, [22, 16, 26, 14, 20, 15, 24, 13, 18, 16],
                                                [.95, .7, 1, .55, .85, .6, .9, .5, .75, .65]))
    diff = ('<div class="diff"><div class="dr"><div class="ic ic-add">+</div><div><b>新增智能核验环节</b>'
            '<br>当前版引入自动比对，减少人工核验工作量。</div></div>'
            '<div class="dr"><div class="ic ic-chg">~</div><div><b>复核时限前移</b><br>'
            '复核由第 5 环节前移至第 4 环节，缩短整体时长。</div></div>'
            '<div class="dr"><div class="ic ic-del">−</div><div><b>精简重复材料</b><br>'
            '合并两份重复佐证清单，避免重复提交。</div></div></div>')
    move = _table(['事项', '性质', '上一版', '当前版', '说明'],
                  '<tr><td>智能核验</td><td>必办</td><td>第 5 环节</td><td>第 4 环节</td><td>提前 1 环节</td></tr>'
                  '<tr><td>材料补正</td><td>选办</td><td>第 2 环节</td><td>第 3 环节</td><td>延后 1 环节</td></tr>'
                  '<tr><td>电子归档</td><td>必办</td><td>—</td><td>第 8 环节</td><td>新增环节</td></tr>')
    return ('<div class="grid2">' +
            _card('对比配置', '<div class="frm"><label>基准版本<span class="ctl"><select>'
                  '<option>2025 版 %s规程</option><option>2024 版 %s规程</option></select></span></label>'
                  '<label>对比版本<span class="ctl"><select><option>2026 版 %s规程</option>'
                  '<option>2025 版 %s规程</option></select></span></label>'
                  '<button class="btn btn-pri" style="width:100%%" data-act="gen">开始对比</button></div>'
                  % (_esc(fp), _esc(fp), _esc(fp), _esc(fp))) +
            _card('关键指标变化', delta) + '</div><div class="grid2">' +
            _card('高频关键词', '<div class="wc">%s</div>' % wc) +
            _card('差异分析与优化建议', diff) + '</div>' +
            _card('环节迁移明细', move) + _tail())


def _pg_monitor(fp, k):
    return (_tiles6([(fp + '· 在线节点', '24', '个', '异常 0 个'), ('今日事件', '136', '条',
                     '预警 2 条'), ('平均响应', '1.8', '秒', 'P95 3.2 秒')]) +
            '<div class="grid2" style="margin-top:16px">' +
            _card('实时趋势', _minis([('今日事件', '136'), ('预警', '2'), ('已处置', '131')]) +
                  _bars([42, 66, 50, 82, 58, 92, 70],
                        ['周一', '周二', '周三', '周四', '周五', '周六', '周日']), sub='近 7 日事件量') +
            _card('办理节点时间线', _tl(), extra='<span class="stt"><span class="d"></span>运行正常</span>') +
            '</div>' +
            _card('事件明细', _toolbar(fp, acts=('export', 'query')) +
                  _table(['时间', '节点', '事件', '级别', '状态'],
                         '<tr><td>2026-10-15 09:40</td><td>规则校验</td><td>批量校验通过</td>'
                         '<td>%s</td><td>%s</td></tr><tr><td>2026-10-15 08:56</td><td>数据接入</td>'
                         '<td>增量同步完成</td><td>%s</td><td>%s</td></tr>'
                         '<tr><td>2026-10-14 22:13</td><td>预警中心</td><td>超时风险提醒</td>'
                         '<td>%s</td><td>%s</td></tr><tr><td>2026-10-14 18:02</td><td>人工复核</td>'
                         '<td>复核结论回填</td><td>%s</td><td>%s</td></tr>'
                         % (_tag('已完成'), _tag('已完成'), _tag('已完成'), _tag('已完成'),
                            _tag('待审核'), _tag('进行中'), _tag('已完成'), _tag('已完成'))) +
                  _pager()))


def _pg_board(fp, k):
    return (_tiles6([(fp + '· 总量', '1,860', '件', '近 30 天 342 件'), ('在办', '12', '项',
                     '本周新增 5 项'), ('办结率', '96', '%', '环比 +1.2%')]) +
            '<div class="grid2" style="margin-top:16px">' +
            _card('办理量趋势', _bars([120, 168, 142, 190, 176, 210],
                                      ['5月', '6月', '7月', '8月', '9月', '10月']), sub='近六个月') +
            _card('分组活跃度', _hbars([(fp, 46), ('综合一部', 37), ('运营中心', 28),
                                        ('服务中心', 19)])) + '</div><div class="grid2">' +
            _card('TOP 事项', _hbars([('材料核验', 58), ('复核归档', 44), ('受理登记', 39),
                                      ('结论出具', 27), ('预警处置', 16)]), sub='按办理量排序') +
            _card('待办清单', _table(['事项', '责任', '截止', '状态'],
                  '<tr><td>%s复核</td><td>%s</td><td>2026-10-17</td><td>%s</td></tr>'
                  '<tr><td>材料补正</td><td>%s</td><td>2026-10-18</td><td>%s</td></tr>'
                  '<tr><td>归档抽查</td><td>%s</td><td>2026-10-20</td><td>%s</td></tr>'
                  % (_esc(fp), _OWNERS[0], _tag('进行中'), _OWNERS[1], _tag('已完成'),
                     _OWNERS[2], _tag('待审核'))), extra=_more()) + '</div>' + _tail())


_ARCH = [
    (('助手', '咨询', '问答', '对话', '客服', '答疑', '陪伴'), _pg_chat),
    (('上传', '生成', '识别', '导入', '课件'), _pg_upload),
    (('分析', '评估', '评价', '达成', '画像', '诊断', '匹配'), _pg_analysis),
    (('路径', '规划', '编排', '流程', '计划', '配置'), _pg_path),
    (('录入', '填报', '提交', '成绩', '登记', '采集'), _pg_entry),
    (('对比', '比较', '版本', '变更', '迁移'), _pg_compare),
    (('监控', '跟踪', '进度', '预警', '巡检', '日志'), _pg_monitor),
    (('概览', '统计', '洞察', '看板', '报表', '大屏'), _pg_board),
]


def _pick_arch(fp, idx):
    for keys, fn in _ARCH:
        if any(k in fp for k in keys):
            return fn
    return [_pg_manage, _pg_board, _pg_monitor, _pg_entry][idx % 4]


# 导航容量硬规则：大导航（模块）最多 10 条，小导航（功能项）最多 30 条；
# 超出时把多个功能合并到同一页/同一页签，内容全部保留不丢弃。
MAX_MODS = 10
MAX_KIDS = 30


def _fit_nav(mods):
    """[(模块名, [功能点])] → [(模块名, [([功能点], 组名)])]，满足导航容量规则。"""
    if len(mods) > MAX_MODS:
        per = -(-len(mods) // MAX_MODS)
        merged = []
        for i in range(0, len(mods), per):
            chunk = mods[i:i + per]
            if len(chunk) == 1:
                merged.append((chunk[0][0], list(chunk[0][1])))
            else:
                name = min((n for n, _ in chunk), key=lambda s: (len(s), s))
                merged.append((name, [f for _n, fs in chunk for f in fs]))
        mods = merged
    total = sum(len(fs) for _, fs in mods)
    groups = []
    for name, fps in mods:
        if total <= MAX_KIDS:
            groups.append((name, [([f], _short_name(f)) for f in fps]))
            continue
        target = max(1, min(len(fps), round(len(fps) * MAX_KIDS / total)))
        size = -(-len(fps) // target)
        gs = []
        for i in range(0, len(fps), size):
            chunk = fps[i:i + size]
            gs.append((chunk, _short_name(chunk[0])))
        groups.append((name, gs))
    while sum(len(gs) for _, gs in groups) > MAX_KIDS:
        mi = max(range(len(groups)), key=lambda j: len(groups[j][1]))
        nm, gs = groups[mi]
        if len(gs) < 2:
            break
        a, b = gs.pop(), gs[-1]
        gs[-1] = (a[0] + b[0], b[1])
    return groups


def _ledger(fp):
    return _card('办理记录', _toolbar(fp, acts=('export', 'query')) + _biz_table(fp, 6) + _pager(),
                 sub='全程留痕，支持追溯', extra=_more())


def _extra_rot(fp, k):
    r = k % 4
    if r == 0:
        return _card('质量指标', '<div class="ringwrap">%s</div>' %
                     ''.join(_ring(p, l) for p, l in [(96, '按期率'), (92, '规范率'),
                                                      (88, '满意度')]))
    if r == 1:
        return _card('构成分布', _hbars([('重点事项', 38), ('常规事项', 26), ('临时事项', 14)]),
                     sub='按事项性质')
    if r == 2:
        return _card('常用操作', ''.join(
            '<div class="qc" data-act="%s"><b>%s</b><span>%s</span></div>' % (a, t, d)
            for a, t, d in [('export', '导出台账', '生成当前筛选条件的导出任务'),
                            ('calc', '重新计算', '刷新指标与分布数据'),
                            ('tpl', '下载模板', '获取标准模板用于线下填报')]))
    return _card('业务说明', _note([
        '办理遵循现行业务规则，关键步骤需双人确认。',
        '材料与结论自动电子归档，支持全链路追溯。',
        '异常自动预警并指派专人跟进。']) +
        '<div class="chips"><span>全程留痕</span><span>双人复核</span>'
        '<span>时限提醒</span><span>电子归档</span></div>')


def _pg_records(fp, k):
    return (_card('办理记录台账', _toolbar(fp, acts=('export', 'query')) + _biz_table(fp, 6) + _pager(),
                  sub='全程留痕，支持审计', extra=_more()) +
            '<div class="grid2">' +
            _card('办理流程', _tl(), extra=_more()) +
            _card('本月统计', _minis([('已办理', '342'), ('环比', '+12%'), ('平均时长', '1.8 天')]) +
                  _bars([18, 26, 22, 31, 27, 35], ['5月', '6月', '7月', '8月', '9月', '10月'])) +
            '</div><div class="grid2">' +
            _card('待办提醒', '<div class="chips">%s</div>'
                  '<p style="font-size:12px;color:#86909c;margin-top:10px">以上事项即将到期，已发送提醒。</p>'
                  % ''.join('<span>%s · 2026-10-%02d</span>' % (_OWNERS[j], 17 + j) for j in range(4))) +
            _card('归档说明', _note([
                '结论与材料 T+1 自动电子归档。',
                '归档编号与台账关联，一键追溯。',
                '保存期限遵循现行业务规则。']) +
                '<div class="chips"><span>电子归档</span><span>关联追溯</span>'
                '<span>期限合规</span></div>') + '</div>')


def _pg_stats(fp, k):
    return (_tiles6([(fp + '·已办理', '342', '件', '环比 +12%'), ('在办', '12', '项', '今日到期 3 项'),
                     ('办结率', '96', '%', '环比 +1.2%'), ('平均时长', '1.8', '天', '环比 -0.3 天'),
                     ('预警', '2', '条', '均已处置'), ('满意度', '98', '%', '环比 +0.4%')]) +
            '<div class="grid2" style="margin-top:16px">' +
            _card('办理量趋势', _bars([18, 26, 22, 31, 27, 35],
                                      ['5月', '6月', '7月', '8月', '9月', '10月']), sub='近六个月') +
            _card('构成分布', _hbars([('重点事项', 38), ('常规事项', 26), ('临时事项', 14)]),
                  sub='按事项性质') + '</div><div class="grid2">' +
            _card('质量指标', '<div class="ringwrap">%s</div>' %
                  ''.join(_ring(p, l) for p, l in [(96, '按期率'), (92, '规范率'), (88, '满意度')])) +
            _card('分析结论', _note([
                '办理量整体上升，10 月为峰值，需提前配置资源。',
                '按期率持续高于 95%，办理质量稳定。',
                '临时事项占比偏高，建议纳入计划管理。']),
                  extra='<button class="btn btn-gh btn-sm" data-act="export">导出报表</button>') +
            '</div>' + _tail())


# 业务语料仅使用短标签；导航由实际能力信号组合，不按固定倍数补菜单。
_PROFILES = [
    (('教案', '教学环节', '课件'), '教学设计',
     ('项目式教学导入', '核心概念讲解', '分组任务实践', '课堂讨论反馈', '课后巩固练习', '综合作品评价'),
     ('目标对齐', '时间分配', '活动参与', '评价覆盖'),
     ('选择课时', '编排教学环节', '配置评价方式', '发布教案'),
     ('课时教案', '课堂活动清单', '教学评价量规')),
    (('图谱', '知识点', '关联关系'), '知识建设',
     ('基础概念单元', '方法应用单元', '综合实践单元', '前置知识单元', '拓展阅读单元', '能力训练单元'),
     ('节点完整', '关系有效', '资源挂接', '路径连通'),
     ('抽取知识节点', '建立先修关系', '绑定学习资源', '发布知识网络'),
     ('知识节点清单', '关联关系矩阵', '学习路径图')),
    (('毕业要求', '培养方案'), '培养评价',
     ('工程知识应用', '问题分析能力', '方案设计能力', '团队协作能力', '沟通表达能力', '持续学习能力'),
     ('指标支撑', '考核覆盖', '目标达成', '持续改进'),
     ('拆解指标点', '映射支撑课程', '核验评价证据', '形成改进报告'),
     ('指标点矩阵', '课程支撑清单', '达成评价报告')),
    (('岗位', '能力画像'), '岗位能力',
     ('业务分析岗位', '产品设计岗位', '技术研发岗位', '质量测试岗位', '运营服务岗位', '项目协同岗位'),
     ('技能匹配', '实践经验', '协作表现', '成长潜力'),
     ('维护岗位标准', '映射能力要素', '识别能力差距', '制定提升计划'),
     ('岗位能力画像', '技能差距清单', '能力提升路径')),
    (('课程', '学生', '成绩', '教学', '课表', '作业'), '课程教学',
     ('课程基础目标', '应用实践目标', '综合设计目标', '自主学习目标', '团队项目目标', '学习反思目标'),
     ('目标明确', '考核有效', '成果达成', '反馈闭环'),
     ('选择课程班级', '维护课程目标', '关联考核任务', '汇总达成结果'),
     ('课程目标表', '考核支撑矩阵', '教学改进记录')),
    (('订单', '库存', '商品', '采购', '物流', '销售'), '经营管理',
     ('日常补货批次', '重点商品批次', '促销备货批次', '客户交付批次', '售后回访批次', '盘点核对批次'),
     ('库存准确', '交付及时', '质检合格', '客户满意'),
     ('核对业务单据', '分配可用库存', '跟进交付节点', '核销对账归档'),
     ('订单明细单', '库存流水表', '交付核对单')),
    (('工单', '报修', '巡检', '设备', '宿舍'), '运维服务',
     ('公共区域巡检', '设备故障排查', '重点设施检修', '服务回访任务', '安全隐患整改', '备件补充任务'),
     ('响应及时', '修复完成', '复检合格', '服务评价'),
     ('登记故障位置', '指派处理人员', '提交处理结果', '验收回访归档'),
     ('服务工单', '检修记录表', '回访评价单')),
]
_DEFAULT_PROFILE = ((), '业务协同', ('资料核对', '任务分派', '执行跟进', '结果复核', '质量抽查', '归档回访'),
                    ('内容完整', '执行及时', '质量合格', '过程可追溯'),
                    ('建立任务', '分配责任', '跟进进展', '确认成果'),
                    ('任务清单', '执行记录', '成果报告'))


def _profile(fp):
    primary = _short_name(fp)
    return next((p for p in _PROFILES if any(k in primary for k in p[0])),
                next((p for p in _PROFILES if any(k in fp for k in p[0])), _DEFAULT_PROFILE))


def _feature_views(fp):
    views = [('workspace', '业务工作台'), ('records', '成果清单')]
    rules = [
        (('上传', '导入', '识别', '材料', '课件', '资源'), 'materials', '材料资源库'),
        (('图谱', '知识点', '关联', '支撑', '映射', '匹配'), 'relations', '关联网络'),
        (('生成', '配置', '编排', '规划', '路径', '规则', '目标'), 'design', '设计配置'),
        (('评估', '分析', '统计', '评价', '达成', '看板', '报表'), 'analysis', '评估洞察'),
        (('对比', '版本', '变更', '历史'), 'versions', '版本对比'),
        (('审核', '审批', '监控', '预警', '校验', '质量', '巡检'), 'quality', '质量核验'),
        (('导出', '归档', '发布'), 'archive', '发布归档'),
    ]
    for keys, code, title in rules:
        if any(k in fp for k in keys):
            views.append((code, title))
    if len(views) == 2:
        views.append(('progress', '执行跟踪'))
    return views


def _seed(text):
    return int(content_key(text)[:8], 16)


def _scenario_rows(fp, name, variant):
    p = _profile(fp)
    seed = _seed(fp)
    rows = []
    for j, item in enumerate(p[2]):
        status = ('已完成', '进行中', '待审核')[(j + seed) % 3]
        rows.append('<tr><td>%02d</td><td><b>%s</b><small>%s · %s</small></td>'
                    '<td>%s</td><td><span class="progress-track"><i style="width:%d%%"></i></span></td>'
                    '<td>%s</td><td><button class="btn btn-gh btn-sm" data-act="view" '
                    'data-name="%s" data-owner="%s" data-time="2026-10-16" '
                    'data-status="%s">详情</button> <button class="btn btn-gh btn-sm" '
                    'data-act="edit" data-name="%s" data-owner="%s">编辑</button></td></tr>' %
                    (j + 1, _esc(item), _esc(name), _esc(p[5][j % 3]), _OWNERS[j],
                     100 if status == '已完成' else 42 + (seed + j * 7) % 48,
                     _tag(status), _esc(item), _OWNERS[j], status, _esc(item), _OWNERS[j]))
    return _table(['序号', '业务对象 / 交付成果', '责任人', '进度', '状态', '操作'], ''.join(rows))


def _relation_card(fp, name):
    p = _profile(fp)
    points = [(90, 55), (410, 55), (90, 165), (410, 165), (90, 275), (410, 275)]
    links, nodes = [], []
    for (x, y), label in zip(points, p[2]):
        links.append('<path d="M250 165 Q250 %d %d %d"/>' % (y, x, y))
        nodes.append('<g><rect x="%d" y="%d" width="142" height="36" rx="9"/>'
                     '<text x="%d" y="%d" text-anchor="middle">%s</text></g>' %
                     (x - 71, y - 18, x, y + 5, _esc(label)))
    graph = ('<svg class="relation-map" viewBox="0 0 500 330" role="img" aria-label="业务关联网络">'
             '<g class="edges">%s</g>%s<g class="hub"><circle cx="250" cy="165" r="48"/>'
             '<text x="250" y="169" text-anchor="middle">%s</text></g></svg>' %
             (''.join(links), ''.join(nodes), _esc(p[1])))
    return _card(name + ' · 关联全景', graph, sub='业务对象与成果之间的连接')


def _context_cards(fp, name, variant):
    p, seed = _profile(fp), _seed(fp)
    tasks = ''.join('<div class="action-line"><span class="step-dot">%02d</span><div><b>%s</b>'
                    '<small>%s · %s</small></div><button class="btn btn-gh btn-sm" '
                    'data-act="edit" data-name="%s">跟进</button></div>' %
                    (i + 1, _esc(stage), _OWNERS[i], _esc(p[5][i % 3]), _esc(stage))
                    for i, stage in enumerate(p[4]))
    quality = ''.join('<tr><td>%s</td><td>%d%%</td><td>%s</td></tr>' %
                      (_esc(label), 83 + (seed + i * 3) % 17, _tag('已完成' if i < 3 else '待审核'))
                      for i, label in enumerate(p[3]))
    artifacts = ''.join('<div class="resource-tile"><span class="file-icon">%s</span>'
                        '<div><b>%s</b><small>%s · 修订 %d 次</small></div>'
                        '<button class="btn btn-gh btn-sm" data-act="export">导出</button></div>' %
                        (('DOC', 'XLS', 'PDF')[i], _esc(item), _esc(name), 2 + i)
                        for i, item in enumerate(p[5]))
    advice = _note(['围绕%s检查%s，补齐未关联对象。' % (name, p[3][0]),
                    '优先复核%s，确认对应的%s。' % (p[2][seed % 6], p[5][1]),
                    '将%s的变更同步给责任人，保留确认记录。' % p[5][0],
                    '下一轮重点跟踪%s，汇总到%s。' % (p[3][2], p[5][2])])
    cards = [_card('待办行动', tasks, sub=p[1] + '执行路径'),
             _card('质量检查矩阵', _table(['检查维度', '完成度', '核验状态'], quality),
                   extra='<button class="btn btn-gh btn-sm" data-act="calc">重新核验</button>'),
             _card('成果与交付', artifacts, sub='可导出当前页面数据'),
             _card('业务建议', advice, sub='按当前业务对象组织')]
    if variant == 2:
        cards = [cards[2], cards[1], cards[0], cards[3]]
    elif variant == 3:
        cards = [cards[0], cards[3], cards[1], cards[2]]
    return '<div class="context-grid">%s</div>' % ''.join(cards)


GC = ['#3370ff', '#00b96b', '#ff9f00', '#7b61ff', '#14c0ff', '#f53f3f']


def _ws_kind(fp):
    """根据功能语义选择专属版式，让不同功能页长得不一样。"""
    rules = [
        ('graph', ('图谱', '知识点', '关联', '网络', '关系', '支撑', '映射', '矩阵')),
        ('live', ('课堂', '直播', '互动', '授课', '在线', '演示', '大屏', '会议')),
        ('chat', ('助手', '问答', '咨询', '客服', '对话', '智能体', '答疑', '陪伴')),
        ('gen', ('备课', '生成', '识别', '导入', '上传', '课件', '教案', '方案', '编排', '规划')),
        ('board', ('作业', '考试', '试卷', '出题', '答题', '错题', '考题', '练习', '组卷', '批阅', '清单', '台账')),
        ('dash', ('概览', '统计', '看板', '分析', '评估', '评价', '监控', '洞察', '报表', '诊断', '画像')),
    ]
    for kind, keys in rules:
        if any(k in fp for k in keys):
            return kind
    return 'manage'


def _chip(text, color='b'):
    return '<span class="wchip wchip-%s">%s</span>' % (color, _esc(text))


def _ws_graph(fp, name, variant):
    """知识图谱画布＋节点详情面板（对标参考图：中心辐射＋右侧详情）。"""
    import math
    p, seed = _profile(fp), _seed(fp)
    pos = [(50 + 33 * math.cos(-math.pi / 2 + i * math.pi / 3),
            50 + 33 * math.sin(-math.pi / 2 + i * math.pi / 3)) for i in range(6)]
    edges = ''.join('<line x1="50" y1="50" x2="%.1f" y2="%.1f"/>' % (x, y) for x, y in pos)
    dots = ''.join('<button class="gnode" style="left:%.1f%%;top:%.1f%%;--c:%s"><b>%s</b>'
                   '<small>%d 个节点</small></button>' %
                   (pos[i][0], pos[i][1], GC[i], _esc(n), 4 + (seed + i * 3) % 6)
                   for i, n in enumerate(p[2]))
    canvas = ('<div class="gcanvas"><svg class="gedges" viewBox="0 0 100 100" '
              'preserveAspectRatio="none">%s</svg>'
              '<div class="ghub"><b>%s</b><small>%s</small></div>%s</div>' %
              (edges, _esc(p[1]), _esc(name), dots))
    focus = p[2][seed % 6]
    detail = ('<div class="gdetail"><span class="eyebrow">一级节点 · 难度 %d/5</span>'
              '<h3>%s</h3><small>%s / %s</small>'
              '<div class="dl">知识说明</div><p>掌握「%s」的核心概念、典型实现与适用场景，'
              '能够在实际问题中选择恰当的组织方式。</p>'
              '<div class="dl">掌握度</div>'
              '<div class="mbar"><i style="width:%d%%"></i></div><em>%d%% · 年级平均 %d%%</em>'
              '<div class="gstat-grid">%s</div>'
              '<div class="dl">学习情况</div>%s'
              '<div class="gbtns"><button class="btn btn-pri btn-sm" data-act="edit" '
              'data-name="%s">编辑节点</button><button class="btn btn-gh btn-sm" '
              'data-act="ask" data-prompt="%s">AI 讲解</button></div></div>' % (
                  3 + seed % 3, _esc(focus), _esc(name), _esc(p[1]), _esc(focus),
                  58 + seed % 30, 58 + seed % 30, 70 + seed % 12,
                  ''.join('<div><span>%s</span><b>%s</b></div>' % (lb, v) for lb, v in
                          [('关联资源', '%d' % (120 + seed % 60)), ('关联习题', '%d' % (280 + seed % 90)),
                           ('认知维度', p[3][seed % 4][:2]), ('子节点', '%d' % (5 + seed % 4))]),
                  _hbars([('已完成', 62), ('进行中', 24), ('未开始', 14)], '%'),
                  _esc(focus), _esc('讲解「%s」的核心概念' % focus)))
    foot = ''.join('<div class="gmini"><b>%s</b><span>%s</span></div>' % (v, lb) for v, lb in
                   [('128 个', '知识点总量'), ('4 层', '节点层级'), ('216 条', '节点关系'),
                    ('486 份', '关联资源'), ('1,240 道', '关联习题'), ('32 个', '思政点')])
    return ('<div class="ws2">%s%s</div><div class="gfoot">%s</div>' % (canvas, detail, foot))


def _ws_live(fp, name, variant):
    """沉浸式大屏＋右栏签到/学生状态（对标参考图：讲义舞台＋代码块＋出勤环）。"""
    p, seed = _profile(fp), _seed(fp)
    bullets = ''.join('<li>%s</li>' % _esc(s) for s in p[4])
    code = ('<pre class="stage-code">void visit(Node s) {\n'
            '    Queue q = new Queue();\n'
            '    q.push(s); marked[s] = true;\n'
            '    while (!q.isEmpty()) {\n'
            '        Node u = q.pop(); handle(u);\n'
            '        for (Node v : adj(u))\n'
            '            if (!marked[v]) {\n'
            '                marked[v] = true; q.push(v);\n'
            '            }\n    }\n}</pre>')
    stage = ('<div class="stage"><div class="stage-top"><span class="liver">● 正在放映</span>'
             '<b>%s · 第 %d / %d 页</b><span class="on-air">%d 人在线</span></div>'
             '<div class="slide"><div class="slide-no">%02d</div><h3>%s</h3>'
             '<small>%s · %s</small><ul>%s</ul>%s</div>'
             '<div class="stage-foot"><span>%s</span><span>%d / %d</span></div></div>' % (
                 _esc(name), 12, 34, 40 + seed % 30, seed % 20, _esc(p[2][seed % 6]),
                 _esc(p[1]), _esc(p[2][0]), bullets, code, _esc(p[1]), 12, 34))
    attend = _ring(88, '到课率')
    stu = ''.join('<div class="stu"><span class="sd sd-%s"></span>%s</div>' %
                  (('g', 'o', 'b', 'r')[(seed + i) % 4], _OWNERS[i]) for i in range(10))
    rail = ('<div class="rail">'
            '<div class="rcard"><h4>课堂签到 <em>已签到 %d / 52</em></h4>'
            '<div class="attend">%s<div class="attend-n"><div><b>%d</b>已签到</div>'
            '<div><b>%d</b>迟到</div><div><b>%d</b>未签到</div></div></div>'
            '<div class="rbtns"><button class="btn btn-gh btn-sm" data-act="view">查看名单</button>'
            '<button class="btn btn-pri btn-sm" data-act="query">提醒未签到</button></div></div>'
            '<div class="rcard"><h4>学生状态 <em>在线 %d 人</em></h4><div class="stu-grid">%s</div>'
            '<div class="legend"><span><i class="sd sd-g"></i>已签到</span><span><i class="sd sd-o"></i>迟到</span>'
            '<span><i class="sd sd-b"></i>在线</span></div></div>'
            '<div class="rcard"><h4>课堂互动</h4><div class="mini2">%s</div></div></div>' % (
                46, attend, 46, 3, 6, 40 + seed % 30, stu,
                _minis([('弹幕', '128'), ('答题参与', '96'), ('举手', '5'), ('点赞', '214')])))
    return '<div class="ws-live">%s%s</div>' % (stage, rail)


def _ws_board(fp, name, variant):
    """页签＋KPI＋数据大表＋底部双卡（对标参考图：作业与考试）。"""
    p, seed = _profile(fp), _seed(fp)
    tabs = p[2][:4]
    nav = ('<div class="board-tabs">%s<span class="bt-right">待处理 %d 项</span></div>' % (
        ''.join('<button class="bt%s">%s</button>' % (' on' if i == 0 else '', _esc(t))
                for i, t in enumerate(tabs)), 3 + seed % 6))
    kpi = _tiles6([(tabs[0], str(3 + seed % 4), '项', '全部对象 · 本周'),
                   ('待处理', str(5 + seed % 6), '份', '↑ 已关联责任人'),
                   ('本周提交', '{:,}'.format(120 + seed % 260), '份', '较上周 +%d%%' % (3 + seed % 9)),
                   ('平均完成率', str(80 + seed % 15), '%', '目标 90%')])
    table = _card(name + ' · 明细列表', _toolbar(name, acts=('new', 'export', 'query')) +
                  _scenario_rows(fp, name, variant) + _pager(),
                  sub='共 %d 项 · 支持检索与导出' % (6 + seed % 8))
    bottom = ('<div class="grid2">' +
              _card('处理进度', _hbars([(p[2][i], 40 + (seed + i * 9) % 55) for i in range(4)], '份'),
                    sub='已完成自动处理，待人工确认') +
              _card('共性问题分析', _note([
                  '%s：%s偏差较大。' % (p[3][0], p[2][0]),
                  '%s：约 %d%% 对象未达标。' % (p[3][1], 20 + seed % 20),
                  '%s：建议补充过程性记录。' % p[3][2]]),
                  extra='<span class="more" data-act="query">查看全部</span>') + '</div>')
    return nav + kpi + table + bottom


def _ws_gen(fp, name, variant):
    """左侧材料进度列表＋右侧智能生成卡片流（对标参考图：智能备课）。"""
    p, seed = _profile(fp), _seed(fp)
    docs = ''
    for i, obj in enumerate(p[2][:4]):
        st = ('已生成', '待完善', '未开始')[i % 3]
        docs += ('<div class="docitem"><span class="dstat dstat-%s">%s</span>'
                 '<div class="dmain"><b>第 %d 讲 · %s</b><small>%d 页 · 已设计 %d 处交互</small></div>'
                 '<button class="btn btn-gh btn-sm" data-act="view">打开</button>'
                 '<button class="btn btn-gh btn-sm" data-act="calc">分析</button></div>' % (
                     ('ok', 'warn', 'mute')[i % 3], st, i + 6, _esc(obj),
                     20 + (seed + i * 5) % 16, (seed + i) % 6))
    left = _card('材料与进度', '<div class="doclist">%s</div>' % docs,
                 sub='本学期共 %d 讲' % (30 + seed % 8),
                 extra='<span class="more" data-act="query">目录</span>')
    kinds = [('单选题', 'b'), ('多选题', 'g'), ('编程题', 'p'), ('简答题', 'o')]
    cards = ''
    for i, (kind, c) in enumerate(kinds):
        st = ('待确认', '已插入')[ (seed + i) % 2]
        cards += ('<div class="aicard"><div class="ai-top">%s<span>插入位置 第 %d 页后</span>'
                  '<em class="%s">%s</em></div><b>%s：%s的核心要点是什么？</b>'
                  '<small>难度系数 %.2f · 结合%s</small>'
                  '<div class="ai-btns"><button class="btn btn-gh btn-sm" data-act="ins">插入成果</button>'
                  '<button class="btn btn-gh btn-sm" data-act="feedback">意见反馈</button>'
                  '<button class="btn btn-gh btn-sm" data-act="edit">编辑</button></div></div>' % (
                      _chip(kind, c), 11 + i * 3, 'okg' if (seed + i) % 2 else 'warno', st,
                      _esc(p[2][i % 6]), _esc(p[3][i % 4]), 0.5 + (seed % 40) / 100,
                      _esc(p[5][i % 3])))
    right = _card('智能生成', '<div class="aistream">%s</div>' % cards,
                  sub='依据材料生成，可一键插入对应位置',
                  extra='<button class="btn btn-pri btn-sm" data-act="ins">全部插入</button>')
    return '<div class="ws2 ws-gen">%s%s</div>' % (left, right)


def _ws_chat(fp, name, variant):
    """三栏对话工作台：角色列表 / 会话流 / 上下文面板（对标参考图：AI 工作台）。"""
    p, seed = _profile(fp), _seed(fp)
    roles = [('生成助手', '提纲 · 案例 · 习题', 0), ('智能学件', '答疑 · 自测 · 路径', 1),
             ('核验助手', '检查 · 批注 · 共因', 2), ('洞察助手', '预警 · 归因 · 建议', 3)]
    rolelist = ''.join('<div class="role%s"><span class="ric">%s</span><div><b>%s</b>'
                       '<small>%s</small></div></div>' %
                       (' on' if i == 0 else '', _ICO[i % 8], _esc(t), _esc(d))
                       for i, (t, d, _x) in enumerate(roles))
    recent = ''.join('<div class="rec"><i></i>%s · %d 分钟前</div>' % (_esc(p[2][i % 6]), 45 - i * 9)
                     for i in range(4))
    left = ('<div class="chat-l"><div class="cl-cap">助手角色</div>%s'
            '<div class="cl-cap">最近会话 · 7 天</div>%s</div>' % (rolelist, recent))
    outline = '<br>'.join('%d. %s：检查%s，形成%s。' % (j + 1, _esc(s), _esc(p[3][j]), _esc(p[5][j % 3]))
                          for j, s in enumerate(p[4]))
    thread = ('<div class="msg"><div class="av av-ai">AI</div><div class="bub">我是%s。当前聚焦%s，'
              '可基于知识网络与%d 份资源为工作提供支持。</div></div>'
              '<div class="msg me"><div class="bub">如何推进%s并形成可复核成果？</div>'
              '<div class="av av-me">我</div></div>'
              '<div class="msg"><div class="av av-ai">AI</div><div class="bub">可按以下步骤：'
              '<br>%s</div></div>' % (
                  _esc(name), _esc(p[1]), 400 + seed % 120, _esc(name), outline))
    chips = ''.join('<button class="qp" data-act="ask" data-prompt="%s">%s</button>' %
                    (_esc(q), _esc(q)) for q in [
                        '生成提纲框架', '找一个真实案例', '按内容出 5 道分层习题', '分析重难点分布'])
    center = ('<div class="chat-c"><div class="cc-head"><b>%s</b><span>%s · 已同步资源</span>'
              '<span class="sp"></span><span class="model">DeepSeek-V3 ▾</span></div>'
              '<div class="chat" data-chatbox>%s</div><div class="qps">%s</div>'
              '<div class="chatin"><input data-chatin placeholder="输入问题，或使用 / 调起指令">'
              '<button class="btn btn-pri" data-act="send">发送</button></div></div>' % (
                  _esc(roles[0][0]), _esc(p[1]), thread, chips))
    hits = ''.join('<div class="hit"><b>%s · 第 %d 章</b><small>命中相关段落，相似度 %.2f</small></div>' %
                   (_esc(p[1]), i + 6, 0.9 - i * 0.04) for i in range(3))
    sws = ''.join('<div class="sw-row"><span>%s</span><i class="sw on"></i></div>' % t
                  for t in ('联网检索', '引用溯源', '按对象定制'))
    right = ('<div class="chat-r"><div class="ctx"><span class="eyebrow">当前上下文</span>'
             '<div class="ctx-h"><span class="ctx-lg">%s</span><div><b>%s</b>'
             '<small>%s · %d 学时</small></div></div>%s</div>'
             '<div class="cl-cap">知识网络命中 · 3 条</div>%s' '<div class="cl-cap">检索与生成</div>%s</div>' % (
                 _ICO[5], _esc(name), _esc(p[1]), 48 + seed % 20,
                 _note(['%s v3.2' % p[1], '资源库 · %d 份' % (400 + seed % 90), '历年成果 · 5 套']),
                 hits, sws))
    return '<div class="ws-chat">%s%s%s</div>' % (left, center, right)


def _ws_dash(fp, name, variant):
    """驾驶舱：KPI＋健康度环＋趋势＋预警（对标参考图：教学总览）。"""
    p, seed = _profile(fp), _seed(fp)
    kpi = _tiles6([
        (p[2][0], str(120 + seed % 300), '项', '↑ 覆盖 %d 个对象' % (6 + seed % 20)),
        (p[3][0], str(86 + seed % 12), '%', '↑ 环比 +%.1fpt' % (0.5 + seed % 12 / 10)),
        ('今日调用', str(30 + seed % 60), '次', '实时进行中'),
        ('待关注', str(2 + seed % 5), '项', '↓ -1 较昨日')])
    health = _health_card(fp, name, seed)
    trend = _card('业务趋势', _bars([22 + (seed + j * 11) % 26 for j in range(6)],
                                    ['5月', '6月', '7月', '8月', '9月', '10月']), sub='近六个月 · 按月汇总')
    calls = _card('维度构成', _hbars([(p[3][i], 40 + (seed + i * 13) % 55) for i in range(4)], '%'))
    attention = _card('需要关注', _alerts(p[2], seed, _DEPTS), sub='近 7 日',
                      extra='<span class="more" data-act="query">全部预警</span>')
    return kpi + '<div class="grid2">%s%s</div><div class="grid2">%s%s</div>' % (
        health, trend, calls, attention)


def _ws_manage(fp, name, variant):
    """通用办理工作台：配置表单＋成果预览＋动态。"""
    p = _profile(fp)
    options = ''.join('<option>%s</option>' % _esc(n) for n in p[2])
    form = ('<div class="frm"><label>业务对象<select>%s</select></label>'
            '<label>执行周期<select><option>2026 年第四季度</option><option>2026 年第三季度</option></select></label>'
            '<label>重点维度<select>%s</select></label><label>工作目标<textarea rows="3">%s</textarea></label>'
            '<button class="btn btn-pri" data-act="compose">生成工作成果</button></div>' % (
                options, ''.join('<option>%s</option>' % _esc(n) for n in p[3]),
                _esc('完成%s，建立%s与%s的对应关系。' % (name, p[2][0], p[5][0]))))
    outline = ''.join('<div class="action-line"><span class="step-dot">%02d</span><div><b>%s</b>'
                      '<small>检查 %s，形成 %s</small></div></div>' %
                      (j + 1, _esc(stage), _esc(p[3][j]), _esc(p[5][j % 3]))
                      for j, stage in enumerate(p[4]))
    return ('<div class="workspace-grid">' + _card(name + ' · 工作配置', form, sub=p[1]) +
            _card('成果预览', '<div class="result-preview">%s</div>' % outline,
                  extra='<button class="btn btn-gh btn-sm" data-act="export">导出</button>') + '</div>')


_WS = {'graph': _ws_graph, 'live': _ws_live, 'board': _ws_board,
       'gen': _ws_gen, 'chat': _ws_chat, 'dash': _ws_dash, 'manage': _ws_manage}


def _business_workspace(fp, name, variant, kind=None):
    return _WS[kind or _ws_kind(fp)](fp, name, variant)


_WS_ORDER = ('graph', 'live', 'chat', 'gen', 'board', 'dash', 'manage')


def _plan_kinds(mods):
    """全局版式配额：同版式最多 3 页（功能很多时按 7 版式均摊），
    超额或相邻重复时轮换到用得最少的其它版式，杜绝整屏千篇一律。"""
    total = sum(len(gfps) for _, gs in mods for gfps, _gn in gs)
    cap = max(3, -(-total // len(_WS_ORDER)))
    counter = {k: 0 for k in _WS_ORDER}
    plan, prev = {}, None
    for mid, (_nm, gs) in enumerate(mods):
        for i, (gfps, _gn) in enumerate(gs):
            for j, fp in enumerate(gfps):
                first = _ws_kind(fp)
                if counter[first] < cap and first != prev:
                    kind = first
                else:
                    kind = min(_WS_ORDER,
                               key=lambda x: (x == prev, counter[x] >= cap,
                                              counter[x], x != first))
                counter[kind] += 1
                prev = kind
                plan[(mid, i, j)] = kind
    return plan


def _view_content(code, fp, name, variant, kind=None):
    p = _profile(fp)
    seed = _seed(fp)
    if code == 'workspace':
        return _business_workspace(fp, name, variant, kind)
    if code in ('records', 'archive'):
        return (_card(name + ' · ' + ('成果清单' if code == 'records' else '发布归档'),
                      _toolbar(name) + _scenario_rows(fp, name, variant)) +
                _card('处理轨迹', _tl([(s, _OWNERS[i] + ' · ' + p[5][i % 3], i < 2)
                                      for i, s in enumerate(p[4])]), sub='责任与成果对应'))
    if code == 'relations':
        return '<div class="grid2">' + _relation_card(fp, name) + _card('关联明细', _note([
            '%s → %s：%s' % (p[2][j], p[5][j % 3], p[3][j % 4]) for j in range(6)])) + '</div>'
    if code == 'analysis':
        vals = [42 + (seed + j * 13) % 47 for j in range(6)]
        return ('<div class="grid2">' + _card('业务趋势', _bars(vals, ['5月', '6月', '7月', '8月', '9月', '10月']),
                 sub='峰值 %d · 按月汇总' % max(vals)) +
                _card('能力与质量雷达', _radar(list(p[3]), [83, 91, 86, 94])) + '</div>' +
                _card('评估对象明细', _scenario_rows(fp, name, variant)))
    if code == 'design':
        return _business_workspace(fp, name, variant) + _card('环节配置',
            '<div class="steps">%s</div>' % ''.join('<div class="st"><div class="n">%d</div>'
            '<div><b>%s</b><div class="ds">%s · %s</div></div></div>' %
            (i + 1, _esc(s), _esc(p[3][i]), _esc(p[5][i % 3])) for i, s in enumerate(p[4])))
    if code == 'materials':
        return ('<div class="grid2">' + _card('材料导入',
                '<div class="upzone" data-act="upload"><div class="ar">↑</div>选择本地材料'
                '<small>支持文档、表格和图片；仅在当前页面展示文件信息</small></div>'
                '<div class="upload-results"></div>') +
                _card('资源分类', _hbars([(n, 12 + (seed + i * 7) % 35) for i, n in enumerate(p[5])], '份')) +
                '</div>' + _card('材料清单', _scenario_rows(fp, name, variant)))
    if code == 'versions':
        return ('<div class="grid2">' + _card('版本选择', '<div class="frm"><label>基准版本<select>'
                '<option>第一版</option></select></label><label>当前版本<select><option>第二版</option>'
                '</select></label><button class="btn btn-pri" data-act="calc">比较版本</button></div>') +
                _card('修订摘要', _note(['补充%s的说明。' % p[2][0], '调整%s的检查口径。' % p[3][1],
                                     '完善%s的交付清单。' % p[5][2]])) + '</div>' +
                _card('差异明细', _table(['业务维度', '第一版', '第二版', '变更'], ''.join(
                    '<tr><td>%s</td><td>%d 项</td><td>%d 项</td><td>补充 2 项</td></tr>' %
                    (_esc(d), 8 + i, 10 + i) for i, d in enumerate(p[3])))))
    return ('<div class="grid2">' + _card('执行节点', _tl([
        (s, _OWNERS[i] + ' · ' + p[3][i], i < 2) for i, s in enumerate(p[4])])) +
        _card('检查进展', '<div class="ringwrap">%s</div>' % ''.join(
            _ring(82 + i * 4, n) for i, n in enumerate(p[3]))) + '</div>' +
        _card('跟进明细', _scenario_rows(fp, name, variant)))


def _pills():
    return ('<div class="hero-acts"><div class="pills"><button class="on">近 7 日</button>'
            '<button>近 30 日</button><button>本年度</button></div>'
            '<button class="btn btn-gh btn-sm" data-act="export">导出概况</button>'
            '<button class="btn btn-pri btn-sm" data-act="query">查看全景</button></div>')


def _fcard(mid, i, gfps, gname):
    seed = _seed(' '.join(gfps))
    stats = ((6 + seed % 20, '对象'), (10 + seed % 40, '次'),
             ('{:,}'.format(120 + (seed * 7) % 3800), '记录'))
    views = sum(len(_feature_views(f)) for f in gfps)
    multi = '%d 项功能 · ' % len(gfps) if len(gfps) > 1 else ''
    return ('<button class="fcard" data-pg="pg-m%d" data-ft="ft-%d-%d">'
            '<span class="fi">%s</span><b>%s%s</b><div class="fs">%s</div>'
            '<small>%s%d 个业务环节 · 进入 →</small></button>' % (
                mid, mid, i, _ICO[(i + mid + 1) % 8], _esc(gname),
                '等%d项' % len(gfps) if len(gfps) > 1 else '',
                ''.join('<span><b>%s</b>%s</span>' % (v, u) for v, u in stats),
                multi, views))


def _alerts(names, seed, dept):
    tpl = [('连续 2 轮待复核', 'dot-r', 'tag-r', '高优先'),
           ('有材料尚未提交', 'dot-o', 'tag-o', '中优先'),
           ('调用量环比下降', 'dot-o', 'tag-o', '中优先'),
           ('长期未更新，待确认', 'dot-b', 'tag-b', '低优先')]
    out = ''
    for j in range(min(4, max(len(names), 1))):
        nm = names[j] if j < len(names) else dept[j % 6]
        t, dot, cls, pr = tpl[(j + seed) % 4]
        out += ('<div class="alert"><span class="dot %s"></span><div><b>%s·%s</b>'
                '<small>%s · 责任人 %s</small></div><div><span class="tag %s">%s</span>'
                '<time>%d 小时前</time></div></div>' % (
                    dot, _esc(nm), t, _esc(dept[(j + seed) % 6]), _OWNERS[(j + seed) % 10],
                    cls, pr, 1 + (j * 5 + seed) % 23))
    return out


def _health_card(fp_all, name, seed):
    p = _profile(fp_all)
    score = 82 + seed % 13
    dims = [(d, 74 + (seed + j * 9) % 26) for j, d in enumerate(p[3])]
    return _card('运行健康度', '<div class="health">%s<div class="hbars">%s</div></div>'
                 '<div class="note-line">较上月提升 %.1f 分，%s与%s是主要贡献项。</div>' % (
                     _ring(score, '综合健康度'), _hbars(dims, '%'),
                     1.5 + seed % 3 / 10, _esc(p[3][seed % 4]), _esc(p[3][(seed + 1) % 4])),
                 sub='综合四项指标', extra='<span class="tag tag-g">良好</span>')


def _fp_block(mid, i, j, fp, name, variant=1, kind=None):
    """单个功能的完整内容块：指标 + 视图页签 + 各视图面板（页内切换，不拆页）。"""
    views = _feature_views(fp)
    p, seed = _profile(fp), _seed(fp)
    nav = '<div class="view-nav">%s</div>' % ''.join(
        '<button class="btn %s" data-vt="vp-%d-%d-%d-%d">%s</button>' %
        ('btn-pri' if k == 0 else 'btn-gh', mid, i, j, k, _esc(t)) for k, (_, t) in enumerate(views))
    metrics = _tiles6([(name, str(18 + seed % 120), '项', '当前业务对象'),
                      (p[3][0], str(86 + seed % 14), '%', '↑ 当前核验进度'),
                      ('待跟进', str(2 + seed % 5), '项', '已关联责任人'),
                      ('交付成果', str(len(p[5])), '类', '支持导出与复核'),
                      ('业务环节', str(len(views)), '个', '页内签切换'),
                      ('执行节点', str(len(p[4])), '个', '全程跟踪')])
    panels = []
    for k, (code, title) in enumerate(views):
        main = _view_content(code, fp, name, variant, kind)
        context = _context_cards(fp, name, variant)
        if variant == 1:
            body = main + context
        elif variant == 2:
            body = context + main
        else:
            body = '<div class="scheme-split"><div>' + main + '</div>' + context + '</div>'
        panels.append('<div class="vpanel%s" id="vp-%d-%d-%d-%d" data-view="%s">%s</div>' %
                      (' on' if k == 0 else '', mid, i, j, k, code, body))
    return ('<div class="fpblock%s" id="fb-%d-%d-%d"><div class="fphead"><h3>%s</h3>'
            '<span>%s · %d 个业务环节，均在本页展开</span></div>%s%s%s</div>' % (
                ' on' if j == 0 else '', mid, i, j,
                _esc(name), _esc(p[1]), len(views), metrics, nav, ''.join(panels)))


def _fg_panel(mid, i, gfps, gname, variant=1, plan=None):
    """一个导航子项 = 一屏：多功能合并组收进一个大按钮，点开下拉菜单逐项切换。"""
    nav = ''
    if len(gfps) > 1:
        items = ''.join(
            '<button class="fb-item%s" data-fb="fb-%d-%d-%d">%s</button>' %
            (' on' if j == 0 else '', mid, i, j, _esc(_short_name(fp)))
            for j, fp in enumerate(gfps))
        nav = ('<div class="fbbar"><button class="fb-big" type="button">'
               '<span class="fb-ico">☰</span><span class="fb-txt">当前功能 <b class="fb-cur">%s</b>'
               '<span class="fb-meta">共 %d 项 · <i class="fb-idx">第1项</i></span></span>'
               '<span class="fb-caret">切换功能 ▾</span></button>'
               '<div class="fb-menu">%s</div></div>' % (_esc(_short_name(gfps[0])), len(gfps), items))
    blocks = ''.join(
        _fp_block(mid, i, j, fp, _short_name(fp) if len(gfps) > 1 else gname, variant,
                  (plan or {}).get((mid, i, j)))
        for j, fp in enumerate(gfps))
    return ('<div class="fppanel%s" id="ft-%d-%d" data-feature="%d-%d">%s%s</div>' %
            (' on' if i == 0 else '', mid, i, mid, i, nav, blocks))


def _module_page(mid, name, groups, variant=1, plan=None):
    """一个模块 = 一页：指标、功能概览、健康度、趋势、预警与全部功能内容集中呈现。"""
    nm = _short_name(name)
    fps_all = [f for gf, _ in groups for f in gf]
    fp_all = ' '.join(fps_all)
    seed = _seed(name + fp_all)
    p = _profile(fp_all)
    count = sum(len(_feature_views(f)) for f in fps_all)
    kpis = _tiles6([
        ('业务功能', str(len(groups)), '项', '↑ 全部有入口'),
        ('业务环节', str(count), '个', '页内签切换'),
        ('在办事项', str(8 + seed % 15), '项', '↑ 本周新增 %d 项' % (1 + seed % 5)),
        ('按期办结率', str(93 + seed % 6), '%', '↑ 环比 +%.1fpt' % (0.5 + seed % 12 / 10)),
        ('智能调用', '{:,}'.format(620 + seed % 4200), '次', '↑ 近 7 天 %d 次' % (80 + seed % 260)),
        ('需关注事项', str(1 + seed % 4), '项', '↓ -1 较昨日')])
    overview = _card('功能总览', '<div class="fgrid">%s</div>' %
                     ''.join(_fcard(mid, i, gf, gn) for i, (gf, gn) in enumerate(groups)),
                     sub='%d 项功能集中在一页 · 点击卡片直达' % len(groups), extra=_pills())
    trend = _card('业务趋势', _bars([22 + (seed + j * 11) % 26 for j in range(6)],
                                    ['5月', '6月', '7月', '8月', '9月', '10月']),
                  sub='近六个月 · 按月汇总')
    gnames = [gn for _, gn in groups]
    attention = _card('需要关注', _alerts(gnames, seed, _DEPTS), sub='近 7 日',
                      extra='<span class="more" data-act="query">全部预警</span>')
    calls = _card('智能调用详情', _hbars([(n, 30 + (seed + i * 13) % 90)
                                         for i, n in enumerate(gnames[:5])], '次'),
                  sub='按功能统计 · 高峰集中在 20:00-22:00')
    tabs = '<div class="ftabs">%s</div>' % ''.join(
        ['<button class="ftab on" data-pg="pg-m%d" data-ft="">总览</button>' % mid] +
        ['<button class="ftab" data-pg="pg-m%d" data-ft="ft-%d-%d">%s</button>' %
         (mid, mid, i, _esc(gn)) for i, (_gf, gn) in enumerate(groups)])
    panels = ''.join(_fg_panel(mid, i, gf, gn, variant, plan) for i, (gf, gn) in enumerate(groups))
    if variant == 2:
        charts = '<div class="grid2">' + calls + trend + '</div>' + attention
    elif variant == 3:
        charts = '<div class="grid2">' + trend + attention + calls + '</div>'
    else:
        charts = ('<div class="grid2">' + _health_card(fp_all, nm, seed) + trend + '</div>'
                  '<div class="grid2">' + attention + calls + '</div>')
    return ('<section class="page" id="pg-m%d"><div class="hero">'
            '<div class="eyebrow">%s / 业务模块</div><h2>%s</h2>'
            '<p>%d 项功能 · %d 个业务环节 · 每个功能独立一屏，点左侧导航逐项查看 · %s</p></div>'
            '<div class="modsum on">%s%s%s</div>'
            '<div class="mod-block"><div class="mb-head"><b>功能工作区</b>'
            '<span>一个导航项一屏 · 组内多功能用页签逐屏切换</span></div>%s%s</div></section>' % (
                mid, _esc(_page_scheme(variant)), _esc(nm), len(groups), count, _esc(p[1]),
                kpis, overview, charts, tabs, panels))


def _work_page(mods, variant=1):
    """工作台 = 驾驶舱：所有模块的概况、指标、趋势与预警汇聚一页。mods 已按导航规则合并。"""
    total_f = sum(len(gs) for _, gs in mods)
    seed = _seed(''.join(n for n, _ in mods) + str(total_f))
    all_names = [_short_name(n) for n, _ in mods]
    kpis = _tiles6([
        ('业务模块', str(len(mods)), '个', '↑ 全部正常运转'),
        ('业务功能', str(total_f), '项', '↑ 覆盖 %d 个模块' % len(mods)),
        ('今日访问', str(96 + seed % 60), '次', '实时进行中 %d 次' % (12 + seed % 20)),
        ('功能完备率', '96.4', '%', '↑ +1.2pt 全员可体验'),
        ('智能调用', '{:,}'.format(9000 + seed % 5200), '次',
         '↑ +18.6%% 人均 %d 次' % (2 + seed % 4)),
        ('需关注事项', str(2 + seed % 5), '项', '↓ -3 含预警处置')])
    mcards = ''.join(
        '<button class="fcard" data-pg="pg-m%d"><span class="fi">%s</span><b>%s</b>'
        '<div class="fs"><span><b>%d</b>功能</span><span><b>%d</b>环节</span>'
        '<span><b>%d</b>对象</span></div>'
        '<small>%s · 进入 →</small></button>' % (
            mid, _ICO[(mid + 1) % 8], _short_name(name), len(gs),
            sum(len(_feature_views(f)) for gf, _ in gs for f in gf), len(gs) * 6,
            ' / '.join(gn for _, gn in gs[:3]))
        for mid, (name, gs) in enumerate(mods))
    overview = _card('模块运行概况', '<div class="fgrid">%s</div>' % mcards,
                     sub='今日 · 全部模块',
                     extra='<span class="more" data-pg="pg-m0">查看模块明细</span>')
    fp0 = ' '.join(f for gf, _ in mods[0][1] for f in gf)
    nm0 = _short_name(mods[0][0])
    detail = _card('重点业务记录', _toolbar(nm0) + _scenario_rows(fp0, nm0, variant),
                   sub='全量演练数据 · 非实时业务')
    graph = _relation_card(fp0, nm0)
    attention = _card('需要关注', _alerts(all_names, seed, _DEPTS), sub='近 7 日',
                      extra='<span class="more" data-act="query">全部预警</span>')
    calls = _card('智能体会话统计', _minis([('教师调用', str(40 + seed % 40)),
                                          ('员工调用', '{:,}'.format(380 + seed % 260)),
                                          ('日均会话', str(60 + seed % 40))]) +
                  _hbars([(n, 120 + (seed + i * 17) % 160) for i, n in enumerate(all_names[:5])], '次'),
                  sub='近 7 日 · 按模块统计')
    trend = _card('业务总量趋势', _bars([30 + (seed + j * 13) % 40 for j in range(7)],
                                        ['周一', '周二', '周三', '周四', '周五', '周六', '周日']),
                  sub='近 7 日 · 每 5 分钟同步一次')
    health = _health_card(fp0, nm0, seed)
    if variant == 2:
        body = (kpis + overview + '<div class="grid2">' + calls + trend + '</div>' +
                health + detail + attention)
    elif variant == 3:
        body = (kpis + overview + attention +
                '<div class="grid2">' + health + calls + '</div>' + trend + detail)
    else:
        body = (kpis + overview +
                '<div class="grid2">' + health + trend + '</div>' +
                '<div class="grid2">' + calls + attention + '</div>' + detail + graph)
    return ('<section class="page on" id="pg-work"><div class="hero">'
            '<div class="eyebrow">%s · 驾驶舱</div><h2>业务总览</h2>'
            '<p>跨越 %d 个模块 · %d 项功能 · 数据每 5 分钟同步一次，全部概况集中一页</p>%s</div>'
            '%s</section>' % (_esc(_page_scheme(variant)), len(mods), total_f, _pills(), body))


def _sys_page():
    rows = ''
    roles = ['管理员', '业务主管', '操作员']
    for i in range(5):
        st = _STATUS[i % 2]
        rows += ('<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                 % (_OWNERS[i], roles[i % 3], _DEPTS[i % 6], _TIMES[i % 5], _tag(st[0])))
    return ('<section class="page" id="pg-sys"><div class="hero"><h2>系统管理</h2>'
            '<p>账号、角色与权限统一维护</p></div>' +
            _tiles6([('在册账号', '26', '个', '本月新增 3 个'), ('角色', '3', '类', '按角色隔离权限'),
                     ('安全事件', '0', '件', '今日巡检正常'), ('在线会话', '8', '个', '峰值 14 个'),
                     ('密码到期', '2', '个', '已发送提醒'), ('授权变更', '5', '次', '近 30 天')]) +
            '<div class="grid2" style="margin-top:16px">' +
            _card('用户管理', _toolbar('用户', acts=('new', 'export')) +
                  _table(['姓名', '角色', '部门', '最近登录', '状态'], rows)) +
            _card('角色权限', _table(['角色', '权限范围', '成员数'],
                  '<tr><td>管理员</td><td>全部模块配置与授权</td><td>2</td></tr>'
                  '<tr><td>业务主管</td><td>业务审批、统计查看</td><td>6</td></tr>'
                  '<tr><td>操作员</td><td>业务办理、记录填报</td><td>18</td></tr>') +
                  '<div style="margin-top:14px"><button class="btn btn-gh btn-sm" data-act="edit" '
                  'data-name="权限方案">调整权限方案</button></div>') + '</div><div class="grid2">' +
            _card('登录统计', _minis([('今日登录', '18'), ('异地登录', '0'), ('锁定账号', '1')]) +
                  _bars([12, 16, 14, 18, 15, 9, 6],
                        ['周一', '周二', '周三', '周四', '周五', '周六', '周日']), sub='近 7 日登录人次') +
            _card('安全策略', _note([
                '密码 90 天强制轮换，历史密码不可复用。',
                '连续 5 次错密自动锁定 30 分钟。',
                '敏感操作二次确认并记入操作日志。',
                '导出行为水印留痕，支持事后追溯。']) +
                '<div class="chips"><span>密码策略</span><span>登录锁定</span>'
                '<span>二次确认</span><span>导出水印</span></div>') +
            '</div>' + _tail() + '</section>')


def _log_page(mods):
    acts = ['新建记录', '审批通过', '导出报表', '修改配置', '提交办理']
    rows = ''
    for i in range(8):
        nm = sanitize(mods[i % len(mods)][0]) if mods else '业务管理'
        rows += ('<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>10.20.31.%d</td></tr>'
                 % (_TIMES[i % 5], _OWNERS[i % 10], _esc(nm), acts[i % 5], 12 + i * 3))
    return ('<section class="page" id="pg-log"><div class="hero"><h2>操作日志</h2>'
            '<p>全量留痕，支持按模块与操作人追溯</p></div>' +
            _card('日志检索', _minis([('今日操作', '136'), ('涉及模块', '6'), ('异常操作', '0')]) +
                  _toolbar('日志', acts=('export', 'query')) +
                  _table(['时间', '操作人', '模块', '动作', 'IP'], rows)) +
            '<div class="grid2">' +
            _card('操作类型分布', _hbars([('记录填报', 46), ('审批流转', 32), ('导出报表', 21),
                                          ('配置变更', 9)]), sub='近 30 天') +
            _card('高频操作人', _hbars([(_OWNERS[0], 28), (_OWNERS[1], 22), (_OWNERS[2], 17),
                                        (_OWNERS[3], 11)]), sub='按操作次数') +
            '</div>' + _tail() + '</section>')


# ---------------- 风格冻结 v2 CSS ----------------
_CSS = '''
:root{--pri:#3370ff;--pri2:#14c0ff;--pur:#7b61ff;--ok:#00b96b;--warn:#ff9f00;--dan:#f53f3f;--ink:#1f2329;--ink2:#4e5969;--ink3:#86909c;--bg:#f5f7fa;--line:#e5e6eb;--sh:0 2px 12px rgba(31,35,41,.06)}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"PingFang SC","Microsoft YaHei",system-ui,sans-serif;color:var(--ink);background:var(--bg);font-size:14px;line-height:1.6}
button{font-family:inherit;cursor:pointer}
.topbar{height:60px;background:#fff;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:14px;padding:0 20px;position:sticky;top:0;z-index:40}
.crumb b{font-size:16px;font-weight:700;display:block;line-height:1.3}
.crumb span{font-size:12px;color:var(--ink3)}
.spacer{flex:1}
.search{display:flex;align-items:center;gap:8px;background:#f2f3f5;border-radius:999px;padding:8px 16px;width:260px;color:var(--ink3)}
.search.sm{width:200px;padding:6px 14px}
.search input{border:none;outline:none;background:transparent;font-size:13px;width:100%;color:var(--ink)}
.search svg{width:15px;height:15px;stroke:#86909c;fill:none;stroke-width:2;flex-shrink:0}
.icobtn svg{width:16px;height:16px;stroke:currentColor;fill:none;stroke-width:1.8}
.icobtn{width:34px;height:34px;border-radius:50%;border:1px solid var(--line);background:#fff;display:grid;place-items:center;color:var(--ink2);font-size:15px}
.icobtn:hover{color:var(--pri);border-color:var(--pri)}
.layout{display:flex;min-height:calc(100vh - 60px)}
.sidebar{width:220px;background:#fff;border-right:1px solid var(--line);padding:10px 8px;flex-shrink:0;position:sticky;top:60px;height:calc(100vh - 60px);overflow:auto}
.brand{display:flex;gap:10px;align-items:center;padding:12px;border-bottom:1px solid var(--line);margin:2px 0 10px}
.brand .lg{width:38px;height:38px;border-radius:10px;background:linear-gradient(135deg,#3370ff,#7b61ff);display:grid;place-items:center;flex-shrink:0}
.brand .lg svg{width:20px;height:20px;stroke:#fff;fill:none;stroke-width:1.8}
.brand b{font-size:15px;font-weight:700;display:block;line-height:1.3;white-space:nowrap}
.brand span{font-size:10px;color:var(--ink3);letter-spacing:1px;white-space:nowrap}
.cap{font-size:12px;color:var(--ink3);padding:12px 12px 6px}
.mi{display:flex;align-items:center;gap:10px;padding:9px 12px;border-radius:8px;color:var(--ink2);cursor:pointer;margin-bottom:2px;user-select:none;white-space:nowrap;font-size:14px}
.mi svg{width:17px;height:17px;stroke:currentColor;fill:none;stroke-width:1.6;flex-shrink:0}
.mi:hover{background:#f2f3f5}
.mi.on{background:#e8f0ff;color:var(--pri);font-weight:600}
.mi .tx,.kid .tx{overflow:hidden;text-overflow:ellipsis}
.mi .che{margin-left:auto;color:var(--ink3);font-size:12px;transition:transform .2s;flex-shrink:0}
.mi.open .che{transform:rotate(90deg)}
.kids{display:none;margin:0 0 6px 20px;padding-left:8px;border-left:1px solid var(--line)}
.kids.open{display:block}
.kid{display:flex;gap:8px;align-items:center;padding:7px 10px;border-radius:6px;color:var(--ink2);cursor:pointer;font-size:13px;white-space:nowrap;user-select:none}
.kid i{font-style:normal;color:var(--ink3);font-size:12px;min-width:16px;flex-shrink:0}
.kid:hover{background:#f2f3f5}
.kid.on{background:#e8f0ff;color:var(--pri);font-weight:600}
.kid.on i{color:var(--pri)}
.main{flex:1;padding:0 24px 28px;min-width:0}
.hero{background:linear-gradient(115deg,#e8f0ff 0%,#eef6ff 45%,rgba(245,247,250,0) 100%);padding:22px 4px 16px;margin-bottom:16px}
.hero h2{font-size:20px;font-weight:700}
.hero p{color:var(--ink2);font-size:13px;margin-top:4px}
.card{background:#fff;border-radius:12px;box-shadow:var(--sh);padding:20px;margin-bottom:16px}
.card .ch{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:14px}
.ch h3{font-size:16px;font-weight:600}
.ch .sub{font-size:12px;color:var(--ink3)}
.more{color:var(--pri);font-size:12px;cursor:pointer}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
.tiles{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}
.tile{border-radius:10px;padding:16px 18px;position:relative;overflow:hidden}
.tile .lb{font-size:13px;color:var(--ink2)}
.tile .vl{font-size:26px;font-weight:700;margin:4px 0 2px;line-height:1.2}
.tile .vl i{font-size:12px;color:var(--ink3);font-style:normal;margin-left:3px;font-weight:400}
.tile .sb{font-size:12px;color:var(--ink3)}
.tile .gh{position:absolute;right:12px;bottom:4px;font-size:46px;opacity:.10}
.t-b{background:linear-gradient(135deg,#edf4ff,#e2ecff)}.t-p{background:linear-gradient(135deg,#f4f0ff,#ece6ff)}
.t-g{background:linear-gradient(135deg,#e8f9f0,#ddf3e8)}.t-o{background:linear-gradient(135deg,#fff4e8,#ffeadd)}
.t-c{background:linear-gradient(135deg,#e6f8fd,#d9f2fb)}.t-k{background:linear-gradient(135deg,#ffeef5,#ffe3ef)}
.tabs{display:flex;gap:26px;border-bottom:1px solid var(--line);margin-bottom:16px;flex-wrap:wrap}
.tb{padding:10px 2px;font-size:14px;color:var(--ink2);border:none;background:none;border-bottom:2px solid transparent;margin-bottom:-1px}
.tb.on{color:var(--pri);font-weight:600;border-bottom-color:var(--pri)}
.cov{background:linear-gradient(90deg,#3370ff,#7b61ff);color:#fff;border-radius:10px;padding:10px 16px;display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:16px;font-size:13px}
.cov b{font-weight:600;margin-right:4px}
.cov .no{background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.35);border-radius:999px;padding:3px 12px;cursor:pointer;font-size:12px}
.cov .no:hover{background:rgba(255,255,255,.32)}
.frm label{display:block;font-size:13px;color:var(--ink2);margin-bottom:12px}
.frm .ctl{display:block;margin-top:6px}
.frm select,.frm input,.frm textarea,.cell{width:100%;border:1px solid var(--line);border-radius:8px;padding:9px 12px;font-size:13px;font-family:inherit;background:#fff;outline:none;color:var(--ink)}
.frm select:focus,.frm input:focus,.frm textarea:focus{border-color:var(--pri)}
.cell{padding:6px 8px;width:74px}
.tb-row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:14px}
.tb-row .sp{flex:1}
.btn{border-radius:6px;padding:8px 18px;font-size:13px;font-weight:500;border:1px solid transparent}
.btn-pri{background:var(--pri);color:#fff}
.btn-pri:hover{background:#2860e0}
.btn-gh{background:#fff;border-color:var(--line);color:var(--ink2)}
.btn-gh:hover{color:var(--pri);border-color:var(--pri)}
.btn-sm{padding:5px 12px;font-size:12px}
.seg{display:inline-flex;background:#f2f3f5;border-radius:6px;padding:2px}
.sg{padding:5px 14px;font-size:12px;border-radius:4px;border:none;background:none;color:var(--ink2)}
.sg.on{background:#fff;color:var(--ink);font-weight:600;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.tbl{width:100%;border-collapse:collapse;font-size:13px}
.tbl th{background:#fafbfc;color:var(--ink3);font-weight:500;text-align:left;padding:9px 12px;font-size:12px}
.tbl td{padding:10px 12px;border-bottom:1px solid #f2f3f5}
.tbl tr:hover td{background:#fafbff}
.tag{border-radius:4px;padding:2px 8px;font-size:12px}
.tag-b{background:#e8f3ff;color:#3370ff}.tag-g{background:#e8f9f0;color:#00b96b}
.tag-o{background:#fff4e8;color:#d4771f}.tag-r{background:#ffece8;color:#f53f3f}
.minis{display:flex;gap:28px;flex-wrap:wrap;margin-bottom:12px}
.minis div b{font-size:20px;font-weight:700}
.minis div span{display:block;font-size:12px;color:var(--ink3)}
.bars{display:flex;align-items:flex-end;gap:18px;height:180px;padding-top:22px}
.bars .c{flex:1;text-align:center;height:100%;display:flex;flex-direction:column;justify-content:flex-end}
.bars .v{width:58%;margin:0 auto;background:linear-gradient(180deg,#6ea6ff,#3370ff);border-radius:6px 6px 2px 2px;position:relative;min-height:6px}
.bars .v b{position:absolute;top:-18px;left:0;right:0;font-size:11px;color:var(--ink2);font-weight:600}
.bars .lb{font-size:11px;color:var(--ink3);margin-top:6px}
.hbar{display:grid;grid-template-columns:110px 1fr 56px;align-items:center;gap:12px;margin-bottom:14px;font-size:13px}
.hbar .tr{background:#f2f3f5;border-radius:999px;height:8px}
.hbar .fl{background:linear-gradient(90deg,#3370ff,#14c0ff);border-radius:999px;height:8px}
.hbar b{text-align:right;font-size:12px}
.ringwrap{display:flex;gap:14px;flex-wrap:wrap;justify-content:space-around}
.ring svg{width:104px;height:104px}
.upzone{border:2px dashed #c9cdd4;border-radius:10px;padding:24px 12px;text-align:center;color:var(--ink3);cursor:pointer;font-size:13px}
.upzone:hover{border-color:var(--pri);color:var(--pri);background:#f7faff}
.upzone .ar{font-size:26px;color:var(--pri);line-height:1.2}
.steps .st{display:flex;gap:12px;padding:12px 0;border-bottom:1px dashed var(--line)}
.steps .st:last-child{border-bottom:none}
.st .n{width:22px;height:22px;border-radius:50%;background:#e8f0ff;color:var(--pri);font-size:12px;display:grid;place-items:center;font-weight:700;flex-shrink:0;margin-top:2px}
.st .tt{font-weight:600;font-size:14px}
.st .mt{font-size:12px;color:var(--ink3);margin:2px 0 4px}
.st .ds{font-size:13px;color:var(--ink2)}
.chat{background:#f7f8fa;border-radius:10px;padding:16px;height:340px;overflow:auto}
.msg{display:flex;gap:10px;margin-bottom:14px}
.av{width:34px;height:34px;border-radius:50%;display:grid;place-items:center;font-size:12px;flex-shrink:0}
.av-ai{background:#e8f0ff;color:var(--pri);font-weight:700}
.av-me{background:var(--pri);color:#fff}
.bub{background:#fff;border:1px solid var(--line);border-radius:8px;padding:10px 14px;font-size:13px;line-height:1.7;max-width:78%}
.msg.me{justify-content:flex-end}
.msg.me .bub{background:var(--pri);color:#fff;border:none}
.chatin{display:flex;gap:10px;margin-top:12px}
.chatin input{flex:1;border:1px solid var(--line);border-radius:8px;padding:9px 12px;font-size:13px;outline:none}
.chatin input:focus{border-color:var(--pri)}
.qc{background:#fff;border:1px solid var(--line);border-radius:8px;padding:10px 12px;margin-bottom:10px;cursor:pointer}
.qc b{display:block;font-size:13px}
.qc span{font-size:12px;color:var(--ink3)}
.qc:hover{border-color:var(--pri)}
.quick{display:flex;gap:10px;flex-wrap:wrap}
.kv .r{margin-bottom:10px;font-size:13px;line-height:1.7}
.kv .k{color:var(--ink3)}
.kd{display:flex;gap:8px;margin-bottom:10px;font-size:13px;color:var(--ink2)}
.kd .d{width:8px;height:8px;border-radius:50%;background:var(--pri);margin-top:7px;flex-shrink:0}
.chips{display:flex;gap:8px;flex-wrap:wrap}
.chips span{background:#f2f3f5;border-radius:6px;padding:4px 10px;font-size:12px;color:var(--ink2)}
.diff .dr{display:flex;gap:10px;align-items:flex-start;padding:10px 0;border-bottom:1px dashed var(--line);font-size:13px}
.diff .dr:last-child{border-bottom:none}
.diff .ic{width:20px;height:20px;border-radius:50%;display:grid;place-items:center;font-size:12px;color:#fff;flex-shrink:0;margin-top:2px}
.ic-add{background:var(--ok)}.ic-chg{background:var(--warn)}.ic-del{background:var(--ink3)}
.tl .ti{display:flex;gap:12px;padding-bottom:16px;position:relative}
.tl .ti:before{content:'';position:absolute;left:5px;top:16px;bottom:0;width:1px;background:var(--line)}
.tl .ti:last-child:before{display:none}
.tl .dt{width:11px;height:11px;border-radius:50%;background:#fff;border:2px solid var(--pri);margin-top:5px;flex-shrink:0;position:relative;z-index:1}
.tl .ti.done .dt{background:var(--pri)}
.wc{display:flex;flex-wrap:wrap;gap:10px 18px;align-items:baseline;padding:10px 0}
.wc span{color:var(--pri);font-weight:600}
.stt{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:var(--ok)}
.stt .d{width:6px;height:6px;border-radius:50%;background:var(--ok)}
.pager{display:flex;gap:6px;justify-content:flex-end;margin-top:14px}
.pager button{min-width:30px;height:30px;border-radius:6px;border:1px solid var(--line);background:#fff;color:var(--ink2);font-size:13px}
.pager button.on{background:var(--pri);color:#fff;border-color:var(--pri)}
.dim p{font-size:13px;color:var(--ink2)}
.mask{position:fixed;inset:0;background:rgba(31,35,41,.45);display:none;place-items:center;z-index:100}
.mask.on{display:grid}
.modal{background:#fff;border-radius:12px;width:520px;max-width:92vw;max-height:86vh;overflow:auto;padding:24px;box-shadow:0 8px 40px rgba(31,35,41,.18)}
.modal h3{font-size:16px;margin-bottom:16px}
.modal .foot{display:flex;justify-content:flex-end;gap:10px;margin-top:20px}
.toast{position:fixed;top:72px;left:50%;transform:translateX(-50%);background:#1f2329;color:#fff;border-radius:8px;padding:10px 20px;display:none;z-index:200;font-size:13px;box-shadow:0 6px 20px rgba(0,0,0,.2)}
.toast.on{display:block}
.page{display:none}.page.on{display:block;animation:fin .3s ease}
.fppanel{display:none}.fppanel.on{display:block;animation:fin .3s ease}
@keyframes fin{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
@media(max-width:1100px){.sidebar{width:60px}.mi span.tx,.cap,.kids,.che,.brand b,.brand span{display:none}.brand{justify-content:center;padding:12px 4px}.mi{justify-content:center}.grid2,.grid3,.tiles{grid-template-columns:1fr 1fr}}
@media(max-width:760px){.sidebar{display:none}.main{padding:0 14px 20px}.grid2,.grid3,.tiles{grid-template-columns:1fr}.search{width:130px}}
'''

_ENHANCED_CSS = '''
body{--tint:#eaf1ff;--mist:#dcebff;--glow:#def6ff;background:radial-gradient(ellipse at 25% 0%,var(--mist),transparent 52%),radial-gradient(ellipse at 95% 25%,var(--glow),transparent 48%),#f5f7fb;background-attachment:fixed}
body.scheme-2{--pri:#7555d9;--pri2:#ac86f9;--pur:#5a74ea;--tint:#f0eaff;--mist:#e5ddff;--glow:#e4eeff;--sh:0 8px 28px rgba(95,75,153,.07)}
body.scheme-3{--pri:#138778;--pri2:#54baa6;--pur:#2c97b0;--tint:#e3f6ef;--mist:#d5f0e7;--glow:#e7f4db;--sh:0 6px 24px rgba(38,102,89,.06)}
.topbar{background:rgba(255,255,255,.93);backdrop-filter:blur(16px)}
.sidebar{width:248px;scrollbar-width:thin;background:rgba(255,255,255,.95)}
.brand{gap:8px;padding:12px 6px}.brand>div:last-child{min-width:0}.brand b{font-size:14px;white-space:normal}.brand span{font-size:8px;letter-spacing:.4px;white-space:normal}
.brand .lg{background:linear-gradient(135deg,var(--pri),var(--pur))}
.mi .tx{white-space:normal;line-height:1.45}.kid{padding:7px 9px;white-space:normal;line-height:1.35}.kid .tx{overflow:visible;flex:1;font-size:12px;overflow-wrap:anywhere}.kid small{display:block;font-size:11px;color:var(--ink3);margin-top:3px}
.mi.on,.kid.on{background:var(--tint)}.kids{margin-left:15px}.kid.on small{color:var(--pri)}
.main{padding:18px 24px 32px}.hero{border:1px solid rgba(255,255,255,.9);border-radius:16px;position:relative;overflow:hidden;padding:20px 24px;background:linear-gradient(115deg,var(--tint),rgba(255,255,255,.82));box-shadow:var(--sh)}
.hero:after{content:'';position:absolute;right:30px;top:-75px;width:240px;height:240px;border:36px solid var(--mist);border-radius:50%;opacity:.42;pointer-events:none}.hero>*{position:relative;z-index:1}.eyebrow{font-size:10px;letter-spacing:1px;color:var(--pri);margin-bottom:4px;font-weight:600}.hero h2{font-size:22px}.hero p{font-size:12px}
.card{border:1px solid rgba(228,234,242,.9);box-shadow:var(--sh);min-width:0;padding:18px;border-radius:14px}.card>.ch{flex-wrap:wrap;gap:6px}.ch h3{font-size:15px}.ch .sub{font-size:11px}.card>.ch>div{flex-wrap:wrap}
.grid2,.grid3,.overview-grid,.workspace-grid,.context-grid{gap:16px;align-items:start}.grid2>*,.grid3>*{min-width:0}.overview-grid,.workspace-grid,.context-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr))}.overview-grid{grid-template-columns:minmax(0,1.25fr) minmax(0,1fr)}.workspace-grid{grid-template-columns:minmax(0,.8fr) minmax(0,1.2fr)}
.tiles{gap:12px;margin-bottom:16px}.card>.tiles{margin-bottom:0}.tile{padding:14px 16px;border:1px solid rgba(255,255,255,.65);min-width:0}.tile .vl{font-size:25px}.tile .lb{font-size:12px}.tile .sb{font-size:11px}.tile .gh{font-size:40px}
.scheme-2 .hero{background:linear-gradient(120deg,#f2edff,#eaf0ff)}.scheme-2 .tile{border-radius:14px}.scheme-2 .context-grid{grid-template-columns:1.1fr .9fr}.scheme-2 .overview-grid.reverse{grid-template-columns:1fr 1.15fr}.scheme-2 .workspace-grid{grid-template-columns:1fr 1fr}
.scheme-3 .hero{border-left:4px solid var(--pri)}.scheme-3 .card>.ch{border-bottom:1px solid #edf3f0;padding-bottom:10px}.scheme-3 .route-grid{grid-template-columns:repeat(3,minmax(0,1fr))}
.scheme-badge{font-size:11px;border:1px solid var(--mist);color:var(--pri);padding:4px 10px;border-radius:999px;background:var(--tint);white-space:nowrap}
.view-nav{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 16px}.view-nav .btn{font-size:12px;padding:7px 12px}
.route-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.route-card{text-align:left;border:1px solid var(--mist);background:linear-gradient(135deg,var(--tint),#fff);border-radius:10px;padding:12px;color:var(--ink);transition:transform .15s,box-shadow .15s}.route-card:hover{transform:translateY(-2px);box-shadow:var(--sh);border-color:var(--pri)}.route-card>span{float:right;color:var(--pri);opacity:.55;font-size:22px;font-weight:700}.route-card>b{display:block;font-size:12px;padding-right:24px}.route-card>small{font-size:10px;color:var(--ink3)}
.relation-map{width:100%;max-height:310px;display:block}.relation-map .edges path{fill:none;stroke:var(--pri);stroke-width:1.5;opacity:.35}.relation-map rect{fill:var(--tint);stroke:var(--mist)}.relation-map text{font-size:12px;fill:var(--ink2)}.relation-map .hub circle{fill:var(--pri)}.relation-map .hub text{fill:#fff;font-weight:600}
.table-scroll{overflow-x:auto;max-width:100%}.tbl{min-width:520px}.tbl td{padding:10px 9px;font-size:12px}.tbl th{white-space:nowrap;padding:9px;font-size:11px}.tbl td small{display:block;color:var(--ink3);font-size:10px;margin-top:4px}.tbl td b{font-weight:500}.tbl .tag{white-space:nowrap}.progress-track{display:block;width:64px;background:#eef2f7;height:5px;border-radius:8px}.progress-track i{display:block;background:linear-gradient(90deg,var(--pri),var(--pri2));height:100%;border-radius:8px}
.tb-row{flex-wrap:wrap;gap:8px}.seg{flex-wrap:wrap}.sg{padding:4px 8px}.btn-pri,.pager button.on{background:var(--pri);border-color:var(--pri)}.bars .v,.hbar .fl{background:linear-gradient(180deg,var(--pri2),var(--pri))}
.action-line{display:flex;align-items:center;gap:10px;padding:11px 0;border-bottom:1px solid #f0f3f7;font-size:12px}.action-line>div{flex:1;min-width:0}.action-line small{display:block;color:var(--ink3);font-size:11px;margin-top:3px}.step-dot{display:grid;place-items:center;background:var(--tint);color:var(--pri);width:30px;height:30px;border-radius:9px;flex-shrink:0;font-size:12px}.action-line:last-child{border-bottom:0}
.resource-tile{display:flex;align-items:center;gap:10px;padding:14px 0;border-bottom:1px solid #f0f3f7}.resource-tile>div{flex:1;min-width:0}.resource-tile b{font-size:12px;display:block}.resource-tile small{font-size:10px;color:var(--ink3)}.file-icon{background:var(--tint);color:var(--pri);border-radius:8px;padding:10px 5px;font-size:10px;font-weight:700}
.frm input,.frm select,.frm textarea{width:100%;min-width:0;padding:9px 10px;border:1px solid var(--line);border-radius:8px;font:inherit;font-size:12px;color:var(--ink2);background:#fff}.frm label{display:grid;gap:6px}.upzone small{display:block;font-size:11px}.upload-results{font-size:12px;color:var(--ink2);overflow-wrap:anywhere}
#search-results{position:absolute;right:80px;top:54px;background:white;border:1px solid var(--line);border-radius:12px;padding:8px;box-shadow:var(--sh);width:340px;max-height:380px;overflow:auto}#search-results button{display:block;border:0;background:transparent;width:100%;text-align:left;padding:9px;color:var(--ink2)}#search-results button:hover{background:var(--tint)}[hidden]{display:none!important}
.menu-toggle{display:none}.empty-results{padding:14px;color:var(--ink3);font-size:12px}.changed{animation:highlight 1.2s ease}@keyframes highlight{from{background:var(--tint)}to{background:transparent}}
.scheme-split{display:grid;grid-template-columns:minmax(0,1.45fr) minmax(0,1fr);gap:16px;align-items:start}.scheme-split>div{min-width:0}.scheme-split .workspace-grid,.scheme-split .grid2{grid-template-columns:1fr}.qc{width:100%;text-align:left;color:var(--ink)}.chatin input{min-width:0}.scheme-split .context-grid{grid-template-columns:1fr;gap:0}.scheme-split .tbl{min-width:460px}
@media(min-width:1450px){.tiles{grid-template-columns:repeat(6,minmax(0,1fr))}.tile .sb{font-size:10px}}
@media(max-width:1200px){.scheme-split{grid-template-columns:1fr}.scheme-split .context-grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}}
@media(max-width:1100px){.menu-toggle{display:grid}.sidebar{width:60px}.sidebar .brand .lg{width:32px}.sidebar.menu-open{position:fixed;left:0;top:60px;width:258px;z-index:70;background:#fff;box-shadow:0 8px 30px #bcc8da}.sidebar.menu-open .tx,.sidebar.menu-open .cap,.sidebar.menu-open .che,.sidebar.menu-open .brand b,.sidebar.menu-open .brand span{display:block}.sidebar.menu-open .kids.open{display:block}.sidebar.menu-open .mi{justify-content:flex-start}.scheme-badge{display:none}.overview-grid,.overview-grid.reverse,.scheme-2 .overview-grid.reverse,.workspace-grid,.scheme-2 .workspace-grid{grid-template-columns:1fr}.scheme-3 .route-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:760px){.topbar{gap:8px;padding:0 10px}.topbar .search{width:130px}.crumb b{font-size:12px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.crumb span{display:none}.sidebar.menu-open{display:block}.main{padding:12px}.context-grid,.scheme-2 .context-grid,.scheme-split .context-grid{grid-template-columns:1fr}.tiles{grid-template-columns:repeat(2,minmax(0,1fr))}.hero{padding:16px}.hero h2{font-size:18px}.hero:after{right:-90px}.route-grid,.scheme-3 .route-grid{grid-template-columns:1fr 1fr}.topbar .icobtn[title]{display:none}#search-results{right:10px;width:min(340px,95vw)}}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
'''

_CSS_V7 = '''
.hero{position:relative}.hero-acts{position:absolute;right:16px;top:20px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;z-index:2}
.pills{display:inline-flex;background:rgba(255,255,255,.88);border:1px solid var(--mist);border-radius:999px;padding:3px;box-shadow:0 2px 8px rgba(31,35,41,.06)}
.pills button{border:none;background:none;padding:6px 14px;font-size:12px;border-radius:999px;color:var(--ink2)}
.pills button.on{background:var(--pri);color:#fff;font-weight:600;box-shadow:0 2px 8px rgba(51,112,255,.35)}
.tile .tr-up{color:#00a870;font-weight:600}.tile .tr-dn{color:#f53f3f;font-weight:600}
.tile{transition:transform .15s,box-shadow .15s}.tile:hover{transform:translateY(-2px);box-shadow:var(--sh)}
.fgrid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}
.fcard{text-align:left;border:1px solid var(--mist);border-radius:12px;padding:14px 16px;background:linear-gradient(135deg,var(--tint),#fff 70%);cursor:pointer;transition:transform .15s,box-shadow .15s,border-color .15s;color:var(--ink);font:inherit;position:relative;min-width:0}
.fcard:hover{transform:translateY(-3px);box-shadow:var(--sh);border-color:var(--pri)}
.fcard>b{display:block;font-size:14px;font-weight:600;margin:2px 44px 10px 0;line-height:1.4}
.fcard .fi{float:right;width:36px;height:36px;border-radius:11px;background:#fff;border:1px solid var(--mist);display:grid;place-items:center;margin-left:8px}
.fcard .fi svg{width:18px;height:18px;stroke:var(--pri);fill:none;stroke-width:1.7}
.fcard .fs{display:flex;gap:18px;flex-wrap:wrap}
.fcard .fs span{font-size:11px;color:var(--ink3)}
.fcard .fs b{font-size:19px;font-weight:700;color:var(--ink);margin-right:2px}
.fcard small{display:block;font-size:11px;color:var(--ink3);margin-top:10px}
.fcard:hover small{color:var(--pri)}
.health{display:flex;gap:24px;align-items:center;flex-wrap:wrap}
.health .hbars{flex:1;min-width:220px}
.health .hbar{grid-template-columns:96px 1fr 48px;margin-bottom:12px}
.note-line{margin-top:12px;background:var(--tint);border-radius:8px;padding:9px 12px;font-size:12px;color:var(--ink2)}
.alert{display:flex;gap:10px;align-items:flex-start;padding:11px 0;border-bottom:1px solid #f0f3f7;font-size:13px}
.alert:last-child{border-bottom:0}
.alert>div{flex:1;min-width:0}
.alert>div:last-child{flex:none;text-align:right}
.alert b{display:block;font-size:13px;font-weight:600}
.alert small{font-size:11px;color:var(--ink3)}
.alert time{display:block;font-size:11px;color:var(--ink3);margin-top:4px}
.alert .dot{width:8px;height:8px;border-radius:50%;margin-top:7px;flex-shrink:0}
.dot-r{background:var(--dan)}.dot-o{background:var(--warn)}.dot-b{background:var(--pri)}.dot-g{background:var(--ok)}
.alert:hover{background:#fafbff;border-radius:8px}
.badge{background:var(--dan);color:#fff;border-radius:999px;font-size:11px;padding:1px 7px;font-weight:600;flex-shrink:0}
.mi .badge{margin-left:6px}
.mod-block{margin-top:22px}
.mb-head{display:flex;align-items:baseline;gap:12px;margin-bottom:10px;flex-wrap:wrap}
.mb-head b{font-size:16px;font-weight:700}
.mb-head span{font-size:12px;color:var(--ink3)}
.ftabs{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 14px}
.ftab{border:1px solid var(--line);background:#fff;border-radius:999px;padding:8px 18px;font-size:13px;color:var(--ink2);transition:all .15s}
.ftab:hover{color:var(--pri);border-color:var(--pri)}
.ftab.on{background:linear-gradient(120deg,var(--pri),var(--pur));border-color:transparent;color:#fff;font-weight:600;box-shadow:0 4px 14px rgba(51,112,255,.3)}
.fphead{display:flex;align-items:baseline;gap:12px;margin:6px 0 12px;flex-wrap:wrap}
.fphead h3{font-size:17px;font-weight:700}
.fphead h3:before{content:'';display:inline-block;width:4px;height:16px;border-radius:2px;background:linear-gradient(180deg,var(--pri),var(--pur));margin-right:9px;vertical-align:-2px}
.fphead span{font-size:12px;color:var(--ink3)}
.vpanel{display:none}.vpanel.on{display:block;animation:fin .3s ease}
.fppanel{animation:fin .3s ease}
.fpblock{display:none;min-width:0}.fpblock.on{display:block;animation:fin .3s ease}
.modsum{display:none}.modsum.on{display:block;animation:fin .3s ease}
.mod-block{display:none}.mod-block.on{display:block;animation:fin .3s ease}
.fbbar{position:relative;margin-bottom:14px}
.fb-big{display:flex;align-items:center;gap:12px;max-width:460px;width:100%;padding:13px 16px;
 border:1px solid var(--mist);border-radius:14px;cursor:pointer;text-align:left;font-size:13px;color:var(--ink2);
 background:linear-gradient(120deg,var(--tint),#fff);box-shadow:var(--sh);transition:border-color .15s}
.fb-big:hover{border-color:var(--pri)}
.fb-ico{width:36px;height:36px;border-radius:11px;background:linear-gradient(135deg,var(--pri),var(--pur));
 color:#fff;display:grid;place-items:center;font-size:16px;flex-shrink:0}
.fb-txt{min-width:0;line-height:1.5}
.fb-cur{font-size:15px;color:var(--ink);margin:0 4px}
.fb-meta{display:block;font-size:11px;color:var(--ink3)}
.fb-caret{margin-left:auto;color:var(--pri);font-size:12px;white-space:nowrap;flex-shrink:0}
.fb-menu{display:none;position:absolute;z-index:70;top:calc(100% + 6px);left:0;width:100%;max-width:640px;
 max-height:300px;overflow:auto;background:#fff;border:1px solid var(--line);border-radius:13px;
 box-shadow:0 14px 40px rgba(31,35,41,.18);padding:8px;grid-template-columns:repeat(2,minmax(0,1fr));gap:6px}
.fb-menu.on{display:grid}
.fb-item{text-align:left;padding:9px 12px;border:1px solid #eef1f6;border-radius:9px;background:#fafbfe;
 font-size:13px;cursor:pointer;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--ink2)}
.fb-item:hover{border-color:var(--pri);color:var(--pri)}
.fb-item.on{background:linear-gradient(120deg,var(--pri),var(--pur));color:#fff;border-color:transparent;font-weight:600}
@media(max-width:760px){.fb-menu{grid-template-columns:1fr}}
@media(max-width:1200px){.fgrid{grid-template-columns:repeat(2,minmax(0,1fr))}.hero-acts{position:static;margin-top:10px}}
@media(max-width:760px){.fgrid{grid-template-columns:1fr}.ftabs{gap:6px}.ftab{padding:6px 12px;font-size:12px}.hero-acts .btn{display:none}}
'''

_CSS_V8 = '''
/* v8 专属版式：不同功能用完全不同的骨架 */
.ws2{display:grid;grid-template-columns:minmax(0,1fr) 350px;gap:16px;align-items:stretch}
.ws2>.card{margin:0}
.wchip{display:inline-block;border-radius:6px;padding:1px 8px;font-size:11px;font-weight:600}
.wchip-b{background:var(--tint);color:var(--pri)}.wchip-g{background:#e3f7ee;color:#0a915c}
.wchip-p{background:#efeaff;color:var(--pur)}.wchip-o{background:#fff3e0;color:#c87800}
/* 图谱画布 */
.gcanvas{position:relative;background:#fff;border:1px solid var(--line);border-radius:14px;
 min-height:400px;overflow:hidden;box-shadow:var(--sh);
 background-image:radial-gradient(var(--mist) 1.2px,transparent 1.2px);background-size:22px 22px}
.gedges{position:absolute;inset:0;width:100%;height:100%}
.gedges line{stroke:var(--pri);stroke-width:.4;opacity:.45}
.ghub{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:104px;height:104px;
 border-radius:50%;background:linear-gradient(135deg,var(--pri),var(--pur));color:#fff;
 display:grid;place-items:center;text-align:center;box-shadow:0 10px 26px rgba(51,112,255,.35);z-index:2}
.ghub b{font-size:14px}.ghub small{font-size:10px;opacity:.85;display:block;margin-top:2px}
.gnode{position:absolute;transform:translate(-50%,-50%);background:#fff;border:2px solid var(--c,var(--pri));
 border-radius:999px;padding:7px 14px;text-align:center;cursor:pointer;box-shadow:var(--sh);
 transition:transform .15s;color:var(--ink);font:inherit;line-height:1.3;z-index:3}
.gnode:hover{transform:translate(-50%,-50%) scale(1.08)}
.gnode b{display:block;font-size:12px;color:var(--c,var(--pri))}
.gnode small{font-size:10px;color:var(--ink3)}
.gdetail{background:#fff;border:1px solid var(--line);border-radius:14px;padding:16px 18px;box-shadow:var(--sh)}
.gdetail h3{font-size:17px;margin:6px 0 2px}.gdetail>small{color:var(--ink3);font-size:12px}
.gdetail p{font-size:12px;color:var(--ink2);margin:6px 0 2px}
.dl{font-size:11px;color:var(--ink3);margin:12px 0 4px}
.mbar{height:8px;border-radius:99px;background:#eef1f6;overflow:hidden}
.mbar i{display:block;height:100%;border-radius:99px;background:linear-gradient(90deg,var(--pri),var(--pur))}
.gdetail em{font-style:normal;font-size:11px;color:var(--ink3)}
.gstat-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px}
.gstat-grid>div{background:var(--bg);border-radius:10px;padding:8px 12px}
.gstat-grid span{display:block;font-size:11px;color:var(--ink3)}
.gstat-grid b{font-size:16px}
.gbtns{display:flex;gap:8px;margin-top:14px;flex-wrap:wrap}
.gfoot{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-top:16px}
.gmini{background:#fff;border:1px solid var(--line);border-radius:12px;text-align:center;padding:10px 4px;box-shadow:var(--sh)}
.gmini b{display:block;font-size:17px}.gmini span{font-size:11px;color:var(--ink3)}
/* 沉浸大屏 */
.ws-live{display:grid;grid-template-columns:minmax(0,1fr) 330px;gap:16px;align-items:start}
.stage{background:#fff;border:1px solid var(--line);border-radius:14px;overflow:hidden;box-shadow:var(--sh)}
.stage-top{display:flex;align-items:center;gap:12px;padding:10px 16px;border-bottom:1px solid var(--line);font-size:12px}
.stage-top b{color:var(--ink2);font-weight:500}
.liver{color:var(--dan);font-weight:600}.on-air{margin-left:auto;background:var(--tint);color:var(--pri);border-radius:99px;padding:3px 10px;font-weight:600}
.slide{padding:26px 30px;background:linear-gradient(160deg,#fbfcff,#f4f7ff)}
.slide-no{font-size:40px;font-weight:800;color:var(--mist);line-height:1}
.slide h3{font-size:22px;margin:2px 0 4px}.slide>small{color:var(--ink3);font-size:12px}
.slide ul{margin:14px 0 0;list-style:none;display:grid;gap:8px}
.slide li{font-size:13px;color:var(--ink2);padding-left:16px;position:relative}
.slide li:before{content:'';position:absolute;left:0;top:8px;width:6px;height:6px;border-radius:50%;background:var(--pri)}
.stage-code{margin:16px 0 0;background:#1e2430;color:#d8e0ee;border-radius:12px;padding:16px 18px;
 font:12px/1.7 Consolas,monospace;overflow:auto}
.stage-foot{display:flex;justify-content:space-between;padding:10px 16px;border-top:1px solid var(--line);font-size:11px;color:var(--ink3)}
.rail{display:grid;gap:14px}
.rcard{background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px 16px;box-shadow:var(--sh)}
.rcard h4{font-size:14px;margin-bottom:10px;display:flex;align-items:baseline;gap:8px}
.rcard h4 em{font-style:normal;font-size:11px;color:var(--ok);font-weight:400;margin-left:auto}
.attend{display:flex;align-items:center;gap:14px}
.attend .ring{width:92px;flex-shrink:0}
.attend-n{display:grid;gap:6px;font-size:12px;color:var(--ink3)}
.attend-n b{color:var(--ink);font-size:15px;margin-right:4px}
.rbtns{display:flex;gap:8px;margin-top:12px}
.stu-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:8px}
.stu{position:relative;background:var(--bg);border-radius:9px;padding:10px 2px 6px;text-align:center;font-size:10px;color:var(--ink2)}
.sd{display:inline-block;width:6px;height:6px;border-radius:50%;position:absolute;right:5px;top:5px}
.sd-g{background:var(--ok)}.sd-o{background:var(--warn)}.sd-b{background:var(--pri)}.sd-r{background:var(--dan)}
.legend{display:flex;gap:12px;margin-top:10px;font-size:11px;color:var(--ink3)}
.legend .sd{position:static}
.mini2 .minis{grid-template-columns:repeat(2,1fr)}
/* 页签大表 */
.board-tabs{display:flex;gap:22px;border-bottom:1px solid var(--line);margin-bottom:16px;align-items:center}
.bt{border:none;background:none;font:inherit;padding:10px 2px;font-size:14px;color:var(--ink2);cursor:pointer;border-bottom:2px solid transparent}
.bt.on{color:var(--pri);font-weight:600;border-bottom-color:var(--pri)}
.bt-right{margin-left:auto;font-size:12px;color:var(--ink3)}
/* 生成卡片流 */
.doclist{display:grid;gap:10px}
.docitem{display:flex;align-items:center;gap:10px;border:1px solid var(--line);border-radius:11px;padding:10px 12px}
.docitem:hover{border-color:var(--pri)}
.dstat{font-size:11px;border-radius:6px;padding:2px 8px;flex-shrink:0;font-weight:600}
.dstat-ok{background:#e3f7ee;color:#0a915c}.dstat-warn{background:#fff3e0;color:#c87800}.dstat-mute{background:#f0f1f5;color:var(--ink3)}
.dmain{flex:1;min-width:0}.dmain b{display:block;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dmain small{font-size:11px;color:var(--ink3)}
.aistream{display:grid;gap:12px}
.aicard{border:1px solid var(--mist);border-radius:12px;padding:12px 14px;background:linear-gradient(180deg,#fff,#fbfcff)}
.aicard:hover{border-color:var(--pri)}
.ai-top{display:flex;align-items:center;gap:10px;font-size:11px;color:var(--ink3);margin-bottom:6px}
.ai-top em{font-style:normal;margin-left:auto;font-size:11px;font-weight:600}
.ai-top .okg{color:var(--ok)}.ai-top .warno{color:var(--warn)}
.aicard b{display:block;font-size:13px;line-height:1.5}
.aicard small{font-size:11px;color:var(--ink3)}
.ai-btns{display:flex;gap:8px;margin-top:9px;flex-wrap:wrap}
/* 三栏对话 */
.ws-chat{display:grid;grid-template-columns:230px minmax(0,1fr) 290px;gap:14px;align-items:start}
.chat-l,.chat-r{display:grid;gap:10px;align-content:start;background:#fff;border:1px solid var(--line);
 border-radius:14px;padding:14px;box-shadow:var(--sh)}
.cl-cap{font-size:11px;color:var(--ink3);margin:6px 0 2px}
.role{display:flex;gap:10px;align-items:center;border-radius:11px;padding:8px 10px;cursor:pointer}
.role.on{background:var(--tint)}
.ric{width:30px;height:30px;border-radius:9px;background:#fff;border:1px solid var(--mist);display:grid;place-items:center;flex-shrink:0}
.ric svg{width:15px;height:15px;stroke:var(--pri);fill:none;stroke-width:1.7}
.role b{font-size:13px;display:block;line-height:1.3}.role small{font-size:10px;color:var(--ink3)}
.rec{font-size:12px;color:var(--ink2);padding:6px 10px;display:flex;gap:8px;align-items:center;cursor:pointer;border-radius:8px}
.rec:hover{background:var(--bg)}.rec i{width:6px;height:6px;border-radius:50%;background:var(--pri);flex-shrink:0}
.chat-c{background:#fff;border:1px solid var(--line);border-radius:14px;box-shadow:var(--sh);
 display:flex;flex-direction:column;min-height:520px;padding:0 16px 16px}
.cc-head{display:flex;align-items:center;gap:10px;padding:14px 0;border-bottom:1px solid var(--line);margin-bottom:12px}
.cc-head b{font-size:15px}.cc-head span:not(.sp){font-size:11px;color:var(--ink3)}
.model{background:var(--bg);border-radius:8px;padding:4px 10px;font-size:12px;color:var(--ink2)}
.chat-c .chat{flex:1;overflow:auto}
.qps{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}
.qp{border:1px solid var(--mist);background:#fff;border-radius:99px;padding:5px 12px;font-size:12px;color:var(--ink2);cursor:pointer}
.qp:hover{color:var(--pri);border-color:var(--pri)}
.ctx{background:var(--bg);border-radius:12px;padding:12px}
.ctx-h{display:flex;gap:10px;align-items:center;margin-top:6px}
.ctx-lg{width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,var(--pri),var(--pur));display:grid;place-items:center;flex-shrink:0}
.ctx-lg svg{width:17px;height:17px;stroke:#fff;fill:none;stroke-width:1.8}
.ctx-h b{font-size:13px;display:block}.ctx-h small{font-size:11px;color:var(--ink3)}
.ctx .kv{margin-top:8px}
.hit{border-left:3px solid var(--pri);background:var(--bg);border-radius:0 10px 10px 0;padding:8px 12px;margin-bottom:8px}
.hit b{font-size:12px;display:block}.hit small{font-size:11px;color:var(--ink3)}
.sw-row{display:flex;align-items:center;justify-content:space-between;font-size:13px;padding:7px 2px}
.sw{width:36px;height:20px;border-radius:99px;background:#dde1ea;position:relative;cursor:pointer;transition:background .15s}
.sw.on{background:var(--pri)}
.sw:after{content:'';position:absolute;top:2px;left:2px;width:16px;height:16px;border-radius:50%;background:#fff;transition:left .15s;box-shadow:0 1px 3px rgba(0,0,0,.2)}
.sw.on:after{left:18px}
@media(max-width:1280px){.ws2,.ws-live{grid-template-columns:1fr}.ws-chat{grid-template-columns:1fr}
 .chat-l,.chat-r{display:none}.gfoot{grid-template-columns:repeat(3,1fr)}}
'''

_CSS_V9 = '''
/* v9 精修：质感细节统一升级 */
::-webkit-scrollbar{width:9px;height:9px}
::-webkit-scrollbar-thumb{background:#d6dbe6;border-radius:99px;border:2px solid transparent;background-clip:content-box}
::-webkit-scrollbar-thumb:hover{background:#bcc4d4;background-clip:content-box}
::-webkit-scrollbar-track{background:transparent}
.card{transition:box-shadow .2s,border-color .2s}
.card:hover{box-shadow:0 6px 24px rgba(31,35,41,.09)}
.ch h3{display:flex;align-items:center;gap:9px}
.ch h3:before{content:'';width:4px;height:15px;border-radius:2px;flex-shrink:0;
 background:linear-gradient(180deg,var(--pri),var(--pur))}
.btn{border-radius:9px;transition:all .15s}
.btn-pri{background:linear-gradient(120deg,var(--pri),var(--pur));border-color:transparent;
 box-shadow:0 3px 10px rgba(51,112,255,.26)}
.btn-pri:hover{box-shadow:0 5px 16px rgba(51,112,255,.36);transform:translateY(-1px)}
.table-scroll{border:1px solid #eef1f6;border-radius:12px;overflow:hidden}
.tbl th{background:linear-gradient(180deg,#fafcff,#f4f7fd);font-weight:600;color:var(--ink2)}
.tbl td{border-bottom-color:#f4f6fa}
.tbl tr:last-child td{border-bottom:none}
.tbl tr:hover td{background:linear-gradient(90deg,var(--tint),#fafbff 60%)}
.tag{border-radius:999px;font-weight:500}
.pager button{border-radius:8px;transition:all .15s}
.pager button.on{background:linear-gradient(120deg,var(--pri),var(--pur));border-color:transparent}
.mi{transition:all .15s;border-radius:9px}
.mi.on{background:linear-gradient(92deg,var(--tint),rgba(255,255,255,.4))}
.kid{transition:all .15s;border-radius:8px}
.kid.on{background:linear-gradient(92deg,var(--tint),rgba(255,255,255,.4))}
.bub{border-radius:12px}
.chatin input{border-radius:10px}
.qc{border-radius:11px;transition:all .15s}
.qc:hover{transform:translateY(-1px);box-shadow:var(--sh)}
.upzone{border-radius:13px}
.modal{border-radius:16px}
.fcard .fi{transition:transform .2s}
.fcard:hover .fi{transform:scale(1.1) rotate(-4deg)}
.tile{border-radius:13px}
.gnode{backdrop-filter:blur(4px)}
.rcard h4:before{content:'';width:4px;height:13px;border-radius:2px;background:linear-gradient(180deg,var(--pri),var(--pur))}
.rcard h4{display:flex;align-items:center;gap:8px}
.cl-cap:first-child{margin-top:0}
.seg{border-radius:9px;padding:3px}
.sg{border-radius:7px;transition:all .15s}
'''

_JS = r'''
const $=(s,r=document)=>r.querySelector(s), $$=(s,r=document)=>Array.from(r.querySelectorAll(s));
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const UI={onSave:null,toast(m){const t=$('#toast');t.textContent=m;t.classList.add('on');
clearTimeout(this._tt);this._tt=setTimeout(()=>t.classList.remove('on'),2600);},
modal(on,title,html){this.onSave=null;if(title)$('#mtitle').textContent=title;
if(html!==undefined)$('#mbody').innerHTML=html;$('#mask').classList.toggle('on',on);}};
const FORM='<div class="frm"><label>名称<input name="title" required></label><label>负责人<input name="owner"></label><label>说明<textarea name="note" rows="3"></textarea></label></div>';
function navTo(pg,ft){const el=(ft?$('.kid[data-ft="'+ft+'"]')||$('.mi[data-pg="'+pg+'"]')
:$('.mi[data-pg="'+pg+'"],.kid[data-pg="'+pg+'"]'));if(el)go(pg,el,ft);}
function fbsel(pan,blk){const items=$$('[data-fb]',pan);
items.forEach(b=>b.classList.toggle('on',document.getElementById(b.dataset.fb)===blk));
const bar=$('.fbbar',pan);if(bar&&blk){const h=$('.fphead h3',blk),cur=$('.fb-cur',bar),idx=$('.fb-idx',bar);
if(h&&cur)cur.textContent=h.textContent;
if(idx)idx.textContent='第'+(items.indexOf($('[data-fb="'+blk.id+'"]',pan))+1)+'项';
$('.fb-menu',bar).classList.remove('on');}}
function ftab(pg,id){const p=document.getElementById(pg);if(!p)return;
const sum=$('.modsum',p),mb=$('.mod-block',p);
if(!id){$$('.ftab',p).forEach((t,j)=>t.classList.toggle('on',j===0));
$$('.fppanel',p).forEach(x=>x.classList.remove('on'));
if(sum)sum.classList.add('on');if(mb)mb.classList.remove('on');
window.scrollTo({top:0,behavior:'smooth'});return;}
if(sum)sum.classList.remove('on');if(mb)mb.classList.add('on');
$$('.ftab',p).forEach(t=>t.classList.toggle('on',t.dataset.ft===id));
$$('.fppanel',p).forEach(x=>{const on=x.id===id;x.classList.toggle('on',on);
if(on){$$('.fpblock',x).forEach((bl,j)=>{bl.classList.toggle('on',j===0);
$$('.vpanel',bl).forEach((v,k)=>v.classList.toggle('on',k===0));
$$('.view-nav .btn',bl).forEach((b,k)=>{b.classList.toggle('btn-pri',k===0);b.classList.toggle('btn-gh',k!==0);});});
const first=$('.fpblock',x);if(first)fbsel(x,first);}});
window.scrollTo({top:0,behavior:'smooth'});}
function go(pg,el,ft){if(!document.getElementById(pg))return;
$$('.page').forEach(p=>p.classList.toggle('on',p.id===pg));$$('.mi,.kid').forEach(m=>m.classList.remove('on'));
if(el){el.classList.add('on');if(el.classList.contains('kid')){const ks=el.parentElement;ks.classList.add('open');ks.previousElementSibling?.classList.add('open','on');}
$('#pt').textContent=el.dataset.t;$('#pb').textContent=el.dataset.b;}
if(ft){$('#toast').classList.remove('on');$('.sidebar').classList.remove('menu-open');$('#search-results').hidden=true;ftab(pg,ft);return;}
const cur=document.getElementById(pg);if(cur&&$('.modsum',cur)&&!$('.modsum',cur).classList.contains('on'))ftab(pg,'');
$('#toast').classList.remove('on');$('.sidebar').classList.remove('menu-open');$('#search-results').hidden=true;window.scrollTo(0,0);}
$$('.mi,.kid').forEach(el=>{el.tabIndex=0;el.setAttribute('role','button');});
$('.menu-toggle').onclick=()=>$('.sidebar').classList.toggle('menu-open');
function filter(card){if(!card)return;const q=($('.search input',card)?.value||'').trim().toLowerCase();
const status=$('.sg.on',card)?.textContent||'全部';let shown=0;
$$('.tbl tr',card).filter(r=>$('td',r)).forEach(row=>{row.hidden=!(row.textContent.toLowerCase().includes(q)&&(status==='全部'||$('.tag',row)?.textContent===status));if(!row.hidden)shown++;});
let hint=$('.empty-results',card);if(!hint){hint=document.createElement('div');hint.className='empty-results';card.appendChild(hint);}hint.textContent=shown?'当前显示 '+shown+' 条':'没有符合条件的记录';}
function download(text,type,filename){const url=URL.createObjectURL(new Blob([text],{type}));const a=document.createElement('a');a.href=url;a.download=filename;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);}
function exportCard(card){const table=$('.tbl',card);
if(table){const quote=s=>'"'+(/^[=+@\-]/.test(s)?"'":'')+s.replace(/"/g,'""')+'"';
const csv=$$('tr',table).filter(r=>!r.hidden).map(r=>$$('th,td',r).map(c=>quote(c.innerText.replace(/\n/g,' '))).join(',')).join('\r\n');download('\ufeff'+csv,'text/csv;charset=utf-8','业务记录.csv');}
else download(card.innerText,'text/plain;charset=utf-8','业务成果.txt');UI.toast('已下载当前内容');}
const search=$('.topbar .search input');search.addEventListener('input',()=>{const box=$('#search-results'),q=search.value.trim().toLowerCase();box.replaceChildren();box.hidden=!q;if(!q)return;
const hits=$$('.mi[data-pg],.kid[data-pg]').filter(n=>n.dataset.t.toLowerCase().includes(q)).slice(0,30);
if(!hits.length){box.textContent='未找到对应业务';return;}hits.forEach(n=>{const b=document.createElement('button');b.textContent=n.dataset.t;b.dataset.pg=n.dataset.pg;box.appendChild(b);});});
document.addEventListener('input',e=>{if(e.target.matches('.card .search input'))filter(e.target.closest('.card'));});
document.addEventListener('keydown',e=>{if(e.key==='Escape'){UI.modal(false);$('.sidebar').classList.remove('menu-open');$('#search-results').hidden=true;}
if(e.key==='Enter'&&e.target.matches('[data-chatin]'))e.target.parentElement.querySelector('[data-act="send"]').click();
if((e.key==='Enter'||e.key===' ')&&e.target.matches('.mi,.kid')){e.preventDefault();e.target.click();}
if(e.key==='Enter'&&e.target===search)$('#search-results button')?.click();});
$('#modal-save').onclick=()=>{if(UI.onSave){if(UI.onSave()===false)return;}UI.modal(false);};
$('#mask').addEventListener('click',e=>{if(e.target===$('#mask'))UI.modal(false);});
document.addEventListener('click',e=>{
if(!e.target.closest('.fbbar'))$$('.fb-menu.on').forEach(m=>m.classList.remove('on'));
const fbb=e.target.closest('.fb-big');if(fbb){fbb.parentElement.querySelector('.fb-menu').classList.toggle('on');return;}
const che=e.target.closest('.che');if(che){const g=che.closest('.grp');g.classList.toggle('open');g.nextElementSibling?.classList.toggle('open');return;}
const pl=e.target.closest('.pills button');if(pl){$$('.pills button',pl.parentElement).forEach(x=>x.classList.remove('on'));pl.classList.add('on');UI.toast('已切换时间范围（本地演练展示）');return;}
const vt=e.target.closest('[data-vt]');if(vt){const panel=document.getElementById(vt.dataset.vt);
if(panel){const box=panel.closest('.fpblock');$$('.vpanel',box).forEach(x=>x.classList.toggle('on',x===panel));
$$('.view-nav .btn',box).forEach(b=>{const on=b===vt;b.classList.toggle('btn-pri',on);b.classList.toggle('btn-gh',!on);});}return;}
const fb=e.target.closest('[data-fb]');if(fb){const blk=document.getElementById(fb.dataset.fb);
if(blk){const pan=blk.closest('.fppanel');$$('.fpblock',pan).forEach(x=>x.classList.toggle('on',x===blk));
fbsel(pan,blk);window.scrollTo({top:0,behavior:'smooth'});}return;}
const nav=e.target.closest('[data-pg]');if(nav){navTo(nav.dataset.pg,nav.dataset.ft);return;}
const sw=e.target.closest('.sw');if(sw){sw.classList.toggle('on');UI.toast(sw.classList.contains('on')?'已开启（本地演练）':'已关闭（本地演练）');return;}
const bt=e.target.closest('.bt');if(bt){$$('.bt',bt.parentElement).forEach(x=>x.classList.remove('on'));bt.classList.add('on');UI.toast('已切换到「'+bt.textContent+'」（本地演练展示）');return;}
const gn=e.target.closest('.gnode');if(gn){$$('.gnode').forEach(x=>x.style.outline='');gn.style.outline='3px solid var(--pri)';UI.toast('已选中节点：'+gn.querySelector('b').textContent);return;}
const sg=e.target.closest('.sg');if(sg){$$('.sg',sg.parentElement).forEach(x=>x.classList.remove('on'));sg.classList.add('on');filter(sg.closest('.card'));return;}
const b=e.target.closest('[data-act],.upzone');if(!b)return;
const page=b.closest('.page')||$('.page.on'),card=b.closest('.card')||page;
const n=b.dataset.name||$('h2',page)?.textContent||'记录',a=b.dataset.act||'upload';
if(a==='new'||a==='edit'){
UI.modal(true,(a==='new'?'新建':'编辑')+' · '+n,FORM);const row=b.closest('tr');
$('#mbody [name="title"]').value=row?($('td b',row)?.textContent||n):n;
$('#mbody [name="owner"]').value=b.dataset.owner||'张伟';
UI.onSave=()=>{const title=$('#mbody [name="title"]').value.trim(),owner=$('#mbody [name="owner"]').value.trim();if(!title){UI.toast('请填写名称');return false;}
let target=row;const table=$('.tbl',card);
if(!target&&a==='new'&&table){const template=$$('tr',table).find(r=>$('td',r));if(template){target=template.cloneNode(true);target.hidden=false;if(table.rows[0].cells[0].textContent==='序号')target.cells[0].textContent=String(table.rows.length).padStart(2,'0');table.tBodies[0].appendChild(target);}}
if(target){const nameColumn=table?.rows[0].cells[0].textContent==='姓名'?0:1;const c=target.cells[nameColumn]||target.cells[0];const label=$('b',c)||c;label.textContent=title;$$('[data-act]',target).forEach(x=>{x.dataset.name=title;x.dataset.owner=owner;});if(target.cells.length>=6)target.cells[2].textContent=owner;target.classList.add('changed');if($('.search input',card))filter(card);}
else{const line=b.closest('.action-line');if(line){$('b',line).textContent=title;$('small',line).textContent=owner+' · '+$('#mbody [name="note"]').value;}
else{let note=$('.local-change',card);if(!note){note=document.createElement('p');note.className='local-change';card.appendChild(note);}note.textContent=title+' · '+owner;}}
UI.toast('已保存到当前页面，刷新后复原');};
}
else if(a==='view'){const row=b.closest('tr');UI.modal(true,n+' · 详情','<div class="kv">'+
(row?$$('td',row).slice(0,-1).map(c=>'<p>'+esc(c.innerText)+'</p>').join(''):'<p>'+esc(n)+'</p>')+'</div>');}
else if(a==='export')exportCard(card);
else if(a==='exportchat'){download($('[data-chatbox]',page)?.innerText||'','text/plain;charset=utf-8','会话记录.txt');UI.toast('会话记录已下载');}
else if(a==='ask'){const inp=$('[data-chatin]',page);if(inp){inp.value=b.dataset.prompt;$('[data-act="send"]',page).click();}}
else if(a==='query'){if($('.tbl',card))filter(card);else UI.modal(true,n+' · 业务摘要','<div class="kv">'+esc(page.innerText.slice(0,1800)).replace(/\n/g,'<br>')+'</div>');}
else if(a==='compose'||a==='gen'){
const form=b.closest('.frm'),preview=$('.result-preview',page);if(preview&&form){const selections=$$('select',form).map(s=>s.value);const goal=$('textarea',form)?.value||'';
preview.innerHTML='<div class="note"><b>'+esc(selections[0])+' · 工作成果</b><p>'+esc(goal)+'</p><p>执行周期：'+esc(selections[1])+'</p><p>重点维度：'+esc(selections[2])+'</p><p>下一步：确认责任人、核对交付清单并提交复核。</p></div>';preview.classList.add('changed');UI.toast('成果预览已更新');}
else UI.toast('请在工作配置页选择对象后生成');}
else if(a==='calc'){card.dataset.checked=String(Number(card.dataset.checked||0)+1);let result=$('.check-result',card);if(!result){result=document.createElement('p');result.className='check-result';card.appendChild(result);}result.textContent='已完成第 '+card.dataset.checked+' 次本地核验，当前记录 '+$$('.tbl tr',page).filter(r=>$('td',r)).length+' 条。';UI.toast('核验结果已更新');}
else if(a==='upload'){const input=document.createElement('input');input.type='file';input.multiple=true;input.onchange=()=>{let result=$('.upload-results',card);if(!result){result=document.createElement('div');result.className='upload-results';card.appendChild(result);}result.replaceChildren();Array.from(input.files).slice(0,10).forEach(f=>{const p=document.createElement('p');p.textContent=f.name+' · '+Math.ceil(f.size/1024)+' KB · 已选择';result.appendChild(p);});UI.toast('已读取文件信息，未上传到服务器');};input.click();}
else if(a==='copy'){const text=$('.result-preview',page)?.innerText||card.innerText;if(navigator.clipboard)navigator.clipboard.writeText(text).then(()=>UI.toast('内容已复制')).catch(()=>UI.toast('浏览器未允许复制，请选中文字手动复制'));else UI.toast('请选中文字手动复制');}
else if(a==='tpl')download('\ufeff名称,负责人,说明\r\n','text/csv;charset=utf-8','填写模板.csv');
else if(a==='save')UI.toast('已保留当前页面输入，刷新后复原');
else if(a==='ins'){UI.toast('已插入到当前位置（本地演练，刷新后复原）');}
else if(a==='feedback'){UI.toast('反馈已记录，将在下一轮生成时采纳');}
else if(a==='reset'){$$('input,textarea',card).forEach(x=>x.value=x.defaultValue);UI.toast('已恢复初始输入');}
else if(a==='send'){const inp=$('[data-chatin]',page),box=$('[data-chatbox]',page),v=inp?.value.trim();if(!v||!box){UI.toast('请先输入问题');return;}
box.insertAdjacentHTML('beforeend','<div class="msg me"><div class="bub">'+esc(v)+'</div><div class="av av-me">我</div></div>');inp.value='';
const actions=$$('.action-line b',page).map(x=>x.textContent).slice(0,3).join('、');
setTimeout(()=>{box.insertAdjacentHTML('beforeend','<div class="msg"><div class="av av-ai">AI</div><div class="bub">当前业务：'+esc($('h2',page).textContent)+'。建议先核对 '+esc(actions||'对象、责任人和交付清单')+'。这是基于本页数据的本地演练回复。</div></div>');box.scrollTop=box.scrollHeight;},300);box.scrollTop=box.scrollHeight;}
});
'''


def build_html(mods, variant=1):
    """同一内容最多三套确定性方案；计次仅由 generate_demo 成功保存时处理。"""
    if variant not in (1, 2, 3):
        raise ValueError('方案编号必须为 1、2 或 3。')
    if not mods or not any(fps for _, fps in mods):
        raise ValueError('业务内容不能为空。')
    mods = _fit_nav(mods)  # 导航容量规则：大导航≤10条、小导航≤30条，超出合并共页
    plan = _plan_kinds(mods)  # 版式配额：同版式最多3页，超出自动轮换
    menu = ['<div class="brand"><div class="lg">%s</div><div><b>智慧业务管理中心</b>'
            '<span>BUSINESS MANAGEMENT HUB</span></div></div>'
            '<div class="cap">总览</div>' % _ICO[5],
            '<div class="mi on" data-pg="pg-work" data-t="工作台" data-b="总览 / 工作台">%s'
            '<span class="tx">工作台</span></div>' % _ICO[0],
            '<div class="cap">业务模块</div>']
    pages = [_work_page(mods, variant)]
    for mid, (name, groups) in enumerate(mods):
        nm = _short_name(name)
        badge = _seed(name + ' '.join(f for gf, _ in groups for f in gf)) % 3
        menu.append('<div class="mi grp open" data-pg="pg-m%d" data-t="%s" data-b="业务模块 / %s">%s'
                    '<span class="tx">%s</span>%s<span class="che">›</span></div>'
                    % (mid, _esc(nm), _esc(nm), _ICO[(mid + 1) % 8], _esc(nm),
                       '<span class="badge">%d</span>' % (2 + badge * 2) if badge else ''))
        children = []
        for i, (gfps, gname) in enumerate(groups):
            views = sum(len(_feature_views(f)) for f in gfps)
            children.append('<div class="kid" role="button" tabindex="0" data-pg="pg-m%d" '
                            'data-ft="ft-%d-%d" data-t="%s" data-b="业务模块 / %s" title="%s">'
                            '<i>%d</i><span class="tx">%s<small>%d 个业务环节%s</small></span></div>' %
                            (mid, mid, i, _esc(gname), _esc(nm), _esc(gname),
                             len(children) + 1, _esc(gname), views,
                             ' · 含%d项功能' % len(gfps) if len(gfps) > 1 else ''))
        menu.append('<div class="kids open">%s</div>' % ''.join(children))
        pages.append(_module_page(mid, name, groups, variant, plan))
    menu.append('<div class="cap">系统</div>')
    menu.append('<div class="mi" data-pg="pg-sys" data-t="系统管理" data-b="系统 / 系统管理">%s'
                '<span class="tx">系统管理</span></div>' % _ICO[6])
    menu.append('<div class="mi" data-pg="pg-log" data-t="操作日志" data-b="系统 / 操作日志">%s'
                '<span class="tx">操作日志</span></div>' % _ICO[7])
    pages.append(_sys_page())
    pages.append(_log_page(mods))
    pages[-1] = pages[-1].replace('</section>', _card('审计说明', _note([
        '当前日志用于本地交互演练，不连接实际账号。', '下载内容取自当前页面，筛选后可重新导出。',
        '重要操作请先核对业务对象与责任人。'])) + '</section>')
    return '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>业务管理系统</title>
<style>%s</style>
</head>
<body class="scheme-%d" data-engine="v6" data-variant="%d">
<div class="topbar"><button class="icobtn menu-toggle" aria-label="打开导航">☰</button><div class="crumb"><b id="pt">工作台</b><span id="pb">总览 / 工作台</span></div>
<div class="spacer"></div><span class="scheme-badge">%s</span>
<div class="search"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg><input placeholder="搜索功能、记录…"></div>
<button class="icobtn" title="消息" onclick="UI.toast('暂无新消息')"><svg viewBox="0 0 24 24"><path d="M6 9a6 6 0 1 1 12 0c0 5 2 6 2 6H4s2-1 2-6"/><path d="M10 20a2 2 0 0 0 4 0"/></svg></button>
<button class="icobtn" title="帮助" onclick="UI.toast('点击左侧业务导航；表格支持检索、编辑和导出。数据保存在当前页面内存中，刷新后复原。')">?</button><div id="search-results" hidden></div></div>
<div class="layout">
<div class="sidebar">%s</div>
<div class="main">%s</div>
</div>
<div class="mask" id="mask"><div class="modal"><h3 id="mtitle">标题</h3><div id="mbody"></div>
<div class="foot"><button class="btn btn-gh" onclick="UI.modal(false)">取消</button>
<button class="btn btn-pri" id="modal-save">确定</button></div></div></div>
<div class="toast" id="toast"></div>
<script>%s</script>
</body>
</html>''' % (_CSS + _ENHANCED_CSS + _CSS_V7 + _CSS_V8 + _CSS_V9, variant, variant, _esc(_page_scheme(variant)),
              ''.join(menu), ''.join(pages), _JS)
