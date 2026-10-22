# -*- coding: utf-8 -*-
"""不依赖 openpyxl / pyexpat 的极简 xlsx 填表模块。

原理：xlsx 本质是 zip 包内的 XML。本模块直接用 zipfile 读取 sheet1.xml 与
sharedStrings.xml（正则解析，不走标准库 XML 解析器，绕开损坏的 pyexpat），
在表头之后追加数据行（用 inlineStr 内联字符串，无需改 sharedStrings），
再原样打包其余部件写回，从而保留原表格的格式与结构。

仅适用于"表头 + 追加数据行"的简单二维表（本项目的固定表格即此类）。
"""
import re
import zipfile


def colkey(col):
    """列字母排序键：A<B<...<Z<AA<AB（先比长度再比字符串）。"""
    return (len(col), col)


def _esc(s):
    """XML 文本转义。"""
    return (str(s).replace('&', '&amp;')
                  .replace('<', '&lt;')
                  .replace('>', '&gt;'))


def _unescape(s):
    s = (s.replace('&lt;', '<')
           .replace('&gt;', '>')
           .replace('&quot;', '"')
           .replace('&apos;', "'"))

    def _cr(m):
        try:
            return chr(int(m.group(1), 16) if m.group(1)[0] in 'xX'
                       else int(m.group(1)))
        except (ValueError, OverflowError):
            return m.group(0)
    # 数字字符引用 &#35838; / &#x8bfe;（openpyxl/WPS 常把中文存成这种形式）
    s = re.sub(r'&#(x[0-9a-fA-F]+|\d+);', _cr, s)
    return s.replace('&amp;', '&')


def _sheet_name(z):
    return _pick_sheet(z)[0]


def list_sheets(z):
    """按 workbook 顺序返回 [(表名, sheetN.xml 路径)]；无元数据时退回文件名序。"""
    cand = sorted(n for n in z.namelist()
                  if n.startswith('xl/worksheets/') and n.endswith('.xml'))
    if not cand:
        raise ValueError('xlsx 中找不到工作表')
    try:
        wb = z.read('xl/workbook.xml').decode('utf-8')
        rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    except KeyError:
        return [(n, n) for n in cand]
    rid = {}
    for rm in re.finditer(r'<Relationship\b[^>]*>', rels):
        tag = rm.group(0)
        i = re.search(r'Id="([^"]+)"', tag)
        t = re.search(r'Target="([^"]+)"', tag)
        if i and t:
            rid[i.group(1)] = t.group(1)
    out = []
    for name, r in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        t = rid.get(r, '')
        path = t[1:] if t.startswith('/') else ('xl/' + t)   # /xl/.. → xl/..；worksheets/.. → xl/worksheets/..
        if path not in cand:
            path = t if t in cand else None
        if path:
            out.append((_unescape(name), path))
    return out or [(n, n) for n in cand]


def _pick_sheet(z):
    """选第一张“表头下有数据行”的工作表（多表模板数据常不在第一张），
    全是空骨架时退回第一张。返回 (路径, [(表名, 路径)])。"""
    sheets = list_sheets(z)
    shared = _shared_strings(z)          # 只解一次：大表重复解码会假死
    for name, path in sheets:
        try:
            sheet = z.read(path).decode('utf-8')
        except KeyError:
            continue
        hr = _detect_header_row(sheet, shared)
        for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sheet, flags=re.S):
            if int(rm.group(1)) <= hr:
                continue
            for cm in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*)>(.*?)</c>',
                                  rm.group(2), flags=re.S):
                if _cell_value(cm.group(2), cm.group(3), shared):
                    return path, sheets
    return sheets[0][1], sheets


def _first_sheet(z):
    """第一张工作表路径（按 workbook 顺序，即最左边的标签页）。

    填表流程（①②③）默认往这张表写：模板表常是空骨架，必须按位置选，
    不能用“第一张有数据的表”去猜，否则会跳到示例/说明表上。
    """
    return list_sheets(z)[0][1]


def _resolve_sheet(z, sheet=None):
    """按表名取工作表路径；未指定时用第一张工作表（最左边标签页）。"""
    if not sheet:
        return _first_sheet(z)
    for name, path in list_sheets(z):
        if name == sheet:
            return path
    raise ValueError('找不到工作表“%s”' % sheet)


def sheet_info(path):
    """返回 (工作表名列表, 默认表名)；默认表 = 第一张“表头下有数据”的表。"""
    with zipfile.ZipFile(path) as z:
        sheets = list_sheets(z)
        default_path = _pick_sheet(z)[0]
    names = [n for n, _ in sheets]
    default = next((n for n, p in sheets if p == default_path),
                   names[0] if names else '')
    return names, default


def _shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    ss = z.read('xl/sharedStrings.xml').decode('utf-8')
    items = re.findall(r'<si>.*?<t[^>]*>(.*?)</t>.*?</si>', ss, flags=re.S)
    return [_unescape(x) for x in items]


def _row_cells(sheet, shared, rno):
    """取指定行号的非空单元格 {列字母: 文本}。"""
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rno, sheet, flags=re.S)
    cells = {}
    if not m:
        return cells
    for cm in re.finditer(r'<c r="([A-Z]+)%d"([^>]*)>(.*?)</c>' % rno, m.group(1), flags=re.S):
        col, attrs, inner = cm.groups()
        v = _cell_value(attrs, inner, shared)
        if v:
            cells[col] = v
    return cells


def _detect_header_row(sheet, shared):
    """自动定位表头行：跳过说明/标题等长文本行，取第一个“多个短标签”的行；找不到回退第1行。

    很多官方模板第1行是一整段填写说明（合并单元格），真正的表头在第2行。"""
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sheet, flags=re.S):
        rno = int(rm.group(1))
        texts = []
        for cm in re.finditer(r'<c r="[A-Z]+\d+"([^>]*)>(.*?)</c>', rm.group(2), flags=re.S):
            v = _cell_value(cm.group(1), cm.group(2), shared)
            if v:
                texts.append(v)
        if len(texts) >= 2 and all(len(t) <= 25 for t in texts):
            return rno
    return 1


def read_header(path, sheet=None):
    """读取表头行（自动跳过说明/标题行），返回 {列字母: 表头文本}。"""
    with zipfile.ZipFile(path) as z:
        sheet = z.read(_resolve_sheet(z, sheet)).decode('utf-8')
        shared = _shared_strings(z)
    return _row_cells(sheet, shared, _detect_header_row(sheet, shared))


def read_rows(path, sheet=None):
    """读取整张表，返回 (表头{列字母:文本}, 数据行[ {列字母: 值}, ... ])。

    支持共享字符串(t="s")、内联字符串(inlineStr)与纯数值单元格。
    全空行自动跳过。sheet 不指定时自动选第一张有数据的表。
    """
    with zipfile.ZipFile(path) as z:
        sheet = z.read(_resolve_sheet(z, sheet)).decode('utf-8')
        shared = _shared_strings(z)
    hr = _detect_header_row(sheet, shared)
    header = {}
    rows = []
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sheet, flags=re.S):
        rno = int(rm.group(1))
        vals = {}
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*)>(.*?)</c>', rm.group(2), flags=re.S):
            letter, attrs, inner = cm.groups()
            v = _cell_value(attrs, inner, shared)
            if v:
                vals[letter] = v
        if not vals:
            continue
        if rno == hr:
            header = vals
        elif rno > hr:
            rows.append(vals)
    return header, rows


ROW_RE = re.compile(r'<row r="(\d+)"([^>]*)>(.*?)</row>', flags=re.S)
CELL_RE = re.compile(r'<c r="([A-Z]+)\d+"([^>]*)>(.*?)</c>', flags=re.S)


def _cell_value(attrs, inner, shared):
    """取单元格当前文本（共享字符串/内联串/纯数值）。"""
    vm = re.search(r'<v>(.*?)</v>', inner, flags=re.S)
    if 't="s"' in attrs and vm:
        idx = int(vm.group(1))
        return _unescape(shared[idx]) if idx < len(shared) else ''
    if 't="inlineStr"' in attrs:
        im = re.search(r'<t[^>]*>(.*?)</t>', inner, flags=re.S)
        return _unescape(im.group(1)) if im else ''
    return _unescape(vm.group(1)) if vm else ''


def add_text(path, output, text, before=True, col=None, row=None, sheet=None):
    """给已填写的文字统一加前缀/后缀，返回修改的单元格数。

    col：None=全部列 / 单个列字母 / 列字母集合（多列或区间）；
    row：None=全部数据行 / 单个行号 / 行号集合（多行或区间）。
    只改非空单元格，空单元格不动；原样式(s=)保留。未指定行时自动跳过表头行。
    sheet 指定工作表名；不指定时用第一张工作表。
    """
    if not text:
        raise ValueError('要添加的文字不能为空')
    if isinstance(col, str):
        cols = {col}
    elif col is None:
        cols = None
    else:
        cols = set(col)
    if isinstance(row, int):
        rows = {row}
    elif row is None:
        rows = None
    else:
        rows = set(row)
    with zipfile.ZipFile(path) as z:
        sname = _resolve_sheet(z, sheet)
        sheet = z.read(sname).decode('utf-8')
        shared = _shared_strings(z)
        entries = [(item, z.read(item.filename)) for item in z.infolist()]
    hr = _detect_header_row(sheet, shared)
    state = {'n': 0}

    def repl_row(m):
        rno, attrs, body = m.groups()
        rn = int(rno)
        if rows is not None:
            if rn not in rows:
                return m.group(0)        # 按行：只动指定行
        elif rn <= hr:
            return m.group(0)            # 按列/全部：跳过表头及其上方说明行

        def repl_cell(cm):
            letter, cattrs, inner = cm.groups()
            if cols is not None and letter not in cols:
                return cm.group(0)
            old = _cell_value(cattrs, inner, shared)
            if not old:
                return cm.group(0)
            new = (text + old) if before else (old + text)
            sm = re.search(r's="(\d+)"', cattrs)
            s = ' s="%s"' % sm.group(1) if sm else ''
            state['n'] += 1
            return '<c r="%s%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (
                letter, rno, s, _esc(new))
        return '<row r="%s"%s>%s</row>' % (rno, attrs, CELL_RE.sub(repl_cell, body))

    sheet = ROW_RE.sub(repl_row, sheet)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item, data in entries:
            if item.filename == sname:
                data = sheet.encode('utf-8')
            zo.writestr(item, data)
    return state['n']


def transform_cells(path, output, fn, sheet=None, skip_header=False):
    """把每个非空**文本**单元格的值经 fn(old)->new 改写（只改文字，原样式/结构不动）。
    返回实际改动的单元格数。数字/公式/日期等非文本单元格一律不动（避免把数值转成文本）。
    skip_header=True 时跳过表头及其上方行。sheet 不指定时用第一张工作表。"""
    with zipfile.ZipFile(path) as z:
        sname = _resolve_sheet(z, sheet)
        sheet_xml = z.read(sname).decode('utf-8')
        shared = _shared_strings(z)
        entries = [(item, z.read(item.filename)) for item in z.infolist()]
    hr = _detect_header_row(sheet_xml, shared)
    state = {'n': 0}

    def repl_row(m):
        rno, attrs, body = m.groups()
        if skip_header and int(rno) <= hr:
            return m.group(0)

        def repl_cell(cm):
            letter, cattrs, inner = cm.groups()
            if 't="s"' not in cattrs and 't="inlineStr"' not in cattrs:
                return cm.group(0)          # 非文本（数字/公式/日期）不动
            old = _cell_value(cattrs, inner, shared)
            if not old:
                return cm.group(0)
            new = fn(old)
            if new == old:
                return cm.group(0)
            sm = re.search(r's="(\d+)"', cattrs)
            s = ' s="%s"' % sm.group(1) if sm else ''
            state['n'] += 1
            return ('<c r="%s%s"%s t="inlineStr"><is>'
                    '<t xml:space="preserve">%s</t></is></c>') % (
                letter, rno, s, _esc(new))
        return '<row r="%s"%s>%s</row>' % (rno, attrs, CELL_RE.sub(repl_cell, body))

    sheet_xml = ROW_RE.sub(repl_row, sheet_xml)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item, data in entries:
            if item.filename == sname:
                data = sheet_xml.encode('utf-8')
            zo.writestr(item, data)
    return state['n']


def split_rows(path, out_prefix, per, sheet=None):
    """按每份 per 行切割表格，返回生成的文件路径列表。

    每份都保留第 1 行表头与原有单元格样式/格式，行号重新从 2 编起。
    """
    if not per or per < 1:
        raise ValueError('每份行数必须≥1')
    with zipfile.ZipFile(path) as z:
        sname = _resolve_sheet(z, sheet)
        sheet = z.read(sname).decode('utf-8')
        entries = [(item, z.read(item.filename)) for item in z.infolist()]
    header = None
    rows = []
    for m in ROW_RE.finditer(sheet):
        rno, attrs, body = m.groups()
        if rno == '1':
            header = (attrs, body)
        else:
            rows.append((attrs, body))
    if header is None:
        raise ValueError('没读到表头行')
    if not rows:
        raise ValueError('没有数据行可切割')
    letters = [cm.group(1) for cm in CELL_RE.finditer(header[1])]
    lastcol = max(letters, key=colkey) if letters else 'A'
    head = sheet.index('<sheetData>')
    tail = sheet.index('</sheetData>') + len('</sheetData>')
    files = []
    for i in range(0, len(rows), per):
        chunk = rows[i:i + per]
        blocks = ['<row r="1"%s>%s</row>' % header]
        for j, (attrs, body) in enumerate(chunk):
            newno = j + 2
            body2 = re.sub(r'(<c r=")([A-Z]+)\d+(")',
                           lambda mm, n=newno: mm.group(1) + mm.group(2) + str(n) + mm.group(3),
                           body)
            blocks.append('<row r="%d"%s>%s</row>' % (newno, attrs, body2))
        newsheet = (sheet[:head] + '<sheetData>' + ''.join(blocks) +
                    '</sheetData>' + sheet[tail:])
        newsheet = re.sub(r'<dimension ref="[^"]*"\s*/>',
                          '<dimension ref="A1:%s%d"/>' % (lastcol, len(chunk) + 1),
                          newsheet, count=1)
        fname = '%s-第%d部分.xlsx' % (out_prefix, i // per + 1)
        with zipfile.ZipFile(fname, 'w', zipfile.ZIP_DEFLATED) as zo:
            for item, data in entries:
                if item.filename == sname:
                    data = newsheet.encode('utf-8')
                zo.writestr(item, data)
        files.append(fname)
    return files


def read_rows_numbered(path):
    """同 read_rows，但数据行带行号：返回 (表头, [(行号, {列字母: 值}), ...])。"""
    with zipfile.ZipFile(path) as z:
        sheet = z.read(_sheet_name(z)).decode('utf-8')
        shared = _shared_strings(z)
    header = {}
    rows = []
    for rm in ROW_RE.finditer(sheet):
        rno, _attrs, body = rm.groups()
        vals = {}
        for cm in CELL_RE.finditer(body):
            letter, cattrs, inner = cm.groups()
            v = _cell_value(cattrs, inner, shared)
            if v:
                vals[letter] = v
        if not vals:
            continue
        if rno == '1':
            header = vals
        else:
            rows.append((int(rno), vals))
    return header, rows


def _is_yellow(rgb6):
    """判断 6 位 hex 颜色是否属于黄色系（红绿高、蓝低）。"""
    try:
        r = int(rgb6[0:2], 16)
        g = int(rgb6[2:4], 16)
        b = int(rgb6[4:6], 16)
    except ValueError:
        return False
    return r >= 180 and g >= 140 and b <= 150 and r - b >= 40 and g - b >= 20


# Excel 传统索引色表（0~63），indexed="13" 即纯黄 FFFF00
_DEFAULT_INDEXED = (
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333')


def _indexed_palette(styles):
    """索引色表：优先用 styles.xml 自带的 indexedColors，否则用默认表。"""
    pal = list(_DEFAULT_INDEXED)
    m = re.search(r'<indexedColors>(.*?)</indexedColors>', styles, flags=re.S)
    if m:
        for i, c in enumerate(re.findall(r'<rgbColor rgb="([0-9A-Fa-f]{6,8})"', m.group(1))):
            if i < len(pal):
                pal[i] = c[-6:].upper()
    return pal


def _theme_colors(z):
    """theme1.xml 的主题色表，按 theme= 下标顺序（0=lt1,1=dk1,2=lt2,3=dk2,4=accent1…）。"""
    name = next((n for n in z.namelist() if n.startswith('xl/theme/theme')
                 and n.endswith('.xml')), None)
    if not name:
        return []
    t = z.read(name).decode('utf-8', 'ignore')
    m = re.search(r'<a:clrScheme.*?</a:clrScheme>', t, flags=re.S)
    if not m:
        return []
    cols = {}
    for nm, sysc, srgb in re.findall(
            r'<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>\s*'
            r'<a:(?:sysClr val="\w+" lastClr="([0-9A-Fa-f]{6})"|srgbClr val="([0-9A-Fa-f]{6})")'
            r'\s*/>\s*</a:\1>', m.group(0)):
        cols[nm] = (sysc or srgb).upper()
    seq = ('lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3',
           'accent4', 'accent5', 'accent6', 'hlink', 'folHlink')
    return [cols.get(n, 'FFFFFF') for n in seq]


def _apply_tint(rgb6, tint):
    """按 Excel 规则把主题色加 tint（HLS 亮度调整）后转回 6 位 hex。"""
    if not tint:
        return rgb6
    import colorsys
    try:
        t = float(tint)
        r, g, b = (int(rgb6[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except ValueError:
        return rgb6
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 - t) + t if t > 0 else l * (1 + t)
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return '%02X%02X%02X' % (round(r * 255), round(g * 255), round(b * 255))


def _fill_rgb(fill_xml, pal, theme):
    """取一个 fill 的实际填充色（6 位 hex）：rgb / indexed / theme+tint 三种存法都认。"""
    m = re.search(r'<fgColor\b([^>]*)/?>', fill_xml)
    if not m:
        return None
    attrs = m.group(1)
    rm = re.search(r'rgb="([0-9A-Fa-f]{6,8})"', attrs)
    if rm:
        return rm.group(1)[-6:].upper()
    im = re.search(r'indexed="(\d+)"', attrs)
    if im:
        i = int(im.group(1))
        return pal[i] if i < len(pal) else None
    tm = re.search(r'theme="(\d+)"', attrs)
    if tm and theme:
        i = int(tm.group(1))
        if i >= len(theme):
            return None
        tt = re.search(r'tint="(-?[\d.]+)"', attrs)
        return _apply_tint(theme[i], tt.group(1) if tt else None)
    return None


def _yellow_style_indices(z):
    """返回 styles.xml 中填充色为黄色系的单元格样式索引(s=)集合。

    兼容三种颜色存法：rgb="FFFFFF00"、indexed="13"（传统索引色）、
    theme="N"+tint（主题色），WPS/Excel 各种版本标黄都能认出。
    """
    if 'xl/styles.xml' not in z.namelist():
        return set()
    styles = z.read('xl/styles.xml').decode('utf-8')
    pal = _indexed_palette(styles)
    theme = _theme_colors(z)
    fills = []
    fm = re.search(r'<fills[^>]*>(.*?)</fills>', styles, flags=re.S)
    if fm:
        for fillm in re.finditer(r'<fill>(.*?)</fill>', fm.group(1), flags=re.S):
            fills.append(fillm.group(1))
    out = set()
    xm = re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', styles, flags=re.S)
    if not xm:
        return out
    for i, xfm in enumerate(re.finditer(r'<xf\b[^>]*>', xm.group(1))):
        fid = re.search(r'fillId="(\d+)"', xfm.group(0))
        if not fid:
            continue
        fidx = int(fid.group(1))
        if fidx >= len(fills):
            continue
        rgb = _fill_rgb(fills[fidx], pal, theme)
        if rgb and _is_yellow(rgb):
            out.add(i)
    return out


def marked_rows(path):
    """扫描**所有工作表**中带黄色底纹标识的行，
    返回 [(行号, {列字母: 值}), ...]。各表表头行不参与。"""
    with zipfile.ZipFile(path) as z:
        shared = _shared_strings(z)
        yellow = _yellow_style_indices(z)
        if not yellow:
            return []
        out = []
        for _name, sname in list_sheets(z):
            sheet = z.read(sname).decode('utf-8')
            hr = _detect_header_row(sheet, shared)
            for rm in ROW_RE.finditer(sheet):
                rno, _attrs, body = rm.groups()
                if int(rno) <= hr:
                    continue
                hit = False
                vals = {}
                for cm in CELL_RE.finditer(body):
                    letter, cattrs, inner = cm.groups()
                    v = _cell_value(cattrs, inner, shared)
                    if v:
                        vals[letter] = v
                    sm = re.search(r'(?:^|\s)s="(\d+)"', cattrs)
                    if sm and int(sm.group(1)) in yellow:
                        hit = True
                if hit and vals:
                    out.append((int(rno), vals))
    return out


def extract_marked_rows(path, out_path):
    """把带黄色底纹标识的行整行抄到新表，返回抄出的行数。

    新表 = 原表头 + 黄标行；行 XML 原样搬走（含 s= 样式索引），
    其余包内部件原样打包，格式不变。
    """
    with zipfile.ZipFile(path) as z:
        sname = _sheet_name(z)
        sheet = z.read(sname).decode('utf-8')
        yellow = _yellow_style_indices(z)
        entries = [(item, z.read(item.filename)) for item in z.infolist()]
    if not yellow:
        raise ValueError('表格里没扫到黄色底纹的标识')
    header = None
    kept = []
    for rm in ROW_RE.finditer(sheet):
        rno, attrs, body = rm.groups()
        if rno == '1':
            header = (attrs, body)
            continue
        hit = False
        for cm in CELL_RE.finditer(body):
            _letter, cattrs, _inner = cm.groups()
            sm = re.search(r'(?:^|\s)s="(\d+)"', cattrs)
            if sm and int(sm.group(1)) in yellow:
                hit = True
                break
        if hit:
            kept.append((attrs, body))
    if header is None:
        raise ValueError('没读到表头行')
    if not kept:
        raise ValueError('没扫到带黄色底纹标识的行')
    newsheet = _rebuild_sheet(sheet, header, kept)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item, data in entries:
            if item.filename == sname:
                data = newsheet.encode('utf-8')
            zo.writestr(item, data)
    return len(kept)


def _rebuild_sheet(sheet, header, kept):
    """用表头 + 选中行重建 sheetData，行号从 2 重编，样式原样保留。"""
    letters = [cm.group(1) for cm in CELL_RE.finditer(header[1])]
    lastcol = max(letters, key=colkey) if letters else 'A'
    blocks = ['<row r="1"%s>%s</row>' % header]
    for j, (attrs, body) in enumerate(kept):
        newno = j + 2
        body2 = re.sub(r'(<c r=")([A-Z]+)\d+(")',
                       lambda mm, n=newno: mm.group(1) + mm.group(2) + str(n) + mm.group(3),
                       body)
        blocks.append('<row r="%d"%s>%s</row>' % (newno, attrs, body2))
    head = sheet.index('<sheetData>')
    tail = sheet.index('</sheetData>') + len('</sheetData>')
    newsheet = (sheet[:head] + '<sheetData>' + ''.join(blocks) +
                '</sheetData>' + sheet[tail:])
    return re.sub(r'<dimension ref="[^"]*"\s*/>',
                  '<dimension ref="A1:%s%d"/>' % (lastcol, len(kept) + 1),
                  newsheet, count=1)


def extract_rows(path, out_path, keyword, col=None, sheet=None):
    """把含标识（关键字）的行整行抄到新表，返回抄出的行数。

    新表 = 原表头 + 命中行；行 XML 原样搬走（含 s= 样式索引），
    其余包内部件原样打包，格式不变。col 指定只在某列里找标识，
    默认整行任意单元格包含即命中。表头行不参与匹配。
    """
    if not keyword:
        raise ValueError('标识不能为空')
    with zipfile.ZipFile(path) as z:
        sname = _resolve_sheet(z, sheet)
        sheet = z.read(sname).decode('utf-8')
        shared = _shared_strings(z)
        entries = [(item, z.read(item.filename)) for item in z.infolist()]
    header = None
    kept = []
    for rm in ROW_RE.finditer(sheet):
        rno, attrs, body = rm.groups()
        if rno == '1':
            header = (attrs, body)
            continue
        if header is None:
            continue
        hit = False
        for cm in CELL_RE.finditer(body):
            letter, cattrs, inner = cm.groups()
            if col is not None and letter != col:
                continue
            v = _cell_value(cattrs, inner, shared)
            if v and keyword in v:
                hit = True
                break
        if hit:
            kept.append((attrs, body))
    if header is None:
        raise ValueError('没读到表头行')
    if not kept:
        raise ValueError('没有找到含“%s”的行' % keyword)
    newsheet = _rebuild_sheet(sheet, header, kept)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item, data in entries:
            if item.filename == sname:
                data = newsheet.encode('utf-8')
            zo.writestr(item, data)
    return len(kept)


def mark_red(path, out, keyword):
    """把表格里所有含关键字的单元格中，关键字部分标成红色字体，返回标红处数。

    命中单元格改写为 inlineStr 富文本：<is><r>…</r></is>，命中段带
    <rPr><color rgb="FFFF0000"/></rPr>；单元格原 s= 样式索引保留，
    未命中单元格一个字节不动。
    """
    if not keyword:
        raise ValueError('关键字不能为空')
    with zipfile.ZipFile(path) as z:
        sname = _sheet_name(z)
        sheet = z.read(sname).decode('utf-8')
        shared = _shared_strings(z)
        entries = [(item, z.read(item.filename)) for item in z.infolist()]
    state = {'n': 0}

    def repl_row(m):
        rno, attrs, body = m.groups()

        def repl_cell(cm):
            letter, cattrs, inner = cm.groups()
            val = _cell_value(cattrs, inner, shared)
            if keyword not in val:
                return cm.group(0)
            sm = re.search(r'(?:^|\s)s="(\d+)"', cattrs)
            s = ' s="%s"' % sm.group(1) if sm else ''
            runs = []
            rest = val
            while keyword in rest:
                i = rest.index(keyword)
                if i:
                    runs.append('<r><t xml:space="preserve">%s</t></r>' % _esc(rest[:i]))
                runs.append('<r><rPr><color rgb="FFFF0000"/></rPr>'
                            '<t xml:space="preserve">%s</t></r>' % _esc(keyword))
                rest = rest[i + len(keyword):]
                state['n'] += 1
            if rest:
                runs.append('<r><t xml:space="preserve">%s</t></r>' % _esc(rest))
            return '<c r="%s%s"%s t="inlineStr"><is>%s</is></c>' % (
                letter, rno, s, ''.join(runs))
        return '<row r="%s"%s>%s</row>' % (rno, attrs, CELL_RE.sub(repl_cell, body))

    sheet = ROW_RE.sub(repl_row, sheet)
    if not state['n']:
        raise ValueError('没有找到“%s”' % keyword)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item, data in entries:
            if item.filename == sname:
                data = sheet.encode('utf-8')
            zo.writestr(item, data)
    return state['n']


def extract_red_xlsx(src, out):
    """把标红（含红色字体片段）的单元格整格文字提取到一个新的空白 Word，返回格数。

    兼容本工具标红产物与 Excel/WPS 手工标红：富文本走 openpyxl
    rich_text 读取；任意片段命中红色族即整格带走。
    """
    from docx import Document
    from openpyxl import load_workbook
    from openpyxl.cell.rich_text import CellRichText

    def is_red(rgb):
        val = (rgb or '')[-6:]
        if not re.fullmatch(r'[0-9A-Fa-f]{6}', val):
            return False
        r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
        return r >= 0xB0 and g <= 0x50 and b <= 0x50

    wb = load_workbook(src, rich_text=True, data_only=True)
    new = Document()
    n = 0
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if not isinstance(v, CellRichText):
                        continue
                    red = False
                    for part in v:
                        font = getattr(part, 'font', None)
                        color = getattr(font, 'color', None) if font is not None else None
                        rgb = getattr(color, 'rgb', None) if color is not None else None
                        if is_red(rgb if isinstance(rgb, str) else ''):
                            red = True
                            break
                    if red and str(v).strip():
                        new.add_paragraph(str(v).strip())
                        n += 1
    finally:
        wb.close()
    if not n:
        raise ValueError('表格中没有标红（红色字体）内容；请先完成标红。')
    new.save(out)
    return n


def _sheet_parts(z):
    """{工作表名: 包内路径}。"""
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    out = {}
    for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        t = rid2target.get(m.group(2), '')
        if t:
            out[m.group(1)] = t if t.startswith('xl/') else 'xl/' + t.lstrip('/')
    return out


def _col_values(z, part, col, r1, r2, shared):
    """读指定工作表某列 r1~r2 的非空值。"""
    sheet = z.read(part).decode('utf-8')
    vals = []
    for rm in ROW_RE.finditer(sheet):
        rno, _attrs, body = rm.groups()
        n = int(rno)
        if n < r1 or n > r2:
            continue
        for cm in CELL_RE.finditer(body):
            letter, attrs, inner = cm.groups()
            if letter != col:
                continue
            v = _cell_value(attrs, inner, shared)
            if v:
                vals.append(v)
    return vals


def read_validations(path):
    """读目标表自带的下拉校验选项，返回 {列字母: 允许值集合}。

    支持两种 formula1：字面量列表（"考试,考查"）与定义名称
   （如 OPTIONS_H → datasheet!$H$1:$H$2）。用于填写前校验，
    避免把源里的脏值（如代码 29）写进有下拉限制的列。
    """
    with zipfile.ZipFile(path) as z:
        sheet = z.read(_sheet_name(z)).decode('utf-8')
        wb = z.read('xl/workbook.xml').decode('utf-8')
        shared = _shared_strings(z)
        parts = _sheet_parts(z)
        names = dict(re.findall(r'<definedName name="([^"]+)">(.*?)</definedName>',
                                wb, flags=re.S))

        def resolve(f):
            f = f.strip()
            if f.startswith('"') and f.endswith('"'):
                return set(x.strip() for x in f[1:-1].replace('，', ',').split(',')
                           if x.strip())
            ref = names.get(f)
            if not ref:
                return None
            m = re.match(r"^(?:'([^']+)'|([^'!]+))!\$?([A-Z]+)\$?(\d+):\$?[A-Z]+\$?(\d+)$",
                         ref.strip())
            if not m:
                return None
            part = parts.get(m.group(1) or m.group(2))
            if not part:
                return None
            return set(_col_values(z, part, m.group(3), int(m.group(4)),
                                   int(m.group(5)), shared))

        out = {}
        for m in re.finditer(r'<dataValidation\b([^>]*)>(.*?)</dataValidation>',
                             sheet, flags=re.S):
            attrs, body = m.groups()
            fm = re.search(r'<formula1>(.*?)</formula1>', body, flags=re.S)
            if not fm:
                continue
            allowed = resolve(_unescape(fm.group(1)))
            if not allowed:
                continue
            sq = re.search(r'sqref="([^"]+)"', attrs)
            if not sq:
                continue
            for token in sq.group(1).split():
                cm = re.match(r'\$?([A-Z]+)\$?\d*', token)
                if cm:
                    out.setdefault(cm.group(1), set()).update(allowed)
        return out


def _max_row(sheet):
    n = 0
    for rm in re.finditer(r'<row r="(\d+)"', sheet):
        n = max(n, int(rm.group(1)))
    return n


def _last_data_row(sheet, shared, hr):
    """最后一行“真正有内容”的行号（忽略只有格式无值的空占位行）。
    若表头以下全空则返回 hr。用于不清空时从数据紧接着开始写，避开模板尾部空行。"""
    last = hr
    for rm in ROW_RE.finditer(sheet):
        rno = int(rm.group(1))
        if rno <= hr:
            continue
        for cm in CELL_RE.finditer(rm.group(3)):
            if _cell_value(cm.group(2), cm.group(3), shared):
                last = max(last, rno)
                break
    return last


def append_row(template_path, output_path, col_values):
    """在模板表头后追加一行数据。col_values: {列字母: 值}。保留原格式写回 output_path。

    支持 output_path 与 template_path 相同（原地追加）：先把所有部件读入内存再写。
    """
    with zipfile.ZipFile(template_path) as z:
        sname = _sheet_name(z)
        sheet = z.read(sname).decode('utf-8')
        entries = [(item, z.read(item.filename)) for item in z.infolist()]

    newrow = _max_row(sheet) + 1
    sidx = _center_style_index(entries)

    cells = []
    for col in sorted(col_values, key=colkey):
        val = col_values[col]
        if val is None:
            val = ''
        cells.append(_cell_xml(col, newrow, val, sidx))
    spans = max((len(col) for col in col_values), default=1)
    rowxml = '<row r="%d" spans="1:%d">%s</row>' % (newrow, spans, ''.join(cells))
    sheet = sheet.replace('</sheetData>', rowxml + '</sheetData>')

    lastcol = max(col_values, key=colkey) if col_values else 'A'
    sheet = re.sub(r'<dimension ref="[^"]*"\s*/>',
                   '<dimension ref="A1:%s%d"/>' % (lastcol, newrow),
                   sheet, count=1)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item, data in entries:
            if item.filename == sname:
                data = sheet.encode('utf-8')
            zo.writestr(item, data)
    return newrow


def _center_style_index(entries):
    """在 styles.xml 的 cellXfs 末尾追加一个居中样式，返回其索引（失败返回 None）。

    直接修改 entries 中 styles 部件的字节，新写入的单元格用 s=索引 引用它。
    """
    for i, (item, data) in enumerate(entries):
        if item.filename != 'xl/styles.xml':
            continue
        styles = data.decode('utf-8')
        m = re.search(r'(<cellXfs[^>]*>)(.*?)(</cellXfs>)', styles, flags=re.S)
        if not m:
            return None
        opening, block, closing = m.groups()
        n = len(re.findall(r'<xf[ >]', block))
        new_opening = re.sub(r'count="\d+"', 'count="%d"' % (n + 1), opening)
        newxf = ('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" '
                 'applyAlignment="1">'
                 '<alignment horizontal="center" vertical="center"/></xf>')
        styles = styles.replace(opening + block + closing,
                                new_opening + block + newxf + closing, 1)
        entries[i] = (item, styles.encode('utf-8'))
        return n
    return None


def _cell_xml(col, rno, val, sidx):
    s = ' s="%d"' % sidx if sidx is not None else ''
    return '<c r="%s%d"%s t="inlineStr"><is><t>%s</t></is></c>' % (col, rno, s, _esc(val))


def _drop_data_rows(sheet, keep_upto=1):
    """删除表头行(keep_upto)及其以上的说明行以外的所有数据行，用于“清空旧数据后重填”。"""
    pat = re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', flags=re.S)

    def repl(m):
        return m.group(0) if int(m.group(1)) <= keep_upto else ''
    return pat.sub(repl, sheet)


def append_rows(template_path, output_path, rows, clear=False, center=True, fill=None):
    """在表头后一次性追加多行。rows: [ {列字母: 值}, ... ]。返回 (起始行, 结束行)。

    clear=True  先删掉表头以下所有旧行，只保留表头 + 本次新数据；
    center=True 新写入的文字水平+垂直居中（原表格其余格式不变）；
    fill        非空时，空单元格统一填该文字（文字填充）。
    """
    if not rows:
        raise ValueError('没有可填写的数据行')
    with zipfile.ZipFile(template_path) as z:
        sname = _first_sheet(z)
        sheet = z.read(sname).decode('utf-8')
        shared = _shared_strings(z)
        entries = [(item, z.read(item.filename)) for item in z.infolist()]

    hr = _detect_header_row(sheet, shared)
    if clear:
        keep = hr
        sheet = _drop_data_rows(sheet, keep)
    else:
        # 不清空：保留到最后一行有数据的行，删掉表尾空占位行（避免数据前出现空洞、
        # 也避免新行与旧空行 r= 重号）。
        keep = _last_data_row(sheet, shared, hr)
        sheet = _drop_data_rows(sheet, keep)
    start = keep + 1
    allcols = set()
    for rv in rows:
        allcols.update(rv.keys())
    lastcol = max(allcols, key=colkey) if allcols else 'A'
    spans = len(lastcol)
    sidx = _center_style_index(entries) if center else None

    blocks = []
    for i, rv in enumerate(rows):
        rno = start + i
        cells = []
        for col in sorted(rv, key=colkey):
            val = rv[col]
            if val is None or val == '':
                val = fill if fill else ''
            cells.append(_cell_xml(col, rno, val, sidx))
        blocks.append('<row r="%d" spans="1:%d">%s</row>' % (rno, spans, ''.join(cells)))
    sheet = sheet.replace('</sheetData>', ''.join(blocks) + '</sheetData>')

    end = start + len(rows) - 1
    sheet = re.sub(r'<dimension ref="[^"]*"\s*/>',
                   '<dimension ref="A1:%s%d"/>' % (lastcol, end),
                   sheet, count=1)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item, data in entries:
            if item.filename == sname:
                data = sheet.encode('utf-8')
            zo.writestr(item, data)
    return start, end
