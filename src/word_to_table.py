# -*- coding: utf-8 -*-
"""Word 文档信息扫描 -> 固定表格填充工具（配置驱动）

功能：
    扫描一个 Word 源文档，按“标签/关键词”定位并提取信息，
    把结果填入一个固定结构的 Word 表格（或已有 Word 模板中的占位符）。

用法：
    python word_to_table.py                       # 使用同目录 extract_config.json
    python word_to_table.py 自定义配置.json        # 指定配置文件

字段清单在 extract_config.json 的 "fields" 中维护，改配置即可，无需改代码。
"""
import os
import re
import sys
import json

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

import xlsx_fill

HERE = os.path.dirname(os.path.abspath(__file__))


# ============================================================
# 一、读取源文档：按文档原始顺序取出所有文本块
# ============================================================
def iter_block_texts(doc):
    """按 body 原始顺序遍历段落与表格，返回文本行列表（保留出现顺序）。"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    lines = []
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            lines.append(Paragraph(child, doc).text)
        elif child.tag == qn('w:tbl'):
            tbl = Table(child, doc)
            for row in tbl.rows:
                # 同一行的单元格用制表符拼接，方便“标签:值”跨列识别
                cell_texts = [c.text.replace('\n', ' ').strip() for c in row.cells]
                lines.append('\t'.join(cell_texts))
    return lines


def load_source_lines(path):
    doc = Document(to_docx(path))   # .doc/.wps 自动转 .docx，避免旧格式读不了
    lines = iter_block_texts(doc)
    # 去掉纯空白行两端空格，但保留空行用于分段判断
    return [ln.rstrip() for ln in lines]


# ============================================================
# 二、提取逻辑：按标签/关键词
# ============================================================
# 允许的分隔符：中文冒号、英文冒号，标签与分隔符间可有空白
_SEP = r'\s*[:：]\s*'


def _clean(value):
    """清洗提取到的值：去掉首尾空白、结尾多余标点。"""
    value = value.strip()
    value = value.strip('\u3000').strip()
    # 去掉行尾用于分隔的制表符残留
    value = value.rstrip('\t ')
    return value


def extract_inline(lines, labels):
    """行内模式：在同一行里找 “标签: 值”，取分隔符之后到行尾的内容。"""
    for label in labels:
        pat = re.compile(re.escape(label) + _SEP + r'([^\n]*)')
        for ln in lines:
            m = pat.search(ln)
            if m:
                val = _clean(m.group(1))
                # 若一行里有多个“标签:值”，只取到下一个已知分隔标签前
                if val:
                    return val
    return ''


def extract_block(lines, labels, stop_labels, max_lines=20):
    """区块模式：标签独占一行（当作小标题），取其后正文，直到遇到下一个标签或空行。"""
    stop_set = {s for s in stop_labels}
    for label in labels:
        for i, ln in enumerate(lines):
            stripped = ln.strip()
            # 标签行：完全等于标签，或“标签：”后为空
            is_heading = (stripped == label) or bool(
                re.fullmatch(re.escape(label) + _SEP, stripped)
            )
            if not is_heading:
                continue
            collected = []
            for nxt in lines[i + 1:i + 1 + max_lines]:
                s = nxt.strip()
                if not s:            # 空行视为区块结束
                    break
                if s in stop_set:    # 命中另一个字段标签，结束
                    break
                if stop_set and re.match(r'^(' + '|'.join(re.escape(x) for x in stop_set) + r')' + _SEP, s):
                    break
                collected.append(s)
            if collected:
                return _clean(' '.join(collected))
    return ''


def extract_field(lines, field, all_labels):
    """根据字段配置的 mode 提取值。"""
    labels = field.get('labels') or [field.get('name')]
    mode = field.get('mode', 'inline')
    value = ''
    if mode == 'inline':
        value = extract_inline(lines, labels)
    elif mode == 'block':
        value = extract_block(lines, labels, all_labels,
                              max_lines=field.get('max_lines', 20))
    elif mode == 'auto':
        value = extract_inline(lines, labels) or extract_block(lines, labels, all_labels)
    else:
        raise ValueError('未知的 mode: %s' % mode)

    # 可选：默认值 / 未提取到时的占位
    if not value:
        value = field.get('default', '')
    return value


def run_extraction(config):
    src = config.get('source_docx')
    if not src or not os.path.isabs(src):
        src = os.path.join(HERE, src or '')
    if not os.path.exists(src):
        raise FileNotFoundError('找不到源文档：%s' % src)

    lines = load_source_lines(src)
    fields = config.get('fields', [])
    all_labels = []
    for f in fields:
        all_labels.extend(f.get('labels') or [f.get('name')])

    results = {}
    for f in fields:
        name = f['name']
        results[name] = extract_field(lines, f, all_labels)
    return results


# ============================================================
# 二·补、动态扫描：以 Excel 表头作为“要扫描什么”的依据
# ============================================================
def scan_fields(src_path, xlsx_path):
    """读取 Excel 表头，表头每一列即一个待扫描字段；返回 (header, values)。

    header: {列字母: 表头文本}；values: {列字母: 从 Word 提取的值}。
    不依赖任何固定字段配置，换一张表格扫描内容自动跟着变。
    """
    lines = load_source_lines(src_path)
    header = xlsx_fill.read_header(xlsx_path)
    labels = list(header.values())
    values = {}
    for col, text in header.items():
        field = {'name': text, 'labels': [text], 'mode': 'auto'}
        values[col] = extract_field(lines, field, labels)
    return header, values


# ============================================================
# 二·补补、扫描 Word 内部的课程清单表（多行）
# ============================================================
def _norm_cell(s):
    return re.sub(r'\s+', '', s or '')


def _has_cjk(s):
    return any('\u4e00' <= ch <= '\u9fff' for ch in (s or ''))


# 统计/汇总行特征（这些行不是课程，按课程名/代码命中后跳过）
_SUMMARY_PAT = re.compile(
    r'合计|总计|小计|总学时|总学分|开课门数|开课学分|开课学时|门数'
    r'|最低总学分|至少修|毕业要求|说明')


def _num(s):
    try:
        return float(str(s).strip())
    except (TypeError, ValueError):
        return None


def _fmt(v):
    return str(int(v)) if float(v).is_integer() else str(v)


# 课程标识列与课程名称列的各种叫法（不同学校方案表头不一），统一归一
_CODE_ALIAS = ('课程编号', '课程代码', '课程代号', '编号', '代码', '课程序号')
_NAME_ALIAS = ('课程名称', '课程中文名称', '名称')
# 判断“这是不是课程表”：除名称列外，还需出现真正的排课属性列（学分/学时/学期）
# 或课程代码列；仅有 课程类别/毕业要求 等的是“课程→毕业要求”映射表，不当课程。
_DATA_HINTS = ('学分', '学时', '学时数', '总学时', '周学时', '开课学期', '学期')


def _find_col(head, aliases):
    for a in aliases:
        if a in head:
            return a
    return None


def _canon(head, col_name, target):
    """把某列改名（col_name→target），返回新列表。"""
    if col_name and col_name != target and col_name in head:
        head = list(head)
        head[head.index(col_name)] = target
    return head


def _norm_key(s):
    """合并键归一：去空白、罗马数字全角转半角、字母小写，便于跨表/写法合并。"""
    for k, v in (('Ⅰ', 'I'), ('Ⅱ', 'II'), ('Ⅲ', 'III'), ('Ⅳ', 'IV'),
                 ('Ⅴ', 'V'), ('Ⅵ', 'VI'), ('Ⅶ', 'VII'), ('Ⅷ', 'VIII'),
                 ('Ⅸ', 'IX'), ('Ⅹ', 'X')):
        s = s.replace(k, v)
    return re.sub(r'\s+', '', s).lower()


def _table_header(table):
    """在前几行内定位表头行（跳过表内标题行），识别/归一 课程名称、课程编号列，
    支持两行合并表头。返回 (有效表头, 数据起始行号)；识别不到返回 (None, 1)。"""
    rows = table.rows
    h = None
    for i in range(min(len(rows), 4)):
        cells = [_norm_cell(c.text) for c in rows[i].cells]
        if _find_col(cells, _NAME_ALIAS) and (
                _find_col(cells, _CODE_ALIAS)
                or any(x in cells for x in _DATA_HINTS)):
            h = i
            break
    if h is None:
        return None, 1
    head0 = [_norm_cell(c.text) for c in rows[h].cells]
    head0 = _canon(head0, _find_col(head0, _NAME_ALIAS), '课程名称')
    head0 = _canon(head0, _find_col(head0, _CODE_ALIAS), '课程编号')
    start = h + 1
    if h + 1 >= len(rows):
        return head0, start
    sub = [_norm_cell(c.text) for c in rows[h + 1].cells]
    ci = head0.index('课程编号') if '课程编号' in head0 else -1
    has_sub_word = any(x in ('理论', '实践') for x in sub)
    is_sub = (0 <= ci < len(sub) and sub[ci] == '课程编号') or has_sub_word
    if not is_sub:
        return head0, start
    eff = []
    for idx, top in enumerate(head0):
        s = sub[idx] if idx < len(sub) else ''
        if not s or s == top:
            eff.append(top)
        elif top == '学分':
            eff.append(s + '学分')
        elif top in ('学时或周数', '学时'):
            eff.append(s + '学时' if s in ('理论', '实践') else s)
        elif head0.count(top) > 1:
            eff.append(s)
        else:
            eff.append(top + s)
    return eff, h + 2


# 旧版 Word/WPS 二进制格式（python-docx 读不了）自动转 .docx 的 PowerShell(COM) 脚本
_PS_TO_DOCX = r'''$ErrorActionPreference='Stop'
$src=$args[0]; $dst=$args[1]
$app=$null
foreach($p in 'Word.Application','KWPS.Application','Kwps.Application'){
  try{ $app=New-Object -ComObject $p; break }catch{}
}
if($null -eq $app){ Write-Error 'NO_WORD_COM'; exit 3 }
$app.Visible=$false
try{ $app.DisplayAlerts=0 }catch{}
try{
  $doc=$app.Documents.Open($src,$false,$true)
  try{ $doc.SaveAs2($dst,12) }catch{ $doc.SaveAs($dst,12) }
  $doc.Close($false)
}finally{
  try{ $app.Quit() }catch{}
}
Write-Output $dst
'''


def to_docx(src):
    """确保得到 python-docx 能读的 .docx：.doc/.wps 用本机 Word/WPS(COM) 自动转换。

    无 Word/WPS 时抛 ValueError（带手动另存提示）；转换失败抛 RuntimeError。"""
    ext = os.path.splitext(src)[1].lower()
    if ext in ('.docx', ''):
        return src
    if ext not in ('.doc', '.wps'):
        raise ValueError('不支持的 Word 格式：%s' % ext)
    import subprocess
    import tempfile
    src = os.path.abspath(src)
    dst = os.path.join(tempfile.mkdtemp(prefix='doc2docx_'),
                       os.path.splitext(os.path.basename(src))[0] + '.docx')
    fd, ps = tempfile.mkstemp(suffix='.ps1')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(_PS_TO_DOCX)
        p = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                            '-File', ps, src, dst],
                           capture_output=True, timeout=180)
    finally:
        try:
            os.remove(ps)
        except OSError:
            pass
    if p.returncode != 0 or not os.path.isfile(dst):
        err = p.stderr.decode('utf-8', errors='replace')
        if 'NO_WORD_COM' in err or p.returncode == 3:
            raise ValueError('本机未检测到 Word/WPS，无法自动读取 .doc/.wps。\n'
                             '请用 Word/WPS 打开后“另存为 → .docx”再扫描。')
        raise RuntimeError('自动转换 .doc 失败：%s' % (err.strip()[:200] or '未知错误'))
    return dst


def scan_courses(src_path):
    """扫描 Word 内所有课程表，按课程编号合并各表信息。

    返回 (word_cols, rows)：
        word_cols: 所有课程表出现过的列名集合（含拆解后的 理论学分/实践学时…）
        rows:      [ {列名: 值}, ... ]，同一课程编号的多张表信息合并到一行；
                   并自动合计 学分=理论+实践、总学时=理论学时+实践学时。
    """
    doc = Document(to_docx(src_path))
    order = []
    by_code = {}
    name_to_key = {}        # 归一课程名 → 合并键（供无代码的行并入已有课程）
    word_cols = set()
    for t in doc.tables:
        try:
            if not t.rows:
                continue
            head, start = _table_header(t)
            if head is None or '课程名称' not in head:
                continue
            ni = head.index('课程名称')
            ci = head.index('课程编号') if '课程编号' in head else None
            for r in t.rows[start:]:
                cells = [_norm_cell(c.text) for c in r.cells]
                name = cells[ni] if ni < len(cells) else ''
                code = cells[ci] if (ci is not None and ci < len(cells)) else ''
                # 无课程名/是杂散表头/命中统计行/代码格填中文 → 跳过
                if (not name or name in _NAME_ALIAS or name in _CODE_ALIAS
                        or _SUMMARY_PAT.search(name)
                        or _SUMMARY_PAT.search(code) or _has_cjk(code)):
                    continue
                nk = _norm_key(name)
                if code:
                    key = code
                    name_to_key.setdefault(nk, code)
                else:                            # 无代码（如选修课/学期安排表）：
                    key = name_to_key.get(nk) or nk   # 能对上已有课程就并入
                rowdict = {}
                for idx, h in enumerate(head):
                    if h:
                        rowdict[h] = cells[idx] if idx < len(cells) else ''
                if key not in by_code:
                    by_code[key] = rowdict
                    order.append(key)
                else:
                    for k, v in rowdict.items():      # 多表合并：只补空缺
                        if not by_code[key].get(k):
                            by_code[key][k] = v
                word_cols.update(rowdict.keys())
        except Exception:                            # 单张表结构异常不影响其它表
            continue
    rows = [by_code[c] for c in order]
    for rd in rows:                                # 自动合计
        if not rd.get('学分') and rd.get('理论学分') and rd.get('实践学分'):
            a, b = _num(rd['理论学分']), _num(rd['实践学分'])
            if a is not None and b is not None:
                rd['学分'] = _fmt(a + b)
        if not rd.get('总学时') and rd.get('理论学时') and rd.get('实践学时'):
            a, b = _num(rd['理论学时']), _num(rd['实践学时'])
            if a is not None and b is not None:
                rd['总学时'] = _fmt(a + b)
        if not rd.get('是否有实践环节') and rd.get('实践学时'):
            p = _num(rd['实践学时'])
            if p is not None:
                rd['是否有实践环节'] = '是' if p > 0 else '否'
        word_cols.update(rd.keys())
    return word_cols, rows


# ============================================================
# 三、输出：生成固定表格 或 填充已有模板占位符
# ============================================================
def _set_cell(cell, text, bold=False):
    cell.text = ''
    run = cell.paragraphs[0].add_run(text if text is not None else '')
    run.bold = bold
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(11)


def replace_placeholders(doc, results):
    """把模板中的 {{字段名}} 占位符替换为提取值（正文段落 + 表格单元格）。"""
    def do_replace(paragraph):
        full = paragraph.text
        if '{{' not in full:
            return
        for k, v in results.items():
            full = full.replace('{{%s}}' % k, v)
        # 重写段落文本
        for run in paragraph.runs:
            run.text = ''
        if paragraph.runs:
            paragraph.runs[0].text = full
        else:
            paragraph.add_run(full)

    for p in doc.paragraphs:
        do_replace(p)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    do_replace(p)


def write_output(config, results):
    # 模式一：填写固定 Excel 表格（在表头后追加一行）
    excel_template = config.get('excel_template')
    if excel_template:
        if not os.path.isabs(excel_template):
            excel_template = os.path.join(HERE, excel_template)
        out_x = config.get('excel_output') or (
            os.path.splitext(excel_template)[0] + '-已填写.xlsx')
        if not os.path.isabs(out_x):
            out_x = os.path.join(HERE, out_x)
        header = xlsx_fill.read_header(excel_template)
        # 字段 -> 列 映射：表头文本等于字段 name 或命中其 labels 别名
        col_values = {}
        for f in config.get('fields', []):
            names = [f.get('name')] + (f.get('labels') or [])
            for col, htext in header.items():
                if htext in names:
                    col_values[col] = results.get(f.get('name'), '')
                    break
        xlsx_fill.append_row(excel_template, out_x, col_values)
        return out_x

    out = config.get('output_docx') or '提取结果.docx'
    if not os.path.isabs(out):
        out = os.path.join(HERE, out)

    template = config.get('template_docx')
    if template:
        if not os.path.isabs(template):
            template = os.path.join(HERE, template)
        doc = Document(template)
        replace_placeholders(doc, results)
        doc.save(out)
        return out

    # 未提供模板 -> 自动生成“字段 | 内容”两列固定表格
    doc = Document()
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    title = config.get('table_title', '文档信息提取表')
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    trun = tp.add_run(title)
    trun.bold = True
    trun.font.size = Pt(16)
    trun.font.name = 'Times New Roman'
    trun.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    doc.add_paragraph()

    headers = config.get('table_headers', ['项目', '内容'])
    tb = doc.add_table(rows=1, cols=2)
    tb.style = 'Table Grid'
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_cell(tb.rows[0].cells[0], headers[0], bold=True)
    _set_cell(tb.rows[0].cells[1], headers[1], bold=True)

    for f in config.get('fields', []):
        name = f['name']
        cells = tb.add_row().cells
        _set_cell(cells[0], name)
        _set_cell(cells[1], results.get(name, ''))

    # 列宽
    for r in tb.rows:
        r.cells[0].width = Cm(4.5)
        r.cells[1].width = Cm(11.5)

    doc.save(out)
    return out


# ============================================================
# 四、主入口
# ============================================================
def load_config(path):
    with open(path, 'r', encoding='utf-8') as fp:
        return json.load(fp)


def mark_red_docx(src, out, keyword):
    """把 Word（含表格内文字）里所有关键字标成红色字体，返回标红处数。

    只拆段重写命中段落的 run：每段深拷贝原 run 的 rPr（加粗/字号等保留），
    命中段额外加 w:color=FF0000；未命中段落一个字节不动。
    """
    import copy
    from docx.oxml import OxmlElement
    if not keyword:
        raise ValueError('关键字不能为空')
    # rPr 里排在 w:color 之后的子标签（schema 顺序，插入时不能乱）
    after_color = [qn(x) for x in (
        'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight',
        'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl',
        'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath')]
    doc = Document(src)
    paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paras.extend(cell.paragraphs)
    n = 0
    for p in paras:
        text = ''.join(r.text or '' for r in p.runs)
        if keyword not in text:
            continue
        # 切成 (文本, 是否命中) 段，跨 run 命中也能整段标红
        segs = []
        rest = text
        while keyword in rest:
            i = rest.index(keyword)
            if i:
                segs.append((rest[:i], False))
            segs.append((keyword, True))
            rest = rest[i + len(keyword):]
            n += 1
        if rest:
            segs.append((rest, False))
        # 取原段落第一个带格式的 run 作样式模板
        tpl = None
        for r in p.runs:
            rpr = r._r.find(qn('w:rPr'))
            if rpr is not None:
                tpl = rpr
                break
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        for seg_text, hit in segs:
            newr = OxmlElement('w:r')
            rpr = copy.deepcopy(tpl) if tpl is not None else OxmlElement('w:rPr')
            if hit:
                for old in rpr.findall(qn('w:color')):
                    rpr.remove(old)
                color = OxmlElement('w:color')
                color.set(qn('w:val'), 'FF0000')
                pos = None
                for tag in after_color:
                    found = rpr.find(tag)
                    if found is not None:
                        pos = found
                        break
                if pos is not None:
                    pos.addprevious(color)
                else:
                    rpr.append(color)
            newr.append(rpr)
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = seg_text
            newr.append(t)
            p._p.append(newr)
    if not n:
        raise ValueError('没有找到“%s”' % keyword)
    doc.save(out)
    return n


def _red_hex(val):
    """判断 w:color 的十六进制值是否属于红色族（含深红）。"""
    if not val or not re.fullmatch(r'[0-9A-Fa-f]{6}', val):
        return False
    r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
    return r >= 0xB0 and g <= 0x50 and b <= 0x50


def extract_red_docx(src, out):
    """把标红（红色字体）文字所在的整段提取到一个新的空白 Word，返回段数。

    段内任意 run 命中红色即整段带走；逐 run 深拷贝原 rPr，
    红色/加粗/字号在提取稿里原样保留；合并单元格不重复提取。
    """
    import copy
    doc = Document(src)
    paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paras.extend(cell.paragraphs)
    new = Document()
    seen = set()
    n = 0
    for p in paras:
        if id(p._p) in seen:
            continue
        seen.add(id(p._p))
        has_red = False
        for r in p.runs:
            rpr = r._r.find(qn('w:rPr'))
            color = rpr.find(qn('w:color')) if rpr is not None else None
            if color is not None and _red_hex(color.get(qn('w:val'))):
                has_red = True
                break
        if not has_red:
            continue
        target = new.add_paragraph()
        for r in p.runs:
            newr = target.add_run(r.text or '')
            rpr = r._r.find(qn('w:rPr'))
            if rpr is not None:
                newr._r.insert(0, copy.deepcopy(rpr))
        n += 1
    if not n:
        raise ValueError('文档中没有标红（红色字体）文字；请先完成标红。')
    new.save(out)
    return n


def main():
    cfg_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'extract_config.json')
    if not os.path.isabs(cfg_path):
        cfg_path = os.path.join(HERE, cfg_path)
    config = load_config(cfg_path)

    results = run_extraction(config)
    out = write_output(config, results)

    print('==== 提取结果 ====')
    for f in config.get('fields', []):
        name = f['name']
        val = results.get(name, '')
        flag = '' if val else '  (未提取到)'
        print('  %-12s : %s%s' % (name, val, flag))
    print('==================')
    print('已输出到：', out)


if __name__ == '__main__':
    main()
