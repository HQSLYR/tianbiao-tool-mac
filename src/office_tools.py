# -*- coding: utf-8 -*-
"""日常办公增强：多表合并、批量套打、两表对账、格式清洗、文档互转、批量重命名。

全部复用 xlsx_fill(zipfile 读写，不依赖 openpyxl) 与 word_to_table(python-docx)，
只改文字/结构，尽量保留原表格与文档格式。函数都是纯逻辑，便于脱离界面单测。
"""
import csv
import os
import re
import shutil

import xlsx_fill
import word_to_table as core


def _norm(s):
    """表头/键归一化：去星号与所有空白，便于匹配 *课程编号 之类。"""
    return re.sub(r'\s+', '', (s or '').replace('*', '').replace('＊', ''))


def _safe_name(s):
    s = re.sub(r'[\\/:*?\"<>|\r\n\t]', '', (s or '')).strip()
    return s[:60]


# ---------------- ⑪ 多表合并汇总 ----------------
def merge_xlsx(paths, output, dedup=False):
    """把多份**同格式** Excel 合并成一张总表（以第一份的表头/格式为模板）。

    列按“表头文字”对齐（不同文件同名列自动对上）。dedup=True 时全行相同只留一条。
    返回 (合并后行数, [(文件名, 该文件贡献行数)], 去重删除数)。
    """
    if not paths:
        raise ValueError('请先选择要合并的表格')
    canon = None                 # {归一化表头: 目标列字母}
    all_rows, seen, dup, per = [], set(), 0, []
    for p in paths:
        header, rows = xlsx_fill.read_rows(p)
        if canon is None:
            canon = {_norm(t): L for L, t in header.items() if _norm(t)}
        mapped = []
        for rv in rows:
            m = {}
            for L, v in rv.items():
                tgt = canon.get(_norm(header.get(L, '')))
                if tgt:
                    m[tgt] = v
            if not m:
                continue
            if dedup:
                key = tuple(sorted(m.items()))
                if key in seen:
                    dup += 1
                    continue
                seen.add(key)
            mapped.append(m)
        all_rows.extend(mapped)
        per.append((os.path.basename(p), len(mapped)))
    if not all_rows:
        raise ValueError('这些表里没读到可合并的数据行（确认它们有相同的表头）')
    xlsx_fill.append_rows(paths[0], output, all_rows, clear=True, center=True)
    return len(all_rows), per, dup


# ---------------- ⑫ 批量套打 / 邮件合并 ----------------
def mail_merge(template, data_xlsx, out_dir, name_field=None):
    """Word 模板里的 {{列标题}} 占位符，按名单表逐行替换，批量生成 N 份 docx。

    template：.docx/.doc/.wps（旧格式自动转换）；data_xlsx：名单表；
    name_field：用哪一列的值给文件命名（留空则用“第N份”）。返回生成文件路径列表。
    """
    from docx import Document
    if not (template and os.path.exists(template)):
        raise ValueError('请先选择模板文档')
    os.makedirs(out_dir, exist_ok=True)
    tpl = core.to_docx(template)
    header, rows = xlsx_fill.read_rows(data_xlsx)
    if not rows:
        raise ValueError('名单表里没有数据行')
    outs = []
    for i, rv in enumerate(rows, 1):
        results = {}
        for L, title in header.items():
            val = rv.get(L, '')
            results[title] = val
            results[_norm(title)] = val
        doc = Document(tpl)
        core.replace_placeholders(doc, results)
        nm = ''
        if name_field:
            nm = (results.get(name_field) or results.get(_norm(name_field)) or '').strip()
        nm = _safe_name(nm) or ('第%d份' % i)
        fp = os.path.join(out_dir, '%s.docx' % nm)
        k = 1
        while os.path.exists(fp):
            k += 1
            fp = os.path.join(out_dir, '%s(%d).docx' % (nm, k))
        doc.save(fp)
        outs.append(fp)
    return outs


# ---------------- ⑬ 两表对账 ----------------
def reconcile(a, b, key_title, report_out=None):
    """按关键列比对两张表：找出 只在A / 只在B / 同键但值不同。

    返回 {'only_a','only_b','diffs','both','total_a','total_b'}；
    给了 report_out 则同时写一份 .txt 对账报告。
    """
    ha, ra = xlsx_fill.read_rows(a)
    hb, rb = xlsx_fill.read_rows(b)
    ka = next((L for L, t in ha.items() if _norm(t) == _norm(key_title)), None)
    kb = next((L for L, t in hb.items() if _norm(t) == _norm(key_title)), None)
    if not ka:
        raise ValueError('表A里找不到关键列“%s”' % key_title)
    if not kb:
        raise ValueError('表B里找不到关键列“%s”' % key_title)

    def build(header, rows, keycol):
        idx = {}
        for rv in rows:
            k = _norm(rv.get(keycol, ''))
            if not k:
                continue
            idx[k] = {_norm(header.get(L, '')): v for L, v in rv.items()
                      if _norm(header.get(L, ''))}
            idx[k]['\u0000key'] = rv.get(keycol, '')
        return idx

    A, B = build(ha, ra, ka), build(hb, rb, kb)
    only_a = [A[k]['\u0000key'] for k in A if k not in B]
    only_b = [B[k]['\u0000key'] for k in B if k not in A]
    diffs = []
    for k in A:
        if k not in B:
            continue
        da = {c: v for c, v in A[k].items()
              if c != '\u0000key' and B[k].get(c, '') != v}
        db = {c: v for c, v in B[k].items()
              if c != '\u0000key' and A[k].get(c, '') != v}
        if da or db:
            diffs.append((A[k]['\u0000key'], da, db))
    result = {'only_a': only_a, 'only_b': only_b, 'diffs': diffs,
              'both': len(set(A) & set(B)), 'total_a': len(A), 'total_b': len(B)}
    if report_out:
        with open(report_out, 'w', encoding='utf-8-sig') as f:
            f.write('两表对账报告\n关键列：%s\n表A %d 条 / 表B %d 条\n\n' % (
                key_title, result['total_a'], result['total_b']))
            f.write('【两表都有】%d 条\n' % result['both'])
            f.write('【只在表A】%d 条\n' % len(only_a))
            for x in only_a:
                f.write('   %s\n' % x)
            f.write('【只在表B】%d 条\n' % len(only_b))
            for x in only_b:
                f.write('   %s\n' % x)
            f.write('【同键但内容不同】%d 条\n' % len(diffs))
            for key, da, db in diffs:
                f.write('   %s：\n' % key)
                for c in sorted(set(da) | set(db)):
                    f.write('       %s → A=%s ｜ B=%s\n' % (
                        c, da.get(c, ''), db.get(c, '')))
        result['report'] = report_out
    return result


# ---------------- ⑭ 格式清洗 ----------------
def _to_half(s):
    out = []
    for ch in s:
        code = ord(ch)
        if code == 0x3000:
            out.append(' ')
        elif 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        else:
            out.append(ch)
    return ''.join(out)


def _clean_val(v):
    """全角转半角、去零宽字符、首尾去空白、连续空格压一个。"""
    s = _to_half(v).replace('\u200b', '').replace('\ufeff', '')
    s = re.sub(r'[ \t]+', ' ', s)
    return s.strip()


def clean_xlsx(path, output):
    """清洗 Excel 文本单元格（不动数值/公式/格式），返回改动的单元格数。"""
    if not (path and os.path.exists(path)):
        raise ValueError('请先选择要清洗的表格')
    n = xlsx_fill.transform_cells(path, output, _clean_val)
    return n


# ---------------- ⑮ 文档互转 ----------------
def convert_to_docx(src, out_dir):
    """.doc/.wps → .docx（本机 Word/WPS COM），输出到 out_dir。返回目标路径。"""
    tmp = core.to_docx(src)                       # 可能抛 ValueError/RuntimeError
    out = os.path.join(out_dir,
                       os.path.splitext(os.path.basename(src))[0] + '.docx')
    k = 1
    while os.path.exists(out):
        k += 1
        out = os.path.join(out_dir,
                           '%s(%d).docx' % (os.path.splitext(os.path.basename(src))[0], k))
    shutil.copy2(tmp, out)
    return out


def xlsx_to_csv(path, out_csv):
    """Excel 首表 → CSV（utf-8-sig，Excel 双击不乱码）。返回 (csv路径, 行数)。"""
    header, rows = xlsx_fill.read_rows(path)
    cols = sorted(header, key=xlsx_fill.colkey)
    with open(out_csv, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow([header.get(c, '') for c in cols])
        for rv in rows:
            w.writerow([rv.get(c, '') for c in cols])
    return out_csv, len(rows)


# ---------------- ⑯ 批量重命名 ----------------
def rename_plan(folder, mode, find='', repl='', prefix='', start=1, ext_filter=None):
    """生成重命名预览方案 [(旧全路径, 新全路径, 旧名, 新名)]，不落盘。

    mode='replace'：文件名里 find→repl；mode='prefix'：加前缀+序号。
    ext_filter 形如 '.docx'（只处理该类型，None=全部）。
    """
    if not (folder and os.path.isdir(folder)):
        raise ValueError('请先选择要整理的文件夹')
    files = [f for f in os.listdir(folder)
             if os.path.isfile(os.path.join(folder, f)) and not f.startswith('~$')]
    if ext_filter:
        files = [f for f in files if os.path.splitext(f)[1].lower() == ext_filter.lower()]
    files.sort()
    if not files:
        raise ValueError('文件夹里没有符合条件的文件')
    width = max(2, len(str(start + len(files))))
    plan, used = [], set()
    for i, f in enumerate(files):
        stem, ext = os.path.splitext(f)
        if mode == 'replace':
            if not find:
                raise ValueError('请填写要查找的文字')
            new = f.replace(find, repl)
        elif mode == 'prefix':
            new = '%s%0*d%s' % (prefix, width, start + i, ext)
        else:
            raise ValueError('未知的重命名方式')
        if new == f or new in used or os.path.exists(os.path.join(folder, new)):
            continue
        used.add(new)
        plan.append((os.path.join(folder, f), os.path.join(folder, new), f, new))
    return plan


def apply_rename(plan):
    """执行重命名，返回成功条数。"""
    n = 0
    for old, new, _of, _nf in plan:
        try:
            os.rename(old, new)
            n += 1
        except OSError:
            continue
    return n
