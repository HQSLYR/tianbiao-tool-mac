# -*- coding: utf-8 -*-
"""共享文档库：任何人都能上传自己的文档，保存后大家打开工具都能看到。

原理：库就是一个文件夹（默认在程序旁边的 shared_notes），上传即把文件
复制进库并登记上传人/时间。把库路径改到局域网共享盘（\\电脑名\共享夹），
所有同事用同一台机器或同一共享路径，就能看到彼此上传的文档。
"""
import getpass
import json
import os
import re
import shutil
import sys
import time

_INDEX = '_index.json'


def base_dir():
    if getattr(sys, 'frozen', False):        # 打包成 exe 后 = exe 所在目录
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def config_file():
    return os.path.join(base_dir(), 'shared_notes.cfg')


def lib_dir():
    """当前共享库目录；未设置时在程序旁自动创建 shared_notes。"""
    try:
        with open(config_file(), encoding='utf-8') as f:
            p = f.read().strip()
        if p and os.path.isdir(p):
            return p
    except OSError:
        pass
    d = os.path.join(base_dir(), 'shared_notes')
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        return os.path.join(os.getcwd(), 'shared_notes')
    return d


def set_dir(path):
    """把共享库指到别的文件夹（如局域网共享盘），目录必须可访问。"""
    if not path or not os.path.isdir(path):
        raise ValueError('路径不存在或不可访问：%s' % path)
    with open(config_file(), 'w', encoding='utf-8') as f:
        f.write(path)


def user():
    try:
        return getpass.getuser() or '同事'
    except Exception:
        return '同事'


def _index_path():
    return os.path.join(lib_dir(), _INDEX)


def _load_index():
    try:
        with open(_index_path(), encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_index(index):
    try:
        with open(_index_path(), 'w', encoding='utf-8') as f:
            json.dump(index, f, ensure_ascii=False, indent=1)
    except OSError:
        pass


def upload(paths, who=None):
    """复制文件进共享库并登记上传人/时间；同名自动加序号。返回入库文件名列表。"""
    who = (who or user()).strip() or user()
    d = lib_dir()
    index = _load_index()
    saved = []
    for src in paths:
        src = os.path.abspath(src)
        if not os.path.isfile(src):
            continue
        stem, ext = os.path.splitext(os.path.basename(src))
        name, n = os.path.basename(src), 1
        while os.path.exists(os.path.join(d, name)):
            n += 1
            name = '%s(%d)%s' % (stem, n, ext)
        dst = os.path.join(d, name)
        shutil.copy2(src, dst)
        index[name] = {'who': who,
                       'time': time.strftime('%Y-%m-%d %H:%M'),
                       'size': os.path.getsize(dst)}
        saved.append(name)
    _save_index(index)
    return saved


def list_items():
    """库内全部文档 [(文件名, 上传人, 时间, 大小, 全路径)]，按时间新在前。

    别人手动拷进这个文件夹的文件也会出现（上传人记“手动放入”）。
    """
    d = lib_dir()
    index = _load_index()
    items = []
    try:
        names = os.listdir(d)
    except OSError:
        return items
    for name in sorted(names):
        fp = os.path.join(d, name)
        if not os.path.isfile(fp) or name == _INDEX or name.startswith('~$'):
            continue
        meta = index.get(name) or {}
        try:
            size = os.path.getsize(fp)
        except OSError:
            size = 0
        items.append((name, meta.get('who', '手动放入'),
                      meta.get('time', ''), size, fp))
    items.sort(key=lambda t: (t[2], t[0]), reverse=True)
    return items


def fmt_size(n):
    """字节数转可读字符串（B/KB/MB/GB）。"""
    n = float(n or 0)
    for unit in ('B', 'KB', 'MB', 'GB'):
        if n < 1024 or unit == 'GB':
            return ('%.0f %s' % (n, unit)) if unit == 'B' else ('%.1f %s' % (n, unit))
        n /= 1024.0


def usage():
    """统计共享库占用：返回 (总字节, 文件数, [(上传人, 个数, 字节)...] 按占用降序)。"""
    items = list_items()
    total = 0
    per = {}
    for _name, who, _t, size, _fp in items:
        total += size
        c, s = per.get(who, (0, 0))
        per[who] = (c + 1, s + size)
    rank = sorted(((w, c, s) for w, (c, s) in per.items()),
                  key=lambda t: t[2], reverse=True)
    return total, len(items), rank


def delete(name):
    """从库里删除一个文件并销账。"""
    try:
        os.remove(os.path.join(lib_dir(), os.path.basename(name)))
    except OSError:
        pass
    index = _load_index()
    index.pop(name, None)
    _save_index(index)


def add_text_note(text, who=None, title=None):
    """把一段文字笔记存成 .txt 放进共享库（大家打开同一共享文件夹就能看到）。
    自动写上上传人/时间抬头；同名自动加序号。返回入库文件名。"""
    text = (text or '').strip()
    if not text:
        raise ValueError('笔记内容不能为空')
    who = (who or user()).strip() or user()
    d = lib_dir()
    os.makedirs(d, exist_ok=True)
    stamp = time.strftime('%Y-%m-%d %H:%M')
    stem = re.sub(r'[\\/:*?\"<>|\r\n]', '', (title or '').strip())[:20] or '笔记'
    name, n = '%s_%s.txt' % (stem, time.strftime('%Y%m%d-%H%M%S')), 1
    while os.path.exists(os.path.join(d, name)):
        n += 1
        name = '%s(%d).txt' % (stem, n)
    body = '【文字笔记】上传人：%s   时间：%s\n%s\n\n%s\n' % (
        who, stamp, '-' * 30, text)
    with open(os.path.join(d, name), 'w', encoding='utf-8-sig') as f:
        f.write(body)
    index = _load_index()
    index[name] = {'who': who, 'time': stamp,
                   'size': os.path.getsize(os.path.join(d, name))}
    _save_index(index)
    return name
