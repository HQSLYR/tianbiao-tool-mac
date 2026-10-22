# -*- coding: utf-8 -*-
"""⬓ 视频合成：把最多 5 个视频按顺序拼成一个。

依赖 ffmpeg 引擎（打包进 exe 时随包；源码运行时可用本机 imageio-ffmpeg 或 PATH）。
只读取源视频、另存新文件，绝不改动源视频。

策略（兼容不同分辨率/帧率/有无音轨的混合素材）：
    1) 逐段规整成统一参数（分辨率取第一段、30fps、yuv420p、AAC 立体声 44.1k，
       无音轨的自动补静音）为临时 mp4；
    2) 用 concat demuxer 以 -c copy 无损快速拼接；
    3) 万一 copy 失败，回退到 concat 滤镜重编码。
"""
import os
import re
import sys
import glob
import shutil
import tempfile
import subprocess

MAX_VIDEOS = 5
_VID_EXT = ('.mp4', '.mov', '.mkv', '.avi', '.flv', '.wmv', '.m4v',
            '.webm', '.ts', '.mpg', '.mpeg', '.3gp', '.m2ts')


class VideoError(Exception):
    pass


def _candidate_dirs():
    """按优先级列出可能存放 ffmpeg.exe 的目录（随包 / exe 旁 / 脚本目录 / 当前目录）。"""
    dirs = []
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
        dirs += [os.path.join(base, 'tools'), base]
        meipass = getattr(sys, '_MEIPASS', None)
        if meipass:
            dirs += [os.path.join(meipass, 'tools'), meipass]
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        dirs += [os.path.join(here, 'tools'), here]
    except Exception:
        pass
    dirs += [os.path.join(os.getcwd(), 'tools'), os.getcwd()]
    return dirs


def _maybe_chmod(path):
    """mac/linux 下确保随包二进制有可执行权限（Windows 上无 os.X_OK 语义，忽略）。"""
    if os.name == 'nt':
        return
    try:
        if not os.access(path, os.X_OK):
            os.chmod(path, 0o755)
    except Exception:
        pass


def find_ffmpeg():
    """定位 ffmpeg 可执行文件（跨平台：Win 找 *.exe，mac/linux 找无后缀或带平台名的）；找不到返回 None。"""
    names = ('ffmpeg.exe', 'ffmpeg') if os.name == 'nt' else ('ffmpeg', 'ffmpeg.exe')
    for d in _candidate_dirs():
        if not d or not os.path.isdir(d):
            continue
        for nm in names:
            exact = os.path.join(d, nm)
            if os.path.isfile(exact):
                _maybe_chmod(exact)
                return exact
        hits = sorted(glob.glob(os.path.join(d, 'ffmpeg*')))
        hits = [h for h in hits if os.path.isfile(h)
                and not h.lower().endswith(('.py', '.pyc', '.txt', '.spec', '.md'))]
        if hits:
            _maybe_chmod(hits[0])
            return hits[0]
    try:
        import imageio_ffmpeg
        exe = imageio_ffmpeg.get_ffmpeg_exe()
        if exe and os.path.exists(exe):
            _maybe_chmod(exe)
            return exe
    except Exception:
        pass
    w = shutil.which('ffmpeg')
    if w:
        return w
    for p in (r'C:\ffmpeg\bin\ffmpeg.exe', r'C:\Program Files\ffmpeg\bin\ffmpeg.exe',
              '/usr/bin/ffmpeg', '/usr/local/bin/ffmpeg', '/opt/homebrew/bin/ffmpeg'):
        if os.path.exists(p):
            return p
    return None


def is_video(path):
    return os.path.splitext(str(path))[1].lower() in _VID_EXT


def _run(cmd):
    flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
    return subprocess.run(cmd, stdout=subprocess.PIPE,
                          stderr=subprocess.PIPE, creationflags=flags)


def _probe(ffmpeg, path):
    """用 `ffmpeg -i` 解析（不依赖 ffprobe）：返回 has_video/has_audio/width/height/duration秒。"""
    r = _run([ffmpeg, '-hide_banner', '-i', str(path)])
    err = r.stderr.decode('utf-8', 'ignore')
    has_v = has_a = False
    w = h = 0
    dur = 0.0
    dm = re.search(r'Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)', err)
    if dm:
        dur = int(dm.group(1)) * 3600 + int(dm.group(2)) * 60 + float(dm.group(3))
    for line in err.splitlines():
        if 'Stream #' not in line:
            continue
        if 'Video:' in line:
            has_v = True
            m = re.search(r'(\d{2,5})\s*x\s*(\d{2,5})', line)
            if m:
                w, h = int(m.group(1)), int(m.group(2))
        elif 'Audio:' in line:
            has_a = True
    return {'video': has_v, 'audio': has_a, 'w': w, 'h': h, 'dur': dur}


def _even(n):
    n = int(n or 0)
    if n <= 2:
        return 0
    return n if n % 2 == 0 else n - 1


def _tail(err_bytes, n=500):
    return err_bytes.decode('utf-8', 'ignore').strip()[-n:]


def concat_videos(paths, output, progress_cb=None, crf=18, preset='veryfast',
                  target_h=None, allow_lossless=True):
    """把 paths（2..MAX_VIDEOS 个）按顺序合成到 output（.mp4），返回 output 路径。

    progress_cb(idx, total, msg) 可选，用于回报进度。
    画质/分辨率参数：
      crf   —— 重编码画质，越小越清晰（16=更高画质、18=高画质、20=标准）；
      preset—— 编码速度/压缩权衡（veryfast 快、medium 更省空间更清）；
      target_h—— 不为空时将输出放大到该高度（如 1080/1440/2160，lanczos 缩放）；
      allow_lossless—— 同格式时优先无损 -c copy 直拼（保持原画质且最快）。
      一旦 target_h 有值或 allow_lossless=False，则强制重编码。
    """
    ffmpeg = find_ffmpeg()
    if not ffmpeg:
        raise VideoError(
            '未找到视频引擎 ffmpeg。\n\n'
            '· 用打包好的“填表工具.exe”通常已内置，若仍提示此消息，说明本次打包未包含 ffmpeg；\n'
            '· 也可把 ffmpeg.exe 放到本程序所在目录的 tools 文件夹后重试；\n'
            '· 或在装有 Python 的电脑上执行：pip install imageio-ffmpeg')
    paths = [str(p).strip() for p in paths if str(p).strip()]
    if len(paths) < 2:
        raise VideoError('请至少选择 2 个视频再合成。')
    if len(paths) > MAX_VIDEOS:
        raise VideoError('最多支持 %d 个视频。' % MAX_VIDEOS)
    for p in paths:
        if not os.path.exists(p):
            raise VideoError('文件不存在：%s' % p)
        if not is_video(p):
            raise VideoError('不是支持的视频格式：%s' % os.path.basename(p))

    output = str(output)
    if not output.lower().endswith('.mp4'):
        output += '.mp4'
    out_dir = os.path.dirname(output)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    infos = [_probe(ffmpeg, p) for p in paths]
    for p, inf in zip(paths, infos):
        if not inf['video']:
            raise VideoError('该文件没有视频轨，无法合成：%s' % os.path.basename(p))
    base = next((i for i in infos if i['w'] and i['h']), None)
    if target_h:
        H = _even(target_h) or 1080
        W = _even(round(H * base['w'] / base['h'])) if base else _even(round(H * 16 / 9))
        W = W or 1920
    else:
        W = _even(next((i['w'] for i in infos if i['w']), 0)) or 1280
        H = _even(next((i['h'] for i in infos if i['h']), 0)) or 720
    FPS = 30
    vf = ('scale=%d:%d:force_original_aspect_ratio=decrease:flags=lanczos,'
          'pad=%d:%d:(ow-iw)/2:(oh-ih)/2:color=black,setsar=1,fps=%d,format=yuv420p'
          % (W, H, W, H, FPS))

    tmpdir = tempfile.mkdtemp(prefix='vidmerge_')
    segs = []
    try:
        n = len(paths)
        # 快速通道：仅当所有输入参数完全一致（同宽高、同有无音轨）时，
        # 才直接对原文件无损 -c copy 拼接（不重编码，长视频秒级）；
        # 否则 copy 会产出“中途变分辨率/音频缺失错位”的坏文件，必须走逐段规整。
        sigs = {(i['w'], i['h'], i['audio']) for i in infos}
        can_fast = (allow_lossless and target_h is None and len(sigs) == 1
                    and next(iter(sigs))[0] > 0 and next(iter(sigs))[1] > 0)
        total_dur = sum(i['dur'] for i in infos)
        if can_fast:
            fast_lst = os.path.join(tmpdir, 'fast_list.txt')
            with open(fast_lst, 'w', encoding='utf-8') as f:
                for p in paths:
                    f.write("file '%s'\n" % p.replace("'", "'\\''"))
            if progress_cb:
                progress_cb(0, n, '尝试无损快速拼接…')
            rf = _run([ffmpeg, '-hide_banner', '-y', '-f', 'concat', '-safe', '0',
                       '-i', fast_lst, '-c', 'copy', '-movflags', '+faststart', output])
            if rf.returncode == 0 and os.path.exists(output) and os.path.getsize(output) >= 1024:
                ok = total_dur <= 0
                if not ok:
                    od = _probe(ffmpeg, output)['dur']
                    ok = abs(od - total_dur) <= max(2.0, total_dur * 0.02)
                if ok:
                    if progress_cb:
                        progress_cb(n, n, '完成（无损直拼）')
                    return output
            # 直拼不成：清掉半成品，退回逐段规整重编码
            if os.path.exists(output):
                try:
                    os.remove(output)
                except OSError:
                    pass
        for idx, (p, inf) in enumerate(zip(paths, infos), 1):
            if progress_cb:
                progress_cb(idx, n, '规整第 %d/%d 段…' % (idx, n))
            seg = os.path.join(tmpdir, 'seg_%02d.mp4' % idx)
            cmd = [ffmpeg, '-hide_banner', '-y', '-i', p]
            if inf['audio']:
                cmd += ['-map', '0:v:0', '-map', '0:a:0?', '-vf', vf,
                        '-c:v', 'libx264', '-preset', preset, '-crf', str(crf),
                        '-c:a', 'aac', '-ar', '44100', '-ac', '2', seg]
            else:
                cmd += ['-f', 'lavfi', '-i', 'anullsrc=r=44100:cl=stereo',
                        '-map', '0:v:0', '-map', '1:a:0', '-shortest', '-vf', vf,
                        '-c:v', 'libx264', '-preset', preset, '-crf', str(crf),
                        '-c:a', 'aac', '-ar', '44100', '-ac', '2', seg]
            r = _run(cmd)
            if r.returncode != 0 or not os.path.exists(seg) or os.path.getsize(seg) < 1024:
                raise VideoError('第 %d 段（%s）处理失败：\n%s' % (
                    idx, os.path.basename(p), _tail(r.stderr, 400)))
            segs.append(seg)

        if progress_cb:
            progress_cb(n, n, '拼接输出…')
        lst = os.path.join(tmpdir, 'list.txt')
        with open(lst, 'w', encoding='utf-8') as f:
            for s in segs:
                f.write("file '%s'\n" % s.replace("'", "'\\''"))
        r = _run([ffmpeg, '-hide_banner', '-y', '-f', 'concat', '-safe', '0',
                  '-i', lst, '-c', 'copy', '-movflags', '+faststart', output])
        if r.returncode == 0 and os.path.exists(output) and os.path.getsize(output) >= 1024:
            return output
        # 兜底：concat 滤镜重编码
        inp = []
        for s in segs:
            inp += ['-i', s]
        fc = ''.join('[%d:v][%d:a]' % (i, i) for i in range(len(segs))) + \
            'concat=n=%d:v=1:a=1[outv][outa]' % len(segs)
        r2 = _run([ffmpeg, '-hide_banner', '-y'] + inp +
                  ['-filter_complex', fc, '-map', '[outv]', '-map', '[outa]',
                   '-c:v', 'libx264', '-preset', preset, '-crf', str(crf),
                   '-c:a', 'aac', '-ar', '44100', '-ac', '2',
                   '-movflags', '+faststart', output])
        if r2.returncode != 0 or not os.path.exists(output) or os.path.getsize(output) < 1024:
            raise VideoError('拼接失败：\n' + _tail(r2.stderr or r.stderr, 500))
        return output
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
