# -*- coding: utf-8 -*-
"""⬔ 去水印：照片用 OpenCV inpaint 智能修补（无 cv2 时退回 ffmpeg delogo）；
视频用 ffmpeg delogo 对固定区域做淡化修复。都只读取源文件、另存新文件，不改源。

水印区域用矩形 box=(x, y, w, h)（原图像素坐标）表示，由界面框选得到。
"""
import os
import re
import importlib
import urllib.request

import video_tools
from video_tools import VideoError  # 复用异常类型

_IMG_EXT = ('.png', '.jpg', '.jpeg', '.bmp', '.webp', '.tif', '.tiff')
_VID_EXT = video_tools._VID_EXT


def is_image(path):
    return os.path.splitext(str(path))[1].lower() in _IMG_EXT


def image_size(path):
    from PIL import Image
    with Image.open(path) as im:
        return im.size  # (w, h)


def _clamp_box(box, W, H):
    x, y, w, h = [int(v) for v in box]
    x = max(0, min(x, W - 1))
    y = max(0, min(y, H - 1))
    w = max(2, min(w, W - x))
    h = max(2, min(h, H - y))
    return x, y, w, h


def _try_cv2():
    """动态导入 cv2/numpy（用 importlib 以免被 PyInstaller 打进 exe 增大体积）。"""
    try:
        cv2 = importlib.import_module('cv2')
        np = importlib.import_module('numpy')
        return cv2, np
    except Exception:
        return None


def download(url, dest_dir):
    """下载网络文件到 dest_dir，返回本地路径。"""
    url = url.strip()
    if not re.match(r'^https?://', url, re.I):
        raise VideoError('请输入以 http:// 或 https:// 开头的网址。')
    os.makedirs(dest_dir, exist_ok=True)
    name = url.split('?')[0].rstrip('/').split('/')[-1]
    name = re.sub(r'[\\/:*?"<>|]', '', name)
    if not name or not os.path.splitext(name)[1]:
        # 无扩展名时按响应头猜
        ext = ''
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            ct = urllib.request.urlopen(req, timeout=20).headers.get('Content-Type', '')
            if 'mp4' in ct:
                ext = '.mp4'
            elif 'jpeg' in ct:
                ext = '.jpg'
            elif 'png' in ct:
                ext = '.png'
            elif 'quicktime' in ct or 'video' in ct:
                ext = '.mov'
        except Exception:
            pass
        name = '下载文件' + (ext or '.mp4')
    dest = os.path.join(dest_dir, name)
    i = 1
    while os.path.exists(dest):
        base, e = os.path.splitext(name)
        dest = os.path.join(dest_dir, '%s(%d)%s' % (base, i, e))
        i += 1
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp, open(dest, 'wb') as f:
            while True:
                chunk = resp.read(1 << 20)
                if not chunk:
                    break
                f.write(chunk)
    except Exception as e:
        if os.path.exists(dest):
            try:
                os.remove(dest)
            except OSError:
                pass
        raise VideoError('下载失败：%s' % e)
    if not os.path.exists(dest) or os.path.getsize(dest) < 1024:
        raise VideoError('下载内容为空或过小，可能该网址不允许直接下载。')
    return dest


def grab_frame(path, out_png, at_sec=0):
    """从视频抽一帧存成 png（供界面预览与框选坐标）。"""
    ff = video_tools.find_ffmpeg()
    if not ff:
        raise VideoError('未找到 ffmpeg 引擎，无法处理视频。')
    seek = ['-ss', str(at_sec)] if at_sec else []
    r = video_tools._run([ff, '-hide_banner', '-y'] + seek +
                         ['-i', path, '-frames:v', '1', out_png])
    if r.returncode != 0 or not os.path.exists(out_png):
        # 有些视频第 0 秒抽帧失败，退而求其次不 seek
        r2 = video_tools._run([ff, '-hide_banner', '-y', '-i', path,
                               '-frames:v', '1', out_png])
        if not os.path.exists(out_png):
            raise VideoError('无法从视频取画面帧：\n' +
                             video_tools._tail(r2.stderr or r.stderr, 300))
    return out_png


def _video_size(path):
    ff = video_tools.find_ffmpeg()
    info = video_tools._probe(ff, path)
    if not info['w'] or not info['h']:
        raise VideoError('读不到视频分辨率，无法定位水印区域。')
    return info['w'], info['h']


def _imread(cv2, np, path):
    data = np.fromfile(path, dtype=np.uint8)
    return cv2.imdecode(data, cv2.IMREAD_COLOR)


def _imwrite(cv2, np, path, img):
    ext = os.path.splitext(path)[1] or '.png'
    ok, buf = cv2.imencode(ext, img)
    if ok:
        buf.tofile(path)
    return bool(ok)


def _delogo(path, out, box, W, H, crf=18, preset='veryfast', image=False):
    """ffmpeg delogo：要求框至少离边 1px，会自动收缩。image=True 时按图片输出（不加视频编码参数）。"""
    ff = video_tools.find_ffmpeg()
    if not ff:
        raise VideoError('未找到 ffmpeg 引擎。')
    x, y, w, h = _clamp_box(box, W, H)
    x = max(1, x); y = max(1, y)
    w = min(w, W - 1 - x); h = min(h, H - 1 - y)
    if w < 2 or h < 2:
        raise VideoError('水印框太靠边或太小，delogo 需要框离画面边缘至少 1 像素。')
    vf = 'delogo=x=%d:y=%d:w=%d:h=%d' % (x, y, w, h)
    cmd = [ff, '-hide_banner', '-y', '-i', path, '-vf', vf]
    if not image:
        cmd += ['-c:v', 'libx264', '-preset', preset, '-crf', str(crf), '-c:a', 'copy']
    cmd += [out]
    r = video_tools._run(cmd)
    if r.returncode != 0 or not os.path.exists(out) or os.path.getsize(out) < 1024:
        raise VideoError('去除失败：\n' + video_tools._tail(r.stderr, 400))
    return out


def remove_image(path, out, box, method='auto'):
    """照片去水印。method：'auto'/'inpaint' 优先 OpenCV 修补，'delogo' 用 ffmpeg。"""
    W, H = image_size(path)
    x, y, w, h = _clamp_box(box, W, H)
    cv = _try_cv2() if method in ('auto', 'inpaint') else None
    if cv is not None:
        cv2, np = cv
        img = _imread(cv2, np, path)
        if img is None:
            raise VideoError('无法读取图片，可能格式不支持。')
        pad = 3
        x0 = max(0, x - pad); y0 = max(0, y - pad)
        x1 = min(W, x + w + pad); y1 = min(H, y + h + pad)
        mask = np.zeros(img.shape[:2], np.uint8)
        mask[y0:y1, x0:x1] = 255
        res = cv2.inpaint(img, mask, 3, cv2.INPAINT_TELEA)
        if not _imwrite(cv2, np, out, res):
            raise VideoError('写出图片失败。')
        return out
    # 无 cv2 或指定 delogo：图片用 delogo（按图片输出，不加视频编码参数）
    return _delogo(path, out, (x, y, w, h), W, H, image=True)


def remove_video(path, out, box, crf=18, preset='veryfast', progress_cb=None):
    """视频去水印（固定区域 delogo 淡化）。box 用原视频分辨率坐标。"""
    W, H = _video_size(path)
    if progress_cb:
        progress_cb('用 delogo 淡化水印区域…')
    return _delogo(path, out, box, W, H, crf=crf, preset=preset, image=False)
