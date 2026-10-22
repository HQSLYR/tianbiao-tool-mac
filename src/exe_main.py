# -*- coding: utf-8 -*-
"""打包入口（PyInstaller 用它生成 exe）。

与直接打包 gui.py 的区别：启动过程任何异常都会先写一份
《填表工具-启动错误.log》到 exe 同目录，再弹窗，方便在别的电脑上排障。
"""
import os
import sys
import traceback


def _log_path():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, '填表工具-启动错误.log')


try:
    import gui
    gui.main()
except Exception:
    try:
        with open(_log_path(), 'w', encoding='utf-8') as f:
            f.write('启动失败， traceback：\n')
            traceback.print_exc(file=f)
    except OSError:
        pass
    raise
