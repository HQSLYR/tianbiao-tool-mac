#!/bin/bash
# ============================================================
#  填表工具 · macOS 一键启动（方案B：源码运行版）
#  若双击没反应，请打开「终端」App，输入 bash 加一个空格，
#  再把本文件拖进终端窗口，按回车即可。
# ============================================================
cd "$(dirname "$0")/src" || { echo "找不到 src 目录"; read -n 1 -s -r -p "按任意键关闭..."; exit 1; }

# 出错时暂停，避免窗口一闪而过看不到原因
trap 'echo; echo "（如上方有报错，请截图反馈；按任意键关闭...）"; read -n 1 -s -r' EXIT

PY="$(command -v python3)"
if [ -z "$PY" ]; then
  echo "❌ 未检测到 Python3。请先安装（任选其一）："
  echo "   1) 官网下载安装包： https://www.python.org/downloads/  （推荐，自带 Tkinter 图形库）"
  echo "   2) 或终端执行：      brew install python-tk"
  exit 1
fi
echo "使用 Python：$($PY --version)"

# 检查 tkinter（图形界面依赖）
if ! "$PY" -c "import tkinter" >/dev/null 2>&1; then
  echo "❌ 当前 Python 缺少 Tkinter 图形库。"
  echo "   请改用官网安装包版 Python，或执行：brew install python-tk"
  exit 1
fi

# 首次创建虚拟环境并安装依赖
if [ ! -d ".venv" ]; then
  echo "📦 首次运行，正在创建虚拟环境并安装依赖（约 1-2 分钟，请稍候）..."
  "$PY" -m venv .venv
  ./.venv/bin/pip install --upgrade pip >/dev/null 2>&1 || true
  ./.venv/bin/pip install -r requirements.txt || { echo "❌ 依赖安装失败，请检查网络后重试"; exit 1; }
fi

echo "🚀 启动填表工具..."
trap - EXIT
exec ./.venv/bin/python exe_main.py
