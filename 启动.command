#!/bin/bash
# 座位轮转助手 · 双击启动
cd "$(dirname "$0")" || exit 1

PY="$(command -v python3 || true)"
if [ -z "$PY" ]; then
  echo "未找到 python3。"
  echo "请先在终端运行：xcode-select --install"
  echo "或从 https://www.python.org 安装 Python 3，然后重新双击本文件。"
  echo ""
  read -r -p "按回车键退出…"
  exit 1
fi

echo "正在启动「座位轮转助手」…"
echo "浏览器会自动打开；关闭这个终端窗口即停止。"
echo ""
exec "$PY" app/server.py
