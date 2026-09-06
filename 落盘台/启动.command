#!/bin/bash
# 双击这个文件就能打开落盘台。终端窗口关掉，服务就停。
cd "$(dirname "$0")" || exit 1
PY="$(command -v python3 || echo /usr/bin/python3)"
if [ ! -x "$PY" ]; then
  echo "没找到 python3。macOS 上装一次 Xcode 命令行工具即可：xcode-select --install"
  read -r -p "按回车关闭…" _
  exit 1
fi
exec "$PY" server.py
