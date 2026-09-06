#!/usr/bin/env bash
# 把落盘台的引擎、核心、一键落盘推到两个技能里（技能是分发副本，以本目录为准）。
set -euo pipefail
cd "$(dirname "$0")"
R=".."

for S in "$R/skills/contract-redline/scripts" "$R/skills/contract-review-sop/scripts"; do
  mkdir -p "$S/core"
  cp 引擎/*.py       "$S/"
  cp 核心/*.py       "$S/core/"
  cp 一键落盘.py     "$S/oneshot.py"
  cp config.py       "$S/config.py"
  echo "已同步 → $S"
done
echo
echo "同步完毕。推之前记得跑过自检："
echo "  python3 $R/skills/contract-review-sop/scripts/自检.py 某份真实来件.docx"
