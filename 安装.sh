#!/usr/bin/env bash
# 把 skills/ 下的技能装进 Claude 的技能目录。
#   bash 安装.sh              → 装到 ~/.claude/skills/
#   bash 安装.sh /某个/skills  → 装到指定目录
set -euo pipefail
cd "$(dirname "$0")"
ROOT="$(pwd)"
DEST="${1:-$HOME/.claude/skills}"

mkdir -p "$DEST"
echo "工具包：$ROOT"
echo "技能装到：$DEST"
echo

for d in skills/*/; do
  name="$(basename "$d")"
  if [ -e "$DEST/$name" ]; then
    echo "  ⚠️  已存在，先备份：$name → $name.bak-$(date +%Y%m%d%H%M%S)"
    mv "$DEST/$name" "$DEST/$name.bak-$(date +%Y%m%d%H%M%S)"
  fi
  cp -R "$d" "$DEST/$name"
  # 别把运行产物带过去（装之前在工具包里跑过脚本就会有）
  find "$DEST/$name" \( -name '__pycache__' -o -name '*.pyc' -o -name '.DS_Store' \) \
       -exec rm -rf {} + 2>/dev/null || true
  echo "  ✅ $name"
done

# 技能被拷走后就找不到工具包根目录的 配置.json 了，给它一份稳定的落脚点
CFGDIR="$HOME/.contract-review"
if [ ! -f "$CFGDIR/配置.json" ]; then
  mkdir -p "$CFGDIR"
  cp 配置.json "$CFGDIR/配置.json"
  echo
  echo "  📝 配置已放到 $CFGDIR/配置.json"
else
  echo
  echo "  📝 $CFGDIR/配置.json 已存在，没有覆盖"
fi

echo
echo "———"
echo "装出去的技能读的是：$CFGDIR/配置.json"
python3 "$DEST/contract-review-sop/scripts/config.py" 2>&1 | grep -E '^配置来源|署名仍是' | head -2
echo "（在工具包目录里直接跑脚本时，读的是工具包自己的 配置.json——两份互不影响）"
echo "———"
echo
echo "下一步："
echo "  1. 把 $CFGDIR/配置.json 里的「署名」改成你自己的；"
echo "  2. 在 Claude 里说「审一下这份合同」并附上 .docx 文件。"
echo
echo "想让它更懂你这行：规则卡怎么加、怎么改，看工具包里的 CLAUDE.md。"
