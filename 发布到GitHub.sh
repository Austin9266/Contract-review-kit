#!/usr/bin/env bash
# 发布到 GitHub —— 你自己跑，因为登录 GitHub 要你本人授权。
#
#   bash 发布到GitHub.sh <你的GitHub用户名> [仓库名]
#
# 仓库名不给就用 contract-review-kit。
#
# 它做四件事：
#   1. 再跑一次脱敏自检，不过关就停（这是最后一道闸）
#   2. 用 GitHub 的 noreply 邮箱做提交身份——你的真实邮箱不会进公开的提交历史
#   3. git init + 首次提交
#   4. 有 gh 就直接建仓库并推送；没有就打印手工推送的两条命令
set -euo pipefail
cd "$(dirname "$0")"

USER_NAME="${1:-}"
REPO="${2:-contract-review-kit}"
if [ -z "$USER_NAME" ]; then
  echo "用法：bash 发布到GitHub.sh <你的GitHub用户名> [仓库名]"
  echo "例：  bash 发布到GitHub.sh zhangsan contract-review-kit"
  exit 1
fi

echo "▸ 1/4 脱敏自检（最后一道闸）"
if ! python3 脱敏自检.py --quiet; then
  echo
  echo "自检没过，先把上面带 ❌ 的处理掉再发。已中止，什么都没提交。"
  exit 1
fi
echo "  ✅ 通过"
echo

echo "▸ 2/4 提交身份"
git init -q 2>/dev/null || true
# 只在这个仓库里设，不动你的全局 git 配置
git config core.quotepath false        # 中文文件名正常显示，不是一串八进制
git config user.name "$USER_NAME"
git config user.email "${USER_NAME}@users.noreply.github.com"
echo "  作者：$(git config user.name) <$(git config user.email)>"
echo "  （用 GitHub 的 noreply 邮箱，你的真实邮箱不会出现在公开提交历史里）"
echo

echo "▸ 3/4 首次提交"
git add -A
if git diff --cached --quiet; then
  echo "  没有要提交的改动"
else
  git commit -q -m "合同审查工具包：规则卡装配 + Word 修订批注落盘 + 交付前体检

- 138 张规则卡，按「通用核 ＋ 情形包 ＋ 类型包」按件装配，不全量套用
- 落盘引擎把纯文本结论写成 Word 真修订与真批注，C0–C7 体检不过不出文件
- 落盘台：本地网页版，人工逐条指认、改口径、跳过
- 工作台：离线网页，生成当次专用的审查 prompt
- 脱敏自检：发布前按特征查九类泄露
- 不含任何客户数据、机构标识或行业专属内容"
  echo "  ✅ 已提交"
fi
git branch -M main
echo

echo "▸ 4/4 推送"
if command -v gh >/dev/null 2>&1; then
  if gh auth status >/dev/null 2>&1; then
    gh repo create "$REPO" --public --source=. --remote=origin \
       --description "中文合同审查工具包：规则卡按件装配 + Word 修订批注确定性落盘 + 交付前体检" \
       --push
    echo
    echo "✅ 完事：https://github.com/${USER_NAME}/${REPO}"
    exit 0
  fi
  echo "  gh 装了但没登录。先跑：gh auth login"
else
  echo "  这台机器没有 gh。两条路："
fi
cat <<EOF

  A. 装 gh 再跑一次本脚本（最省事）
       brew install gh && gh auth login

  B. 在 github.com 上手工新建一个空仓库「${REPO}」（Public、不要勾 README），然后：
       git remote add origin https://github.com/${USER_NAME}/${REPO}.git
       git push -u origin main

  本地提交已经做好了，上面任选一条即可。
EOF
