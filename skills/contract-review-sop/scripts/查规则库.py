#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""查规则库.py —— 规则库体检。改过卡片就跑一次，别等装配时才发现。

    python3 查规则库.py                    # 查本技能的 rules/规则库.json
    python3 查规则库.py 别的库.json         # 查指定文件（工作台导出的 rules.js 也认）
    python3 查规则库.py --stats            # 只看统计，不查错

查这些：
  E1 卡片 id 重复            —— 装配时后一张会把前一张顶掉，最隐蔽的一种坏法
  E2 id 前缀与 pack 对不上   —— 迁卡时最容易漏改（CORE 包的卡以 G 开头）
  E3 卡片引用了不存在的 pack —— 这张卡永远装不进来
  E4 必填字段缺失或为空      —— id/t/when/sig/act/pack
  E5 pack 键重复或缺 label
  W1 有 fail 但 failReal 不同步 —— 页面读 failReal，装配器读 fail，两边会不一致
  W2 空包（一张卡都没有）
  W3 ask 里带了句号以外的长篇说理 —— 批注该短，说理进报告
返回码：有 E 级问题 → 1；只有 W → 0。
"""
import json
import re
import sys
from collections import Counter
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT = HERE.parent / "rules" / "规则库.json"
REQUIRED = ("id", "t", "when", "sig", "act", "pack")


def load(p: Path):
    t = p.read_text(encoding="utf-8")
    if not t.lstrip().startswith(("{", "[")):        # 工作台导出的 rules.js
        m = re.search(r"=\s*(\{.*\})\s*;?\s*$", t, re.S)
        if not m:
            raise SystemExit(f"认不出这个文件的格式：{p}")
        t = m.group(1)
    return json.loads(t)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    stats_only = "--stats" in sys.argv
    path = Path(args[0]) if args else DEFAULT
    if not path.is_file():
        raise SystemExit(f"找不到规则库：{path}")
    lib = load(path)
    cards, packs = lib.get("cards", []), lib.get("packs", [])
    pack_keys = [p.get("k") for p in packs]

    print(f"规则库：{path}")
    print(f"版本 {lib.get('version','?')} · 更新 {lib.get('updated','?')} · "
          f"{len(cards)} 张卡 · {len(packs)} 个包\n")

    for p in packs:
        n = sum(1 for c in cards if c.get("pack") == p.get("k"))
        print(f"  {str(p.get('k')):6s} {str(p.get('label','')):34s} {n:3d}")
    if stats_only:
        return 0

    errs, warns = [], []

    dup = {k: v for k, v in Counter(c.get("id") for c in cards).items() if v > 1}
    for k, v in dup.items():
        errs.append(f"E1 卡片 id 重复 {v} 次：{k}"
                    f"（装配时只会留下最后一张，前面的静默丢失）")

    for c in cards:
        cid, pk = c.get("id", ""), c.get("pack", "")
        want = "G" if pk == "CORE" else f"{pk}-"
        if pk == "CORE":
            if not cid.startswith("G"):
                errs.append(f"E2 {cid} 在 CORE 包，id 应以 G 开头")
        elif not cid.startswith(want):
            errs.append(f"E2 {cid} 的 id 前缀与 pack「{pk}」对不上（应形如 {want}N）")
        if pk not in pack_keys:
            errs.append(f"E3 {cid} 引用了不存在的 pack「{pk}」——这张卡永远装不进来")
        for f in REQUIRED:
            if not str(c.get(f, "")).strip():
                errs.append(f"E4 {cid or '(无 id)'} 缺字段或字段为空：{f}")

    pdup = {k: v for k, v in Counter(pack_keys).items() if v > 1}
    for k, v in pdup.items():
        errs.append(f"E5 pack 键重复 {v} 次：{k}")
    for p in packs:
        if not str(p.get("label", "")).strip():
            errs.append(f"E5 pack「{p.get('k')}」没有 label")

    for c in cards:
        fail, real = c.get("fail", ""), c.get("failReal", "")
        if fail and fail != real and fail != "无。":
            warns.append(f"W1 {c.get('id')} 的 fail 与 failReal 不一致"
                         f"（页面读 failReal，装配器读 fail，两边口径会不同）")
    for p in packs:
        if not any(c.get("pack") == p.get("k") for c in cards):
            warns.append(f"W2 包「{p.get('k')} {p.get('label','')}」一张卡都没有")
    for c in cards:
        ask = c.get("ask", "")
        if len(ask) > 400:
            warns.append(f"W3 {c.get('id')} 的 ask 有 {len(ask)} 字，"
                         f"批注示例宜短——说理应写进审查报告，不进 Word")

    print()
    if errs:
        print(f"❌ {len(errs)} 个必须修的问题：")
        for e in errs:
            print("   ", e)
    if warns:
        print(f"\n⚠️  {len(warns)} 个提示：")
        for w in warns[:20]:
            print("   ", w)
        if len(warns) > 20:
            print(f"    …另有 {len(warns)-20} 条")
    if not errs and not warns:
        print("✅ 规则库没问题。")
    elif not errs:
        print("\n✅ 没有必须修的问题。")
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main())
