#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""config.py —— 全套工具的配置读取（署名、时段、交付口径）。

只用标准库。任何脚本都可以 `from config import CFG`，拿不到配置也不会崩——
读不到就用内置默认值，并在 stderr 提示一次「还没配置署名」。

查找顺序（先命中先用）：
  1. 环境变量 CONTRACT_REVIEW_CONFIG 指向的 json 文件
  2. 沿本文件所在目录逐级向上找 配置.json —— 也就是工具包根目录那份。
     在工具包里跑，就用工具包自己的配置，改了立刻生效。
  3. ~/.contract-review/配置.json —— 技能被 安装.sh 拷进 ~/.claude/skills/ 之后，
     上一条找不到工具包了，就落到这份。
  4. 内置默认值

所以：**在工具包里跑改工具包的 配置.json，装出去之后改 ~/.contract-review/配置.json。**
拿不准就跑 `python3 config.py`，它会打印这次是从哪读的。

命令行用法：
    python3 config.py            # 打印当前生效的配置与来源
    python3 config.py --path     # 只打印配置文件路径
"""
import json
import os
import sys
from pathlib import Path

DEFAULTS = {
    "署名": "审查人",
    "署名缩写": "审",
    "工作时段": "09:00-18:00",
    "含周末": False,
    "交付文件名后缀": "（修订稿）",
    "我方默认立场": "甲方",
}

_PLACEHOLDER_SIGNS = {"审查人", "", None}


def _candidates():
    env = os.environ.get("CONTRACT_REVIEW_CONFIG")
    if env:
        yield Path(env).expanduser()
    here = Path(__file__).resolve().parent
    home = Path.home()
    for d in [here, *here.parents]:
        if d == home:            # 家目录不算「工具包根目录」，跳过，交给下面那条
            continue
        yield d / "配置.json"
    yield home / ".contract-review" / "配置.json"


def load():
    """返回 (配置 dict, 来源路径 or None)。"""
    cfg = json.loads(json.dumps(DEFAULTS))  # deep copy
    for p in _candidates():
        try:
            if not p.is_file():
                continue
            raw = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        for k, v in raw.items():
            if k.startswith("_"):
                continue
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update({kk: vv for kk, vv in v.items() if not kk.startswith("_")})
            else:
                cfg[k] = v
        return cfg, p
    return cfg, None


CFG, CFG_PATH = load()

AUTHOR = CFG.get("署名") or DEFAULTS["署名"]
INITIALS = CFG.get("署名缩写") or DEFAULTS["署名缩写"]
WINDOWS = CFG.get("工作时段") or DEFAULTS["工作时段"]
WEEKEND = bool(CFG.get("含周末", False))
SUFFIX = CFG.get("交付文件名后缀", DEFAULTS["交付文件名后缀"])
ROLE = CFG.get("我方默认立场") or DEFAULTS["我方默认立场"]

_warned = False


def warn_if_unconfigured():
    """署名还是占位值时，往 stderr 提醒一次。不阻断运行。"""
    global _warned
    if _warned or AUTHOR not in _PLACEHOLDER_SIGNS:
        return
    _warned = True
    where = CFG_PATH or "（未找到配置文件）"
    print(
        f"[提示] 署名仍是默认占位值「{AUTHOR}」，交付件上会这样署名。\n"
        f"       改这里：{where}\n"
        f"       或本次运行加 --author \"某某律所 张三\" --initials \"张\"",
        file=sys.stderr,
    )


if __name__ == "__main__":
    if "--path" in sys.argv:
        print(CFG_PATH or "")
    else:
        print(f"配置来源：{CFG_PATH or '内置默认值（没找到 配置.json）'}")
        print(json.dumps(CFG, ensure_ascii=False, indent=2))
        warn_if_unconfigured()
