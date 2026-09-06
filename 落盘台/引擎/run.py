#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""run.py —— 一条命令跑完全程：准备 → 批量落盘 → 署名与时间 → 打包体检（→ 回转 .doc）。

用法:
    python3 scripts/run.py --input "来件.docx" --plan plan.json \\
        --name "XX合同（修订稿）" --start "2026-08-21T09:20" \\
        [--end "..."] [--work work/] [--author ...] [--initials ...] \\
        [--windows "09:00-18:00"] [--weekend] [--seed 7] [--dry-run]

--dry-run 只做"准备 + 锚定校验"，不写修订、不出文件；确认清单能全部命中后再正式跑。
任何一步失败即停止并返回非零退出码。
"""
import argparse
import subprocess
import sys
from pathlib import Path

# —— 署名等口径从工具包根目录的 配置.json 读；读不到就用内置默认值 ——
try:
    from config import AUTHOR as _CFG_AUTHOR, INITIALS as _CFG_INITIALS, \
        WINDOWS as _CFG_WINDOWS, WEEKEND as _CFG_WEEKEND, warn_if_unconfigured
except Exception:  # 单独拷出去用时也能跑
    import sys as _sys, json as _json, os as _os
    from pathlib import Path as _Path
    _CFG_AUTHOR, _CFG_INITIALS = "审查人", "审"
    _CFG_WINDOWS, _CFG_WEEKEND = "09:00-18:00", False
    def warn_if_unconfigured():
        pass


HERE = Path(__file__).resolve().parent


def step(title, cmd):
    print(f"\n===== {title} =====")
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print(f"[STOP] 「{title}」失败，流程中止。")
        sys.exit(r.returncode)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--plan", required=True)
    ap.add_argument("--name", required=True)
    ap.add_argument("--start", required=True)
    ap.add_argument("--end")
    ap.add_argument("--work", default="work")
    ap.add_argument("--author", default=_CFG_AUTHOR)
    ap.add_argument("--initials", default=_CFG_INITIALS)
    ap.add_argument("--windows", default=_CFG_WINDOWS)
    ap.add_argument("--weekend", action="store_true")
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    py = sys.executable
    unp = str(Path(a.work) / "unpacked")
    wk = ["--weekend"] if a.weekend else []

    step("1/4 准备（解包 + 记录基线）",
         [py, str(HERE / "prep.py"), a.input, "--work", a.work])

    step("2/4 批量落盘修订与批注" + ("（dry-run）" if a.dry_run else ""),
         [py, str(HERE / "apply.py"), unp, "--plan", a.plan,
          "--author", a.author, "--initials", a.initials]
         + (["--dry-run"] if a.dry_run else []))

    if a.dry_run:
        print("\n[dry-run 结束] 清单全部命中，可去掉 --dry-run 正式执行。")
        return

    step("3/4 统一署名与北京时间戳",
         [py, str(HERE / "stamp.py"), unp, "--author", a.author, "--initials", a.initials,
          "--start", a.start, "--windows", a.windows, "--seed", str(a.seed)]
         + (["--end", a.end] if a.end else []) + wk)

    step("4/4 打包 + 体检" + ("（+ 回转 .doc）" if a.input.lower().endswith(".doc") else ""),
         [py, str(HERE / "finish.py"), "--work", a.work, "--name", a.name,
          "--author", a.author, "--windows", a.windows] + wk)


if __name__ == "__main__":
    main()
