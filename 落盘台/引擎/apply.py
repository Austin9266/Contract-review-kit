#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""apply.py —— 一次性批量执行整张修订/批注清单（document.xml 只读写一次）。

用法:
    python3 scripts/apply.py work/unpacked --plan plan.json [--dry-run]

plan.json:
{
  "edits": [
    {"find": "××市仲裁委员会", "replace": "××仲裁委员会", "note": "序号4"},
    {"find": "并按日万分之五支付违约金", "replace": ""},
    {"find": "验收合格之日起", "insert_after": " 30 日内", "occurrence": 2},
    {"after_para": 37, "insert_paragraph": "9.4 ……"},
    {"delete_para": 41},
    {"comment": "提请注意：……", "anchor": "甲方应在收到发票后 15 日内付款"},
    {"comment": "目录未修改，请定稿后更新目录域。", "para": 2}
  ]
}

规则:
  * 段落序号一律按**来件原始编号**（prep 之后 wredline.py --list 看到的那个），
    脚本内部会把改变段落数的操作放到最后、按序号从大到小执行，序号不会串位。
  * 逐条报告 OK/FAIL 并继续往下跑 —— 一次就能看到所有锚定失败，不必一条条试。
  * --dry-run 只校验锚定能否命中，不写文件。
  * 时间戳与署名仍留 PENDING，由 stamp.py 统一分配。
"""
import argparse
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import random                   # noqa: E402

import wredline as R           # noqa: E402
import wcomment as C           # noqa: E402
from wpara import tokenize, index_map, find_span, slice_tokens, render  # noqa: E402

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



def label(e):
    if "comment" in e:
        where = f'--para {e["para"]}' if "para" in e else f'锚「{str(e.get("anchor"))[:18]}…」'
        return f'批注 {where}'
    if "find" in e:
        if "insert_after" in e:
            return f'插入 「{e["find"][:14]}…」后 +「{str(e["insert_after"])[:14]}」'
        return (f'删除 「{e["find"][:20]}…」' if e.get("replace") == ""
                else f'替换 「{e["find"][:16]}…」→「{str(e.get("replace"))[:16]}…」')
    if "insert_paragraph" in e:
        return f'新增段（在第 {e["after_para"]} 段之后）'
    if "delete_para" in e:
        return f'删除第 {e["delete_para"]} 段'
    return "未知操作"


def do_find_edit(doc, e):
    wid = R.new_id(doc)
    pi, pm, (head, ppr, toks), span = R.locate(doc, e["find"], e.get("occurrence", 1))
    i0, o0, i1, o1 = span
    pre, mid, post = slice_tokens(toks, i0, o0, i1, o1)
    rpr = toks[i0].rpr
    if "insert_after" in e:
        middle = render(mid) + R.ins_block(wid, rpr, e["insert_after"])
    else:
        middle = R.del_block(wid, mid) + R.ins_block(wid + 1, rpr, e.get("replace", ""))
    body_new = (render(toks[:i0]) + render(pre) + middle + render(post)
                + render(toks[i1 + 1:]))
    return doc[:pm.start()] + head + ppr + body_new + "</w:p>" + doc[pm.end():], f"段落 {pi}"


def do_insert_paras(doc, n, texts):
    """在第 n 段之后一次性插入 k 段。链式打 ins 标记；锚点段的段落标记已带
    来件既有 w:ins（他人/前次插入）时，锚点段不动、新段各自携带 w:ins 段落标记，
    拒绝修订时整链合并可完整还原来件 —— 逻辑统一在 wredline.insert_paras。"""
    rnd = random.Random(hash((n, tuple(texts))) & 0xFFFFFFFF)
    return R.insert_paras(doc, n, texts, rnd)


def do_delete_para(doc, e):
    n = e["delete_para"]
    return R.delete_para(doc, n), f"段落 {n}"


def check_anchor(doc, needle, occurrence):
    seen = 0
    for pm in R.PARA.finditer(doc):
        _, _, body = R.split_para(pm.group(0))
        text, _ = index_map(tokenize(body))
        seen += text.count(needle)
    if seen < occurrence:
        raise SystemExit(f"全文只找到 {seen} 处，取不到第 {occurrence} 处")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("unpacked")
    ap.add_argument("--plan", required=True)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--author", default=_CFG_AUTHOR)
    ap.add_argument("--initials", default=_CFG_INITIALS)
    a = ap.parse_args()

    unp = Path(a.unpacked)
    docp = unp / "word" / "document.xml"
    doc = docp.read_text(encoding="utf-8")
    plan = json.loads(Path(a.plan).read_text(encoding="utf-8"))
    edits = plan["edits"] if isinstance(plan, dict) else plan

    # 只有 insert_paragraph 会改变段落总数 → 放到最后、按段号从大到小执行，序号不串位；
    # 同一个 after_para 的多条合并成一组一次插入（否则会给同一段落标记打两个 w:ins）。
    groups = {}
    for i, e in enumerate(edits):
        if "insert_paragraph" in e:
            groups.setdefault(e["after_para"], []).append(i)
    order = [i for i, e in enumerate(edits) if "insert_paragraph" not in e]
    ins_groups = sorted(groups.items(), key=lambda kv: -kv[0])

    okn = failn = 0
    fails = []
    for i in order:
        e = edits[i]
        tag = f"[{i + 1}/{len(edits)}] {label(e)}"
        try:
            if "comment" in e:
                if a.dry_run:
                    if "para" not in e:
                        check_anchor(doc, e["anchor"], e.get("occurrence", 1))
                    where = "（dry-run）"
                else:
                    cid = C.next_id(C.read(unp / "word" / "comments.xml", ""))
                    C.add_comment_part(unp, cid, e["comment"], a.author, a.initials)
                    ref_style, _ = C.comment_styles(unp)
                    doc, where = C.anchor_in_document(
                        doc, cid, e.get("anchor"), e.get("occurrence", 1), e.get("para"),
                        ref_style)
            elif "find" in e:
                if a.dry_run:
                    check_anchor(doc, e["find"], e.get("occurrence", 1))
                    where = "（dry-run）"
                else:
                    doc, where = do_find_edit(doc, e)
            elif "delete_para" in e:
                doc, where = (doc, "（dry-run）") if a.dry_run else do_delete_para(doc, e)
            else:
                raise SystemExit("plan 条目缺少 find / insert_paragraph / delete_para / comment")
            okn += 1
            print(f"  OK   {tag} @ {where}")
        except SystemExit as ex:
            failn += 1
            fails.append((i + 1, label(e), str(ex)))
            print(f"  FAIL {tag} :: {ex}")

    for after, idxs in ins_groups:
        texts = [edits[i]["insert_paragraph"] for i in idxs]
        tag = f"[{'/'.join(str(i + 1) for i in idxs)}] 新增 {len(texts)} 段（在第 {after} 段之后）"
        try:
            if a.dry_run:
                if after >= len(list(R.PARA.finditer(doc))):
                    raise SystemExit(f"段落序号 {after} 超界")
                where = "（dry-run）"
            else:
                doc, where = do_insert_paras(doc, after, texts)
            okn += len(idxs)
            print(f"  OK   {tag} @ {where}")
        except SystemExit as ex:
            failn += len(idxs)
            fails.append((idxs[0] + 1, tag, str(ex)))
            print(f"  FAIL {tag} :: {ex}")

    if not a.dry_run:
        # 切分带 rPrChange 的 run 会复制出重复的修订 id（第三方格式修订常见于 WPS 来件），
        # 写回前统一去重 —— 只改 id 数字，作者/时间/旧格式不动
        docp.write_text(R.dedupe_change_ids(doc), encoding="utf-8")

    print(f"\n[汇总] 成功 {okn} / 失败 {failn}"
          + ("（dry-run，未写文件）" if a.dry_run else ""))
    if fails:
        print("[待处理] 逐条修正锚定串后，把失败项单独重跑（用 wredline.py --grep 看真实原文）：")
        for n, lb, msg in fails:
            print(f"  #{n} {lb} :: {msg}")
        sys.exit(1)


if __name__ == "__main__":
    main()
