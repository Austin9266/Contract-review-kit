#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""planner.py —— 条目 × 来件 → 逐条定位结果 + 引擎吃的 plan.json。

一条条目在这里会经历：定位（locator）→ 收窄修订范围 → 回验锚定串真的落在
我们要的那一处 → 生成引擎条目。任何一步不确定，状态就标成「需确认」并带上
候选段落，交给界面上的人点一下，绝不猜着往下走。
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from locator import (Doc, locate, narrow, engine_locate, norm, spacing_style,
                     fit_spacing, locate_clause)            # noqa: E402
from parser import KIND_LABEL, summary                            # noqa: E402

NEWPARA = re.compile(
    r"^\s*(?:第[一二三四五六七八九十百]+条|[0-9]+(?:[.．][0-9]+)*[\s、.．]|[（(][一二三四五六七八九十0-9]+[)）])")


def body_paras(body: str):
    """新增条款正文 → 段落列表。空行必分段；换行后像新条款号的也分段，
    否则视为 AI 的软换行，接回上一段。"""
    out = []
    for chunk in re.split(r"\n\s*\n", (body or "").strip()):
        cur = ""
        for line in chunk.splitlines():
            if not cur:
                cur = line.strip()
            elif NEWPARA.match(line):
                out.append(cur)
                cur = line.strip()
            else:
                cur += line.strip()
        if cur:
            out.append(cur)
    return [p for p in out if p]


def engine_occurrence(doc: Doc, s: str, pi: int, sx: int):
    """与引擎一致的「全文第几处」；sx 是落盘当时坐标。"""
    seen = sum(doc.sim(j)[0].count(s) for j in range(pi))
    text = doc.sim(pi)[0]
    k, at = 0, text.find(s)
    while 0 <= at < sx:
        k += 1
        at = text.find(s, at + 1)
    return seen + k + 1


def anchor_ok(doc: Doc, find: str, pi: int, fx: int):
    """→ (occurrence, 是否回验通过)。回验＝引擎按这个串和序号定位，落点与我们一致。"""
    sx = doc.sim_pos(pi, fx)
    if sx is None:
        return 0, False
    occ = engine_occurrence(doc, find, pi, sx)
    return occ, engine_locate(doc, find, occ) == (pi, sx)


def settle_anchor(doc: Doc, pi: int, x: int, y: int, left_only=False, cap=80):
    """回验不过就就近扩锚定串，直到引擎能唯一命中我们要的这一处。
    → (a, b, occurrence) 或 None。"""
    text = doc.paras[pi]
    a, b = x, y
    for _ in range(cap):
        occ, ok = anchor_ok(doc, text[a:b], pi, a)
        if ok:
            return a, b, occ
        if not left_only and b < len(text):
            b += 1
            continue
        if a > 0:
            a -= 1
            continue
        return None
    return None


def resolve_one(doc: Doc, it: dict, prefer_after: int, force=None):
    r = {"uid": it.get("uid") or it["no"], "no": it["no"], "kind": it["kind"], "label": KIND_LABEL.get(it["kind"], "?"),
         "loc": it["loc"], "why": it["why"], "summary": summary(it),
         "status": "miss", "msg": "", "cands": [], "para": -1,
         "entries": [], "preview": None, "skip": False}
    if it.get("error"):
        r["msg"] = it["error"]
        r["status"] = "bad"
        return r

    needle = it["src"] if it["kind"] in ("edit", "del") else it["anchor"]
    if not needle and it["kind"] == "note" and it["loc"] and not force:
        pi = locate_clause(doc, it["loc"])
        if pi is None:
            r["msg"] = f'批注既没有锚点，位置「{it["loc"]}」也对不上任何条款号，挂不上。'
            return r
        r.update(status="ok", para=pi, msg=f'按位置「{it["loc"]}」挂在第 {pi} 段整段上',
                 entries=[{"comment": it["note"], "para": pi, "note": f'批注{it["no"]}'}],
                 preview={"text": doc.paras[pi], "s": 0, "e": 0, "dst": it["note"], "mode": "note"})
        return r
    if force:
        fp, fs, fe = int(force["para"]), int(force["start"]), int(force["end"])
        if fs == fe and it["kind"] in ("edit", "del"):
            # 只指认了段落、没框出位置：在这一段里再找一次最像的那一句
            b = doc.best_in(fp, needle)
            if not b or b[2] < 0.45:
                r["status"] = "ask"
                r["msg"] = (f"第 {fp} 段里也找不到这句原文"
                            + (f"（最像的一段只有 {b[2]:.0%} 像）" if b else "")
                            + "。换一段试试，或者把这一条跳过。")
                return r
            fs, fe = b[0], b[1]
            if b[2] < 0.8:
                r["msg"] = f"人工指认，但这一段与原文只有 {b[2]:.0%} 像 —— 务必看一眼下面的预览再落盘。"
        loc = {"status": "manual", "para": fp, "start": fs, "end": fe,
               "note": r["msg"] or "人工指认", "cands": []}
    else:
        loc = locate(doc, needle, it["loc"], prefer_after, it.get("occ") or 0)
    r["cands"] = loc.get("cands", [])
    r["msg"] = loc.get("note", "")
    if loc["status"] in ("miss", "ambiguous"):
        r["status"] = "ask" if loc["status"] == "ambiguous" else "miss"
        return r

    pi, x, y = loc["para"], loc["start"], loc["end"]
    r["para"] = pi
    r["status"] = {"exact": "ok", "norm": "ok", "fuzzy": "check", "manual": "ok"}[loc["status"]]
    if loc["status"] == "norm" and not r["msg"]:
        r["msg"] = "原文有空格/全半角差异，已按文档里的原样定位"
    k = it["kind"]

    if k in ("edit", "del"):
        if doc.eaten(pi, x, y):
            r["status"] = "ask"
            r["msg"] = "这一处和前面某一条修订改到同一段文字上了，请合并成一条，或指认别处。"
            return r
        if not doc.span_clean(pi, x, y):
            r["status"] = "miss"
            r["msg"] = "这一段原文里混着图片、域代码或已有的第三方修订，不能直接下刀。请把原文缩到更小的一句，或留给人工处理。"
            return r
        style = getattr(doc, "_style", None)
        if style is None:
            style = doc._style = spacing_style(doc.paras)
        dst = fit_spacing(style, it["dst"])
        find, rep, mode = narrow(doc, pi, x, y, dst)
        if not rep and mode == "insert_after":
            find, rep, mode = "", "", "replace"          # 落到"没有实质改动"分支
        if find == rep or norm(doc.paras[pi][x:y])[0] == norm(dst)[0]:
            r["status"] = "check"
            r["skip"] = True
            r["msg"] = "按文档里的原样逐字比对，这一条没有实质改动（多半是原文抄漏或抄错了字），已跳过。"
            return r
        text = doc.paras[pi]
        # 收窄之后的 find 一定落在 [x, y] 这一段里面 —— **必须从 x 开始找**。
        # 从 x-80 开始找会命中同一段里更靠前的同一串字（「汇入下列帐户…并按月对帐」这种，
        # 一段里改两处时第二处必挂：找到的是第一处、而第一处已被前一条修订吃掉，
        # sim_pos 返回 None，最后报成「这段文字在全文里重复太多」）。
        fx = text.find(find, x)
        if fx < 0 or fx + len(find) > y:
            fx = text.find(find)
        if fx < 0:
            r["status"] = "miss"
            r["msg"] = "收窄修订范围之后反而定不住了，请把这一条的原文写短一点、只留改动的那一句。"
            return r
        st = settle_anchor(doc, pi, fx, fx + len(find), left_only=(mode == "insert_after"))
        if not st:
            r["status"] = "ask"
            r["msg"] = ("这一处要改的字，落在前面某一条修订已经改过的地方上了，请把两条合并成一条。"
                        if doc.eaten(pi, fx, fx + len(find))
                        else "这段文字在全文里重复太多，定不到唯一一处，请指认段落。")
            r["cands"] = doc.candidates(needle)
            return r
        a, b, occ = st
        if (a, b) != (fx, fx + len(find)):
            if mode == "insert_after":
                find = text[a:b]
            else:
                rep = text[a:fx] + rep + text[fx + len(find):b]
                find = text[a:b]
            fx = a
        if mode == "replace" and rep:
            if not find[:1].isspace():
                rep = rep.lstrip(" \u3000")
            if not find[-1:].isspace():
                rep = rep.rstrip(" \u3000")
        e = {"find": find, "occurrence": occ,
             "note": f'{r["label"]}{it["no"]} {it["loc"]}'.strip()}
        e["insert_after" if mode == "insert_after" else "replace"] = rep
        r["entries"] = [e]
        if mode != "insert_after":
            r["eat"] = (pi, fx, fx + len(find))     # 这段原文将进入 <w:del>，后面就检索不到了
        r["preview"] = {"text": text, "s": fx, "e": fx + len(find), "dst": rep, "mode": mode}

    elif k == "note":
        text = doc.paras[pi]
        if needle and y > x:
            st = settle_anchor(doc, pi, x, y)
            if not st:
                r["status"] = "ask"
                r["msg"] = "锚点文字在全文里重复太多，请指认段落。"
                return r
            a, b, occ = st
            r["entries"] = [{"comment": it["note"], "anchor": text[a:b], "occurrence": occ,
                             "note": f'批注{it["no"]}'}]
            r["preview"] = {"text": text, "s": a, "e": b, "dst": it["note"], "mode": "note"}
        else:
            r["entries"] = [{"comment": it["note"], "para": pi, "note": f'批注{it["no"]}'}]
            r["preview"] = {"text": text, "s": 0, "e": 0, "dst": it["note"], "mode": "note"}

    elif k == "insert":
        paras = body_paras(it["body"])
        r["entries"] = [{"after_para": pi, "insert_paragraph": t, "note": f'新增{it["no"]}'}
                        for t in paras]
        r["preview"] = {"text": doc.paras[pi], "s": 0, "e": len(doc.paras[pi]),
                        "dst": "\n".join(paras), "mode": "insert"}
        r["msg"] = (r["msg"] + "；" if r["msg"] else "") + f"插在第 {pi} 段之后，共 {len(paras)} 段"

    elif k == "delpara":
        r["entries"] = [{"delete_para": pi, "note": f'删段{it["no"]}'}]
        r["eat"] = (pi, 0, len(doc.paras[pi]))
        r["preview"] = {"text": doc.paras[pi], "s": 0, "e": len(doc.paras[pi]),
                        "dst": "", "mode": "delpara"}
    return r


def resolve_all(doc: Doc, items, forces=None, skips=None):
    """按引擎的实际执行顺序逐条解算：一条修订吃掉的原文，后面的条目就看不见了。"""
    forces, skips = forces or {}, set(skips or [])
    doc.reset_sim()
    out, prefer = [], -1
    for it in items:
        uid = str(it.get("uid") or it["no"])
        r = resolve_one(doc, it, prefer, forces.get(uid))
        if uid in skips:
            r["skip"] = True
        elif r["status"] in ("ok", "check") and r.get("eat"):
            doc.eat(*r["eat"])
        if r["para"] >= 0:
            prefer = r["para"]
        out.append(r)
    return out


def build_plan(results):
    edits = []
    for r in results:
        if r.get("skip") or r["status"] in ("miss", "ask", "bad"):
            continue
        edits.extend(r["entries"])
    return {"edits": edits}


def stats(results):
    d = {"ok": 0, "check": 0, "ask": 0, "miss": 0, "bad": 0, "skip": 0,
         "edits": 0, "notes": 0}
    for r in results:
        if r.get("skip"):
            d["skip"] += 1
            continue
        d[r["status"]] = d.get(r["status"], 0) + 1
        if r["status"] in ("ok", "check"):
            if r["kind"] == "note":
                d["notes"] += 1
            else:
                d["edits"] += len(r["entries"])
    return d
