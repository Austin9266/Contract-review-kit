#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一键落盘 —— 来件 ＋ 落盘块 → 带修订与批注的交付件，一条命令跑完。

这是落盘台那套确定性代码的无界面版本，**给 AI 自主执行用**：
解析落盘块 → 逐条定位 → 自动收敛存疑条目 → 生成 plan → 批量落盘 → 署名与北京时间 →
打包 → C0–C7 体检 → 交付，全程一个进程，.docx 约 1 秒、.doc 约 4 秒。
落盘台界面上能做的事（改口径、改字段、指认段落、跳过、导出落盘块），这里都有对应参数，
所以 AI 不需要人点一下也能把一件事做完。

    # 0）先看来件情况，决定直落还是退回纯文字（不写任何文件）
    python3 一键落盘.py --input 来件.docx --inspect --json -

    # 1）试跑：只定位不写文件，把存疑条目和候选段落一次列清
    python3 一键落盘.py --input 来件.docx --block 落盘块.txt --dry-run --json -

    # 2）自主跑：能自动收敛的自动收敛（落不了刀的修订降级为批注、高把握的候选自动指认），
    #    收不了的列进 needs_human
    python3 一键落盘.py --input 来件.docx --block 落盘块.txt --auto \
        --start "2026-09-01T09:20" --json 报告.json --emit-block 实际落盘块.txt

    # 3）你自己判断完，用一个 patch 文件把决定一次交回去（改口径/改原文/指认/跳过）
    python3 一键落盘.py --input 来件.docx --block 落盘块.txt --patch 决定.json --json -

    # 4）抄错原文时先查来件真实原文
    python3 一键落盘.py --input 来件.docx --search "履约保证金"

patch 文件（uid → 决定，一个文件把所有决定交回去）：

    {"3": {"kind": "note", "note": "请确认审核时限。"},
     "7": {"src": "更短的一句原文"},
     "9": {"pick": {"para": 132}},
     "11": {"skip": true}}

条目编号（uid）＝落盘块里条目出现的先后顺序，1 起；不认 @修订 后面那个允许重复的序号。
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent


def _dir_with(name, cands):
    for c in cands:
        if (c / name).exists():
            return c
    raise SystemExit(f"[FAIL] 找不到 {name}，请把本脚本放在落盘台目录或技能 scripts/ 目录下。")


CORE = _dir_with("planner.py", [HERE / "核心", HERE / "core", HERE.parent / "核心"])
ENGINE = _dir_with("prep.py", [HERE / "引擎", HERE, HERE.parent / "引擎"])
sys.path.insert(0, str(CORE))
sys.path.insert(0, str(ENGINE))

from locator import Doc                       # noqa: E402
import parser as blockparser                  # noqa: E402
import planner                                 # noqa: E402

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


PY = sys.executable
LOG = []
QUIET = False


def log(line=""):
    LOG.append(line)
    if not QUIET:
        print(line, flush=True)


def run(cmd):
    """跑一个引擎脚本，输出并进日志。→ 返回码"""
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                         text=True, bufsize=1)
    for line in p.stdout:
        log(line.rstrip())
    p.wait()
    return p.returncode


def uid_of(it):
    return str(it.get("uid") or it["no"])


# ── 来件体检（决定直落还是退回纯文字） ────────────────────────────────────
TBL = re.compile(r"<w:tbl[ >]|</w:tbl>")
PARA_AT = re.compile(r"<w:p[ >]")


def table_para_count(xml: str) -> int:
    """有多少段落长在表格里（这些段落的问题按 SOP 归人工，不进落盘块）。"""
    depth, n = 0, 0
    marks = sorted([(m.start(), m.group(0)) for m in TBL.finditer(xml)]
                   + [(m.start(), "P") for m in PARA_AT.finditer(xml)])
    for _, kind in marks:
        if kind == "P":
            n += depth > 0
        elif kind.startswith("</"):
            depth = max(0, depth - 1)
        else:
            depth += 1
    return n


def do_prep(src: Path, work: Path):
    work.mkdir(parents=True, exist_ok=True)
    if run([PY, str(ENGINE / "prep.py"), str(src), "--work", str(work)]) != 0:
        raise SystemExit("[FAIL] 解包失败，见上面的输出。")
    base = json.loads((work / "baseline.json").read_text(encoding="utf-8"))
    doc = Doc((work / "unpacked" / "word" / "document.xml").read_text(encoding="utf-8"))
    return base, doc


def inspect(src: Path, work: Path, base: dict, doc: Doc) -> dict:
    xml = (work / "unpacked" / "word" / "document.xml").read_text(encoding="utf-8")
    tp = table_para_count(xml)
    body = [t for t in doc.paras if t.strip()]
    info = {
        "path": str(src), "name": src.name, "ext": src.suffix,
        "paras": len(doc.paras), "body_paras": len(body),
        "chars": sum(len(t) for t in body),
        "table_paras": tp,
        "table_ratio": round(tp / max(1, len(doc.paras)), 3),
        "pre_authors": base.get("pre_authors", []),
        "pre_marks": len(base.get("pre_stamps", [])),
        "soffice": bool(shutil.which("soffice")),
        "head": body[:3],
    }
    flags = []
    if info["ext"].lower() == ".doc" and not info["soffice"]:
        flags.append("来件是 .doc 但这台机器没有 LibreOffice，转不了格式 —— 换一台有的机器跑，"
                     "或只交落盘块纯文字让人在落盘台上手落")
    if info["pre_authors"]:
        flags.append("来件已带第三方修订/批注（作者：" + "、".join(info["pre_authors"]) +
                     "）：本次只叠加不改动；要改到别人修订块上的条目会被拒绝，"
                     "用 --on-blocked note 降级为批注，或单列给人工")
    if info["table_ratio"] > 0.5:
        flags.append("过半段落长在表格里：表格内的问题不进落盘块，另列给人工")
    info["flags"] = flags
    return info


# ── 落盘块与决定 ──────────────────────────────────────────────────────────
def read_block(arg: str) -> str:
    if arg == "-":
        return sys.stdin.read()
    p = Path(arg).expanduser()
    return p.read_text(encoding="utf-8") if p.exists() else arg


PATCH_FIELDS = ("loc", "src", "dst", "anchor", "note", "body", "why")


def refinalize(it):
    it["occ"] = str(it.get("occ") or "")
    it["error"] = ""
    blockparser.finalize(it)
    return it


def apply_patch(items, patch, forces, skips):
    """把「你的决定」一次性打进条目：改口径、改字段、指认、跳过。"""
    notes = []
    by = {uid_of(it): it for it in items}
    for uid, spec in (patch or {}).items():
        uid = str(uid)
        it = by.get(uid)
        if it is None:
            notes.append(f"patch：没有 #{uid} 这一条，已忽略")
            continue
        if not isinstance(spec, dict):
            notes.append(f"patch #{uid}：要给一个对象，已忽略")
            continue
        if spec.get("skip"):
            if uid not in skips:
                skips.append(uid)
        p = spec.get("pick")
        if p:
            forces[uid] = {"para": int(p["para"]), "start": int(p.get("start", 0)),
                           "end": int(p.get("end", 0))}
        touched = False
        if spec.get("kind"):
            k = spec["kind"]
            if k not in ("edit", "del", "delpara", "insert", "note"):
                notes.append(f"patch #{uid}：认不出类型 {k}")
                continue
            it["kind"] = k
            it["kw"] = blockparser.KIND_LABEL.get(k, it.get("kw", ""))
            touched = True
        for k in PATCH_FIELDS:
            if k in spec:
                it[k] = (spec[k] or "").strip()
                touched = True
        if "occ" in spec:
            it["occ"] = str(spec["occ"] or "")
            touched = True
        if touched:
            refinalize(it)
            if it["error"]:
                notes.append(f"patch #{uid}：{it['error']}")
    return notes


def clip(s: str, n: int) -> str:
    s = (s or "").strip()
    return s if len(s) <= n else s[:n] + "…"


def auto_note_text(it) -> str:
    """修订落不了刀时，自动写一句符合本方口径的批注（占位口径；AI 应当用 patch 写自己的话）。"""
    if it["kind"] == "del" or not it.get("dst"):
        return f'建议删除「{clip(it["src"], 24)}」，请核实。'
    return f'建议改为「{clip(it["dst"], 24)}」，请核实。' 


BLOCKED = ("不能直接下刀", "拒绝叠改", "既有修订")


def auto_settle(doc, items, forces, skips, on_blocked, auto_pick, actions):
    """自动收敛存疑条目：落不了刀的修订降级为批注／跳过；高把握的候选自动指认。
    最多三轮（改一条会影响后面条目的检索），仍收不住的留给 needs_human。"""
    results = planner.resolve_all(doc, items, forces, skips)
    for _ in range(3):
        by = {uid_of(it): it for it in items}
        changed = False
        for r in results:
            if r["skip"] or r["status"] in ("ok", "check"):
                continue
            uid, it, msg = r["uid"], by[r["uid"]], (r["msg"] or "")
            if any(w in msg for w in BLOCKED) and it["kind"] in ("edit", "del") \
                    and on_blocked != "none":
                if on_blocked == "skip":
                    if uid not in skips:
                        skips.append(uid)
                    actions.append({"uid": uid, "do": "跳过",
                                    "why": "原文落在来件既有修订块里，修订下不了刀"})
                else:
                    it["kind"] = "note"
                    it["kw"] = "批注"
                    it["anchor"] = it.get("anchor") or it.get("src") or ""
                    it["note"] = it.get("note") or auto_note_text(it)
                    refinalize(it)
                    actions.append({"uid": uid, "do": "降级为批注",
                                    "why": "原文落在来件既有修订块里，修订下不了刀",
                                    "note": it["note"]})
                changed = True
                continue
            if auto_pick and r["cands"]:
                c0 = r["cands"][0]
                c1 = r["cands"][1] if len(r["cands"]) > 1 else None
                if c0["score"] >= auto_pick and (not c1 or c0["score"] - c1["score"] >= 0.08):
                    forces[uid] = {"para": c0["para"], "start": c0["start"], "end": c0["end"]}
                    actions.append({"uid": uid, "do": "自动指认",
                                    "para": c0["para"], "score": round(c0["score"], 3)})
                    changed = True
        if not changed:
            break
        results = planner.resolve_all(doc, items, forces, skips)
    return results


def view(results, doc: Doc):
    out = []
    for r in results:
        p = r.get("preview")
        out.append({
            "uid": r["uid"], "no": r["no"], "label": r["label"], "kind": r["kind"],
            "loc": r["loc"], "summary": r["summary"], "status": r["status"],
            "msg": r["msg"], "para": r["para"], "skip": bool(r["skip"]),
            "why": r["why"], "entries": len(r["entries"]),
            "preview": p and {"before": p["text"][max(0, p["s"] - 40):p["s"]],
                              "hit": p["text"][p["s"]:p["e"]][:200],
                              "after": p["text"][p["e"]:p["e"] + 40],
                              "dst": p["dst"][:300], "mode": p["mode"]},
            "cands": [{"para": c["para"], "score": round(c["score"], 3),
                       "start": c["start"], "end": c["end"], "text": c["text"][:300]}
                      for c in (r["cands"] or [])[:5]],
        })
    return out


def parse_picks(vals):
    forces = {}
    for v in vals or []:
        uid, _, rhs = v.partition("=")
        bits = [b for b in rhs.split(":") if b != ""]
        if not uid.strip() or not bits:
            raise SystemExit(f"[FAIL] --pick 写法不对：{v}（应为 uid=段号 或 uid=段号:起:止）")
        forces[uid.strip()] = {"para": int(bits[0]),
                               "start": int(bits[1]) if len(bits) > 2 else 0,
                               "end": int(bits[2]) if len(bits) > 2 else 0}
    return forces


def parse_skips(vals):
    out = []
    for v in vals or []:
        out += [x.strip() for x in re.split(r"[,，\s]+", v) if x.strip()]
    return out


# ── 主流程 ───────────────────────────────────────────────────────────────
def main():
    global QUIET
    ap = argparse.ArgumentParser(description="一键落盘：来件＋落盘块 → 修订与批注")
    ap.add_argument("--input", required=True, help="来件 .doc/.docx")
    ap.add_argument("--block", help="落盘块：文件路径、'-'（读标准输入）或直接给整段文字")
    ap.add_argument("--inspect", action="store_true", help="只看来件情况，不落盘")
    ap.add_argument("--search", help="在来件里搜一段文字，看真实原文与段号")
    ap.add_argument("--dry-run", action="store_true", help="只定位与校验，不写任何文件")
    ap.add_argument("--start", help="时间线起点（北京时间，如 2026-09-01T09:20），默认此刻")
    ap.add_argument("--end", help="时间线止点")
    ap.add_argument("--windows", default=_CFG_WINDOWS)
    ap.add_argument("--weekend", action="store_true")
    ap.add_argument("--author", default=_CFG_AUTHOR)
    ap.add_argument("--initials", default=_CFG_INITIALS)
    ap.add_argument("--name", help="交付文件名（不含后缀），默认「来件名（修订稿）」")
    ap.add_argument("--outdir", help="交付目录，默认来件所在目录")
    ap.add_argument("--work", help="中间产物目录，默认系统临时目录")
    ap.add_argument("--patch", help="决定文件（JSON）：uid → {kind/src/dst/anchor/note/body/loc/occ/pick/skip}")
    ap.add_argument("--pick", action="append", help="指认位置：uid=段号 或 uid=段号:起:止，可多次")
    ap.add_argument("--skip", action="append", help="跳过条目：uid，逗号分隔，可多次")
    ap.add_argument("--on-blocked", choices=("note", "skip", "none"), default="none",
                    help="修订落在来件既有修订块上时怎么办：note＝降级为批注，skip＝跳过，none＝留着报给人（默认）")
    ap.add_argument("--auto-pick", type=float, default=0.0,
                    help="候选相似度达到这个值且明显领先时自动指认（如 0.9），0＝不自动")
    ap.add_argument("--auto", action="store_true",
                    help="自主模式：等于 --on-blocked note --auto-pick 0.9，并把仍收不住的列进 needs_human")
    ap.add_argument("--emit-block", help="把实际落盘用的落盘块（含你的 patch 与自动改动）写出来")
    ap.add_argument("--strict", action="store_true", help="有任何存疑条目就停下不落盘")
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--quiet", action="store_true", help="只出 JSON，不打日志（日志仍在 JSON 的 log 里）")
    ap.add_argument("--json", help="把完整报告写成 JSON；给 '-' 则打到标准输出")
    a = ap.parse_args()
    QUIET = bool(a.quiet)
    if a.auto:
        a.on_blocked = "note" if a.on_blocked == "none" else a.on_blocked
        a.auto_pick = a.auto_pick or 0.9

    src = Path(a.input).expanduser().resolve()
    if not src.exists():
        raise SystemExit(f"[FAIL] 找不到来件：{src}")
    if src.suffix.lower() not in (".doc", ".docx"):
        raise SystemExit("[FAIL] 只认 .doc 和 .docx。")

    work = Path(a.work).expanduser().resolve() if a.work else \
        Path(tempfile.gettempdir()) / "落盘台工作区" / f"{datetime.now():%Y%m%d-%H%M%S}-{src.stem[:20]}"
    if work.exists():
        shutil.rmtree(work, ignore_errors=True)

    rep = {"ok": False, "mode": "run", "input": {}, "items": [], "stats": {},
           "actions": [], "needs_human": [], "unresolved": [], "checks": [],
           "deliver": "", "work": str(work), "log": LOG}

    log(f"【一键落盘】{datetime.now():%Y-%m-%d %H:%M:%S}")
    log(f"来件：{src}")
    base, doc = do_prep(src, work)
    rep["input"] = inspect(src, work, base, doc)
    for f in rep["input"]["flags"]:
        log("[!] " + f)

    if a.search:
        q = a.search
        hits = [{"para": i, "text": t} for i, t in enumerate(doc.paras) if q in t]
        if not hits:
            nq = re.sub(r"[\s　]", "", q)
            hits = [{"para": i, "text": t} for i, t in enumerate(doc.paras)
                    if nq and nq in re.sub(r"[\s　]", "", t)]
        rep.update(ok=True, mode="search", search={"q": q, "hits": hits[:40]})
        log(f"命中 {len(hits)} 段：")
        for h in hits[:40]:
            log(f"  第 {h['para']} 段：{h['text'][:160]}")
        return done(rep, a, work, keep=False)

    if a.inspect or not a.block:
        rep.update(ok=True, mode="inspect")
        i = rep["input"]
        log(f"段落 {i['paras']}（表格内 {i['table_paras']}）· 约 {i['chars']} 字 · "
            f"既有第三方痕迹 {i['pre_marks']} 处")
        return done(rep, a, work, keep=False)

    items, warns = blockparser.parse(read_block(a.block))
    rep["warns"] = warns
    for w in warns:
        log("[!] " + w)
    if not items:
        raise SystemExit("[FAIL] 落盘块里一条条目也没解析出来。检查 <<<落盘>>> 标记与 @类型 行。")

    forces, skips = parse_picks(a.pick), parse_skips(a.skip)
    patch = json.loads(Path(a.patch).expanduser().read_text(encoding="utf-8")) if a.patch else {}
    pnotes = apply_patch(items, patch, forces, skips)
    for n in pnotes:
        log("[!] " + n)
    rep["patch_notes"] = pnotes

    results = auto_settle(doc, items, forces, skips, a.on_blocked, a.auto_pick, rep["actions"])
    st = planner.stats(results)
    rep["stats"], rep["items"] = st, view(results, doc)
    rep["unresolved"] = [r["uid"] for r in results
                         if r["status"] in ("miss", "ask", "bad") and not r["skip"]]
    rep["needs_human"] = [i for i in rep["items"] if i["uid"] in rep["unresolved"]]

    if a.emit_block:
        Path(a.emit_block).expanduser().write_text(blockparser.dump(items), encoding="utf-8")
        log(f"（实际落盘块：{a.emit_block}）")

    log("")
    log(f"清单 {len(results)} 条：可落盘 {st['ok'] + st['check']}，需指认 {st['ask']}，"
        f"找不到 {st['miss']}，格式有问题 {st['bad']}，跳过 {st['skip']}")
    for act in rep["actions"]:
        log(f"  [自动] #{act['uid']} {act['do']}"
            + (f"：{act.get('note') or act.get('why') or ''}" if act.get('note') or act.get('why') else "")
            + (f"（第 {act['para']} 段，{act['score']:.0%}）" if act.get("para") is not None and "score" in act else ""))
    for r in results:
        tag = {"ok": "命中", "check": "请核对", "ask": "要指认", "miss": "找不到",
               "bad": "格式有问题"}[r["status"]]
        line = f"  #{r['uid']} {r['label']}{r['no']} {r['loc']} [{'跳过' if r['skip'] else tag}]"
        if r["para"] >= 0:
            line += f" 第 {r['para']} 段"
        if r["msg"]:
            line += f" —— {r['msg']}"
        log(line)
        if r["status"] in ("ask", "miss") and not r["skip"]:
            for c in (r["cands"] or [])[:3]:
                log(f"      候选 第 {c['para']} 段 {c['score']:.0%}：{c['text'][:100]}")

    if rep["unresolved"]:
        log("")
        log(f"[!] {len(rep['unresolved'])} 条要你定：#{'、#'.join(rep['unresolved'])}")
        log("    看 needs_human 里的候选段落；--search 核对来件真实原文；"
            "决定写进 --patch（改口径 kind／改原文 src／指认 pick／跳过 skip）重跑。")
        if a.strict:
            log("[STOP] --strict：有存疑条目，不落盘。")
            return done(rep, a, work, keep=True, code=2)

    plan = planner.build_plan(results)
    if not plan["edits"]:
        log("[STOP] 没有一条能落盘的条目。")
        return done(rep, a, work, keep=True, code=1)
    (work / "plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=1),
                                    encoding="utf-8")

    log("")
    log("== 批量落盘" + ("（试跑，不写文件）" if a.dry_run else "") + " ==")
    cmd = [PY, str(ENGINE / "apply.py"), str(work / "unpacked"), "--plan", str(work / "plan.json"),
           "--author", a.author, "--initials", a.initials] + (["--dry-run"] if a.dry_run else [])
    if run(cmd) != 0:
        log("[STOP] 有条目没落上，见上面的 FAIL 行。")
        return done(rep, a, work, keep=True, code=1)
    if a.dry_run:
        rep.update(ok=True, mode="dry")
        log("")
        log("试跑通过：清单能全部锚定。去掉 --dry-run 正式落盘。")
        return done(rep, a, work, keep=True)

    log("")
    log("== 署名与北京时间戳 ==")
    start = a.start or datetime.now().strftime("%Y-%m-%dT%H:%M")
    cmd = [PY, str(ENGINE / "stamp.py"), str(work / "unpacked"), "--author", a.author,
           "--initials", a.initials, "--start", start, "--windows", a.windows,
           "--seed", str(a.seed)] + (["--end", a.end] if a.end else []) \
        + (["--weekend"] if a.weekend else [])
    if run(cmd) != 0:
        log("[STOP] 署名与时间戳失败。")
        return done(rep, a, work, keep=True, code=1)

    name = a.name or (src.stem + "（修订稿）")
    log("")
    log("== 打包 + C0–C7 体检" + ("（+ 回转 .doc）" if src.suffix.lower() == ".doc" else "") + " ==")
    cmd = [PY, str(ENGINE / "finish.py"), "--work", str(work), "--name", name,
           "--author", a.author, "--windows", a.windows] + (["--weekend"] if a.weekend else [])
    code = run(cmd)
    rep["checks"] = [l for l in LOG if re.search(r"\bC[0-7]\b", l)]
    if code != 0:
        log("[STOP] 体检没过，按规矩不出文件；来件本身一个字没动。")
        log("    C2 挂＝有改动没被修订标记包住；C7 挂＝动到了别人的痕迹 —— 这两条别硬发，"
            "改交落盘块纯文字，由人在落盘台上手落。")
        return done(rep, a, work, keep=True, code=1)

    made = work / "deliver" / (name + src.suffix)
    if not made.exists():
        made = work / "deliver" / (name + ".docx")
    outdir = Path(a.outdir).expanduser() if a.outdir else src.parent
    outdir.mkdir(parents=True, exist_ok=True)
    final = outdir / made.name
    n = 1
    while final.exists() and final.resolve() != made.resolve():
        final = outdir / f"{made.stem}({n}){made.suffix}"
        n += 1
    shutil.copy2(made, final)
    spare = work / "deliver" / (name + ".docx")
    if src.suffix.lower() == ".doc" and spare.exists():
        shutil.copy2(spare, outdir / spare.name)
        log(f"[备查] {outdir / spare.name}")
    rep.update(ok=True, deliver=str(final))
    log("")
    log(f"【交付】{final}")
    log(f"修订 {rep['stats']['edits']} 处、批注 {rep['stats']['notes']} 条"
        + (f"；自动处理 {len(rep['actions'])} 条" if rep["actions"] else "")
        + (f"；未落盘 {len(rep['unresolved'])} 条（#{'、#'.join(rep['unresolved'])}）留给人工"
           if rep["unresolved"] else ""))
    return done(rep, a, work, keep=False)


def done(rep, a, work: Path, keep=True, code=0):
    if a.json:
        txt = json.dumps(rep, ensure_ascii=False, indent=1)
        if a.json == "-":
            if not QUIET:
                print("\n<<<JSON>>>")
            print(txt)
        else:
            Path(a.json).expanduser().write_text(txt, encoding="utf-8")
            if not QUIET:
                print(f"（报告：{a.json}）")
    if not keep and not a.work:
        shutil.rmtree(work, ignore_errors=True)
    sys.exit(code)


if __name__ == "__main__":
    main()
