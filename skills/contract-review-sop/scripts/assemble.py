#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""assemble.py —— 按本件的实际情况，装配本次专用的审查指令。

《审查工作台》（工具包里的那个 HTML）做的就是这件事：138 张规则卡按「通用核 ＋ 情形包 ＋
类型包」装配，只把用得上的那些发给模型。这个脚本是它的无界面版，规则库、文本积木都在
本技能里，所以 AI 自己就能走完「识别 → 装配 → 审查」，不必去粘一份写死的大 prompt。

    # 1）看看有哪些包可选（每个包什么时候装）
    python3 assemble.py --list --json -

    # 2）机械扫一遍来件，拿线索（只是线索，装不装由你判断）
    python3 assemble.py --scan 来件.docx --json -

    # 3）装配本次专用审查指令
    python3 assemble.py --mode contract --track B --type T1 --situ S3,S5 \
        --role 发包人 --note "我方为发包人，一侧 3 家项目公司；本件是施工补充协议；已履行招标" \
        --file 来件.docx --out 本次审查指令.md

    # 4）审查过程中想查某条规则的原话
    python3 assemble.py --card G4

规则库：rules/规则库.json（＝工作台内置库）。你在工作台里改过规则、存成 rules.js 的，
用 --rules 指过来，两边就还是同一套判断口径。
"""
import argparse
import json
import re
import sys
import tempfile
import subprocess
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
import prompt_parts as P                                   # noqa: E402

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


RULES = ROOT / "rules" / "规则库.json"
HINTS = ROOT / "rules" / "识别线索.json"
DOCKIND = [("zc", "公司章程／章程修正案"), ("jy", "股东会／董事会决议"),
           ("hj", "函件（告知·催告·承诺·律师函·复函）"), ("zd", "内部制度／议事规则"),
           ("bs", "合规审查基础信息表／内部报审表"), ("qt", "其他文书")]
DOC_TYPE_PACKS = {"zc": ["T7"], "jy": ["T7"], "hj": ["T8"], "zd": [], "bs": [], "qt": []}


# ── 规则库 ────────────────────────────────────────────────────────────────
def load_rules(path=None):
    p = Path(path).expanduser() if path else RULES
    t = p.read_text(encoding="utf-8").strip()
    if not t.startswith(("{", "[")):                     # 工作台导出的 rules.js
        t = re.sub(r";\s*$", "", t[t.index("=") + 1:].strip())
    lib = json.loads(t)
    for c in lib["cards"]:                                # 与工作台 adoptSource 同规矩
        if c.get("failReal") is None:
            f = (c.get("fail") or "").strip()
            c["failReal"] = "" if re.fullmatch(r"(无。?|——)?", f) else f
    return lib


def label(lib, k):
    return next((p["label"] for p in lib["packs"] if p["k"] == k), k)


def short(lib, k):
    return re.sub(r"^(情形|类型)｜", "", label(lib, k))


def type_keys(lib):
    return [p["k"] for p in lib["packs"] if re.fullmatch(r"T\d+", p["k"])]


def situ_keys(lib):
    return [p["k"] for p in lib["packs"] if re.fullmatch(r"S\d+", p["k"]) and p["k"] != "S6"]


def active_packs(lib, mode, track, type_, situ, doc):
    have = {p["k"] for p in lib["packs"]}
    if mode == "doc":
        ks = ["CORE"] + DOC_TYPE_PACKS.get(doc, [])
    else:
        s = set(situ or [])
        if track == "B":
            s.add("S6")
        ks = ["CORE"] + sorted(s) + ([type_] if type_ else [])
    out = []
    for k in ks:                                          # 去重保序
        if k in have and k not in out:
            out.append(k)
    return out


def active_cards(lib, packs):
    s = set(packs)
    return [c for c in lib["cards"] if c.get("pack") in s]


def rule_block(lib, packs):
    chunks = []
    for k in packs:
        l = [c for c in lib["cards"] if c.get("pack") == k]
        if not l:
            continue
        body = []
        for c in l:
            t = f'【{c["id"]}】{c["t"]}\n  在文本里看到：{c["sig"]}\n  怎么写批注：{c["ask"]}'
            if c.get("failReal"):
                t += f'\n  例外：{c["failReal"]}'
            body.append(t)
        chunks.append(f'〔{label(lib, k)}〕\n' + "\n".join(body))
    return "\n\n".join(chunks)


# ── 来件取文 ─────────────────────────────────────────────────────────────
def doc_text(path: Path) -> str:
    if path.suffix.lower() == ".docx":
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
        xml = re.sub(r"</w:p>", "\n", xml)
        return re.sub(r"<[^>]+>", "", xml)
    if path.suffix.lower() in (".txt", ".md"):
        return path.read_text(encoding="utf-8", errors="ignore")
    raise SystemExit("[FAIL] --scan 只认 .docx/.txt；.doc 请先用 oneshot.py --inspect 转过再扫，"
                     "或直接按你读到的内容自己判断装配哪些包。")


ROLE = "(甲方|乙方|丙方|发包人|承包人|出租方|承租方|买方|卖方|委托方|受托方|出借人|借款人)"
NUMBERED = re.compile(ROLE + r"\s*[（(]?\s*([一二三四五六七八九十]|[0-9]{1,2}|[A-Z])\s*[)）]?\s*[：:、]")
PLURAL = ["各甲方", "各发包人", "甲方各方", "各方甲方", "各甲方公司", "各项目公司",
          "分别开票", "分别结算", "分别支付", "连带清偿", "连带责任", "各自项目"]


def detect_multiparty(text: str):
    """我方一侧是不是两个及以上主体——这一条最容易漏，单独做一个检测器。

    漏了它就等于漏掉 S1 那一整包（分别结算/分别开票/不承担连带责任、签署栏分设、
    管辖落到具体主体、价格口径、设备归属、份数），这几条是实战里反复出现的硬伤。"""
    flat = re.sub(r"\s+", "", text)
    labels = {}
    for role, idx in NUMBERED.findall(flat):
        labels.setdefault(role, set()).add(idx)
    multi_roles = {r: sorted(v) for r, v in labels.items() if len(v) >= 2}
    plural = [{"w": w, "n": flat.count(w)} for w in PLURAL if flat.count(w)]
    # 签署页上同一角色出现多个盖章栏
    seals = len(re.findall(r"甲方[（(]盖章[)）]", flat))
    ev = []
    for r, v in multi_roles.items():
        ev.append({"w": "".join(f"{r}{x}" for x in v[:4]), "n": len(v)})
    ev += plural
    if seals >= 2:
        ev.append({"w": "甲方（盖章）栏", "n": seals})
    return {"multi": bool(multi_roles) or seals >= 2 or len(plural) >= 2,
            "roles": multi_roles, "evidence": ev[:8]}


# 强触发：机械上足够明确的，装配时默认装上；不装就得在【清单证伪】里写理由。
# title＝出现在标题（前 200 字）里就算；sure＝出现一次就算；weak＝要两次以上。
# 宁滥勿缺：装错了顶多多读几条规则，每条都带「例外」告诉你何时不适用；漏了就是实打实的漏审。
STRONG = {
    "S2": {"title": ["框架协议"], "sure": ["具体项目合同", "集采框架"], "weak": ["框架协议", "集采"]},
    "S3": {"title": ["补充协议", "变更协议", "终止协议", "解除协议", "补充合同"],
           "sure": [], "weak": ["补充协议", "变更协议", "终止协议", "解除协议", "原合同", "原协议"]},
    "S4": {"title": [], "sure": ["人民政府", "管委会", "管理委员会", "村民委员会", "村委会",
                                 "股份经济合作社", "村民小组", "街道办事处"], "weak": []},
    "S5": {"title": ["招标", "投标"], "sure": ["中标通知书", "招标文件", "投标人", "投标文件"],
           "weak": ["招标", "投标", "评标"]},
    "S7": {"title": [], "sure": ["美元", "USD", "欧元", "港币", "日元"], "weak": ["汇率", "等值"]},
    "S8": {"title": ["租赁"],
           "sure": ["建设工程规划许可证", "土地使用权", "农村土地", "屋面", "屋顶"],
           "weak": ["厂房", "在建", "不动产", "承包地"]},
}


def strong_triggers(text: str):
    flat = re.sub(r"\s+", "", text)
    title = flat[:200]
    out = {}
    mp = detect_multiparty(text)
    if mp["multi"]:
        out["S1"] = mp["evidence"]
    for k, rule in STRONG.items():
        hit = []
        for w in rule["title"]:
            if w in title:
                hit.append({"w": w + "（在标题里）", "n": flat.count(w)})
        for w in rule["sure"]:
            n = flat.count(w)
            if n >= 1:
                hit.append({"w": w, "n": n})
        for w in rule["weak"]:
            n = flat.count(w)
            if n >= 2:
                hit.append({"w": w, "n": n})
        if hit:
            seen, uniq = set(), []
            for h in hit:
                if h["w"] not in seen:
                    seen.add(h["w"])
                    uniq.append(h)
            out[k] = uniq[:4]
    return out, mp


def typo_scan(path: Path):
    """错别字与衍字：能机械判定的部分先扫掉，并直接出一份可并进落盘块的 @修订。"""
    script = HERE / "typo.py"
    if not script.exists() or path.suffix.lower() not in (".doc", ".docx", ".txt"):
        return None
    blk = Path(tempfile.gettempdir()) / f"错别字块-{path.stem[:20]}.txt"
    r = subprocess.run([sys.executable, str(script), "--input", str(path),
                        "--emit-block", str(blk), "--json", "-"],
                       capture_output=True, text=True)
    if r.returncode != 0:
        return {"ok": False, "err": (r.stderr or r.stdout)[-200:]}
    d = json.loads(r.stdout)
    return {"ok": True, "total": d["total"], "sure": d["sure"],
            "entries": d["block_entries"], "by_kind": d["by_kind"], "block": str(blk)}


def xref_count(path: Path):
    """交叉引用是默认动作：扫的时候顺手把引用清单提出来，省得忘。"""
    script = HERE / "xref" / "extract_index.py"
    if not script.exists() or path.suffix.lower() not in (".docx", ".txt"):
        return None
    out = Path(tempfile.gettempdir()) / f"xref-{path.stem[:20]}.json"
    r = subprocess.run([sys.executable, str(script), str(path), "-o", str(out)],
                       capture_output=True, text=True)
    if r.returncode != 0 or not out.exists():
        return {"ok": False, "err": (r.stderr or r.stdout)[-200:]}
    d = json.loads(out.read_text(encoding="utf-8"))
    return {"ok": True, "path": str(out), "scopes": len(d.get("scopes", [])),
            "index": len(d.get("index", [])), "references": len(d.get("references", []))}


def scan(lib, text: str):
    hints = json.loads(HINTS.read_text(encoding="utf-8"))
    flat = re.sub(r"\s+", "", text)

    def hits(words):
        out = [(w, flat.count(re.sub(r"\s+", "", w))) for w in words]
        return [{"w": w, "n": n} for w, n in out if n]

    res = {"mode": {}, "doc_kind": {}, "type": {}, "situ": {}, "card_hits": []}
    for grp in ("mode", "doc_kind", "type", "situ"):
        for k, words in hints[grp].items():
            h = hits(words)
            if h:
                res[grp][k] = {"score": sum(x["n"] for x in h), "hit": h[:6]}
    # 卡片 sig 里的字面词也扫一遍，作为"为什么建议装这个包"的证据
    for c in lib["cards"]:
        words = [w for w in re.split(r"[、，,；;／/｜|\s]+", c.get("sig", ""))
                 if 2 <= len(re.sub(r"[「」『』\"'“”‘’（）()【】]", "", w)) <= 14]
        h = hits([re.sub(r"[「」『』\"'“”‘’]", "", w) for w in words])[:3]
        if h:
            res["card_hits"].append({"id": c["id"], "pack": c.get("pack"), "t": c["t"],
                                     "hit": h, "score": sum(x["n"] for x in h)})
    res["card_hits"].sort(key=lambda x: -x["score"])
    res["card_hits"] = res["card_hits"][:40]

    mode = "doc" if res["mode"].get("doc", {}).get("score", 0) >= \
        max(3, res["mode"].get("contract", {}).get("score", 0) // 6) else "contract"
    sug = {
        "mode": mode,
        "doc_kind": max(res["doc_kind"], key=lambda k: res["doc_kind"][k]["score"], default=None),
        "type": max(res["type"], key=lambda k: res["type"][k]["score"], default=None),
        "situ": sorted([k for k, v in res["situ"].items() if v["score"] >= 2 and k != "S6"]),
        "track_b_hint": bool(hits(hints["track_b_hints"])),
    }
    strong, mp = strong_triggers(text)
    res["multiparty"] = mp
    res["strong"] = strong
    sug["situ"] = sorted(set(sug["situ"]) | {k for k in strong if k != "S6"})
    res["suggest"] = sug
    res["head"] = [l.strip() for l in text.splitlines() if l.strip()][:6]
    return res


# ── 装配 ─────────────────────────────────────────────────────────────────
def mcp_block(law, tyc):
    on = ([] if not law else ["法律法规检索"]) + ([] if not tyc else ["企业信息查询"])
    off = ([] if law else ["法律法规检索"]) + ([] if tyc else ["企业信息查询"])
    s = ""
    if on:
        s += ("已接入：" + "、".join(on) + "。凡涉及法条编号、法规时效、仲裁机构全称、"
              "企业登记与涉诉信息，一律实查后再写，不得凭记忆。\n")
    if off:
        s += ("未接入：" + "、".join(off) + "。凡需要这些信息才能下判断的，不得凭记忆断言，"
              "改为写进「需经办部门确认事项」，逐项写明查什么、查到不同结果分别怎么处理。"
              "宁可留待核，不许编。\n")
    return s


def fill(t, **kw):
    for k, v in kw.items():
        t = t.replace("«%s»" % k, v)
    left = re.findall(r"«(.*?)»", t)
    assert not left, "还有没填的槽：%s" % left
    return t


def build(lib, a):
    mode = a.mode
    warn = []
    if getattr(a, "_strong", None) and mode == "contract":
        auto = [k for k in a._strong if k not in a.situ and k not in (a.no_situ or [])]
        if auto and not a.no_auto_situ:
            a.situ = sorted(set(a.situ) | set(auto))
            warn.append("已按来件扫描自动装上：" + "；".join(
                f'{k}（{short(lib, k)}：' + "、".join(f'{h["w"]}×{h["n"]}' for h in a._strong[k][:3]) + "）"
                for k in auto))
        elif auto:
            warn.append("⚠ 扫描到强触发但本次未装：" + "；".join(
                f'{k}（{short(lib, k)}）' for k in auto)
                + "。若确不适用，必须在【清单证伪】逐条写明理由。")
        blocked = [k for k in a._strong if k in (a.no_situ or [])]
        if blocked:
            warn.append("⚠ 扫描到强触发、但你明确排除了：" + "；".join(
                f'{k}（{short(lib, k)}：' + "、".join(f'{h["w"]}×{h["n"]}' for h in a._strong[k][:2]) + "）"
                for k in blocked) + "。这几条必须在【清单证伪】里写明为什么不适用。")
    for k in (a.no_situ or []):
        if k in a.situ:
            a.situ.remove(k)
            warn.append(f"按你的要求排除了 {k}（{short(lib, k)}）——记得在【清单证伪】里说明为什么不适用。")
    packs = active_packs(lib, mode, a.track, a.type, a.situ, a.doc)
    cs = active_cards(lib, packs)
    timeline = (f'{a.t1 or "【起】"}—{a.t2 or "【止】"}（北京时间）' if (a.t1 or a.t2)
                else "未指定，从收到文件的时刻起算，落在配置的工作时段内（北京时间），不落周末")
    land = fill(P.LAND, sign=a.sign, timeline=timeline,
                file=a.file or "【来件文件名】", blockspec=P.BLOCKSPEC)
    common = dict(mcp=mcp_block(a.law, a.tyc), org=a.org, seq=P.SEQ, tone=P.TONE, restraint=P.RESTRAINT,
                  pack_names="、".join(short(lib, k) for k in packs), n_cards=str(len(cs)),
                  rules=rule_block(lib, packs), role=a.role,
                  note=(a.note.strip() if a.note and a.note.strip() else "（未提供）"),
                  mcp_tail=("" if (a.law and a.tyc) else "、以及本次因未接工具而留待核实的事项"),
                  land=land)
    if mode == "doc":
        kind = dict(DOCKIND).get(a.doc, "文书")
        text = fill(P.DOC, doc_kind=kind,
                    docpoints=P.DOCPOINTS.get(a.doc) or P.DOCPOINTS["qt"], **common)
    else:
        eff_situ = [k for k in packs if re.fullmatch(r"S\d+", k)]
        track_txt = ("A 类（我方制式模板）——集团法务部统一制定、已内部审核，立场天然偏甲方。"
                     "默认信任模板正文，怀疑一切「填进去的」和「改过的」内容；重点在第六节确定性校验九查。"
                     if a.track == "A" else
                     "B 类（对方模板／双方自拟／政府或垄断行业格式文本）——条款体系可能不完整、立场中性或偏对方。"
                     "不预设任何条款可信，逐条过。")
        text = fill(P.CONTRACT, track_txt=track_txt,
                    type_name=short(lib, a.type) if a.type in type_keys(lib) else "未指定",
                    source_txt=("我方制式模板" if a.track == "A" else "对方模板／双方自拟"),
                    situ_names=("、".join(short(lib, k) for k in eff_situ) or "无"),
                    **common)
    head = (f"（本次装配：{'、'.join(short(lib, k) for k in packs)}；共 {len(cs)} 条规则；"
            f"规则库 {lib.get('version', '?')} {lib.get('updated', '')}）\n")
    if getattr(a, "_xref", None) and a._xref.get("ok"):
        x = a._xref
        head += (f"（交叉引用已机械提取：{x['references']} 处引用、{x['index']} 个条号、"
                 f"{x['scopes']} 个编区 → {x['path']}。**逐条核完再写报告**，"
                 f"第九项要写核了多少处、发现几处。）\n")
    t = getattr(a, "_typo", None)
    if t and t.get("ok"):
        head += (f"（错别字/衍字已机械扫过：{t['total']} 处"
                 + (f"（{'、'.join(f'{k}{v}' for k, v in t['by_kind'].items())}）" if t["by_kind"] else "")
                 + f"，其中 {t['entries']} 处是确定性错误、已写成落盘块 → {t['block']}。"
                 "**这几条直接并进你的落盘块**；其余的自己看一眼再定。）\n"
                 if t["total"] else "（错别字/衍字已机械扫过：没扫到。）\n")
    for w in warn:
        head += f"（{w}）\n"
    return head + "\n" + text, packs, cs


def main():
    ap = argparse.ArgumentParser(description="按本件情况装配审查指令")
    ap.add_argument("--list", action="store_true", help="列出可选的包、文书类型与说明")
    ap.add_argument("--scan", help="机械扫来件，给装配线索（.docx/.txt）")
    ap.add_argument("--card", help="查一条规则的原话，如 G4")
    ap.add_argument("--rules", help="外部规则库（工作台导出的 rules.js 或 json）")
    ap.add_argument("--mode", choices=("contract", "doc"), default="contract")
    ap.add_argument("--track", choices=("A", "B"), default="A", help="A＝我方制式模板，B＝对方／自拟")
    ap.add_argument("--type", help="类型包，如 T1；不给＝未指定")
    ap.add_argument("--situ", help="情形包，逗号分隔，如 S1,S3,S5")
    ap.add_argument("--scan-input", dest="scan_input",
                    help="装配前先扫这份来件：强触发的情形包自动装上，并顺手提取交叉引用清单")
    ap.add_argument("--no-situ", help="明确排除某些情形包（会要求你在清单证伪里说明理由）")
    ap.add_argument("--no-auto-situ", action="store_true",
                    help="扫到强触发也不自动装，只在指令顶部打警示")
    ap.add_argument("--doc", default="zc", help="文书类型：zc/jy/hj/zd/bs/qt")
    ap.add_argument("--role", default="甲方", help="我方角色，如 发包人／买方／股东")
    ap.add_argument("--note", default="", help="经办人给的交易背景，原样写进指令")
    ap.add_argument("--sign", default=_CFG_AUTHOR)
    ap.add_argument("--org", default="我方",
                    help="我方是谁（写进指令开头的身份句），如「一家制造业集团」「某医疗器械公司」；默认「我方」")
    ap.add_argument("--t1", help="时间线起，如 2026-09-01T09:20")
    ap.add_argument("--t2", help="时间线止")
    ap.add_argument("--file", help="来件文件名（写进落盘口径）")
    ap.add_argument("--no-law", dest="law", action="store_false", help="没接法规检索工具")
    ap.add_argument("--no-tyc", dest="tyc", action="store_false", help="没接企业信息查询")
    ap.add_argument("--out", help="把装配好的指令写到文件")
    ap.add_argument("--json", help="机器可读输出；'-' 打到标准输出")
    a = ap.parse_args()
    a.situ = [x.strip() for x in re.split(r"[,，\s]+", a.situ or "") if x.strip()]
    a.no_situ = [x.strip() for x in re.split(r"[,，\s]+", a.no_situ or "") if x.strip()]
    lib = load_rules(a.rules)
    a._strong, a._xref, a._typo = None, None, None
    if a.scan_input:
        f = Path(a.scan_input).expanduser()
        a._strong = strong_triggers(doc_text(f))[0]
        a._xref = xref_count(f)
        a._typo = typo_scan(f)

    def emit(obj, text=None):
        if a.json:
            t = json.dumps(obj, ensure_ascii=False, indent=1)
            print(t) if a.json == "-" else Path(a.json).expanduser().write_text(t, encoding="utf-8")
        elif text is not None:
            print(text)
        else:
            print(json.dumps(obj, ensure_ascii=False, indent=1))

    if a.list:
        return emit({
            "version": lib.get("version"), "updated": lib.get("updated"),
            "cards": len(lib["cards"]),
            "packs": [{"k": p["k"], "label": p["label"], "when": p.get("desc", ""),
                       "cards": sum(1 for c in lib["cards"] if c.get("pack") == p["k"])}
                      for p in lib["packs"]],
            "doc_kinds": [{"k": k, "label": v} for k, v in DOCKIND],
            "how": "先 --scan 拿线索，自己判断后用 --mode/--track/--type/--situ 装配。"
                   "CORE 永远装；情形包按交易实际装；类型包最多一个。",
        })
    if a.card:
        c = next((c for c in lib["cards"] if c["id"].lower() == a.card.lower()), None)
        if not c:
            raise SystemExit(f"[FAIL] 没有这条规则：{a.card}")
        return emit(c, "\n".join(f"{k}：{v}" for k, v in c.items() if v and not k.startswith("_")))
    if a.scan:
        r = scan(lib, doc_text(Path(a.scan).expanduser()))
        if not a.json:
            s = r["suggest"]
            print("首几行：" + " / ".join(r["head"][:3]))
            print(f'像是：{"非合同文书" if s["mode"] == "doc" else "合同"}'
                  + (f'（{dict(DOCKIND).get(s["doc_kind"], "?")}）' if s["mode"] == "doc" and s["doc_kind"] else "")
                  + (f'　类型包 {s["type"]}（{short(lib, s["type"])}）' if s["type"] else "")
                  + (f'　情形包 {"、".join(s["situ"])}' if s["situ"] else ""))
            if r.get("multiparty", {}).get("multi"):
                mp = r["multiparty"]
                print("【我方多主体】看起来是多个主体并列 —— S1 必装（分别结算/分别开票/"
                      "不承担连带责任、签署栏分设、管辖落到具体主体…）：",
                      "、".join(f'{h["w"]}×{h["n"]}' for h in mp["evidence"][:4]))
            if r.get("strong"):
                print("强触发（默认就该装）：" + "、".join(
                    f'{k} {short(lib, k)}' for k in r["strong"]))
            if s["track_b_hint"]:
                print("有对方模板的味道（出现了单方有利表述），track 建议先按 B 看一眼")
            print("\n证据（词：出现次数）：")
            for grp in ("type", "situ"):
                for k, v in sorted(r[grp].items(), key=lambda kv: -kv[1]["score"])[:6]:
                    print(f'  {k} {short(lib, k):<24}{v["score"]:>3}  '
                          + "、".join(f'{h["w"]}×{h["n"]}' for h in v["hit"][:4]))
            print("\n这些只是线索，装不装你自己判断；拿不准就把包装上（规则里都有「例外」告诉你何时不适用）。")
            return
        return emit(r)

    if not a.scan_input and a.mode == "contract":
        print("[提示] 没给 --scan-input：强触发的情形包（多甲方、涉政府村集体、涉招投标…）"
              "不会自动装上，交叉引用清单也没提取。除非你已经手动确认过，否则加上它。",
              file=sys.stderr)
    text, packs, cs = build(lib, a)
    if a.out:
        Path(a.out).expanduser().write_text(text, encoding="utf-8")
    if a.json:
        emit({"packs": packs, "pack_labels": [short(lib, k) for k in packs],
              "cards": [c["id"] for c in cs], "n_cards": len(cs),
              "chars": len(text), "out": a.out or "", "text": None if a.out else text})
    elif not a.out:
        print(text)
    else:
        print(f"已装配 {len(cs)} 条规则（{'、'.join(short(lib, k) for k in packs)}）→ {a.out}")


if __name__ == "__main__":
    main()
