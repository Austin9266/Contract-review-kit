#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""typo.py —— 错别字、衍字、复制粘贴残留的机械扫描。

九查里「错别字近形字」一直写着要修订，但没有工具，全靠肉眼——肉眼在四十页合同上必漏。
这个脚本把能机械判定的那部分一次扫完，并且**直接产出可以并进落盘块的 @修订 条目**
（错别字属确定性错误，只有一个正确答案，按 SOP 就该修订而不是批注）。

    python3 typo.py --input 来件.docx --json -                 # 看清单
    python3 typo.py --input 来件.docx --emit-block 错别字块.txt  # 出落盘块（只出 sure 档）
    python3 typo.py --input 来件.docx --level all              # 连提示级一起看

扫六类：

  A 衍字        叠字叠词（「的的」「应当当」「提交提交」），叠词白名单外的都报   sure
  B 段内重复    同一段里重复出现的 ≥8 字长句（复制粘贴没删干净）              check
  C 易混词      rules/易混词.json 里的错别字与已废止法名                      sure/check
  D 大写金额    大写金额里混进阿拉伯数字、「元/圆」混用、「另」当「零」用        check
  E 括号        中英文括号混用、括号不配对                                    sure
  F 词中空格    汉字之间的多余空格（排版残留）                                 提示

判词表在 rules/易混词.json，想加词直接加，不用改代码。
"""
import argparse
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE / "core"))
from locator import Doc                                    # noqa: E402

TABLE = ROOT / "rules" / "易混词.json"
CJK = r"一-龥"


def load_doc(path: Path) -> Doc:
    if path.suffix.lower() == ".docx":
        with zipfile.ZipFile(path) as z:
            return Doc(z.read("word/document.xml").decode("utf-8"))
    if path.suffix.lower() == ".doc":
        work = Path(tempfile.mkdtemp(prefix="typo-"))
        r = subprocess.run([sys.executable, str(HERE / "prep.py"), str(path), "--work", str(work)],
                           capture_output=True, text=True)
        if r.returncode != 0:
            raise SystemExit("[FAIL] .doc 转换失败（要装 LibreOffice）：" + r.stderr[-200:])
        return Doc((work / "unpacked" / "word" / "document.xml").read_text(encoding="utf-8"))
    if path.suffix.lower() in (".txt", ".md"):
        return Doc("".join(f"<w:p><w:r><w:t>{l}</w:t></w:r></w:p>"
                           for l in path.read_text(encoding="utf-8").splitlines()))
    raise SystemExit("[FAIL] 只认 .doc/.docx/.txt")


def scan(doc: Doc, table: dict):
    dup_ok = set(table["dup_ok"])
    pairs = [p for p in table["pairs"] if p.get("level") in ("sure", "check")]
    hits = []

    def add(pi, s, e, kind, level, msg, right=None):
        hits.append({"para": pi, "start": s, "end": e, "kind": kind, "level": level,
                     "text": doc.paras[pi][s:e], "msg": msg, "right": right})

    for pi, t in enumerate(doc.paras):
        if not t.strip():
            continue
        # A 衍字：单字叠、双字叠
        for m in re.finditer(r"([%s])\1" % CJK, t):
            if m.group(0) in dup_ok:
                continue
            add(pi, m.start(), m.end(), "衍字", "sure",
                f"「{m.group(0)}」重复了一个字", m.group(1))
        for m in re.finditer(r"([%s]{2,4})\1" % CJK, t):
            if m.group(1) in dup_ok or len(set(m.group(1))) == 1:
                continue
            add(pi, m.start(), m.end(), "衍字", "sure",
                f"「{m.group(1)}」整词重复了一遍", m.group(1))
        # B 段内重复长片段（复制粘贴残留）
        seen = {}
        for w in (10,):
            for i in range(len(t) - w + 1):
                s = t[i:i + w]
                if not re.fullmatch(r"[%s，、；：（）()]+" % CJK, s):
                    continue
                if s in seen and i - seen[s] >= w:
                    add(pi, i, i + w, "段内重复", "check",
                        f"这一段里「{s}」出现了两次，检查是不是复制粘贴没删干净")
                    break
                seen.setdefault(s, i)
        # C 易混词
        for p in pairs:
            for m in re.finditer(re.escape(p["wrong"]), t):
                add(pi, m.start(), m.end(), "易混词", p["level"],
                    f'「{p["wrong"]}」→「{p["right"]}」（{p["why"]}）', p["right"])
        # D 大写金额
        for m in re.finditer(r"[壹贰叁肆伍陆柒捌玖拾佰仟萬亿零元圆整角分]{4,}", t):
            seg = m.group(0)
            around = t[max(0, m.start() - 2):m.end() + 2]
            if re.search(r"[0-9０-９]", around):
                add(pi, m.start(), m.end(), "大写金额", "check",
                    "大写金额附近混着阿拉伯数字，核对大小写是否一致")
            if "另" in t[max(0, m.start() - 1):m.end() + 1]:
                add(pi, m.start(), m.end(), "大写金额", "check", "「另」疑为「零」")
        # E 括号
        pairs_b = [("（", "）"), ("(", ")"), ("【", "】"), ("《", "》")]
        for a, b in pairs_b:
            if t.count(a) != t.count(b):
                i = t.find(a) if t.count(a) > t.count(b) else t.find(b)
                add(pi, max(0, i), max(0, i) + 1, "括号", "sure",
                    f"这一段里「{a}」{t.count(a)} 个、「{b}」{t.count(b)} 个，括号没配对")
        for m in re.finditer(r"（[^（）]*\)|\([^()]*）", t):
            add(pi, m.start(), m.end(), "括号", "sure", "中英文括号混用")
        # F 词中空格
        sp = list(re.finditer(r"(?<=[%s])[ 　]+(?=[%s])" % (CJK, CJK), t))
        if len(sp) >= 3:
            add(pi, sp[0].start(), sp[0].end(), "词中空格", "提示",
                f"这一段汉字之间有 {len(sp)} 处空格，多半是排版残留（一般不动，只在影响阅读时提）")
    return hits


def occurrence(doc: Doc, s: str, pi: int, x: int):
    """这串字在全文是第几处——落盘块靠它消歧。"""
    n = sum(doc.paras[j].count(s) for j in range(pi))
    t, k, at = doc.paras[pi], 0, doc.paras[pi].find(s)
    while 0 <= at < x:
        k += 1
        at = t.find(s, at + 1)
    return n + k + 1


def widen(doc: Doc, pi: int, s: int, e: int, pad=4, rounds=8):
    """把原文左右先各扩几个字（太短的串引擎不好落刀），再扩到全文唯一为止。"""
    t = doc.paras[pi]
    a, b = max(0, s - pad), min(len(t), e + pad)
    for _ in range(rounds):
        if sum(p.count(t[a:b]) for p in doc.paras) <= 1:
            break
        if a == 0 and b == len(t):
            break
        a, b = max(0, a - 3), min(len(t), b + 3)
    return a, b


def to_block(doc: Doc, hits):
    """sure 档 → @修订 条目；错别字属确定性错误，按 SOP 就该修订。"""
    out, n = ["<<<落盘>>>"], 0
    for h in hits:
        if h["level"] != "sure" or not h.get("right") or h["kind"] == "括号":
            continue
        pi, s, e = h["para"], h["start"], h["end"]
        a, b = widen(doc, pi, s, e)
        src = doc.paras[pi][a:b]
        dst = src[:s - a] + h["right"] + src[e - a:]
        if src == dst:
            continue
        n += 1
        occ = occurrence(doc, src, pi, a)
        out += [f"@修订 {n} | 第 {pi} 段（机械扫描）",
                f"原文：{src}", f"改为：{dst}"]
        if occ > 1 or sum(p.count(src) for p in doc.paras) > 1:
            out.append(f"第几处：{occ}")
        out.append(f"理由：{h['msg']}")
    out.append("<<<落盘完>>>")
    return "\n".join(out), n


def main():
    ap = argparse.ArgumentParser(description="错别字、衍字、复制粘贴残留的机械扫描")
    ap.add_argument("--input", required=True)
    ap.add_argument("--table", help="判词表，默认 rules/易混词.json")
    ap.add_argument("--level", choices=("sure", "check", "all"), default="check",
                    help="报到哪一档：sure＝只报确定的；check＝确定的＋要核实的（默认）；all＝连提示一起")
    ap.add_argument("--emit-block", help="把 sure 档写成落盘块，可直接并进你的落盘块")
    ap.add_argument("--json", help="完整结果；'-' 打到标准输出")
    a = ap.parse_args()

    doc = load_doc(Path(a.input).expanduser())
    table = json.loads(Path(a.table or TABLE).expanduser().read_text(encoding="utf-8"))
    hits = scan(doc, table)
    keep = {"sure": ["sure"], "check": ["sure", "check"],
            "all": ["sure", "check", "提示"]}[a.level]
    shown = [h for h in hits if h["level"] in keep]

    block, n = to_block(doc, hits)
    if a.emit_block:
        Path(a.emit_block).expanduser().write_text(block, encoding="utf-8")

    rep = {"input": a.input, "total": len(hits), "shown": len(shown),
           "by_kind": {k: sum(1 for h in hits if h["kind"] == k)
                       for k in sorted({h["kind"] for h in hits})},
           "sure": sum(1 for h in hits if h["level"] == "sure"),
           "block_entries": n, "block": None if a.emit_block else block,
           "hits": [{k: h[k] for k in ("para", "kind", "level", "text", "msg", "right")}
                    for h in shown]}
    if a.json:
        t = json.dumps(rep, ensure_ascii=False, indent=1)
        print(t) if a.json == "-" else Path(a.json).expanduser().write_text(t, encoding="utf-8")
        return
    if not shown:
        print("没扫到错别字/衍字（扫了 %d 段）。" % len(doc.paras))
        return
    print(f"扫到 {len(shown)} 处（{'、'.join(f'{k} {v}' for k, v in rep['by_kind'].items())}）：")
    for h in shown:
        t = doc.paras[h["para"]]
        lo, hi = max(0, h["start"] - 14), min(len(t), h["end"] + 14)
        print(f'  [{h["level"]}] 第 {h["para"]} 段 {h["kind"]}：{h["msg"]}')
        print(f'        …{t[lo:h["start"]]}【{h["text"]}】{t[h["end"]:hi]}…')
    if n:
        print(f"\n其中 {n} 处可直接落盘（确定性错误）——加 --emit-block 出落盘块，"
              "并进你自己的落盘块一起跑 oneshot。")


if __name__ == "__main__":
    main()
