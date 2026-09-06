#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自检.py —— 拿一份真来件，把落盘链路的几条「新建部件」路径都走一遍。

改过引擎、或者换了新的来件类型（WPS 件、老年代 .doc 转来的件）之后跑一次：

    python3 自检.py 某份真实来件.docx

它会用同一份来件造几个变体，分别跑完整落盘并核对 C0–C7：

    ① 原件                          —— 常规路径
    ② 去掉 commentsExtensible.xml   —— 走「新建批注扩展部件」路径（历史上在这里出过
                                       unbound prefix：抄 xmlns 时抄成了 <?xml?> 声明）
    ③ 去掉 people.xml               —— 走「新建作者部件」路径
    ④ 去掉两者                      —— 两条新建路径同时走

任何一项体检 FAIL，或新建出来的部件解析不了，就打 FAIL 并返回非零。
"""
import json
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ONESHOT = HERE / "oneshot.py"
sys.path.insert(0, str(HERE / "core"))
from locator import Doc                                    # noqa: E402
DROP = {
    "ce": ("word/commentsExtensible.xml", "commentsExtensible.xml"),
    "people": ("word/people.xml", "people.xml"),
    "comments": ("word/comments.xml", "comments.xml"),
    "cext": ("word/commentsExtended.xml", "commentsExtended.xml"),
    "cids": ("word/commentsIds.xml", "commentsIds.xml"),
}
# 把正文里的批注标记也摘掉，否则删了部件会留下悬空引用（那是另一种错，不是我们要测的）
COMMENT_MARKS = [
    re.compile(r"<w:commentRange(?:Start|End)\b[^>]*/>"),
    re.compile(r"<w:r\b(?:(?!</w:r>).)*?<w:commentReference\b[^>]*/>.*?</w:r>", re.S),
]


def variant(src: Path, dst: Path, drop_keys):
    """复制一份来件，把指定部件连同它的 rels 与内容类型 Override 一起摘干净。"""
    names = [k for k in drop_keys]
    with zipfile.ZipFile(src) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    gone = {DROP[k][0] for k in names}
    out = []
    for info, data in items:
        if info.filename in gone:
            continue
        if info.filename == "word/_rels/document.xml.rels":
            t = data.decode("utf-8")
            for k in names:
                t = re.sub(r'<Relationship[^>]*Target="%s"[^>]*/>' % DROP[k][1], "", t)
            data = t.encode("utf-8")
        if info.filename == "word/document.xml" and "comments" in names:
            t = data.decode("utf-8")
            for rx in COMMENT_MARKS:
                t = rx.sub("", t)
            data = t.encode("utf-8")
        if info.filename == "[Content_Types].xml":
            t = data.decode("utf-8")
            for k in names:
                t = re.sub(r'<Override[^>]*PartName="/%s"[^>]*/>' % DROP[k][0], "", t)
            data = t.encode("utf-8")
        out.append((info, data))
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)
    return dst


def pickable(src: Path):
    """挑出「全文唯一、且不在别人的修订块里」的短句——两条都得落得上，自检才算数。"""
    with zipfile.ZipFile(src) as z:
        doc = Doc(z.read("word/document.xml").decode("utf-8"))
    out = []
    for t in doc.paras:
        t = t.strip()
        for w in (14, 12, 10, 8, 6):
            if len(t) < w:
                continue
            s = t[:w]
            hits = doc.find_all(s)
            if len(hits) == 1 and doc.span_clean(*hits[0]):
                out.append(s)
                break
    return out


def same_para_pair(src: Path):
    """造一个「同一段里改两处，且两处收窄后是同一个字」的用例。

    这是历史上必挂的那条路径：收窄之后 find 只剩一个字，而这个字在本段更靠前处也出现过，
    旧代码从 x-80 开始找，命中的是前一处、且前一处已被上一条修订吃掉，第二条就报
    「这段文字在全文里重复太多」。→ (原文1, 改为1, 原文2, 改为2) 或 None。"""
    with zipfile.ZipFile(src) as z:
        doc = Doc(z.read("word/document.xml").decode("utf-8"))
    for t in doc.paras:
        if len(t) < 24:
            continue
        for ch in set(re.findall(r"[\u4e00-\u9fa5]", t)):
            at = [m.start() for m in re.finditer(re.escape(ch), t)]
            if len(at) < 2:
                continue
            for a, b in ((at[0], at[1]),):
                if a < 3 or b < a + 8 or b + 4 > len(t):
                    continue
                s1, s2 = t[a - 3:a + 4], t[b - 3:b + 4]
                if len(doc.find_all(s1)) != 1 or len(doc.find_all(s2)) != 1:
                    continue
                if not (doc.span_clean(*doc.find_all(s1)[0])
                        and doc.span_clean(*doc.find_all(s2)[0])):
                    continue
                return (s1, s1.replace(ch, "〇", 1) if s1.index(ch) == 3 else
                        s1[:3] + "〇" + s1[4:], s2, s2[:3] + "〇" + s2[4:])
    return None


def make_block(cands):
    """一条批注 ＋ 一条修订，把「新建批注部件」与「写修订」两条路径都走到。"""
    a = cands[0]
    b = next((s for s in cands[1:] if s != a), None)
    lines = ["<<<落盘>>>", "@批注 1 | 自检", f"锚点：{a}", "批注：自检用批注，请忽略。"]
    if b:
        lines += ["@修订 2 | 自检", f"原文：{b}", f"改为：{b}（自检）"]
    lines.append("<<<落盘完>>>")
    return "\n".join(lines)


def run_case(name, docx: Path, block: Path, work: Path, expect_notes=True):
    out = work / name
    out.mkdir(parents=True, exist_ok=True)
    r = subprocess.run([sys.executable, str(ONESHOT), "--input", str(docx),
                        "--block", str(block), "--auto", "--quiet", "--json", "-",
                        "--outdir", str(out), "--work", str(work / (name + "-w"))],
                       capture_output=True, text=True)
    try:
        rep = json.loads(r.stdout)
    except json.JSONDecodeError:
        return False, f"脚本没给出 JSON：{(r.stderr or r.stdout)[-300:]}"
    fails = [c for c in rep.get("checks", []) if "FAIL" in c]
    if not rep.get("ok"):
        return False, "落盘没跑通：" + (fails[0] if fails else rep["log"][-1] if rep.get("log") else "?")
    if fails:
        return False, "体检 FAIL：" + fails[0]
    deliver = Path(rep["deliver"])
    with zipfile.ZipFile(deliver) as z:
        for n in z.namelist():
            if not n.endswith(".xml") and not n.endswith(".rels"):
                continue
            try:
                ET.fromstring(z.read(n))
            except ET.ParseError as e:
                return False, f"{n} 解析不了：{e}"
        ce = "word/commentsExtensible.xml"
        if ce in z.namelist():
            root = z.read(ce).decode("utf-8")
            if "xmlns:w16cex=" not in root:
                return False, "commentsExtensible.xml 没有绑定 w16cex 命名空间"
    st = rep["stats"]
    tail = ""
    if expect_notes and not st["notes"]:
        return False, "批注没落上——「新建批注部件」这条路径没走到，自检不算数"
    if not st["edits"]:
        tail = "（这份来件里没有第二段可用来做修订，只验了批注路径）"
    if rep.get("actions"):
        tail += "（自动处理 %d 条）" % len(rep["actions"])
    return True, (f"修订 {st['edits']} 处、批注 {st['notes']} 条，"
                  f"体检 {len(rep['checks'])} 项全过" + tail)


def main():
    if len(sys.argv) < 2:
        raise SystemExit("用法：python3 自检.py <一份真实来件.docx>")
    src = Path(sys.argv[1]).expanduser().resolve()
    cands = pickable(src)
    if not cands:
        raise SystemExit("[FAIL] 这份来件里找不到「全文唯一且不在既有修订块里」的短句，"
                         "换一份来件自检。")
    work = Path(tempfile.mkdtemp(prefix="落盘自检-"))
    block = work / "block.txt"
    block.write_text(make_block(cands), encoding="utf-8")

    pair = same_para_pair(src)
    if pair:
        (work / "同段两条.txt").write_text(
            "<<<落盘>>>\n"
            f"@修订 1 | 同段第一处\n原文：{pair[0]}\n改为：{pair[1]}\n"
            f"@修订 2 | 同段第二处\n原文：{pair[2]}\n改为：{pair[3]}\n"
            "<<<落盘完>>>", encoding="utf-8")

    cases = [("原件", []),
             ("无commentsExtensible", ["ce"]),
             ("无people", ["people"]),
             ("无commentsExtensible+people", ["ce", "people"]),
             ("全无批注部件（最常见的干净来件）",
              ["comments", "cext", "cids", "ce", "people"])]
    ok_all = True
    for name, drop in cases:
        docx = src if not drop else variant(src, work / f"{name}.docx", drop)
        ok, msg = run_case(name, docx, block, work)
        ok_all &= ok
        print(f"{'[PASS]' if ok else '[FAIL]'} {name}：{msg}")
    if pair:
        ok, msg = run_case("同段两条修订", src, work / "同段两条.txt", work,
                           expect_notes=False)
        if ok and "修订 2 处" not in msg:
            ok, msg = False, "同一段里的两条修订没有全部落上（" + msg + "）"
        ok_all &= ok
        print(f"{'[PASS]' if ok else '[FAIL]'} 同段两条修订：{msg}")
    else:
        print("[跳过] 同段两条修订：这份来件里找不到同一段中的两处唯一短句")
    shutil.rmtree(work, ignore_errors=True)
    print("\n全部通过。" if ok_all else "\n有不通过的，别交付，先修脚本。")
    sys.exit(0 if ok_all else 1)


if __name__ == "__main__":
    main()
