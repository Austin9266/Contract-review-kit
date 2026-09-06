#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify.py —— 交付前的强制体检。任何一项 FAIL 都不许交付。

用法:
    python3 scripts/verify.py --work work/ --out 交付.docx \\
        --author "某某律所 张三" [--windows "09:00-18:00"] [--weekend]

检查项:
  C0 XML 结构合法 —— 交付包内每个 XML 部件都能被解析器完整解析（正则手术若切出
     非法 XML，Word 直接报文件损坏；本项还核对 [Content_Types] 的 Override 与
     document.xml.rels 的内部 Target 都指向真实存在的部件，杜绝悬空引用）。
  C1 结构完整性 —— 除 document.xml / comments*.xml / rels / [Content_Types] 外，
     所有 zip 条目必须与来件逐字节一致（样式、编号、页眉页脚、图片、字体一个都不许动）。
  C2 正文零漂移 —— "拒绝全部修订后的正文" 必须与来件正文完全相同；
     不同则说明有改动没被 w:ins/w:del 包住（Word 里看不出来，最阴的错）。
  C3 署名一致 —— 本次新增的修订与批注均为指定署名，无 PENDING/Claude/Author 等杂署名
     （来件既有的第三方署名允许存在，见 C7）。
  C4 时间戳 —— 本次新增的时间与 Word 同形（本地钟面 + Z、秒为 00），钟面是北京时间未经折算，
     落在工作时间窗内、单调不减；写了 w16du:dateUtc 的，其值必须正好是钟面 −8h。
  C7 来件既有修订/批注零改写 —— 客户或对方律师原有的署名与时间必须一条不少、一字未改。
  C5 后缀一致 —— 交付文件后缀与来件一致。
  C6 批注可见 —— comments.xml 里每条批注在 document.xml 中都有 commentReference 锚点。
"""
import argparse
import difflib
import hashlib
import json
import re
import sys
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from wtext import document_text  # noqa: E402

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


MUTABLE = re.compile(r"^(word/document\.xml|word/comments.*\.xml|word/people\.xml|"
                     r"word/_rels/document\.xml\.rels|\[Content_Types\]\.xml|docProps/.*)$")
AUTHORED = re.compile(r"<w:[A-Za-z]+\b[^>]*?\sw:author=\"([^\"]*)\"[^>]*?/?>")
DATE = re.compile(r'\sw:date="([^"]*)"')
# Word 写修订时间的真实形态：本地钟面 + 一个 Z 后缀，秒恒为 00（如 2026-08-25T14:17:00Z）。
# 那个 Z 是 Word 的历史包袱，它并不换算时区 —— 所以"北京时间不折算"与"跟 Word 一致"是同一件事。
BARE = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:00Z$")
DATEUTC = re.compile(r'w16du:dateUtc="([^"]*)"')

ok = True


def check(cond, tag, msg):
    global ok
    print(("  [PASS] " if cond else "  [FAIL] ") + f"{tag} {msg}")
    if not cond:
        ok = False
    return cond


def in_windows(dt, windows, weekend):
    if not weekend and dt.weekday() >= 5:
        return False
    for (sh, sm), (eh, em) in windows:
        if (dt.hour, dt.minute) >= (sh, sm) and (dt.hour, dt.minute) < (eh, em):
            return True
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--work", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--author", default=_CFG_AUTHOR)
    ap.add_argument("--windows", default=_CFG_WINDOWS)
    ap.add_argument("--weekend", action="store_true")
    ap.add_argument("--expect-ext", help="最终交付给客户的后缀，如 .doc；默认取来件后缀")
    a = ap.parse_args()

    work = Path(a.work)
    base = json.loads((work / "baseline.json").read_text(encoding="utf-8"))
    windows = [tuple(tuple(int(v) for v in p.split(":")) for p in w.split("-"))
               for w in a.windows.split(",")]
    out = Path(a.out)
    expect_ext = a.expect_ext or base["source_ext"]

    print(f"== 体检（严格，针对 docx 结构）：{out.name} ==")
    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())

        # C0：逐部件 XML 结构解析（正则手术的兜底；C2 的正则实现抓不到结构错误）
        import xml.etree.ElementTree as ET
        broken = []
        for n in sorted(names):
            if not (n.endswith(".xml") or n.endswith(".rels")):
                continue
            try:
                ET.fromstring(z.read(n))
            except ET.ParseError as ex:
                broken.append(f"{n}（{ex}）")
        check(not broken, "C0", "全部 XML 部件可完整解析"
              + (f"，非法: {broken[:3]}" if broken else ""))
        # C0：无悬空引用（删/漏部件却留着 Override / Relationship，Word 报文件损坏）
        ct = z.read("[Content_Types].xml").decode("utf-8")
        dangling = [p for p in re.findall(r'<Override PartName="/([^"]+)"', ct)
                    if p not in names]
        rels = (z.read("word/_rels/document.xml.rels").decode("utf-8")
                if "word/_rels/document.xml.rels" in names else "")
        import posixpath
        for m in re.finditer(r'<Relationship\s[^>]*Target="([^"]+)"[^>]*/?>', rels):
            if 'TargetMode="External"' in m.group(0) or m.group(1).startswith("http"):
                continue
            tgt = (m.group(1).lstrip("/") if m.group(1).startswith("/")
                   else posixpath.normpath("word/" + m.group(1)))
            if tgt not in names:
                dangling.append(tgt)
        check(not dangling, "C0", "无悬空引用（Content_Types/rels 指向的部件都在）"
              + (f"，悬空: {dangling[:3]}" if dangling else ""))
        # C0：修订 id 唯一（ECMA-376 要求）。切分带 rPrChange 的 run 会复制出重复 id，
        # apply/wredline 写回前已去重；来件本身就重复的，只提示不拦截。
        import collections
        doc_raw = z.read("word/document.xml").decode("utf-8")
        rev_ids = re.findall(
            r'<w:(?:ins|del|rPrChange|pPrChange|moveFrom|moveTo|sectPrChange|tblPrChange|'
            r'trPrChange|tcPrChange)\b[^>]*?\sw:id="(\d+)"', doc_raw)
        dup_ids = sorted(k for k, v in collections.Counter(rev_ids).items() if v > 1)
        pre_dup = set(base.get("pre_dup_ids", []))
        new_dup = [d for d in dup_ids if d not in pre_dup]
        if dup_ids and not new_dup:
            print(f"  [INFO] C0 修订 id 有 {len(dup_ids)} 个重复，均为来件既有，未新增")
        else:
            check(not new_dup, "C0", f"修订 w:id 唯一（共 {len(rev_ids)} 处）"
                  + (f"，本次新产生重复: {new_dup[:5]}" if new_dup else ""))

        # C1（注意：先跳过可变部件，再判缺失 —— 可变部件由 C0/C6/C7 管辖）
        diffs, missing, missing_dirs = [], [], []
        for n, h in base["entries"].items():
            if MUTABLE.match(n):
                continue
            if n not in names:
                # 纯目录条目（WPS 生成的包带 'word/' 之类）不承载内容，
                # 丢了只提示不拦截 —— finish.rezip 正常会原样保留
                (missing_dirs if n.endswith("/") else missing).append(n)
                continue
            if hashlib.sha256(z.read(n)).hexdigest() != h:
                diffs.append(n)
        if missing_dirs:
            print(f"  [INFO] C1 目录条目缺失（不影响内容）: {', '.join(missing_dirs[:5])}")
        added = [n for n in names - set(base["entries"]) if not MUTABLE.match(n)]
        check(not diffs, "C1", f"非正文部件零改动 {('，被改动: ' + ', '.join(diffs)) if diffs else ''}")
        check(not missing, "C1", f"部件无缺失 {('，缺失: ' + ', '.join(missing)) if missing else ''}")
        check(not added, "C1", f"无意外新增部件 {('：' + ', '.join(added)) if added else ''}")

        doc = z.read("word/document.xml").decode("utf-8")
        comments = z.read("word/comments.xml").decode("utf-8") if "word/comments.xml" in names else ""
        cext = (z.read("word/commentsExtensible.xml").decode("utf-8")
                if "word/commentsExtensible.xml" in names else "")
        cids_xml = (z.read("word/commentsIds.xml").decode("utf-8")
                    if "word/commentsIds.xml" in names else "")

    # C2
    rejected = document_text(doc, mode="reject")
    baseline = (work / "baseline.txt").read_text(encoding="utf-8")
    same = rejected == baseline
    check(same, "C2", "拒绝全部修订后的正文 == 来件正文")
    if not same:
        d = list(difflib.unified_diff(baseline.split("\n"), rejected.split("\n"),
                                      "来件", "交付(拒绝修订后)", lineterm="", n=1))
        print("\n".join("      " + l for l in d[:60]))

    # C3 / C4：只针对"本次新增"的署名，来件既有的第三方署名走 C7
    pre_authors = set(base.get("pre_authors", []))
    all_authors = (set(AUTHORED.findall(doc))
                   | set(re.findall(r'<w:comment\s[^>]*w:author="([^"]*)"', comments)))
    mine = all_authors - pre_authors
    check(mine <= {a.author}, "C3",
          f"本次署名唯一：{a.author}"
          + (f"（另有杂署名 {sorted(mine - {a.author})}）" if mine - {a.author} else "")
          + (f"；来件既有作者 {sorted(pre_authors)} 保留" if pre_authors else ""))
    check("PENDING" not in all_authors, "C3", "无 PENDING 占位残留（忘跑 stamp.py 会留下）")

    pre_stamps = {(x[0], x[1]) for x in base.get("pre_stamps", [])}

    def own(pairs):
        return [d for au, d in pairs if au == a.author and (au, d) not in pre_stamps]

    pairs = (re.findall(r'w:author="([^"]*)"\s+w:date="([^"]*)"', doc + comments)
             + [(au, d) for d, au in
                re.findall(r'w:date="([^"]*)"\s+w:author="([^"]*)"', doc + comments)])
    dates = own(pairs)
    bad_fmt = [d for d in dates if not BARE.match(d)]
    check(not bad_fmt, "C4",
          "时间与 Word 同形（北京钟面+Z、秒为 00）"
          + (f"，异常: {bad_fmt[:5]}" if bad_fmt else ""))
    parsed = [datetime.fromisoformat(d[:-1]) for d in dates if BARE.match(d)]

    # dateUtc 必须是钟面 −8h：写反了就等于把北京时间当 UTC 存，正是要防的那个错
    pairs_utc = re.findall(r'w:date="([^"]*)"\s+w16du:dateUtc="([^"]*)"', doc + comments)
    bad_utc = [(l, u) for l, u in pairs_utc
               if BARE.match(l) and BARE.match(u)
               and datetime.fromisoformat(l[:-1]) - datetime.fromisoformat(u[:-1])
               != timedelta(hours=8)]
    check(not bad_utc, "C4",
          f"w16du:dateUtc = 钟面 −8h（共 {len(pairs_utc)} 处）"
          + (f"，异常: {bad_utc[:3]}" if bad_utc else ""))
    outw = [d for d in parsed if not in_windows(d, windows, a.weekend)]
    check(not outw, "C4", f"全部落在工作时间窗 {a.windows}"
          + (f"，越界: {[str(d) for d in outw[:5]]}" if outw else ""))
    seq = [datetime.fromisoformat(d[:-1]) for au, d in
           re.findall(r'w:author="([^"]*)"\s+w:date="([^"]*)"', doc)
           if au == a.author and BARE.match(d) and (au, d) not in pre_stamps]
    # 真人改文档并不按文档顺序从头到尾（对照修订模式模板实件：时间在文档顺序上来回跳），
    # 所以乱序不是错，只提示不拦截；stamp.py 默认单调是"一遍从头看到尾"的合理形态。
    if seq != sorted(seq):
        print("  [INFO] C4 时间在文档顺序上非单调（真人编辑本就如此，不拦截）")

    # C4：批注的真 UTC 存在 commentsExtensible.xml（对照实件校准；批注元素本身不带 dateUtc）。
    # 本次新增批注的条目必须存在且 = 钟面 −8h；该部件绝不允许被整体删除（S-2 教训：
    # 删部件不清 rels/Override 会悬空 —— 悬空由 C0 拦，这里管条目的对与全）。
    fresh_comments = [(cid2, d) for cid2, au2, d in re.findall(
        r'<w:comment\s[^>]*w:id="(\d+)"[^>]*w:author="([^"]*)"[^>]*w:date="([^"]*)"', comments)
        if au2 == a.author and (au2, d) not in pre_stamps]
    if fresh_comments and "xmlns:w16cex=" in doc[:6000]:
        para_of = {m.group(1): (re.search(r'w14:paraId="([0-9A-Fa-f]+)"', m.group(2)) or [None])
                   for m in re.finditer(r'<w:comment\s[^>]*w:id="(\d+)"[^>]*>(.*?)</w:comment>',
                                        comments, re.S)}
        dur_of = dict(re.findall(r'w16cid:paraId="([0-9A-Fa-f]+)"\s+w16cid:durableId="([0-9A-Fa-f]+)"',
                                 cids_xml))
        cx_utc = dict(re.findall(r'w16cex:durableId="([0-9A-Fa-f]+)"\s+w16cex:dateUtc="([^"]*)"',
                                 cext))
        bad_ce = []
        for cid2, d in fresh_comments:
            pm2 = para_of.get(cid2)
            pid2 = pm2.group(1) if hasattr(pm2, "group") else None
            utc = cx_utc.get(dur_of.get(pid2, ""), None)
            if utc is None:
                bad_ce.append(f"批注{cid2}: 缺 commentsExtensible 条目")
            elif BARE.match(d) and BARE.match(utc) and \
                    datetime.fromisoformat(d[:-1]) - datetime.fromisoformat(utc[:-1]) \
                    != timedelta(hours=8):
                bad_ce.append(f"批注{cid2}: dateUtc≠钟面−8h（{d} / {utc}）")
        check(not bad_ce, "C4", f"新增批注的 commentsExtensible.dateUtc 正确（共 {len(fresh_comments)} 条）"
              + (f"，异常: {bad_ce[:3]}" if bad_ce else ""))
    pre_cext = base.get("pre_cext", {})
    if pre_cext:
        cur = dict(re.findall(r'w16cex:durableId="([0-9A-Fa-f]+)"\s+w16cex:dateUtc="([^"]*)"',
                              cext))
        lost_ce = [k for k, v in pre_cext.items() if cur.get(k) != v]
        check(not lost_ce, "C7", f"来件既有 commentsExtensible 条目 {len(pre_cext)} 个原封未动"
              + (f"，被改/丢失: {lost_ce[:3]}" if lost_ce else ""))

    # C5
    check(expect_ext == base["source_ext"], "C5",
          f"计划交付后缀 {expect_ext} == 来件后缀 {base['source_ext']}")

    # C6
    cids = re.findall(r'<w:comment\s[^>]*w:id="(\d+)"', comments)
    refs = set(re.findall(r'<w:commentReference\s[^>]*w:id="(\d+)"', doc))
    orphan = [c for c in cids if c not in refs]
    check(not orphan, "C6", f"批注 {len(cids)} 条全部有锚点"
          + (f"，孤立: {orphan}" if orphan else ""))

    # C7
    if pre_stamps:
        present = set(pairs)
        lost = sorted(pre_stamps - present)
        check(not lost, "C7", f"来件既有 {len(pre_stamps)} 处第三方署名原封未动"
              + (f"，被改写/丢失: {lost[:3]}" if lost else ""))
    else:
        check(True, "C7", "来件本无第三方修订/批注")

    print("== 结论：" + ("全部通过，可交付 ==" if ok else "存在 FAIL，禁止交付 =="))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
