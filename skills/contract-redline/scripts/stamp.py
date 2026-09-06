#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""stamp.py —— 统一署名 + 按"像人干活"的规律分配北京时间戳。

用法:
    python3 scripts/stamp.py work/unpacked \\
        --author "某某律所 张三" --initials 张 \\
        --start "2026-08-21T09:20" [--end "2026-08-22T17:00"] \\
        [--windows "09:00-18:00"] [--weekend] [--seed 7]

规则（照抄真实 Word 的写法，见 references/word-fingerprints.md）:
  * `w:date` 写**北京时间的钟面值再加一个 Z**，如 `2026-08-25T14:17:00Z`。
    这个 Z 在 Word 里是历史包袱：它写的其实是本地钟面时间，Z 只是个后缀，
    Word 显示时不做换算。所以"北京时间不折算 UTC"与"跟 Word 一模一样"在这里是同一件事。
  * 文档根节点声明了 w16du（Word 2023 以后）时，另写 `w16du:dateUtc` 放**真 UTC**
    （北京时间 −8h）。根节点没声明就不写 —— 给一份 Word 2010 的老文件塞 2023 的
    命名空间，反而是更扎眼的破绽。
  * **秒恒为 :00**，Word 的修订时间只精确到分钟。
  * **同一次编辑动作（同一段落）内的所有修订与批注共享同一个时间戳**，不做秒级微差。
  * 段落之间按随机间隔递进，只落在工作时间窗内，默认跳过周末；文档顺序上单调不减。
  * 删除的 run 补 `w:rsidDel`，新段补 `w14:paraId`，并确保 people.xml 里有作者条目。
"""
import argparse
import random
import re
import xml.etree.ElementTree as ET
import sys
from datetime import datetime, timedelta
from pathlib import Path
from xml.sax.saxutils import escape

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


PARA = re.compile(r"<w:p(?:\s[^>]*)?>|</w:p>")
# 只认本次落盘留下的占位署名 —— 来件里第三方（客户/对方律师/前手）的修订与批注
# 一个字都不许改：改了就等于把别人的修改冒认成我方的，且时间也会被改写。
PENDING_AUTHOR = "PENDING"
AUTHORED = re.compile(r"<w:[A-Za-z]+\b[^>]*?\sw:author=\"PENDING\"[^>]*?/?>")
CREF = re.compile(r'<w:commentReference\s[^>]*w:id="(\d+)"')
COMMENT = re.compile(r'(<w:comment\s[^>]*?)(/?>)', re.S)


def parse_windows(s):
    out = []
    for part in s.split(","):
        a, b = part.strip().split("-")
        out.append((tuple(int(x) for x in a.split(":")), tuple(int(x) for x in b.split(":"))))
    return out


class Clock:
    def __init__(self, windows, weekend=False):
        self.w = windows
        self.weekend = weekend

    def snap(self, dt):
        """把 dt 挪进最近的（当前或下一个）工作时间窗。"""
        for _ in range(400):
            if not self.weekend and dt.weekday() >= 5:
                dt = datetime.combine(dt.date() + timedelta(days=1), datetime.min.time())
                continue
            for (sh, sm), (eh, em) in self.w:
                s = dt.replace(hour=sh, minute=sm, second=0, microsecond=0)
                e = dt.replace(hour=eh, minute=em, second=0, microsecond=0)
                if dt < s:
                    return s
                if s <= dt < e:
                    return dt
            dt = datetime.combine(dt.date() + timedelta(days=1), datetime.min.time())
        raise SystemExit("[FAIL] 无法在工作时间窗内定位时间")

    def add(self, dt, minutes):
        dt = self.snap(dt)
        left = float(minutes)
        while left > 0:
            end = None
            for (sh, sm), (eh, em) in self.w:
                s = dt.replace(hour=sh, minute=sm, second=0, microsecond=0)
                e = dt.replace(hour=eh, minute=em, second=0, microsecond=0)
                if s <= dt < e:
                    end = e
                    break
            room = (end - dt).total_seconds() / 60.0
            if left < room:
                return dt + timedelta(minutes=left)
            left -= room
            dt = self.snap(end + timedelta(seconds=1))
        return dt

    def between(self, a, b):
        a, mins = self.snap(a), 0.0
        while a < b:
            end = None
            for (sh, sm), (eh, em) in self.w:
                s = a.replace(hour=sh, minute=sm, second=0, microsecond=0)
                e = a.replace(hour=eh, minute=em, second=0, microsecond=0)
                if s <= a < e:
                    end = e
                    break
            stop = min(end, b)
            mins += (stop - a).total_seconds() / 60.0
            if stop >= b:
                break
            a = self.snap(end + timedelta(seconds=1))
        return mins


def para_index_map(xml):
    """位置 → 段落序号。"""
    marks = [(m.start(), m.group(0)) for m in PARA.finditer(xml)]
    starts = [pos for pos, tag in marks if not tag.startswith("</")]
    return starts


def which_para(starts, pos):
    lo, hi = 0, len(starts)
    while lo < hi:
        mid = (lo + hi) // 2
        if starts[mid] <= pos:
            lo = mid + 1
        else:
            hi = mid
    return lo - 1


def root_tag(xml: str, name: str = "w:document") -> str:
    """取根元素的起始标签，如 <w:document xmlns:w="…" xmlns:w16cex="…" mc:Ignorable="…">。

    **不要写成 xml[:xml.index(">")+1]** —— 文件开头是 `<?xml version="1.0"?>` 声明，
    第一个 ">" 是声明的结尾，那样抓到的是声明、一个命名空间都没有。新建部件时
    照着它抄 xmlns 声明，抄空了就会产出 unbound prefix 的无效 XML（C0 会拦下）。"""
    m = re.search(r"<%s\b[^>]*>" % re.escape(name), xml)
    return m.group(0) if m else ""


def ns_decls(root: str) -> str:
    return " ".join(re.findall(r'xmlns:\w+="[^"]*"', root))


def has_w16du(xml: str) -> bool:
    root = root_tag(xml)
    if root:
        return "xmlns:w16du=" in root
    return 'xmlns:w16du=' in xml[:4000]


def stamp_pair(dt, offset_hours):
    """返回 (w:date, w16du:dateUtc)：前者是本地钟面 + Z，后者是真 UTC。秒恒为 00。"""
    dt = dt.replace(second=0, microsecond=0)
    local = dt.strftime("%Y-%m-%dT%H:%M:00Z")
    utc = (dt - timedelta(hours=offset_hours)).strftime("%Y-%m-%dT%H:%M:00Z")
    return local, utc


def rsid_root(unp: Path) -> str:
    """取文档自己的 rsidRoot，用于给删除的 run 补 w:rsidDel（Word 会带）。"""
    sp = unp / "word" / "settings.xml"
    if sp.exists():
        m = re.search(r'<w:rsidRoot w:val="([0-9A-Fa-f]+)"', sp.read_text(encoding="utf-8"))
        if m:
            return m.group(1)
        m = re.search(r'<w:rsid w:val="([0-9A-Fa-f]+)"', sp.read_text(encoding="utf-8"))
        if m:
            return m.group(1)
    return ""


def add_rsid_del(doc: str, rsid: str) -> str:
    """给本次删除块里的 w:r 补 w:rsidDel —— 真 Word 的删除 run 都带这个属性。"""
    if not rsid:
        return doc

    def fix(m):
        block = m.group(0)
        return re.sub(r"<w:r>", f'<w:r w:rsidDel="{rsid}">', block)
    return re.sub(r'<w:del\b[^>]*w:author="PENDING"[^>]*>.*?</w:del>', fix, doc, flags=re.S)


def ensure_people(unp: Path, author: str):
    """确保 word/people.xml 里有该作者条目 —— Word 每次保存修订都会写这个部件。"""
    from xml.sax.saxutils import escape as _esc
    au = _esc(author, {'"': "&quot;"})
    pp = unp / "word" / "people.xml"
    if pp.exists():
        x = pp.read_text(encoding="utf-8")
        if f'w15:author="{au}"' in x:
            return False
        entry = (f'<w15:person w15:author="{au}">'
                 f'<w15:presenceInfo w15:providerId="None" w15:userId="{au}"/></w15:person>')
        pp.write_text(x.replace("</w15:people>", entry + "</w15:people>"), encoding="utf-8")
        return True
    # 没有就照 Word 的样子新建一份，并挂上关系与内容类型
    docx = (unp / "word" / "document.xml").read_text(encoding="utf-8")
    decls = ns_decls(root_tag(docx)) or (
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"')
    if "xmlns:w15=" not in decls:
        decls += ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
    pp.write_text('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                  f'<w15:people {decls}>'
                  f'<w15:person w15:author="{au}">'
                  f'<w15:presenceInfo w15:providerId="None" w15:userId="{au}"/>'
                  f'</w15:person></w15:people>', encoding="utf-8")
    must_parse(pp)
    rels = unp / "word" / "_rels" / "document.xml.rels"
    rx = rels.read_text(encoding="utf-8")
    if 'Target="people.xml"' not in rx:
        nid = max([int(i) for i in re.findall(r'Id="rId(\d+)"', rx)] or [0]) + 1
        rx = rx.replace("</Relationships>",
                        f'<Relationship Id="rId{nid}" Type="http://schemas.microsoft.com/'
                        f'office/2011/relationships/people" Target="people.xml"/></Relationships>')
        rels.write_text(rx, encoding="utf-8")
    ct = unp / "[Content_Types].xml"
    cx = ct.read_text(encoding="utf-8")
    if 'PartName="/word/people.xml"' not in cx:
        ct.write_text(cx.replace("</Types>",
                      '<Override PartName="/word/people.xml" ContentType="application/vnd.'
                      'openxmlformats-officedocument.wordprocessingml.people+xml"/></Types>'),
                      encoding="utf-8")
    return True


def must_parse(path: Path):
    """新建的 XML 部件当场解析一遍：前缀没绑上、标签没闭合，这里就炸，
    错误信息直接指到文件，不必等 C0 体检倒推。"""
    try:
        ET.parse(path)
    except ET.ParseError as e:
        raise SystemExit(f"[FAIL] 新建的 {path.name} 不是合法 XML：{e}。"
                         "多半是命名空间声明没抄全 —— 见 root_tag() 的注释。")


def sync_comments_extensible(unp: Path, fresh_cids, stamped, default_pair):
    """真 Word 把批注的真 UTC 存在 word/commentsExtensible.xml：
    <w16cex:commentExtensible w16cex:durableId="…" w16cex:dateUtc="真UTC"/>，
    经 commentsIds.xml 的 paraId→durableId 关联到批注（对照修订模式模板实件校准；
    w:comment 元素本身不带 dateUtc 属性）。

    这里给本次新增的批注补条目，来件既有条目一字不动。**绝不删除该部件** ——
    早期版本删部件但不清理 word/_rels/document.xml.rels 的 Relationship 与
    [Content_Types].xml 的 Override，产出包成悬空引用，Word 报文件损坏。
    部件不存在时仅当文档根节点声明了 w16cex 才新建（保持与来件同年代）。"""
    cp = unp / "word" / "comments.xml"
    ci = unp / "word" / "commentsIds.xml"
    if not fresh_cids or not cp.exists() or not ci.exists():
        return False
    cx = cp.read_text(encoding="utf-8")
    ix = ci.read_text(encoding="utf-8")
    dur = dict(re.findall(
        r'<w16cid:commentId\s[^>]*w16cid:paraId="([0-9A-Fa-f]+)"\s+'
        r'w16cid:durableId="([0-9A-Fa-f]+)"', ix))
    entries = []          # (durableId, utc) 本次新增批注
    for m in re.finditer(r'<w:comment\s[^>]*w:id="(\d+)"[^>]*>(.*?)</w:comment>', cx, re.S):
        if m.group(1) not in fresh_cids:
            continue
        pm = re.search(r'w14:paraId="([0-9A-Fa-f]+)"', m.group(2))
        d = dur.get(pm.group(1)) if pm else None
        if d:
            entries.append((d, stamped.get(m.group(1), default_pair)[1]))
    if not entries:
        return False

    ce = unp / "word" / "commentsExtensible.xml"
    if not ce.exists():
        droot = root_tag((unp / "word" / "document.xml").read_text(encoding="utf-8"))
        if 'xmlns:w16cex=' not in droot:
            return False          # 老年代文档：真 Word 也未必写，保持同年代
        decls = ns_decls(droot)
        # mc:Ignorable 只在 mc 前缀确实声明了的时候才写，否则又是一个未绑定前缀
        ign = re.search(r'mc:Ignorable="([^"]*)"', droot) if 'xmlns:mc=' in decls else None
        ce.write_text('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                      f'<w16cex:commentsExtensible {decls}'
                      + (f' mc:Ignorable="{ign.group(1)}"' if ign else "")
                      + '></w16cex:commentsExtensible>', encoding="utf-8")
        must_parse(ce)            # 新建的部件当场验一次，别等到 C0 才发现
        rels = unp / "word" / "_rels" / "document.xml.rels"
        rx = rels.read_text(encoding="utf-8")
        if 'Target="commentsExtensible.xml"' not in rx:
            nid = max([int(i) for i in re.findall(r'Id="rId(\d+)"', rx)] or [0]) + 1
            rx = rx.replace("</Relationships>",
                            f'<Relationship Id="rId{nid}" Type="http://schemas.microsoft.com/'
                            f'office/2018/08/relationships/commentsExtensible" '
                            f'Target="commentsExtensible.xml"/></Relationships>')
            rels.write_text(rx, encoding="utf-8")
        ct = unp / "[Content_Types].xml"
        c = ct.read_text(encoding="utf-8")
        if 'PartName="/word/commentsExtensible.xml"' not in c:
            ct.write_text(c.replace("</Types>",
                          '<Override PartName="/word/commentsExtensible.xml" '
                          'ContentType="application/vnd.openxmlformats-officedocument.'
                          'wordprocessingml.commentsExtensible+xml"/></Types>'),
                          encoding="utf-8")

    x = ce.read_text(encoding="utf-8")
    close = "</w16cex:commentsExtensible>"
    for d, utc in entries:
        if f'w16cex:durableId="{d}"' in x:
            continue                       # 已有条目（不该发生在新批注上）——不动
        x = x.replace(close,
                      f'<w16cex:commentExtensible w16cex:durableId="{d}" '
                      f'w16cex:dateUtc="{utc}"/>' + close)
    ce.write_text(x, encoding="utf-8")
    return True


def touch_core_props(unp: Path, author: str, last_utc: str):
    """把"最后保存人/保存时间"改成我方 —— 不改的话文档属性里仍写着客户，
    等于这份文件从没被我方保存过，和满篇我方修订对不上。
    注意 docProps/core.xml 的 dcterms:modified 是**真 UTC**（与 w:date 的规矩相反）。"""
    cp = unp / "docProps" / "core.xml"
    if not cp.exists():
        return False
    from xml.sax.saxutils import escape as _esc
    x = cp.read_text(encoding="utf-8")
    au = _esc(author)
    if "<cp:lastModifiedBy>" in x:
        x = re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>",
                   f"<cp:lastModifiedBy>{au}</cp:lastModifiedBy>", x, flags=re.S)
    else:
        x = x.replace("</cp:coreProperties>",
                      f"<cp:lastModifiedBy>{au}</cp:lastModifiedBy></cp:coreProperties>")
    if "<dcterms:modified" in x:
        x = re.sub(r"(<dcterms:modified[^>]*>).*?(</dcterms:modified>)",
                   lambda m: m.group(1) + last_utc + m.group(2), x, flags=re.S)
    else:
        x = x.replace("</cp:coreProperties>",
                      '<dcterms:modified xsi:type="dcterms:W3CDTF">'
                      f"{last_utc}</dcterms:modified></cp:coreProperties>")
    m = re.search(r"<cp:revision>(\d+)</cp:revision>", x)
    if m:
        x = x.replace(m.group(0), f"<cp:revision>{int(m.group(1)) + 1}</cp:revision>")
    cp.write_text(x, encoding="utf-8")
    return True


def set_attr(tag, name, value):
    if re.search(rf'\s{name}="', tag):
        return re.sub(rf'(\s{name}=")[^"]*(")', lambda m: m.group(1) + value + m.group(2), tag, count=1)
    close = "/>" if tag.endswith("/>") else ">"
    return tag[: -len(close)] + f' {name}="{value}"' + close


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("unpacked")
    ap.add_argument("--author", default=_CFG_AUTHOR)
    ap.add_argument("--initials", default=_CFG_INITIALS)
    ap.add_argument("--start", required=True, help="北京时间，如 2026-08-21T09:20")
    ap.add_argument("--end", help="北京时间上限；给了就把时间戳压缩到该区间内")
    ap.add_argument("--windows", default=_CFG_WINDOWS)
    ap.add_argument("--weekend", action="store_true", help="允许落在周末")
    ap.add_argument("--seed", type=int, default=None)
    ap.add_argument("--utc-offset", type=float, default=8.0,
                    help="本地时区相对 UTC 的小时数，用于 w16du:dateUtc（北京为 8）")
    a = ap.parse_args()

    rnd = random.Random(a.seed)
    clock = Clock(parse_windows(a.windows), a.weekend)
    start = clock.snap(datetime.fromisoformat(a.start))
    unp = Path(a.unpacked)
    docp = unp / "word" / "document.xml"
    doc = docp.read_text(encoding="utf-8")
    starts = para_index_map(doc)

    # 1) 收集修订元素与批注引用，按文档顺序聚成"段落簇"
    events = []           # (pos, kind, match)
    for m in AUTHORED.finditer(doc):
        events.append((m.start(), "rev", m))
    cp = unp / "word" / "comments.xml"
    cx0 = cp.read_text(encoding="utf-8") if cp.exists() else ""
    PLACEHOLDER_DATE = "1999-01-01T00:00:00"
    fresh_cids = {m.group(1) for m in re.finditer(
        r'<w:comment\s[^>]*w:id="(\d+)"[^>]*w:date="' + PLACEHOLDER_DATE + '"', cx0)}
    for m in CREF.finditer(doc):
        if m.group(1) in fresh_cids:
            events.append((m.start(), "cmt", m))
    events.sort(key=lambda e: e[0])
    if not events:
        print("[WARN] 没有待署名的新修订/新批注（无 PENDING 占位），无需 stamp")
        return

    clusters = []
    for pos, kind, m in events:
        pi = which_para(starts, pos)
        if clusters and clusters[-1][0] == pi:
            clusters[-1][1].append((kind, m))
        else:
            clusters.append((pi, [(kind, m)]))

    # 2) 生成每簇的起始时刻
    n = len(clusters)
    if a.end:
        end = clock.snap(datetime.fromisoformat(a.end))
        total = clock.between(start, end)
        if total <= 0:
            raise SystemExit("[FAIL] start/end 之间没有工作时间")
        step = total / max(n, 1)
        offs, acc = [], 0.0
        for _ in range(n):
            offs.append(acc)
            acc += step * rnd.uniform(0.6, 1.4)
        scale = (total * 0.92) / acc if acc > total * 0.92 else 1.0
        times = [clock.add(start, o * scale) for o in offs]
    else:
        times, cur = [], start
        for i in range(n):
            if i:
                cur = clock.add(cur, rnd.uniform(2, 20))
            times.append(cur)

    # 3) 回写（从后往前替换，避免位置漂移）
    doc = add_rsid_del(doc, rsid_root(unp))
    starts = para_index_map(doc)
    events = []
    for m in AUTHORED.finditer(doc):
        events.append((m.start(), "rev", m))
    for m in CREF.finditer(doc):
        if m.group(1) in fresh_cids:
            events.append((m.start(), "cmt", m))
    events.sort(key=lambda e: e[0])
    clusters = []
    for pos, kind, m in events:
        pi = which_para(starts, pos)
        if clusters and clusters[-1][0] == pi:
            clusters[-1][1].append((kind, m))
        else:
            clusters.append((pi, [(kind, m)]))

    utc_ok = has_w16du(doc)
    stamped_comments = {}
    author = escape(a.author, {'"': "&quot;"})
    initials = escape(a.initials, {'"': "&quot;"})
    edits = []
    for (pi, items), t0 in zip(clusters, times):
        # 同一次编辑动作内的所有修订与批注共享同一个时间戳（Word 只精确到分钟）
        ts, ts_utc = stamp_pair(t0, a.utc_offset)
        for kind, m in items:
            if kind == "rev":
                tag = set_attr(set_attr(m.group(0), "w:author", author), "w:date", ts)
                if utc_ok:
                    tag = set_attr(tag, "w16du:dateUtc", ts_utc)
                edits.append((m.start(), m.end(), tag))
            else:
                stamped_comments[m.group(1)] = (ts, ts_utc)
    for s, e, txt in sorted(edits, reverse=True):
        doc = doc[:s] + txt + doc[e:]
    docp.write_text(doc, encoding="utf-8")

    # 4) comments.xml 同步署名与时间
    if cp.exists():
        cx = cx0

        def fix(m):
            head, close = m.group(1), m.group(2)
            cid = re.search(r'w:id="(\d+)"', head)
            if (cid.group(1) if cid else "") not in fresh_cids:
                return head + close          # 来件既有批注：原样保留
            ts, ts_utc = stamped_comments.get(
                cid.group(1) if cid else "", stamp_pair(start, a.utc_offset))
            # 真 Word 的 w:comment 不带 w16du:dateUtc 属性 —— 批注的真 UTC
            # 存在 commentsExtensible.xml 里（见 sync_comments_extensible）。
            head = set_attr(head + close, "w:author", author)
            head = set_attr(head, "w:initials", initials)
            head = set_attr(head, "w:date", ts)
            return head
        cx = COMMENT.sub(fix, cx)
        cp.write_text(cx, encoding="utf-8")

    if sync_comments_extensible(unp, fresh_cids, stamped_comments,
                                stamp_pair(start, a.utc_offset)):
        print("[OK] commentsExtensible.xml 已为本次新增批注补 dateUtc（真 UTC；既有条目未动）")

    kept = len(re.findall(r'\sw:author="', doc)) - len(edits)
    print(f"[OK] 署名 {a.author} ｜ 本次修订/批注 {len(events)} 处，编辑动作 {n} 次")
    if kept > 0:
        print(f"[OK] 来件既有署名 {kept} 处，原封未动")
    _, last_utc = stamp_pair(times[-1], a.utc_offset)
    if touch_core_props(unp, a.author, last_utc):
        print(f"[OK] 文档属性：最后保存人 → {a.author}，保存时间 → {last_utc}（此处是真 UTC）")
    if ensure_people(unp, a.author):
        print("[OK] people.xml 已补作者条目（Word 保存修订时会写这个部件）")
    print(f"[OK] 时间跨度 {times[0]:%Y-%m-%d %H:%M} → {times[-1]:%Y-%m-%d %H:%M}"
          f"（北京钟面写入 w:date，{'另写 w16du:dateUtc 真 UTC' if utc_ok else '文档未声明 w16du，不写 dateUtc'}）")


if __name__ == "__main__":
    main()
