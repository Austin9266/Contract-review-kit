#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""locator.py —— 把 AI 写的原文/锚点，落到来件里的**精确串 + 全局第几处**。

为什么需要这一层：AI 看到的是聊天窗口里粘贴的纯文本，Word 里的真实原文常常
差那么一点点——全角半角、数字与汉字之间的空格、弯引号直引号、破折号、
AI 顺手补的标点。旧流程里这种差异会让 apply.py 一条条 FAIL，然后靠 AI 反复
--grep 试错，慢就慢在这里。

这里做三级定位：
    1) 精确：原文原样出现在某段里 → 直接用；
    2) 归一：忽略空格/全半角/引号/破折号后命中 → 取回**文档里的原样切片**；
    3) 相似：都不中，按相似度给候选段落，交给人点一下。

定位到之后再做一件事——**收窄修订范围**：把原文与改为的公共前后缀剥掉，
只把真正变了的那几个字包进修订。这样 Word 里看到的是"15 日→30 日"，
而不是整句划线重写；同时也顺带绕开了"AI 抄整句时抄漏一个空格"的问题。
"""
import difflib
import re
import sys
from xml.sax.saxutils import unescape
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "引擎"))

from wpara import tokenize, index_map          # noqa: E402
from wredline import PARA, split_para          # noqa: E402

# ── 归一化：严格 1→1 或 1→0（丢弃），保证能映射回原始下标 ────────────────
DROP = set(" \t\r\n　 ​‌‍﻿")
MAP = {}
for i in range(0x21, 0x7F):                      # 全角 ASCII → 半角
    MAP[chr(0xFF00 + i - 0x20)] = chr(i)
MAP.update({
    "“": '"', "”": '"', "„": '"', "‟": '"', "「": '"', "」": '"', "『": '"', "』": '"',
    "‘": "'", "’": "'", "‛": "'",
    "—": "-", "–": "-", "‒": "-", "―": "-", "－": "-", "─": "-", "〜": "~", "～": "~",
    "．": ".", "。": ".", "，": ",", "；": ";", "：": ":", "！": "!", "？": "?",
    "（": "(", "）": ")", "【": "[", "】": "]", "〔": "[", "〕": "]",
})


def norm(s: str):
    """→ (归一串, 下标表)；下标表[i] 是归一串第 i 个字符在原串里的位置。"""
    out, idx = [], []
    for i, ch in enumerate(s):
        if ch in DROP:
            continue
        c = MAP.get(ch, ch)
        out.append(c.lower() if c.isascii() else c)
        idx.append(i)
    return "".join(out), idx


def raw_slice(idx, s: str, a: int, b: int):
    """归一串区间 [a,b) → 原串区间 [x,y)。"""
    if a >= len(idx):
        return len(s), len(s)
    x = idx[a]
    y = (idx[b - 1] + 1) if b - 1 < len(idx) and b > a else x
    return x, y


CLAUSE = re.compile(r"(?:第\s*)?([0-9０-９]+(?:\s*[.．、]\s*[0-9０-９]+)*)\s*(?:条|款|项)?")
DIG = str.maketrans("０１２３４５６７８９", "0123456789")


def clause_key(s: str):
    """从「第7.2条」「7.2」「第九条」里取出可比对的条款号；取不到返回 ''。"""
    if not s:
        return ""
    m = CLAUSE.search(s.translate(DIG))
    return re.sub(r"[\s．、]", ".", m.group(1)) if m else ""


TXT = re.compile(r"<w:(t|delText)(?:\s[^>]*)?>(.*?)</w:(?:t|delText)>", re.S)


def rev_view(body: str, drop: str) -> str:
    """drop='del' → 接受全部修订后的正文；drop='ins' → 拒绝全部修订后的正文。"""
    b = re.sub(r"<w:%s\b[^>]*>.*?</w:%s>" % (drop, drop), "", body, flags=re.S)
    return "".join(unescape(m.group(2)) for m in TXT.finditer(b))


class Doc:
    """来件正文的只读视图（与引擎看到的完全一致：同一套 tokenize/index_map）。"""

    def __init__(self, xml: str):
        self.paras, self.toks, self.norms, self.withrev = [], [], [], []
        for pm in PARA.finditer(xml):
            _, _, body = split_para(pm.group(0))
            tk = tokenize(body)
            text, _ = index_map(tk)
            self.paras.append(text)
            self.toks.append(tk)
            self.norms.append(norm(text))
            # 另存"接受全部修订"与"拒绝全部修订"两个视角的正文：
            # 用来把"全文压根没有"和"这段字在别人的修订块里"分开说
            self.withrev.append((norm(rev_view(body, "del"))[0],
                                 norm(rev_view(body, "ins"))[0]))

    def in_revision(self, needle: str):
        """这段文字是不是躺在既有修订块里？→ 段号，否则 -1。"""
        nn, _ = norm(needle)
        if not nn:
            return -1
        for i, (plain, (acc, rej)) in enumerate(zip(self.norms, self.withrev)):
            if nn not in plain[0] and (nn in acc or nn in rej):
                return i
        return -1

    # ── 一段之内的精确/归一命中 ──────────────────────────────────────────
    def _hits_in(self, pi: int, needle_n: str):
        np_, idx = self.norms[pi]
        hits, at = [], np_.find(needle_n)
        while at >= 0:
            x, y = raw_slice(idx, self.paras[pi], at, at + len(needle_n))
            hits.append((pi, x, y))
            at = np_.find(needle_n, at + 1)
        return hits

    def find_all(self, needle: str):
        nn, _ = norm(needle)
        if not nn:
            return []
        out = []
        for pi in range(len(self.paras)):
            out.extend(self._hits_in(pi, nn))
        return out

    # ── 相似度候选 ───────────────────────────────────────────────────────
    def candidates(self, needle: str, top=6):
        nn, _ = norm(needle)
        if not nn:
            return []
        scored = []
        for pi, (np_, idx) in enumerate(self.norms):
            if not np_:
                continue
            sm = difflib.SequenceMatcher(None, np_, nn, autojunk=False)
            if sm.real_quick_ratio() < 0.25:
                continue
            blocks = [b for b in sm.get_matching_blocks() if b.size >= 2]
            if not blocks or max(b.size for b in blocks) < max(4, len(nn) // 6):
                continue
            # 用首尾匹配块反推真实边界：AI 抄漏字时，定长窗口会把结尾切掉，
            # 收窄修订范围时就会误判成"要补一段字"，这是最阴的一种错。
            a0 = max(0, blocks[0].a - blocks[0].b)
            last = blocks[-1]
            a1 = min(len(np_), last.a + last.size + (len(nn) - last.b - last.size))
            if a1 <= a0:
                continue
            win = np_[a0:a1]
            r = difflib.SequenceMatcher(None, win, nn, autojunk=False).ratio()
            x, y = raw_slice(idx, self.paras[pi], a0, a1)
            scored.append({"para": pi, "score": round(r, 3), "start": x, "end": y,
                           "text": self.paras[pi][:160],
                           "hit": self.paras[pi][x:y]})
        scored.sort(key=lambda d: -d["score"])
        return scored[:top]

    # ── 落盘过程的"文本已被吃掉"模拟 ────────────────────────────────────
    # 引擎是顺着清单一条条改的：一条替换/删除做完，那段原文就进了 <w:del>，
    # 后面的条目再检索时**看不见它了**——"第 2 处"会变成"第 1 处"。
    # 所以算「第几处」必须按落盘当时的状态算，这里用一组"已吃掉的区间"模拟。
    def reset_sim(self):
        self.cut = [[] for _ in self.paras]
        self._simcache = {}

    def sim(self, pi):
        c = getattr(self, "_simcache", None)
        if c is None:
            self.reset_sim()
            c = self._simcache
        if pi in c:
            return c[pi]
        text, cuts = self.paras[pi], sorted(self.cut[pi])
        if not cuts:
            v = (text, list(range(len(text))))
        else:
            keep, idx, pos = [], [], 0
            for a, b in cuts:
                if a > pos:
                    keep.append(text[pos:a])
                    idx.extend(range(pos, a))
                pos = max(pos, b)
            keep.append(text[pos:])
            idx.extend(range(pos, len(text)))
            v = ("".join(keep), idx)
        c[pi] = v
        return v

    def sim_pos(self, pi, x):
        """原始坐标 → 落盘当时的坐标；那一段已被吃掉则返回 None。"""
        _, idx = self.sim(pi)
        lo, hi = 0, len(idx)
        while lo < hi:
            mid = (lo + hi) // 2
            if idx[mid] < x:
                lo = mid + 1
            else:
                hi = mid
        return lo if lo < len(idx) and idx[lo] == x else (lo if x >= len(self.paras[pi]) else None)

    def eat(self, pi, a, b):
        self.cut[pi].append((a, b))
        self._simcache.pop(pi, None)

    def eaten(self, pi, a, b):
        return any(not (b <= ca or a >= cb) for ca, cb in self.cut[pi])

    def best_in(self, pi: int, needle: str):
        """在指定段落里找最像 needle 的一段（人工指认了段落、但没框出具体位置时用）。"""
        nn, _ = norm(needle)
        np_, idx = self.norms[pi]
        if not nn or not np_:
            return None
        sm = difflib.SequenceMatcher(None, np_, nn, autojunk=False)
        blocks = [b for b in sm.get_matching_blocks() if b.size >= 2]
        if not blocks:
            return None
        a0 = max(0, blocks[0].a - blocks[0].b)
        last = blocks[-1]
        a1 = min(len(np_), last.a + last.size + (len(nn) - last.b - last.size))
        if a1 <= a0:
            return None
        r = difflib.SequenceMatcher(None, np_[a0:a1], nn, autojunk=False).ratio()
        x, y = raw_slice(idx, self.paras[pi], a0, a1)
        return x, y, r

    def span_clean(self, pi, x, y):
        """区间里是否混有图片/域代码/既有修订块——混了就不能做修订手术。"""
        text = self.paras[pi]
        if x >= y or y > len(text):
            return False
        tk = self.toks[pi]
        _, idx = index_map(tk)
        i0 = idx[x][0]
        i1 = idx[y - 1][0]
        return all(tk[i].kind == "run" for i in range(i0, i1 + 1))

    def unique_enough(self, s: str):
        return sum(self.sim(i)[0].count(s) for i in range(len(self.paras)))


# ── 定位一条 ──────────────────────────────────────────────────────────────
def locate(doc: Doc, needle: str, loc_hint: str = "", prefer_after: int = -1, occ_hint: int = 0):
    """→ dict(status, para, start, end, note, cands)
    status: exact | norm | fuzzy | ambiguous | miss"""
    if not needle:
        return {"status": "miss", "note": "没有可用于定位的文字", "cands": []}

    hits = [h for h in doc.find_all(needle)]
    if hits:
        exact = [h for h in hits if doc.paras[h[0]][h[1]:h[2]] == needle]
        pool = exact or hits
        kind = "exact" if exact else "norm"
        if occ_hint and occ_hint <= len(pool):
            pi, x, y = pool[occ_hint - 1]
            return {"status": kind, "para": pi, "start": x, "end": y,
                    "note": f"按「第几处：{occ_hint}」取第 {occ_hint} 处", "cands": []}
        if len(pool) == 1:
            pi, x, y = pool[0]
            return {"status": kind, "para": pi, "start": x, "end": y, "note": "", "cands": []}
        # 多处命中：先按条款号消歧，再按"紧跟上一条之后"消歧
        key = clause_key(loc_hint)
        ranked = []
        for k, (pi, x, y) in enumerate(pool):
            s = 0.0
            if key:
                for j in range(pi, max(-1, pi - 6), -1):
                    kj = clause_key(doc.paras[j][:24])
                    if kj:                       # 最近的一个带条款号的段落说了算
                        if kj == key:
                            s += 2.0
                        break
            if prefer_after >= 0 and pi >= prefer_after:
                s += 1.0 - min(0.9, (pi - prefer_after) / 200.0)
            ranked.append((s, k, pi, x, y))
        ranked.sort(key=lambda t: (-t[0], t[1]))
        best, second = ranked[0], ranked[1]
        cands = [{"para": pi, "score": round(s, 2), "start": x, "end": y,
                  "text": doc.paras[pi][:160], "hit": doc.paras[pi][x:y]}
                 for s, _, pi, x, y in ranked]
        if best[0] - second[0] >= 1.0:
            return {"status": kind, "para": best[2], "start": best[3], "end": best[4],
                    "note": f"全文 {len(pool)} 处，按条款位置取第 {best[1] + 1} 处", "cands": cands}
        return {"status": "ambiguous", "note": f"全文有 {len(pool)} 处一样的文字，需要指认",
                "cands": cands}

    cands = doc.candidates(needle)
    if cands and cands[0]["score"] >= 0.92 and (
            len(cands) == 1 or cands[0]["score"] - cands[1]["score"] >= 0.04):
        c = cands[0]
        return {"status": "fuzzy", "para": c["para"], "start": c["start"], "end": c["end"],
                "note": f"原文对不上，按相似度 {c['score']:.0%} 定到此处（请核对）", "cands": cands}
    inrev = doc.in_revision(needle)
    if inrev >= 0:
        return {"status": "miss", "cands": cands,
                "note": f"这段文字在第 {inrev} 段，但它整个落在来件既有的修订块里"
                        "（客户、对方律师或上一版留下的）。不在别人的修订上叠改——"
                        "先跟经办人确认那处修订是否接受，再决定怎么写。"}
    return {"status": "miss", "note": "全文找不到这段原文" + ("，下面是最像的几段" if cands else ""),
            "cands": cands}


# ── 收窄修订范围 ──────────────────────────────────────────────────────────
MIN_ANCHOR = 1


def narrow(doc: Doc, pi: int, x: int, y: int, dst: str):
    """把 (段 pi 的 [x,y) 原文) → dst 收窄成最小改动区间。
    → (find, replace, mode)  mode ∈ replace | insert_after | delete"""
    src = doc.paras[pi][x:y]
    ns, isx = norm(src)
    nd, idd = norm(dst)
    if not ns and not nd:
        return src, dst, "replace"
    p = 0
    while p < len(ns) and p < len(nd) and ns[p] == nd[p]:
        p += 1
    q = 0
    while q < len(ns) - p and q < len(nd) - p and ns[len(ns) - 1 - q] == nd[len(nd) - 1 - q]:
        q += 1
    if p == 0 and q == 0:
        return src, dst, ("delete" if not dst else "replace")

    sx = isx[p] if p < len(ns) else len(src)
    sy = (isx[len(ns) - q] if q and len(ns) - q < len(ns) else len(src)) if q else len(src)
    dx = idd[p] if p < len(nd) else len(dst)
    dy = (idd[len(nd) - q] if q and len(nd) - q < len(nd) else len(dst)) if q else len(dst)
    if sx > sy or dx > dy:
        return src, dst, ("delete" if not dst else "replace")

    core_s, core_d = src[sx:sy], dst[dx:dy]

    if not core_s:                                   # 纯插入：锚在前半段
        a0 = x + sx
        lo = max(x, a0 - max(MIN_ANCHOR, 10))
        anchor = doc.paras[pi][lo:a0]
        while len(anchor) < MIN_ANCHOR and lo > 0:
            lo -= 1
            anchor = doc.paras[pi][lo:a0]
        if not anchor:
            return src, dst, "replace"
        return anchor, core_d, "insert_after"

    # 收窄后太短容易撞车：往两边补回一点上下文
    while len(core_s) < MIN_ANCHOR and (sx > 0 or sy < len(src)):
        if sx > 0:
            sx -= 1
            dx = max(0, dx - 1)
        elif sy < len(src):
            sy += 1
            dy = min(len(dst), dy + 1)
        core_s, core_d = src[sx:sy], dst[dx:dy]
    return core_s, core_d, ("delete" if not core_d else "replace")


def engine_locate(doc: "Doc", needle: str, occurrence: int):
    """复刻 wredline.locate 的定位语义（按落盘当时的文本状态），用来回验「第几处」。"""
    seen = 0
    for pi in range(len(doc.paras)):
        text = doc.sim(pi)[0]
        n = text.count(needle)
        if n and seen + n >= occurrence:
            start = -1
            for _ in range(occurrence - seen):
                start = text.find(needle, start + 1)
                if start < 0:
                    return None
            return pi, start
        seen += n
    return None


# ── 让插入的文字随来件的排版习惯 ──────────────────────────────────────────
CJK = r"\u4e00-\u9fff\u3000-\u303f\uff00-\uffef"
_D_SP = re.compile(f"[0-9]\\s+[{CJK}]")
_D_NO = re.compile(f"[0-9][{CJK}]")
_C_SP = re.compile(f"[{CJK}]\\s+[0-9]")
_C_NO = re.compile(f"[{CJK}][0-9]")


def spacing_style(paras):
    """来件把数字和汉字之间写不写空格？→ 'space' / 'tight' / ''（看不出来）。"""
    sp = no = 0
    for t in paras:
        t = re.sub(r"^\s*[0-9.．、]+\s*", "", t)   # 条款号后面那个空格不算排版习惯
        sp += len(_D_SP.findall(t)) + len(_C_SP.findall(t))
        no += len(_D_NO.findall(t)) + len(_C_NO.findall(t))
    if sp + no < 4:
        return ""
    if no >= sp * 3:
        return "tight"
    if sp >= no * 3:
        return "space"
    return ""


def fit_spacing(style: str, text: str) -> str:
    """按来件习惯调整插入文字里"数字↔汉字"之间的空格；其余一个字不动。"""
    if not text or not style:
        return text
    if style == "tight":
        text = re.sub(f"(?<=[0-9])[ \u3000]+(?=[{CJK}])", "", text)
        text = re.sub(f"(?<=[{CJK}])[ \u3000]+(?=[0-9])", "", text)
    else:
        text = re.sub(f"(?<=[0-9])(?=[{CJK}])", " ", text)
        text = re.sub(f"(?<=[{CJK}])(?=[0-9])", " ", text)
    return text


def locate_clause(doc: "Doc", loc_hint: str):
    """只给了「第5.2条」这种位置、没给锚点时，按条款号找段落。"""
    key = clause_key(loc_hint)
    if not key:
        return None
    hits = [i for i, t in enumerate(doc.paras) if clause_key(t[:24]) == key and t.strip()]
    if len(hits) == 1:
        return hits[0]
    return hits[0] if hits else None
