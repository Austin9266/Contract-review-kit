#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""wpara.py —— 段落级的 run 模型：支持跨 run 定位原文。

Word（以及 LibreOffice 转出来的 docx）会因为字体、拼写标记、rsid 把一句话拆成
好几个 <w:r>，所以"整串原文在单个 w:t 里"经常不成立。本模块把段落切成 token，
只在"可编辑的正文 run"上做匹配与切分，其余（已有的 w:ins/w:del 块、超链接、
书签、批注标记）整块保留、原样不动。
"""
import re
from xml.sax.saxutils import escape, unescape

BLOCK = re.compile(
    r"<w:ins\b[^>]*?/>|<w:del\b[^>]*?/>|"
    r"<w:ins\b[^>]*?>.*?</w:ins>|<w:del\b[^>]*?>.*?</w:del>|"
    r"<w:moveFrom\b[^>]*?>.*?</w:moveFrom>|<w:moveTo\b[^>]*?>.*?</w:moveTo>|"
    r"<w:hyperlink\b[^>]*?>.*?</w:hyperlink>|"
    r"<w:r\b[^>]*?>.*?</w:r>|<w:r\b[^>]*?/>", re.S)
T = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)

# --- pPr 的配对提取与段落标记工具 -------------------------------------------
# 注意：pPr 里可能嵌套 <w:pPrChange>（其中又有一层 <w:pPr>），rPr 里可能嵌套
# <w:rPrChange>（其中又有一层 <w:rPr>）。用非贪婪正则 <w:pPr>.*?</w:pPr> 会在
# 内层 </w:pPr> 处截断，切分后产出非法 XML —— 必须按开闭配对计数。
PPR_TAG = re.compile(r"<w:pPr(?:\s[^>]*)?/?>|</w:pPr>")
RPR_TAG = re.compile(r"<w:rPr(?:\s[^>]*)?/?>|</w:rPr>")
PPRCHANGE = re.compile(r"<w:pPrChange\b[^>]*/>|<w:pPrChange\b[^>]*>.*?</w:pPrChange>", re.S)
RPRCHANGE = re.compile(r"<w:rPrChange\b[^>]*/>|<w:rPrChange\b[^>]*>.*?</w:rPrChange>", re.S)


def split_ppr(body: str):
    """从段落正文开头按配对计数取出 pPr。返回 (ppr, rest)；没有 pPr 时 ppr 为空串。"""
    m = re.match(r"\s*<w:pPr(?:\s[^>]*)?(/?)>", body)
    if not m:
        return "", body
    if m.group(1) == "/":                     # <w:pPr/> 自闭合
        return body[m.start():m.end()].lstrip(), body[m.end():]
    depth = 1
    for mm in PPR_TAG.finditer(body, m.end()):
        tag = mm.group(0)
        if tag.startswith("</"):
            depth -= 1
            if depth == 0:
                return body[m.start():mm.end()].lstrip(), body[mm.end():]
        elif not tag.endswith("/>"):
            depth += 1
    # 配对不上：文件本身有问题，宁可当作没有 pPr 也不能切出半截
    return "", body


def extract_rpr(seg: str) -> str:
    """从 run 内容开头按配对计数取出完整 rPr（含嵌套的 rPrChange 内层 rPr）。
    WPS/Word 的格式修订会在 run 级 rPr 里嵌 <w:rPrChange><w:rPr>…</w:rPr></w:rPrChange>，
    非贪婪正则在内层 </w:rPr> 处截断，render 出的 run 就是非法 XML（Word 报文件损坏）。
    没有 rPr 时返回空串。"""
    m = re.match(r"\s*<w:rPr(?:\s[^>]*)?(/?)>", seg)
    if not m:
        return ""
    if m.group(1) == "/":
        return seg[m.start():m.end()].lstrip()
    depth = 1
    for mm in RPR_TAG.finditer(seg, m.end()):
        tag = mm.group(0)
        if tag.startswith("</"):
            depth -= 1
            if depth == 0:
                return seg[m.start():mm.end()].lstrip()
        elif not tag.endswith("/>"):
            depth += 1
    return ""          # 配对不上：宁可不带 rPr 也不能带半截


def pmark_flags(ppr: str):
    """段落标记（pPr 里第一层 rPr）的修订状态：(有 w:ins, 有 w:del, 是否 PENDING)。
    先剥离 pPrChange / rPrChange，避免把历史格式记录里的标记当成现在的。"""
    core = PPRCHANGE.sub("", ppr or "")
    m = re.search(r"<w:rPr>", core)
    if not m:
        return False, False, False
    seg = RPRCHANGE.sub("", core[m.start():])
    m2 = re.search(r"<w:rPr>(.*?)</w:rPr>", seg, re.S)
    inner = m2.group(1) if m2 else ""
    ins = re.search(r"<w:ins\s[^>]*/>", inner)
    dele = re.search(r"<w:del\s[^>]*/>", inner)
    pending = bool((ins and 'w:author="PENDING"' in ins.group(0))
                   or (dele and 'w:author="PENDING"' in dele.group(0)))
    return bool(ins), bool(dele), pending


def mark_pmark(ppr: str, marker: str) -> str:
    """把 w:ins/w:del 空元素插到段落标记 rPr 的第一个子元素位置。
    pPrChange 是 pPr 的最后一个子元素，先剥离、插完再放回，避免插进它嵌套的内层。"""
    m = PPRCHANGE.search(ppr)
    change = m.group(0) if m else ""
    core = (ppr[:m.start()] + ppr[m.end():]) if m else ppr
    if "<w:rPr>" in core:
        core = core.replace("<w:rPr>", "<w:rPr>" + marker, 1)
    else:
        core = core.replace("</w:pPr>", f"<w:rPr>{marker}</w:rPr></w:pPr>")
    if change:
        core = core.replace("</w:pPr>", change + "</w:pPr>")
    return core


def clean_pmark(ppr: str) -> str:
    """给新段落复用的 pPr：去掉修订标记与 pPrChange/rPrChange 历史记录
    （新段落不该背着参照段的'格式曾被改过'的历史）。"""
    core = PPRCHANGE.sub("", ppr or "")
    core = RPRCHANGE.sub("", core)
    return re.sub(r"<w:ins\s[^>]*/>|<w:del\s[^>]*/>", "", core)


class Tok:
    __slots__ = ("kind", "raw", "rpr", "text")

    def __init__(self, kind, raw, rpr="", text=""):
        self.kind, self.raw, self.rpr, self.text = kind, raw, rpr, text

    def render(self):
        if self.kind != "run":
            return self.raw
        if self.text == "":
            return ""
        return f'<w:r>{self.rpr}<w:t xml:space="preserve">{escape(self.text)}</w:t></w:r>'


def tokenize(body: str):
    """把段落正文（去掉 pPr 之后的部分）切成 token 序列。"""
    toks, pos = [], 0
    for m in BLOCK.finditer(body):
        if m.start() > pos:
            toks.append(Tok("raw", body[pos:m.start()]))
        raw = m.group(0)
        if raw.startswith("<w:r") and not raw.startswith("<w:rPr"):
            tm = T.search(raw)
            # 只有"纯文字 run"才可切分；含图片/域代码/制表符的按原样保留
            simple = tm and not re.search(
                r"<w:(drawing|pict|object|fldChar|instrText|tab|br|sym|footnoteReference|"
                r"endnoteReference|commentReference)\b", raw)
            if simple:
                rpr = extract_rpr(m.group(0)[m.group(0).index(">") + 1:])
                toks.append(Tok("run", raw, rpr, unescape(tm.group(1))))
            else:
                toks.append(Tok("raw", raw))
        else:
            toks.append(Tok("raw", raw))
        pos = m.end()
    if pos < len(body):
        toks.append(Tok("raw", body[pos:]))
    return toks


def index_map(toks):
    text, idx = [], []
    for i, t in enumerate(toks):
        if t.kind != "run":
            continue
        for j, ch in enumerate(t.text):
            text.append(ch)
            idx.append((i, j))
    return "".join(text), idx


def find_span(toks, needle, nth=1):
    """返回 (i0, o0, i1, o1)：[i0,o0) 起、到 (i1,o1] 止（o1 为结束后一位）。"""
    text, idx = index_map(toks)
    start = -1
    for _ in range(nth):
        start = text.find(needle, start + 1)
        if start < 0:
            return None
    end = start + len(needle) - 1
    i0, o0 = idx[start]
    i1, o1 = idx[end]
    return i0, o0, i1, o1 + 1


def slice_tokens(toks, i0, o0, i1, o1):
    """按跨 run 的区间切成 (前, 中, 后) 三段 token 列表；中段保留各自 rPr。"""
    head = [Tok("run", "", toks[i0].rpr, toks[i0].text[:o0])] if o0 else []
    tail = [Tok("run", "", toks[i1].rpr, toks[i1].text[o1:])] if o1 < len(toks[i1].text) else []
    mid = []
    for i in range(i0, i1 + 1):
        t = toks[i]
        if t.kind != "run":
            mid.append(t)
            continue
        s = o0 if i == i0 else 0
        e = o1 if i == i1 else len(t.text)
        if t.text[s:e]:
            mid.append(Tok("run", "", t.rpr, t.text[s:e]))
    return head, mid, tail


def render(toks):
    return "".join(t.render() for t in toks)
