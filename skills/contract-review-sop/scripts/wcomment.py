#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""wcomment.py —— 给已解包的 docx 目录加 Word 批注（锚定到指定原文）。

用法:
    python3 scripts/wcomment.py work/unpacked --anchor "乙方应于验收合格后30日内" \\
        --text "提请注意：……" [--occurrence 1] [--author "某某律所 张三"] [--initials 张]
    python3 scripts/wcomment.py work/unpacked --para 12 --text "……"      # 按段落序号锚定
    python3 scripts/wcomment.py work/unpacked --list-para 10 14           # 查看段落纯文本

设计取舍（都是为了"只改该改的地方"）:
  * 写 comments.xml / commentsExtended.xml / commentsIds.xml + rels + [Content_Types]。
    commentsExtensible.xml（真 Word 存放批注真 UTC 的地方）由 stamp.py 在分配时间时
    统一补条目 —— 来件既有的该部件与其中条目原样保留，绝不删除（删了会留下悬空的
    Relationship/Override，Word 直接报文件损坏）。
  * 批注样式（annotation reference / annotation text）按来件 styles.xml 解析真实 styleId
    （中文模板里常是 'ad' 这类压缩 id），查不到就不写样式，绝不写死 'CommentReference'。
  * 锚定时把命中的 w:r 拆成 前/中/后 三个 run 并原样复制 rPr，不改动任何文字与格式。
  * 时间戳先写占位值，最后统一由 stamp.py 分配（见 SKILL.md）。
"""
import argparse
import random
import re
from pathlib import Path
from xml.sax.saxutils import escape

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from wpara import tokenize, index_map, find_span, slice_tokens, render, split_ppr  # noqa: E402

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


W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
W16 = 'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"'
MC = ('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
PLACEHOLDER_DATE = "1999-01-01T00:00:00"

PARA = re.compile(r"<w:p(?:\s[^>]*)?>.*?</w:p>|<w:p(?:\s[^>]*)?/>", re.S)
RUN = re.compile(r"<w:r(?:\s[^>]*)?>(.*?)</w:r>", re.S)
T = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)


def para_id():
    return "%08X" % random.randint(0x10000000, 0x7FFFFFFF)


def read(p: Path, default=None):
    return p.read_text(encoding="utf-8") if p.exists() else default


def ensure_part(unp: Path, name: str, root_open: str, root_close: str):
    p = unp / "word" / name
    if not p.exists():
        p.write_text('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     + root_open + root_close, encoding="utf-8")
    return p


def ensure_rel_and_ct(unp: Path):
    rels = unp / "word" / "_rels" / "document.xml.rels"
    x = read(rels, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '</Relationships>')
    base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    ms = "http://schemas.microsoft.com/office/2011/relationships/"
    ms16 = "http://schemas.microsoft.com/office/2016/09/relationships/"
    want = [(base + "comments", "comments.xml"),
            (ms + "commentsExtended", "commentsExtended.xml"),
            (ms16 + "commentsIds", "commentsIds.xml")]
    ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', x)] or [0]
    nid = max(ids) + 1
    for typ, tgt in want:
        if f'Target="{tgt}"' in x:
            continue
        x = x.replace("</Relationships>",
                      f'<Relationship Id="rId{nid}" Type="{typ}" Target="{tgt}"/></Relationships>')
        nid += 1
    rels.parent.mkdir(parents=True, exist_ok=True)
    rels.write_text(x, encoding="utf-8")

    ct = unp / "[Content_Types].xml"
    c = read(ct)
    over = {
        "/word/comments.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        "/word/commentsExtended.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
        "/word/commentsIds.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
    }
    for part, tp in over.items():
        if f'PartName="{part}"' not in c:
            c = c.replace("</Types>", f'<Override PartName="{part}" ContentType="{tp}"/></Types>')
    ct.write_text(c, encoding="utf-8")


def next_id(comments_xml: str) -> int:
    ids = [int(i) for i in re.findall(r'<w:comment\s[^>]*w:id="(\d+)"', comments_xml)]
    return (max(ids) + 1) if ids else 0


def comment_styles(unp: Path):
    """从 styles.xml 解析批注样式的真实 styleId。
    中文环境的 Word 文件里批注引用样式的 id 往往是 'ad' 这类压缩 id（对照
    修订模式模板实件校准），写死 'CommentReference' 会引用一个不存在的样式。
    按 w:name 的内部名（annotation reference / annotation text）查；查不到就不写样式。"""
    sp = unp / "word" / "styles.xml"
    ref = txt = None
    if sp.exists():
        s = sp.read_text(encoding="utf-8")
        for m in re.finditer(
                r'<w:style\s[^>]*w:styleId="([^"]+)"[^>]*>\s*<w:name w:val="([^"]+)"', s):
            if m.group(2) == "annotation reference":
                ref = m.group(1)
            elif m.group(2) == "annotation text":
                txt = m.group(1)
    return ref, txt


def add_comment_part(unp: Path, cid: int, text: str, author: str, initials: str, parent=None):
    # 文档是 Word 2023+（声明了 w16du）时，批注部件也跟着声明，好让 stamp.py 写 dateUtc；
    # 老文档就不声明 —— 版本特征要和来件保持同一个年代。
    dx = (unp / "word" / "document.xml").read_text(encoding="utf-8")[:4000]
    du = (' xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"'
          if "xmlns:w16du=" in dx else "")
    ign = "w14 w16du" if du else "w14"
    cp = ensure_part(unp, "comments.xml",
                     f'<w:comments {W} {W14} {MC}{du} mc:Ignorable="{ign}">', "</w:comments>")
    ce = ensure_part(unp, "commentsExtended.xml",
                     f'<w15:commentsEx {W} {W14} {W15} {MC} mc:Ignorable="w14 w15">',
                     "</w15:commentsEx>")
    ci = ensure_part(unp, "commentsIds.xml",
                     f'<w16cid:commentsIds {W} {W14} {W16} {MC} mc:Ignorable="w14 w16cid">',
                     "</w16cid:commentsIds>")
    ref_style, txt_style = comment_styles(unp)
    pstyle = f'<w:pPr><w:pStyle w:val="{txt_style}"/></w:pPr>' if txt_style else ""
    aref = (f'<w:r><w:rPr><w:rStyle w:val="{ref_style}"/></w:rPr><w:annotationRef/></w:r>'
            if ref_style else '<w:r><w:annotationRef/></w:r>')
    pid = para_id()
    body = "".join(
        f'<w:p w14:paraId="{pid if i == 0 else para_id()}" w14:textId="77777777">'
        f'{pstyle}{aref}'
        f'<w:r><w:t xml:space="preserve">{escape(line)}</w:t></w:r></w:p>'
        if i == 0 else
        f'<w:p w14:paraId="{para_id()}" w14:textId="77777777">'
        f'{pstyle}'
        f'<w:r><w:t xml:space="preserve">{escape(line)}</w:t></w:r></w:p>'
        for i, line in enumerate(text.split("\n")))
    x = read(cp)
    x = x.replace("</w:comments>",
                  f'<w:comment w:id="{cid}" w:author="{escape(author, {chr(34): "&quot;"})}" '
                  f'w:date="{PLACEHOLDER_DATE}" w:initials="{escape(initials)}">'
                  f'{body}</w:comment></w:comments>')
    cp.write_text(x, encoding="utf-8")

    e = read(ce)
    pa = f' w15:paraIdParent="{parent}"' if parent else ""
    e = e.replace("</w15:commentsEx>",
                  f'<w15:commentEx w15:paraId="{pid}"{pa} w15:done="0"/></w15:commentsEx>')
    ce.write_text(e, encoding="utf-8")

    i = read(ci)
    i = i.replace("</w16cid:commentsIds>",
                  f'<w16cid:commentId w16cid:paraId="{pid}" '
                  f'w16cid:durableId="{para_id()}"/></w16cid:commentsIds>')
    ci.write_text(i, encoding="utf-8")
    ensure_rel_and_ct(unp)
    return pid


def marker_ref(cid, ref_style=None):
    if ref_style:
        return (f'<w:r><w:rPr><w:rStyle w:val="{ref_style}"/></w:rPr>'
                f'<w:commentReference w:id="{cid}"/></w:r>')
    return f'<w:r><w:commentReference w:id="{cid}"/></w:r>'


def split_para(chunk):
    """pPr 用配对计数提取 —— 非贪婪正则遇 pPrChange 嵌套会截断出非法 XML。"""
    head_end = chunk.index(">") + 1
    head, rest = chunk[:head_end], chunk[head_end:]
    body = rest[: rest.rindex("</w:p>")] if "</w:p>" in rest else rest
    ppr, remainder = split_ppr(body)
    return head, ppr, remainder


def anchor_in_document(doc, cid, anchor=None, occurrence=1, para_index=None, ref_style=None):
    """把 commentRange 标记插入 document.xml（支持跨 run 锚定）。返回 (新 xml, 位置说明)。"""
    paras = list(PARA.finditer(doc))
    if para_index is not None:
        if para_index >= len(paras):
            raise SystemExit(f"[FAIL] 段落序号超界（共 {len(paras)} 段）")
        pm = paras[para_index]
        head, ppr, body = split_para(pm.group(0))
        newp = (head + ppr + f'<w:commentRangeStart w:id="{cid}"/>' + body
                + f'<w:commentRangeEnd w:id="{cid}"/>' + marker_ref(cid, ref_style) + "</w:p>")
        return doc[:pm.start()] + newp + doc[pm.end():], f"整段 {para_index}"

    seen = 0
    for pi, pm in enumerate(paras):
        head, ppr, body = split_para(pm.group(0))
        toks = tokenize(body)
        text, _ = index_map(toks)
        n = text.count(anchor)
        if not n or seen + n < occurrence:
            seen += n
            continue
        i0, o0, i1, o1 = find_span(toks, anchor, occurrence - seen)
        pre, mid, post = slice_tokens(toks, i0, o0, i1, o1)
        newbody = (render(toks[:i0]) + render(pre)
                   + f'<w:commentRangeStart w:id="{cid}"/>' + render(mid)
                   + f'<w:commentRangeEnd w:id="{cid}"/>' + marker_ref(cid, ref_style)
                   + render(post) + render(toks[i1 + 1:]))
        return doc[:pm.start()] + head + ppr + newbody + "</w:p>" + doc[pm.end():], \
            f"段落 {pi}「{anchor[:20]}…」"
    # 兜底：锚定串落在来件既有的修订块（w:ins/w:del）里 —— 批注不改内容，
    # 把标记放在整块的两侧是安全且合法的，只是范围粗一点（覆盖整个修订块）。
    for pi, pm in enumerate(paras):
        head, ppr, body = split_para(pm.group(0))
        toks = tokenize(body)
        for ti, t in enumerate(toks):
            if t.kind != "raw" or anchor not in "".join(T.findall(t.raw)):
                continue
            newbody = (render(toks[:ti]) + f'<w:commentRangeStart w:id="{cid}"/>'
                       + t.raw + f'<w:commentRangeEnd w:id="{cid}"/>' + marker_ref(cid, ref_style)
                       + render(toks[ti + 1:]))
            return doc[:pm.start()] + head + ppr + newbody + "</w:p>" + doc[pm.end():], \
                f"段落 {pi}（锚在来件既有修订块上，范围为整块）"

    raise SystemExit(f"[FAIL] 全文未找到锚定原文「{anchor}」的第 {occurrence} 处（共 {seen} 处）。"
                     "先用 wredline.py --grep 定位，或改用 --para 整段锚定。")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("unpacked")
    ap.add_argument("--text")
    ap.add_argument("--anchor")
    ap.add_argument("--para", type=int)
    ap.add_argument("--occurrence", type=int, default=1)
    ap.add_argument("--author", default=_CFG_AUTHOR)
    ap.add_argument("--initials", default=_CFG_INITIALS)
    ap.add_argument("--parent", help="父批注的 paraId（回复用）")
    ap.add_argument("--list-para", nargs=2, type=int, metavar=("FROM", "TO"))
    a = ap.parse_args()

    unp = Path(a.unpacked)
    docp = unp / "word" / "document.xml"
    doc = docp.read_text(encoding="utf-8")

    if a.list_para:
        paras = list(PARA.finditer(doc))
        for i in range(a.list_para[0], min(a.list_para[1], len(paras))):
            txt = "".join(T.findall(paras[i].group(0)))
            print(f"{i}\t{txt[:120]}")
        return

    if not a.text or (a.anchor is None and a.para is None):
        raise SystemExit("[FAIL] 需要 --text，且 --anchor 或 --para 二选一")

    cid = next_id(read(unp / "word" / "comments.xml", ""))
    pid = add_comment_part(unp, cid, a.text, a.author, a.initials, a.parent)
    ref_style, _ = comment_styles(unp)
    doc, where = anchor_in_document(doc, cid, a.anchor, a.occurrence, a.para, ref_style)
    docp.write_text(doc, encoding="utf-8")
    print(f"[OK] 批注 id={cid} paraId={pid} 锚定于 {where}（时间戳待 stamp.py 分配）")


if __name__ == "__main__":
    main()
