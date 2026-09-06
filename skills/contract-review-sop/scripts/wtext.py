#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""wtext.py —— 从 word/document.xml 原始文本中抽取正文纯文本。

mode="reject" : 拒绝全部修订后的文本（= 来件原文）
mode="accept" : 接受全部修订后的文本（= 定稿文本）

"拒绝全部修订后 == 来件原文" 是本技能最重要的完整性校验：
只要相等，就证明除修订标记外没有动过任何正文文字。
"""
import re
import sys
from pathlib import Path
from xml.sax.saxutils import unescape

sys.path.insert(0, str(Path(__file__).resolve().parent))
from wpara import split_ppr, pmark_flags  # noqa: E402

TAG = re.compile(
    r"<w:p(?:\s[^>]*)?>|</w:p>|"
    r"<w:ins\s[^>]*?/>|<w:del\s[^>]*?/>|"
    r"<w:ins(?:\s[^>]*)?>|</w:ins>|<w:del(?:\s[^>]*)?>|</w:del>|"
    r"<w:moveFrom(?:\s[^>]*)?>|</w:moveFrom>|<w:moveTo(?:\s[^>]*)?>|</w:moveTo>|"
    r"<w:t(?:\s[^>]*)?>(.*?)</w:t>|<w:delText(?:\s[^>]*)?>(.*?)</w:delText>|"
    r"<w:tab\s*/>|<w:br\s*/>|<w:noBreakHyphen\s*/>",
    re.S)

PARA = re.compile(r"<w:p(?:\s[^>]*)?>(.*?)</w:p>", re.S)


def document_text(xml: str, mode: str = "reject") -> str:
    body = xml.split("<w:body>", 1)[-1]
    out = []
    for pm in PARA.finditer(body):
        chunk = pm.group(1)
        # pPr 按配对计数提取、段落标记状态剥离 pPrChange/rPrChange 后判断
        # （非贪婪正则遇嵌套会截断，历史格式记录里的旧标记会被误当成现在的）
        ppr, rest = split_ppr(chunk)
        ins_mark, del_mark, _ = pmark_flags(ppr)

        buf = []
        ins_depth = 0
        del_depth = 0
        for m in TAG.finditer(rest):
            s = m.group(0)
            if m.group(1) is not None:                # <w:t>
                if not (mode == "reject" and ins_depth):
                    buf.append(unescape(m.group(1)))
                continue
            if m.group(2) is not None:                # <w:delText>
                if mode != "accept":
                    buf.append(unescape(m.group(2)))
                continue
            if s.startswith("<w:ins") and s.endswith("/>"):
                continue
            if s.startswith("<w:del") and s.endswith("/>"):
                continue
            if s.startswith("<w:ins") or s.startswith("<w:moveTo"):
                ins_depth += 1
            elif s.startswith("</w:ins") or s.startswith("</w:moveTo"):
                ins_depth = max(0, ins_depth - 1)
            elif s.startswith("<w:del") or s.startswith("<w:moveFrom"):
                del_depth += 1
            elif s.startswith("</w:del") or s.startswith("</w:moveFrom"):
                del_depth = max(0, del_depth - 1)
            elif s.startswith("<w:tab"):
                if not (mode == "reject" and ins_depth) and not (mode == "accept" and del_depth):
                    buf.append("\t")
            elif s.startswith("<w:br") or s.startswith("<w:noBreakHyphen"):
                if not (mode == "reject" and ins_depth) and not (mode == "accept" and del_depth):
                    buf.append("\n" if s.startswith("<w:br") else "-")
        out.append("".join(buf))

        # 段落标记本身被插入 → 拒绝时这一段与下一段本是同一段
        if mode == "reject" and ins_mark:
            out.append("\x00JOIN\x00")
        if mode == "accept" and del_mark:
            out.append("\x00JOIN\x00")

    text = "\n".join(out)
    text = re.sub(r"\n?\x00JOIN\x00\n?", "", text)
    return text


if __name__ == "__main__":
    import sys
    import zipfile
    src, mode = sys.argv[1], (sys.argv[2] if len(sys.argv) > 2 else "reject")
    if src.endswith(".docx"):
        with zipfile.ZipFile(src) as z:
            xml = z.read("word/document.xml").decode("utf-8")
    else:
        xml = open(src, encoding="utf-8").read()
    print(document_text(xml, mode))
