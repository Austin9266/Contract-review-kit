#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_index.py — 从合同/招标文件（.docx 或 .txt）中抽取两类信息：

1. 编号索引（index）：每个条号（如 "18.1"）出现在文档的哪个"编区"（scope，
   如"通用合同条款"/"专用合同条款"/"第二章 投标人须知"），条号标题与正文。
2. 交叉引用清单（refs）：每一处"第X.Y款/项/条/目"式引用，记录其所在编区、
   行号、引用编号、限定语（如"本章""专用合同条款""第二章"）与上下文。

输出 JSON 到 stdout 或文件，供后续语义核对使用。

用法：
  python3 extract_index.py <合同文件.docx|.txt> [-o 输出.json]
  python3 extract_index.py <合同文件.docx|.txt> --refs-only   # 只输出引用清单
  python3 extract_index.py <合同文件.docx|.txt> --index-only  # 只输出编号索引

注意：本脚本只做"机械提取与定位"，不做语义判断。语义核对（被引用条款的内容
是否与引用语境匹配）由使用本脚本的 agent 完成。
"""

import argparse
import html
import json
import re
import sys
import zipfile


# ---------- 文本抽取 ----------

def docx_to_paragraphs(path: str):
    """从 .docx 提取段落文本列表（不依赖第三方库）。"""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    # 源码中的换行会打断正则匹配，先去除
    xml = xml.replace("\n", "").replace("\r", "")
    paras = re.findall(r"<w:p\b.*?</w:p>", xml, re.S)
    out = []
    for para in paras:
        texts = re.findall(r"<w:t\b[^>]*>([^<]*)</w:t>", para, re.S)
        out.append(html.unescape("".join(texts)))
    return out


def txt_to_paragraphs(path: str):
    with open(path, encoding="utf-8") as f:
        return f.read().split("\n")


# ---------- 编区（scope）识别 ----------
# 合同常由多个独立编号的板块组成：各"章"、通用条款、专用条款等。
# 每个板块内部条号自成体系，不同板块可有相同条号（如通用 18.1 与专用 18.1）。

SCOPE_PATTERNS = [
    # "第一章 招标公告" / "第二章  投标人须知"（整行，允许前置空白）
    re.compile(r"^\s*第([一二三四五六七八九十百]+)章[ 　]*(.*)$"),
    # "第一部分 合同协议书" / "第二部分 通用合同条款"
    re.compile(r"^\s*第([一二三四五六七八九十百]+)部分[ 　]*(.*)$"),
    # 独立的"通用合同条款"/"专用合同条款"标题行
    re.compile(r"^\s*(通用合同条款|专用合同条款|通用条款|专用条款)\s*$"),
    # "附件一"/"附录A" 等
    re.compile(r"^\s*(附件|附录)\s*([一二三四五六七八九十百A-Z0-9]+)[ 　]*(.*)$"),
]

CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7,
          "八": 8, "九": 9, "十": 10, "百": 100}


def cn_to_int(s: str) -> int:
    if s in CN_NUM:
        return CN_NUM[s]
    if s.startswith("十"):
        return 10 + CN_NUM.get(s[1:], 0)
    if "十" in s:
        a, b = s.split("十", 1)
        return CN_NUM.get(a, 1) * 10 + CN_NUM.get(b, 0)
    return 0


def detect_scopes(lines):
    """返回 [(行号, scope名)]，scope名如 '第二章 投标人须知'。

    目录页中的条目（"第一章 招标公告2" 这种标题后紧跟页码的行）会被排除；
    若同一编区名出现多次，以首次"干净"出现的为准，目录行被跳过即可。
    """
    marks = []
    for i, raw in enumerate(lines):
        s = raw.strip()
        if not s or len(s) > 45:
            continue
        # 目录行特征：标题文字后直接跟页码数字（可能带点线残留），如 "第一章  招标公告2"
        if re.search(r"[^\d]\d{1,4}\s*$", s):
            continue
        for pat in SCOPE_PATTERNS:
            m = pat.match(s)
            if m:
                marks.append((i, s))
                break
    return marks


def scope_of(line_no: int, marks) -> str:
    cur = "文首（未识别编区）"
    for start, name in marks:
        if line_no >= start:
            cur = name
        else:
            break
    return cur


# ---------- 条号索引 ----------

# 条号标题行：以数字编号开头，如 "18.1 竣工试验"、"1一般约定"、"1.1.1.4 投标函：..."
# 编号深度 1-4 级；标题行通常较短或以中文紧跟。
CLAUSE_HEAD = re.compile(
    r"^\s*(\d{1,2}(?:\.\d{1,2}){0,3})[ 　]*([^\d\s\.].*)$"
)

# 交叉引用：第X款 / 第X.Y款 / 第X.Y.Z项 / 第X条 / 第X.Y.Z目
# 同时捕获引用前的限定语（"本章""本合同""专用合同条款""第二章..."等，最多20字）
REF_PAT = re.compile(
    r"(?P<qual>本章|本合同|本招标文件|招标文件|合同|专用合同条款|通用合同条款|"
    r"专用条款|通用条款|第[一二三四五六七八九十百]+章[^，。；,;]{0,12}?)?"
    r"第\s*(?P<num>\d+(?:\.\s*\d+){0,3})\s*(?P<unit>条|款|项|目)"
)


def build_index(lines, marks):
    """建立 {(scope, 条号): {line, title, text}}。text 取条号行至下一个同级或更高级条号前的内容。"""
    entries = []  # (line, scope, num, title)
    for i, raw in enumerate(lines):
        s = raw.strip()
        if not s or len(s) > 120:
            # 条号标题行一般不长；过长的行多半是正文，跳过可提高精度
            # 但不少合同条号与正文同段，故仍记录候选
            pass
        m = CLAUSE_HEAD.match(s)
        if not m:
            continue
        num = m.group(1)
        rest = m.group(2) or ""
        # 排除明显非条号的行：纯数字行、比例、日期等
        if re.match(r"^[\d\s%.、，,-]+$", s):
            continue
        # 排除"第X章"标题行已被 scope 捕获的（条号正则不会匹配中文章号，安全）
        entries.append({
            "line": i,
            "scope": scope_of(i, marks),
            "num": num,
            "title": rest[:60],
        })
    return entries


def collect_refs(lines, marks):
    refs = []
    for i, raw in enumerate(lines):
        s = raw.strip()
        if not s:
            continue
        for m in REF_PAT.finditer(s):
            num = re.sub(r"\s+", "", m.group("num"))
            qual = (m.group("qual") or "").strip()
            refs.append({
                "line": i,
                "scope": scope_of(i, marks),
                "ref": f"{num}{m.group('unit')}",
                "num": num,
                "unit": m.group("unit"),
                "qualifier": qual,
                "context": s[:120],
            })
    return refs


def main():
    ap = argparse.ArgumentParser(description="提取合同编号索引与交叉引用清单")
    ap.add_argument("file", help="合同文件路径（.docx 或 .txt）")
    ap.add_argument("-o", "--output", help="输出 JSON 文件路径（默认 stdout）")
    ap.add_argument("--refs-only", action="store_true", help="只输出交叉引用清单")
    ap.add_argument("--index-only", action="store_true", help="只输出条号索引")
    args = ap.parse_args()

    if args.file.lower().endswith(".docx"):
        lines = docx_to_paragraphs(args.file)
    else:
        lines = txt_to_paragraphs(args.file)

    marks = detect_scopes(lines)
    index = build_index(lines, marks)
    refs = collect_refs(lines, marks)

    result = {
        "file": args.file,
        "paragraphs": len(lines),
        "scopes": [{"line": ln, "name": name} for ln, name in marks],
    }
    if not args.refs_only:
        result["index"] = index
    if not args.index_only:
        result["references"] = refs

    out = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(out)
        print(f"已写入 {args.output}（编区 {len(marks)} 个，条号 {len(index)} 条，引用 {len(refs)} 处）",
              file=sys.stderr)
    else:
        print(out)


if __name__ == "__main__":
    main()
