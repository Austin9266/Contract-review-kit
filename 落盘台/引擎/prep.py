#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""prep.py —— 落盘前的准备：规范输入、解包、留存基线。

用法:
    python3 scripts/prep.py 来件.doc  --work work/
    python3 scripts/prep.py 来件.docx --work work/

产物（work/ 目录下）:
    origin.<ext>        来件原件的只读副本（永远不动它）
    base.docx           工作用 docx（来件是 .docx 时即原件副本；是 .doc 时为 soffice 转换结果）
    unpacked/           base.docx 解包目录，后续所有编辑只改这里的 word/document.xml
    baseline.json       基线清单：来件后缀、各 zip 条目的 sha256、正文纯文本
    baseline.txt        基线正文纯文本（供人工比对）

注意:
    - 来件是 .doc 时，交付仍需回转为 .doc（见 SKILL.md 第 6 步），本脚本只负责入口方向。
    - 本脚本不修改来件，不做 merge_runs 之外的任何内容改动。
"""
import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from wtext import document_text  # noqa: E402


def sh(cmd, **kw):
    return subprocess.run(cmd, check=True, capture_output=True, text=True, **kw)


def convert(src: Path, target: str, outdir: Path) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    sh(["soffice", "--headless", "--norestore", "--convert-to", target,
        "--outdir", str(outdir), str(src)])
    out = outdir / (src.stem + "." + target.split(":")[0])
    if not out.exists():
        raise SystemExit(f"[FAIL] soffice 未生成 {out}")
    return out


SIMPLE_RUN = re.compile(
    r"<w:r>(<w:rPr>.*?</w:rPr>)?<w:t(?: xml:space=\"preserve\")?>(.*?)</w:t></w:r>", re.S)


def merge_runs(doc_xml: Path):
    """合并相邻的、格式完全相同的 w:r，使正文可被整串检索。不改变渲染结果。

    单趟线性扫描：把"首尾相接且 rPr 相同"的整串 run 一次并成一个。
    早期版本是"每轮只并一对、最多 40 轮"的多轮全文重扫，在 1.5MB/千余段的
    文档上要跑约 2 分钟；现在同样输入 1 秒内完成，输出一致。"""
    xml = doc_xml.read_text(encoding="utf-8")
    out = []
    pos = 0                       # 已输出到的位置
    run_rpr = None                # 当前累积串的 rPr
    run_texts = []                # 当前累积串的文本
    run_start = run_end = -1

    def flush():
        nonlocal run_texts, run_rpr, run_start, run_end
        if not run_texts:
            return
        if len(run_texts) == 1:
            out.append(xml[run_start:run_end])          # 单个 run 原样保留
        else:
            out.append('<w:r>' + (run_rpr or "") + '<w:t xml:space="preserve">'
                       + "".join(run_texts) + '</w:t></w:r>')
        run_texts, run_rpr, run_start, run_end = [], None, -1, -1

    for m in SIMPLE_RUN.finditer(xml):
        rpr = m.group(1) or ""
        if run_texts and m.start() == run_end and rpr == run_rpr:
            run_texts.append(m.group(2))
            run_end = m.end()
            pos = m.end()
            continue
        flush()
        if pos < m.start():
            out.append(xml[pos:m.start()])
        run_rpr, run_texts = rpr, [m.group(2)]
        run_start, run_end = m.start(), m.end()
        pos = m.end()
    flush()
    out.append(xml[pos:])
    doc_xml.write_text("".join(out), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--work", default="work")
    ap.add_argument("--no-merge-runs", action="store_true")
    a = ap.parse_args()

    src = Path(a.input).resolve()
    work = Path(a.work).resolve()
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)

    ext = src.suffix.lower()
    if ext not in (".doc", ".docx"):
        raise SystemExit(f"[FAIL] 只支持 .doc/.docx，收到 {ext}")

    origin = work / ("origin" + ext)
    shutil.copy2(src, origin)

    if ext == ".docx":
        base = work / "base.docx"
        shutil.copy2(origin, base)
        converted = False
    else:
        base = convert(origin, "docx", work / "_conv")
        base = shutil.copy2(base, work / "base.docx") and (work / "base.docx")
        converted = True

    unpacked = work / "unpacked"
    with zipfile.ZipFile(base) as z:
        names = z.namelist()
        z.extractall(unpacked)
    for p in unpacked.rglob("*"):
        if p.is_symlink():
            p.unlink()

    doc_xml = unpacked / "word" / "document.xml"
    if not a.no_merge_runs:
        merge_runs(doc_xml)

    entries = {}
    with zipfile.ZipFile(base) as z:
        for n in names:
            entries[n] = hashlib.sha256(z.read(n)).hexdigest()

    dx = doc_xml.read_text(encoding="utf-8")
    # 来件里可能已经带着第三方（客户/对方律师/前手）的修订与批注 —— 这些是别人的署名，
    # 一个字都不许动。这里把它们逐条记下来，交付前由 verify.py C7 核对是否原封不动。
    pre = sorted(set(re.findall(r'<w:[A-Za-z]+\b[^>]*?\sw:author="([^"]*)"[^>]*?/?>', dx)))
    cxml = ""
    cpath = unpacked / "word" / "comments.xml"
    if cpath.exists():
        cxml = cpath.read_text(encoding="utf-8")
        pre = sorted(set(pre) | set(re.findall(r'<w:comment\s[^>]*w:author="([^"]*)"', cxml)))
    pre_pairs = sorted(set(
        re.findall(r'w:author="([^"]*)"\s+w:date="([^"]*)"', dx + cxml)
        + [(a, d) for d, a in re.findall(r'w:date="([^"]*)"\s+w:author="([^"]*)"', dx + cxml)]))
    # 来件既有 commentsExtensible 条目（durableId → dateUtc），交付前核对一字未改
    cextp = unpacked / "word" / "commentsExtensible.xml"
    pre_cext = (dict(re.findall(
        r'w16cex:durableId="([0-9A-Fa-f]+)"\s+w16cex:dateUtc="([^"]*)"',
        cextp.read_text(encoding="utf-8"))) if cextp.exists() else {})

    text = document_text(dx, mode="reject")
    (work / "baseline.txt").write_text(text, encoding="utf-8")
    (work / "baseline.json").write_text(json.dumps({
        "source_name": src.name,
        "source_ext": ext,
        "converted_from_doc": converted,
        "entries": entries,
        "text_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
        "paragraphs": text.count("\n") + 1,
        "pre_authors": pre,
        "pre_stamps": [list(t) for t in pre_pairs],
        "pre_cext": pre_cext,
        # 来件本身就重复的修订 id（第三方文件不保证干净）——交付时只对"本次新产生的
        # 重复"拦截，见 verify C0
        "pre_dup_ids": sorted(
            k for k, v in __import__("collections").Counter(re.findall(
                r'<w:(?:ins|del|rPrChange|pPrChange|moveFrom|moveTo|sectPrChange|tblPrChange|'
                r'trPrChange|tcPrChange)\b[^>]*?\sw:id="(\d+)"', dx)).items() if v > 1),
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[OK] 来件后缀 : {ext}{'（已转 docx 用于编辑，交付需回转 .doc）' if converted else ''}")
    print(f"[OK] 工作目录 : {unpacked}")
    print(f"[OK] 待编辑   : {doc_xml}")
    print(f"[OK] 基线段落 : {text.count(chr(10)) + 1}")
    if pre:
        print(f"[!] 来件已带修订/批注，作者：{'、'.join(pre)}（共 {len(pre_pairs)} 处署名）")
        print("    这些是别人的痕迹，本流程只增不改；stamp.py 不会碰它们，"
              "verify.py C7 会逐条核对原封不动。")


if __name__ == "__main__":
    main()
