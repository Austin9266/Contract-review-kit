#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""wredline.py —— 在已解包的 docx 目录里做"带修订痕迹"的替换/删除/插入。

用法:
    # 改字：原文 → 新文（生成 w:del + w:ins）
    python3 scripts/wredline.py work/unpacked --find "××市仲裁委员会" --replace "××仲裁委员会"

    # 纯删除
    python3 scripts/wredline.py work/unpacked --find "并按日万分之五支付违约金" --replace ""

    # 在某段原文之后插入一句
    python3 scripts/wredline.py work/unpacked --find "验收合格之日起" --insert-after "30 日内"

    # 整段新增（沿用参照段的 pPr，编号/样式随之一致）
    python3 scripts/wredline.py work/unpacked --after-para 37 --insert-paragraph "9.4 ……"

    # 整段删除
    python3 scripts/wredline.py work/unpacked --delete-para 41

    # 定位辅助
    python3 scripts/wredline.py work/unpacked --grep "仲裁"
    python3 scripts/wredline.py work/unpacked --list 30 45

为什么必须用它而不是手改 XML:
    它只把命中的那几个 w:r 拆开重组，各自的 rPr 原样复制，前后文一个字符都不碰；
    手改极易顺手动到相邻 run 的格式或漏掉 w:delText，这正是"偶发性改动"的来源。
时间戳与署名一律先占位（PENDING），最后由 stamp.py 统一分配。
"""
import argparse
import re
import sys
from pathlib import Path
from xml.sax.saxutils import escape

sys.path.insert(0, str(Path(__file__).resolve().parent))
from wpara import (tokenize, index_map, find_span, slice_tokens, render,  # noqa: E402
                   split_ppr, pmark_flags, mark_pmark, clean_pmark)

PARA = re.compile(r"<w:p(?:\s[^>]*)?>.*?</w:p>|<w:p(?:\s[^>]*)?/>", re.S)
T = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)
PLACEHOLDER = 'w:author="PENDING" w:date="1999-01-01T00:00:00"'


def split_para(chunk):
    """<w:p …> 头 / pPr / 正文 三段拆开。
    pPr 用配对计数提取（wpara.split_ppr）——非贪婪正则遇到 pPrChange 嵌套的
    内层 </w:pPr> 会截断，切分后产出非法 XML。"""
    head_end = chunk.index(">") + 1
    head, rest = chunk[:head_end], chunk[head_end:]
    body = rest[: rest.rindex("</w:p>")] if "</w:p>" in rest else rest
    ppr, remainder = split_ppr(body)
    return head, ppr, remainder


def new_id(doc):
    ids = [int(i) for i in re.findall(r'<w:(?:ins|del)\s[^>]*w:id="(\d+)"', doc)]
    return (max(ids) + 1) if ids else 0   # Word 自己就是从 0 开始编号


def del_block(wid, toks):
    inner = []
    for t in toks:
        if t.kind != "run":
            raise SystemExit("[FAIL] 待删除区间里混有图片/域代码/已有修订，"
                             "请缩小 --find 范围或改用 --delete-para")
        if t.text:
            inner.append(f'<w:r>{t.rpr}<w:delText xml:space="preserve">'
                         f'{escape(t.text)}</w:delText></w:r>')
    return f'<w:del w:id="{wid}" {PLACEHOLDER}>{"".join(inner)}</w:del>' if inner else ""


def ins_block(wid, rpr, text):
    if not text:
        return ""
    return (f'<w:ins w:id="{wid}" {PLACEHOLDER}><w:r>{rpr}'
            f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:ins>')


CHANGE_ID = re.compile(
    r'<w:(ins|del|rPrChange|pPrChange|moveFrom|moveTo|sectPrChange|tblPrChange|'
    r'trPrChange|tcPrChange)\b[^>]*?\sw:id="(\d+)"')


def dedupe_change_ids(doc: str) -> str:
    """把重复的修订 w:id 重新编号（只改后出现的副本，保留首次出现的原 id）。

    为什么需要：切分一个带 <w:rPrChange> 的 run 时，每一份都要复制完整的 rPr
    （否则被切开的文字就丢了第三方的格式修订记录），于是同一个 rPrChange 的
    w:id 会出现好几次。ECMA-376 要求修订 id 唯一，真 Word 拆 run 时也是给新
    副本分配新 id。这里只动 id 这个数字 —— 作者、时间、被记录的旧格式一字不动，
    而 id 本来就是 Word 每次保存时统一重编的（见 references 第 12 条）。"""
    seen, used = set(), set()
    for m in CHANGE_ID.finditer(doc):
        used.add(int(m.group(2)))
    nxt = (max(used) + 1) if used else 0
    out, pos = [], 0
    for m in CHANGE_ID.finditer(doc):
        wid = m.group(2)
        if wid not in seen:
            seen.add(wid)
            continue
        nxt += 1
        s, e = m.span()
        out.append(doc[pos:s])
        out.append(re.sub(r'(\sw:id=")\d+(")', lambda x: x.group(1) + str(nxt) + x.group(2),
                          m.group(0), count=1))
        pos = e
    if not out:
        return doc
    out.append(doc[pos:])
    return "".join(out)


def new_para_tag(doc, rnd):
    """文档在用 w14:paraId 时，新段也要带 —— 只有 LibreOffice 转出来的文件才是光秃秃的 <w:p>。"""
    if "w14:paraId" not in doc:
        return "<w:p>"
    pid = "%08X" % rnd.randint(0x10000000, 0x7FFFFFFF)
    tid = "%08X" % rnd.randint(0x10000000, 0x7FFFFFFF)
    return f'<w:p w14:paraId="{pid}" w14:textId="{tid}">'


def insert_paras(doc, n, texts, rnd):
    """在第 n 段之后一次性插入 k 段（apply.py 与本文件 CLI 共用）。

    常规情形（锚点段的段落标记干净）：Word 的真实动作是在第 n 段末尾连敲 k 次
    回车 —— 新产生的 k 个段落标记依次属于第 n 段和前 k-1 个新段，最后一个新段
    沿用第 n 段原有的段落标记。所以 ins 标记顺着链条往下打。

    锚点段的段落标记已带 w:ins（来件里客户/对方律师/前次交付的插入）时：
    锚点段一个字都不动 —— 让 k 个新段各自携带自己的 w:ins 段落标记，正文也在
    w:ins 里。拒绝全部修订时 n → new1 → … → newk → next 整链合并，可完整还原来件。
    """
    wid = new_id(doc)
    paras = list(PARA.finditer(doc))
    if n >= len(paras):
        raise SystemExit(f"段落序号 {n} 超界（共 {len(paras)} 段）")
    pm = paras[n]
    head, ppr, body = split_para(pm.group(0))
    ppr = ppr or "<w:pPr></w:pPr>"
    has_ins, _, pending = pmark_flags(ppr)
    clean = clean_pmark(ppr)
    k = len(texts)
    if has_ins and pending:
        raise SystemExit("该段的段落标记带本次刚打的 w:ins 占位，"
                         "请把要插的几段写在同一个 after_para 条目组里，一次插入")
    if has_ins:
        # 锚点段落标记是来件既有的第三方插入：不动它，新段自带 w:ins 段落标记
        out = head + ppr + body + "</w:p>"
        for t in texts:
            p_ppr = mark_pmark(clean, f'<w:ins w:id="{wid}" {PLACEHOLDER}/>')
            wid += 1
            out += new_para_tag(doc, rnd) + p_ppr + ins_block(wid, "", t) + "</w:p>"
            wid += 1
        how = f"段落 {n} 之后新增 {k} 段（锚点段标记为来件既有插入，原样保留）"
    else:
        out = head + mark_pmark(ppr, f'<w:ins w:id="{wid}" {PLACEHOLDER}/>') + body + "</w:p>"
        for i, t in enumerate(texts):
            wid += 1
            # 最后一段承接原段落标记，其余继续往下传 ins
            p_ppr = clean if i == k - 1 else mark_pmark(clean, f'<w:ins w:id="{wid}" {PLACEHOLDER}/>')
            wid += 1
            out += new_para_tag(doc, rnd) + p_ppr + ins_block(wid, "", t) + "</w:p>"
        how = f"段落 {n} 之后新增 {k} 段"
    return doc[:pm.start()] + out + doc[pm.end():], how


def delete_para(doc, n):
    """整段删除（含段落标记）。"""
    wid = new_id(doc)
    paras = list(PARA.finditer(doc))
    if n >= len(paras):
        raise SystemExit(f"段落序号 {n} 超界（共 {len(paras)} 段）")
    pm = paras[n]
    head, ppr, body = split_para(pm.group(0))
    ppr = ppr or "<w:pPr></w:pPr>"
    has_ins, _, pending = pmark_flags(ppr)
    if has_ins and pending:
        raise SystemExit("该段是本次新增的，不要再叠一层删除修订——直接改插入内容")
    if has_ins:
        raise SystemExit("该段是来件既有的插入段（他人修订）。整段删除会叠改别人的痕迹，"
                         "先与经办人确认；需要表达意见可改用批注。")
    ppr = mark_pmark(ppr, f'<w:del w:id="{wid}" {PLACEHOLDER}/>')
    k, out = wid + 1, []
    for t in tokenize(body):
        if t.kind == "run" and t.text:
            out.append(del_block(k, [t]))
            k += 1
        else:
            out.append(t.raw)
    return doc[:pm.start()] + head + ppr + "".join(out) + "</w:p>" + doc[pm.end():]


def locate(doc, needle, occurrence):
    """跨 run 定位；返回 (段落序号, 段落 match, (head, ppr, tokens), span)。"""
    seen = 0
    for pi, pm in enumerate(PARA.finditer(doc)):
        head, ppr, body = split_para(pm.group(0))
        toks = tokenize(body)
        text, _ = index_map(toks)
        n = text.count(needle)
        if n and seen + n >= occurrence:
            span = find_span(toks, needle, occurrence - seen)
            return pi, pm, (head, ppr, toks), span
        seen += n
    raise SystemExit(
        f"[FAIL] 全文未找到「{needle}」的第 {occurrence} 处（共 {seen} 处）。"
        "注意原文里的空格、全角/半角；或先用 --grep 定位。")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("unpacked")
    ap.add_argument("--find")
    ap.add_argument("--replace")
    ap.add_argument("--insert-after")
    ap.add_argument("--occurrence", type=int, default=1)
    ap.add_argument("--after-para", type=int)
    ap.add_argument("--insert-paragraph")
    ap.add_argument("--delete-para", type=int)
    ap.add_argument("--grep")
    ap.add_argument("--list", nargs=2, type=int, metavar=("FROM", "TO"))
    a = ap.parse_args()

    unp = Path(a.unpacked)
    docp = unp / "word" / "document.xml"
    doc = docp.read_text(encoding="utf-8")

    if a.grep or a.list:
        for i, pm in enumerate(PARA.finditer(doc)):
            txt = "".join(T.findall(pm.group(0)))
            if (a.grep and a.grep in txt) or (a.list and a.list[0] <= i < a.list[1]):
                print(f"{i}\t{txt[:140]}")
        return

    wid = new_id(doc)

    if a.find is not None and (a.replace is not None or a.insert_after is not None):
        pi, pm, (head, ppr, toks), span = locate(doc, a.find, a.occurrence)
        i0, o0, i1, o1 = span
        pre, mid, post = slice_tokens(toks, i0, o0, i1, o1)
        rpr = toks[i0].rpr
        if a.insert_after is not None:
            middle = render(mid) + ins_block(wid, rpr, a.insert_after)
            what = f"在「{a.find}」后插入「{a.insert_after}」"
        else:
            middle = del_block(wid, mid) + ins_block(wid + 1, rpr, a.replace)
            what = (f"删除「{a.find}」" if a.replace == "" else f"「{a.find}」→「{a.replace}」")
        body_new = (render(toks[:i0]) + render(pre) + middle + render(post)
                    + render(toks[i1 + 1:]))
        doc = doc[:pm.start()] + head + ppr + body_new + "</w:p>" + doc[pm.end():]
        print(f"[OK] 段落 {pi}：{what}")

    elif a.after_para is not None and a.insert_paragraph:
        import random
        rnd = random.Random(hash((a.after_para, a.insert_paragraph)) & 0xFFFFFFFF)
        doc, how = insert_paras(doc, a.after_para, [a.insert_paragraph], rnd)
        print(f"[OK] {how}（{len(a.insert_paragraph)} 字）")

    elif a.delete_para is not None:
        doc = delete_para(doc, a.delete_para)
        print(f"[OK] 段落 {a.delete_para} 整段删除（含段落标记）")
    else:
        raise SystemExit("[FAIL] 参数组合无效，见文件头用法说明")

    docp.write_text(dedupe_change_ids(doc), encoding="utf-8")


if __name__ == "__main__":
    main()
