#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""parser.py —— 把 AI 产出的「落盘块」纯文本解析成结构化条目。

设计前提：这段文字是聊天窗口里复制粘贴过来的，必然带各种脏东西——
代码围栏、markdown 列表符号、全角冒号、引号包裹、软换行、行首引用符。
所以解析器一律**宽进**：能猜出来的都认，猜不出来的逐条报错并指出行号，
绝不静默丢条目。

块格式（见《落盘块格式》说明）：

    <<<落盘>>>
    @修订 1 | 第7.2条
    原文：甲方应在收到发票后15日内付款
    改为：甲方应在收到发票且经审核无误后30日内付款
    理由：付款条件不明确
    @批注 2 | 第7.2条
    锚点：甲方应在收到发票后15日内付款
    批注：请确认审核时限。
    @新增 3 | 第9条之后
    锚点：本合同一式肆份
    正文：9.5 因不可抗力……
    @删段 4 | 第11.3条
    锚点：本条第一句原文
    <<<落盘完>>>
"""
import re

# ── 块边界 ────────────────────────────────────────────────────────────────
OPEN = re.compile(r"<{2,4}\s*落盘(?:开始)?\s*>{2,4}")
CLOSE = re.compile(r"<{2,4}\s*落盘(?:完|结束|完毕)\s*>{2,4}")

# ── 类型别名 ──────────────────────────────────────────────────────────────
KINDS = {
    "修订": "edit", "修改": "edit", "改": "edit", "replace": "edit",
    "删除": "del", "删": "del", "删字": "del", "delete": "del",
    "删段": "delpara", "删除整段": "delpara", "整段删除": "delpara",
    "新增": "insert", "增加": "insert", "补充": "insert", "新增条款": "insert",
    "插入": "insert", "insert": "insert",
    "批注": "note", "注": "note", "comment": "note",
}
# ── 字段别名 ──────────────────────────────────────────────────────────────
FIELDS = {
    "位置": "loc", "条款": "loc", "条款位置": "loc", "定位": "loc", "所在": "loc",
    "原文": "src", "原句": "src", "原条款": "src", "原表述": "src", "旧": "src",
    "改为": "dst", "修改为": "dst", "改成": "dst", "新文": "dst", "修订为": "dst",
    "替换为": "dst", "新": "dst",
    "批注": "note", "批注文字": "note", "批注内容": "note", "注": "note",
    "锚点": "anchor", "锚": "anchor", "挂在": "anchor", "定位串": "anchor",
    "正文": "body", "条款文本": "body", "新增文本": "body", "文本": "body",
    "内容": "body", "建议条款": "body",
    "理由": "why", "说明": "why", "依据": "why", "说理": "why", "法律依据": "why",
    "第几处": "occ", "出现次数": "occ", "第几个": "occ", "序次": "occ",
}
FIELD_LINE = re.compile(r"^\s*([一-龥A-Za-z]{1,6})\s*[:：]\s*(.*)$")
# 行首可能挂着 markdown 列表符或全角 ＠——AI 常干这事，认下来，别把整条吞掉
HEAD_LINE = re.compile(r"^\s*(?:[-*·•]\s*|\d+[.、)]\s*)?[@＠]\s*([一-龥A-Za-z]{1,6})\s*"
                       r"([0-9０-９]{0,3})\s*(?:[|｜/、,，－\-]\s*(.*))?$")
KIND_ANY = re.compile(r"[@＠]\s*(修订|修改|删除|删段|新增|增加|补充|插入|批注)")

FENCE = re.compile(r"^\s*(```|~~~).*$")
BULLET = re.compile(r"^\s*(?:[>＞]\s*)?(?:[-*·•]\s+|\d+[.、)]\s+)?")
EMPTY_WORDS = {"无", "（无）", "(无)", "空", "删除", "（删除）", "(删除)", "[删除]", "【删除】", "—", "－", "-", "/"}
QUOTE_PAIRS = [("「", "」"), ("『", "』"), ("“", "”"), ("‘", "’"), ('"', '"'), ("'", "'"), ("《", "》")]


def _dequote(s: str) -> str:
    s = s.strip()
    changed = True
    while changed and len(s) >= 2:
        changed = False
        for a, b in QUOTE_PAIRS:
            if s.startswith(a) and s.endswith(b) and len(s) > len(a) + len(b) - 1:
                inner = s[len(a):-len(b)]
                # 只有成对包住整串才剥；串内还有同款引号的（如引用了带书名号的法规）不动
                if a not in inner and b not in inner:
                    s, changed = inner.strip(), True
                    break
    return s


def _clean_lines(text: str):
    """去代码围栏、行首引用符与列表符；保留行号（1 起，对应原文）。"""
    out = []
    for i, raw in enumerate(text.splitlines(), 1):
        if FENCE.match(raw):
            continue
        line = raw.replace("　", " ").rstrip()
        if line.strip().startswith(("＞", ">")):
            line = re.sub(r"^\s*[>＞]\s?", "", line)
        out.append((i, line))
    return out


def extract_block(text: str):
    """取出落盘块。没有标记时退回整段文本，并给一条提示。"""
    warn = []
    mo = OPEN.search(text)
    mc = CLOSE.search(text, mo.end() if mo else 0)
    if mo and mc:
        return text[mo.end():mc.start()], warn
    if mo:
        warn.append("只找到 <<<落盘>>> 开始标记、没找到结束标记，按到文末处理。")
        return text[mo.end():], warn
    if "@" not in text:
        return "", ["没找到落盘块，也没找到任何以 @ 开头的条目。请把审查报告里 <<<落盘>>> 那一整段贴进来。"]
    warn.append("没找到 <<<落盘>>> 标记，按整段文字里的 @ 条目解析。")
    return text, warn


def finalize(it):
    """清洗字段、补默认值、做必填校验。解析与界面改写共用同一套规矩。"""
    for k in ("loc", "src", "dst", "note", "anchor", "body", "why"):
        it[k] = it[k].strip()
    it["src"] = _dequote(it["src"])
    it["anchor"] = _dequote(it["anchor"])
    it["note"] = _dequote(it["note"])
    if it["dst"].strip() in EMPTY_WORDS:
        it["dst"] = ""
    else:
        it["dst"] = _dequote(it["dst"])
    m = re.search(r"\d+", it["occ"] or "")
    it["occ"] = int(m.group(0)) if m else 0
    if it["error"]:
        return it
    k = it["kind"]
    if k in ("edit", "del"):
        if not it["src"]:
            it["error"] = f"第 {it['line']} 行：@{it['kw']} 缺「原文」。"
        elif k == "del":
            it["dst"] = ""
        elif it["src"] == it["dst"]:
            it["error"] = f"第 {it['line']} 行：原文与改为一模一样，没有改动。"
        elif not it["dst"]:
            it["kind"] = "del"       # 改为留空 = 纯删除
    elif k == "note":
        if not it["note"]:
            it["error"] = f"第 {it['line']} 行：@批注 缺「批注」文字。"
        if not it["anchor"]:
            it["anchor"] = it["src"]         # 有人把锚点写在「原文」里
        if not it["anchor"] and not it["loc"]:
            it["error"] = f"第 {it['line']} 行：@批注 既没有「锚点」也没有「位置」，挂不上。"
    elif k == "insert":
        if not it["body"]:
            it["body"] = it["dst"]
        if not it["body"]:
            it["error"] = f"第 {it['line']} 行：@新增 缺「正文」。"
        if not it["anchor"]:
            it["anchor"] = it["src"]
        if not it["anchor"]:
            it["error"] = f"第 {it['line']} 行：@新增 缺「锚点」（要插在哪一段之后）。"
    elif k == "delpara":
        if not it["anchor"]:
            it["anchor"] = it["src"]
        if not it["anchor"]:
            it["error"] = f"第 {it['line']} 行：@删段 缺「锚点」（这一段里的任一句原文）。"
    return it


def parse(text: str):
    """→ (items, warnings)。items 是逐条 dict，errors 挂在条目的 'error' 上。"""
    block, warns = extract_block(text or "")
    lines = _clean_lines(block)

    items, cur, field = [], None, None

    def flush():
        if cur is not None:
            items.append(cur)

    for ln, line in lines:
        if not line.strip():
            field = None
            continue
        mh = HEAD_LINE.match(line)
        if mh and (mh.group(1) in KINDS or KIND_ANY.match(line.lstrip(" -*·•"))):
            flush()
            kw = mh.group(1)
            kind = KINDS.get(kw)
            num = mh.group(2) or ""
            num = num.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
            cur = {"line": ln, "kw": kw, "kind": kind, "no": num or str(len(items) + 1),
                   "uid": str(len(items) + 1),
                   "loc": (mh.group(3) or "").strip(), "src": "", "dst": "", "note": "",
                   "anchor": "", "body": "", "why": "", "occ": "", "error": ""}
            if kind is None:
                cur["error"] = f"第 {ln} 行：认不出条目类型「@{kw}」，只认 修订／删除／删段／新增／批注。"
            field = None
            continue
        if cur is None:
            continue
        mf = FIELD_LINE.match(line)
        key = FIELDS.get(mf.group(1)) if mf else None
        if key:
            field = key
            val = mf.group(2)
            cur[key] = (cur[key] + "\n" + val).strip("\n") if cur[key] else val
        elif field:                      # 续行：AI 把长条款换了行
            cur[field] = cur[field] + "\n" + BULLET.sub("", line, count=1)
        # 其余散行（AI 的旁白）直接忽略
    flush()

    # 有 @条目 的行没被认成条目 → 明说，绝不静默丢
    seen = {it["line"] for it in items}
    for ln, line in lines:
        if KIND_ANY.search(line) and ln not in seen and not FIELD_LINE.match(line):
            warns.append(f"第 {ln} 行像是一条条目却没认出来，已忽略：{line.strip()[:40]}")

    for it in items:
        finalize(it)
    return items, warns


KIND_LABEL = {"edit": "修订", "del": "删除", "delpara": "删段", "insert": "新增", "note": "批注"}
DUMP_FIELDS = {
    "edit":    [("src", "原文"), ("dst", "改为"), ("occ", "第几处"), ("why", "理由")],
    "del":     [("src", "原文"), ("occ", "第几处"), ("why", "理由")],
    "delpara": [("anchor", "锚点"), ("why", "理由")],
    "insert":  [("anchor", "锚点"), ("body", "正文"), ("why", "理由")],
    "note":    [("anchor", "锚点"), ("note", "批注"), ("occ", "第几处"), ("why", "理由")],
}


def dump(items) -> str:
    """条目 → 落盘块文字（界面上改过口径之后，用它导出留档或贴回报告）。"""
    out = ["<<<落盘>>>"]
    for it in items:
        head = f'@{KIND_LABEL.get(it["kind"], it.get("kw", "?"))} {it["no"]}'
        if it.get("loc"):
            head += f' | {it["loc"]}'
        out.append(head)
        for k, label in DUMP_FIELDS.get(it["kind"], []):
            v = it.get(k)
            if k == "occ":
                v = "" if not v or str(v) in ("0", "") else str(v)
            if not v:
                continue
            out.append(f"{label}：{v}")
    out.append("<<<落盘完>>>")
    return "\n".join(out)


def summary(it):
    k = it["kind"]
    if k == "edit":
        return f'「{it["src"][:22]}」→「{it["dst"][:22]}」'
    if k == "del":
        return f'删去「{it["src"][:30]}」'
    if k == "delpara":
        return f'整段删除（锚「{it["anchor"][:20]}」）'
    if k == "insert":
        return f'插在「{it["anchor"][:16]}」之后：{it["body"][:26]}'
    if k == "note":
        return f'{it["note"][:30]}'
    return "?"


if __name__ == "__main__":
    import json
    import sys
    its, ws = parse(sys.stdin.read())
    print(json.dumps({"items": its, "warns": ws}, ensure_ascii=False, indent=1))
