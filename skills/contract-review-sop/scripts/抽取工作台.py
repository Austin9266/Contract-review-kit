#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""抽取工作台.py —— 把《审查工作台》里的规则库与文本积木抽进本技能。

工作台是判断口径的源头（规则卡在那里改最顺手）。改完之后跑一次这个，
技能里的 rules/规则库.json 与 scripts/prompt_parts.py 就跟工作台对齐了：

    python3 抽取工作台.py ../../../工作台/审查工作台.html

抽完建议做一次逐字比对：同一组设置下，assemble.py 的输出应当与工作台页面上的
prompt 一模一样（只多一行「本次装配」表头）。
"""
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent


def unesc(t):
    return t.replace("\\\\", "\\").replace("\\`", "`").replace("\\$", "$")


SUB = [
    ('${mcpBlock()}', '«mcp»'), ('${SEQ}', '«seq»'), ('${TONE}', '«tone»'),
    ('${RESTRAINT}', '«restraint»'), ('${trackTxt}', '«track_txt»'),
    ('${packs.map(SHORT).join("、")}', '«pack_names»'), ('${cs.length}', '«n_cards»'),
    ('${ruleBlock()}', '«rules»'), ('${st.role}', '«role»'), ('${t}', '«type_name»'),
    ('${st.track === "A" ? "我方制式模板" : "对方模板／双方自拟"}', '«source_txt»'),
    ('${st.situ.size ? [...st.situ].map(SHORT).join("、") : "无"}', '«situ_names»'),
    ('${st.note.trim() || "（未提供）"}', '«note»'),
    ('${(!st.law||!st.tyc)?"、以及本次因未接工具而留待核实的事项":""}', '«mcp_tail»'),
    ('${LAND()}', '«land»'), ('${kind}', '«doc_kind»'),
    ('${DOCPOINTS[st.doc] || DOCPOINTS.qt}', '«docpoints»'),
    ('${SIGN()}', '«sign»'), ('${TL()}', '«timeline»'), ('${FILE()}', '«file»'),
    ('${BLOCKSPEC}', '«blockspec»'),
    ('${st.org||"我方"}', '«org»'),
]


def slots(t):
    for a, b in SUB:
        t = t.replace(a, b)
    left = re.findall(r"\$\{.*?\}", t, re.S)
    if left:
        raise SystemExit("[FAIL] 工作台里出现了没见过的插值，先在 SUB 里补一条：%s" % left[:3])
    return t


def main():
    if len(sys.argv) < 2:
        raise SystemExit("用法：python3 抽取工作台.py <审查工作台.html>")
    src = Path(sys.argv[1]).expanduser()
    s = src.read_text(encoding="utf-8")

    lib = json.loads(re.search(r"const BUILTIN = (\{.*?\});\n", s, re.S).group(1))
    (ROOT / "rules" / "规则库.json").write_text(
        json.dumps(lib, ensure_ascii=False, indent=1), encoding="utf-8")

    def lit(n):
        return unesc(re.search(r"const %s = `(.*?)`;" % n, s, re.S).group(1))

    parts = {n: lit(n) for n in ("SEQ", "TONE", "RESTRAINT", "BLOCKSPEC")}
    parts["LAND"] = slots(unesc(re.search(
        r"const LAND = \(\) => st\.who === \"ai\" \? `(.*?)` : `", s, re.S).group(1)))

    def body(fn):
        i = s.index("function %s()" % fn)
        j = s.index("\n}", s.index("return `", i))
        return slots(unesc(s[s.index("return `", i) + len("return `"): s.rindex("`;", i, j)]))

    parts["CONTRACT"], parts["DOC"] = body("promptContract"), body("promptDoc")
    blk = s[s.index("const DOCPOINTS = {"): s.index("};", s.index("const DOCPOINTS = {"))]
    dp = {k: unesc(v) for k, v in re.findall(r"(\w+):`(.*?)`,?\n", blk, re.S)}

    def q(t):
        assert '"""' not in t
        return '"""' + t.replace("\\", "\\\\") + '"""'

    out = ["#!/usr/bin/env python3", "# -*- coding: utf-8 -*-",
           '"""prompt_parts.py —— 审查指令的文本积木。', "",
           "自动从《审查工作台》HTML 抽出，**勿手改**：要改判断口径就改工作台"
           "（或 rules/规则库.json），再跑一次 scripts/抽取工作台.py。",
           f"规则库版本 {lib.get('version')}（{lib.get('updated')}），"
           f"共 {len(lib['cards'])} 张卡。",
           "«…» 是插值槽，由 assemble.py 填。", '"""', ""]
    for k in ("SEQ", "TONE", "RESTRAINT", "BLOCKSPEC", "LAND", "CONTRACT", "DOC"):
        out += [f"{k} = {q(parts[k])}", ""]
    out += ["DOCPOINTS = {"] + [f"    {k!r}: {q(v)}," for k, v in dp.items()] + ["}", ""]
    (HERE / "prompt_parts.py").write_text("\n".join(out), encoding="utf-8")
    print(f"[OK] 规则库 {len(lib['cards'])} 张卡 → rules/规则库.json")
    print(f"[OK] 文本积木 {len(parts)} 块 + {len(dp)} 个文书要点 → scripts/prompt_parts.py")
    print("下一步：拿同一组设置比一比 assemble.py 与工作台页面上的 prompt 是否逐字相同。")


if __name__ == "__main__":
    main()
