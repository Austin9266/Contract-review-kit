#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""finish.py —— 打包 → 体检 → （来件是 .doc 时）回转 .doc 并复检。

用法:
    python3 scripts/finish.py --work work/ --name "XX合同（修订稿）" \\
        --author "某某律所 张三" [--windows "09:00-18:00"] [--weekend]

产物:
    work/deliver/<name><来件后缀>     ← 交给客户的文件，后缀与来件一致
    work/deliver/<name>.docx         ← 来件是 .doc 时同时留一份 docx 备查
"""
import argparse
import re
import shutil
import subprocess
import sys
import zipfile
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from wtext import document_text  # noqa: E402

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


HERE = Path(__file__).resolve().parent


def rezip(unp: Path, out: Path, base_entries=None):
    """按来件的条目顺序重建 zip：
    - [Content_Types].xml 必须在包首（Word 的硬要求）；
    - 来件里的**目录条目**（如 'word/'，WPS 生成的包都带）原样保留 ——
      丢了目录条目，C1 会把它们判成"部件缺失"；
    - 来件没有而本次新增的部件（people.xml、commentsExtensible.xml 等）排在最后。"""
    if out.exists():
        out.unlink()
    files = {str(p.relative_to(unp)).replace("\\", "/"): p
             for p in unp.rglob("*") if p.is_file()}
    order = list(base_entries) if base_entries else []
    if "[Content_Types].xml" in order:
        order.remove("[Content_Types].xml")
    order = ["[Content_Types].xml"] + order
    extra = [n for n in sorted(files) if n not in order]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order + extra:
            if name.endswith("/"):
                zi = zipfile.ZipInfo(name)
                zi.external_attr = (0o40775 << 16) | 0x10   # 目录条目
                z.writestr(zi, b"")
            elif name in files:
                z.write(files[name], name)
            # 来件有、解包目录里没有的文件条目：不该发生（本流程绝不删部件），
            # 真缺了由体检 C1 拦截，这里不做静默补偿


def soffice(src: Path, target: str, outdir: Path) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", target,
                    "--outdir", str(outdir), str(src)],
                   check=True, capture_output=True, text=True)
    out = outdir / (src.stem + "." + target.split(":")[0])
    if not out.exists():
        raise SystemExit(f"[FAIL] soffice 未生成 {out}")
    return out


def count(xml):
    return (len(re.findall(r"<w:ins\s", xml)), len(re.findall(r"<w:del\s", xml)),
            len(re.findall(r"<w:commentReference\s", xml)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--work", required=True)
    ap.add_argument("--name", required=True, help="交付文件名（不含后缀）")
    ap.add_argument("--author", default=_CFG_AUTHOR)
    ap.add_argument("--windows", default=_CFG_WINDOWS)
    ap.add_argument("--weekend", action="store_true")
    a = ap.parse_args()

    work = Path(a.work).resolve()
    base = json.loads((work / "baseline.json").read_text(encoding="utf-8"))
    ext = base["source_ext"]
    deliver = work / "deliver"
    deliver.mkdir(exist_ok=True)

    docx_out = deliver / (a.name + ".docx")
    rezip(work / "unpacked", docx_out, list(base["entries"].keys()))
    print(f"[OK] 已打包 {docx_out.name}（沿用来件条目顺序，目录条目保留）")

    cmd = [sys.executable, str(HERE / "verify.py"), "--work", str(work),
           "--out", str(docx_out), "--author", a.author, "--windows", a.windows,
           "--expect-ext", ext]
    if a.weekend:
        cmd.append("--weekend")
    r = subprocess.run(cmd)
    if r.returncode != 0:
        raise SystemExit("[FAIL] 严格体检未通过，已停在 docx 阶段，不生成交付件")

    if ext == ".docx":
        print(f"\n[交付] {docx_out}")
        return

    # 来件是 .doc：回转 .doc，并对回转结果做复检（LibreOffice 转换会有格式漂移风险）
    doc_out = soffice(docx_out, "doc", deliver)
    print(f"\n== 复检（.doc 回转件：{doc_out.name}）==")
    probe = soffice(doc_out, "docx", work / "_probe")
    with zipfile.ZipFile(probe) as z:
        pxml = z.read("word/document.xml").decode("utf-8")
        names = z.namelist()
        pcomments = z.read("word/comments.xml").decode("utf-8") if "word/comments.xml" in names else ""
    with zipfile.ZipFile(docx_out) as z:
        dxml = z.read("word/document.xml").decode("utf-8")
        dnames = z.namelist()
        dcomments = z.read("word/comments.xml").decode("utf-8") if "word/comments.xml" in dnames else ""

    good = True

    def chk(c, msg):
        nonlocal good
        print(("  [PASS] " if c else "  [FAIL] ") + msg)
        good = good and c

    ci, cd, cc = count(dxml)
    pi, pd, pc = count(pxml)
    def norm(t):
        return [l.strip() for l in t.split("\n") if l.strip()]
    probe_rej = document_text(pxml, "reject")
    baseline_txt = (work / "baseline.txt").read_text(encoding="utf-8")
    chk(norm(probe_rej) == norm(baseline_txt), "回转后拒绝修订的正文 == 来件正文")
    chk((pi > 0) == (ci > 0) and (pd > 0) == (cd > 0), f"修订标记存活（插入 {ci}→{pi}，删除 {cd}→{pd}）")
    chk(pc == cc, f"批注锚点存活（{cc}→{pc}）")
    pa = set(re.findall(r'w:author="([^"]*)"', pxml)) | set(re.findall(r'w:author="([^"]*)"', pcomments))
    allowed = {a.author} | set(base.get("pre_authors", []))
    chk(pa <= allowed or not pa, f"署名存活：{sorted(pa)}"
        + (f"（含来件既有作者 {sorted(set(base.get('pre_authors', [])))}）"
           if base.get("pre_authors") else ""))
    pdates = set(re.findall(r'w:date="([^"]*)"', pxml)) | set(re.findall(r'w:date="([^"]*)"', pcomments))
    # .doc 的 DTTM 只精确到分钟，且 LibreOffice 回读时会补一个 Z 后缀；
    # 但"钟面数字"没有被折算 —— Word 打开 .doc 显示的仍是北京时间。
    # 故只按"年月日时分"校验数字没被 ±8 小时改写。
    dwall = {d.rstrip("Z")[:16] for d in
             set(re.findall(r'w:date="([^"]*)"', dxml)) | set(re.findall(r'w:date="([^"]*)"', dcomments))}
    pwall = {d.rstrip("Z")[:16] for d in pdates}
    chk(pwall <= dwall, f"回转后钟面时间未被折算{('，异常: ' + str(sorted(pwall - dwall)[:3])) if pwall - dwall else ''}")

    print("\n[交付] " + str(doc_out) + "（后缀与来件一致）")
    print("[备查] " + str(docx_out))
    if not good:
        print("\n[!] .doc 回转复检有 FAIL：.doc 是二进制旧格式，Word 97 版式与修订元数据"
              "在转换中可能丢失。请按 SKILL.md 第 6 步处理——改为交付 docx 备查件，"
              "并在回复中向客户说明原因。")
        sys.exit(1)


if __name__ == "__main__":
    main()
