#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""落盘台 —— 本地服务。

只用标准库：双击「启动.command」即可，不装任何东西。
界面在浏览器里，服务只跑在 127.0.0.1，不对外开放，文件不出本机。

它做的事：把 AI 写的《落盘块》文字，落到来件 Word 上，成为真正的修订与批注，
署名和时间线按配置里的口径写，交付前跑完 C0–C7 体检，不过关不出文件。
"""
import base64
import http.server
import json
import os
import platform
import re
import shutil
import socketserver
import subprocess
import sys
import tempfile
import threading
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
try:
    from config import AUTHOR as CFG_AUTHOR, INITIALS as CFG_INITIALS, \
        WINDOWS as CFG_WINDOWS, WEEKEND as CFG_WEEKEND, SUFFIX as CFG_SUFFIX
except Exception:                                        # 拿不到配置也要能跑
    CFG_AUTHOR, CFG_INITIALS = "审查人", "审"
    CFG_WINDOWS, CFG_WEEKEND, CFG_SUFFIX = "09:00-18:00", False, "（修订稿）"
ENGINE = HERE / "引擎"
CORE = HERE / "核心"
# 中间产物放系统临时目录，不往用户的文件夹里堆垃圾；日志留在身边，随时可查
WORKROOT = Path(tempfile.gettempdir()) / "落盘台工作区"
LOGDIR = HERE / "日志"
sys.path.insert(0, str(CORE))
sys.path.insert(0, str(ENGINE))

from locator import Doc                      # noqa: E402
import parser as blockparser                 # noqa: E402
import planner                               # noqa: E402

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


WORKROOT.mkdir(parents=True, exist_ok=True)
LOGDIR.mkdir(exist_ok=True)

# LibreOffice（.doc 来件要用它进出格式）在 macOS 上不在 PATH 里
for cand in ("/Applications/LibreOffice.app/Contents/MacOS",
             "/opt/homebrew/bin", "/usr/local/bin"):
    if Path(cand, "soffice").exists() and cand not in os.environ.get("PATH", ""):
        os.environ["PATH"] = cand + os.pathsep + os.environ.get("PATH", "")

HAS_SOFFICE = bool(shutil.which("soffice"))

S = {"src": None, "work": None, "doc": None, "items": [], "results": [],
     "forces": {}, "skips": [], "warns": [], "text": ""}
JOB = {"lines": [], "running": False, "done": False, "ok": None, "deliver": "",
       "title": ""}
LOCK = threading.Lock()
LOGF = LOGDIR / f"{datetime.now():%Y%m%d-%H%M%S}.log"


def log(line=""):
    with LOCK:
        JOB["lines"].append(line)
    try:
        with LOGF.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def friendly(exc: BaseException) -> str:
    t = f"{type(exc).__name__}: {exc}"
    if "soffice" in t or "FileNotFoundError" in t and "soffice" in t:
        return ("没找到 LibreOffice——.doc 这种旧格式要靠它进出。"
                "请装一个 LibreOffice，或者先把来件在 Word 里另存为 .docx 再来。")
    if "BadZipFile" in t:
        return "这个文件打不开：后缀虽然是 Word，里面却不是。请在 Word 里打开另存一次再试。"
    if "PermissionError" in t:
        return "没有权限读写这个位置。把文件挪到桌面或文稿里再试。"
    if "FileNotFoundError" in t:
        return "找不到这个文件，路径可能改过了。重新选一次来件。"
    return t


# ── 业务 ─────────────────────────────────────────────────────────────────
def do_prep(src: Path) -> Path:
    work = WORKROOT / f"{datetime.now():%Y%m%d-%H%M%S}-{src.stem[:20]}"
    if work.exists():
        shutil.rmtree(work, ignore_errors=True)
    keep_recent(WORKROOT, 20)
    r = subprocess.run([sys.executable, str(ENGINE / "prep.py"), str(src),
                        "--work", str(work)], capture_output=True, text=True)
    for ln in (r.stdout + r.stderr).splitlines():
        log("  " + ln)
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout).strip()[-400:] or "解包失败")
    return work


def keep_recent(root: Path, n: int):
    """只留最近 n 次的中间产物，旧的清掉。"""
    try:
        dirs = sorted([d for d in root.iterdir() if d.is_dir() and d.name != "上传"],
                      key=lambda d: d.name)
        for d in dirs[:-n]:
            shutil.rmtree(d, ignore_errors=True)
    except OSError:
        pass


def load_doc(work: Path) -> Doc:
    return Doc((work / "unpacked" / "word" / "document.xml").read_text(encoding="utf-8"))


def resolve_now():
    S["results"] = planner.resolve_all(S["doc"], S["items"], S["forces"], S["skips"])
    return S["results"]


def view(results):
    raw = {str(it.get("uid") or it["no"]): it for it in S["items"]}
    out = []
    for r in results:
        p = r.get("preview")
        it = raw.get(str(r["uid"]), {})
        out.append({
            "uid": r["uid"],
            "item": {k: it.get(k, "") for k in
                     ("kw", "kind", "loc", "src", "dst", "note", "anchor", "body", "why", "occ")},
            "no": r["no"], "kind": r["kind"], "label": r["label"], "loc": r["loc"],
            "summary": r["summary"], "status": r["status"], "msg": r["msg"],
            "para": r["para"], "skip": bool(r["skip"]), "why": r["why"],
            "n": len(r["entries"]),
            "preview": p and {"text": p["text"][:400], "s": p["s"], "e": p["e"],
                              "dst": p["dst"][:300], "mode": p["mode"]},
            "cands": [{"para": c["para"], "score": c["score"], "start": c["start"],
                       "end": c["end"], "text": c["text"][:160], "hit": c["hit"][:80]}
                      for c in (r["cands"] or [])[:6]],
        })
    return out


def run_pipeline(cfg):
    """重新解包 → 重算清单 → 落盘 → 署名 → 打包体检 →（.doc 则回转）。"""
    JOB.update(lines=[], running=True, done=False, ok=None, deliver="", title=cfg["name"])
    try:
        src = Path(S["src"])
        log(f"【落盘台】{datetime.now():%Y-%m-%d %H:%M:%S}")
        log(f"来件：{src}")
        log(f"署名：{cfg['author']}（{cfg['initials']}）　时间线：{cfg['start']}"
            + (f" → {cfg['end']}" if cfg.get("end") else "") + f"　工作时段：{cfg['windows']}")
        log("")
        log("== 1/4 解包并记录基线 ==")
        work = do_prep(src)
        S["work"] = str(work)
        S["doc"] = load_doc(work)
        results = resolve_now()
        st = planner.stats(results)
        log(f"  清单 {len(results)} 条：可落盘 {st['ok'] + st['check']}，"
            f"需确认 {st['ask']}，找不到 {st['miss']}，格式有问题 {st['bad']}，跳过 {st['skip']}")
        for r in results:
            if r["status"] in ("miss", "ask", "bad") and not r["skip"]:
                log(f"  [未落盘] {r['label']}{r['no']} {r['loc']}：{r['msg']}")
            elif r["skip"]:
                log(f"  [已跳过] {r['label']}{r['no']} {r['loc']}：{r['msg'] or '手动跳过'}")
        plan = planner.build_plan(results)
        if not plan["edits"]:
            raise RuntimeError("没有一条能落盘的条目，先把上面的问题处理掉。")
        (work / "plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=1),
                                        encoding="utf-8")
        log("")
        log("== 2/4 批量落盘修订与批注" + ("（试跑，不写文件）" if cfg["dry"] else "") + " ==")
        cmd = [sys.executable, str(ENGINE / "apply.py"), str(work / "unpacked"),
               "--plan", str(work / "plan.json"), "--author", cfg["author"],
               "--initials", cfg["initials"]] + (["--dry-run"] if cfg["dry"] else [])
        code = stream(cmd)
        if code != 0:
            raise RuntimeError("有条目没落上，看上面的 FAIL 行。")
        if cfg["dry"]:
            log("")
            log("试跑通过：清单全部能锚定。去掉试跑，正式落盘。")
            JOB.update(ok=True)
            return
        log("")
        log("== 3/4 统一署名与北京时间戳 ==")
        cmd = [sys.executable, str(ENGINE / "stamp.py"), str(work / "unpacked"),
               "--author", cfg["author"], "--initials", cfg["initials"],
               "--start", cfg["start"], "--windows", cfg["windows"], "--seed", "7"]
        if cfg.get("end"):
            cmd += ["--end", cfg["end"]]
        if cfg.get("weekend"):
            cmd.append("--weekend")
        if stream(cmd) != 0:
            raise RuntimeError("署名与时间戳这一步失败了。")
        log("")
        log("== 4/4 打包 + 体检" + ("（+ 回转 .doc）" if src.suffix.lower() == ".doc" else "") + " ==")
        cmd = [sys.executable, str(ENGINE / "finish.py"), "--work", str(work),
               "--name", cfg["name"], "--author", cfg["author"], "--windows", cfg["windows"]]
        if cfg.get("weekend"):
            cmd.append("--weekend")
        code = stream(cmd)
        if code != 0:
            raise RuntimeError("体检没过，按规矩不出文件。上面的 FAIL 行是原因。")
        ext = src.suffix
        made = work / "deliver" / (cfg["name"] + ext)
        if not made.exists():
            made = work / "deliver" / (cfg["name"] + ".docx")
        outdir = Path(cfg["outdir"]).expanduser() if cfg.get("outdir") else src.parent
        outdir.mkdir(parents=True, exist_ok=True)
        final = outdir / made.name
        n = 1
        while final.exists() and final.resolve() != made.resolve():
            final = outdir / f"{made.stem}({n}){made.suffix}"
            n += 1
        shutil.copy2(made, final)
        spare = work / "deliver" / (cfg["name"] + ".docx")
        if ext.lower() == ".doc" and spare.exists():
            shutil.copy2(spare, outdir / spare.name)
            log(f"[备查] {outdir / spare.name}")
        JOB["deliver"] = str(final)
        log("")
        log(f"【交付】{final}")
        JOB.update(ok=True)
    except Exception as e:                                  # noqa: BLE001
        log("")
        log("【停下了】" + friendly(e))
        log(traceback.format_exc().strip().splitlines()[-1])
        JOB.update(ok=False)
    finally:
        JOB.update(running=False, done=True)
        log(f"（日志：{LOGF}）")


def stream(cmd):
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                         text=True, bufsize=1)
    for line in p.stdout:
        log(line.rstrip())
    p.wait()
    return p.returncode


# ── HTTP ─────────────────────────────────────────────────────────────────
class H(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            data = (HERE / "ui.html").read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if self.path.startswith("/api/log"):
            since = 0
            m = re.search(r"since=(\d+)", self.path)
            if m:
                since = int(m.group(1))
            with LOCK:
                lines = JOB["lines"][since:]
                n = len(JOB["lines"])
            return self._send({"lines": lines, "n": n, "running": JOB["running"],
                               "done": JOB["done"], "ok": JOB["ok"],
                               "deliver": JOB["deliver"]})
        if self.path == "/api/state":
            return self._send({"src": S["src"], "soffice": HAS_SOFFICE,
                               "mac": platform.system() == "Darwin",
                               "items": len(S["items"]), "log": str(LOGF),
                               "cfg": {"author": CFG_AUTHOR, "initials": CFG_INITIALS,
                                       "windows": CFG_WINDOWS, "weekend": CFG_WEEKEND,
                                       "suffix": CFG_SUFFIX}})
        self.send_error(404)

    def do_POST(self):
        n = int(self.headers.get("Content-Length") or 0)
        try:
            body = json.loads(self.rfile.read(n) or b"{}")
        except Exception:                                   # noqa: BLE001
            body = {}
        try:
            self._send(self.route(self.path, body))
        except Exception as e:                              # noqa: BLE001
            traceback.print_exc()
            self._send({"error": friendly(e)}, 200)

    # ── 路由 ─────────────────────────────────────────────────────────────
    def route(self, path, b):
        if path == "/api/pick":
            return api_pick()
        if path == "/api/open":
            return api_open(Path(b["path"]).expanduser())
        if path == "/api/upload":
            up = WORKROOT / "上传"
            up.mkdir(exist_ok=True)
            f = up / b["name"]
            f.write_bytes(base64.b64decode(b["b64"].split(",")[-1]))
            return api_open(f)
        if path == "/api/parse":
            S["text"] = b.get("text", "")
            S["forces"], S["skips"] = {}, []
            items, warns = blockparser.parse(S["text"])
            S["items"], S["warns"] = items, warns
            if not S["doc"]:
                return {"error": "先选来件，再解析——定位要对着来件做。"}
            return {"warns": warns, "items": view(resolve_now()),
                    "stats": planner.stats(S["results"])}
        if path == "/api/force":
            S["forces"][key(b)] = {"para": int(b["para"]), "start": int(b["start"]),
                                   "end": int(b["end"])}
            return {"items": view(resolve_now()), "stats": planner.stats(S["results"])}
        if path == "/api/unforce":
            S["forces"].pop(key(b), None)
            return {"items": view(resolve_now()), "stats": planner.stats(S["results"])}
        if path == "/api/skip":
            uid = key(b)
            if b.get("on"):
                if uid not in S["skips"]:
                    S["skips"].append(uid)
            elif uid in S["skips"]:
                S["skips"].remove(uid)
            return {"items": view(resolve_now()), "stats": planner.stats(S["results"])}
        if path == "/api/edit":
            # 改口径：AI 写的是修订、你看了想改成批注；或者原文抄漏了要就地补。
            uid = key(b)
            it = next((x for x in S["items"] if str(x.get("uid") or x["no"]) == uid), None)
            if not it:
                return {"error": "没有这一条了，重新解析一次落盘块。"}
            kind = (b.get("kind") or it["kind"]).strip()
            if kind not in ("edit", "del", "delpara", "insert", "note"):
                return {"error": "认不出这个类型。"}
            it["kind"] = kind
            it["kw"] = blockparser.KIND_LABEL.get(kind, it.get("kw", ""))
            for k in ("loc", "src", "dst", "note", "anchor", "body", "why"):
                if k in b:
                    it[k] = (b.get(k) or "").strip()
            it["occ"] = str(b.get("occ") or it.get("occ") or "")
            it["error"] = ""
            blockparser.finalize(it)
            S["forces"].pop(uid, None)          # 改了内容，旧的指认作废
            r = {"items": view(resolve_now()), "stats": planner.stats(S["results"])}
            if it["error"]:
                r["warn"] = it["error"]
            return r
        if path == "/api/export":
            return {"text": blockparser.dump(S["items"])}
        if path == "/api/paras":
            q = (b.get("q") or "").strip()
            doc = S["doc"]
            if not doc:
                return {"paras": []}
            out = []
            for i, t in enumerate(doc.paras):
                if not t.strip():
                    continue
                if q and q not in t:
                    continue
                out.append({"i": i, "text": t[:200]})
                if len(out) >= 60:
                    break
            return {"paras": out}
        if path == "/api/run":
            if JOB["running"]:
                return {"error": "上一趟还没跑完。"}
            if not S["src"]:
                return {"error": "还没选来件。"}
            if not S["items"]:
                return {"error": "还没解析落盘块。"}
            cfg = {
                "dry": bool(b.get("dry")),
                "name": (b.get("name") or "").strip() or (Path(S["src"]).stem + "（修订稿）"),
                "author": (b.get("author") or "").strip() or _CFG_AUTHOR,
                "initials": (b.get("initials") or "").strip() or _CFG_INITIALS,
                "start": (b.get("start") or "").strip() or datetime.now().strftime("%Y-%m-%dT09:20"),
                "end": (b.get("end") or "").strip(),
                "windows": (b.get("windows") or "").strip() or _CFG_WINDOWS,
                "weekend": bool(b.get("weekend")),
                "outdir": (b.get("outdir") or "").strip(),
            }
            threading.Thread(target=run_pipeline, args=(cfg,), daemon=True).start()
            return {"ok": True}
        if path == "/api/reveal":
            p = b.get("path") or JOB.get("deliver")
            if not p:
                return {"error": "还没有可交付的文件。"}
            if platform.system() == "Darwin":
                subprocess.run(["open", "-R", p])
            elif platform.system() == "Windows":
                subprocess.run(["explorer", "/select,", p])
            else:
                subprocess.run(["xdg-open", str(Path(p).parent)])
            return {"ok": True}
        if path == "/api/quit":
            threading.Timer(0.4, lambda: os._exit(0)).start()
            return {"ok": True}
        return {"error": "没有这个接口"}


def key(b):
    return str(b.get("uid") or b.get("no") or "")


def api_pick():
    if platform.system() != "Darwin":
        return {"error": "这台机器不是 macOS，请直接把文件拖进来或粘贴路径。"}
    scr = ('POSIX path of (choose file with prompt "选择来件（.doc / .docx）" '
           'of type {"org.openxmlformats.wordprocessingml.document","com.microsoft.word.doc"})')
    r = subprocess.run(["osascript", "-e", scr], capture_output=True, text=True)
    if r.returncode != 0:
        return {"error": "没选文件。"}
    return api_open(Path(r.stdout.strip()))


def api_open(src: Path):
    if not src.exists():
        return {"error": "这个文件不在：" + str(src)}
    if src.suffix.lower() not in (".doc", ".docx"):
        return {"error": "只认 .doc 和 .docx。"}
    if src.suffix.lower() == ".doc" and not HAS_SOFFICE:
        return {"error": "来件是 .doc，本机没装 LibreOffice，转不了格式。"
                         "请先在 Word 里另存为 .docx，或装一个 LibreOffice。"}
    JOB.update(lines=[], done=False, ok=None, deliver="")
    log(f"打开来件：{src}")
    work = do_prep(src)
    S.update(src=str(src), work=str(work), doc=load_doc(work), forces={}, skips=[])
    doc = S["doc"]
    body = [t for t in doc.paras if t.strip()]
    out = {"path": str(src), "name": src.name, "stem": src.stem, "ext": src.suffix,
           "dir": str(src.parent), "paras": len(doc.paras), "words": sum(len(t) for t in body),
           "head": body[:3]}
    if S["items"]:
        out["items"] = view(resolve_now())
        out["stats"] = planner.stats(S["results"])
    return out


class Server(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


def main():
    port = int(os.environ.get("LUOPAN_PORT", "0")) or 0
    for p in ([port] if port else range(8730, 8760)):
        try:
            httpd = Server(("127.0.0.1", p), H)
            break
        except OSError:
            continue
    else:
        print("端口都被占着，换一个：LUOPAN_PORT=8790 python3 server.py")
        return
    url = f"http://127.0.0.1:{httpd.server_address[1]}/"
    print("落盘台已启动：" + url)
    print("（这个终端窗口关掉，服务就停了。）")
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")


if __name__ == "__main__":
    main()
