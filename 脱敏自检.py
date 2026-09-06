#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# <<<SELF-EXEMPT  以下到 SELF-EXEMPT>>> 之间是本脚本自己的说明与词表——
# 它天然要写着"要查的那些词长什么样"，所以自查时把这一段挖掉再扫。
# 除此之外本脚本的每一行（含注释）都参与自查：检测器把自己排除在外，
# 就是这类工具最典型的坏法。
"""脱敏自检.py —— 把工具包发给别人、或开源之前，先跑这个。

    python3 脱敏自检.py            # 查整个工具包
    python3 脱敏自检.py 某目录      # 查指定目录
    python3 脱敏自检.py --quiet     # 只报问题，不打印统计

它不靠"我记得改过哪些词"，而是按**几类特征**去找——这类东西最容易漏在
预填值、示例、注释里，肉眼过一遍必漏：

  L1 真实金额       千分位（1,234,567.89）、3 位以上小数，**以及中文数字写法**
                    （叁拾万元、贰佰万）—— 合同里金额本来就常写成中文
  L2 预填/默认值里的私人内容  HTML 的 value=/placeholder=，**以及代码里的兜底常量**
                    （`or "某"`、写死的作息时段）——页面一打开就顶在脸上，
                    或者用户一留空就落进交付件
  L3 具体日期       写死的年月日；示例日期可接受，但预填值里的多半是真实收件时刻
  L4 机构与人名     常见机构后缀 + 已知要清掉的词表 + 从词表派生的**单姓**
  L5 行业专属词     这套工具包声明是通用的，出现行业黑话说明没洗干净
  L6 私人路径       /Users/xxx、C:\\Users\\xxx —— 会暴露你的用户名与目录结构
  L7 运行产物       日志、工作区、__pycache__、.DS_Store、.bak
  L8 像逐字抄来的条款  引号内既长又带具体数字的片段——多半是从真实来件抄下来的原句。
                    示例要保留「长什么样」，但数字一律换成 X/N/【】。
                    中文直角引号「」也算；**所有文本文件都查**，不只 .json/.py
  L9 地名           指向具体乡镇/区县/园区的名字。通用工具包里不该出现任何真实地名，
                    这一类几乎没有正当理由留着

返回码：有 L1/L2/L4/L6/L7/L8/L9 → 1（必须处理）；只有 L3/L5 提示 → 0。

L8 与 L9 靠特征猜，会误伤；确认没问题的整行加进 脱敏白名单.txt 放行。
"""
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

SKIP_DIRS = {".git", "node_modules"}
BIN_SUFFIX = {".docx", ".doc", ".zip", ".dotx", ".xlsx", ".pptx"}

# —— L4 已知要清掉的词 ——
# 你的机构名、同事名、常合作的客户名写进 脱敏词表.local.txt（一行一条，# 开头是注释）。
# **不要写在这个文件里**：这个文件是要随包分发的，把真名写进来等于把刚洗掉的东西
# 又发出去一遍——而且它在自查豁免区内，脚本永远查不到自己这一行。
# .gitignore 已经挡住 *.local.txt，那份词表只留在你自己机器上。
# 写成中文全名（两到四个字）的，脚本会自动派生单姓一起盯——真实泄露常常只剩一个姓。
NAMES_FILE = "脱敏词表.local.txt"
NAMES = []                      # 兜底为空；真正的词表从 NAMES_FILE 读
# 机构口吻：任何机构都会有的说法，与你是哪一行、哪一家无关，所以固定带着。
# 反映师承关系或内部班组称呼的说法，请自己往 NAMES 里加——把它们写死在这里，
# 等于告诉读者作者的稿子里出现过这些词（连举例都别写在注释里，同理）。
TONE_WORDS = ["所内", "所里", "本所", "我们所", "部门内"]
ORG_SUFFIX = r"(?:律师事务所|律所|有限公司|股份有限公司|集团有限公司|管理委员会|管委会)"

# —— L5 行业专属词（这套包声明通用，出现即需要洗）——
# 出厂给的是一张**跨行业**示例表：留下你自己那一行，其余按需删改。
# ⚠️ 别把它改成只剩一个行业的词——那样这张表本身就等于宣布了你是做哪一行的，
#    正文洗得再干净也白搭（这是本包自己踩过的坑）。
INDUSTRY = ["临床数据", "适应症", "药品批件", "注册证",          # 医药医疗
            "埋点", "SDK", "用户画像", "接口调用量",                # 互联网
            "提单", "信用证", "报关单", "保税",                     # 外贸物流
            "分账", "署名权", "信息网络传播权", "母带",             # 影视文化
            "并网", "装机", "上网电价", "绿证",                     # 能源电力
            "容积率", "预售许可", "网签", "五证",                   # 房地产
            ]

# —— L9 地名：通名（乡镇政府、地方政府这类）不算，只抓看起来像专名的 ——
# 专名里不会出现这些字，出现了多半是把动词/连词误切进来了（"甲方本身系乡镇政府"）
NOT_IN_PLACE = set("的是系为在与及或由本该此各和向对从并且到了不无有所属甲乙丙丁方等如按照根据前后另其两三四五于位处名叫称谓设含即将也还就都")
# SELF-EXEMPT>>>

RE_MONEY = re.compile(r"\b\d{1,3}(?:,\d{3})+(?:\.\d+)?\b|\b\d+\.\d{3,}\b")
# 中文数字金额：合同里的法定写法，阿拉伯数字的正则一个都抓不到
RE_CN_MONEY = re.compile(r"[一二三四五六七八九十百千零壹贰叁肆伍陆柒捌玖拾佰仟两]{1,8}(?:万元|万|亿元|亿)")
RE_DATE = re.compile(r"\b20\d{2}[-年]\d{1,2}[-月]\d{1,2}\b")
RE_HTML_VALUE = re.compile(r'\bvalue="([^"]{1,60})"')
RE_HTML_PH = re.compile(r'\bplaceholder="([^"]{1,60})"')
# 占位符里写「如 …」「例：…」的是示例，不算泄露
RE_EG = re.compile(r"如|例|示例|e\.g\.")

# —— L2 之二：代码里的兜底常量 ——
# 用户把界面上的框留空时落进交付件的就是这些值。典型坏法：兜底值写死成上一个人的姓氏缩写、
# 写死的作息时段。时段以 配置.json 的「工作时段」为唯一正解，代码里出现别的值就是硬编码。
RE_WINDOW = re.compile(r"\d{1,2}:\d{2}\s*[-–]\s*\d{1,2}:\d{2}(?:\s*,\s*\d{1,2}:\d{2}\s*[-–]\s*\d{1,2}:\d{2})*")
RE_SIGN_LINE = re.compile(r"author|initials|sign|署名|缩写|作者", re.I)
RE_SHORT_CJK_LIT = re.compile(r'["\u2018\u201c\']([\u4e00-\u9fa5]{1,2})["\u2019\u201d\']')
CODE_SUFFIX = {".py", ".js", ".html", ".htm", ".json", ".sh", ".command"}

# —— L8：引号内 ≥10 字、且含具体数字（不是 X/N/【】占位）的片段 ——
# 引号必须包含中文直角引号：本仓库的批注示例几乎都用「」，早先漏掉它，
# 结果「没有这个附件，附件 X 是另一份文件。」这类真实发现一直没被查出来。
RE_QUOTE = re.compile(r"[「『\"'‘“]([^」』\"'’”\n]{10,})[」』\"'’”]")
RE_LAWNAME = re.compile(r"《[^》\n]{0,40}》")
RE_REALNUM = re.compile(
    r"\d+\s*(?:台|万元|万|亿元|亿|年|个月|%|份|元|名|日|次|人|天|个工作日)"  # 带单位
    r"|\d+\.\d+\s*[条款项]"                                              # 第 19.4.2 条
    r"|(?:附件|附表|第)\s*\d+\s*(?:是|为)"                                # 附件 8 是…
)

# —— L9 地名 ——
RE_PLACE_GOV = re.compile(
    r"([\u4e00-\u9fa5]{2,4})(?:镇|乡|县|市|区)(?=人民政府|政府|人民法院|党委|管委会|村民委员会|村委会)")
RE_PLACE_PARK = re.compile(
    r"([\u4e00-\u9fa5]{2,6})(?:经济开发区|经济技术开发区|高新区|工业园区|工业园|产业园)")
# 光秃秃出现在正文里的乡镇村名（"项目位于××市××区××镇，"）——后面跟标点或行尾
RE_PLACE_BARE = re.compile(
    r"([\u4e00-\u9fa5]{2,4})(?:镇|乡|村|县)(?=[，。、；：）】」』\s]|$)")

RE_HOME = re.compile(r"/Users/[A-Za-z0-9_.-]+|/home/[A-Za-z0-9_.-]+|C:\\\\Users\\\\[A-Za-z0-9_.-]+")
RE_CJK_NAME = re.compile(r"^[\u4e00-\u9fa5]{1,4}$")

# 界面上的字段标签，不是人名
UI_LABELS = {"缩写", "版本名", "第几处", "改了什么", "条款位置",
             "文件名", "输出文件夹", "起", "止", "署名", "批注", "原文", "改为", "锚点", "正文"}

# value="…" 里这些是合法的（模板占位、代码片段、通用词）
VALUE_OK = re.compile(r"^\s*$|[\'\"+{}<>$]|^(审查人|审|甲方|乙方|是|否|on|off|true|false)$")

HARD = {"L1", "L2", "L4", "L6", "L7", "L8", "L9"}

ALLOW_FILE = "脱敏白名单.txt"
CONFIG_FILE = "配置.json"

SELF = Path(__file__).resolve()
RE_SELF_MASK = re.compile(r"# <<<SELF-EXEMPT.*?# SELF-EXEMPT>>>", re.S)


def surnames_of(names):
    """从词表里的中文全名派生单姓：真实泄露常常只剩一个姓（"模板里 某 的时间戳…"）。"""
    out = set()
    for w in names:
        if 2 <= len(w) <= 4 and RE_CJK_NAME.match(w):
            out.add(w[0])
    return out


def surname_hits(text, sn):
    """单姓只在两种形态下算数，否则「张冠李戴」这种成语会满屏误报：
    ① 单独被引号括起来（代码里的兜底常量、XML 示例）；② "模板里 X 的时间戳" 这类叙述。"""
    for s in sn:
        if re.search(r"[\"'「『>]\s*%s\s*[\"'」』<]" % re.escape(s), text):
            yield s
        elif re.search(r"(?:模板|文件|来件|样本|实件)里\s*%s\s*的" % re.escape(s), text):
            yield s


def config_window(root: Path):
    """配置.json 的「工作时段」是唯一正解；代码里出现别的时段值就是硬编码。"""
    f = root / CONFIG_FILE
    if not f.is_file():
        return None
    m = re.search(r'"工作时段"\s*:\s*"([^"]+)"', f.read_text(encoding="utf-8"))
    return m.group(1).replace(" ", "") if m else None


def load_names(root: Path):
    """从 脱敏词表.local.txt 读你自己的机构名与人名。文件不存在就只用内置的空表——
    此时 L4 只挡机构口吻词和带机构后缀的专名，挡不住你自己的名字。"""
    f = root / NAMES_FILE
    if not f.is_file():
        return list(NAMES), False
    out = list(NAMES)
    for line in f.read_text(encoding="utf-8").split("\n"):
        line = line.strip()
        if line and not line.startswith("#"):
            out.append(line)
    return out, True


def load_allow(root: Path):
    """看过、确认没问题的片段写进 脱敏白名单.txt，一行一条，# 开头是注释。

    只对 L8 / L9 生效——这两类靠特征猜，必然有正常内容被误伤
    （公开范本的条号、你自己编的示例）。其余各类不设白名单，宁可手动改掉。"""
    f = root / ALLOW_FILE
    if not f.is_file():
        return set()
    out = set()
    for line in f.read_text(encoding="utf-8").split("\n"):
        line = line.strip()
        if line and not line.startswith("#"):
            out.add(line)
    return out


def allowed(q, allow):
    """整句相等才放行；子串放行只在**白名单条目本身足够长**时才允许。
    早先是双向裸子串匹配，像「第 X.Y.Z 项」这种短条目会顺带放行任何包含它的长句。"""
    for a in allow:
        if a == q:
            return True
        if len(a) >= 12 and a in q:
            return True
        if len(q) >= 12 and q in a:
            return True
    return False


def iter_files(root: Path):
    for p in sorted(root.rglob("*")):
        if any(part in SKIP_DIRS for part in p.parts):
            continue
        if p.is_file():
            yield p


def texts_of(p: Path):
    """产出 (显示名, 文本)。docx/zip 会拆开看里面的 XML。"""
    if p.suffix in BIN_SUFFIX:
        try:
            z = zipfile.ZipFile(p)
        except Exception:
            return
        for n in z.namelist():
            try:
                yield f"{p}::{n}", z.read(n).decode("utf-8", "ignore")
            except Exception:
                continue
        return
    try:
        t = p.read_bytes().decode("utf-8", "ignore")
    except Exception:
        return
    if p.resolve() == SELF:
        t = RE_SELF_MASK.sub("", t)      # 只挖掉自己的说明与词表，其余照扫
    yield str(p), t


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    quiet = "--quiet" in sys.argv
    root = Path(args[0]) if args else SELF.parent
    hits = defaultdict(list)   # level -> [(file, detail)]
    allow = load_allow(root)
    win_ok = config_window(root)
    names, has_names = load_names(root)
    sn = surnames_of(names)
    n_allowed = 0

    n_files = 0
    for p in iter_files(root):
        # L7 运行产物
        name = p.name
        if (name == ".DS_Store" or name.endswith(".pyc") or name.endswith(".log")
                or ".bak" in name or "__pycache__" in p.parts
                or "日志" in p.parts or "工作区" in p.parts):
            hits["L7"].append((str(p.relative_to(root)), "运行产物/临时文件，不该随包分发"))
            continue
        if name == NAMES_FILE:
            # 词表本身当然写满了要查的词。它不参与内容扫描，
            # 但下面会核对 .gitignore 是否真的挡住了它——挡不住才是事故。
            continue
        n_files += 1
        is_html = p.suffix in {".html", ".htm"}
        is_code = p.suffix in CODE_SUFFIX
        for shown, text in texts_of(p):
            rel = shown.replace(str(root) + "/", "")
            for m in RE_MONEY.finditer(text):
                hits["L1"].append((rel, m.group()))
            for m in RE_CN_MONEY.finditer(text):
                hits["L1"].append((rel, m.group() + "（中文数字）"))
            if is_html:
                # value= 是「预填内容」：人名、日期、作息时段出现在这里 = 页面一开就顶在脸上
                for m in RE_HTML_VALUE.finditer(text):
                    v = m.group(1)
                    if VALUE_OK.search(v) or v in UI_LABELS:
                        continue
                    if (RE_CJK_NAME.match(v) or RE_DATE.search(v) or RE_WINDOW.search(v)
                            or re.search(r"\b20\d{2}-\d{2}-\d{2}T\d{2}:\d{2}", v)):
                        hits["L2"].append((rel, f'value="{v}"'))
                # placeholder= 是提示文案：只有「没标成示例的真实日期/时段」才算问题
                for m in RE_HTML_PH.finditer(text):
                    v = m.group(1)
                    if RE_EG.search(v):
                        continue
                    if (RE_DATE.search(v) or RE_WINDOW.search(v)
                            or re.search(r"\b20\d{2}-\d{2}-\d{2}T\d{2}:\d{2}", v)):
                        hits["L2"].append((rel, f'placeholder="{v}"'))
            if is_code:
                # 写死的作息时段：以 配置.json 为唯一正解，别的值都是硬编码
                if win_ok and p.name != CONFIG_FILE:
                    for m in RE_WINDOW.finditer(text):
                        if m.group().replace(" ", "") != win_ok:
                            hits["L2"].append((rel, f"写死的作息时段 {m.group()}（配置.json 是 {win_ok}）"))
                # 兜底常量里的姓氏缩写（上一个人的姓落进别人的交付件，历史上真踩过）
                for line in text.split("\n"):
                    if not RE_SIGN_LINE.search(line):
                        continue
                    for m in RE_SHORT_CJK_LIT.finditer(line):
                        v = m.group(1)
                        if v in UI_LABELS or v in ("审查人", "审", "甲", "乙"):
                            continue
                        hits["L2"].append((rel, f'署名类兜底常量 "{v}"'))
            for m in RE_DATE.finditer(text):
                hits["L3"].append((rel, m.group()))
            for w in names + TONE_WORDS:
                if w in text:
                    hits["L4"].append((rel, w))
            for s in surname_hits(text, sn):
                hits["L4"].append((rel, f"单姓「{s}」"))
            for m in re.finditer(r"([\u4e00-\u9fa5]{2,10})(" + ORG_SUFFIX + r")", text):
                pre = m.group(1)
                # 前缀里出现虚词/占位符的不是专名（"政府或管委会""某某律师事务所"）
                if NOT_IN_PLACE & set(pre) or re.search(r"某某|××|xx|XX|你所在|本|该|项目|平台|集团|下属", pre):
                    continue
                hits["L4"].append((rel, m.group()))
            for w in INDUSTRY:
                if w in text:
                    hits["L5"].append((rel, w))
            for m in RE_HOME.finditer(text):
                hits["L6"].append((rel, m.group()))
            # L8 —— 正文可能藏在任何一种文本文件里，不挑后缀
            for m in RE_QUOTE.finditer(text):
                q = m.group(1)
                # 引法规是常识，但不能因为句子里提了一部法就整句免检：
                # 只把《…》挖掉，剩下的照查
                probe = RE_LAWNAME.sub("", q)
                if "废止" in probe or "并入" in probe:
                    continue
                if not re.search(r"[\u4e00-\u9fa5]", probe):        # 纯代码/英文
                    continue
                if "{" in probe or "}" in probe:                    # f-string 格式串
                    continue
                if RE_REALNUM.search(probe):
                    if allowed(q, allow):
                        n_allowed += 1
                        continue
                    hits["L8"].append((rel, q[:60]))
            # L9 地名 —— 通用工具包里没有任何理由出现真实地名
            for rx in (RE_PLACE_GOV, RE_PLACE_PARK, RE_PLACE_BARE):
                for m in rx.finditer(text):
                    # 前缀往回取到第一个虚词为止，剩下的才是专名候选
                    tok = re.split(r"[%s]" % "".join(NOT_IN_PLACE), m.group(1))[-1]
                    if len(tok) < 2:
                        continue
                    hit = tok + m.group()[len(m.group(1)):]
                    if allowed(hit, allow) or allowed(m.group(), allow):
                        n_allowed += 1
                        continue
                    hits["L9"].append((rel, hit))

    if has_names:
        ig = root / ".gitignore"
        pats = ig.read_text(encoding="utf-8") if ig.is_file() else ""
        if not any(x in pats for x in ("*.local.txt", NAMES_FILE)):
            hits["L4"].append((NAMES_FILE,
                               ".gitignore 没挡住这份词表——一次 git add -A 就会把里面的名字推上去"))

    LABEL = {"L1": "真实金额（千分位、多位小数或中文数字）",
             "L2": "预填值/兜底常量带了人名·日期·作息时段",
             "L3": "写死的具体日期", "L4": "机构名或人名（含单姓）", "L5": "行业专属词",
             "L6": "私人路径", "L7": "运行产物", "L8": "像逐字抄来的条款（带具体数字）",
             "L9": "指向具体地方的地名"}

    if not quiet:
        print(f"扫描目录：{root}")
        print(f"检查文件：{n_files} 个"
              + (f"（白名单放行 {n_allowed} 处 L8/L9）" if n_allowed else "")
              + (f"｜词表 {NAMES_FILE}：{len(names)} 条" if has_names
                 else f"｜⚠️ 没有 {NAMES_FILE}，L4 挡不住你自己的名字")
              + "\n")

    bad = False
    for lv in ("L1", "L2", "L4", "L6", "L7", "L8", "L9", "L3", "L5"):
        rows = hits.get(lv)
        if not rows:
            continue
        mark = "❌" if lv in HARD else "⚠️ "
        if lv in HARD:
            bad = True
        agg = defaultdict(set)
        for f, d in rows:
            agg[f].add(d)
        total = sum(len(v) for v in agg.values())
        print(f"{mark} {lv} {LABEL[lv]}：{total} 处，{len(agg)} 个文件")
        for f, ds in list(agg.items())[:8]:
            shown = "、".join(sorted(ds)[:5])
            print(f"      {f}\n        → {shown}")
        if len(agg) > 8:
            print(f"      …另有 {len(agg)-8} 个文件")
        print()

    if not bad and not hits.get("L3") and not hits.get("L5"):
        print("✅ 没有发现任何一类问题，可以发出去了。")
    elif not bad:
        print("✅ 没有必须处理的问题（L3/L5 只是提示，自己判断一下是不是示例）。")
    else:
        print("上面带 ❌ 的必须处理完再发。")
        if hits.get("L8") or hits.get("L9"):
            print(f"L8/L9 是靠特征猜的，确认没问题的片段整行加进 {ALLOW_FILE} 即可放行。")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
