# Word 写修订的真实形态（对照实件校准）

校准样本：
1. Word（w16du 世代，2023+）在 Windows 中文环境下手工做的一处删除、一处插入；
2. 一份混合修订样本——同一文档里两个作者的修订与若干条批注，
   含段落插入/删除标记、rPrChange、commentsExtensible 全套部件。
3. 一份 WPS 生成的文档（带较多格式修订）——见第 13 条的 WPS 包特征。
下面每一条都是"和 Word 一模一样"的要求，脚本已按此实现；改动脚本时不要破坏。

## 1. 时间：w:date 是本地钟面 + 一个 Z

```xml
<w:del w:id="0" w:author="某某律所 张三"
       w:date="2026-08-25T14:17:00Z"
       w16du:dateUtc="2026-08-25T06:17:00Z">
```

- `w:date` 写的是**北京时间的钟面值**（14:17），后面那个 `Z` 是 Word 的历史包袱——
  它并不做时区换算，显示时原样呈现。所以"北京时间不折算 UTC"与"跟 Word 一致"在这里
  是同一件事；**光秃秃不带 Z 反而不像 Word**。
- `w16du:dateUtc` 才是真 UTC（钟面 −8h），Word 2023 起才写。
- **只有当文档根节点声明了 `xmlns:w16du` 时才写 dateUtc**。给一份 Word 2010 世代的
  老文件塞 2023 的命名空间，比不写更扎眼——版本特征要和来件同一个年代。
- **秒恒为 `:00`**。Word 的修订时间只到分钟，出现 `14:17:49` 一眼假。

## 2. 一次编辑动作 = 一个时间戳

样本里删除与插入是同一次编辑，两者时间完全相同。所以同一段落内的所有修订与批注
共享同一个时间戳，不做秒级微差；段落之间才递进。

## 3. w:id 从 0 开始

Word 在新文档里就是 0、1、2……而不是从某个大数起跳。

## 4. 删除的 run 带 w:rsidDel

```xml
<w:del ...><w:r w:rsidDel="008F0880"><w:rPr>…</w:rPr><w:delText>甲</w:delText></w:r></w:del>
```

值取文档自己 `settings.xml` 里的 `w:rsidRoot`。插入的 run 反而**不带** rsid（样本如此）。

## 5. 新段落带 w14:paraId / w14:textId

文档本身在用 paraId 时（Word 出品的文件都用），新插入的段落也要带；
LibreOffice 转出来的文件没有 paraId，那就别加。

## 6. word/people.xml 要有作者条目

```xml
<w15:person w15:author="某某律所 张三">
  <w15:presenceInfo w15:providerId="None" w15:userId="某某律所 张三"/>
</w15:person>
```

Word 每次保存修订都会写这个部件；缺了它是"这文件没被 Word 编辑过"的痕迹。

## 7. docProps/core.xml 要显示是我方保存的

`cp:lastModifiedBy` 改为作者，`cp:revision` +1，`dcterms:modified` 更新。
**注意此处的 `dcterms:modified` 是真 UTC**（06:17Z 对应北京 14:17），
与 `w:date` 的规矩正好相反——两处别写反。

## 8. 保留插入点的字体提示

样本里每个 run 的 `<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>` 都在。
脚本按"复制被改动 run 自己的 rPr"处理，天然满足。

## 9. 批注的真 UTC 在 commentsExtensible.xml，不在批注元素上

模板实件（文档声明了 w16du/w16cex，修订元素带 dateUtc）里，`<w:comment>` **不带**
`w16du:dateUtc` 属性。批注的真 UTC 存放在 `word/commentsExtensible.xml`：

```xml
<w16cex:commentExtensible w16cex:durableId="00000001" w16cex:dateUtc="2026-08-25T09:25:00Z"/>
```

durableId 经 `commentsIds.xml` 的 paraId→durableId 关联到批注。dateUtc 是真 UTC
（09:25Z ↔ 北京 17:25）。所以：
- stamp.py 为**本次新增批注**补 commentsExtensible 条目（真 UTC = 钟面 −8h）；
- 来件既有的该部件与其中条目**一字不动，绝不删除**——删部件不清理 rels 的
  Relationship 与 [Content_Types] 的 Override 就是悬空引用，Word 报"文件损坏"。

## 10. 批注样式的 styleId 是按文档解析的

中文环境 Word 文件里样式 id 常被压缩：模板里"批注引用"（annotation reference）的
styleId 是 `ad`，不是 `CommentReference`。要按 styles.xml 里 `<w:name w:val="annotation
reference"/>` 反查真实 id；查不到就不写样式，绝不写死英文 id 引用一个不存在的样式。
模板里批注正文段落甚至没写 pStyle——Word 也这么干，宁缺毋滥。

## 11. 真人的编辑时间在文档顺序上不是单调的

模板里同一作者的时间戳按文档顺序是 17:24 → 17:25 → 17:13 → **14:17** → 17:13 → 17:14……
真人改文档来回跳，先改正文末尾、再回头改开头是常态。所以"文档顺序上时间单调"
只是 stamp.py 默认生成的合理形态（一遍从头看到尾），**不是**校验硬性要求，
verify.py 对乱序只提示不拦截。

## 12. w:id 是 Word 每次保存时按文档顺序重编的

模板里修订与批注锚点的 w:id 共用一个序列（0,1,2,…,41），严格按文档顺序递增而与
作者、时间无关——这是 Word 保存时统一重编号的结果。脚本用"取当前最大值 +1"追加，
虽然新 id 不一定落在文档顺序位置上，但客户下次用 Word 保存时会被自动重编，无伤。
不要为了模仿这一点去改别人修订的 id。

## 13. WPS 来件的两个包特征（对照真实来件校准）

- **zip 里带目录条目**（`word/`、`_rels/` 这类以 `/` 结尾的空条目）。重打包必须原样
  保留并沿用来件条目顺序（finish.rezip 已做），否则 C1 会把它们判成"部件缺失"。
- **正文 run 级到处是 rPrChange**（格式修订）：`<w:rPr>…<w:rPrChange…><w:rPr>…</w:rPr>
  </w:rPrChange></w:rPr>`，一份来件里能有上百处；且 WPS 写的修订时间**带秒**
  （如 `14:47:18Z`）——那是第三方痕迹，照样一字不动，别按我方"秒恒为 00"的规矩去纠正。
- 提取 run 级 rPr 必须配对计数（wpara.extract_rpr）——非贪婪正则在 rPrChange 的
  内层 `</w:rPr>` 截断，render 出的 run 直接是非法 XML。
- 切分一个带 rPrChange 的 run 时，每一份都要复制完整 rPr（否则被切开的文字丢了
  第三方的格式修订记录），于是同一个 rPrChange 的 `w:id` 会出现好几次。ECMA-376 要求
  修订 id 唯一，真 Word 拆 run 时也给新副本分配新 id ——
  `wredline.dedupe_change_ids` 在写回前只改重复副本的 id 数字，作者、时间、被记录的
  旧格式一字不动；C0 复核"本次没新产生重复 id"。

## 14. 段落标记的 rPr 里可能嵌套 rPrChange / pPr 里可能嵌套 pPrChange

模板段落标记实例：`<w:rPr><w:ins …/><w:rFonts…/><w:rPrChange …><w:rPr><w:ins …/>…
</w:rPr></w:rPrChange></w:rPr>`。pPrChange 里还会再嵌一层完整的 `<w:pPr>`。
一切"从 pPr/rPr 里取内容"的代码都必须按开闭配对计数（wpara.split_ppr /
pmark_flags），非贪婪正则在内层闭合标签处截断，切出非法 XML——C0 会拦，
但别指望 C2 抓到（它是纯文本比对）。

## 14. 新建部件时抄命名空间：别抄到 `<?xml?>` 声明上（2026-09 实件踩过）

`document.xml` 的第一行是 `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`，
**它的 `>` 在根元素之前**。所以

```python
head = xml[:xml.index(">") + 1]        # 错：抓到的是 XML 声明，一个 xmlns 都没有
```

拿这个 `head` 里的 `xmlns:*` 去写新建的 `commentsExtensible.xml`，写出来的就是
`<w16cex:commentsExtensible >`——前缀没绑定，XML 非法，Word 报文件损坏，C0 会拦下
（报 `unbound prefix`）。正确做法是找真正的根元素起始标签：

```python
root = re.search(r"<w:document\b[^>]*>", xml).group(0)   # stamp.py: root_tag()
decls = " ".join(re.findall(r'xmlns:\w+="[^"]*"', root))
```

配套两条：

- `mc:Ignorable="…"` 只在 `xmlns:mc=` 确实抄到了的时候才写，否则又是一个未绑定前缀；
- 新建的部件当场 `ET.parse()` 验一次（stamp.py 的 `must_parse()`），错误信息直接指到文件，
  不必等 C0 倒推。

这一类「新建部件」路径平时跑不到（来件多半已有这些部件），改完脚本务必跑
`scripts/自检.py 某份真实来件.docx`：它会造出「没有 commentsExtensible」「没有 people」
「整份来件一条批注都没有」等变体，把这几条路径都走一遍。
