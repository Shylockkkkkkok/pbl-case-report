---
name: pbl-case-report
description: "Use this skill whenever the user hands you a PBL (problem-based learning) medical case — one or more 情景 PDFs or pasted case text — and wants the coursework deliverables built around a research question the case raises: a Word report that answers that question, plus a short presentation deck (usually 5 minutes). The report is an answer to a question, not a formal literature review: textbook explanations and plain reasoning are fine, journal citations are used where they carry the evidence. Trigger on PBL, 案例, 情景, 病例, case, 文献综述, 汇报PPT, 小讲课, or when the user says they will only give you the case and asks you to 调研 a topic and produce 1 份 word + 1 份 PPT. Also use it when revising formatting or content of a PBL report/deck you or a previous session produced. For decks that must be editable native PowerPoint built from research materials, it routes to a bundled engine with 10 colour schemes, 23 layouts and light/solid fills."
---

# PBL 案例作业生成

把「一个临床案例」+「一个研究问题」变成「一份 Word 报告 + 一份汇报 PPT」。本 skill 记录了一套已经过用户验收的做法：工作流、硬性排版要求、以及踩过的坑。

**文中出现的甲状腺、睡眠、UK Biobank、`fT3` 等内容，全部来自首个验收案例（内分泌 case 1，2026-10），只是示例，不是模板。换成别的系统、别的题型，骨架照样用，但术语、例子、配色必须换。`reference/` 里代码的示例数据同样来自那个案例，替换内容时整段换掉。**

**报告是「答题」，不是「写综述」。** 结构服务于把问题讲清楚，不服务于体例；教材、生理/病理基础、常识性推理都是合法内容，期刊文献用在该用的地方（数字、结论、争议）。

**先读本文件，再按需打开 `reference/` 里的模板代码。** 模板是可运行的真实代码，直接复制改内容，不要从零重写。

| 文件 | 用途 |
|---|---|
| `reference/word_helpers.py` | python-docx 排版库（字体/标题/表格/题注/参考文献），含可直接跑的示例 |
| `reference/ppt_helpers.js` | pptxgenjs 版式库（标题、要点卡、图表卡、图标、数据图适配） |
| `reference/figures.py` | matplotlib 中文图表 + 机制通路图（含中文字体设置） |
| `reference/qa.py` | PPT 版面几何自检 + 纯 Python 预览渲染（不启动 Office） |
| `.vendor/editable-research-ppt/` | 内置第三方引擎：可编辑科研 PPT（MIT）。需要时读它自己的 `SKILL.md` |

---

## 一、整体流程

1. **读案例** → 2. **定问题** → 3. **调研** → 4. **Word** → 5. **PPT** → 6. **自查** → 7. **交付**

---

## 二、读案例

- PDF 用 Read 工具（超过 10 页加 `pages` 参数）。
- **把案例里的临床图片抠出来复用**（查体照片、超声、X 线、病理图），这是 PPT 最真实、最有说服力的素材：
  ```python
  import fitz
  d = fitz.open(path)
  for pno in range(len(d)):
      for i, img in enumerate(d[pno].get_images(full=True)):
          pix = fitz.Pixmap(d, img[0])
          if pix.n - pix.alpha >= 4: pix = fitz.Pixmap(fitz.csRGB, pix)
          pix.save(f"assets/case_p{pno+1}_{i}.png")
  ```
  存下来后**逐张 Read 看一眼**再命名（`fig_eye.png` / `fig_goiter.png` …），不要凭文件名猜。
- 提取要素：人口学、时间线、诱因、症状、查体、实验室（含**参考值和箭头方向**）、影像、家族史、医生的关键一句话。
- **研究问题通常藏在以下几类地方**（下面每类只举一个例子，不要因为例子是内分泌的就以为只适用于内分泌）：
  - 医生的叮嘱或感叹：「别熬夜了」「记得复查」「这个病很容易复发」
  - 某个「看似矛盾」的检查：与主诊断不完全吻合的那一项
  - 患者自述的生活习惯或社会因素：熬夜、饮酒、节食、职业暴露、情绪压力
  - 案例/情景标题的暗示：「女强人」暗示长期高压 + 缺觉
  - 一个两难选择：治疗方案冲突、要不要手术、要不要终止妊娠、要不要告知家属

### 读完案例，先判断这是哪一类题目

同一个骨架，不同题型侧重完全不同。先归类，再决定报告写什么、PPT 放什么图：

| 题型 | 典型问法 | 报告侧重 | PPT 侧重的图 |
|---|---|---|---|
| 机制类 | 「X 为什么会导致 Y」 | 通路拆解 + 证据强度 | 机制总览图 + 分通路拆解 |
| 诊断类 | 「怎么鉴别 / 还要查什么」 | 诊断标准、鉴别表 | 诊断流程图、对比表 |
| 治疗类 | 「用不用 / 怎么选 / 要不要手术」 | 循证证据分级、一线二线 | 疗效对比图、决策树 |
| 预后/风险类 | 「会不会复发 / 危险因素」 | 队列数据、HR 与置信区间 | 生存曲线、森林图 |
| 心理社会/伦理类 | 「该不该告诉 / 怎么沟通」 | 文献 + 伦理框架 | 案例回放、观点对照 |
| 基础+临床混合 | 大多数 PBL 案例 | 前两类各占一半 | 机制图 + 临床图各一张 |

**不要每道题都硬套机制图。** 伦理、沟通、卫生经济类题目，机制图反而是减分项。

## 三、调研

### 网络受限时的唯一可靠路径

本机环境实测（2026-10）：`WebSearch` 返回畸形结果、`WebFetch` 全部域名被拦截、Bing 经 curl 会丢掉除首个词以外的查询词、维基媒体不可达。**但 NCBI E-utilities 经 curl/python 完全可用**，这是唯一可靠的可核查文献来源。

```bash
# 检索
curl -s "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term=sleep+deprivation+AND+thyroid&retmax=25&retmode=json&sort=relevance"
# 取摘要（efetch, retmode=xml）
curl -s "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&retmode=xml&id=42635905,37733750"
```

建议写两个小脚本：`pm.py`（esearch + esummary，打印 PMID/期刊/年份/作者）和 `pmab.py`（efetch 解析摘要）。**Windows 下务必 `PYTHONIOENCODING=utf-8`**，否则 GBK 编码会在打印特殊字符时崩溃。

按证据层级检索，中英文关键词组合都要试：

| 层级 | 作用 | 关键词模式 |
|---|---|---|
| Meta 分析 / 系统综述 | 主结论的数字来源 | `<暴露> AND <结局>`, `systematic review`, `meta-analysis` |
| RCT | 因果证据 | `randomized`, `clinical trial` |
| 前瞻队列 | 发病风险 | `cohort`, `incidence`, `hazard ratio` |
| 机制综述 | 通路描述 | `mechanism`, `review`（优先 Nat Rev / Endocr Rev / Lancet 系列） |
| 指南 | 临床标准 | `guideline`, `consensus` |

### 调研纪律

**哪些内容需要文献，哪些不需要**

| 内容 | 来源 | 要求 |
|---|---|---|
| 基本概念、解剖生理、经典机制 | **教材**（《内科学》《生理学》《病理学》《诊断学》等，人卫版） | 写明教材名、版次、出版社、年份；页码不确定可不写，**不要编** |
| 具体数字（患病率、RR/HR、效应量、样本量） | 期刊原文 | 必须有一手来源，能核到 PMID |
| 该领域的主流结论与争议 | 系统综述 / Meta 分析 / 指南 | 优先高影响力期刊 |
| 案例的具体表现 | case 原文 | 直接引用，标注情景号 |

- **不是每个论点都要挂引文**。教材级、常识级的内容直说不必强行引文；关键数据必须有一手来源。
- **不允许编造文献**，也不允许编造卷期页码。中文库（知网/万方/中华医学会期刊）在本环境无法在线核验，只能引你确实知道的信息；不确定的具体出处宁可省略。
- **如实报告阴性/不显著结果**。示例（首个案例）：UK Biobank 队列中「健康睡眠降低甲状腺疾病风险」在**甲亢亚型上未达统计学显著**——报告和 PPT 都要写出来，不能只挑阳性结果。换成你的题目时同理：亚组不显著、效应量很小、文献结论互相矛盾，都要照实写。这既是科研诚信，也是答辩被追问时的护身符。
- 区分**证据强度**：动物实验、机制推断要说"提示/推测"，不能写成定论。

## 四、Word 报告

### 结构：按「回答问题」组织，不要套综述体例

**先想清楚一句话答案，再决定分几节。** 下面只是常用骨架，可增删合并：

```
标题（一行，不写副标题）
（不写摘要、不写关键词）

1  引言          病例摘要 + 研究问题（1–2 段）
2  总体判断      先给结论，再列支撑它的几条证据线
3  证据与机制    教材讲机制 + 文献给数据（1–2 张表、1 张图）
4  案例分析      ← 必写（1–2 段）
5  结论          落到可操作的判断与建议
参考文献
```

各节要点：

- **问题引入**：不要复述整份病历，只保留与研究问题相关的线索（时间线、化验、医生的那句话）。
- **核心回答**：一段话讲清"是/不是、通过什么、程度如何"。老师最先看这里，别把结论埋到最后一节。
- **依据与展开**：基础机制用教材讲透即可，不必每句挂引文；关键数字（患病率、HR、效应量）要有原文。**图表宁少勿滥**——一张把机理讲清的图，胜过五张装饰图。
- **结论与建议**：给出可执行的判断或干预，不要停在"有待进一步研究"。

篇幅：完整报告 **3000–5000 字**，宁精勿滥。

**写作硬指标**：

- **问题引入（引言）1–2 段；案例联系（案例分析）1–2 段**，不要写成病历复述
- 每个机制小节 1–2 段：教材讲机制 + 至少一条带区间的数据（g、OR、HR、CI）
- 至少 1 张表 + 1 张图；表格列「证据类型 → 结果 → 推断意义」，不是罗列数字
- 单独一节讲**现有证据的局限性**
- **不写摘要、不写关键词**
- **标题用陈述性短语**：写「昼夜节律与 TSH 夜间峰值」，不写「昼夜节律：被削平的峰值？」
  更不要「A：B」式对仗
- **语体模仿医学论文**：第三人称、陈述句、少修辞；不用反问，不写「你会不会…」「是不是…」

**「案例联系」必写**，回答三件事：暴露与时间线是否吻合；症状能否一元论解释；医生的那句话有没有循证依据。长度控制在 4 段左右。

### 硬性排版要求（用户明确指定，不得更改）

| 项目 | 要求 |
|---|---|
| 字体颜色 | **全文纯黑**，包括标题、题注、表头。表头用浅灰底（`E9EDF1`）+ 黑字 |
| 副标题 | **不要**。标题下只保留一行课程标识 |
| 各级标题 | 中文 **黑体**（Heading 1/2/3），西文 **Times New Roman** |
| 正文 | 中文 **宋体**，西文 **Times New Roman**，小四（12pt），1.5 倍行距，首行缩进 2 字符 |
| 表格 | 表头跨页重复（`w:tblHeader`）、单行不跨页断开（`w:cantSplit`）、三线感/网格线、题注在表上方 |
| 图 | 居中，宽度约 15 cm（页心宽），题注在图下方 |

### 参考文献：按正文首次引用顺序编号（GB/T 7714 顺序编码制）

**这是硬性规则，不是可选项。** 编号取决于该文献在正文中**第一次被引用**的先后，不是字母序、不是重要性、也不是写作时的顺手顺序。

- 正文标 `[1]`、`[1,3]`、`[2-4]`；三个及以上连续编号压成区间
- 文献表按 `[1] [2] [3] …` 连续排列，**不允许跳号**
- 文献表里**不允许出现正文从未引用的条目**
- 边写边编号几乎必错——正确做法是**先写正文，最后统一重排**

`word_helpers.py` 已内置自动重排与校验：

```python
new_document()
... 正文先随便编号，把内容写完 ...
h1("参考文献")
ref("[1] ...")
ref("[2] ...")
renumber_refs()      # ← 保存前必须调用：按首次引用顺序重排
save("xxx.docx")
audit()              # ← 内部会调用 check_refs()
```

`renumber_refs()` 会同时改写正文里的 `[n]`、`[a,b]`、`[a-b]`，并把文献表原地重排；未被引用的条目排在最后并给出警告。已有文档可直接修复：

```python
load("旧报告.docx")
renumber_refs()
check_refs()
BLACK_DOC.save("旧报告.docx")
```

**条目格式：GB/T 7714-2015**

- 期刊：`作者. 题名[J]. 刊名, 年, 卷(期): 起页-止页.`
- 在线优先（尚无卷期页码）：`作者. 题名[J/OL]. 刊名, 年. DOI: xxx.`
- 专著：`主要责任者. 书名[M]. 版本. 出版地: 出版者, 出版年.`
- 作者：西文**姓全大写**、名用缩写且字母间留空格（`WANG S`、`LEE S Y`）；超过 3 位取前 3 位 + `, et al.`
- **不允许缺页码**：期刊条目必须带 `卷(期): 页码`；确实还没有页码的在线优先文献，必须用 DOI。查不到完整题录就换一篇能查全的，不要留半条。
- 教材按整本引用时可省略页码（GB/T 7714 对整本专著不要求引文页码）；若要求具体页码，必须查实体书，**不能编**。

**验收不能靠肉眼**：`check_refs()` 会断言首次引用顺序为 1,2,3…、编号连续、正文与文献表双向一一对应，返回 False 就不要交付。

### ⚠️ 最大的坑：主题字体覆盖

Word 内置 Heading 样式的 `<w:rFonts>` 同时带有：

```xml
w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"
w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"
```

**主题属性的优先级高于显式属性**。只设 `w:eastAsia="黑体"` 而不删主题属性，Word 仍会渲染主题字体（等线 Light）——标题看起来根本不是黑体，而且从属性上检查还以为改对了。

正确做法（两层保险）：

```python
for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
    if rf.get(qn(attr)) is not None:
        del rf.attrib[qn(attr)]
rf.set(qn("w:ascii"), "Times New Roman")
rf.set(qn("w:hAnsi"), "Times New Roman")
rf.set(qn("w:cs"), "Times New Roman")
rf.set(qn("w:eastAsia"), "黑体")
```

然后**再给每个标题段落的 run 显式写一遍字体**，不依赖样式继承。

> 验收时不能只读 `w:eastAsia`，必须同时断言主题属性残留数为 0。

### 生成方式

用 `python-docx`（本机可用）。`reference/word_helpers.py` 提供了 `set_font` / `h1..h3` / `title_block` / `body` / `bullet` / `keywords` / `caption` / `picture` / `make_table` / `ref` / `audit` 等函数，`set_font` 与标题样式已经内置主题属性清理。直接 `from word_helpers import *` 改内容即可。

⚠️ 用 heredoc 写超长 Python 脚本会被 shell 截断（报 `unexpected EOF`）。**长脚本一律用 Write 工具落盘再运行。**

## 五、PPT

### 需要"可编辑科研 PPT"时

本仓库内置了第三方 skill `skills/editable-research-ppt/`（MIT，v1.7.1）。当用户要求
**可编辑的原生形状 PPT、十套配色、按论文资料自动匹配版式**，或明确指名该 skill 时，
按它自己的 `SKILL.md` 执行，不要用本文件的版式强行套。

### PPT 生产：必须走引擎，模板按主题选（都不要默认）

**硬要求：做 PPT 一律调用 `editable-research-ppt`。** 不要凭本文件的开发经验手搓版式——
先读 `.vendor/editable-research-ppt/SKILL.md`，用它的 `generate.py` 生成、`audit.py --strict` 验收。

#### 引擎的完整流程

1. **三问**：场景（`academic` / `coursework` / `progress`）→ 配色 → 填充（`light` / `solid`）。
   配色**按主题选**，见下表，**不设默认**；用户说"你定"时也要先按主题判断一次再选。
2. **输入 JSON**：逐页写 `layout`（`L1`–`L23`）与 `variant`，显式声明 `demo: false`、`fill_style`。
3. **命令**（在 `.vendor/editable-research-ppt/` 下执行）：
   ```console
   python scripts/check_environment.py
   python scripts/preflight.py --input INPUT.json
   python scripts/generate.py --input INPUT.json --output OUTPUT.pptx --force
   python scripts/audit.py OUTPUT.pptx --strict     # 必须 0 Error、0 Warning
   ```
4. **字号**：常规版 L1–L20 标题 42px、正文 21–25px；紧凑版 L21–L23 标题 36px、正文 20–21px。
5. **图片**：有论文原图时必须从原文裁取（`scripts/crop_paper_figure.py`），不得自绘替代、不得编数据。
6. 不得修改 `.vendor/` 内文件；交付前 `verify_release.py` 必须为 `ok`。

#### 配色按主题选（P1–P10）

| 主题 | 建议配色 | 理由 |
|---|---|---|
| 临床医学、内科学、诊断 | P1 蓝白朱红 / P9 孔雀青白 | 冷静、不抢影像图 |
| 基础医学、生理、病理机制 | P3 石墨青金 / P6 湖蓝珊瑚 | 高对比，适合通路图 |
| 公共卫生、流行病、队列 | P2 松绿米白 | 自然、适合地图与曲线 |
| 神经、精神、认知 | P5 靛紫芥黄 | 层级清晰 |
| 药理学、毒理、分子 | P8 深蓝香槟金 | 强调剂量/结构 |
| 护理、人文、伦理 | P4 酒红灰蓝 / P7 赭橙米白 | 柔和 |
| 需要极简黑白 | P10 黑白朱红 | 印刷友好 |

拿不准时，**先说明你的判断和理由**，再问用户要不要换；不要默默套一个固定配色。

#### 学校模板：只在需要校徽/校色时用，同样按主题选

引擎**不能加载外部 .pptx 模板**，所以学校模板与引擎版式不能叠在同一份 deck 上。
需要交大品牌时走模板路线，从 `templates/xjtu/` 里选：

| 主题 / 场合 | 模板 | 理由 |
|---|---|---|
| 正式开题、毕业答辩 | `西安交通大学汇报答辩通用模板`（校红） | 官方感最强 |
| 工科、创新港相关 | `2022 西安交通大学创新港风格模板` | 主题吻合 |

**没有任何"默认模板"。** 每次都要给出选型理由（一句话），让用户有机会改。现有 3 套全部为 16:9。

#### 引擎 + 校徽：用 `apply_school_skin.py` 组合

引擎不能直接加载外部 `.pptx` 模板，但可以先出稿、再套壳：

```bash
python reference/apply_school_skin.py   --deck OUTPUT.pptx   --template templates/xjtu/飞上云霄.pptx   --out FINAL.pptx
```

它做三件事：用所选模板的**封面页**替换第 1 页、用**致谢页**替换末页、给每个内容页贴**校徽**（保留 `srcRect` 裁剪）并在底部加一条校色线；页码若被校徽压住会自动挪到右下角。

套壳后**必须再跑一次** `python reference/qa.py FINAL.pptx`，确认没有新的重叠或溢出。

**图文并茂**：纯文字页偏多时，用 `reference/add_image_slide.py` 插入配图页（病例照片、效应量图、机制示意图）。图片来源必须是真实材料——病例图从案例 PDF 抠、数据图用真实数值、示意图需标注「示意图」。

#### 学校模板的使用细则

`templates/xjtu/` 内置 3 套交大模板，详见该目录的 `README.md`。**按主题选，不设默认。**

用模板的正确姿势：

1. 以模板文件为底稿，不要另起炉灶只借配色。
2. **保留模板的封面页与结尾页**，只替换文字——封面文字常在组合形状（GROUP）里，要递归遍历 `shape.shapes`。
3. 内容页用模板的 `空白` / `1_空白` 版式新建：版式自带校徽（右上，约 3.4×1.0 英寸）和页眉线（y≈0.91），**不要再自己贴 logo**。
4. 页面元素避让右上角校徽：标题宽度不超过约 9 英寸，页码放右下角。
5. 换页序后参考文献编号会变——必须重新按首次引用顺序编号，见第四节。

用户没有指定时：先按主题给出一套推荐并说明理由，再问是否更换。



### 页数与时长的对应

5 分钟 → **12–13 页**（含封面与参考文献页），每页 20–30 秒。宁可少一页，也不要塞满。

### 推荐页序

一套通用骨架，括号里是「这一页该放什么视觉元素」，按题目替换：

```
1  封面（深色）
2  案例导入（时间线 + 案例里的真实图片 + 医生的原话）
3  问题背景 + 核心问题框
4  证据一（最强的一手证据：Meta 分析 / RCT / 指南）
5  证据二（人群或机制证据：队列图 / 动物实验 / 机制综述）
6  核心机制或核心流程总览（满版大图）
7  展开①（示意图 / 流程图 / 对比表）
8  展开②（双栏：两条通路 / 两种方案对照）
9  展开③（双栏；论点少于 3 条时与上一页合并）
10 关键关系图（环形图 / 决策树 / 时间轴——本页是重点）
11 案例联系与临床启示（左「回到患者」，右「临床启示」）
12 参考文献（整页）
13 谢谢大家（深色，与封面呼应）
```

7–9 页的条数按你的机制/论点数量增减：3–5 条最常见。只有 1 条核心机制时，把第 6 页做满，7–9 合并成一页「证据与边界」。

### 设计系统

**固定的部分**（不随选题变）：

| 元素 | 取值 |
|---|---|
| 标题字体 | `Source Han Sans CN`（思源黑体），26pt 加粗 |
| 正文字体 | `Noto Sans SC` |
| 视觉母题 | 标题左侧 **0.085 英寸**竖条；封面/结尾一条主题线条（心电、DNA、时间轴…） |
| 结构 | 深色封面 → 浅色内容页 → 深色参考文献/致谢，构成「三明治」 |

**颜色必须按选题换。** 下表是首个案例（内分泌/睡眠）用的，只是参考，不要照搬到所有题目：

| 元素 | 内分泌示例 |
|---|---|
| 主色（夜晚） | `0E2233` / `17384F` / `24506B` |
| 临床青 | `1C7293`，浅 `4FA3BF` |
| 强调琥珀（高代谢/热） | `E8871E` |
| 警示红 | `C0392B` |
| 辅助 | 绿 `0D9488`、紫 `7A4E9B` |
| 背景 | 内容页 `F7FAFC`，封面/结尾/参考文献页 `0E2233` |

**选色规则**：

- 深色背景 = 题目里最「重」的那个意象（夜、深海、血液、旧照片、消毒灯）
- 一个主色占 60–70% 视觉重量，辅助 1–2 个，**强调色只留 1 个**，给最想让人记住的数字或通路
- 不要默认蓝色。说不出「为什么是这个颜色」，就换一个
- 同一份 deck 里颜色语义要稳定：红永远是警示，不要一处红表危险、一处红表强调

换色起手式（自行调整，别照抄）：

| 选题类型 | 深色 | 主色 | 强调 |
|---|---|---|---|
| 内分泌 / 睡眠 | `0E2233` | `1C7293` | `E8871E` |
| 心血管 / 急诊 | `1A1F2B` | `8C1C13` | `F2A65A` |
| 肿瘤 / 血液 | `221B2E` | `6D2E46` | `D4A373` |
| 感染 / 呼吸 | `132A2B` | `2C7A7B` | `E9C46A` |

改配色时，`ppt_helpers.js` 顶部的 `C` 对象和 `figures.py` 顶部的颜色常量要**一起改**，否则图里的颜色和文字色条对不上。

**不要**用标题下划线（AI 味很重），用左侧色块代替。深色页与浅色页交替构成「三明治」结构。

### 硬性要求

- **每页只要引用了文献，页脚就要有一行小字参考文献**（距底 0.7 英寸左右，8.8pt，灰色），参考文献页除外。页脚可省题名，但必须含 **刊名, 年, 卷(期): 页码**；没有页码的写 DOI，不能只写「作者 + 年份」。
- **页脚编号与文献表编号一致**，并按幻灯片顺序递增；不要每页从 `[1]` 重新开始，否则和末页文献表对不上。
- **页脚不能压住任何内容**：放在卡片下沿与要点条之间的空带（约 y = 6.5–6.8，7.8–8.8pt）。交付前跑 `qa.py`，出现 `FOOTER-OVERLAP` 必须调整位置或缩短题录。
- **最后一页前放参考文献整页**，再放**「谢谢大家」**页。
- 参考文献页**不要**写"注：……"之类的说明行。
- **字号优先级高于留白**：宁可少写两行字，也要把字放大。默认基线：幻灯片标题 **28pt**、正文 **≥14pt**、卡片标题 **≥15.5pt**、小标签 ≥12pt、页脚文献 ≥8.8pt。
- **唯一的例外是参考文献页**：题录密集时可用 9pt，并按实际行数把条目均分到三栏，不要按条数硬分。
- 图表自带的字号如果被缩放后小于 ~9pt，**要重新出图而不是硬塞**（见"图表适配"）。
- **字与字之间不要有多余空白**：不要用 `charSpacing`；不要用全角空格 `　` 分隔数值（用 `·` 或逗号）。

### 生成方式

`pptxgenjs`（全局已安装）。运行前设置：

```bash
export NODE_PATH="C:/Users/86188/AppData/Roaming/npm/node_modules"
node build_ppt.js
```

`reference/ppt_helpers.js` 提供了全套版式函数（`contentSlide` / `title` / `refLine` / `card` / `iconHead` / `chartCard` / `pointCards` / `fitBox`），复制后替换内容即可。

图表用手绘：`matplotlib` 出 PNG（透明背景），或直接用 pptxgenjs 形状画（环形图、流程图）。

### ⚠️ 两个必踩的坑

1. **pptxgenjs 不能画负的宽/高。** 用 `x/y/w/h` 画斜线箭头时，若终点在起点左上方，会写出 `cx="-1309494"`，PowerPoint **直接打不开文件**（报"PowerPoint could not open the file"）。必须改成：
   ```js
   const dx = x2 - x1, dy = y2 - y1;
   s.addShape(pptx.ShapeType.line, {
     x: Math.min(x1, x2), y: Math.min(y1, y2),
     w: Math.abs(dx), h: Math.abs(dy),
     flipH: dx < 0, flipV: dy < 0,
     line: { color, width, endArrowType: "triangle" },
   });
   ```
   交付前务必扫一遍所有 slide XML 里有没有负的 `cx`/`cy`。

2. **图注会和图体重叠。** 只用 `sizing:{type:"contain"}` 时图片在框内居中，按固定 y 写图注会压到坐标轴标签。要**先按真实宽高比算实际渲染尺寸**，把图注放在图片底边之下：
   ```js
   let dw = boxW, dh = dw / ar;
   if (dh > boxH) { dh = boxH; dw = dh * ar; }
   const dy = boxY + (boxH - dh) / 2;
   // 图注 y = dy + dh + 0.06
   ```
   宽高比用 `PIL` 或 `sharp` 读实际像素，不要手算。

### 图表适配（字号问题的根源）

宽图（宽高比 ≈ 2.0）放进竖向卡片会被高度限制，两侧留白、整体缩小，图内文字跟着变小。**解决办法是重新出图，把宽高比调大**（如 1.87 → 2.19），同时把图内字号整体上调，让它在 PPT 里的实际字号 ≥ 9pt。不要靠调 PPT 文字字号来补救图片里的字。

## 六、自查（交付前必做）

**绝对不要用 PowerPoint COM 自动化来导出图片** —— 脚本结尾的 `$ppt.Quit()` 会关掉整个 PowerPoint 进程，把用户正在编辑的演示文稿一起关掉（用户明确反对过）。

改用 `reference/qa.py`：

1. **几何检查**（python-pptx 读坐标）：形状越界、文字溢出（按字号估算所需行高）、文字框互相重叠。装饰性出血圆算已知例外。
2. **预览渲染**（PIL 直接按形状几何绘制）：用于肉眼确认。注意这是近似渲染，**不支持**：斜线形状（`prst="line"` 需单独识别）、空 `<a:ln/>`（无描边，别画边框）、PowerPoint 的断词规则（英文可能显示为断词）。

Word 同理：读 XML 断言字体/颜色，不要靠肉眼。

## 七、交付

- 两个文件放在 case 目录下，中文文件名：`<研究问题>.docx`、`<研究问题>-汇报PPT.pptx`。
- 中间产物（build_*.py / build_ppt.js / assets/ 图片 / 预览图）保留 build 脚本和 assets，删掉预览图和临时 ps1。
- 汇报时说明：页数、每页讲多少秒、哪几页可以快速带过、参考文献语言与局限。

## 八、检查清单

复制到回复里逐条确认：

**Word**
- [ ] **研究问题有明确答案**，且在前半篇就能看到（不是埋在结论里）
- [ ] 正文 ≥4000 字，每个机制小节有教材机制 + 至少一条带区间的数据
- [ ] 基础机制用教材讲清，关键数字有文献来源
- [ ] 有「案例联系」一节
- [ ] 全文文字为纯黑，无彩色标题/题注/表头
- [ ] 无副标题
- [ ] Heading 1/2/3 中文黑体、西文 Times New Roman，且**主题字体属性已清除**
- [ ] 正文宋体 12pt，1.5 倍行距，首行缩进 2 字符
- [ ] 表头跨页重复、行不跨页断；图表有编号题注
- [ ] 参考文献按正文首次引用顺序编号、连续无跳号、无未引用条目（`check_refs()` 通过）
- [ ] 条目符合 GB/T 7714：有 `[J]`/`[M]` 标识、西文作者姓全大写、**卷(期): 页码齐全**，无页码的在线优先文献带 DOI

**PPT**
- [ ] 12–13 页，5 分钟节奏
- [ ] 每页都有视觉元素（图/表/图标/形状）
- [ ] 引用当页文献的页面，页脚有参考文献，且含刊名、年、卷(期): 页码（或 DOI）
- [ ] 倒数第二页为参考文献整页，**无"注"字说明**
- [ ] 最后一页为「谢谢大家」
- [ ] 标题 28pt、正文 ≥14pt、卡片标题 ≥15.5pt；文献页可 9pt；无 `charSpacing`，无全角空格分隔
- [ ] 无负的 `cx/cy`
- [ ] 图注未压到图体
- [ ] 未使用 PowerPoint COM
- [ ] 若用了内置引擎：`verify_release.py` 通过，且未改动 `.vendor/` 内文件
- [ ] 若套了学校外观：`apply_school_skin.py` 之后重新跑过 `qa.py`
