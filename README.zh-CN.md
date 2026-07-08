# Word Typesetting（中文 Word 排版 Skill）

[English](README.md) · **简体中文**

[![Release](https://img.shields.io/github/v/release/AllenWang2005/Word-typesetting?label=release)](https://github.com/AllenWang2005/Word-typesetting/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![CI](https://github.com/AllenWang2005/Word-typesetting/actions/workflows/ci.yml/badge.svg)](https://github.com/AllenWang2005/Word-typesetting/actions/workflows/ci.yml)
[![Stars](https://img.shields.io/github/stars/AllenWang2005/Word-typesetting?style=flat)](https://github.com/AllenWang2005/Word-typesetting/stargazers)
![Skill](https://img.shields.io/badge/skill-Codex%20%7C%20Claude%20Code-7c3aed)

一个可复用的 **Codex / Claude Code 技能(skill)**,把正式的中文 Word(`.docx`)报告
排成统一的学术规范——适用于课程设计、实习报告、论文式文档、工程计算报告等。

它由两部分协同工作:

1. **一套书面规范**(`SKILL.md` + `references/`),告诉模型一份合规报告到底长什么样——
   字体排版、文档结构、表格、图件、公式、引用、数字与单位。
2. **一套工具**(`scripts/`),把排好的 `.docx` 按规范中"机器可检"的部分做**检查**,
   并对最安全的问题做**自动修复**。

真正动手排版的是模型(配合通用的 DOCX 工具);本 skill 负责提供规则、完成度清单和一道
护栏,保证明显的问题不漏网。

---

## 能干什么

**字体排版与语言**
- 中文正文宋体/SimSun,英文与数字 Times New Roman;一、二级标题黑体(不加粗),三级标题宋体加粗。
- 完整字号层级——见下方[字号层级表](#字号层级)。
- 按主语言归一化标点(中文 vs 英文),专门符号(`……`、`——`、`《 》`),并保护网址、DOI、代码、公式、小数和引用方括号。
- 正文两端对齐,首行缩进 2 字符,行距默认 1.5 倍。

**文档结构与页面设置**
- 标准章节顺序(封面 → 声明 → 中英文摘要 → 目录 → 正文 → 参考文献 → 附录 → 致谢)。
- A4 页面、页边距/装订线、页眉页脚,通过分节实现"罗马数字→阿拉伯数字"页码。
- 自动目录(可含图/表目录),标题多级编号(`1 / 1.1 / 1.1.1`),孤行控制。
- 标题左对齐(题目居中)、末尾不加标点、层级不超过 3–4 级。

**表格、图件、公式**
- 白底三线表(顶/底线约 1.5 磅、栏目线约 0.75 磅),表内文字五号(10.5 pt,比正文小一号)且居中,**表题在上**,跨页重复表头,表注用小五置于表下。
- 图件**图题在下**、按章编号(`图 1-1` / `表 2-3` / `(3-1)`),遵循"先引用后出现"。
- 公式用 LaTeX 编写并渲染为**原生 Word OMML**——变量斜体,单位/运算符/函数正体;公式编号按章右对齐,正文写"由式 (3-1) 可得"。

**引用与参考文献**
- 正文引用上标,**把整个 `[1]`(含中括号)做成 Word 交叉引用**(而不是只引数字)。
- 参考文献按 **GB/T 7714—2015** 著录(默认顺序编码制,支持著者—出版年制),带正确的类型标识(`[J] [M] [D] [C] [S] [P] [EB/OL]` 等)。

**数字、单位与附录代码**
- 数字与单位间留**半角空格**(`20 m³/s`)——但 `%`、`°`、`℃`(`°C`)紧跟数字不留空格(`50%`、`30°`、`25℃`);单位遵循 GB 3100/3101/3102,数字用法遵循 GB/T 15835。
- 附录代码放在正文之后,**注释标红**(`C00000`)。

## 字号层级

以下为默认值;学校 / 期刊模板优先。

| 结构元素 | 字号 | 字体 / 字形 |
| --- | --- | --- |
| 封面题目 | 二号(22 pt) | 黑体加粗,居中 |
| 部分标题(摘要 / 目录 / 参考文献等) | 小二(18 pt) | 黑体,不加粗 |
| 一级标题(章) | 三号(16 pt) | 黑体,不加粗;段前段后 0.5 行 |
| 二级标题(节) | 四号(14 pt) | 黑体,不加粗 |
| 三级标题(小节) | 小四(12 pt) | 宋体**加粗** |
| 正文 | 小四(12 pt) | 中文宋体 / 西文与数字 Times New Roman |
| 摘要、关键词 | 小四(12 pt) | 宋体 |
| 图题、表题 | 五号(10.5 pt) | 宋体(中)/ Times New Roman(英),居中 |
| 表格内文字 | 五号(10.5 pt) | 宋体,比正文小一号 |
| 表注 | 小五(9 pt) | 宋体 |
| 参考文献条目 | 五号(10.5 pt) | 宋体,悬挂缩进 |
| 页眉、页脚 | 小五(9 pt) | 宋体 |
| 脚注 | 小五(9 pt) | 宋体 |

## 排出来什么样

一份排版整洁、可导航的报告:可点击的标题大纲和自动更新的目录;左对齐的黑体标题压在统一字号的
宋体正文之上;干净的白底三线表(表题在上)、图件(图题在下)且都按章编号;真正的 Word 公式对象
而不是截图或纯文本;上标 `[1]` 式引用因为是 Word 交叉引用字段而能自动重排编号;以及一份
GB/T 7714—2015 的参考文献表。对成品运行审计脚本会输出 `PASS: no machine-detected guardrail issues.`

可以看 [`examples/`](examples/):一个故意**不合规**的样例和它产生的审计报告——直观展示规则能抓到什么。

## 安装

从 GitHub 安装:

```text
AllenWang2005/Word-typesetting
```

或把文件夹复制到你的 skills 目录:

```text
~/.codex/skills/word-report-formatting                 # Codex,macOS / Linux
%USERPROFILE%\.codex\skills\word-report-formatting     # Codex,Windows
~/.claude/skills/word-report-formatting                # Claude Code
```

## 使用

在创建或润色 Word 报告时让助手使用该 skill:

```text
Use $word-report-formatting to format this course design report.
```

涉及正式中文 Word 报告、三线表、OMML 公式、图表题注、引用交叉引用、附录代码时,它也会自然触发。

## 审计脚本

排好 DOCX 后,运行这道护栏:

```text
python scripts/audit_docx_format.py path/to/report.docx
python scripts/audit_docx_format.py path/to/report.docx --json   # 机器可读
```

它会区分 `FAIL`(机器可确定的违规)和 `WARN`(需人工复核),结尾给出 `FAIL/WARN` 汇总,并在某类问题过多时提示省略条数。检查项包括:

| 代码 | 级别 | 抓什么 |
| --- | --- | --- |
| `ZH_PUNCT` / `EN_PUNCT` | FAIL | 正文里用错语言的标点 |
| `ABSTRACT_INDENT` | FAIL | 摘要/关键词居中或缩进 |
| `H1_CENTER` | FAIL | 一级标题居中(应左对齐) |
| `H1_GAP` | WARN | 一级标题前缺少空行 |
| `HEADING_PUNCT` | WARN | 标题末尾带标点 |
| `HEADING_FONT` / `BODY_FONT` | WARN | 一二级标题用了宋体 / 正文用了黑体(仅查直接字体) |
| `HEADING_NO_STYLE` | WARN | 看着像标题但没用 Word 标题样式 |
| `TABLE_SIZE` | FAIL | 表内文字或表内公式非五号/10.5 pt(小四亦可;不得大于正文) |
| `TABLE_BORDERS` | FAIL | 表格有竖线/内部网格,含引用的表格样式画出来的网格 |
| `TABLE_SHADING` | FAIL | 表格/单元格带底纹,含表格样式 `firstRow` 条件格式带来的表头底纹(三线表必须白底) |
| `TABLE_RULES` | FAIL | 三线不对:顶/底线缺失、行间横线、或表头下线不比顶/底线细 |
| `TABLE_HEADER_REPEAT` | WARN | 多行表格表头未设跨页重复(`w:tblHeader`) |
| `CAPTION_POSITION` | WARN | 表题在表下 / 图题在图上 |
| `CAPTION_ALIGN` | FAIL | 表题/图题段落没有居中 |
| `FLOAT_ORDER` | WARN | 图/表出现在正文首次引用之前(应先引用后出现) |
| `COLOR` | WARN | 多余的非黑色字体(已排除超链接/主题色) |
| `CITATION_BRACKETS` | FAIL | 全角引用括号 `［1］` |
| `CITATION_FIELDS` | FAIL | 有引用但无 `REF ref_###` 字段 |
| `CITATION_NO_BRACKETS` | WARN | 裸上标数字引用、丢了中括号 |
| `VISIBLE_LATEX` | FAIL | 残留可见 LaTeX 源码而非 OMML |
| `FORMULA_TEXT` | FAIL/WARN | 本应是 OMML 的纯文本公式/量符号 |
| `MATH_DUPLICATE` | FAIL | 公式对象的内容仍以纯文本形式留在同一段落(追加而非原位替换) |
| `FORMULA_DIGIT_ITALIC` | FAIL | 公式里的数字/运算符被斜体,含斜体 `F=44.5` 这类混合 run |
| `FORMULA_MULTILETTER_ITALIC` | FAIL | 公式里 2+ 相邻字母停留在 OMML 默认斜体(单位/函数名需 `m:sty="p"` 正体;`CI` 式系数需拆开) |
| `MANUAL_ITALIC_MATH` | FAIL | 用普通文本手工斜体冒充公式(如斜体 `F = 44.5 km²`) |
| `NUMBER_UNIT_SPACING` | WARN | 数字与单位粘连(`20km` 应为 `20 km`),或 `%`/`°`/`℃` 前多了空格 |
| `EQUATION_NUMBER_CENTER` | FAIL | 带编号的独立公式整体居中(编号应靠右) |
| `EQUATION_NUMBER_TABS` | FAIL | 带编号的独立公式没有右对齐制表位放编号 |
| `EQUATION_UNNUMBERED` | FAIL | 独立公式没有按章编号 `(3-1)` |
| `EQUATION_NOT_REFERENCED` | WARN | 公式编号从未在正文中被引用(应写"由式 (3-1) 可得") |
| `FORMULA_TEXT_TABLE` | WARN | 表格单元格里残留本应是 OMML 的纯文本公式/符号 |
| `FIELDS_UPDATE` | WARN | 有 TOC/REF 域但未设 `w:updateFields`(域结果可能过期) |
| `FIRSTLINE_FIXED` | WARN | 首行缩进用固定 twip 而非字符(`firstLineChars`) |
| `PACKAGE_INTEGRITY` | FAIL | zip 成员损坏或缺少必需包部件 |

它刻意保持**领域中立**——不写死任何学科专有符号表。

**审计范围:** 脚本检查 `word/document.xml` 主文档流(另读 `word/styles.xml` 核查表格样式底纹/边框、`word/settings.xml` 核查域刷新),不检查页眉、页脚、脚注、尾注、批注或嵌入部件——这些需人工或更深的 OOXML 检查。表格单元格内的文字查字号、边框、底纹和纯文本公式(WARN 级),不查标点/标题(以避免单元格里的数字和短片段造成误报)。

## 自动修复脚本

`scripts/normalize_docx.py` 机械修复安全问题,其余字节原样保留。默认开启:全角引用括号(`［1］` → `[1]`)和 CJK 之间误用的 ASCII 句读(`中文,中文` → `中文，中文`)。可选开关:`--units`(`20km` → `20 km`、`50 %` → `50%`)、`--tables`(清表内底纹、归零 `w:tblLook`、表头行跨页重复)、`--update-fields`(让 MS Word 打开时自动刷新 TOC/REF 域)、`--all`:

```text
python scripts/normalize_docx.py report.docx -o report.fixed.docx
python scripts/normalize_docx.py report.docx --all --in-place
```

它**不会**改字体、样式或交叉引用——那些需要判断,留给模型 + 主规范;公式用下面的专用脚本。

## 交付门禁

`scripts/finalize_docx.py` 是交付前必须运行的唯一命令:原地应用全部安全机械修复、跑完整审计、给出 `DELIVERY GATE: PASS/FAIL` 判定(FAIL 时退出码非零)。所有 Must-Fix 规则都是 FAIL 级,门禁 FAIL 就不许交付:

```text
python scripts/finalize_docx.py report.docx            # 修复 + 审计 + 判定
python scripts/finalize_docx.py report.docx --no-fix   # 只审计不改文件
```

## 公式替换脚本

`scripts/replace_math.py` 是公式规则背后的确定性工具:把 LaTeX 批量转换为原生 OMML(Pandoc),并把每条公式**拼接到纯文本原文的确切位置**——跨 run 的 token 也能处理,前后字符原样保留,自动盖上字号(小四/五号),display 公式自动排成"居中制表位 + 右对齐编号"版式。输出的 JSON 摘要里 `not_found` 与 `still_plain_text` 必须为空:

```text
python scripts/replace_math.py report.docx registry.json --in-place
python scripts/replace_math.py --convert 'Q = \\frac{W}{T}'   # 打印单条公式的 OMML
```

## 测试

仅用标准库(无第三方依赖):

```text
python -m unittest discover -s tests -v
```

每次推送时 CI 会在 Python 3.9 和 3.12 上跑 `py_compile` 加测试套件(见 `.github/workflows/ci.yml`)。

## 目录结构

```text
.
├── SKILL.md
├── LICENSE
├── CHANGELOG.md
├── agents/
│   └── openai.yaml
├── references/
│   ├── formatting-standard.md                  # 主清单
│   ├── document-structure-and-page-setup.md    # 结构、页面设置、目录、字号
│   ├── latex-omml-formula-workflow.md          # LaTeX → Word OMML 工作流
│   ├── reference-style-gbt7714.md              # GB/T 7714—2015 参考文献格式
│   ├── citation-crossrefs-ooxml.md             # 正文 REF 交叉引用 OOXML
│   └── three-line-table-ooxml.md               # 三线表 OOXML 实现配方
├── scripts/
│   ├── audit_docx_format.py                    # 只读护栏
│   ├── finalize_docx.py                        # 交付门禁:修复+审计+判定
│   ├── normalize_docx.py                       # 安全自动修复
│   └── replace_math.py                         # LaTeX→OMML 公式原位替换
├── examples/
│   ├── README.md
│   ├── make_sample.py                          # 生成合规 / 不合规样例
│   ├── sample-audit-output.txt
│   └── sample-compliant-audit-output.txt
└── tests/
    ├── test_audit_docx_format.py
    └── test_normalize_docx.py
```

版本历史见 [`CHANGELOG.md`](CHANGELOG.md)。

## 维护说明

- 保持 `SKILL.md` 精简,触发时加载更快。
- 详细规则放在 `references/`。
- 审计器保持保守:`FAIL` 只给机器可确定的违规,`WARN` 给需人工复核的项。
- 不要在本仓库存放私密信息、凭据或一次性聊天记录。
- 规范变更时,更新对应的参考文件,并同步刷新 `SKILL.md` 里的摘要和路径。

## 许可

[MIT](LICENSE)。
