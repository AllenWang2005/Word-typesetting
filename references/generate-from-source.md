# 从源生成 DOCX（source-first 工作流）

适用于**新建文档**或**需要大改重排**的文档。核心思想：不在别人的 OOXML 里做手术，而是先把内容整理成"Markdown 正文 + LaTeX 公式"的源文件，再一次性编译成 DOCX——公式、样式在编译时**天生正确**，不存在"追加到段尾""整段斜体"这类手术事故。编辑既有文档的小改动仍走编辑路径（`replace_math.py` + 审计器）。

## 基本管线

```bash
# 1. 编译：Markdown（含 $...$ / $$...$$ 数学）→ DOCX
pandoc source.md -o report.docx --reference-doc=reference.docx

# 2. 机械修复：表格白底/表头重复/域刷新/数字-单位空格
python scripts/normalize_docx.py report.docx --all --in-place

# 3. 审计
python scripts/audit_docx_format.py report.docx
```

- `--reference-doc=reference.docx` 是关键：Pandoc 从这个参考文档继承所有样式（正文宋体小四、标题黑体、题注五号……）。参考文档做一次、反复使用：拿任何一份审计通过的合格报告，删掉正文只留样式即可充当 reference.docx。
- Markdown 里的 `$F = 44.5\,\mathrm{km^2}$` 编译后就是带正体标记的原生 OMML；`\mathrm`/`\text` 自动变成 `m:sty val="p"`。
- 三线表边框 Pandoc 默认给的不是三线表，编译后按 `references/three-line-table-ooxml.md` 显式设边框（`normalize_docx.py --tables` 已处理底纹/tblLook/表头重复，边框需另设）。

## 已有公式怎么办：先收割、再重建

**可以。**Pandoc 反向转换会把既有 DOCX 里的 OMML 公式转回 LaTeX：

```bash
pandoc old.docx -t markdown --wrap=none -o harvest.md
```

- `harvest.md` 里原来的 Word 公式对象会变成 `$...$` / `$$...$$` LaTeX——这就是现成的公式清单（formula registry）初稿。
- 已经是纯文本的伪公式（`F=44.5`）不会被识别为数学，收割后仍是普通文本；要人工把它们改写成 LaTeX 并入源文件。
- 反向转换是**有损**的：复杂排版（多级表、题注、域、分节）会退化，收割的产物当"内容底稿"用，不当成品。

两条路线按改动量选：

| 场景 | 路线 |
| --- | --- |
| 新建文档 / 全文重排 | 收割内容 → 整理 Markdown+LaTeX 源 → `pandoc --reference-doc` 编译 → normalize → 审计 |
| 既有文档小修（公式补转、表格修复） | 建公式注册表 → `scripts/replace_math.py` 原位替换 → normalize → 审计 |

## 源文件里的公式写法约定

- 行内量符号：`$Q$`、`$N_p$`、`$R_{\mathrm{SN}}$`。
- 显示公式用 `$$...$$` 独立成段；编号不写进公式，编译后由编辑路径或手工加"居中制表位 + 右制表位 + (3-1)"版式（`replace_math.py` 的 display 模式即此版式）。
- 单位、函数名、说明性下标一律 `\mathrm{}`/`\text{}`；变量裸写。多字母系数写 `C_{\mathrm{I}}`，不写 `CI`。
- 数值与单位之间 `\,`：`43.5\,\mathrm{万\,kW}`、`216\,\mathrm{m^3/s}`。

## 环境

- 需要 Pandoc（macOS：`brew install pandoc`，或从 GitHub releases 下载二进制放入 `~/.local/bin`；脚本会自动在 PATH、`~/.local/bin`、`/opt/homebrew/bin`、`/usr/local/bin` 查找）。
- 渲染验收以 **MS Word** 为准；交付前在 Word 中打开一次目视复核（不做 PDF 渲染批量检查）。
