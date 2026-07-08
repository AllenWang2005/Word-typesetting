# 三线表实现指南（OOXML）

用于把表格做成规范的**三线表**。只描述"线宽 1.5/0.75 磅"不够——必须明确**哪条边设线、哪条边设无、用什么数值**，否则模型容易套用带网格或带粗标题线的表格样式，排出"上下无线、中间一条粗线"的错误结果。三线表同时必须是**纯白底**：任何单元格（尤其是表头行）都不能有底纹。

## 三条线到底是哪三条

一张三线表**只有三条横线**，从上到下：

1. **表顶线**：整张表最上沿。**粗**（约 1.5 磅）。
2. **表头下线（栏目线）**：表头行与表体之间的那**一条**线。**细**（约 0.75 磅）。← 唯一的中间线，必须是细的。
3. **表底线**：整张表最下沿。**粗**（约 1.5 磅）。

除此之外**全部设为无**：所有竖线（left/right/insideV）、表体各行之间的横线（insideH）都不画。

> 常见错误：表顶线/表底线缺失、中间那条反而最粗——这通常是**套用了"网格型表格 / 带标题行底纹的表格样式"**，或把 `insideH`（行间横线）设成了粗线而没设 top/bottom。务必按下面的方式**显式**设边框，不要依赖内置带框样式。

## 线宽数值（关键）

OOXML 里边框宽度 `w:sz` 的单位是**八分之一磅**（eighths of a point）：

| 目标线宽 | `w:sz` 值 |
| --- | --- |
| 1.5 磅（顶线 / 底线，较粗） | `12` |
| 0.75 磅（表头下线，较细） | `6` |
| 0.5 磅（更细的备选） | `4` |

所以：**顶线、底线 `w:sz="12"`；表头下线 `w:sz="6"`**。中间线一定要比上下线细。

## 推荐实现

第一步：在表级 `w:tblPr/w:tblBorders` 里把整张表设成"只有粗的上下线、其余无"：

```xml
<w:tblBorders>
  <w:top    w:val="single" w:sz="12" w:space="0" w:color="000000"/>
  <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>
  <w:left     w:val="none"/>
  <w:right    w:val="none"/>
  <w:insideH  w:val="none"/>
  <w:insideV  w:val="none"/>
</w:tblBorders>
```

第二步：给**表头行**的每个单元格 `w:tcPr/w:tcBorders` 加一条细的下边线，作为表头下线：

```xml
<w:tcBorders>
  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
</w:tcBorders>
```

结果：粗顶线（表顶）+ 细中线（表头下）+ 粗底线（表底），无竖线、无行间线。

## 白底：底纹必须显式清掉（关键）

三线表必须**全表白底**。最常见的"表头带蓝/灰底"不是谁手动填的色，而是**表格样式的条件格式**（`firstRow` 标题行底纹）带进来的——`word/document.xml` 里往往看不到任何 `w:shd`，底纹藏在 `styles.xml` 的表格样式定义里。所以要做三件事：

1. **不要引用带格式的表格样式**：删掉 `w:tblPr/w:tblStyle`，或将其指向素样式（`TableNormal`）。不要用 `Table Grid`、`Grid Table 4 Accent 1`、"浅色网格－强调文字颜色1"这类内置样式。
2. **关掉条件格式开关**：把 `w:tblPr/w:tblLook` 的 `firstRow`/`lastRow`/`firstColumn`/`lastColumn`/`noHBand`/`noVBand` 全部置 0（或 `w:val="0000"`），防止残留的样式条件格式生效。
3. **逐单元格显式清底纹**：删除所有 `w:tcPr/w:shd`、`w:tblPr/w:shd`，或显式写成无底纹：

```xml
<w:shd w:val="clear" w:color="auto" w:fill="auto"/>
```

`w:fill` 只允许 `auto` 或 `FFFFFF`；`w:val` 只允许 `clear`/`nil`。任何其他 `fill` 色值或 `pct10` 之类的图案值都算带底纹，必须清除。python-docx 里同样没有底纹 API，用 `tc.tcPr` 注入/删除 `w:shd` 节点即可。

> 审计脚本会对文档直接设置的底纹和表格样式（含 `basedOn` 继承链、`tblStylePr` 条件格式）里带来的底纹报 `TABLE_SHADING`（FAIL）。

## python-docx 提示

python-docx 没有边框 API，用底层 XML 设置。思路：

- 先清掉默认样式带来的边框（不要用 `table.style = 'Table Grid'` 这类带网格的样式）。
- 给 `table._tbl` 注入上面的 `w:tblBorders`；
- 给表头行各单元格的 `tc.tcPr` 注入 `w:tcBorders/w:bottom`（`w:sz="6"`）。

可用 `from docx.oxml.ns import qn` 和 `from docx.oxml import OxmlElement` 构造这些节点。

## 交付前自查

- 表**最上沿和最下沿各有一条粗线**（约 1.5 磅），且清晰可见。
- 表头与表体之间**只有一条细线**（约 0.75 磅），且比上下线细。
- **没有任何竖线**，表体行与行之间**没有横线**。
- **全表白底**：表头行和所有单元格没有任何底纹（含表格样式条件格式带来的底纹）。
- 不是套用的网格 / 带标题底纹样式，而是显式设的边框 + 显式清的底纹。
- 审计脚本：`TABLE_BORDERS`（竖线/内部网格，含表格样式从 styles.xml 带来的网格）、`TABLE_SHADING`（底纹）、`TABLE_RULES`（顶/底线缺失、行间横线、中间线不比上下线细）都是 **FAIL 级**、直接拦截交付；最终在 MS Word 中打开目视复核。
