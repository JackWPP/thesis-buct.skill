# BUCT 毕业论文 DOCX 格式规范参考

## 1. 页面设置

```javascript
// 页面：A4，单面
page: {
  size: { width: 11906, height: 16838 },  // A4 in DXA
  margin: {
    top: 1418,     // 2.5cm
    bottom: 1418,  // 2.5cm
    left: 1531,    // 2.7cm
    right: 1531,   // 2.7cm
    header: 794,   // 1.4cm
    footer: 1134,  // 2.0cm
  }
}
// 版心宽度 = 11906 - 1531 - 1531 = 8844 DXA
const CONTENT_WIDTH = 8844;
```

## 2. 字体规则

- 中文正文：宋体
- 中文标题：黑体
- 英文/数字/符号：Times New Roman（贯穿全文）
- 摘要标题、关键词：黑体
- 图题/表题：五号宋体（中文）+ 五号Times New Roman（英文）

在 docx.js 中为每个 TextRun 显式指定字体：
```javascript
// 正文混排示例
new TextRun({ text: "深度学习", font: { name: "宋体" }, size: 24 })
new TextRun({ text: "Deep Learning", font: { name: "Times New Roman" }, size: 24 })
```

## 3. 段落样式定义（docx.js）

### 3.1 通用行距常量
```javascript
const LINE_SPACING_BODY = { value: 400, rule: LineRuleType.EXACT };  // 20磅
// 20磅 = 20pt = 400 twips (1pt = 20 twips)
```

### 3.2 正文段落
```javascript
// 正文（首行缩进2字符）
{
  alignment: AlignmentType.BOTH,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 0 },
  indent: { firstLine: 480 },  // 2字符 × 12pt × 20 = 480 twips
}
// 字号：sz: 24 (= 12pt, 小四号)
```

### 3.3 章标题（H1 = 第X章）
```javascript
{
  alignment: AlignmentType.CENTER,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 800 },
  // after: 800 = 2行 × 20磅 × 20twips/pt = 800 twips
}
// 字号：sz: 32 (= 16pt, 三号)，黑体，加粗
```

### 3.4 节标题（H2 = 1.x）
```javascript
{
  alignment: AlignmentType.LEFT,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 400 },
  // after: 400 = 1行 × 20磅
}
// 字号：sz: 28 (= 14pt, 四号)，黑体
```

### 3.5 小节标题（H3 = 1.x.x）
```javascript
{
  alignment: AlignmentType.LEFT,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 400 },
}
// 字号：sz: 24 (= 12pt, 小四号)，黑体
```

### 3.6 条标题（H4 = 1.x.x.x）
```javascript
{
  alignment: AlignmentType.LEFT,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 0 },
}
// 字号：sz: 24 (= 12pt, 小四号)，黑体
```

### 3.7 摘要标题
```javascript
{
  alignment: AlignmentType.CENTER,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 400 },
}
// "摘 要"：两字中间空一格，黑体小三(15pt=sz:30)
```

### 3.8 参考文献条目
```javascript
{
  alignment: AlignmentType.BOTH,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 0 },
  indent: { left: 480, hanging: 480 },  // 悬挂缩进，序号对齐
}
// 字号：sz: 24 (= 12pt, 小四号)
```

## 4. 图的排版规范

### 4.1 图片本身（居中）
```javascript
new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: 400, lineRule: "exact", before: 200, after: 0 },
  children: [
    new ImageRun({
      data: imageBuffer,
      transformation: { width: imageWidthEMU, height: imageHeightEMU },
      // EMU: 1 inch = 914400 EMU; 建议图片宽度不超过版心 (8844 DXA = 8844/1440 inches)
      type: "png",  // 或 "jpg"
    })
  ]
})
```

### 4.2 图题（中文行）
```javascript
new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 0 },
  children: [
    new TextRun({ text: "图1-1 ", font: { name: "宋体" }, size: 21 }),       // 五号 = 10.5pt
    new TextRun({ text: "捏合盘元件输送特性", font: { name: "宋体" }, size: 21 }),
  ]
})
```

### 4.3 图题（英文行）
```javascript
new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 200 },  // 图后稍留空
  children: [
    new TextRun({
      text: "Figure 1-1 Conveying Characteristics of the Kneading Disc Element",
      font: { name: "Times New Roman" }, size: 21
    }),
  ]
})
```

**图编号规则**：分章编序，如图1-1, 1-2; 图2-1, 2-2。编号数字用 Times New Roman。

## 5. 表的排版规范

### 5.1 表题（中文，置于表上方）
```javascript
new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: 400, lineRule: "exact", before: 200, after: 0 },
  children: [
    new TextRun({ text: "表1-1 四种不同捏合块构型下流道的物理参数", font: { name: "宋体" }, size: 21 }),
  ]
})
// 紧接英文表题
new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 0 },
  children: [
    new TextRun({ text: "Table 1-1 Physical Parameters...", font: { name: "Times New Roman" }, size: 21 }),
  ]
})
```

### 5.2 表格本身（支持封闭式或三线式）

**封闭式（推荐）**：所有边框均显示
```javascript
const cellBorder = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const allBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder, insideH: cellBorder, insideV: cellBorder };
```

**三线式**：仅顶线、底线、表头底线
```javascript
// 顶线/底线粗（size: 12），表头底线细（size: 6），无内部竖线
const thickLine = { style: BorderStyle.SINGLE, size: 12, color: "000000" };
const thinLine  = { style: BorderStyle.SINGLE, size: 6,  color: "000000" };
const noLine    = { style: BorderStyle.NONE };
```

**表格宽度**：
```javascript
new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },  // 8844 DXA
  columnWidths: [...],  // 各列 DXA 之和 = 8844
  alignment: AlignmentType.CENTER,
  // ...
})
```

**表格内文字**：五号宋体（sz:21 = 10.5pt），Times New Roman for English

### 5.3 续表处理

当表格跨页时，需要在新页顶部重复表头并显示"续表"：

**策略**：在 docx.js 中，对超长表格，在表格 `TableRow` 上设置 `tableHeader: true`（使表头行自动重复），并手动在每个续页前插入"续表"标注段落。

```javascript
// 表头行设置
new TableRow({
  tableHeader: true,  // 跨页重复此行
  children: [/* 表头单元格 */]
})
```

**续表标注段落**（在 Table 之前插入，仅当检测到表格会跨页时）：
```javascript
// 续表标注右对齐
new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 0 },
  children: [new TextRun({ text: "续表", font: { name: "宋体" }, size: 21 })]
})
```

**实践建议**：构建脚本将表格拆分为"首页部分"和"续页部分"（`<!-- TABLE_CONTINUED: tab_x_x -->`），并在续页部分前自动插入续表标注和表头。

## 6. 参考文献上标引用

在文中，引用标注为右上角方括号上标：

```javascript
// 正文段落中的引用上标
new TextRun({ text: "该方法已被广泛验证" }),
new TextRun({
  text: "[1]",
  superScript: true,
  font: { name: "Times New Roman" }, size: 24
}),
```

多个引用：
- 连续：`[1,2,3]` 或 `[1-3]`
- 非连续：`[1,3,5]`

## 7. 特殊页面元素

### 7.1 页眉
```javascript
// 奇偶同页眉：宋体小五（9pt）居中，内容为"北京化工大学毕业论文（设计）"
```

### 7.2 页码
```javascript
// 页脚居中，阿拉伯数字，宋体小五（9pt）
// 正文从第1章开始编页码，封面/声明/任务书不编
```

### 7.3 目录
```javascript
// 目录标题：黑体三号居中，段前0行段后2行
// 目录条目：四号黑体（一级），小四黑体（二级及以下）
// 页码：宋体小四，右对齐，用前导符"……"
// 段前段后0，固定值20磅
```

## 8. 公式排版

公式居中，编号右对齐：

```javascript
// 使用 Tab Stop 实现居中+右对齐编号
new Paragraph({
  tabStops: [
    { type: TabStopType.CENTER, position: CONTENT_WIDTH / 2 },
    { type: TabStopType.RIGHT, position: CONTENT_WIDTH }
  ],
  children: [
    new TextRun({ text: "\t" }),          // tab to center
    new TextRun({ text: "E = mc²" }),     // 公式内容（建议用公式图片）
    new TextRun({ text: "\t（1-1）" }),   // tab to right + 编号
  ]
})
```

## 9. 模板文件

`assets/template.docx` 是北京化工大学官方模板（2025年单面版）。

**使用策略**：用 `scripts/office/unpack.py` 解包模板，保留其样式定义（styles.xml），然后向 document.xml 中写入内容。这样可继承模板中的所有预定义样式，比从头创建更可靠。

```bash
# 解包模板
python3 /path/to/scripts/office/unpack.py assets/template.docx output_unpacked/
# 编辑 output_unpacked/word/document.xml 写入内容
# 打包
python3 /path/to/scripts/office/pack.py output_unpacked/ output/thesis_final.docx --original assets/template.docx
```
