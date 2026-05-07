---
name: thesis-buct
description: "北京化工大学毕业论文写作与排版技能。当用户需要写作毕业论文、按章节生成论文 Markdown、管理图表注册表、插入参考文献、或将论文 Markdown 构建为符合 BUCT 模板的 DOCX 时，必须使用此技能。触发关键词：毕业论文、BUCT论文、论文排版、论文写作、图表注册表、续表、参考文献插入、论文章节。即使用户只是说「帮我写论文第X章」或「生成图表注册表」也应立即触发本技能。"
---

# 北京化工大学毕业论文写作技能 (BUCT Thesis Skill)

## 概览

本技能覆盖完整的论文写作→排版流水线：

```
[用户提供图表规划 + 文献列表]
        ↓
  Phase 1: 初始化工作区
  生成 figure_table_registry.json
        ↓
  Phase 2: 逐章写作 (Markdown)
  按标准格式写每章，引用注册表中的图/表占位符
        ↓
  Phase 3: 参考文献处理
  维护 references.json，确保文中引用顺序一致
        ↓
  Phase 4: 构建 DOCX
  运行 build_thesis.js 生成符合模板的 Word 文档
```

**工作目录约定**（每次使用此技能都在此目录操作）：
```
thesis_workspace/
├── chapters/           # 每章 Markdown
│   ├── 00_abstract.md
│   ├── 01_intro.md
│   └── ...
├── figure_table_registry.json   # 图表主注册表
├── references.json              # 参考文献库
├── figures/            # 图片文件（PNG/JPG）
└── output/             # 生成的 DOCX
```

---

## Phase 1：初始化工作区

### 1.1 读取图表规划

用户提供图表规划（Markdown 格式），例如：
```markdown
## 第1章 绪论
- 图1-1: 系统整体架构图 | system_architecture.png
- 图1-2: 研究流程图 | research_flow.png

## 第2章 相关技术
- 表2-1: 主流框架对比 | 3列×5行
- 图2-1: 模型结构示意图 | model_structure.png
```

### 1.2 生成图表注册表

将规划转换为 `figure_table_registry.json`，此文件是写作和 DOCX 构建的单一事实来源：

```json
{
  "figures": {
    "fig_1_1": {
      "label": "图1-1",
      "caption_zh": "系统整体架构图",
      "caption_en": "Overall System Architecture",
      "file": "figures/system_architecture.png",
      "chapter": 1,
      "planned_after": "介绍完系统设计思路后",
      "status": "pending"
    }
  },
  "tables": {
    "tab_2_1": {
      "label": "表2-1",
      "caption_zh": "主流框架对比",
      "caption_en": "Comparison of Mainstream Frameworks",
      "chapter": 2,
      "columns": ["框架名称", "语言", "性能", "社区活跃度", "适用场景"],
      "planned_after": "介绍各框架特点之后",
      "status": "pending"
    }
  }
}
```

**status** 字段：`pending`（规划中）→ `draft`（已在 Markdown 中插入）→ `file_ready`（图片已就位）→ `built`（已入 DOCX）

---

## Phase 2：逐章写作（Markdown 格式规范）

> **详细格式规范见** `references/formatting.md`

### 2.1 Markdown 章节结构

```markdown
<!-- CHAPTER: 1 -->
<!-- TITLE: 绪论 -->

# 第1章 绪论

## 1.1 研究背景

正文段落使用普通段落写作。首行缩进由构建脚本自动添加，不需要手动空格。
中文正文用宋体小四，英文/数字用Times New Roman小四，由构建脚本自动处理字体切换。

引用文献在行内用 `[^1]` 标注（上标方括号），如"该技术已被广泛验证[^1]"。
多篇文献连续引用写作 `[^1][^2]` 或 `[^1,2]`。

<!-- FIGURE: fig_1_1 -->
<!-- 此占位符告诉构建脚本在此处插入图1-1 -->

## 1.2 研究现状

<!-- FIGURE: fig_1_2 -->

## 1.3 本文主要工作

<!-- TABLE: tab_1_1 -->
```

### 2.2 图、表、公式占位符语法

| 类型 | Markdown 占位符 | 说明 |
|------|----------------|------|
| 图 | `<!-- FIGURE: fig_1_1 -->` | 从注册表读取 caption |
| 表 | `<!-- TABLE: tab_2_1 -->` | 从注册表读取结构 |
| 表格数据 | 见下方 | 在注册表中或 Markdown 内嵌 |
| 公式 | `<!-- EQ: 1-1 -->` + LaTeX | 居中排版，自动编号 |

**表格数据内嵌方式**（如果表格内容在写作时已知）：
```markdown
<!-- TABLE: tab_2_1
| 框架名称 | 语言 | 性能 | 社区活跃度 | 适用场景 |
|---------|------|------|-----------|---------|
| PyTorch | Python | 高 | 极高 | 研究 |
| TensorFlow | Python | 高 | 高 | 生产 |
-->
```

### 2.3 写作时的注意事项

- **参考文献仅在"国内外研究现状"等综述性章节大量出现**，其他章节少量引用即可
- 每个引用 `[^n]` 的 `n` 在整篇文章中连续递增，不能跨章节重置
- 标题层级：H1=章，H2=节（1.1），H3=小节（1.1.1），H4=条（1.1.1.1）
- 不要在 Markdown 中手动加首行空格

---

## Phase 3：参考文献管理

> **详细说明见** `references/citation-guide.md`

### 3.1 references.json 格式

```json
{
  "refs": [
    {
      "id": "ref_001",
      "cite_key": 1,
      "type": "journal",
      "authors": "Yongjian A, Xie R, Xiong J, et al.",
      "title": "Microfluidics for Bio-Synthesizing: From Droplets and Vesicles to Artificial Cells",
      "journal": "Small",
      "year": 2019,
      "volume": 16,
      "issue": 9,
      "first_appearance_chapter": 1,
      "first_appearance_section": "1.2"
    }
  ]
}
```

### 3.2 引用顺序维护

在写完所有章节后，运行顺序检查：检查各章 Markdown 中 `[^n]` 的实际出现顺序，确保 `cite_key` 按文中首次出现顺序从 1 开始连续编号。如有乱序，更新 `references.json` 中的 `cite_key` 并同步更新 Markdown。

---

## Phase 4：构建 DOCX（三阶段管线）

> **环境要求**：Node.js (npm docx), Python 3 (lxml, Pillow), Pandoc (用于公式 OMML 转换)

### 4.1 三阶段构建流程

```
build_thesis.js          inject_omml.py           fix_crossrefs.py
(docx npm 格式化)  →   (Pandoc 公式 OMML)   →   (REF 域代码交叉引用)
       ↓                       ↓                         ↓
 thesis_styled.docx     thesis_omml.docx          thesis_final.docx
```

### 4.2 执行构建

```bash
cd thesis_workspace
npm install docx                                          # 一次性
node build_thesis.js --workspace . --output ./output/styled.docx
python inject_omml.py                                     # 公式 → OMML
python fix_crossrefs.py                                   # 引用 → REF 域代码
# 最终输出：output/thesis_final.docx（或 thesis_crossref.docx）
```

### 4.3 各阶段职责

**Stage 1 — build_thesis.js (Node.js)**
- 按章顺序读取 `chapters/*.md`，调用 `docx` npm 包构建 DOCX
- 解析 `<!-- FIGURE/TABLE -->` 占位符，嵌入实际图片和数据表格
- `[^ref_RXX]` 引用 → `InternalHyperlink` 上标（Stage 3 转为 REF 域代码）
- `$...$` 内联公式 → 隐藏标记 + 灰色 italic fallback（Stage 2 转为 OMML）
- `$$...$$` + `<!-- EQ: X-Y -->` → 公式段落标记（Stage 2 转为 OMML）
- `**text**` 加粗解析（含递归内联公式处理）
- `- item` 列表解析
- BUCT 格式：宋体/黑体/TNR 字体，A4 页面，2.5/2.7cm 边距，`atLeast` 22pt 行距
- 图题/表题 中英双行 五号，图片 AUTO 行距（避免裁剪）
- 参考文献列表 Bookmark + 手动编号

**Stage 2 — inject_omml.py (Python + Pandoc)**
- 读取 `formulas_map.json`，查找 DOCX 中所有 `[[FRM_X_Y]]` / `[[INL_N]]` 标记
- 每个公式：Pandoc `$$LaTeX$$` → 原生 OMML (`m:oMath`)　
- 块级公式：居中段落 + `\t（3-2）` RIGHT tab stop 编号
- 内联公式：替换标记 run + 删除灰色 fallback run，插入 OMML inline
- 清理所有残留标记和灰色 fallback 文本

**Stage 3 — fix_crossrefs.py (Python)**
- 添加 `word/numbering.xml` 定义自动编号列表
- 参考文献条目 → `w:numPr` 自动编号（替换手动 `[n]` 文本）
- 文内引用 `w:hyperlink` → `REF ref_RXX \n \h` Word 原生交叉引用域代码
- 用户 Word 中 Ctrl+点击即可跳转到参考文献条目

---

## 构建经验总结（踩坑记录）

### 行距：必须用 `atLeast` 而非 `exact`
- `lineRule: 'exact'` 会裁剪超行高的公式和图片
- 模板使用 `lineRule: 'atLeast'`，`line: 440`（22pt）
- 图片段落使用默认 AUTO 行距（不设置 `lineRule`）

### 公式渲染：Pandoc 是唯一可靠的 LaTeX→OMML 方案
- `docx` npm 包不支持 OMML
- Pandoc 的 `--from markdown+tex_math_dollars --to docx` 能将 `$$...$$` 转为原生 OMML
- **多行 `$$` 块**：需在解析器中累积所有行直到闭合 `$$`，不能被单行匹配漏掉
- **内联 `$...$`**：需单独处理，同样通过 Pandoc stdin 转换
- **加粗文本内的公式**：`**$...$**` 中的 `$...$` 会被加粗正则吞掉，必须在加粗 handler 内递归处理内联公式

### 交叉引用：Word 原生 REF 域代码 > w:hyperlink
- `w:hyperlink` + `w:bookmark` 在某些 Word 版本中不可点击
- 正确做法：参考文献用 `w:numPr` 自动编号列表 + 文内用 `REF ref_RXX \n \h` 域代码
- 需添加 `word/numbering.xml` 定义编号格式

### 内联公式双重渲染
- OMML 注入时必须**同时删除**标记 run 和灰色 fallback run
- 删除操作必须在树修改**之前**定位目标 run（修改后索引会失效）
- 使用最终清理 pass 处理遗漏标记和灰色文本

### Markdown 写作规范
- 引用使用 `[^ref_RXX]` id 格式（非 cite_key），构建时自动编号
- 公式使用 `$$...$$` + `<!-- EQ: X-Y -->` 块级，`$...$` 内联
- 每个公式必须有 `<!-- EQ: X-Y -->` 标记，否则不编号
- 列表使用 `- item`（已支持解析）
- 不在正文中出现"参考文献分类统计"等元数据表格

---

## 快速参考：关键格式数据

| 元素 | 中文字体 | 英文字体 | 字号 | 行距 | 对齐 | 段前/后 |
|------|---------|---------|------|------|------|---------|
| 正文 | 宋体 | Times New Roman | 小四(12pt) | 最小22磅(atLeast) | 两端 | 0/0 |
| 章标题(H1) | 黑体 | Times New Roman | 三号(16pt) | 最小22磅 | 居中 | 0/2行 |
| 节标题(H2) | 黑体 | Times New Roman | 四号(14pt) | 最小22磅 | 左 | 0/1行 |
| 小节标题(H3) | 黑体 | Times New Roman | 小四(12pt) | 最小22磅 | 左 | 0/1行 |
| 条标题(H4) | 黑体 | Times New Roman | 小四(12pt) | 最小22磅 | 左 | 0/0 |
| 图题/表题 | 宋体 | Times New Roman | 五号(10.5pt) | 固定20磅 | 居中 | 0/0 |
| 参考文献条目 | 宋体 | Times New Roman | 小四(12pt) | 固定20磅 | 两端 | 0/0 |

**页面参数（A4）**：上下边距 2.5cm (1418 DXA)，左右边距 2.7cm (1531 DXA)，页眉 1.4cm，页脚 2cm
