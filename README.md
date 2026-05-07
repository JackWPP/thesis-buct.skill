# thesis-buct.skill

<p align="center">
  <img src="./assets/readme-hero.png" alt="北京化工大学毕业论文构建工具链插图" width="100%">
</p>

> 一个面向北京化工大学毕业论文场景的技能包：把分章 Markdown、图表注册表、参考文献和公式处理流程，组织成可落地的 `DOCX` 构建管线。

## 这是什么

这个仓库不是通用的论文模板集合，而是一个更聚焦的工具包：

- `SKILL.md` 定义了面向 Codex 的论文写作/排版工作流。
- `scripts/` 提供从 Markdown 到 Word 的构建脚本。
- `references/` 汇总写作格式、图表注册表、参考文献管理约定。
- `assets/template.docx` 提供模板资产，便于和学校格式要求对齐。

如果你要做的是：

- 按章节维护论文 Markdown
- 管理图、表、公式、参考文献编号
- 输出更接近学校要求的 `.docx`
- 尽量减少手工调 Word 的时间

那么这个项目就是为这类任务准备的。

## 项目解决的问题

普通的 Markdown 导出 Word，通常会在下面几件事上失真：

- 中文/英文混排字体不统一
- 图表题注和编号容易手工漂移
- 公式很难稳定转成 Word 原生公式
- 文内引用和文末参考文献难以保持一致
- 交叉引用在 Word 里不够“原生”

这个仓库的思路是把这些问题拆成几层：

1. 写作层：章节内容用 Markdown 维护。
2. 数据层：图表注册表和 `references.json` 作为单一事实源。
3. 构建层：脚本负责版式、图片、表格、公式、引用和交叉引用。
4. 输出层：尽量得到可继续在 Word 中编辑的论文文档。

## 仓库结构

```text
.
├─ SKILL.md
├─ assets/
│  └─ template.docx
├─ references/
│  ├─ citation-guide.md
│  ├─ figure-table-guide.md
│  └─ formatting.md
└─ scripts/
   ├─ build_thesis.js
   ├─ build_thesis.py
   ├─ inject_omml.py
   └─ fix_crossrefs.py
```

## 核心脚本

### `scripts/build_thesis.js`

Node 版主构建器，职责最完整，也最适合当作主入口理解：

- 读取 `chapters/*.md`
- 解析 `<!-- FIGURE: ... -->`、`<!-- TABLE: ... -->`、`<!-- EQ: ... -->`
- 根据 `figure_table_registry.json` 插入图表与题注
- 根据 `references.json` 生成参考文献区
- 先写出样式化 `docx`
- 同时产出 `formulas_map.json`，给后续公式注入脚本使用

它支持通过参数指定工作区，适合从外部目录驱动：

```bash
node ./scripts/build_thesis.js --workspace . --output ./output/thesis_styled.docx
```

### `scripts/inject_omml.py`

把隐藏的公式标记替换成 Word 原生 OMML 公式，覆盖两类场景：

- 行间公式
- 行内公式

脚本内部通过 Pandoc 把 LaTeX 公式转成 OMML，再写回到 `docx` 的 XML 里。

### `scripts/fix_crossrefs.py`

把文内超链接式引用进一步改造成更接近 Word 原生体验的 `REF` 域代码：

- 参考文献条目自动编号
- 文内引用替换为 `REF ... \n \h`
- 在 Word 中更适合继续做交叉引用维护

### `scripts/build_thesis.py`

另一条一体化 Python 管线，思路是：

1. 合并章节 Markdown
2. 交给 Pandoc 先生成基础 `docx`
3. 再对图片、表格、引用、参考文献做后处理

如果你更偏好“单脚本驱动”，它是另一种实现路线。

## 推荐工作区结构

这个仓库本身更像“技能包 + 脚本集”，真实论文通常还需要一个工作目录，例如：

```text
thesis_workspace/
├─ chapters/
│  ├─ 00_abstract.md
│  ├─ 01_intro.md
│  └─ ...
├─ figures/
├─ output/
├─ figure_table_registry.json
├─ references.json
└─ formulas_map.json
```

其中最关键的两个数据文件是：

- `figure_table_registry.json`：图表编号、题注、文件路径、状态
- `references.json`：参考文献条目、首次出现位置、最终编号

## 快速开始

### 1. 准备环境

从当前脚本实现来看，至少需要：

- Node.js
- Python 3
- Pandoc
- Python 包：`lxml`、`Pillow`
- Node 包：`docx`

示例：

```bash
pip install lxml pillow
npm install docx
```

### 2. 准备论文数据

先准备：

- `chapters/*.md`
- `figure_table_registry.json`
- `references.json`
- `figures/` 下的图片文件

格式约定可直接看：

- [`references/formatting.md`](./references/formatting.md)
- [`references/figure-table-guide.md`](./references/figure-table-guide.md)
- [`references/citation-guide.md`](./references/citation-guide.md)

### 3. 运行构建

#### 方案 A：Node 主构建 + Python 后处理

如果你的论文章节和 JSON 数据就在当前工作区根目录，这条路线更容易理解。
但需要注意：当前 Python 后处理脚本默认把“脚本所在目录”当作工作区。
也就是说，直接运行仓库里的 `./scripts/inject_omml.py` 和 `./scripts/fix_crossrefs.py` 时，它们会去找 `./scripts/output/...`。

因此二选一：

- 把这几个 Python 脚本复制到实际论文工作区根目录再运行
- 或者先改脚本里的 `WORKSPACE` 路径逻辑

在“脚本已放到论文工作区根目录”的前提下，命令是：

```bash
node ./scripts/build_thesis.js --workspace . --output ./output/thesis_styled.docx
python ./inject_omml.py
python ./fix_crossrefs.py
```

#### 方案 B：Python 一体化流程

这条路线同样建立在“Python 脚本位于论文工作区根目录”这个假设之上：

```bash
python ./build_thesis.py
```

## 当前状态与已知约束

这个仓库已经把关键思路搭起来了，但从工程化角度看，还有几件事需要使用者留意：

- 目前没有 `package.json` 或 `requirements.txt`，依赖需要手动安装。
- Python 脚本里 Pandoc 路径写死为 `C:/Program Files/Pandoc/pandoc`，跨平台前需要改造。
- Python 管线大量使用“脚本所在目录即工作区”的假设，更适合作为工作区内脚本，而不是直接把本仓库当样例工程运行。
- 仓库没有附带最小可运行示例论文目录，因此第一次接入时最好先搭一个 demo workspace。
- `assets/template.docx` 目前更像是模板资产/参考资产；从当前可见脚本实现来看，并没有直接把它作为主构建入口消费。

这也解释了为什么这个项目更像“论文生产工具箱”，而不是“开箱即用的完整模板仓库”。

## 适合怎么继续演进

如果后面要把它做成更完整的项目，优先级建议是：

1. 增加一个最小示例工作区，降低第一次上手成本。
2. 补 `package.json`、`requirements.txt` 和一键构建命令。
3. 把 Pandoc 路径、工作区路径改成参数化配置。
4. 增加构建校验，提前发现缺图、缺表、缺引用和编号冲突。

## 一句话总结

`thesis-buct` 的价值不在“再造一个 Word 模板”，而在于把论文里最容易失控的编号、题注、公式和引用，收束成一条可重复执行的构建流程。
