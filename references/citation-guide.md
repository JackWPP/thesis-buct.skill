# 参考文献管理指南

## 1. 核心原则

- 参考文献按**文中首次出现顺序**连续编号（[1][2][3]...）
- 参考文献**主要集中在综述性章节**（国内外研究现状、相关工作等），其他章节少量引用
- 文末参考文献列表按编号顺序排列，与文中引用完全对应
- 文中引用格式：右上角上标方括号，如 `XXXX[1]` 或 `XXXX[1,2]`

---

## 2. references.json 完整格式

```json
{
  "refs": [
    {
      "id": "ref_001",
      "cite_key": 1,
      "type": "journal",
      "authors": "Yongjian A, Xie R, Xiong J, et al.",
      "title": "Microfluidics for Bio-Synthesizing",
      "journal": "Small",
      "year": 2019,
      "volume": "16",
      "issue": "9",
      "pages": "",
      "doi": "",
      "first_appearance_chapter": 1,
      "first_appearance_section": "1.2",
      "formatted": "[1] Yongjian A, Xie R, Xiong J, et al. Microfluidics for Bio-Synthesizing: From Droplets and Vesicles to Artificial Cells[J]. Small, 2019, 16(9)."
    },
    {
      "id": "ref_002",
      "cite_key": 2,
      "type": "thesis",
      "authors": "宗薇",
      "title": "基于磷脂的人造细胞构建及其功能研究",
      "university": "哈尔滨工业大学",
      "year": 2018,
      "first_appearance_chapter": 1,
      "first_appearance_section": "1.2",
      "formatted": "[2] 宗薇. 基于磷脂的人造细胞构建及其功能研究[D]. 哈尔滨工业大学, 2018."
    }
  ]
}
```

**type 枚举**：`journal`（期刊）, `book`（专著）, `thesis`（学位论文）, `conference`（会议论文集）, `patent`（专利）, `standard`（技术标准）, `newspaper`（报纸）, `report`（科技报告）, `electronic`（电子文献）

---

## 3. 各类文献格式模板

### 3.1 期刊 [J]
```
[n] 作者. 文献名称[J]. 期刊名称, 年, 卷(期): 页码范围
```
示例：
```
[1] XXX, XXX, XXX, 等. 一种用于在线检测局部放电的数字滤波技术[J]. 清华大学学报(自然科学版), 1993, 33(4): 62-67.
```

### 3.2 专著 [M]
```
[n] 作者. 专著名称[M]. 版本(第1版不注). 出版地: 出版者, 出版年. 参考页码
```
示例：
```
[2] XXX, XXX, XXX. 图书馆目录[M]. 北京: 高等教育出版社, 1957. 15-18.
```

### 3.3 学位论文 [D]
```
[n] 作者. 题目[D]. 保存地点: 保存单位, 年份
```
示例：
```
[3] XXX. 微分半动力系统的不变集[D]. 北京: 北京大学数学系数学研究所, 1983.
```

### 3.4 会议论文集析出 [A][C]
```
[n] 作者. 论文题目[A]. 见: 主编. 论文集名[C]. 出版地: 出版者, 出版年, 页码范围
```

### 3.5 专利 [P]
```
[n] 专利发明者. 题目[P]. 国别, 专利号. 批准日期
```

### 3.6 技术标准 [S]
```
[n] 标准代号和标准号. 标准名称[S]. 出版年
```

### 3.7 电子文献 [EB/OL] [DB/CD]
```
[n] 主要责任者. 电子文献题名[文献类型/载体类型]. 电子文献地址, 发表或更新日期/引用日期
```

---

## 4. 作者姓名规则

- 作者数量 ≤ 3 人：全部列出
- 作者数量 > 3 人：只列前 3 位，其后加"等"（中文）或"et al."（英文）
- 作者姓名：姓在前，名在后；名可用缩写字母，缩写名后**不加点号**
- 作者间用逗号分开
- 年份优先用近 5 年文献，远年份不超过总文献量的 20%

---

## 5. 写作流程中的引用工作流

### Step 1：写作时临时标注
在 Markdown 写作时，使用 `[^ref_001]` 形式的 ID 引用（对应 references.json 中的 id 字段），不必关心最终编号：

```markdown
该方法在早期研究中已得到验证[^ref_001]，后续工作进一步拓展了其应用范围[^ref_002][^ref_003]。
```

### Step 2：完成全稿后排序

运行引用排序脚本（或手动操作）：
1. 扫描所有 `chapters/*.md`，按文件顺序（章节顺序）找出每个 `[^ref_xxx]` 的**首次出现位置**
2. 按首次出现顺序为每个 ref 分配 `cite_key`（1, 2, 3...）
3. 更新 `references.json` 中的 `cite_key` 和 `first_appearance_*` 字段
4. 构建脚本将 `[^ref_xxx]` 替换为 `[cite_key]` 上标

### Step 3：生成参考文献列表

构建脚本按 `cite_key` 升序读取 `references.json`，生成"参考文献"章节。

---

## 6. 参考文献章节格式

```
参考文献（黑体三号，居中，段前0段后2行）

[1] ...（小四宋体，两端对齐，固定20磅，悬挂缩进：序号占位后正文对齐）
[2] ...
```

docx.js 实现：
```javascript
// 参考文献条目（悬挂缩进，让正文行与第一行对齐）
new Paragraph({
  alignment: AlignmentType.BOTH,
  spacing: { line: 400, lineRule: "exact", before: 0, after: 0 },
  indent: { left: 480, hanging: 480 },  // 悬挂缩进 = 序号宽度
  children: [
    new TextRun({ text: "[1] ", font: { name: "Times New Roman" }, size: 24 }),
    new TextRun({ text: "Yongjian A, Xie R, ...", font: { name: "Times New Roman" }, size: 24 }),
    // 中文部分用宋体 TextRun
  ]
})
```

---

## 7. 文献综述写作建议

当用户提供文献综述草稿和参考文献列表时：
1. 先读取文献列表，构建 `references.json`（`cite_key` 暂时按列表顺序）
2. 将综述内容转化为符合格式的 Markdown，将文献引用替换为 `[^ref_xxx]` 标注
3. 完成后进行排序检查，确保引用顺序与列表一致（综述性章节通常顺序较规整）
4. 若有提供文献原文，可调用 Agent 子任务对综述内容进行事实核验和丰富
