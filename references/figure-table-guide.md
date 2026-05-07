# 图表注册表管理指南

## 1. 注册表结构说明

`figure_table_registry.json` 是写作流程的核心数据文件。它：
- 作为图/表编号的**唯一事实来源**，避免编号混乱
- 跟踪每个图/表的插入状态
- 为构建脚本提供图题、文件路径、表格列结构

## 2. 完整 JSON 结构

```json
{
  "_meta": {
    "created": "2025-01-01",
    "author": "学号-班级-姓名",
    "thesis_title": "论文题目"
  },
  "figures": {
    "fig_1_1": {
      "label": "图1-1",
      "caption_zh": "系统整体架构图",
      "caption_en": "Overall System Architecture",
      "file": "figures/system_architecture.png",
      "chapter": 1,
      "planned_section": "1.2",
      "planned_after": "介绍完系统设计思路的第一段之后",
      "width_ratio": 0.7,
      "height_ratio": null,
      "actual_ratio": 0.6,
      "status": "pending",
      "notes": "需要从Visio导出PNG，分辨率300dpi"
    }
  },
  "tables": {
    "tab_2_1": {
      "label": "表2-1",
      "caption_zh": "主流深度学习框架对比",
      "caption_en": "Comparison of Mainstream Deep Learning Frameworks",
      "chapter": 2,
      "planned_section": "2.3",
      "planned_after": "介绍各框架特点之后，对比总结前",
      "style": "closed",
      "columns": ["框架名称", "主要语言", "性能", "社区活跃度", "适用场景"],
      "data": [
        ["PyTorch", "Python", "★★★★★", "极高", "科研/生产"],
        ["TensorFlow", "Python", "★★★★☆", "高", "大规模生产"],
        ["PaddlePaddle", "Python", "★★★★☆", "中", "国内工业"]
      ],
      "status": "draft",
      "notes": "style 可选: closed(封闭式) 或 three_line(三线式)"
    }
  }
}
```

## 3. 图片尺寸规范

### 3.1 推荐尺寸

- `width_ratio`：图片宽度占版心宽度（8844 DXA ≈ 15.6cm）的比例
  - 单栏图：0.7（推荐，约10.9cm）
  - 全宽图：0.95
  - 小图/示意图：0.4~0.5
- `actual_ratio`：图片的高宽比（height/width），用于计算渲染高度
- 构建脚本将根据这两个值自动计算 EMU 尺寸

### 3.2 导出图片建议

- 格式：PNG（透明背景）或 JPG（照片）
- 分辨率：300 DPI（印刷标准）
- 命名：与注册表 `file` 字段一致，存放于 `figures/` 目录

### 3.3 实际尺寸更新（写作完成后）

当图片文件就位后，运行以下命令获取实际尺寸并更新注册表：
```bash
python3 -c "
from PIL import Image
import json, sys
img_path = sys.argv[1]
img = Image.open(img_path)
w, h = img.size
print(f'宽:{w}px 高:{h}px 比例:{h/w:.3f}')
" figures/system_architecture.png
```
将 `actual_ratio` 更新为实际高宽比。

## 4. 表格 data 字段

如果表格数据在规划阶段就已知，可以直接写入 `data` 字段（二维数组）。
如果数据在写作时才确定，可以：
1. 在 Markdown 中用内联注释写入（见 SKILL.md Phase 2.2）
2. 写作完成后回填 `data` 字段

## 5. 续表（跨页表格）处理

### 5.1 续表标准

当表格行数较多（超过约15行），估计会跨页时：
1. 在 `notes` 中标记 `"续表: 可能需要"`
2. 构建时通过设置 `tableHeader: true` 实现表头自动重复
3. 手动在适当位置拆分表格（在注册表中新增 `tab_x_x_continued` 条目）

### 5.2 续表注册表条目示例

```json
"tab_3_1_cont": {
  "label": "表3-1（续）",
  "caption_zh": "（续表，不再重复标题）",
  "caption_en": "",
  "is_continuation_of": "tab_3_1",
  "chapter": 3,
  "columns": ["与主表相同"],
  "data": [["第16行数据..."]],
  "status": "draft"
}
```

构建脚本识别 `is_continuation_of` 字段，在此表格前自动插入右对齐的"续表"标注，并复用主表的表头。

## 6. 状态管理

写作过程中随时更新 `status` 字段：

| status | 含义 | 下一步 |
|--------|------|--------|
| `pending` | 已规划，尚未在 Markdown 中写入占位符 | 在合适位置插入 `<!-- FIGURE: xxx -->` |
| `draft` | 已插入占位符，图片/数据尚未就绪 | 准备图片文件或表格数据 |
| `file_ready` | 图片文件已放置，数据已填写 | 运行构建脚本测试 |
| `built` | 已成功入 DOCX | 检查效果 |

## 7. 图表编号命名规则

- Key 格式：`fig_章号_序号`（如 `fig_1_1`）或 `tab_章号_序号`（如 `tab_2_1`）
- label 格式：`图章号-序号`（如 `图1-1`）、`表章号-序号`（如 `表2-1`）
- 数字部分始终使用 Times New Roman（由构建脚本处理）
- 编号从每章从 1 开始，跨章不重置

## 8. 生成注册表的 Agent 提示词

当用户提供图表规划时，Agent 应：

> 请根据以下图表规划生成 `figure_table_registry.json`。对于每个条目：
> - 中文和英文 caption 都需要填写（英文可以根据中文翻译）
> - file 字段统一使用 `figures/` 前缀加小写下划线文件名
> - width_ratio 默认为 0.7，actual_ratio 可先设为 0.6 等实际图片就位后更新
> - 表格的 columns 字段要列出所有列名，data 可先为 []
> - status 全部初始化为 "pending"
