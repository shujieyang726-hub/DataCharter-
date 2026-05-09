# DataCharter — 数据驱动的智能 PPT 图表生成器

> 从 Excel 数据到专业 PPT 报告，一条命令搞定。

DataCharter 是一个纯 Python 实现的自动化 PPT 生成工具。它能读取 Excel/CSV/JSON 数据文件，**自动推荐最合适的图表类型**，并生成包含**原生可编辑图表**的多页 PowerPoint 报告。所有图表在 PowerPoint 中均可直接双击编辑数据，无需任何图片或截图。

---

## 核心特性

- **数据驱动**：直接读取 Excel (.xlsx) 文件，自动识别章节结构、表格数据和趋势数据
- **智能推荐**：基于数据特征（时间序列、占比、排名、对比等）自动推荐最佳图表类型
- **7 种原生图表【自学习】**：柱状图、横向条形图、折线图、环形图、堆叠柱状图、面积图、雷达图
- **自动分页**：按数据章节分组，每页最多 4 个图表，自动拆分多页
- **图表多样性**：同一页内同类型图表不超过 50%，自动替换为备选类型
- **8 套配色方案**：按页码和图表序号自动轮换，避免视觉单调
- **自学习图表库**：从 PPTX 模板提取图表样式，增量扫描，持久化缓存到图表库
- **报告元信息提取**：自动从 Excel 提取报告标题和摘要作为封面内容
- **跨平台**：支持 Windows / macOS / Linux

---

## 快速开始

### 环境要求

- Python 3.8+
- pip

### 安装依赖

```bash
pip install -r requirements.txt
```

### 一键生成报告

```bash
python scripts/generate_from_template.py --data data/data3.xlsx -o exports/报告.pptx
```

生成的 PPTX 文件包含封面页 + 多页图表页，所有图表均为 PowerPoint 原生对象，可直接编辑。

---

## 项目结构

```
ppt-master-release/
├── scripts/                          # 核心脚本
│   ├── generate_from_template.py     # 主流程：数据解析 → 推荐 → 生成 PPT
│   ├── native_charts.py              # 原生 PowerPoint 图表创建引擎
│   ├── chart_recommender.py          # 数据特征分析与图表类型推荐
│   ├── chart_composer.py             # SVG 图表模块布局组合工具
│   ├── chart_extractor.py            # 图表模块提取工具
│   └── config.py                     # 统一配置（配色方案、尺寸常量等）
├── tests/                            # 单元测试
│   ├── test_chart_composer.py        # 布局组合测试
│   ├── test_chart_recommender.py     # 推荐引擎测试
│   └── test_chart_extractor.py       # 模块提取测试
├── templates/
│   └── charts/                       # 52 个 SVG 图表模板
│       └── charts_index.json         # 模板索引
├── data/                             # 示例数据文件
│   ├── data1.xlsx
│   ├── data2.xlsx
│   └── data3.xlsx
├── exports/                          # 输出目录（含示例输出）
│   └── data3_分析报告.pptx           # 示例生成结果
├── chart_library.json                # 持久化图表样式库（自学习）
├── requirements.txt                  # Python 依赖
└── README.md                         # 本文件
```

---

## 业务流程

PPT Master 的完整流程分为 5 个步骤：

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  Step 1     │     │  Step 2     │     │  Step 3     │     │  Step 4     │     │  Step 5     │
│  加载图表库 │ ──► │  解析数据   │ ──► │  图表推荐   │ ──► │  创建 PPT   │ ──► │  自学习     │
│             │     │             │     │             │     │             │     │             │
│ PPTX模板扫描│     │ Excel解析   │     │ 特征分析    │     │ 多页分页    │     │ 样式提取    │
│ 样式提取    │     │ 章节识别    │     │ 类型匹配    │     │ 原生图表    │     │ 库更新      │
│ 增量缓存    │     │ 元信息提取  │     │ 多样性约束  │     │ 自动布局    │     │             │
└─────────────┘     └─────────────┘     └─────────────┘     └─────────────┘     └─────────────┘
```

### Step 1：加载图表库

从 PPTX 模板文件中提取图表样式（类型、颜色、尺寸），建立持久化图表库。

- **增量扫描**：通过文件签名（`文件名:大小:修改时间`）检测新增/修改的模板
- **持久化缓存**：样式信息存储在 `chart_library.json` 中，无需重复扫描
- **自动发现**：默认扫描当前目录下 `ppt模板*.pptx` 文件

```python
# 文件签名机制
def _file_signature(path: Path) -> str:
    stat = path.stat()
    return f"{path.name}:{stat.st_size}:{int(stat.st_mtime)}"
```

### Step 2：解析数据

通用 Excel 解析器，自动识别数据结构：

| 数据类型 | 识别方式 | 解析结果 |
|---------|---------|---------|
| **报告标题** | 第一行单列文本 | 封面大标题 |
| **报告摘要** | 标题后、第一章节前的文本行 | 封面副标题 |
| **章节标题** | `N. 章节名` 格式（如 `1. 品牌销量汇总`） | 页面标题 |
| **数据表格** | 多列表头 + 数值行 | 柱状图/条形图/环形图 |
| **趋势数据** | `月份\|品牌1\|品牌2\|...` 表头 | 多系列折线图 |
| **键值对数据** | `标签 \| 数值` 两列格式 | 柱状图 |

**智能列角色检测**：根据列名关键词自动判断数据角色：

| 角色 | 关键词 | 对应图表 |
|------|--------|---------|
| composition | 占比、渗透率、份额 | 环形图 |
| ranking | 排名、TOP | 横向条形图 |
| comparison_negative | 差值、变动 | 横向条形图 |
| score | 评分、满意度、指数 | 雷达图 |
| trend | 趋势、月度、走势 | 折线图 |
| bar | 默认 | 柱状图 |

### Step 3：图表推荐

基于决策树的图表类型推荐引擎（`chart_recommender.py`）：

```
                    ┌── 占比数据 ≤6类 ──► 环形图 (95%)
                    │
                    ├── 时间序列
数据特征分析   ──────┤   ├── 单系列 ──► 折线图 (92%)
                    │   ├── 双系列 ──► 双轴折线图 (88%)
                    │   └── 多系列 ──► 堆叠面积图 (82%)
                    │
                    ├── 含负数 ──► 瀑布图 (85%)
                    │
                    ├── 单系列 ≤8类
                    │   ├── 标签长 ──► 横向条形图 (90%)
                    │   └── 标签短 ──► 柱状图 (92%)
                    │
                    └── 多系列 ──► 分组柱状图 (88%)
```

**50% 多样性约束**：同一页面内，同类型图表不超过总数的一半，超出部分自动替换为备选类型。

### Step 4：创建 PPT

使用 `python-pptx` 从零创建 PowerPoint 文件：

1. **封面页**：报告标题（36pt 加粗） + 摘要副标题（14pt 灰色）
2. **自动分页**：按 Excel 章节分组，每页最多 4 个图表
3. **智能布局**：根据图表数量自动选择布局（1列/2列/3列/2×2网格）
4. **原生图表**：7 种图表类型，全部使用 PowerPoint 原生 Chart 对象
5. **图表标题**：每个图表上方显示数据标题
6. **配色轮换**：8 套配色方案按 `(页码 + 图表序号) % 8` 轮换

```python
# 画布尺寸（16:9 标准）
CANVAS_W = 1280  # px
CANVAS_H = 720   # px
PX_TO_EMU = 9525  # 1px = 9525 EMU

# 每页最多图表数
MAX_CHARTS_PER_SLIDE = 4
```

**支持的原生图表类型**：

| 类型 | 函数 | python-pptx 图表类型 |
|------|------|---------------------|
| 柱状图 | `add_native_bar_chart()` | `XL_CHART_TYPE.COLUMN_CLUSTERED` |
| 横向条形图 | `add_native_horizontal_bar_chart()` | `XL_CHART_TYPE.BAR_CLUSTERED` |
| 折线图 | `add_native_line_chart()` | `XL_CHART_TYPE.LINE_MARKERS` |
| 环形图 | `add_native_pie_chart()` | `XL_CHART_TYPE.DOUGHNUT` |
| 堆叠柱状图 | `add_native_stacked_bar_chart()` | `XL_CHART_TYPE.COLUMN_STACKED` |
| 面积图 | `add_native_area_chart()` | `XL_CHART_TYPE.AREA` |
| 雷达图 | `add_native_radar_chart()` | `XL_CHART_TYPE.RADAR` |

### Step 5：自学习

生成 PPT 后，自动从输出文件中提取图表样式，反馈到图表库：

```python
def learn_from_output(pptx_path: Path) -> int:
    """从生成的PPT中学习图表样式，扩充图表库"""
    # 去重：基于 source:slide:shape_name 签名
    # 增量：只添加库中不存在的新样式
```

每次生成都会让图表库更丰富，后续生成的 PPT 有更多样式参考。

---

## 数据格式说明

### Excel 文件结构

PPT Master 期望的 Excel 文件格式：

```
行 1: 报告标题（单列文本）
行 2: 报告摘要（单列文本，可选，可多行）
行 N: 空行
行 N+1: 1. 第一章节标题
行 N+2: 表头行（品牌 | 销量 | 占比 | ...）
行 N+3: 数据行1
行 N+4: 数据行2
...
行 M: 2. 第二章节标题
行 M+1: 月份 | 品牌A | 品牌B | ...
行 M+2: 1月 | 12000 | 8000 | ...
...
```

**章节标题格式**：`数字. 标题文字`（如 `1. 品牌累计销量汇总`）

**支持的数据格式**：

| 格式 | 说明 |
|------|------|
| 排名表格 | `排名 \| 品牌 \| 销量 \| 占比` — 自动识别排名列并跳过 |
| 多列数值表 | `品牌 \| 指标1 \| 指标2` — 每个数值列生成独立图表 |
| 趋势表 | `月份 \| 系列1 \| 系列2` — 生成多系列折线图 |
| 百分比列 | 0~1 之间的小数自动识别为百分比，乘以 100 显示 |

---

## 命令行参数

```bash
python scripts/generate_from_template.py [选项]
```

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--data` | 数据文件路径（必填） | — |
| `--templates` | 模板文件通配符 | `ppt模板*.pptx` |
| `-o, --output` | 输出 PPTX 路径 | `exports/{文件名}_分析报告_{日期}.pptx` |

### 示例

```bash
# 基本用法
python scripts/generate_from_template.py --data data/data1.xlsx

# 指定输出路径
python scripts/generate_from_template.py --data data/data3.xlsx -o exports/新能源报告.pptx

# 指定模板
python scripts/generate_from_template.py --data data/data2.xlsx --templates "templates/*.pptx"
```

---

## 辅助工具

除主流程外，项目还包含以下辅助模块：

### chart_recommender.py — 独立推荐工具

```bash
# 分析数据并输出推荐结果
python scripts/chart_recommender.py data/data1.xlsx

# 输出到 JSON 文件
python scripts/chart_recommender.py data/data3.xlsx -o recommendation.json
```

支持 JSON / CSV / Excel 三种输入格式，输出包含推荐图表类型、置信度和布局建议。

### chart_composer.py — SVG 布局组合

```bash
# 将多个 SVG 图表模块组合成一页
python scripts/chart_composer.py --modules modules/ --layout grid_2x2 -o output.svg
```

支持 6 种布局：`single`、`two_column`、`three_column`、`grid_2x2`、`grid_top2_bottom3`、`grid_2x3`。

### chart_extractor.py — 图表模块提取

从复合 SVG 中提取独立的图表模块，支持基于 `<g>` 元素结构、面积阈值和标签检测的智能分割。

---

## 运行测试

```bash
# 运行所有测试
pytest tests/ -v

# 运行单个测试文件
pytest tests/test_chart_recommender.py -v
```

---

## 配色方案

内置 8 套专业配色方案，按页码自动轮换：

| # | 名称 | 主色调 |
|---|------|--------|
| 1 | 商务经典 | 深蓝 + 金色 + 绿色 |
| 2 | 科技蓝紫 | 紫蓝 + 品红 + 青色 |
| 3 | 自然暖色 | 棕色 + 橄榄绿 + 米色 |
| 4 | 活力对比 | 红色 + 蓝色 + 深蓝 |
| 5 | 深沉专业 | 深青 + 绿色 + 金色 |
| 6 | 清新明亮 | 绿色 + 蓝色 + 粉红 |
| 7 | 渐变蓝绿 | 深青 + 金棕 + 深红 |
| 8 | 柔和粉彩 | 粉红 + 薄荷绿 + 浅蓝 |

---

## 技术实现细节

### EMU 单位转换

PowerPoint 内部使用 EMU（English Metric Unit）作为坐标单位：

```python
PX_TO_EMU = 9525  # 1 像素 = 9525 EMU

def px(val):
    """将像素值转换为 EMU"""
    return int(val * PX_TO_EMU)
```

### 图表库数据结构

`chart_library.json` 存储从模板提取的图表样式：

```json
{
  "styles": {
    "column_chart": [
      {
        "source": "ppt模板1.pptx",
        "slide": 3,
        "shape_name": "Chart 1",
        "pptx_type": "COLUMN_CLUSTERED (51)",
        "series_colors": ["#2B4570", "#4682B4"],
        "left": 40, "top": 90,
        "width": 600, "height": 400
      }
    ]
  },
  "scanned_files": {
    "ppt模板1.pptx": "ppt模板1.pptx:245760:1714500000"
  }
}
```

### 分页算法

```python
MAX_CHARTS_PER_SLIDE = 4

# 按 Excel 章节分组
# 每个章节内超过 4 个图表则拆分为多页
# 页标题格式：章节名 (页码/总页数)
```

---

## 示例输出

使用 `data3.xlsx`（新能源汽车行业分析报告）生成的 PPT 包含：

| 页码 | 内容 | 图表数 |
|------|------|--------|
| 1 | 封面：报告标题 + 摘要 | 0 |
| 2 | 品牌累计销量汇总 | 3 |
| 3 | 品牌月度销量趋势 | 1（多系列折线图） |
| 4 | 细分市场销量占比 | 3 |
| 5 | 品牌同比增长率排名 | 1 |
| 6 | 城市终端销量 TOP10 | 3 |
| 7 | 动力类型市场分布 | 2 |
| 8 | 品牌单车均价排名 | 1 |
| 9 | 品牌出口量排名 | 1 |
| 10 | 用户满意度评分排名 | 2 |
| 11 | 投融资事件汇总 | 1 |

共 11 页，17 个原生可编辑图表。

---

## License

MIT
