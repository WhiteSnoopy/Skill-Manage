---
name: excel-master
description: "全能Excel智能处理引擎。当用户需要对Excel文件进行任何操作时触发，包括但不限于：(1)数据清洗和转换，(2)公式计算和批量填充，(3)数据筛选、排序和去重，(4)格式化和样式设置，(5)数据透视和统计分析，(6)多表合并和拆分，(7)条件格式和数据验证，(8)图表生成，(9)查找替换，(10)行列操作。使用自然语言描述需求，自动分析意图→生成指令→执行→输出结果到原文件同目录。"
---

# Excel Master - 全能Excel智能处理引擎

## 核心工作流程

```
用户请求 → 意图分析 → 指令编排 → 执行处理 → 输出结果
```

### 处理流程

1. **意图分析**: 解析用户自然语言，识别操作类型和目标
2. **指令编排**: 将意图转换为可执行的指令序列
3. **执行处理**: 按顺序执行指令，处理Excel文件
4. **输出结果**: 保存到用户文件同目录，生成处理报告

## 执行步骤

### Step 1: 一次性探测文件结构（关键！）

**必须先执行此探测脚本**，获取完整文件信息后再进行任何操作：

```python
import pandas as pd
from openpyxl import load_workbook
import json

file_path = "用户文件路径"

# 一次性获取所有信息
wb = load_workbook(file_path, read_only=True, data_only=True)
info = {
    "sheets": [],
    "file": file_path
}

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    # 获取表头（第一行）
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    # 估算行数
    row_count = ws.max_row - 1 if ws.max_row else 0
    info["sheets"].append({
        "name": sheet_name,  # 使用实际名称，不要假设大小写！
        "columns": [h for h in headers if h],
        "rows": row_count
    })

wb.close()
print(json.dumps(info, ensure_ascii=False, indent=2))
```

**重要**：后续所有操作必须使用探测到的**实际sheet名称**，不要假设！

### Step 2: 分析意图并生成执行计划

基于Step 1的探测结果，输出执行计划：

```
【文件结构】
- Sheet: [实际名称] → 列: [...], 行数: N
- Sheet: [实际名称] → 列: [...], 行数: N

【意图分析】
- 操作类型: [数据清洗/格式化/计算/筛选/合并/...]
- 目标范围: [全表/指定列/指定行/指定区域]
- 预期结果: [描述输出效果]

【执行计划】
1. [具体指令1]
2. [具体指令2]
...
```

### Step 3: 一次性执行处理（避免多次调用）

将所有操作合并到**一个脚本**中执行，减少交互次数：

```python
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

file_path = "用户文件路径"

# === 读取数据（使用实际sheet名称！）===
df1 = pd.read_excel(file_path, sheet_name='实际名称1')
df2 = pd.read_excel(file_path, sheet_name='实际名称2')

# === 执行所有处理逻辑 ===
# ... 具体操作 ...

# === 写入结果 ===
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    result.to_excel(writer, sheet_name='结果sheet', index=False)

# === 输出摘要 ===
print(f"处理完成: {len(result)} 行")
```

### Step 4: 输出结果摘要

```
【处理完成】
✅ 文件: xxx.xlsx
✅ 操作: [具体操作描述]
✅ 结果: [行数/匹配率等关键指标]
```

## 操作类型速查

### 数据操作
| 操作 | 触发词 | 示例 |
|------|--------|------|
| 筛选 | 筛选、过滤、保留 | "筛选销售额>1000的行" |
| 排序 | 排序、升序、降序 | "按日期降序排列" |
| 去重 | 去重、删除重复 | "按姓名去重" |
| 查找替换 | 替换、改成 | "把N/A替换为空" |
| 拆分 | 拆分、分列 | "按逗号拆分地址" |
| 合并列 | 合并、拼接 | "合并姓和名" |

### 计算操作
| 操作 | 触发词 | 示例 |
|------|--------|------|
| 求和 | 求和、总计 | "计算月销售总额" |
| 平均 | 平均、均值 | "计算平均分" |
| 公式 | 添加公式、计算列 | "添加利润率=利润/收入" |
| 条件 | 如果则、SUMIF | "状态完成则标1" |

### 格式操作
| 操作 | 触发词 | 示例 |
|------|--------|------|
| 数字格式 | 货币、百分比 | "金额设为货币格式" |
| 日期 | 转日期 | "转为YYYY-MM-DD" |
| 条件格式 | 标红、高亮 | "销售<500标红" |
| 样式 | 加粗、居中、颜色 | "表头加粗居中" |

### 结构操作
| 操作 | 触发词 | 示例 |
|------|--------|------|
| 增删行列 | 添加、删除、插入 | "删除空行" |
| 移动 | 移动、调整 | "备注列移到最后" |
| 重命名 | 重命名 | "Sheet1改名销售数据" |
| 合并单元格 | 合并单元格 | "合并A1:C1为标题" |

### 多表操作
| 操作 | 触发词 | 示例 |
|------|--------|------|
| 关联合并 | vlookup、关联 | "按订单号合并两表" |
| 拆分 | 按...分表 | "按月份拆分sheet" |

### 分析操作
| 操作 | 触发词 | 示例 |
|------|--------|------|
| 透视 | 透视、分组统计 | "按类别月份统计销售" |
| 图表 | 图表、柱状图 | "生成月度趋势图" |

## 代码模板

### 数据清洗
```python
import pandas as pd
from pathlib import Path

input_path = Path(input_file)
df = pd.read_excel(input_path)

df = df.dropna(how='all')  # 删空行
df = df.drop_duplicates(subset=['姓名'])  # 去重
df['状态'] = df['状态'].replace({'N/A': ''})  # 替换

output_path = input_path.parent / f"{input_path.stem}_cleaned.xlsx"
df.to_excel(output_path, index=False)
```

### 格式化
```python
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from pathlib import Path

wb = load_workbook(input_file)
ws = wb.active

# 表头样式
for cell in ws[1]:
    cell.font = Font(bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor='4472C4')
    cell.alignment = Alignment(horizontal='center')

output_path = Path(input_file).parent / f"{Path(input_file).stem}_formatted.xlsx"
wb.save(output_path)
```

### 条件格式
```python
from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

wb = load_workbook(input_file)
ws = wb.active

red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
ws.conditional_formatting.add('B2:B100', CellIsRule(operator='lessThan', formula=['500'], fill=red_fill))

wb.save(output_path)
```

### 数据透视
```python
import pandas as pd

df = pd.read_excel(input_file)
pivot = pd.pivot_table(df, values='销售额', index='类别', columns='月份', 
                       aggfunc='sum', fill_value=0, margins=True)
pivot.to_excel(output_path)
```

### 多表合并
```python
import pandas as pd

# 注意：使用探测到的实际sheet名称！
df1 = pd.read_excel(input_file, sheet_name='sheet1')  # 小写
df2 = pd.read_excel(input_file, sheet_name='sheet2')

# 关联合并（指定需要的列，避免列名冲突）
merged = pd.merge(
    df1, 
    df2[['OMS订单号', '渠道来源', '出库单号']],  # 只取需要的列
    on='OMS订单号', 
    how='left'
)

# 写入新sheet（保留原有sheets）
with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    merged.to_excel(writer, sheet_name='sheet3', index=False)
```

## ⚠️ 常见陷阱

| 陷阱 | 正确做法 |
|------|----------|
| 假设Sheet名为`Sheet1` | **先探测实际名称**，可能是`sheet1`或自定义名 |
| 多次读取文件 | 一次探测获取所有信息，一次执行完成所有操作 |
| 覆盖原文件丢失数据 | 使用`mode='a'`追加写入，或输出新文件 |
| merge后列名冲突 | 指定`suffixes`或只取需要的列 |
| 大文件内存溢出 | 使用`chunksize`分块读取 |

### 添加公式（使用Excel公式而非硬编码）
```python
from openpyxl import load_workbook

wb = load_workbook(input_file)
ws = wb.active

# ✅ 使用Excel公式
ws['D2'] = '=B2*C2'  # 金额=数量*单价
for row in range(3, ws.max_row + 1):
    ws[f'D{row}'] = f'=B{row}*C{row}'

ws[f'D{ws.max_row + 1}'] = f'=SUM(D2:D{ws.max_row})'  # 总计

wb.save(output_path)
```

## 输出规范

### 文件命名
- 默认: `{原文件名}_processed.xlsx`
- 按操作: `_cleaned`, `_formatted`, `_merged`, `_pivot`

### 保存位置
1. **优先**: 原文件同目录
2. **备用**: `/mnt/user-data/outputs/`

### 处理报告
```
【处理完成】
✅ 输入: xxx.xlsx
✅ 输出: xxx_processed.xlsx
✅ 位置: /path/to/output/

【操作摘要】
- 删除空行 → 移除15行
- 去重 → 移除8项
- 格式化表头 → 10列

【数据概览】
- 行数: 1,234 | 列数: 15
```

## 错误处理

| 错误 | 原因 | 处理 |
|------|------|------|
| `Worksheet Sheet1 does not exist` | Sheet名称大小写不匹配 | **先用探测脚本获取实际名称** |
| 文件不存在 | 路径错误 | 提示用户确认路径 |
| 列名不存在 | 用户描述与实际不符 | 列出现有列名供选择 |
| PermissionError | 文件被其他程序打开 | 提示关闭Excel后重试 |
| 内存不足 | 文件过大 | 使用chunksize分块处理 |

## 最佳实践

1. **先探测再操作**：永远先运行探测脚本获取实际sheet名称和列名
2. **一次性执行**：将所有操作合并到一个脚本，减少交互次数
3. **不假设命名**：Sheet名可能是 `Sheet1`、`sheet1` 或中文名
4. **保留原数据**：使用 `mode='a'` 追加或输出新文件
5. **使用Excel公式**：尽量用公式而非硬编码值
6. **明确输出**：处理完成后给出行数、匹配率等关键指标
