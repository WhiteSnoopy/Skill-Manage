# Excel Master 指令参考

## ExcelProcessor 类方法

### 数据操作

#### filter_rows(column, condition, value)
筛选符合条件的行
- condition: `>`, `>=`, `<`, `<=`, `==`, `!=`, `contains`, `startswith`, `endswith`

#### sort_data(columns, ascending)
按指定列排序
- columns: 列名列表 `['日期', '金额']`
- ascending: 升序标志列表 `[True, False]`

#### remove_duplicates(columns)
去除重复行
- columns: 去重依据列，None表示全部列

#### replace_values(column, old_value, new_value)
替换指定值

#### remove_empty_rows()
删除所有空行

#### split_column(column, delimiter, new_columns)
按分隔符拆分列
- new_columns: 拆分后的新列名列表

#### merge_columns(columns, new_column, separator)
合并多列为一列

### 计算操作

#### add_sum_row(columns)
添加求和行
- columns: 要求和的列，None表示所有数字列

#### add_calculated_column(new_column, formula)
添加计算列（Python表达式）
- formula: 使用列名的表达式 `"销售额 * 利润率"`

#### add_excel_formula_column(column_letter, formula_template, start_row, end_row)
添加Excel公式列
- formula_template: 使用`{row}`占位符 `"=B{row}*C{row}"`

### 格式操作

#### format_header(bold, bg_color, font_color, center)
格式化表头
- bg_color: 十六进制颜色 `"4472C4"`

#### set_column_width(column_widths)
设置列宽
- column_widths: `{"A": 15, "B": 20}`

#### auto_column_width()
自动调整所有列宽

#### add_conditional_format(range_str, rule_type, value, fill_color)
添加条件格式
- rule_type: `lessThan`, `greaterThan`, `equal`

#### format_as_currency(columns, symbol)
货币格式化

#### format_as_percentage(columns, decimals)
百分比格式化

### 结构操作

#### add_row(position, values)
插入行

#### delete_rows(condition_column, condition_value, empty_only)
删除行

#### rename_columns(mapping)
重命名列
- mapping: `{"旧名": "新名"}`

#### reorder_columns(new_order)
重排列顺序

### 多表操作

#### merge_sheets(sheet_names, on, how)
合并多个sheet
- how: `left`, `right`, `inner`, `outer`

#### split_by_column(column)
按列值拆分为多个sheet

### 分析操作

#### create_pivot_table(values, index, columns, aggfunc)
创建数据透视表
- aggfunc: `sum`, `mean`, `count`, `max`, `min`

#### add_chart(chart_type, data_range, title, position)
添加图表
- chart_type: `bar`, `line`, `pie`

#### get_summary()
获取数据摘要

## JSON指令格式

```json
[
  {
    "method": "remove_empty_rows",
    "params": {}
  },
  {
    "method": "filter_rows",
    "params": {
      "column": "销售额",
      "condition": ">",
      "value": 1000
    }
  },
  {
    "method": "format_header",
    "params": {
      "bold": true,
      "bg_color": "4472C4"
    }
  }
]
```

## 命令行用法

```bash
# 获取文件信息
python excel_processor.py input.xlsx --info

# 执行操作
python excel_processor.py input.xlsx --operations '[{"method":"remove_empty_rows","params":{}}]'
```
