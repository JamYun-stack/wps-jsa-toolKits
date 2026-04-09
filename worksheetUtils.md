# worksheetUtils

## 模块用途与适用场景

用于工作表定位与边界计算：按名称/索引取表、确保工作表存在、列出工作表、读取 `UsedRange` 边界、查找最后行列。

适用场景：
- 自动创建缺失配置页
- 动态判断数据区域大小
- 批量循环全部工作表

## 公共函数目录

- `getWorksheet`
- `ensureWorksheet`
- `listWorksheetNames`
- `getUsedRangeBounds`
- `findLastRow`
- `findLastColumn`

## 函数说明

### 1. getWorksheet

作用：按名称或索引获取工作表。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| workbook | Object | 是 | 工作簿对象 |
| sheetNameOrIndex | string\|number | 是 | 名称或 1 基索引 |

返回值：工作表对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var ws = getWorksheet(wb, "操作");
```

示例目的：统一工作表获取方式，避免散落 `try/catch`。

### 2. ensureWorksheet

作用：确保工作簿中存在指定名称工作表，不存在则新建。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| workbook | Object | 是 | 工作簿对象 |
| sheetName | string | 是 | 目标名称 |
| position | number | 否 | 新建插入位置 |

返回值：工作表对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var ws = ensureWorksheet(wb, "结果");
```

示例目的：执行前自动补齐目标输出页。

### 3. listWorksheetNames

作用：列出工作簿内全部工作表名称。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| workbook | Object | 是 | 工作簿对象 |

返回值：名称数组。  
返回格式：`Array`

示例代码：

```javascript
var names = listWorksheetNames(wb);
```

示例目的：用于日志输出和白名单校验。

### 4. getUsedRangeBounds

作用：读取 `UsedRange` 边界信息。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |

返回值：边界信息对象，失败返回 `null`。  
返回格式：

```javascript
{
    row: 1,
    col: 1,
    rows: 10,
    cols: 5,
    lastRow: 10,
    lastCol: 5,
    address: "$A$1:$E$10"
}
```

示例代码：

```javascript
var bounds = getUsedRangeBounds(ws);
if (bounds) {
    MsgBox(bounds.address);
}
```

示例目的：动态计算数据读取范围。

### 5. findLastRow

作用：查找指定列最后一个非空行。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| columnIndex | number | 否 | 列号，默认 1 |
| startRow | number | 否 | 最小返回行号 |

返回值：最后行号，失败返回 `0`。  
返回格式：`number`

示例代码：

```javascript
var lastRow = findLastRow(ws, 1, 2);
```

示例目的：定位明细数据末尾。

### 6. findLastColumn

作用：查找指定行最后一个非空列。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| rowIndex | number | 否 | 行号，默认 1 |
| startColumn | number | 否 | 最小返回列号 |

返回值：最后列号，失败返回 `0`。  
返回格式：`number`

示例代码：

```javascript
var lastCol = findLastColumn(ws, 1, 1);
```

示例目的：根据表头自动识别字段数量。

## 依赖说明与 WPS-first / Windows fallback

- 完全基于 WPS/JSA 表格对象：`Workbook`、`Worksheets`、`UsedRange`、`Cells`。
- 不依赖 Windows ActiveX 文件系统对象。
