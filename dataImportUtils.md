# dataImportUtils

## 模块用途与适用场景

用于把工作表数据读取为二维数组或对象数组，并支持按主键建立索引。

适用场景：
- 外部工作簿首表导入
- 按表头映射读取字段
- 明细对象化后做业务计算

## 公共函数目录

- `readWorksheetMatrix`
- `readFirstSheetByHeaderMap`
- `indexRowsByKey`
- `readRowsAsObjects`

## 函数说明

### 1. readWorksheetMatrix

作用：按范围读取工作表二维数组。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| startRow | number | 否 | 起始行 |
| startCol | number | 否 | 起始列 |
| endRow | number | 否 | 结束行 |
| endCol | number | 否 | 结束列 |

返回值：二维数组。  
返回格式：`Array`

示例代码：

```javascript
var matrix = readWorksheetMatrix(ws, 1, 1, 100, 8);
```

示例目的：把固定区域一次性加载到内存。

### 2. readFirstSheetByHeaderMap

作用：按“字段 -> 表头名”映射读取第一张工作表。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| workbook | Object | 是 | 工作簿对象 |
| headerMap | Object | 是 | 字段映射对象 |
| headerRow | number | 否 | 表头行 |
| startRow | number | 否 | 数据起始行 |

返回值：对象数组。  
返回格式：

```javascript
[
    { shopName: "店铺A", amount: 1200 },
    { shopName: "店铺B", amount: 900 }
]
```

示例代码：

```javascript
var rows = readFirstSheetByHeaderMap(wb, {
    shopName: "店铺",
    amount: "销售额"
}, 1, 2);
```

示例目的：按目标字段名读取外部模板不一致的列结构。

### 3. indexRowsByKey

作用：将对象数组按键字段建立索引对象。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| rows | Array | 是 | 对象数组 |
| keyField | string | 是 | 键字段名 |

返回值：索引对象。  
返回格式：

```javascript
{
    "A001": { id: "A001", name: "苹果" },
    "A002": { id: "A002", name: "香蕉" }
}
```

示例代码：

```javascript
var map = indexRowsByKey(rows, "id");
```

示例目的：后续按主键 O(1) 查询行对象。

### 4. readRowsAsObjects

作用：按表头行将工作表读取为对象数组。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| headerRow | number | 否 | 表头行 |
| startRow | number | 否 | 数据起始行 |
| endRow | number | 否 | 数据结束行 |
| endCol | number | 否 | 数据结束列 |

返回值：对象数组。  
返回格式：

```javascript
[
    { "店铺": "店铺A", "销售额": 1200 },
    { "店铺": "店铺B", "销售额": 900 }
]
```

示例代码：

```javascript
var list = readRowsAsObjects(ws, 1, 2);
```

示例目的：将表格明细转成可直接处理的对象结构。

## 依赖说明与 WPS-first / Windows fallback

- 完全基于 WPS/JSA 的 `Worksheet`、`Range`、`Value2`。
- 不依赖 Windows ActiveX。
