# rangeUtils

## 模块用途与适用场景

用于区域与单元格的统一读写，包括矩阵读写、公式写入、清空、偏移、缩放和冻结窗格。

适用场景：
- 将二维数组批量写入工作表
- 读取区域数据后进行二次处理
- 统一处理冻结表头逻辑

## 公共函数目录

- `getRange`
- `readCell`
- `writeCell`
- `readMatrix`
- `writeMatrix`
- `writeFormulaR1C1`
- `clearRange`
- `resizeFrom`
- `offsetFrom`
- `freezePanesAt`

## 函数说明

### 1. getRange

作用：按地址或行列参数获取区域对象。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| row | number\|string | 是 | 起始行或地址字符串 |
| col | number | 否 | 起始列 |
| rowCount | number | 否 | 行数 |
| colCount | number | 否 | 列数 |

返回值：区域对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var rg = getRange(ws, 2, 1, 100, 8);
var rg2 = getRange(ws, "A1:C10");
```

示例目的：统一区域定位入口。

### 2. readCell

作用：读取单元格值（优先 `Value2`）。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| rangeOrWorksheet | Object | 是 | 区域对象或工作表对象 |
| row | number | 否 | 当传工作表时使用 |
| col | number | 否 | 当传工作表时使用 |

返回值：单元格值，失败返回 `null`。  
返回格式：`any|null`

示例代码：

```javascript
var v1 = readCell(ws, 2, 3);
var v2 = readCell(ws.Range("A1"));
```

示例目的：兼容两种调用方式，提高复用性。

### 3. writeCell

作用：写入单元格值（优先 `Value2`）。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| rangeOrWorksheet | Object | 是 | 区域对象或工作表对象 |
| row | number | 否 | 当传工作表时使用 |
| col | number | 否 | 当传工作表时使用 |
| value | any | 是 | 要写入的值 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
writeCell(ws, 2, 3, "完成");
```

示例目的：在流程中快速更新状态单元格。

### 4. readMatrix

作用：将区域值读取为二维数组。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 区域对象 |

返回值：二维数组。  
返回格式：`Array`

示例代码：

```javascript
var data = readMatrix(ws.Range("A1:D20"));
```

示例目的：将表格数据加载到内存进行排序/筛选。

### 5. writeMatrix

作用：将二维数组批量写入区域。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| targetRangeOrWorksheet | Object | 是 | 目标区域或工作表 |
| row | number | 否 | 当传工作表时使用 |
| col | number | 否 | 当传工作表时使用 |
| matrix | Array | 是 | 二维数组 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
writeMatrix(ws, 2, 1, [["a", 1], ["b", 2]]);
```

示例目的：一次性写入结果集，减少逐单元格性能开销。

### 6. writeFormulaR1C1

作用：写入 R1C1 公式。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| targetRangeOrWorksheet | Object | 是 | 目标区域或工作表 |
| row | number | 否 | 当传工作表时使用 |
| col | number | 否 | 当传工作表时使用 |
| formulaR1C1 | string | 是 | R1C1 公式 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
writeFormulaR1C1(ws, 2, 5, "=RC[-2]*RC[-1]");
```

示例目的：按相对列快速批量填充计算逻辑。

### 7. clearRange

作用：清空区域内容。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 目标区域 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
clearRange(ws.Range("A2:Z1000"));
```

示例目的：写入新数据前清空旧数据。

### 8. resizeFrom

作用：基于区域起点调整大小。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 起始区域 |
| rowCount | number | 是 | 行数 |
| colCount | number | 是 | 列数 |

返回值：新区域对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var target = resizeFrom(ws.Range("A2"), 200, 6);
```

示例目的：根据数据尺寸动态生成目标写入范围。

### 9. offsetFrom

作用：基于区域偏移并可选调整大小。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 基准区域 |
| rowOffset | number | 是 | 行偏移 |
| colOffset | number | 是 | 列偏移 |
| rowCount | number | 否 | 调整后行数 |
| colCount | number | 否 | 调整后列数 |

返回值：新区域对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var detail = offsetFrom(ws.Range("A1"), 1, 0, 100, 8);
```

示例目的：从表头区域快速定位到明细区域。

### 10. freezePanesAt

作用：在指定单元格位置冻结窗格。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| row | number | 是 | 冻结分隔行 |
| col | number | 是 | 冻结分隔列 |
| app | Object | 否 | 应用对象 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
freezePanesAt(ws, 2, 2);
```

示例目的：固定首行首列，便于查看大表。

## 依赖说明与 WPS-first / Windows fallback

- 完全依赖 WPS/JSA 表格对象：`Range`、`Cells`、`Value2`、`FormulaR1C1`、`ActiveWindow`。
- 不依赖 Windows ActiveX 文件或网络对象。
