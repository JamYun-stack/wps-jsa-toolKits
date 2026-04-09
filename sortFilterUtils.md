# sortFilterUtils

## 模块用途与适用场景

用于二维数组和工作表区域的排序、筛选、去重处理。

适用场景：
- 数据写入前在内存排序
- 工作表区域按指定列重排
- 自动筛选并导出目标行

## 公共函数目录

- `sort2DByColumn`
- `sortRangeByColumn`
- `applyAutoFilter`
- `clearAutoFilter`
- `dedupeMatrix`

## 函数说明

### 1. sort2DByColumn

作用：按指定列对二维数组排序，支持保留表头。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| matrix | Array | 是 | 二维数组 |
| columnIndex | number | 是 | 0 基列索引 |
| ascending | boolean | 否 | 是否升序，默认 true |
| hasHeader | boolean | 否 | 是否包含表头 |

返回值：排序后的新二维数组。  
返回格式：`Array`

示例代码：

```javascript
var sorted = sort2DByColumn(data, 2, false, true);
```

示例目的：按第 3 列降序排列，并保留首行表头。

### 2. sortRangeByColumn

作用：对区域数据按指定列排序并写回。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 区域对象 |
| columnIndex | number | 是 | 1 基列号 |
| ascending | boolean | 否 | 是否升序 |
| hasHeader | boolean | 否 | 是否包含表头 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
sortRangeByColumn(ws.Range("A1:F100"), 4, true, true);
```

示例目的：按第 4 列升序重新整理明细区域。

### 3. applyAutoFilter

作用：对区域应用自动筛选条件。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 区域对象 |
| fieldIndex | number | 是 | 1 基字段索引 |
| criteria1 | any | 否 | 条件 1 |
| operator | number | 否 | 条件操作符 |
| criteria2 | any | 否 | 条件 2 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
applyAutoFilter(ws.Range("A1:F1000"), 2, "华东");
```

示例目的：筛选区域字段中“华东”数据。

### 4. clearAutoFilter

作用：清除工作表筛选状态。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
clearAutoFilter(ws);
```

示例目的：下一轮筛选前重置状态。

### 5. dedupeMatrix

作用：按指定键列对二维数组去重。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| matrix | Array | 是 | 二维数组 |
| keyColumns | Array\|number | 是 | 0 基键列 |
| keepFirst | boolean | 否 | 保留首条或末条 |

返回值：去重后的新二维数组。  
返回格式：`Array`

示例代码：

```javascript
var deduped = dedupeMatrix(data, [0, 2], true);
```

示例目的：按“店铺+分类”保留首条记录。

## 依赖说明与 WPS-first / Windows fallback

- 优先使用 WPS 区域对象 `Range` 的 `AutoFilter`。
- 区域排序采用“读入内存排序再写回”，不依赖 Windows 对象。
