# tableQueryUtils

## 模块用途与适用场景

用于管理表对象（ListObject）和查询表刷新，包括创建表、按名获取、追加行、增列、排序与刷新。

适用场景：
- 将普通区域升级为结构化表
- 给表批量追加数据
- 刷新与外部数据连接关联的查询表

## 公共函数目录

- `createTable`
- `getTableByName`
- `appendTableRow`
- `addTableColumn`
- `sortTable`
- `refreshQueryTable`

## 函数说明

### 1. createTable

作用：在工作表上创建表对象。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| range | Object\|string | 是 | 区域对象或地址 |
| tableName | string | 否 | 表名 |
| hasHeaders | boolean | 否 | 是否有表头 |

返回值：表对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var tb = createTable(ws, "A1:F100", "SalesTable", true);
```

示例目的：把数据区域转换为结构化表，方便后续引用。

### 2. getTableByName

作用：按名称在工作簿范围查找表对象。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| workbook | Object | 是 | 工作簿对象 |
| tableName | string | 是 | 表名 |

返回值：表对象，未找到返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var tb = getTableByName(wb, "SalesTable");
```

示例目的：跨工作表定位目标表。

### 3. appendTableRow

作用：向表对象追加一行数据。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| table | Object | 是 | 表对象 |
| rowValues | Array | 是 | 行值数组 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
appendTableRow(tb, ["2026-04-09", "店铺A", 1200]);
```

示例目的：把计算结果逐行写入结构化表。

### 4. addTableColumn

作用：给表对象新增列并可选写默认值。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| table | Object | 是 | 表对象 |
| columnName | string | 是 | 列名 |
| defaultValue | any | 否 | 默认值 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
addTableColumn(tb, "处理状态", "待处理");
```

示例目的：运行时动态扩展业务字段。

### 5. sortTable

作用：按列名或列号对表进行排序。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| table | Object | 是 | 表对象 |
| keyColumnName | string\|number | 是 | 列名或 1 基列号 |
| ascending | boolean | 否 | 是否升序 |
| hasHeader | boolean | 否 | 是否有表头 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
sortTable(tb, "销售额", false, true);
```

示例目的：按关键指标快速排序输出。

### 6. refreshQueryTable

作用：刷新查询表数据。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| tableOrWorksheet | Object | 是 | 表对象或工作表对象 |
| tableName | string | 否 | 当传工作表时使用 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
refreshQueryTable(tb);
```

示例目的：同步外部连接数据后再执行后续逻辑。

## 依赖说明与 WPS-first / Windows fallback

- 基于 WPS/JSA `ListObjects`、`ListRows`、`ListColumns`、`QueryTable`。
- 不依赖 Windows ActiveX。
