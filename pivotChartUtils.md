# pivotChartUtils

## 模块用途与适用场景

用于透视表和图表的基础封装，包括创建透视表、刷新、设置字段、创建图表和修改图表属性。

适用场景：
- 从明细表生成汇总透视
- 自动刷新透视结果
- 按固定模板生成图表

## 公共函数目录

- `createPivotTable`
- `refreshPivotTable`
- `setPivotField`
- `createChart`
- `setChartType`
- `setChartTitle`

## 函数说明

### 1. createPivotTable

作用：根据源区域和目标位置创建透视表。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| sourceRange | Object | 是 | 源数据区域 |
| destinationRange | Object | 是 | 目标起始区域 |
| pivotName | string | 否 | 透视表名 |
| cacheName | string | 否 | 缓存名 |

返回值：透视表对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var pt = createPivotTable(
    wsData.Range("A1:H1000"),
    wsPivot.Range("A3"),
    "ptSales"
);
```

示例目的：用明细区域快速生成透视表骨架。

### 2. refreshPivotTable

作用：刷新透视表数据。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| pivotTable | Object | 是 | 透视表对象 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
refreshPivotTable(pt);
```

示例目的：更新数据源后强制刷新透视结果。

### 3. setPivotField

作用：设置透视字段方向和位置。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| pivotTable | Object | 是 | 透视表对象 |
| fieldName | string | 是 | 字段名 |
| orientation | number | 是 | 字段方向常量 |
| position | number | 否 | 字段位置 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
setPivotField(pt, "店铺", 1, 1);
setPivotField(pt, "销售额", 4, 1);
```

示例目的：定义透视表行字段和数据字段布局。

### 4. createChart

作用：创建图表并绑定数据源。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| sourceRange | Object | 是 | 数据区域 |
| chartType | number | 否 | 图表类型 |
| left | number | 否 | 左位置 |
| top | number | 否 | 上位置 |
| width | number | 否 | 宽度 |
| height | number | 否 | 高度 |

返回值：ChartObject 对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var chartObj = createChart(wsPivot, wsPivot.Range("A3:D20"), 51, 80, 60, 560, 320);
```

示例目的：自动生成固定尺寸图表用于看板。

### 5. setChartType

作用：设置图表类型。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| chartObjectOrChart | Object | 是 | ChartObject 或 Chart |
| chartType | number | 是 | 图表类型常量 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
setChartType(chartObj, 4);
```

示例目的：按场景切换柱状图/折线图。

### 6. setChartTitle

作用：设置图表标题。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| chartObjectOrChart | Object | 是 | ChartObject 或 Chart |
| title | string | 是 | 标题文本 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
setChartTitle(chartObj, "2026年4月销售趋势");
```

示例目的：输出图表时自动填充标题。

## 依赖说明与 WPS-first / Windows fallback

- 基于 WPS/JSA 的 `PivotCaches`、`PivotTable`、`ChartObjects`。
- 不依赖 Windows ActiveX。
