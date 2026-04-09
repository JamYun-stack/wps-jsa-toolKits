# formatUtils

## 模块用途与适用场景

用于复制和设置单元格格式：数字格式、字体、填充、边框以及自适应行高列宽。

适用场景：
- 模板格式复制到明细区域
- 输出报表统一样式
- 自动调整列宽避免内容被遮挡

## 公共函数目录

- `copyRowFormats`
- `copyRangeFormats`
- `setNumberFormat`
- `setFontStyle`
- `setFillColor`
- `setBorderStyle`
- `autoFitColumns`
- `autoFitRows`

## 函数说明

### 1. copyRowFormats

作用：复制行格式到目标行。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| sourceRowRange | Object | 是 | 源行区域 |
| targetRowRange | Object | 是 | 目标行区域 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
copyRowFormats(ws.Rows(2), ws.Rows(3));
```

示例目的：将模板行样式复制到新插入行。

### 2. copyRangeFormats

作用：复制区域格式。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| sourceRange | Object | 是 | 源区域 |
| targetRange | Object | 是 | 目标区域 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
copyRangeFormats(ws.Range("A1:F1"), ws.Range("A2:F2"));
```

示例目的：快速继承标题行样式。

### 3. setNumberFormat

作用：设置数字格式。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 目标区域 |
| formatText | string | 是 | 格式字符串 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
setNumberFormat(ws.Range("C2:C200"), "#,##0.00");
```

示例目的：统一金额显示格式。

### 4. setFontStyle

作用：设置字体名称、大小、粗斜体与颜色。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 目标区域 |
| fontName | string | 否 | 字体 |
| fontSize | number | 否 | 字号 |
| bold | boolean | 否 | 是否加粗 |
| italic | boolean | 否 | 是否斜体 |
| color | number | 否 | RGB 颜色整数 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
setFontStyle(ws.Range("A1:F1"), "微软雅黑", 10, true, false, 0xFFFFFF);
```

示例目的：统一表头字体风格。

### 5. setFillColor

作用：设置单元格填充色。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 目标区域 |
| color | number | 是 | RGB 颜色整数 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
setFillColor(ws.Range("A1:F1"), 0x2F75B5);
```

示例目的：高亮标题区域。

### 6. setBorderStyle

作用：设置边框线型、粗细、颜色。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| range | Object | 是 | 目标区域 |
| lineStyle | number | 否 | 线型 |
| weight | number | 否 | 粗细 |
| color | number | 否 | 颜色 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
setBorderStyle(ws.Range("A1:F100"), 1, 2, 0xD9D9D9);
```

示例目的：生成标准网格样式。

### 7. autoFitColumns

作用：自动调整列宽。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| rangeOrWorksheet | Object | 是 | 区域或工作表 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
autoFitColumns(ws.UsedRange);
```

示例目的：避免文字显示被截断。

### 8. autoFitRows

作用：自动调整行高。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| rangeOrWorksheet | Object | 是 | 区域或工作表 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
autoFitRows(ws.UsedRange);
```

示例目的：自动匹配换行文本高度。

## 依赖说明与 WPS-first / Windows fallback

- 全部基于 WPS/JSA 表格对象：`Range`、`Font`、`Interior`、`Borders`、`PasteSpecial`。
- 不依赖 Windows ActiveX。
