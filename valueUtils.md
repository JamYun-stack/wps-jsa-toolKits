# valueUtils.js 基础值与日期工具模块

`valueUtils.js` 用于处理 WPS JSA 宏里最常见的值转换、空值判断，以及 WPS/Excel 日期序列与 JavaScript `Date` 之间的互转。

## 使用约定

- 以“宏脚本直接可调用”为目标，返回值尽量稳定，不依赖现代 JavaScript 语法。
- 日期序列默认按 WPS/Excel 常见的 1900 日期系统处理。
- 无法解析日期时，多数日期函数返回 `null`、`NaN` 或空字符串，方便业务代码先判断再处理。

## 公开函数

- `isBlank(value)`
- `toNumber(value, defaultValue)`
- `toStringSafe(value, defaultValue)`
- `toBoolean(value, defaultValue)`
- `toVal(value)`
- `isDateSerial(value)`
- `serialToDate(serial)`
- `dateToSerial(value)`
- `parseDateInput(value)`
- `formatDate(value, pattern)`
- `formatDateTime(value, pattern)`
- `startOfDay(value)`
- `endOfDay(value)`
- `startOfMonth(value)`
- `endOfMonth(value)`
- `diffDays(startValue, endValue)`
- `diffSeconds(startValue, endValue)`

### isBlank(value)

作用：判断值是否为空白。

参数：`value` 为任意值。

返回值：`boolean`。`null`、`undefined`、空字符串或只包含空白字符的字符串返回 `true`。

示例代码：

```js
if (isBlank(Cells(2, 1).Value)) {
    MsgBox("A2 为空");
}
```

示例代码完成的目的：在读取单元格数据后，快速判断该值是否需要跳过。

### toNumber(value, defaultValue)

作用：把值转换为数字。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `value` | `*` | 要转换的值。 |
| `defaultValue` | `number` | 转换失败时返回的默认值。默认是 `0`。 |

返回值：`number`。

示例代码：

```js
var amount = toNumber(Cells(2, 3).Value, 0);
```

示例代码完成的目的：把表格中的数量、金额等值安全转成数字。

### toStringSafe(value, defaultValue)

作用：把值安全转换为字符串。

参数：`value` 为任意值；`defaultValue` 为值为空时使用的默认字符串。

返回值：`string`。

示例代码：

```js
var shopName = toStringSafe(Cells(2, 2).Value, "");
```

示例代码完成的目的：避免单元格值为空时出现 `null`、`undefined` 字样。

### toBoolean(value, defaultValue)

作用：把常见文本、数字或布尔值转换为布尔结果。

参数：`value` 为任意值；`defaultValue` 为无法识别时的默认布尔值。

返回值：`boolean`。

示例代码：

```js
var enabled = toBoolean("yes", false);
```

示例代码完成的目的：把配置中的开关值统一成 `true` 或 `false`。

### toVal(value)

作用：按项目里常见习惯把值转成数字；失败时直接返回 `0`。

参数：`value` 为任意值。

返回值：`number`。

示例代码：

```js
var count = toVal(Cells(2, 5).Value);
```

示例代码完成的目的：与现有宏脚本里“空值按 0 处理”的习惯保持一致。

### isDateSerial(value)

作用：判断值是否可以视为日期序列值。

参数：`value` 为任意值。

返回值：`boolean`。

示例代码：

```js
if (isDateSerial(45291.5)) {
    MsgBox("它可以当成日期序列处理");
}
```

示例代码完成的目的：在日期转换前先判断输入值是否适合按序列值处理。

### serialToDate(serial)

作用：把 WPS/Excel 日期序列值转换为 `Date`。

参数：`serial` 为数字或可转成数字的字符串。

返回值：`Date` 或 `null`。

示例代码：

```js
var dateValue = serialToDate(45291.5);
```

示例代码完成的目的：把单元格中的日期序列值转换成可继续格式化和比较的日期对象。

### dateToSerial(value)

作用：把日期值转换为 WPS/Excel 日期序列值。

参数：`value` 支持 `Date`、日期字符串或日期序列值。

返回值：`number`。失败时返回 `NaN`。

示例代码：

```js
var serial = dateToSerial("2026-04-09 08:30:00");
```

示例代码完成的目的：把外部输入的日期转成表格更容易直接写入的序列值。

### parseDateInput(value)

作用：把常见日期输入统一解析为 `Date`。

参数：`value` 支持 `Date`、日期序列值、`yyyy-MM-dd`、`yyyy/MM/dd`、`yyyyMMdd`、带时分秒的日期字符串等格式。

返回值：`Date` 或 `null`。

示例代码：

```js
var dateValue = parseDateInput("20260409 083000");
```

示例代码完成的目的：统一处理来自单元格、配置文件和手工输入的日期。

### formatDate(value, pattern)

作用：把日期值格式化为日期字符串。

参数：`value` 为日期输入值；`pattern` 为格式模板，默认 `yyyy-MM-dd`。

返回值：`string`。失败时返回空字符串。

示例代码：

```js
var text = formatDate(45291, "yyyy/MM/dd");
```

示例代码完成的目的：生成适合展示或拼接文件名的日期文本。

### formatDateTime(value, pattern)

作用：把日期值格式化为日期时间字符串。

参数：`value` 为日期输入值；`pattern` 默认 `yyyy-MM-dd HH:mm:ss`。

返回值：`string`。失败时返回空字符串。

示例代码：

```js
var text = formatDateTime(new Date(), "yyyy-MM-dd HH:mm:ss");
```

示例代码完成的目的：输出带时分秒的日志时间。

### startOfDay(value)

作用：获取指定日期当天的开始时刻。

参数：`value` 为日期输入值。

返回值：`Date` 或 `null`。

示例代码：

```js
var fromTime = startOfDay("2026-04-09");
```

示例代码完成的目的：构造按天统计时的开始时间边界。

### endOfDay(value)

作用：获取指定日期当天的结束时刻。

参数：`value` 为日期输入值。

返回值：`Date` 或 `null`。

示例代码：

```js
var toTime = endOfDay("2026-04-09");
```

示例代码完成的目的：构造按天统计时的结束时间边界。

### startOfMonth(value)

作用：获取指定日期当月第一天的开始时刻。

参数：`value` 为日期输入值。

返回值：`Date` 或 `null`。

示例代码：

```js
var monthStart = startOfMonth("2026-04-09");
```

示例代码完成的目的：生成月度报表的起始日期。

### endOfMonth(value)

作用：获取指定日期当月最后一天的结束时刻。

参数：`value` 为日期输入值。

返回值：`Date` 或 `null`。

示例代码：

```js
var monthEnd = endOfMonth("2026-04-09");
```

示例代码完成的目的：生成月度报表的结束日期。

### diffDays(startValue, endValue)

作用：计算两个日期之间相差的天数。

参数：`startValue` 和 `endValue` 都为日期输入值。

返回值：`number`。失败时返回 `NaN`。

示例代码：

```js
var days = diffDays("2026-04-01", "2026-04-09");
```

示例代码完成的目的：计算报表周期跨度。

### diffSeconds(startValue, endValue)

作用：计算两个日期之间相差的秒数。

参数：`startValue` 和 `endValue` 都为日期输入值。

返回值：`number`。失败时返回 `NaN`。

示例代码：

```js
var seconds = diffSeconds("2026-04-09 08:00:00", "2026-04-09 08:05:30");
```

示例代码完成的目的：计算宏执行耗时或接口请求间隔。

## 依赖说明

- 本模块不依赖 WPS 对象即可工作。
- 日期格式化和解析只使用 JScript 兼容语法，方便直接在 WPS JSA 宏环境中复用。
