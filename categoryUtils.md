# categoryUtils.js 分类处理工具模块

`categoryUtils.js` 用于拆分分类字符串、拼接分类信息，以及对分类值做稳定排序。

## 使用约定

- 默认按第一个 `-` 拆分分类字符串。
- 默认把分隔符右侧开头的连续数字识别为排序号，其余部分视为分类名称。
- 适合处理类似 `零食-12薯片`、`鞋服-2上衣` 这类业务分类值。

## 公开函数

- `splitCategory(category, separator)`
- `joinCategoryParts(group, order, name, separator)`
- `buildCategoryKey(category, separator)`
- `compareCategoryKey(left, right, separator)`
- `sortCategoryItems(items, getCategory, separator, desc)`

### splitCategory(category, separator)

作用：拆分分类字符串。

参数：`category` 为原始分类字符串；`separator` 为分隔符，默认 `-`。

返回值：对象，包含 `raw`、`group`、`order`、`name`。

返回格式：

```js
{
    raw: "零食-12薯片",
    group: "零食",
    order: 12,
    name: "薯片"
}
```

示例代码：

```js
var info = splitCategory("零食-12薯片");
```

示例代码完成的目的：把分类字符串拆成更适合排序和展示的结构。

### joinCategoryParts(group, order, name, separator)

作用：按组名、排序号和名称重新拼接分类字符串。

参数：`group` 为组名；`order` 为排序号；`name` 为名称；`separator` 为分隔符。

返回值：`string`。

示例代码：

```js
var category = joinCategoryParts("零食", 12, "薯片", "-");
```

示例代码完成的目的：根据拆分后的字段重新生成标准分类名称。

### buildCategoryKey(category, separator)

作用：生成用于排序和比较的分类键。

参数：`category` 为分类字符串或 `splitCategory` 结果对象；`separator` 为分隔符。

返回值：`string`。

示例代码：

```js
var key = buildCategoryKey("零食-12薯片");
```

示例代码完成的目的：为分类排序生成稳定、可比较的键值。

### compareCategoryKey(left, right, separator)

作用：比较两个分类值的顺序。

参数：`left`、`right` 为分类字符串或分类对象；`separator` 为分隔符。

返回值：`number`。左小于右返回 `-1`，相等返回 `0`，左大于右返回 `1`。

示例代码：

```js
var result = compareCategoryKey("零食-2糖果", "零食-12薯片");
```

示例代码完成的目的：让脚本按分类顺序做自定义排序，而不是只按字符串字典序。

### sortCategoryItems(items, getCategory, separator, desc)

作用：对分类数组或包含分类字段的对象数组排序。

参数：`items` 为原始数组；`getCategory` 为分类字段名或取值函数；`separator` 为分隔符；`desc` 为是否倒序。

返回值：`Array`。

示例代码：

```js
var sorted = sortCategoryItems(rows, "category", "-", false);
```

示例代码完成的目的：按分类规则对对象数组重新排序，适合生成有序报表。

## 依赖说明

- 本模块不依赖 WPS 对象。
- 常用于配置解析、数据导入分组和报表输出排序。
