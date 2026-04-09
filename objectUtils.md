# objectUtils.js 对象处理工具模块

`objectUtils.js` 用于安全读取和写入对象属性，以及做浅拷贝、深拷贝、键筛选等操作。

## 使用约定

- `safeGet` 优先用于读嵌套对象，避免逐层判空。
- `deepClone` 只保证普通对象、数组和 `Date` 的递归拷贝；不处理循环引用，也不克隆 WPS/ActiveX 特殊对象。

## 公开函数

- `safeGet(obj, path, defaultValue)`
- `safeSet(obj, path, value, createMissing)`
- `hasPath(obj, path)`
- `pickKeys(obj, keys)`
- `omitKeys(obj, keys)`
- `shallowClone(value)`
- `cloneArray(arrayValue, deep)`
- `deepClone(value)`

### safeGet(obj, path, defaultValue)

作用：安全读取对象中的嵌套属性。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `obj` | `*` | 要读取的对象。 |
| `path` | `string` 或 `Array<string>` | 路径字符串，如 `"a.b.c"`，或路径数组。 |
| `defaultValue` | `*` | 路径不存在时返回的默认值。 |

返回值：`*`。

示例代码：

```js
var city = safeGet(config, "user.address.city", "");
```

示例代码完成的目的：从复杂配置对象中安全读取城市字段，不需要层层判空。

### safeSet(obj, path, value, createMissing)

作用：安全写入对象中的嵌套属性。

参数：`obj` 为目标对象；`path` 为路径；`value` 为要写入的值；`createMissing` 表示中间节点不存在时是否自动创建。

返回值：`boolean`。

示例代码：

```js
safeSet(config, "output.folder.path", "D:\\report\\output", true);
```

示例代码完成的目的：按路径把输出目录写入配置对象。

### hasPath(obj, path)

作用：判断对象中是否存在指定路径。

参数：`obj` 为目标对象；`path` 为路径字符串或数组。

返回值：`boolean`。

示例代码：

```js
if (hasPath(config, "user.token")) {
    MsgBox("配置里有 token");
}
```

示例代码完成的目的：在读取配置前先判断关键字段是否存在。

### pickKeys(obj, keys)

作用：从对象中挑选指定键。

参数：`obj` 为原始对象；`keys` 为单个键或键数组。

返回值：`Object`。

返回格式：

```js
{
    name: "门店A",
    city: "上海"
}
```

示例代码：

```js
var simpleInfo = pickKeys(shopInfo, ["name", "city"]);
```

示例代码完成的目的：从大对象中抽取少量需要参与输出的字段。

### omitKeys(obj, keys)

作用：从对象中排除指定键。

参数：`obj` 为原始对象；`keys` 为单个键或键数组。

返回值：`Object`。

示例代码：

```js
var cleanInfo = omitKeys(shopInfo, ["secret", "token"]);
```

示例代码完成的目的：在导出对象前去掉敏感字段。

### shallowClone(value)

作用：对对象或数组做浅拷贝。

参数：`value` 为任意值。

返回值：`*`。对象和数组返回新的浅拷贝，其它值原样返回。

示例代码：

```js
var copied = shallowClone(rowInfo);
```

示例代码完成的目的：复制一份对象再做局部改动，避免直接影响原对象引用。

### cloneArray(arrayValue, deep)

作用：复制数组，可选择是否深拷贝数组元素。

参数：`arrayValue` 为数组；`deep` 为是否递归拷贝元素。

返回值：`Array`。

示例代码：

```js
var rows = cloneArray(sourceRows, true);
```

示例代码完成的目的：复制二维数据数组，避免排序或清洗时改动原始数据。

### deepClone(value)

作用：对普通对象、数组和 `Date` 做深拷贝。

参数：`value` 为任意值。

返回值：`*`。

示例代码：

```js
var copiedConfig = deepClone(config);
```

示例代码完成的目的：复制一份完整配置对象后再做修改，避免污染源配置。

## 依赖说明

- 本模块不依赖 WPS 运行时对象。
- 公开函数适合被配置读取、数据导入和 HTTP 响应处理等模块复用。
