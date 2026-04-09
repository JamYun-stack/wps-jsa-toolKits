# configSheetUtils

## 模块用途与适用场景

用于读取“配置页”中的键值、字段映射、路径配置和店铺分类配置。

适用场景：
- 从“操作”页或“配置”页读取运行参数
- 管理字段映射、输入输出路径
- 读取店铺/分类维度控制项

## 公共函数目录

- `readConfigValue`
- `readKeyValueConfig`
- `readFieldMapConfig`
- `readPathConfig`
- `readShopCategoryConfig`

## 函数说明

### 1. readConfigValue

作用：读取指定键名对应的单个配置值。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| key | string | 是 | 键名 |
| keyCol | number | 否 | 键列，默认 1 |
| valueCol | number | 否 | 值列，默认 2 |
| startRow | number | 否 | 起始行 |
| endRow | number | 否 | 结束行 |
| defaultValue | any | 否 | 未命中默认值 |

返回值：命中值或默认值。  
返回格式：`any`

示例代码：

```javascript
var tpl = readConfigValue(ws, "templatePath", 1, 2, 1, 100, "");
```

示例目的：单点读取关键配置项。

### 2. readKeyValueConfig

作用：读取键值配置区域为对象。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| keyCol | number | 否 | 键列 |
| valueCol | number | 否 | 值列 |
| startRow | number | 否 | 起始行 |
| endRow | number | 否 | 结束行 |

返回值：键值对象。  
返回格式：

```javascript
{
    "inputPath": "D:\\input",
    "outputPath": "D:\\output"
}
```

示例代码：

```javascript
var kv = readKeyValueConfig(ws, 1, 2, 1, 200);
```

示例目的：一次性读取全量配置再做后续分发。

### 3. readFieldMapConfig

作用：读取字段映射配置（字段名 -> 表头名）。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| fieldCol | number | 否 | 字段列 |
| headerCol | number | 否 | 表头列 |
| startRow | number | 否 | 起始行 |
| endRow | number | 否 | 结束行 |

返回值：映射对象。  
返回格式：

```javascript
{
    "shopName": "店铺名",
    "amount": "销售额"
}
```

示例代码：

```javascript
var fieldMap = readFieldMapConfig(ws, 1, 2, 2, 100);
```

示例目的：导入外部文件时做字段统一。

### 4. readPathConfig

作用：读取标准路径配置对象（输入/输出/模板）。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| startRow | number | 否 | 起始行 |
| endRow | number | 否 | 结束行 |

返回值：路径对象。  
返回格式：

```javascript
{
    inputPath: "D:\\input",
    outputPath: "D:\\output",
    templatePath: "D:\\tpl\\demo.xlsx"
}
```

示例代码：

```javascript
var pathConfig = readPathConfig(ws, 1, 100);
```

示例目的：统一输入/输出/模板路径来源。

### 5. readShopCategoryConfig

作用：读取店铺分类配置列表。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| startRow | number | 否 | 起始行 |
| endRow | number | 否 | 结束行 |
| shopCol | number | 否 | 店铺列 |
| categoryCol | number | 否 | 分类列 |
| valueCol | number | 否 | 值列 |

返回值：配置数组。  
返回格式：

```javascript
[
    { shop: "店铺A", category: "零食-100薯片", value: "Y" }
]
```

示例代码：

```javascript
var rows = readShopCategoryConfig(ws, 2, 200, 1, 2, 3);
```

示例目的：读取业务规则控制表用于过滤或打标。

## 依赖说明与 WPS-first / Windows fallback

- 完全基于 WPS/JSA 的单元格读取。
- 不依赖 Windows ActiveX。
