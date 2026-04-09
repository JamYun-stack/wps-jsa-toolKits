# shapeUtils

## 模块用途与适用场景

用于形状与图片处理，包括新增形状、移动、缩放、分组和插入图片。

适用场景：
- 报表封面自动放置标注框
- 动态调整图标位置和大小
- 插入本地图片作为说明或徽标

## 公共函数目录

- `addShape`
- `moveShape`
- `resizeShape`
- `groupShapes`
- `insertImageFromFile`

## 函数说明

### 1. addShape

作用：新增形状并可选写入文本。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| type | number | 是 | 形状类型常量 |
| left | number | 是 | 左位置 |
| top | number | 是 | 上位置 |
| width | number | 是 | 宽度 |
| height | number | 是 | 高度 |
| text | string | 否 | 文本内容 |

返回值：形状对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var shp = addShape(ws, 1, 80, 60, 180, 40, "月报");
```

示例目的：在报表顶部自动添加标题形状。

### 2. moveShape

作用：移动形状位置。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| shape | Object | 是 | 形状对象 |
| left | number | 是 | 左位置 |
| top | number | 是 | 上位置 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
moveShape(shp, 100, 100);
```

示例目的：根据内容区域重新布局图形位置。

### 3. resizeShape

作用：调整形状尺寸并可选锁定纵横比。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| shape | Object | 是 | 形状对象 |
| width | number | 是 | 宽度 |
| height | number | 是 | 高度 |
| lockAspectRatio | boolean | 否 | 是否锁定比例 |

返回值：是否成功。  
返回格式：`boolean`

示例代码：

```javascript
resizeShape(shp, 220, 60, false);
```

示例目的：统一图形尺寸风格。

### 4. groupShapes

作用：按名称将多个形状分组。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| shapeNames | Array | 是 | 形状名称数组 |

返回值：分组对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var grp = groupShapes(ws, ["矩形 1", "文本框 2"]);
```

示例目的：将相关元素绑定，便于整体移动。

### 5. insertImageFromFile

作用：从本地文件插入图片。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 工作表对象 |
| imagePath | string | 是 | 图片路径 |
| left | number | 是 | 左位置 |
| top | number | 是 | 上位置 |
| width | number | 否 | 宽度 |
| height | number | 否 | 高度 |
| linkToFile | boolean | 否 | 是否链接文件 |
| saveWithDocument | boolean | 否 | 是否随文档保存 |

返回值：图片形状对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
insertImageFromFile(ws, "D:\\assets\\logo.png", 20, 12, 120, 40, false, true);
```

示例目的：自动插入品牌 logo 或结果截图。

## 依赖说明与 WPS-first / Windows fallback

- 完全基于 WPS/JSA 的 `Shapes`、`Pictures` 对象。
- 不依赖 Windows ActiveX。
