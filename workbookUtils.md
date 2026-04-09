# workbookUtils

## 模块用途与适用场景

用于管理工作簿生命周期：获取应用对象、打开工作簿、另存为、关闭、复制工作表到新工作簿、控制 `DisplayAlerts`。

适用场景：
- 批量读取多个文件并汇总
- 模板复制后另存输出
- 自动化执行时关闭弹窗干扰

## 公共函数目录

- `getApp`
- `getActiveWorkbook`
- `openWorkbook`
- `saveWorkbookAs`
- `closeWorkbookSafe`
- `copyWorksheetToNewWorkbook`
- `setDisplayAlerts`

## 函数说明

### 1. getApp

作用：获取可用的应用对象。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| app | Object | 否 | 外部传入的应用对象 |

返回值：可用应用对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var app = getApp();
if (!app) {
    throw new Error("无法获取 Application");
}
```

示例目的：在宏入口处统一拿到应用对象。

### 2. getActiveWorkbook

作用：获取当前活动工作簿。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| app | Object | 否 | 应用对象 |

返回值：工作簿对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var wb = getActiveWorkbook();
if (wb) {
    MsgBox("当前工作簿：" + wb.Name);
}
```

示例目的：快速读取用户当前正在操作的工作簿。

### 3. openWorkbook

作用：按路径打开工作簿。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| path | string | 是 | 文件完整路径 |
| readOnly | boolean | 否 | 是否只读打开 |
| app | Object | 否 | 应用对象 |

返回值：打开后的工作簿对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var wb = openWorkbook("D:\\data\\report.xlsx", true);
if (!wb) {
    MsgBox("打开失败");
}
```

示例目的：读取外部报表文件而不修改原文件。

### 4. saveWorkbookAs

作用：将工作簿另存到目标路径。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| workbook | Object | 是 | 工作簿对象 |
| savePath | string | 是 | 另存路径 |
| fileFormat | number | 否 | 文件格式常量 |
| overwrite | boolean | 否 | 目标存在是否覆盖 |

返回值：是否保存成功。  
返回格式：`boolean`

示例代码：

```javascript
var ok = saveWorkbookAs(wb, "D:\\output\\new_report.xlsx", null, true);
```

示例目的：生成固定命名的输出文件并支持覆盖旧版本。

### 5. closeWorkbookSafe

作用：安全关闭工作簿。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| workbook | Object | 是 | 工作簿对象 |
| saveChanges | boolean | 否 | 是否保存后关闭 |

返回值：是否关闭成功。  
返回格式：`boolean`

示例代码：

```javascript
closeWorkbookSafe(wb, false);
```

示例目的：处理完临时工作簿后立即释放资源。

### 6. copyWorksheetToNewWorkbook

作用：复制工作表到新工作簿，并返回新工作簿对象。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| worksheet | Object | 是 | 源工作表 |
| newWorkbookName | string | 否 | 新工作簿标题 |
| app | Object | 否 | 应用对象 |

返回值：新工作簿对象，失败返回 `null`。  
返回格式：`Object|null`

示例代码：

```javascript
var newBook = copyWorksheetToNewWorkbook(wb.Worksheets("模板"), "日报副本");
```

示例目的：从模板页快速生成独立交付文件。

### 7. setDisplayAlerts

作用：设置 `DisplayAlerts` 并返回旧值。

参数：

| 参数名 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| app | Object | 否 | 应用对象 |
| value | boolean | 是 | 新值 |

返回值：修改前的值，失败返回 `null`。  
返回格式：`boolean|null`

示例代码：

```javascript
var oldAlerts = setDisplayAlerts(null, false);
try {
    // 批处理逻辑
} finally {
    if (oldAlerts !== null) {
        setDisplayAlerts(null, oldAlerts);
    }
}
```

示例目的：批量运行期间减少弹窗中断，结束后恢复现场。

## 依赖说明与 WPS-first / Windows fallback

- 优先使用 WPS/JSA 宏对象：`Application`、`Workbooks`、`Workbook`、`Worksheet`。
- 文件存在与删除在 WPS 内置失败时，回退 `Scripting.FileSystemObject`。
