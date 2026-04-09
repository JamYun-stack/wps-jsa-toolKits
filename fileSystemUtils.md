# fileSystemUtils.js 文件与文件夹工具模块

`fileSystemUtils.js` 用于集中放置 WPS JSA 宏中常用的文件、文件夹和路径处理函数。

## 使用约定

- 路径按 Windows 路径处理，`/` 会在需要时转成 `\`。
- 大多数函数失败时不抛出异常，而是返回 `false`、`""`、`[]`、`{}` 或 `-1`，方便宏脚本继续做判断。
- `filterName`、`filterIgnore`、`filterExtend` 可以传字符串，也可以传数组；不传或传空值表示不启用该过滤条件。
- 实现上优先使用 WPS/JSA 宏环境已有的对象或兼容函数，例如 `Application.FileDialog`、`GetAttr`、`Dir`、`MkDir`、`FileLen`、`FileDateTime`、`FileCopy`、`Name`、`Kill`、`RmDir`；不可用或失败时再使用 Windows 的 `ActiveXObject` 作为兜底。
- 带有 `_fsu` 前缀的函数是模块内部辅助函数，不建议业务宏直接调用。

## 公开函数

### openFolderPicker()

作用：打开系统文件夹选择器，返回用户选择的文件夹路径。

参数：无。

返回值：`string`。用户取消或选择器失败时返回空字符串 `""`。

示例代码：

```js
var folderPath = openFolderPicker();

if (folderPath !== "") {
    MsgBox("已选择文件夹：" + folderPath);
}
```

示例代码完成的目的：让用户手动选择一个后续要扫描或保存文件的目录。

### openFilePicker(filterDescription, filterPattern, allowMultiSelect)

作用：打开系统文件选择器，返回用户选择的文件路径。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `filterDescription` | `string` | 文件类型说明，例如 `"Excel 文件"`。可为空。 |
| `filterPattern` | `string` | 文件匹配规则，例如 `"*.xlsx;*.xlsm"`。可为空，默认使用 `"*.*"`。 |
| `allowMultiSelect` | `boolean` | 是否允许多选。`true` 返回数组；其它值返回单个路径字符串。 |

返回值：单选时返回 `string`；多选时返回 `Array<string>`。取消或失败时单选返回 `""`，多选返回 `[]`。

示例代码：

```js
var filePath = openFilePicker("Excel 文件", "*.xlsx;*.xlsm", false);

if (filePath !== "") {
    MsgBox("准备处理文件：" + filePath);
}
```

示例代码完成的目的：让用户选择一个 Excel 文件，作为宏后续读取或处理的目标文件。

### fileExists(path)

作用：判断指定路径是否为已经存在的文件。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要检查的文件路径。 |

返回值：`boolean`。文件存在返回 `true`，不存在、路径为空或检查失败返回 `false`。

示例代码：

```js
var reportPath = "D:\\report\\result.xlsx";

if (fileExists(reportPath)) {
    MsgBox("结果文件已存在");
}
```

示例代码完成的目的：在生成文件前判断目标文件是否已经存在，避免误覆盖。

### folderExists(path)

作用：判断指定路径是否为已经存在的文件夹。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要检查的文件夹路径。 |

返回值：`boolean`。文件夹存在返回 `true`，不存在、路径为空或检查失败返回 `false`。

示例代码：

```js
var outputFolder = "D:\\report";

if (!folderExists(outputFolder)) {
    MsgBox("输出目录不存在");
}
```

示例代码完成的目的：在保存文件前确认输出目录是否可用。

### normalizePath(path)

作用：把路径中的 `/` 转为 Windows 风格的 `\`。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要标准化的路径。 |

返回值：`string`。空值会返回 `""`。

示例代码：

```js
var fixedPath = normalizePath("D:/report/result.xlsx");
MsgBox(fixedPath);
```

示例代码完成的目的：把用户输入或配置里的斜杠路径统一成 Windows 路径。

### joinPath()

作用：把多个路径片段拼接成一个 Windows 路径。

参数：可变参数，每个参数都是一个路径片段。

返回值：`string`。

示例代码：

```js
var filePath = joinPath("D:\\report", "2026", "result.xlsx");
MsgBox(filePath);
```

示例代码完成的目的：安全拼接目录和文件名，避免手写 `\` 时出现重复或遗漏。

### getPathName(path)

作用：获取路径最后一段名称。路径是文件时返回文件名，路径是文件夹时返回文件夹名。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 文件或文件夹路径。 |

返回值：`string`。

示例代码：

```js
var name = getPathName("D:\\report\\result.xlsx");
MsgBox(name);
```

示例代码完成的目的：从完整路径中提取文件名 `result.xlsx`。

### getParentFolderPath(path)

作用：获取文件或文件夹路径的父目录。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 文件或文件夹路径。 |

返回值：`string`。没有父目录时返回 `""`。

示例代码：

```js
var parentPath = getParentFolderPath("D:\\report\\result.xlsx");
MsgBox(parentPath);
```

示例代码完成的目的：从文件路径中提取保存目录 `D:\report`。

### getFileBaseName(fileName)

作用：获取不带扩展名的文件名。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `fileName` | `string` | 文件名或完整文件路径。 |

返回值：`string`。

示例代码：

```js
var baseName = getFileBaseName("D:\\report\\result.xlsx");
MsgBox(baseName);
```

示例代码完成的目的：从文件路径中提取基础文件名 `result`，用于生成新文件名。

### getFileExtend(fileName)

作用：获取文件扩展名，不包含点号。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `fileName` | `string` | 文件名或完整文件路径。 |

返回值：`string`。没有扩展名时返回 `""`。

示例代码：

```js
var extend = getFileExtend("D:\\report\\result.xlsx");
MsgBox(extend);
```

示例代码完成的目的：判断文件类型是否为 `xlsx`。

### changeFileExtend(path, newExtend)

作用：把文件路径的扩展名替换成新的扩展名。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 原始文件路径。 |
| `newExtend` | `string` | 新扩展名，可以写 `"csv"` 或 `".csv"`。传空字符串会移除扩展名。 |

返回值：`string`。

示例代码：

```js
var csvPath = changeFileExtend("D:\\report\\result.xlsx", "csv");
MsgBox(csvPath);
```

示例代码完成的目的：根据 Excel 文件路径生成同名 CSV 文件路径。

### createFolder(path)

作用：创建一个文件夹。要求父目录已经存在。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要创建的文件夹路径。 |

返回值：`boolean`。文件夹存在或创建成功返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = createFolder("D:\\report");

if (!ok) {
    MsgBox("创建目录失败");
}
```

示例代码完成的目的：创建一个直接输出目录。

### ensureFolder(path)

作用：创建文件夹以及缺失的父级文件夹。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要确保存在的文件夹路径。 |

返回值：`boolean`。最终目录存在返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = ensureFolder("D:\\report\\2026\\04");

if (ok) {
    MsgBox("目录已准备好");
}
```

示例代码完成的目的：保存月度报表前自动创建多级目录。

### ensureParentFolder(path)

作用：确保某个文件路径的父目录存在。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 文件路径。 |

返回值：`boolean`。父目录存在或创建成功返回 `true`，失败返回 `false`。

示例代码：

```js
var outputFile = "D:\\report\\2026\\04\\result.xlsx";

if (ensureParentFolder(outputFile)) {
    MsgBox("可以写入结果文件");
}
```

示例代码完成的目的：在保存文件前自动准备它所在的目录。

### getTempFolderPath()

作用：获取系统临时目录路径。

参数：无。

返回值：`string`。失败时返回 `""`。

示例代码：

```js
var tempPath = joinPath(getTempFolderPath(), "wps_macro_temp.xlsx");
MsgBox(tempPath);
```

示例代码完成的目的：生成一个位于系统临时目录中的临时文件路径。

### getFileSize(path)

作用：获取文件大小，单位为字节。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 文件路径。 |

返回值：`number`。读取失败或文件不存在时返回 `-1`。

示例代码：

```js
var size = getFileSize("D:\\report\\result.xlsx");

if (size > 0) {
    MsgBox("文件大小：" + size + " 字节");
}
```

示例代码完成的目的：检查结果文件是否已经生成且不是空文件。

### getFileModifiedTime(path)

作用：获取文件最后修改时间。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 文件路径。 |

返回值：可显示的日期时间值。读取失败或文件不存在时返回 `""`。

示例代码：

```js
var modifiedTime = getFileModifiedTime("D:\\report\\result.xlsx");

if (modifiedTime !== "") {
    MsgBox("最后修改时间：" + modifiedTime);
}
```

示例代码完成的目的：确认报表文件最近一次更新时间。

### copyFile(sourcePath, targetPath, overwrite)

作用：复制文件，并在需要时自动创建目标父目录。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `sourcePath` | `string` | 源文件路径。 |
| `targetPath` | `string` | 目标文件路径。 |
| `overwrite` | `boolean` | 目标文件存在时是否覆盖。只有 `true` 会覆盖。 |

返回值：`boolean`。复制成功返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = copyFile(
    "D:\\report\\result.xlsx",
    "D:\\backup\\result.xlsx",
    true
);
```

示例代码完成的目的：把报表文件复制到备份目录，并允许覆盖旧备份。

### moveFile(sourcePath, targetPath, overwrite)

作用：移动文件，并在需要时自动创建目标父目录。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `sourcePath` | `string` | 源文件路径。 |
| `targetPath` | `string` | 目标文件路径。 |
| `overwrite` | `boolean` | 目标文件存在时是否覆盖。只有 `true` 会先删除目标文件再移动。 |

返回值：`boolean`。移动成功返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = moveFile(
    "D:\\download\\result.xlsx",
    "D:\\report\\2026\\04\\result.xlsx",
    true
);
```

示例代码完成的目的：把下载目录中的报表移动到正式归档目录。

### deleteFile(path, force)

作用：删除文件。调用后文件不存在即返回成功。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要删除的文件路径。 |
| `force` | `boolean` | 是否强制删除只读文件。传 `false` 时不强制；其它值默认强制。 |

返回值：`boolean`。文件不存在或删除成功返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = deleteFile("D:\\report\\temp.xlsx", true);
```

示例代码完成的目的：清理宏运行过程中生成的临时文件。

### copyFolder(sourcePath, targetPath, overwrite)

作用：复制文件夹，并在需要时自动创建目标父目录。根目录不会被复制。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `sourcePath` | `string` | 源文件夹路径。 |
| `targetPath` | `string` | 目标文件夹路径。 |
| `overwrite` | `boolean` | 目标文件夹内存在同名文件时是否覆盖。只有 `true` 会覆盖。 |

返回值：`boolean`。复制成功返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = copyFolder("D:\\report\\2026", "D:\\backup\\report_2026", true);
```

示例代码完成的目的：把整年的报表目录复制到备份位置。

### moveFolder(sourcePath, targetPath, overwrite)

作用：移动文件夹，并在需要时自动创建目标父目录。根目录不会被移动。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `sourcePath` | `string` | 源文件夹路径。 |
| `targetPath` | `string` | 目标文件夹路径。 |
| `overwrite` | `boolean` | 目标文件夹存在时是否删除后移动。只有 `true` 会覆盖。 |

返回值：`boolean`。移动成功返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = moveFolder("D:\\temp\\report", "D:\\report\\2026\\04", true);
```

示例代码完成的目的：把临时目录中的报表文件夹移动到正式归档目录。

### deleteFolder(path, force)

作用：删除文件夹。根目录永远不会被删除。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要删除的文件夹路径。 |
| `force` | `boolean` | 是否强制删除只读内容。传 `false` 时不强制；其它值默认强制。 |

返回值：`boolean`。文件夹不存在或删除成功返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = deleteFolder("D:\\report\\temp", true);
```

示例代码完成的目的：清理宏运行后不再需要的临时目录。

### getFilesByPath(path, filterName, filterIgnore, filterExtend)

作用：获取指定目录下的直接子文件，并按名称关键字、忽略关键字和扩展名过滤。不递归子目录。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要扫描的文件夹路径。 |
| `filterName` | `string` 或 `Array<string>` | 文件名必须包含的关键字。为空表示不过滤。 |
| `filterIgnore` | `string` 或 `Array<string>` | 文件名中需要排除的关键字。为空表示不排除。 |
| `filterExtend` | `string` 或 `Array<string>` | 扩展名过滤，例如 `"xlsx"` 或 `["xlsx", "xlsm"]`。为空表示不过滤扩展名。 |

返回值：`Object`。键是文件名，值包含 `fileName`、`path`、`extend`。

返回格式：

```js
{
    "demo.xlsx": {
        fileName: "demo.xlsx",
        path: "D:\\test\\demo.xlsx",
        extend: "xlsx"
    }
}
```

示例代码：

```js
var files = getFilesByPath(
    "D:\\report",
    ["销售", "库存"],
    ["~$"],
    ["xlsx", "xlsm"]
);
```

示例代码完成的目的：找出报表目录中名称包含“销售”或“库存”的 Excel 文件，并排除 WPS/Excel 临时文件。

### listFilesByPath(path, filterName, filterIgnore, filterExtend)

作用：`getFilesByPath` 的别名，语义上更强调“列出文件”。

参数：同 `getFilesByPath(path, filterName, filterIgnore, filterExtend)`。

返回值：同 `getFilesByPath`。

示例代码：

```js
var files = listFilesByPath("D:\\report", null, ["~$"], "xlsx");
```

示例代码完成的目的：用更直观的函数名列出目录下的 Excel 文件。

### getFoldersByPath(path, filterName, filterIgnore)

作用：获取指定目录下的直接子文件夹，并按文件夹名称关键字过滤。不递归子目录。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要扫描的文件夹路径。 |
| `filterName` | `string` 或 `Array<string>` | 文件夹名必须包含的关键字。为空表示不过滤。 |
| `filterIgnore` | `string` 或 `Array<string>` | 文件夹名中需要排除的关键字。为空表示不排除。 |

返回值：`Object`。键是文件夹名，值包含 `folderName`、`path`。

示例代码：

```js
var folders = getFoldersByPath("D:\\report", ["2026"], ["temp"]);
```

示例代码完成的目的：列出报表目录中包含 `2026` 且不包含 `temp` 的子目录。

### listFoldersByPath(path, filterName, filterIgnore)

作用：`getFoldersByPath` 的别名，语义上更强调“列出文件夹”。

参数：同 `getFoldersByPath(path, filterName, filterIgnore)`。

返回值：同 `getFoldersByPath`。

示例代码：

```js
var folders = listFoldersByPath("D:\\report", null, ["temp"]);
```

示例代码完成的目的：用更直观的函数名列出目录下的子文件夹，并排除临时目录。

### getFolderModifiedTime(path)

作用：获取文件夹最后修改时间。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要读取时间的文件夹路径。 |

返回值：`Date` 或 `string`。读取成功时返回时间对象；文件夹不存在或读取失败时返回空字符串 `""`。

示例代码：

```js
var modifiedTime = getFolderModifiedTime("D:\\report\\2026");
```

示例代码完成的目的：在归档目录处理中判断文件夹最近一次被修改的时间。

### readTextFile(path, charset)

作用：读取文本文件内容，支持传入编码名称。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 文本文件路径。 |
| `charset` | `string` | 文本编码，例如 `"utf-8"`、`"unicode"`。为空时默认按 UTF-8 读取。 |

返回值：`string`。读取成功返回文本内容；文件不存在或读取失败时返回空字符串 `""`。

示例代码：

```js
var jsonText = readTextFile("D:\\report\\config.json", "utf-8");
```

示例代码完成的目的：读取配置文件或模板文本，供宏脚本进一步解析。

### writeTextFile(path, text, overwrite, charset)

作用：写入文本文件，并在需要时自动创建目标父目录。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 目标文本文件路径。 |
| `text` | `*` | 要写入的内容，会被转成字符串。 |
| `overwrite` | `boolean` | 目标文件已存在时是否覆盖。只有 `true` 会覆盖。 |
| `charset` | `string` | 文本编码，例如 `"utf-8"`、`"unicode"`。为空时默认按 UTF-8 写入。 |

返回值：`boolean`。写入成功返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = writeTextFile(
    "D:\\report\\logs\\run.txt",
    "任务完成",
    true,
    "utf-8"
);
```

示例代码完成的目的：把宏运行日志写入文本文件，并允许覆盖旧日志。

### readJsonFile(path, charset)

作用：读取 JSON 文件并解析为对象或数组。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | JSON 文件路径。 |
| `charset` | `string` | 文本编码，例如 `"utf-8"`。 |

返回值：`*`。解析成功时返回对象、数组或基础值；读取失败、文件为空或 JSON 解析失败时返回 `null`。

返回格式：

```js
{
    name: "演示配置",
    outputFolder: "D:\\report\\output"
}
```

示例代码：

```js
var config = readJsonFile("D:\\report\\config.json", "utf-8");
```

示例代码完成的目的：把外部配置文件解析成宏可以直接使用的对象。

### writeJsonFile(path, data, overwrite, indent, charset)

作用：把对象、数组或基础值写入 JSON 文件，并在需要时自动创建父目录。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 目标 JSON 文件路径。 |
| `data` | `*` | 要写入的对象、数组或基础值。 |
| `overwrite` | `boolean` | 目标文件存在时是否覆盖。只有 `true` 会覆盖。 |
| `indent` | `number` | JSON 缩进空格数。为空或非法时默认 `4`。 |
| `charset` | `string` | 文本编码，例如 `"utf-8"`。 |

返回值：`boolean`。写入成功返回 `true`，失败返回 `false`。

示例代码：

```js
var ok = writeJsonFile(
    "D:\\report\\output\\summary.json",
    { total: 12, success: true },
    true,
    4,
    "utf-8"
);
```

示例代码完成的目的：把汇总结果输出为 JSON 文件，方便其它宏或系统继续使用。

### appendBaseNameSuffix(path, suffix)

作用：在文件基础名后追加后缀，并保留原扩展名。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 原始文件路径。 |
| `suffix` | `string` | 要追加到基础名后的后缀，例如 `"_bak"`。 |

返回值：`string`。返回追加后的新路径；原始路径为空时返回空字符串 `""`。

示例代码：

```js
var backupPath = appendBaseNameSuffix("D:\\report\\demo.xlsx", "_bak");
```

示例代码完成的目的：为备份文件生成带后缀的新文件名。

### ensureUniqueFilePath(path, separator, startIndex)

作用：生成一个当前不存在的文件路径；如果原路径没有冲突，则直接返回原路径。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 原始目标文件路径。 |
| `separator` | `string` | 文件名和序号之间的分隔符，默认是 `"_"`。 |
| `startIndex` | `number` | 起始序号，默认从 `1` 开始。 |

返回值：`string`。返回不与现有文件或文件夹冲突的路径；原始路径为空时返回空字符串 `""`。

示例代码：

```js
var savePath = ensureUniqueFilePath("D:\\report\\result.xlsx", "_", 1);
```

示例代码完成的目的：避免导出文件时覆盖已存在的同名结果文件。

### isEmptyFolder(path)

作用：判断文件夹是否为空。目录下同时没有直接子文件和直接子文件夹时，才视为空目录。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要检查的文件夹路径。 |

返回值：`boolean`。目录存在且为空时返回 `true`，否则返回 `false`。

示例代码：

```js
if (isEmptyFolder("D:\\report\\temp")) {
    deleteFolder("D:\\report\\temp", true);
}
```

示例代码完成的目的：在删除目录前先确认它是否为空，避免误清理有内容的目录。

### walkFilesRecursive(path, filterName, filterIgnore, filterExtend)

作用：递归列出目录及其所有子目录中的文件，并支持名称、忽略关键字和扩展名过滤。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要扫描的根目录路径。 |
| `filterName` | `string` 或 `Array<string>` | 文件名必须包含的关键字。为空表示不过滤。 |
| `filterIgnore` | `string` 或 `Array<string>` | 文件名中需要排除的关键字。为空表示不排除。 |
| `filterExtend` | `string` 或 `Array<string>` | 扩展名过滤，例如 `"xlsx"` 或 `["xlsx", "xlsm"]`。 |

返回值：`Array<Object>`。返回文件信息数组；失败时返回空数组 `[]`。

返回格式：

```js
[
    {
        fileName: "demo.xlsx",
        path: "D:\\report\\2026\\demo.xlsx",
        extend: "xlsx"
    }
]
```

示例代码：

```js
var files = walkFilesRecursive(
    "D:\\report",
    ["销售"],
    ["~$"],
    ["xlsx", "xlsm"]
);
```

示例代码完成的目的：递归收集报表目录中所有正式的 Excel 报表文件。

### findFilesRecursive(path, filterName, filterIgnore, filterExtend)

作用：`walkFilesRecursive` 的别名，语义上更强调“递归查找文件”。

参数：同 `walkFilesRecursive(path, filterName, filterIgnore, filterExtend)`。

返回值：同 `walkFilesRecursive`。

示例代码：

```js
var files = findFilesRecursive("D:\\report", "日报", ["~$"], "xlsx");
```

示例代码完成的目的：用更接近业务语义的函数名递归搜索目标文件。

### walkFoldersRecursive(path, filterName, filterIgnore)

作用：递归列出目录及其所有子目录中的文件夹。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要扫描的根目录路径。 |
| `filterName` | `string` 或 `Array<string>` | 文件夹名必须包含的关键字。为空表示不过滤。 |
| `filterIgnore` | `string` 或 `Array<string>` | 文件夹名中需要排除的关键字。为空表示不排除。 |

返回值：`Array<Object>`。返回文件夹信息数组；失败时返回空数组 `[]`。

返回格式：

```js
[
    {
        folderName: "2026",
        path: "D:\\report\\2026"
    }
]
```

示例代码：

```js
var folders = walkFoldersRecursive("D:\\report", "2026", ["temp"]);
```

示例代码完成的目的：递归列出报表目录下所有正式归档子目录。

### findNewestFile(path, filterName, filterIgnore, filterExtend, includeSubFolders)

作用：在指定目录中查找最后修改时间最新的文件，可选择是否递归子目录。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要扫描的目录路径。 |
| `filterName` | `string` 或 `Array<string>` | 文件名必须包含的关键字。为空表示不过滤。 |
| `filterIgnore` | `string` 或 `Array<string>` | 文件名中需要排除的关键字。为空表示不排除。 |
| `filterExtend` | `string` 或 `Array<string>` | 扩展名过滤。为空表示不过滤。 |
| `includeSubFolders` | `boolean` | 是否递归扫描子目录；为 `true` 时递归。 |

返回值：`Object` 或 `null`。找到时返回最新文件对象；未找到时返回 `null`。

返回格式：

```js
{
    fileName: "demo.xlsx",
    path: "D:\\report\\demo.xlsx",
    extend: "xlsx",
    modifiedTime: "2026-04-09 10:30:00"
}
```

示例代码：

```js
var newest = findNewestFile(
    "D:\\report",
    ["销售"],
    ["~$"],
    ["xlsx", "xlsm"],
    true
);
```

示例代码完成的目的：找到报表目录及子目录中最新生成的一份销售报表。

### findNewestFileByPattern(path, pattern, includeSubFolders, filterExtend)

作用：按文件名关键字模式查找最新文件，是 `findNewestFile` 的快捷封装。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `path` | `string` | 要扫描的目录路径。 |
| `pattern` | `string` 或 `Array<string>` | 文件名必须包含的关键字模式。 |
| `includeSubFolders` | `boolean` | 是否递归扫描子目录；为 `true` 时递归。 |
| `filterExtend` | `string` 或 `Array<string>` | 扩展名过滤。为空表示不过滤。 |

返回值：`Object` 或 `null`。找到时返回最新文件对象；未找到时返回 `null`。

示例代码：

```js
var newest = findNewestFileByPattern(
    "D:\\report",
    ["日报", "店铺A"],
    true,
    ["xlsx"]
);
```

示例代码完成的目的：根据文件名关键字组合，快速找到最新的一份目标报表文件。

## 内部函数

以下函数只服务于本模块内部实现。业务宏优先调用上面的公开函数。

| 函数 | 作用 | 参数 |
| --- | --- | --- |
| `_fsuGetFilesByPathFallback(basePath, nameFilters, ignoreFilters, extendFilters)` | `Dir` 枚举失败时，用 `Scripting.FileSystemObject` 枚举文件。 | `basePath`：目录；`nameFilters`：名称过滤；`ignoreFilters`：忽略过滤；`extendFilters`：扩展名过滤。 |
| `_fsuGetFoldersByPathFallback(basePath, nameFilters, ignoreFilters)` | `Dir` 枚举失败时，用 `Scripting.FileSystemObject` 枚举子文件夹。 | `basePath`：目录；`nameFilters`：名称过滤；`ignoreFilters`：忽略过滤。 |
| `_fsuWalkFilesRecursiveByDir(basePath, nameFilters, ignoreFilters, extendFilters, result)` | 使用 `Dir` 递归枚举所有子文件。 | `basePath`：目录；`nameFilters`：名称过滤；`ignoreFilters`：忽略过滤；`extendFilters`：扩展名过滤；`result`：输出数组。 |
| `_fsuWalkFoldersRecursiveByDir(basePath, nameFilters, ignoreFilters, result)` | 使用 `Dir` 递归枚举所有子文件夹。 | `basePath`：目录；`nameFilters`：名称过滤；`ignoreFilters`：忽略过滤；`result`：输出数组。 |
| `_fsuWalkFilesRecursiveByFso(basePath, nameFilters, ignoreFilters, extendFilters)` | `Dir` 不可用时，用 `FileSystemObject` 递归枚举文件。 | `basePath`：目录；`nameFilters`：名称过滤；`ignoreFilters`：忽略过滤；`extendFilters`：扩展名过滤。 |
| `_fsuWalkFoldersRecursiveByFso(basePath, nameFilters, ignoreFilters)` | `Dir` 不可用时，用 `FileSystemObject` 递归枚举文件夹。 | `basePath`：目录；`nameFilters`：名称过滤；`ignoreFilters`：忽略过滤。 |
| `_fsuWalkFilesRecursiveByFsoFolder(folder, nameFilters, ignoreFilters, extendFilters, result)` | 递归处理单个 FSO 文件夹对象下的文件。 | `folder`：FSO 文件夹对象；`nameFilters`：名称过滤；`ignoreFilters`：忽略过滤；`extendFilters`：扩展名过滤；`result`：输出数组。 |
| `_fsuWalkFoldersRecursiveByFsoFolder(folder, nameFilters, ignoreFilters, result)` | 递归处理单个 FSO 文件夹对象下的子文件夹。 | `folder`：FSO 文件夹对象；`nameFilters`：名称过滤；`ignoreFilters`：忽略过滤；`result`：输出数组。 |
| `_fsuListEntriesByDir(basePath)` | 用 `Dir` 一次性列出目录下的文件和文件夹条目。 | `basePath`：目录。 |
| `_fsuNormalizeCharset(charset)` | 统一文本流使用的编码名称。 | `charset`：编码字符串。 |
| `_fsuPadNumber(value, size)` | 按固定长度左侧补零。 | `value`：数字；`size`：总位数。 |
| `_fsuGetObjectKeys(obj)` | 获取对象自身可枚举键名数组。 | `obj`：普通对象。 |
| `_fsuObjectToArray(obj)` | 把“以名称为键”的对象转换成值数组。 | `obj`：普通对象。 |
| `_fsuDateToTimestamp(value)` | 把 `Date` 或可解析日期转成时间戳。 | `value`：日期值。 |
| `_fsuNormalizePath(path)` | 把路径中的 `/` 转为 `\`。 | `path`：路径字符串。 |
| `_fsuTrimRightSlash(path)` | 去掉路径右侧多余斜杠，但保留根路径斜杠。 | `path`：路径字符串。 |
| `_fsuTrimLeftSlash(path)` | 去掉路径左侧斜杠，用于拼接路径片段。 | `path`：路径字符串。 |
| `_fsuJoinTwoPathParts(left, right)` | 拼接两个路径片段。 | `left`：左侧路径；`right`：右侧路径。 |
| `_fsuIsRootPath(path)` | 判断路径是否为盘符根目录、`\` 或 UNC 根目录。 | `path`：路径字符串。 |
| `_fsuIsDirectoryAttr(attr)` | 判断 `GetAttr` 返回值是否包含目录标记。 | `attr`：文件属性数字。 |
| `_fsuIsFilePath(path)` | 用 `GetAttr` 判断路径是否为文件。 | `path`：路径字符串。 |
| `_fsuIsFolderPath(path)` | 用 `GetAttr` 判断路径是否为文件夹。 | `path`：路径字符串。 |
| `_fsuToStringArray(value)` | 把空值、字符串或数组统一转成字符串数组。 | `value`：待转换值。 |
| `_fsuToLowerStringArray(value)` | 把过滤值统一转成小写字符串数组，并去掉扩展名前导点号。 | `value`：待转换值。 |
| `_fsuGetFileExtend(fileName)` | 获取扩展名，不包含点号。 | `fileName`：文件名。 |
| `_fsuNameMatched(fileName, nameFilters)` | 判断文件名是否命中名称过滤。 | `fileName`：文件名；`nameFilters`：名称过滤数组。 |
| `_fsuIgnoreMatched(fileName, ignoreFilters)` | 判断文件名是否命中忽略过滤。 | `fileName`：文件名；`ignoreFilters`：忽略过滤数组。 |
| `_fsuExtendMatched(extend, extendFilters)` | 判断扩展名是否命中过滤条件。 | `extend`：扩展名；`extendFilters`：扩展名过滤数组。 |

内部函数调用示例：

```js
// getFilesByPath 内部会把 ["xlsx", ".xlsm"] 统一成 ["xlsx", "xlsm"]，
// 然后再通过 _fsuExtendMatched 判断文件扩展名是否命中。
var files = getFilesByPath("D:\\report", null, ["~$"], ["xlsx", ".xlsm"]);
```

示例代码完成的目的：展示内部过滤函数如何服务于公开的文件枚举函数；业务代码不需要直接调用 `_fsu` 函数。
