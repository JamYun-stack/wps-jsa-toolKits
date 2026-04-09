# httpUtils.js HTTP 与下载工具模块

`httpUtils.js` 用于在 WPS JSA 宏环境里发起通用 HTTP 请求，并完成文本、JSON、普通文件和图片的下载。

## 使用约定

- 网络请求优先使用 `WinHttp.WinHttpRequest.5.1`，不可用时退回 `MSXML2.XMLHTTP`。
- 二进制落盘使用 `ADODB.Stream`。
- 如果已加载 [fileSystemUtils.js](F:/DATA/3_data_analysis/wps-jsa-toolKits/fileSystemUtils.js)，会优先复用其中的路径和写文件能力；否则使用本模块内置的 ActiveX 兜底实现。

## 公开函数

- `buildQueryString(params)`
- `parseResponseHeaders(headerText)`
- `httpRequest(method, url, options)`
- `httpGet(url, params, headers, timeout)`
- `httpPostJson(url, data, headers, timeout)`
- `httpPostForm(url, data, headers, timeout)`
- `downloadFile(url, targetPath, headers, cookies, timeout, autoRename)`
- `downloadImage(url, targetPath, headers, cookies, timeout, autoRename)`
- `downloadText(url, targetPath, headers, cookies, timeout, charset, overwrite)`
- `downloadJson(url, targetPath, headers, cookies, timeout, overwrite, indent, charset)`

### buildQueryString(params)

作用：把对象参数拼成查询字符串。

参数：`params` 为普通对象，值可以是单值或数组。

返回值：`string`。不包含前导 `?`。

示例代码：

```js
var query = buildQueryString({ page: 1, keyword: "销售" });
```

示例代码完成的目的：把接口参数转成 URL 查询串。

### parseResponseHeaders(headerText)

作用：把原始响应头文本解析为对象。

参数：`headerText` 为 HTTP 返回的完整响应头文本。

返回值：`Object`，键统一转为小写。

返回格式：

```js
{
    "content-type": "application/json",
    "content-length": "123"
}
```

示例代码：

```js
var headers = parseResponseHeaders("Content-Type: application/json\r\nContent-Length: 123");
```

示例代码完成的目的：方便脚本按键名读取响应头信息。

### httpRequest(method, url, options)

作用：发送通用 HTTP 请求。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `method` | `string` | 请求方法，例如 `"GET"`、`"POST"`。 |
| `url` | `string` | 请求地址。 |
| `options` | `Object` | 可选项，支持 `query`、`headers`、`cookies`、`timeout`、`body`、`data`、`contentType`、`responseType`。 |

返回值：请求结果对象，包含状态码、响应头、文本、JSON 解析结果等。

返回格式：

```js
{
    success: true,
    method: "GET",
    url: "https://example.com/api?a=1",
    status: 200,
    statusText: "OK",
    headers: { "content-type": "application/json" },
    rawHeaders: "Content-Type: application/json",
    text: "{\"ok\":true}",
    json: { ok: true },
    body: "{\"ok\":true}"
}
```

示例代码：

```js
var response = httpRequest("GET", "https://example.com/api", {
    query: { page: 1 },
    headers: { Accept: "application/json" },
    responseType: "json",
    timeout: 10000
});
```

示例代码完成的目的：用统一入口发起请求并读取结构化结果。

### httpGet(url, params, headers, timeout)

作用：发送 GET 请求。

参数：`url` 为请求地址；`params` 为查询参数；`headers` 为请求头；`timeout` 为毫秒级超时。

返回值：同 `httpRequest`。

示例代码：

```js
var response = httpGet("https://example.com/list", { page: 1 }, null, 8000);
```

示例代码完成的目的：快速读取一个 GET 接口。

### httpPostJson(url, data, headers, timeout)

作用：发送 JSON POST 请求。

参数：`url` 为请求地址；`data` 为 JSON 数据；`headers` 为附加请求头；`timeout` 为毫秒级超时。

返回值：同 `httpRequest`。

示例代码：

```js
var response = httpPostJson("https://example.com/save", { name: "门店A" }, null, 10000);
```

示例代码完成的目的：向接口提交 JSON 结构数据。

### httpPostForm(url, data, headers, timeout)

作用：发送 `application/x-www-form-urlencoded` 表单请求。

参数：`url` 为请求地址；`data` 为表单对象；`headers` 为附加请求头；`timeout` 为毫秒级超时。

返回值：同 `httpRequest`。

示例代码：

```js
var response = httpPostForm("https://example.com/login", {
    username: "demo",
    password: "123456"
}, null, 10000);
```

示例代码完成的目的：调用需要传统表单提交的接口。

### downloadFile(url, targetPath, headers, cookies, timeout, autoRename)

作用：下载普通二进制文件。

参数：`url` 为下载地址；`targetPath` 为保存路径；`headers` 为请求头；`cookies` 为 Cookie；`timeout` 为毫秒级超时；`autoRename` 表示目标文件已存在时是否自动改名。

返回值：下载结果对象。

返回格式：

```js
{
    success: true,
    status: 200,
    statusText: "OK",
    url: "https://example.com/demo.xlsx",
    path: "D:\\temp\\demo.xlsx",
    fileName: "demo.xlsx"
}
```

示例代码：

```js
var result = downloadFile(
    "https://example.com/files/demo.xlsx",
    "D:\\download\\demo.xlsx",
    null,
    null,
    15000,
    true
);
```

示例代码完成的目的：把远程文件直接下载到本地目录，并避免覆盖同名文件。

### downloadImage(url, targetPath, headers, cookies, timeout, autoRename)

作用：下载图片文件。

参数：与 `downloadFile` 相同。

返回值：同 `downloadFile`。

示例代码：

```js
var result = downloadImage(
    "https://example.com/images/a.png",
    "D:\\images\\a.png",
    null,
    null,
    10000,
    true
);
```

示例代码完成的目的：把远程图片下载到本地，供后续插入表格或缓存。

### downloadText(url, targetPath, headers, cookies, timeout, charset, overwrite)

作用：下载文本内容，并可选写入文件。

参数：`url` 为下载地址；`targetPath` 为空时只返回文本；其它参数分别控制请求和写文件行为。

返回值：结果对象，包含 `text` 字段。

示例代码：

```js
var result = downloadText(
    "https://example.com/readme.txt",
    "D:\\download\\readme.txt",
    null,
    null,
    10000,
    "utf-8",
    true
);
```

示例代码完成的目的：下载接口返回的文本，并保存在本地供后续分析。

### downloadJson(url, targetPath, headers, cookies, timeout, overwrite, indent, charset)

作用：下载 JSON 内容，并可选写入文件。

参数：`url` 为下载地址；`targetPath` 为空时只返回解析结果；其余参数控制请求、保存和 JSON 格式。

返回值：结果对象，包含 `data` 字段。

返回格式：

```js
{
    success: true,
    status: 200,
    statusText: "OK",
    url: "https://example.com/config.json",
    path: "D:\\download\\config.json",
    data: { ok: true }
}
```

示例代码：

```js
var result = downloadJson(
    "https://example.com/config.json",
    "D:\\download\\config.json",
    null,
    null,
    10000,
    true,
    4,
    "utf-8"
);
```

示例代码完成的目的：下载 JSON 配置并同时落盘缓存。

## 依赖说明

- 网络层使用 Windows ActiveX 对象：`WinHttp.WinHttpRequest.5.1`、`MSXML2.XMLHTTP`、`ADODB.Stream`。
- 文件落盘时会优先复用 [fileSystemUtils.js](F:/DATA/3_data_analysis/wps-jsa-toolKits/fileSystemUtils.js) 的 `ensureParentFolder`、`ensureUniqueFilePath`、`writeTextFile`、`writeJsonFile` 等函数。
