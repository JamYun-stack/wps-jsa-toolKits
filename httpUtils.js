/**
 * 把对象参数拼接为查询字符串。
 *
 * @param {Object} params 查询参数对象。
 * @returns {string} 查询字符串，不包含前导 `?`。
 */
function buildQueryString(params) {
    var parts = [];
    var key = "";
    var i = 0;
    var value = null;

    if (!params || typeof params !== "object") {
        return "";
    }

    for (key in params) {
        if (params.hasOwnProperty(key)) {
            value = params[key];

            if (value instanceof Array) {
                for (i = 0; i < value.length; i++) {
                    parts.push(_huEncode(key) + "=" + _huEncode(value[i]));
                }
            } else if (value !== undefined) {
                parts.push(_huEncode(key) + "=" + _huEncode(value === null ? "" : value));
            }
        }
    }

    return parts.join("&");
}

/**
 * 解析响应头文本。
 *
 * @param {string} headerText 原始响应头文本。
 * @returns {Object.<string, string>} 解析后的响应头对象；键统一转为小写。
 *
 * 返回格式：
 * {
 *   "content-type": "application/json",
 *   "content-length": "123"
 * }
 */
function parseResponseHeaders(headerText) {
    var result = {};
    var lines = [];
    var i = 0;
    var line = "";
    var index = -1;
    var key = "";
    var value = "";

    if (!headerText) {
        return result;
    }

    lines = String(headerText).replace(/\r/g, "").split("\n");

    for (i = 0; i < lines.length; i++) {
        line = lines[i];

        if (line === "") {
            continue;
        }

        index = line.indexOf(":");
        if (index <= 0) {
            continue;
        }

        key = String(line.substring(0, index)).toLowerCase();
        value = String(line.substring(index + 1)).replace(/^\s+|\s+$/g, "");
        result[key] = value;
    }

    return result;
}

/**
 * 发送通用 HTTP 请求。
 *
 * @param {string} method 请求方法，例如 `"GET"`、`"POST"`。
 * @param {string} url 请求地址。
 * @param {Object} [options] 请求选项。
 * @returns {{success: boolean, method: string, url: string, status: number, statusText: string, headers: Object, rawHeaders: string, text: string, json: *, body: *}} 请求结果对象。
 *
 * 返回格式：
 * {
 *   success: true,
 *   method: "GET",
 *   url: "https://example.com/api?a=1",
 *   status: 200,
 *   statusText: "OK",
 *   headers: { "content-type": "application/json" },
 *   rawHeaders: "Content-Type: application/json",
 *   text: "{\"ok\":true}",
 *   json: { ok: true },
 *   body: "{\"ok\":true}"
 * }
 */
function httpRequest(method, url, options) {
    var requestOptions = options || {};
    var queryString = buildQueryString(requestOptions.query);
    var finalUrl = String(url || "");
    var request = null;
    var requestMethod = String(method || "GET").toUpperCase();
    var headers = requestOptions.headers || {};
    var cookies = _huCookiesToHeader(requestOptions.cookies);
    var timeout = requestOptions.timeout;
    var body = requestOptions.body;
    var contentType = requestOptions.contentType || "";
    var responseType = String(requestOptions.responseType || "text").toLowerCase();
    var result = {
        success: false,
        method: requestMethod,
        url: finalUrl,
        status: 0,
        statusText: "",
        headers: {},
        rawHeaders: "",
        text: "",
        json: null,
        body: null
    };

    if (finalUrl === "") {
        return result;
    }

    if (queryString !== "") {
        finalUrl = finalUrl + (finalUrl.indexOf("?") >= 0 ? "&" : "?") + queryString;
    }

    result.url = finalUrl;

    if (body === undefined && requestOptions.data !== undefined) {
        if (contentType === "application/json") {
            if (typeof JSON !== "undefined" && JSON.stringify) {
                body = JSON.stringify(requestOptions.data);
            } else {
                body = String(requestOptions.data);
            }
        } else if (contentType === "application/x-www-form-urlencoded") {
            body = buildQueryString(requestOptions.data);
        } else {
            body = requestOptions.data;
        }
    }

    request = _huCreateRequest();
    if (!request) {
        return result;
    }

    try {
        request.Open(requestMethod, finalUrl, false);
        _huSetTimeouts(request, timeout);

        if (cookies !== "") {
            headers.Cookie = cookies;
        }

        _huApplyHeaders(request, headers);

        if (contentType !== "") {
            try {
                request.SetRequestHeader("Content-Type", contentType);
            } catch (contentTypeError) {
            }
        }

        if (body === undefined || body === null) {
            request.Send();
        } else {
            request.Send(body);
        }

        result.status = Number(request.Status || 0);
        result.statusText = String(request.StatusText || "");
        result.rawHeaders = _huGetAllResponseHeaders(request);
        result.headers = parseResponseHeaders(result.rawHeaders);
        result.text = _huTryReadResponseText(request);
        result.body = responseType === "binary" ? _huTryReadResponseBody(request) : result.text;
        result.success = result.status >= 200 && result.status < 300;

        if (responseType === "json" || _huIsJsonResponse(result.headers["content-type"])) {
            try {
                result.json = result.text === "" ? null : JSON.parse(result.text);
            } catch (jsonError) {
                result.json = null;
            }
        }
    } catch (error) {
        result.statusText = result.statusText || String(error.message || error.description || error);
    }

    return result;
}

/**
 * 发送 GET 请求。
 *
 * @param {string} url 请求地址。
 * @param {Object} [params] 查询参数对象。
 * @param {Object} [headers] 请求头对象。
 * @param {number} [timeout] 超时时间，单位毫秒。
 * @returns {{success: boolean, method: string, url: string, status: number, statusText: string, headers: Object, rawHeaders: string, text: string, json: *, body: *}} 请求结果对象。
 */
function httpGet(url, params, headers, timeout) {
    return httpRequest("GET", url, {
        query: params,
        headers: headers,
        timeout: timeout
    });
}

/**
 * 发送 JSON POST 请求。
 *
 * @param {string} url 请求地址。
 * @param {*} data 要提交的 JSON 数据。
 * @param {Object} [headers] 请求头对象。
 * @param {number} [timeout] 超时时间，单位毫秒。
 * @returns {{success: boolean, method: string, url: string, status: number, statusText: string, headers: Object, rawHeaders: string, text: string, json: *, body: *}} 请求结果对象。
 */
function httpPostJson(url, data, headers, timeout) {
    return httpRequest("POST", url, {
        data: data,
        headers: headers,
        timeout: timeout,
        contentType: "application/json",
        responseType: "json"
    });
}

/**
 * 发送表单 POST 请求。
 *
 * @param {string} url 请求地址。
 * @param {Object} data 表单数据对象。
 * @param {Object} [headers] 请求头对象。
 * @param {number} [timeout] 超时时间，单位毫秒。
 * @returns {{success: boolean, method: string, url: string, status: number, statusText: string, headers: Object, rawHeaders: string, text: string, json: *, body: *}} 请求结果对象。
 */
function httpPostForm(url, data, headers, timeout) {
    return httpRequest("POST", url, {
        data: data,
        headers: headers,
        timeout: timeout,
        contentType: "application/x-www-form-urlencoded"
    });
}

/**
 * 下载二进制文件。
 *
 * @param {string} url 下载地址。
 * @param {string} [targetPath] 目标文件路径；为空时自动按 URL 文件名保存到临时目录。
 * @param {Object} [headers] 请求头对象。
 * @param {string|Object} [cookies] Cookie 字符串或对象。
 * @param {number} [timeout] 超时时间，单位毫秒。
 * @param {boolean} [autoRename] 目标文件存在时是否自动改名；为 `true` 时自动生成不冲突路径。
 * @returns {{success: boolean, status: number, statusText: string, url: string, path: string, fileName: string}} 下载结果对象。
 *
 * 返回格式：
 * {
 *   success: true,
 *   status: 200,
 *   statusText: "OK",
 *   url: "https://example.com/demo.xlsx",
 *   path: "D:\\temp\\demo.xlsx",
 *   fileName: "demo.xlsx"
 * }
 */
function downloadFile(url, targetPath, headers, cookies, timeout, autoRename) {
    var response = httpRequest("GET", url, {
        headers: headers,
        cookies: cookies,
        timeout: timeout,
        responseType: "binary"
    });
    var savePath = "";
    var result = {
        success: false,
        status: response.status,
        statusText: response.statusText,
        url: String(url || ""),
        path: "",
        fileName: ""
    };

    if (!response.success || !response.body) {
        return result;
    }

    savePath = _huResolveDownloadPath(url, targetPath, autoRename);
    if (savePath === "") {
        return result;
    }

    if (!_huWriteBinary(savePath, response.body)) {
        return result;
    }

    result.success = true;
    result.path = savePath;
    result.fileName = _huGetPathName(savePath);
    return result;
}

/**
 * 下载图片文件。
 *
 * @param {string} url 图片地址。
 * @param {string} [targetPath] 目标文件路径。
 * @param {Object} [headers] 请求头对象。
 * @param {string|Object} [cookies] Cookie 字符串或对象。
 * @param {number} [timeout] 超时时间，单位毫秒。
 * @param {boolean} [autoRename] 目标文件存在时是否自动改名。
 * @returns {{success: boolean, status: number, statusText: string, url: string, path: string, fileName: string}} 下载结果对象。
 */
function downloadImage(url, targetPath, headers, cookies, timeout, autoRename) {
    return downloadFile(url, targetPath, headers, cookies, timeout, autoRename);
}

/**
 * 下载文本内容，并可选地落盘保存。
 *
 * @param {string} url 下载地址。
 * @param {string} [targetPath] 目标文件路径；不传时只返回文本。
 * @param {Object} [headers] 请求头对象。
 * @param {string|Object} [cookies] Cookie 字符串或对象。
 * @param {number} [timeout] 超时时间，单位毫秒。
 * @param {string} [charset] 文本编码，例如 `"utf-8"`。
 * @param {boolean} [overwrite] 目标文件存在时是否覆盖。
 * @returns {{success: boolean, status: number, statusText: string, url: string, path: string, text: string}} 下载结果对象。
 */
function downloadText(url, targetPath, headers, cookies, timeout, charset, overwrite) {
    var response = httpRequest("GET", url, {
        headers: headers,
        cookies: cookies,
        timeout: timeout,
        responseType: "text"
    });
    var result = {
        success: false,
        status: response.status,
        statusText: response.statusText,
        url: String(url || ""),
        path: "",
        text: response.text
    };

    if (!response.success) {
        return result;
    }

    if (targetPath) {
        if (!_huWriteText(targetPath, response.text, charset, overwrite)) {
            return result;
        }
        result.path = String(targetPath);
    }

    result.success = true;
    return result;
}

/**
 * 下载 JSON 内容，并可选地落盘保存。
 *
 * @param {string} url 下载地址。
 * @param {string} [targetPath] 目标文件路径；不传时只返回解析结果。
 * @param {Object} [headers] 请求头对象。
 * @param {string|Object} [cookies] Cookie 字符串或对象。
 * @param {number} [timeout] 超时时间，单位毫秒。
 * @param {boolean} [overwrite] 目标文件存在时是否覆盖。
 * @param {number} [indent] JSON 写入缩进空格数。
 * @param {string} [charset] 文本编码，例如 `"utf-8"`。
 * @returns {{success: boolean, status: number, statusText: string, url: string, path: string, data: *}} 下载结果对象。
 *
 * 返回格式：
 * {
 *   success: true,
 *   status: 200,
 *   statusText: "OK",
 *   url: "https://example.com/config.json",
 *   path: "D:\\temp\\config.json",
 *   data: { ok: true }
 * }
 */
function downloadJson(url, targetPath, headers, cookies, timeout, overwrite, indent, charset) {
    var response = httpRequest("GET", url, {
        headers: headers,
        cookies: cookies,
        timeout: timeout,
        responseType: "json"
    });
    var result = {
        success: false,
        status: response.status,
        statusText: response.statusText,
        url: String(url || ""),
        path: "",
        data: response.json
    };

    if (!response.success || response.json === null) {
        return result;
    }

    if (targetPath) {
        if (!_huWriteJson(targetPath, response.json, overwrite, indent, charset)) {
            return result;
        }
        result.path = String(targetPath);
    }

    result.success = true;
    return result;
}

/**
 * 创建 HTTP 请求对象。
 *
 * @private
 * @returns {Object|null} ActiveX 请求对象。
 */
function _huCreateRequest() {
    try {
        return new ActiveXObject("WinHttp.WinHttpRequest.5.1");
    } catch (error) {
    }

    try {
        return new ActiveXObject("MSXML2.XMLHTTP");
    } catch (fallbackError) {
        return null;
    }
}

/**
 * 设置超时时间。
 *
 * @private
 * @param {Object} request 请求对象。
 * @param {number} timeout 超时时间，单位毫秒。
 * @returns {void}
 */
function _huSetTimeouts(request, timeout) {
    var timeoutValue = Number(timeout);

    if (!timeoutValue || !isFinite(timeoutValue) || timeoutValue < 0) {
        return;
    }

    try {
        if (typeof request.SetTimeouts === "function") {
            request.SetTimeouts(timeoutValue, timeoutValue, timeoutValue, timeoutValue);
        }
    } catch (error) {
    }
}

/**
 * 应用请求头。
 *
 * @private
 * @param {Object} request 请求对象。
 * @param {Object} headers 请求头对象。
 * @returns {void}
 */
function _huApplyHeaders(request, headers) {
    var key = "";

    if (!headers || typeof headers !== "object") {
        return;
    }

    for (key in headers) {
        if (headers.hasOwnProperty(key) && headers[key] !== undefined) {
            try {
                request.SetRequestHeader(String(key), String(headers[key]));
            } catch (error) {
            }
        }
    }
}

/**
 * 把 Cookie 输入值统一转为请求头字符串。
 *
 * @private
 * @param {string|Object} cookies Cookie 字符串或对象。
 * @returns {string} Cookie 请求头值。
 */
function _huCookiesToHeader(cookies) {
    var parts = [];
    var key = "";

    if (!cookies) {
        return "";
    }

    if (typeof cookies === "string") {
        return cookies;
    }

    if (typeof cookies !== "object") {
        return "";
    }

    for (key in cookies) {
        if (cookies.hasOwnProperty(key)) {
            parts.push(String(key) + "=" + String(cookies[key]));
        }
    }

    return parts.join("; ");
}

/**
 * 获取全部响应头。
 *
 * @private
 * @param {Object} request 请求对象。
 * @returns {string} 原始响应头字符串。
 */
function _huGetAllResponseHeaders(request) {
    try {
        if (typeof request.GetAllResponseHeaders === "function") {
            return String(request.GetAllResponseHeaders() || "");
        }
    } catch (error) {
    }

    return "";
}

/**
 * 读取响应文本。
 *
 * @private
 * @param {Object} request 请求对象。
 * @returns {string} 响应文本；读取失败时返回空字符串。
 */
function _huTryReadResponseText(request) {
    try {
        return String(request.ResponseText || "");
    } catch (error) {
        return "";
    }
}

/**
 * 读取响应二进制内容。
 *
 * @private
 * @param {Object} request 请求对象。
 * @returns {*} 响应二进制内容；读取失败时返回 `null`。
 */
function _huTryReadResponseBody(request) {
    try {
        return request.ResponseBody;
    } catch (error) {
        return null;
    }
}

/**
 * 判断是否为 JSON 响应。
 *
 * @private
 * @param {string} contentType Content-Type 值。
 * @returns {boolean} 是 JSON 响应时返回 `true`，否则返回 `false`。
 */
function _huIsJsonResponse(contentType) {
    if (!contentType) {
        return false;
    }

    return String(contentType).toLowerCase().indexOf("json") >= 0;
}

/**
 * 解析下载保存路径。
 *
 * @private
 * @param {string} url 下载地址。
 * @param {string} [targetPath] 目标路径。
 * @param {boolean} [autoRename] 是否自动改名。
 * @returns {string} 保存路径；失败时返回空字符串。
 */
function _huResolveDownloadPath(url, targetPath, autoRename) {
    var savePath = targetPath ? String(targetPath) : "";
    var fileName = "";

    if (savePath === "") {
        fileName = _huGetFileNameFromUrl(url);
        if (fileName === "") {
            fileName = "download.bin";
        }
        savePath = _huJoinPath(_huGetTempFolderPath(), fileName);
    }

    if (autoRename === true) {
        savePath = _huEnsureUniqueFilePath(savePath);
    }

    return savePath;
}

/**
 * 保存二进制内容到文件。
 *
 * @private
 * @param {string} path 目标文件路径。
 * @param {*} body 响应二进制内容。
 * @returns {boolean} 保存成功时返回 `true`，否则返回 `false`。
 */
function _huWriteBinary(path, body) {
    if (!_huEnsureParentFolder(path)) {
        return false;
    }

    try {
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 1;
        stream.Mode = 3;
        stream.Open();
        stream.Write(body);
        stream.Position = 0;
        stream.SaveToFile(String(path), 2);
        stream.Close();
        return true;
    } catch (error) {
        return false;
    }
}

/**
 * 保存文本内容到文件。
 *
 * @private
 * @param {string} path 目标文件路径。
 * @param {string} text 文本内容。
 * @param {string} [charset] 编码。
 * @param {boolean} [overwrite] 是否覆盖。
 * @returns {boolean} 保存成功时返回 `true`，否则返回 `false`。
 */
function _huWriteText(path, text, charset, overwrite) {
    if (typeof writeTextFile === "function") {
        return writeTextFile(path, text, overwrite === true, charset);
    }

    if (!_huEnsureParentFolder(path)) {
        return false;
    }

    try {
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 2;
        stream.Mode = 3;
        stream.Charset = charset || "utf-8";
        stream.Open();
        stream.WriteText(String(text === null || text === undefined ? "" : text));
        stream.Position = 0;
        stream.SaveToFile(String(path), overwrite === true ? 2 : 1);
        stream.Close();
        return true;
    } catch (error) {
        return false;
    }
}

/**
 * 保存 JSON 内容到文件。
 *
 * @private
 * @param {string} path 目标文件路径。
 * @param {*} data JSON 数据。
 * @param {boolean} [overwrite] 是否覆盖。
 * @param {number} [indent] 缩进空格数。
 * @param {string} [charset] 编码。
 * @returns {boolean} 保存成功时返回 `true`，否则返回 `false`。
 */
function _huWriteJson(path, data, overwrite, indent, charset) {
    if (typeof writeJsonFile === "function") {
        return writeJsonFile(path, data, overwrite === true, indent, charset);
    }

    if (typeof JSON === "undefined" || !JSON.stringify) {
        return false;
    }

    return _huWriteText(path, JSON.stringify(data, null, indent === undefined ? 4 : indent), charset, overwrite);
}

/**
 * 确保目标父目录存在。
 *
 * @private
 * @param {string} path 文件路径。
 * @returns {boolean} 目录存在或创建成功时返回 `true`，否则返回 `false`。
 */
function _huEnsureParentFolder(path) {
    if (typeof ensureParentFolder === "function") {
        return ensureParentFolder(path);
    }

    var parentPath = _huGetParentFolderPath(path);

    if (parentPath === "") {
        return true;
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (fso.FolderExists(parentPath)) {
            return true;
        }

        fso.CreateFolder(parentPath);
        return fso.FolderExists(parentPath);
    } catch (error) {
        return false;
    }
}

/**
 * 生成不冲突的文件路径。
 *
 * @private
 * @param {string} path 原始路径。
 * @returns {string} 不冲突的路径。
 */
function _huEnsureUniqueFilePath(path) {
    if (typeof ensureUniqueFilePath === "function") {
        return ensureUniqueFilePath(path, "_", 1);
    }

    var targetPath = String(path);
    var index = 1;

    while (_huFileExists(targetPath)) {
        targetPath = _huAppendBaseNameSuffix(path, "_" + index);
        index = index + 1;
    }

    return targetPath;
}

/**
 * 在文件基础名后追加后缀。
 *
 * @private
 * @param {string} path 原始路径。
 * @param {string} suffix 后缀。
 * @returns {string} 新路径。
 */
function _huAppendBaseNameSuffix(path, suffix) {
    if (typeof appendBaseNameSuffix === "function") {
        return appendBaseNameSuffix(path, suffix);
    }

    var parentPath = _huGetParentFolderPath(path);
    var fileName = _huGetPathName(path);
    var index = fileName.lastIndexOf(".");
    var baseName = index > 0 ? fileName.substring(0, index) : fileName;
    var extend = index > 0 ? fileName.substring(index) : "";
    var nextName = baseName + String(suffix || "") + extend;

    return parentPath === "" ? nextName : _huJoinPath(parentPath, nextName);
}

/**
 * 判断文件是否存在。
 *
 * @private
 * @param {string} path 文件路径。
 * @returns {boolean} 存在时返回 `true`，否则返回 `false`。
 */
function _huFileExists(path) {
    if (typeof fileExists === "function") {
        return fileExists(path);
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        return fso.FileExists(String(path));
    } catch (error) {
        return false;
    }
}

/**
 * 获取临时目录路径。
 *
 * @private
 * @returns {string} 临时目录路径。
 */
function _huGetTempFolderPath() {
    if (typeof getTempFolderPath === "function") {
        return getTempFolderPath();
    }

    try {
        var shell = new ActiveXObject("WScript.Shell");
        return String(shell.ExpandEnvironmentStrings("%TEMP%"));
    } catch (error) {
        return ".";
    }
}

/**
 * 从 URL 中提取文件名。
 *
 * @private
 * @param {string} url 下载地址。
 * @returns {string} 文件名。
 */
function _huGetFileNameFromUrl(url) {
    var text = String(url || "");
    var cleanUrl = text.split("?")[0];
    var cleanHashUrl = cleanUrl.split("#")[0];
    var index = cleanHashUrl.lastIndexOf("/");

    if (index >= 0) {
        return cleanHashUrl.substring(index + 1);
    }

    return cleanHashUrl;
}

/**
 * 获取路径中的文件名。
 *
 * @private
 * @param {string} path 文件路径。
 * @returns {string} 文件名。
 */
function _huGetPathName(path) {
    var text = String(path || "").replace(/\//g, "\\");
    var parts = text.split("\\");

    return parts.length > 0 ? parts[parts.length - 1] : text;
}

/**
 * 获取父目录路径。
 *
 * @private
 * @param {string} path 文件路径。
 * @returns {string} 父目录路径。
 */
function _huGetParentFolderPath(path) {
    var text = String(path || "").replace(/\//g, "\\");
    var index = text.lastIndexOf("\\");

    if (index <= 0) {
        return "";
    }

    return text.substring(0, index);
}

/**
 * 拼接路径。
 *
 * @private
 * @param {string} left 左侧路径。
 * @param {string} right 右侧路径。
 * @returns {string} 拼接后的路径。
 */
function _huJoinPath(left, right) {
    if (typeof joinPath === "function") {
        return joinPath(left, right);
    }

    var leftText = String(left || "").replace(/\//g, "\\");
    var rightText = String(right || "").replace(/\//g, "\\");

    if (leftText === "") {
        return rightText;
    }

    if (rightText === "") {
        return leftText;
    }

    leftText = leftText.replace(/[\\\/]+$/g, "");
    rightText = rightText.replace(/^[\\\/]+/g, "");
    return leftText + "\\" + rightText;
}

/**
 * 编码查询参数。
 *
 * @private
 * @param {*} value 原始值。
 * @returns {string} 编码后的字符串。
 */
function _huEncode(value) {
    try {
        return encodeURIComponent(String(value));
    } catch (error) {
        return escape(String(value));
    }
}
