/**
 * 打开 WPS 文件夹选择器，并返回用户选择的文件夹路径。
 * 优先使用 WPS 的 Application.FileDialog，失败后再使用 Windows Shell 对象作为兜底。
 *
 * @returns {string} 用户选择的文件夹路径；用户取消或选择器不可用时返回空字符串。
 */
function openFolderPicker() {
    try {
        var dialog = Application.FileDialog(4); // 4 = msoFileDialogFolderPicker
        dialog.Title = "请选择文件夹";
        dialog.AllowMultiSelect = false;

        if (dialog.Show() === -1 && dialog.SelectedItems.Count > 0) {
            return String(dialog.SelectedItems.Item(1));
        }
    } catch (error) {
        // 兜底路径：优先依赖 WPS API，仅在不可用时使用 Windows Shell。
        try {
            var shell = new ActiveXObject("Shell.Application");
            var folder = shell.BrowseForFolder(0, "请选择文件夹", 0, 0);

            if (!folder || !folder.Self || !folder.Self.Path) {
                return "";
            }

            return String(folder.Self.Path);
        } catch (fallbackError) {
            return "";
        }
    }

    return "";
}

/**
 * 打开 WPS 文件选择器，并返回用户选择的文件路径。
 *
 * @param {string} [filterDescription] 文件类型说明，例如“Excel 文件”。
 * @param {string} [filterPattern] 文件匹配规则，例如“*.xlsx;*.xlsm”；为空时使用“*.*”。
 * @param {boolean} [allowMultiSelect] 是否允许多选；为 true 时返回路径数组。
 * @returns {string|string[]} 单选时返回文件路径字符串，多选时返回文件路径数组；用户取消或失败时返回空字符串或空数组。
 */
function openFilePicker(filterDescription, filterPattern, allowMultiSelect) {
    var result = [];
    var multiSelect = allowMultiSelect === true;

    try {
        var dialog = Application.FileDialog(3); // 3 = msoFileDialogFilePicker
        dialog.Title = "请选择文件";
        dialog.AllowMultiSelect = multiSelect;

        if (filterDescription || filterPattern) {
            try {
                dialog.Filters.Clear();
                dialog.Filters.Add(String(filterDescription || "文件"), String(filterPattern || "*.*"));
            } catch (filterError) {
            }
        }

        if (dialog.Show() === -1 && dialog.SelectedItems.Count > 0) {
            for (var i = 1; i <= dialog.SelectedItems.Count; i++) {
                result.push(String(dialog.SelectedItems.Item(i)));
            }
        }
    } catch (error) {
        return multiSelect ? [] : "";
    }

    if (multiSelect) {
        return result;
    }

    return result.length > 0 ? result[0] : "";
}

/**
 * 判断指定路径是否为已经存在的文件。
 * 优先使用 WPS/JSA 的 GetAttr，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} path 要检查的文件路径。
 * @returns {boolean} 文件存在且不是文件夹时返回 true，否则返回 false。
 */
function fileExists(path) {
    if (!path) {
        return false;
    }

    var targetPath = String(path);

    try {
        var attr = GetAttr(targetPath);
        return !_fsuIsDirectoryAttr(attr);
    } catch (error) {
        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            return fso.FileExists(targetPath);
        } catch (fallbackError) {
            return false;
        }
    }
}

/**
 * 判断指定路径是否为已经存在的文件夹。
 * 优先使用 WPS/JSA 的 GetAttr，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} path 要检查的文件夹路径。
 * @returns {boolean} 文件夹存在时返回 true，否则返回 false。
 */
function folderExists(path) {
    if (!path) {
        return false;
    }

    var targetPath = _fsuTrimRightSlash(String(path));

    try {
        var attr = GetAttr(targetPath);
        return _fsuIsDirectoryAttr(attr);
    } catch (error) {
        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            return fso.FolderExists(targetPath);
        } catch (fallbackError) {
            return false;
        }
    }
}

/**
 * 标准化路径分隔符，把路径中的正斜杠转换为 Windows 反斜杠。
 *
 * @param {string} path 要标准化的路径。
 * @returns {string} 标准化后的路径；空值返回空字符串。
 */
function normalizePath(path) {
    return _fsuNormalizePath(path);
}

/**
 * 将多个路径片段拼接为一个 Windows 路径。
 *
 * @param {...string} pathParts 路径片段列表。
 * @returns {string} 拼接后的路径。
 */
function joinPath() {
    var result = "";

    for (var i = 0; i < arguments.length; i++) {
        var part = _fsuNormalizePath(arguments[i]);

        if (part === "") {
            continue;
        }

        if (result === "") {
            result = _fsuTrimRightSlash(part);
            continue;
        }

        result = _fsuJoinTwoPathParts(result, part);
    }

    return result;
}

/**
 * 获取路径最后一段名称。
 *
 * @param {string} path 文件或文件夹路径。
 * @returns {string} 路径最后一段名称；例如文件路径返回文件名，文件夹路径返回文件夹名。
 */
function getPathName(path) {
    var targetPath = _fsuTrimRightSlash(path);
    var index = targetPath.lastIndexOf("\\");

    if (index < 0) {
        return targetPath;
    }

    return targetPath.substring(index + 1);
}

/**
 * 获取文件或文件夹路径的父目录路径。
 *
 * @param {string} path 文件或文件夹路径。
 * @returns {string} 父目录路径；没有父目录时返回空字符串。
 */
function getParentFolderPath(path) {
    var targetPath = _fsuTrimRightSlash(path);
    var index = targetPath.lastIndexOf("\\");

    if (index < 0) {
        return "";
    }

    if (index === 2 && targetPath.charAt(1) === ":") {
        return targetPath.substring(0, 3);
    }

    return targetPath.substring(0, index);
}

/**
 * 获取不带扩展名的文件名。
 *
 * @param {string} fileName 文件名或完整文件路径。
 * @returns {string} 不带扩展名的文件名。
 */
function getFileBaseName(fileName) {
    var name = getPathName(fileName);
    var index = name.lastIndexOf(".");

    if (index <= 0) {
        return name;
    }

    return name.substring(0, index);
}

/**
 * 获取文件扩展名，不包含点号。
 *
 * @param {string} fileName 文件名或完整文件路径。
 * @returns {string} 文件扩展名；没有扩展名时返回空字符串。
 */
function getFileExtend(fileName) {
    return _fsuGetFileExtend(String(fileName || ""));
}

/**
 * 替换文件路径的扩展名。
 *
 * @param {string} path 原始文件路径。
 * @param {string} newExtend 新扩展名，可包含或不包含前导点号；为空时移除扩展名。
 * @returns {string} 替换扩展名后的路径；原始路径为空时返回空字符串。
 */
function changeFileExtend(path, newExtend) {
    if (!path) {
        return "";
    }

    var targetPath = _fsuNormalizePath(path);
    var parentPath = getParentFolderPath(targetPath);
    var baseName = getFileBaseName(targetPath);
    var extend = String(newExtend || "");

    if (extend.charAt(0) === ".") {
        extend = extend.substring(1);
    }

    var fileName = extend === "" ? baseName : baseName + "." + extend;

    if (parentPath === "") {
        return fileName;
    }

    return joinPath(parentPath, fileName);
}

/**
 * 创建单级文件夹。
 * 优先使用 WPS/JSA 的 MkDir，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} path 要创建的文件夹路径；父目录需要已经存在。
 * @returns {boolean} 文件夹已存在或创建成功时返回 true，否则返回 false。
 */
function createFolder(path) {
    if (!path) {
        return false;
    }

    var targetPath = _fsuTrimRightSlash(path);

    if (folderExists(targetPath)) {
        return true;
    }

    try {
        MkDir(targetPath);
        return folderExists(targetPath);
    } catch (error) {
        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            fso.CreateFolder(targetPath);
            return fso.FolderExists(targetPath);
        } catch (fallbackError) {
            return false;
        }
    }
}

/**
 * 确保文件夹存在，并递归创建缺失的父级文件夹。
 *
 * @param {string} path 要确保存在的文件夹路径。
 * @returns {boolean} 最终文件夹存在时返回 true，否则返回 false。
 */
function ensureFolder(path) {
    if (!path) {
        return false;
    }

    var targetPath = _fsuTrimRightSlash(path);

    if (folderExists(targetPath)) {
        return true;
    }

    if (_fsuIsRootPath(targetPath)) {
        return folderExists(targetPath);
    }

    var parentPath = getParentFolderPath(targetPath);

    if (parentPath !== "" && !folderExists(parentPath) && !ensureFolder(parentPath)) {
        return false;
    }

    return createFolder(targetPath);
}

/**
 * 确保某个文件路径的父目录存在。
 *
 * @param {string} path 文件路径。
 * @returns {boolean} 父目录存在或创建成功时返回 true，否则返回 false。
 */
function ensureParentFolder(path) {
    var parentPath = getParentFolderPath(path);

    if (parentPath === "") {
        return false;
    }

    return ensureFolder(parentPath);
}

/**
 * 获取系统临时目录路径。
 * 优先使用 WPS/JSA 兼容的 Environ，再使用 Windows Shell 和 FileSystemObject 兜底。
 *
 * @returns {string} 系统临时目录路径；获取失败时返回空字符串。
 */
function getTempFolderPath() {
    try {
        var tempFolder = String(Environ("TEMP") || "");

        if (tempFolder !== "") {
            return tempFolder;
        }
    } catch (error) {
    }

    try {
        var tmpFolder = String(Environ("TMP") || "");

        if (tmpFolder !== "") {
            return tmpFolder;
        }
    } catch (tmpError) {
    }

    try {
        var shell = new ActiveXObject("WScript.Shell");
        var shellTempFolder = String(shell.ExpandEnvironmentStrings("%TEMP%") || "");

        if (shellTempFolder !== "" && shellTempFolder !== "%TEMP%") {
            return shellTempFolder;
        }
    } catch (shellError) {
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        return String(fso.GetSpecialFolder(2)); // 2 = TemporaryFolder
    } catch (error) {
        return "";
    }
}

/**
 * 获取文件大小，单位为字节。
 * 优先使用 WPS/JSA 的 FileLen，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} path 文件路径。
 * @returns {number} 文件大小；文件不存在或读取失败时返回 -1。
 */
function getFileSize(path) {
    if (!fileExists(path)) {
        return -1;
    }

    try {
        return Number(FileLen(String(path)));
    } catch (error) {
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        return Number(fso.GetFile(String(path)).Size);
    } catch (fallbackError) {
        return -1;
    }
}

/**
 * 获取文件最后修改时间。
 * 优先使用 WPS/JSA 的 FileDateTime，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} path 文件路径。
 * @returns {Date|string} 文件最后修改时间；文件不存在或读取失败时返回空字符串。
 */
function getFileModifiedTime(path) {
    if (!fileExists(path)) {
        return "";
    }

    try {
        return FileDateTime(String(path));
    } catch (error) {
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        return fso.GetFile(String(path)).DateLastModified;
    } catch (fallbackError) {
        return "";
    }
}

/**
 * 获取文件夹最后修改时间。
 * 文件夹时间主要依赖 Windows FileSystemObject，因为 WPS/JSA 常用内置函数没有直接的目录时间接口。
 *
 * @param {string} path 文件夹路径。
 * @returns {Date|string} 文件夹最后修改时间；文件夹不存在或读取失败时返回空字符串。
 */
function getFolderModifiedTime(path) {
    if (!folderExists(path)) {
        return "";
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        return fso.GetFolder(String(path)).DateLastModified;
    } catch (error) {
        return "";
    }
}

/**
 * 读取文本文件内容。
 * 文本读取主要依赖 Windows ActiveX 文本流对象，因为 WPS/JSA 没有通用的带编码文本读取接口。
 *
 * @param {string} path 文本文件路径。
 * @param {string} [charset] 文本编码，例如 "utf-8"、"unicode"。
 * @returns {string} 文本内容；文件不存在或读取失败时返回空字符串。
 */
function readTextFile(path, charset) {
    if (!fileExists(path)) {
        return "";
    }

    var targetPath = String(path);
    var fileCharset = _fsuNormalizeCharset(charset);

    try {
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 2;
        stream.Mode = 3;
        stream.Charset = fileCharset;
        stream.Open();
        stream.LoadFromFile(targetPath);
        var text = String(stream.ReadText(-1));
        stream.Close();
        return text;
    } catch (error) {
        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            var format = fileCharset === "unicode" ? -1 : 0;
            var textFile = fso.OpenTextFile(targetPath, 1, false, format);
            var fallbackText = String(textFile.ReadAll());
            textFile.Close();
            return fallbackText;
        } catch (fallbackError) {
            return "";
        }
    }
}

/**
 * 写入文本文件内容，并在需要时自动创建目标父目录。
 * 文本写入主要依赖 Windows ActiveX 文本流对象，因为 WPS/JSA 没有通用的带编码文本写入接口。
 *
 * @param {string} path 文本文件路径。
 * @param {*} text 要写入的文本内容。
 * @param {boolean} [overwrite] 目标文件存在时是否覆盖；只有 true 会覆盖。
 * @param {string} [charset] 文本编码，例如 "utf-8"、"unicode"。
 * @returns {boolean} 写入成功时返回 true，否则返回 false。
 */
function writeTextFile(path, text, overwrite, charset) {
    if (!path) {
        return false;
    }

    var targetPath = String(path);
    var fileCharset = _fsuNormalizeCharset(charset);

    if (fileExists(targetPath) && overwrite !== true) {
        return false;
    }

    if (!ensureParentFolder(targetPath)) {
        return false;
    }

    try {
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 2;
        stream.Mode = 3;
        stream.Charset = fileCharset;
        stream.Open();
        stream.WriteText(String(text === null || text === undefined ? "" : text));
        stream.Position = 0;
        stream.SaveToFile(targetPath, overwrite === true ? 2 : 1);
        stream.Close();
        return fileExists(targetPath);
    } catch (error) {
        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            var createNew = overwrite === true ? true : !fso.FileExists(targetPath);
            var textFile = fso.CreateTextFile(targetPath, createNew, fileCharset === "unicode");
            textFile.Write(String(text === null || text === undefined ? "" : text));
            textFile.Close();
            return fso.FileExists(targetPath);
        } catch (fallbackError) {
            return false;
        }
    }
}

/**
 * 读取 JSON 文件内容并解析为对象。
 *
 * @param {string} path JSON 文件路径。
 * @param {string} [charset] 文本编码，例如 "utf-8"。
 * @returns {*} 解析后的对象或数组；读取失败、文件为空或解析失败时返回 null。
 */
function readJsonFile(path, charset) {
    var text = readTextFile(path, charset);

    if (text === "") {
        return null;
    }

    try {
        if (typeof JSON !== "undefined" && JSON.parse) {
            return JSON.parse(text);
        }

        return eval("(" + text + ")");
    } catch (error) {
        return null;
    }
}

/**
 * 把对象或数组写入 JSON 文件，并在需要时自动创建目标父目录。
 *
 * @param {string} path JSON 文件路径。
 * @param {*} data 要写入的对象、数组或基础值。
 * @param {boolean} [overwrite] 目标文件存在时是否覆盖；只有 true 会覆盖。
 * @param {number} [indent] JSON 缩进空格数；默认 4。
 * @param {string} [charset] 文本编码，例如 "utf-8"。
 * @returns {boolean} 写入成功时返回 true，否则返回 false。
 */
function writeJsonFile(path, data, overwrite, indent, charset) {
    var spaceCount = Number(indent);

    if (!isFinite(spaceCount) || spaceCount < 0) {
        spaceCount = 4;
    }

    try {
        if (typeof JSON === "undefined" || !JSON.stringify) {
            return false;
        }

        return writeTextFile(path, JSON.stringify(data, null, spaceCount), overwrite, charset);
    } catch (error) {
        return false;
    }
}

/**
 * 在文件基础名后追加后缀，并保留原扩展名。
 *
 * @param {string} path 原始文件路径。
 * @param {string} suffix 追加到基础名后的后缀文本。
 * @returns {string} 追加后缀后的文件路径；原始路径为空时返回空字符串。
 */
function appendBaseNameSuffix(path, suffix) {
    if (!path) {
        return "";
    }

    var targetPath = _fsuNormalizePath(path);
    var parentPath = getParentFolderPath(targetPath);
    var fileName = getPathName(targetPath);
    var baseName = getFileBaseName(fileName);
    var extend = getFileExtend(fileName);
    var nextName = baseName + String(suffix || "");

    if (extend !== "") {
        nextName = nextName + "." + extend;
    }

    if (parentPath === "") {
        return nextName;
    }

    return joinPath(parentPath, nextName);
}

/**
 * 生成一个当前不存在的文件路径；若原路径不存在，则直接返回原路径。
 *
 * @param {string} path 原始文件路径。
 * @param {string} [separator] 基础名与序号之间的分隔符；默认 "_"。
 * @param {number} [startIndex] 起始序号；默认 1。
 * @returns {string} 不与现有文件或文件夹冲突的文件路径；原始路径为空时返回空字符串。
 */
function ensureUniqueFilePath(path, separator, startIndex) {
    if (!path) {
        return "";
    }

    var targetPath = _fsuNormalizePath(path);
    var suffixSeparator = separator === undefined ? "_" : String(separator);
    var index = Number(startIndex);

    if (!isFinite(index) || index < 1) {
        index = 1;
    }

    if (!fileExists(targetPath) && !folderExists(targetPath)) {
        return targetPath;
    }

    while (fileExists(targetPath) || folderExists(targetPath)) {
        targetPath = appendBaseNameSuffix(path, suffixSeparator + _fsuPadNumber(index, 3));
        index = index + 1;
    }

    return targetPath;
}

/**
 * 复制文件，并在需要时自动创建目标父目录。
 * 优先使用 WPS/JSA 的 FileCopy，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} sourcePath 源文件路径。
 * @param {string} targetPath 目标文件路径。
 * @param {boolean} [overwrite] 目标文件存在时是否覆盖；只有 true 会覆盖。
 * @returns {boolean} 复制成功时返回 true，否则返回 false。
 */
function copyFile(sourcePath, targetPath, overwrite) {
    if (!fileExists(sourcePath) || !targetPath) {
        return false;
    }

    if (fileExists(targetPath)) {
        if (overwrite !== true || !deleteFile(targetPath, true)) {
            return false;
        }
    }

    if (!ensureParentFolder(targetPath)) {
        return false;
    }

    try {
        FileCopy(String(sourcePath), String(targetPath));
        return fileExists(targetPath);
    } catch (error) {
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        fso.CopyFile(String(sourcePath), String(targetPath), overwrite === true);
        return fileExists(targetPath);
    } catch (fallbackError) {
        return false;
    }
}

/**
 * 移动文件，并在需要时自动创建目标父目录。
 * 优先使用 WPS/JSA 的 Name，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} sourcePath 源文件路径。
 * @param {string} targetPath 目标文件路径。
 * @param {boolean} [overwrite] 目标文件存在时是否覆盖；只有 true 会先删除目标文件再移动。
 * @returns {boolean} 移动成功时返回 true，否则返回 false。
 */
function moveFile(sourcePath, targetPath, overwrite) {
    if (!fileExists(sourcePath) || !targetPath) {
        return false;
    }

    if (fileExists(targetPath)) {
        if (overwrite !== true || !deleteFile(targetPath, true)) {
            return false;
        }
    }

    if (!ensureParentFolder(targetPath)) {
        return false;
    }

    try {
        Name(String(sourcePath), String(targetPath));
        return fileExists(targetPath) && !fileExists(sourcePath);
    } catch (error) {
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        fso.MoveFile(String(sourcePath), String(targetPath));
        return fileExists(targetPath) && !fileExists(sourcePath);
    } catch (fallbackError) {
        return false;
    }
}

/**
 * 删除文件。
 * 优先使用 WPS/JSA 的 Kill，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} path 要删除的文件路径。
 * @param {boolean} [force] 是否强制删除只读文件；传 false 时不强制，其它值默认强制。
 * @returns {boolean} 文件不存在或删除成功时返回 true，否则返回 false。
 */
function deleteFile(path, force) {
    if (!path) {
        return false;
    }

    if (!fileExists(path)) {
        return true;
    }

    try {
        Kill(String(path));
        return !fileExists(path);
    } catch (error) {
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        fso.DeleteFile(String(path), force !== false);
        return !fileExists(path);
    } catch (fallbackError) {
        return false;
    }
}

/**
 * 复制文件夹，并在需要时自动创建目标父目录。
 * 文件夹递归复制主要依赖 Windows FileSystemObject，因为 WPS/JSA 常用内置函数没有等价的递归复制接口。
 *
 * @param {string} sourcePath 源文件夹路径。
 * @param {string} targetPath 目标文件夹路径。
 * @param {boolean} [overwrite] 目标文件夹中存在同名内容时是否覆盖；只有 true 会覆盖。
 * @returns {boolean} 复制成功时返回 true，否则返回 false。
 */
function copyFolder(sourcePath, targetPath, overwrite) {
    if (!folderExists(sourcePath) || !targetPath || _fsuIsRootPath(sourcePath)) {
        return false;
    }

    if (!ensureParentFolder(targetPath)) {
        return false;
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        fso.CopyFolder(_fsuTrimRightSlash(sourcePath), _fsuTrimRightSlash(targetPath), overwrite === true);
        return folderExists(targetPath);
    } catch (error) {
        return false;
    }
}

/**
 * 移动文件夹，并在需要时自动创建目标父目录。
 * 优先使用 WPS/JSA 的 Name，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} sourcePath 源文件夹路径。
 * @param {string} targetPath 目标文件夹路径。
 * @param {boolean} [overwrite] 目标文件夹存在时是否覆盖；只有 true 会先删除目标文件夹再移动。
 * @returns {boolean} 移动成功时返回 true，否则返回 false。
 */
function moveFolder(sourcePath, targetPath, overwrite) {
    if (!folderExists(sourcePath) || !targetPath || _fsuIsRootPath(sourcePath)) {
        return false;
    }

    if (folderExists(targetPath)) {
        if (overwrite !== true || !deleteFolder(targetPath, true)) {
            return false;
        }
    }

    if (!ensureParentFolder(targetPath)) {
        return false;
    }

    try {
        Name(_fsuTrimRightSlash(sourcePath), _fsuTrimRightSlash(targetPath));
        return folderExists(targetPath) && !folderExists(sourcePath);
    } catch (error) {
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        fso.MoveFolder(_fsuTrimRightSlash(sourcePath), _fsuTrimRightSlash(targetPath));
        return folderExists(targetPath) && !folderExists(sourcePath);
    } catch (fallbackError) {
        return false;
    }
}

/**
 * 删除文件夹，根目录永远不会被删除。
 * 优先使用 WPS/JSA 的 RmDir，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} path 要删除的文件夹路径。
 * @param {boolean} [force] 是否强制删除只读内容；传 false 时不强制，其它值默认强制。
 * @returns {boolean} 文件夹不存在或删除成功时返回 true，否则返回 false。
 */
function deleteFolder(path, force) {
    if (!path) {
        return false;
    }

    var targetPath = _fsuTrimRightSlash(path);

    if (_fsuIsRootPath(targetPath)) {
        return false;
    }

    if (!folderExists(targetPath)) {
        return true;
    }

    try {
        RmDir(targetPath);
        return !folderExists(targetPath);
    } catch (error) {
    }

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        fso.DeleteFolder(targetPath, force !== false);
        return !folderExists(targetPath);
    } catch (fallbackError) {
        return false;
    }
}

/**
 * 获取指定目录下的直接子文件，并按文件名关键字、忽略关键字和扩展名过滤。
 * 优先使用 WPS/JSA 的 Dir 和 GetAttr，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} path 要扫描的文件夹路径。
 * @param {string|string[]} [filterName] 文件名必须包含的关键字；为空时不过滤。
 * @param {string|string[]} [filterIgnore] 文件名中需要排除的关键字；为空时不排除。
 * @param {string|string[]} [filterExtend] 扩展名过滤，例如 "xlsx" 或 ["xlsx", "xlsm"]；为空时不过滤扩展名。
 * @returns {Object.<string, {fileName: string, path: string, extend: string}>} 文件信息对象；键为文件名。
 *
 * 返回格式：
 * {
 *   "demo.xlsx": { fileName: "demo.xlsx", path: "D:\\test\\demo.xlsx", extend: "xlsx" }
 * }
 *
 * @example
 * var files = getFilesByPath("D:\\test", ["公司"], ["~$"], ["xlsx", "xls", "xlsm"]);
 */
function getFilesByPath(path, filterName, filterIgnore, filterExtend) {
    var result = {};

    if (!folderExists(path)) {
        return result;
    }

    var basePath = _fsuTrimRightSlash(String(path));
    var nameFilters = _fsuToStringArray(filterName);
    var ignoreFilters = _fsuToLowerStringArray(filterIgnore);
    var extendFilters = _fsuToLowerStringArray(filterExtend);

    try {
        var searchPattern = basePath + "\\*";
        var fileName = Dir(searchPattern);

        while (fileName !== "") {
            var filePath = basePath + "\\" + fileName;
            var extend = _fsuGetFileExtend(fileName);
            var isIgnored = _fsuIgnoreMatched(fileName, ignoreFilters);

            if (!isIgnored && _fsuIsFilePath(filePath)) {
                if (_fsuNameMatched(fileName, nameFilters) && _fsuExtendMatched(extend, extendFilters)) {
                    result[fileName] = {
                        fileName: fileName,
                        path: filePath,
                        extend: extend
                    };
                }
            }

            fileName = Dir();
        }
    } catch (error) {
        return _fsuGetFilesByPathFallback(basePath, nameFilters, ignoreFilters, extendFilters);
    }

    return result;
}

/**
 * 列出指定目录下的直接子文件，是 getFilesByPath 的语义化别名。
 *
 * @param {string} path 要扫描的文件夹路径。
 * @param {string|string[]} [filterName] 文件名必须包含的关键字；为空时不过滤。
 * @param {string|string[]} [filterIgnore] 文件名中需要排除的关键字；为空时不排除。
 * @param {string|string[]} [filterExtend] 扩展名过滤；为空时不过滤扩展名。
 * @returns {Object.<string, {fileName: string, path: string, extend: string}>} 文件信息对象；键为文件名。
 */
function listFilesByPath(path, filterName, filterIgnore, filterExtend) {
    return getFilesByPath(path, filterName, filterIgnore, filterExtend);
}

/**
 * 获取指定目录下的直接子文件夹，并按文件夹名称关键字过滤。
 * 优先使用 WPS/JSA 的 Dir 和 GetAttr，失败后再使用 Windows FileSystemObject。
 *
 * @param {string} path 要扫描的文件夹路径。
 * @param {string|string[]} [filterName] 文件夹名必须包含的关键字；为空时不过滤。
 * @param {string|string[]} [filterIgnore] 文件夹名中需要排除的关键字；为空时不排除。
 * @returns {Object.<string, {folderName: string, path: string}>} 文件夹信息对象；键为文件夹名。
 */
function getFoldersByPath(path, filterName, filterIgnore) {
    var result = {};

    if (!folderExists(path)) {
        return result;
    }

    var basePath = _fsuTrimRightSlash(String(path));
    var nameFilters = _fsuToStringArray(filterName);
    var ignoreFilters = _fsuToLowerStringArray(filterIgnore);

    try {
        var searchPattern = basePath + "\\*";
        var folderName = Dir(searchPattern);

        while (folderName !== "") {
            var folderPath = basePath + "\\" + folderName;

            if (!_fsuIgnoreMatched(folderName, ignoreFilters) && _fsuIsFolderPath(folderPath) && _fsuNameMatched(folderName, nameFilters)) {
                result[folderName] = {
                    folderName: folderName,
                    path: folderPath
                };
            }

            folderName = Dir();
        }
    } catch (error) {
        return _fsuGetFoldersByPathFallback(basePath, nameFilters, ignoreFilters);
    }

    return result;
}

/**
 * 列出指定目录下的直接子文件夹，是 getFoldersByPath 的语义化别名。
 *
 * @param {string} path 要扫描的文件夹路径。
 * @param {string|string[]} [filterName] 文件夹名必须包含的关键字；为空时不过滤。
 * @param {string|string[]} [filterIgnore] 文件夹名中需要排除的关键字；为空时不排除。
 * @returns {Object.<string, {folderName: string, path: string}>} 文件夹信息对象；键为文件夹名。
 */
function listFoldersByPath(path, filterName, filterIgnore) {
    return getFoldersByPath(path, filterName, filterIgnore);
}

/**
 * 判断文件夹是否为空；同时没有直接子文件和直接子文件夹时视为空目录。
 *
 * @param {string} path 文件夹路径。
 * @returns {boolean} 文件夹存在且没有直接子文件和子文件夹时返回 true，否则返回 false。
 */
function isEmptyFolder(path) {
    if (!folderExists(path)) {
        return false;
    }

    var files = getFilesByPath(path);
    var folders = getFoldersByPath(path);

    return _fsuGetObjectKeys(files).length === 0 && _fsuGetObjectKeys(folders).length === 0;
}

/**
 * 递归列出目录及其子目录中的文件。
 * 递归枚举会先尝试 WPS/JSA 的 Dir + GetAttr；若失败则退回 Windows FileSystemObject。
 *
 * @param {string} path 要扫描的文件夹路径。
 * @param {string|string[]} [filterName] 文件名必须包含的关键字；为空时不过滤。
 * @param {string|string[]} [filterIgnore] 文件名中需要排除的关键字；为空时不排除。
 * @param {string|string[]} [filterExtend] 扩展名过滤；为空时不过滤扩展名。
 * @returns {Array.<{fileName: string, path: string, extend: string}>} 文件信息数组；失败时返回空数组。
 *
 * 返回格式：
 * [
 *   { fileName: "demo.xlsx", path: "D:\\test\\2026\\demo.xlsx", extend: "xlsx" }
 * ]
 */
function walkFilesRecursive(path, filterName, filterIgnore, filterExtend) {
    var result = [];

    if (!folderExists(path)) {
        return result;
    }

    var basePath = _fsuTrimRightSlash(String(path));
    var nameFilters = _fsuToStringArray(filterName);
    var ignoreFilters = _fsuToLowerStringArray(filterIgnore);
    var extendFilters = _fsuToLowerStringArray(filterExtend);

    try {
        _fsuWalkFilesRecursiveByDir(basePath, nameFilters, ignoreFilters, extendFilters, result);
        return result;
    } catch (error) {
        return _fsuWalkFilesRecursiveByFso(basePath, nameFilters, ignoreFilters, extendFilters);
    }
}

/**
 * 递归查找目录及其子目录中的文件，是 walkFilesRecursive 的语义化别名。
 *
 * @param {string} path 要扫描的文件夹路径。
 * @param {string|string[]} [filterName] 文件名必须包含的关键字；为空时不过滤。
 * @param {string|string[]} [filterIgnore] 文件名中需要排除的关键字；为空时不排除。
 * @param {string|string[]} [filterExtend] 扩展名过滤；为空时不过滤扩展名。
 * @returns {Array.<{fileName: string, path: string, extend: string}>} 文件信息数组；失败时返回空数组。
 */
function findFilesRecursive(path, filterName, filterIgnore, filterExtend) {
    return walkFilesRecursive(path, filterName, filterIgnore, filterExtend);
}

/**
 * 递归列出目录及其子目录中的文件夹。
 *
 * @param {string} path 要扫描的文件夹路径。
 * @param {string|string[]} [filterName] 文件夹名必须包含的关键字；为空时不过滤。
 * @param {string|string[]} [filterIgnore] 文件夹名中需要排除的关键字；为空时不排除。
 * @returns {Array.<{folderName: string, path: string}>} 文件夹信息数组；失败时返回空数组。
 *
 * 返回格式：
 * [
 *   { folderName: "2026", path: "D:\\test\\2026" }
 * ]
 */
function walkFoldersRecursive(path, filterName, filterIgnore) {
    var result = [];

    if (!folderExists(path)) {
        return result;
    }

    var basePath = _fsuTrimRightSlash(String(path));
    var nameFilters = _fsuToStringArray(filterName);
    var ignoreFilters = _fsuToLowerStringArray(filterIgnore);

    try {
        _fsuWalkFoldersRecursiveByDir(basePath, nameFilters, ignoreFilters, result);
        return result;
    } catch (error) {
        return _fsuWalkFoldersRecursiveByFso(basePath, nameFilters, ignoreFilters);
    }
}

/**
 * 查找最新修改的文件。
 *
 * @param {string} path 要扫描的文件夹路径。
 * @param {string|string[]} [filterName] 文件名必须包含的关键字；为空时不过滤。
 * @param {string|string[]} [filterIgnore] 文件名中需要排除的关键字；为空时不排除。
 * @param {string|string[]} [filterExtend] 扩展名过滤；为空时不过滤扩展名。
 * @param {boolean} [includeSubFolders] 是否递归扫描子目录；为 true 时递归扫描。
 * @returns {{fileName: string, path: string, extend: string, modifiedTime: Date|string}|null} 最新文件对象；未找到时返回 null。
 *
 * 返回格式：
 * {
 *   fileName: "demo.xlsx",
 *   path: "D:\\test\\demo.xlsx",
 *   extend: "xlsx",
 *   modifiedTime: "2026-04-09 10:30:00"
 * }
 */
function findNewestFile(path, filterName, filterIgnore, filterExtend, includeSubFolders) {
    var files = includeSubFolders === true
        ? walkFilesRecursive(path, filterName, filterIgnore, filterExtend)
        : _fsuObjectToArray(getFilesByPath(path, filterName, filterIgnore, filterExtend));
    var newest = null;
    var newestTimestamp = -1;

    for (var i = 0; i < files.length; i++) {
        var modifiedTime = getFileModifiedTime(files[i].path);
        var timestamp = _fsuDateToTimestamp(modifiedTime);

        if (timestamp > newestTimestamp) {
            newestTimestamp = timestamp;
            newest = {
                fileName: files[i].fileName,
                path: files[i].path,
                extend: files[i].extend,
                modifiedTime: modifiedTime
            };
        }
    }

    return newest;
}

/**
 * 按文件名关键字模式查找最新修改的文件。
 *
 * @param {string} path 要扫描的文件夹路径。
 * @param {string|string[]} pattern 文件名必须包含的关键字模式。
 * @param {boolean} [includeSubFolders] 是否递归扫描子目录；为 true 时递归扫描。
 * @param {string|string[]} [filterExtend] 扩展名过滤；为空时不过滤扩展名。
 * @returns {{fileName: string, path: string, extend: string, modifiedTime: Date|string}|null} 最新文件对象；未找到时返回 null。
 */
function findNewestFileByPattern(path, pattern, includeSubFolders, filterExtend) {
    return findNewestFile(path, pattern, null, filterExtend, includeSubFolders);
}

/**
 * 使用 Windows FileSystemObject 枚举目录下的直接子文件，作为 Dir 枚举失败时的兜底实现。
 *
 * @private
 * @param {string} basePath 要扫描的文件夹路径。
 * @param {string[]} nameFilters 文件名必须包含的关键字列表。
 * @param {string[]} ignoreFilters 文件名中需要排除的小写关键字列表。
 * @param {string[]} extendFilters 允许的小写扩展名列表。
 * @returns {Object.<string, {fileName: string, path: string, extend: string}>} 文件信息对象；失败时返回空对象。
 */
function _fsuGetFilesByPathFallback(basePath, nameFilters, ignoreFilters, extendFilters) {
    var result = {};

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        var folder = fso.GetFolder(basePath);
        var files = new Enumerator(folder.Files);

        for (; !files.atEnd(); files.moveNext()) {
            var file = files.item();
            var fileName = String(file.Name);
            var filePath = String(file.Path);
            var extend = _fsuGetFileExtend(fileName);

            if (!_fsuIgnoreMatched(fileName, ignoreFilters) && _fsuNameMatched(fileName, nameFilters) && _fsuExtendMatched(extend, extendFilters)) {
                result[fileName] = {
                    fileName: fileName,
                    path: filePath,
                    extend: extend
                };
            }
        }
    } catch (error) {
        return {};
    }

    return result;
}

/**
 * 使用 Windows FileSystemObject 枚举目录下的直接子文件夹，作为 Dir 枚举失败时的兜底实现。
 *
 * @private
 * @param {string} basePath 要扫描的文件夹路径。
 * @param {string[]} nameFilters 文件夹名必须包含的关键字列表。
 * @param {string[]} ignoreFilters 文件夹名中需要排除的小写关键字列表。
 * @returns {Object.<string, {folderName: string, path: string}>} 文件夹信息对象；失败时返回空对象。
 */
function _fsuGetFoldersByPathFallback(basePath, nameFilters, ignoreFilters) {
    var result = {};

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        var folder = fso.GetFolder(basePath);
        var folders = new Enumerator(folder.SubFolders);

        for (; !folders.atEnd(); folders.moveNext()) {
            var childFolder = folders.item();
            var folderName = String(childFolder.Name);

            if (!_fsuIgnoreMatched(folderName, ignoreFilters) && _fsuNameMatched(folderName, nameFilters)) {
                result[folderName] = {
                    folderName: folderName,
                    path: String(childFolder.Path)
                };
            }
        }
    } catch (error) {
        return {};
    }

    return result;
}

/**
 * 使用 Dir + GetAttr 递归枚举文件。
 *
 * @private
 * @param {string} basePath 当前扫描目录。
 * @param {string[]} nameFilters 文件名过滤列表。
 * @param {string[]} ignoreFilters 忽略过滤列表。
 * @param {string[]} extendFilters 扩展名过滤列表。
 * @param {Array.<{fileName: string, path: string, extend: string}>} result 结果数组。
 * @returns {void}
 */
function _fsuWalkFilesRecursiveByDir(basePath, nameFilters, ignoreFilters, extendFilters, result) {
    var entries = _fsuListEntriesByDir(basePath);
    var i = 0;

    for (i = 0; i < entries.files.length; i++) {
        if (!_fsuIgnoreMatched(entries.files[i].fileName, ignoreFilters) && _fsuNameMatched(entries.files[i].fileName, nameFilters) && _fsuExtendMatched(entries.files[i].extend, extendFilters)) {
            result.push(entries.files[i]);
        }
    }

    for (i = 0; i < entries.folders.length; i++) {
        _fsuWalkFilesRecursiveByDir(entries.folders[i].path, nameFilters, ignoreFilters, extendFilters, result);
    }
}

/**
 * 使用 Dir + GetAttr 递归枚举文件夹。
 *
 * @private
 * @param {string} basePath 当前扫描目录。
 * @param {string[]} nameFilters 名称过滤列表。
 * @param {string[]} ignoreFilters 忽略过滤列表。
 * @param {Array.<{folderName: string, path: string}>} result 结果数组。
 * @returns {void}
 */
function _fsuWalkFoldersRecursiveByDir(basePath, nameFilters, ignoreFilters, result) {
    var entries = _fsuListEntriesByDir(basePath);
    var i = 0;

    for (i = 0; i < entries.folders.length; i++) {
        if (!_fsuIgnoreMatched(entries.folders[i].folderName, ignoreFilters) && _fsuNameMatched(entries.folders[i].folderName, nameFilters)) {
            result.push(entries.folders[i]);
        }

        _fsuWalkFoldersRecursiveByDir(entries.folders[i].path, nameFilters, ignoreFilters, result);
    }
}

/**
 * 使用 Windows FileSystemObject 递归枚举文件。
 *
 * @private
 * @param {string} basePath 当前扫描目录。
 * @param {string[]} nameFilters 文件名过滤列表。
 * @param {string[]} ignoreFilters 忽略过滤列表。
 * @param {string[]} extendFilters 扩展名过滤列表。
 * @returns {Array.<{fileName: string, path: string, extend: string}>} 文件结果数组。
 */
function _fsuWalkFilesRecursiveByFso(basePath, nameFilters, ignoreFilters, extendFilters) {
    var result = [];

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        _fsuWalkFilesRecursiveByFsoFolder(fso.GetFolder(basePath), nameFilters, ignoreFilters, extendFilters, result);
    } catch (error) {
        return [];
    }

    return result;
}

/**
 * 使用 Windows FileSystemObject 递归枚举文件夹。
 *
 * @private
 * @param {string} basePath 当前扫描目录。
 * @param {string[]} nameFilters 名称过滤列表。
 * @param {string[]} ignoreFilters 忽略过滤列表。
 * @returns {Array.<{folderName: string, path: string}>} 文件夹结果数组。
 */
function _fsuWalkFoldersRecursiveByFso(basePath, nameFilters, ignoreFilters) {
    var result = [];

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        _fsuWalkFoldersRecursiveByFsoFolder(fso.GetFolder(basePath), nameFilters, ignoreFilters, result);
    } catch (error) {
        return [];
    }

    return result;
}

/**
 * 使用 Windows FileSystemObject 递归枚举文件。
 *
 * @private
 * @param {*} folder FSO 文件夹对象。
 * @param {string[]} nameFilters 文件名过滤列表。
 * @param {string[]} ignoreFilters 忽略过滤列表。
 * @param {string[]} extendFilters 扩展名过滤列表。
 * @param {Array.<{fileName: string, path: string, extend: string}>} result 结果数组。
 * @returns {void}
 */
function _fsuWalkFilesRecursiveByFsoFolder(folder, nameFilters, ignoreFilters, extendFilters, result) {
    var files = new Enumerator(folder.Files);
    var folders = new Enumerator(folder.SubFolders);
    var i = 0;

    for (; !files.atEnd(); files.moveNext()) {
        var file = files.item();
        var fileName = String(file.Name);
        var extend = _fsuGetFileExtend(fileName);

        if (!_fsuIgnoreMatched(fileName, ignoreFilters) && _fsuNameMatched(fileName, nameFilters) && _fsuExtendMatched(extend, extendFilters)) {
            result.push({
                fileName: fileName,
                path: String(file.Path),
                extend: extend
            });
        }
    }

    for (; !folders.atEnd(); folders.moveNext()) {
        _fsuWalkFilesRecursiveByFsoFolder(folders.item(), nameFilters, ignoreFilters, extendFilters, result);
    }
}

/**
 * 使用 Windows FileSystemObject 递归枚举文件夹。
 *
 * @private
 * @param {*} folder FSO 文件夹对象。
 * @param {string[]} nameFilters 名称过滤列表。
 * @param {string[]} ignoreFilters 忽略过滤列表。
 * @param {Array.<{folderName: string, path: string}>} result 结果数组。
 * @returns {void}
 */
function _fsuWalkFoldersRecursiveByFsoFolder(folder, nameFilters, ignoreFilters, result) {
    var folders = new Enumerator(folder.SubFolders);

    for (; !folders.atEnd(); folders.moveNext()) {
        var childFolder = folders.item();
        var folderName = String(childFolder.Name);

        if (!_fsuIgnoreMatched(folderName, ignoreFilters) && _fsuNameMatched(folderName, nameFilters)) {
            result.push({
                folderName: folderName,
                path: String(childFolder.Path)
            });
        }

        _fsuWalkFoldersRecursiveByFsoFolder(childFolder, nameFilters, ignoreFilters, result);
    }
}

/**
 * 使用 Dir + GetAttr 列出一个目录当前层的文件与文件夹。
 *
 * @private
 * @param {string} basePath 当前扫描目录。
 * @returns {{files: Array.<{fileName: string, path: string, extend: string}>, folders: Array.<{folderName: string, path: string}>}} 当前层文件与文件夹结果。
 */
function _fsuListEntriesByDir(basePath) {
    var result = {
        files: [],
        folders: []
    };
    var name = Dir(basePath + "\\*");

    while (name !== "") {
        var fullPath = basePath + "\\" + name;

        if (_fsuIsFolderPath(fullPath)) {
            if (name !== "." && name !== "..") {
                result.folders.push({
                    folderName: name,
                    path: fullPath
                });
            }
        } else if (_fsuIsFilePath(fullPath)) {
            result.files.push({
                fileName: name,
                path: fullPath,
                extend: _fsuGetFileExtend(name)
            });
        }

        name = Dir();
    }

    return result;
}

/**
 * 标准化文本编码名称。
 *
 * @private
 * @param {string} charset 输入编码名称。
 * @returns {string} 标准化后的编码名称。
 */
function _fsuNormalizeCharset(charset) {
    var text = String(charset || "").toLowerCase();

    if (text === "" || text === "utf8") {
        return "utf-8";
    }

    if (text === "utf-16" || text === "utf16") {
        return "unicode";
    }

    return text;
}

/**
 * 把数字补齐为固定长度字符串。
 *
 * @private
 * @param {number} value 原始数字。
 * @param {number} size 目标长度。
 * @returns {string} 补齐后的数字字符串。
 */
function _fsuPadNumber(value, size) {
    var text = String(Math.floor(Math.abs(Number(value) || 0)));

    while (text.length < size) {
        text = "0" + text;
    }

    return text;
}

/**
 * 获取对象自有键列表。
 *
 * @private
 * @param {Object} obj 要读取键的对象。
 * @returns {string[]} 自有键数组。
 */
function _fsuGetObjectKeys(obj) {
    var result = [];
    var key = "";

    if (!obj) {
        return result;
    }

    for (key in obj) {
        if (Object.prototype.hasOwnProperty.call(obj, key)) {
            result.push(key);
        }
    }

    return result;
}

/**
 * 把键值对象转换为值数组。
 *
 * @private
 * @param {Object} obj 要转换的对象。
 * @returns {Array} 值数组。
 */
function _fsuObjectToArray(obj) {
    var keys = _fsuGetObjectKeys(obj);
    var result = [];
    var i = 0;

    for (i = 0; i < keys.length; i++) {
        result.push(obj[keys[i]]);
    }

    return result;
}

/**
 * 把日期值转换为时间戳。
 *
 * @private
 * @param {*} value 日期对象、日期字符串或可解析日期值。
 * @returns {number} 时间戳；无法解析时返回 -1。
 */
function _fsuDateToTimestamp(value) {
    if (!value) {
        return -1;
    }

    try {
        if (Object.prototype.toString.call(value) === "[object Date]") {
            return value.getTime();
        }

        var parsed = new Date(value);
        return isNaN(parsed.getTime()) ? -1 : parsed.getTime();
    } catch (error) {
        return -1;
    }
}

/**
 * 标准化路径分隔符，把正斜杠转换为反斜杠。
 *
 * @private
 * @param {string} path 要标准化的路径。
 * @returns {string} 标准化后的路径；空值返回空字符串。
 */
function _fsuNormalizePath(path) {
    return String(path || "").replace(/\//g, "\\");
}

/**
 * 去掉路径右侧多余的斜杠，但保留根路径斜杠。
 *
 * @private
 * @param {string} path 要处理的路径。
 * @returns {string} 去掉右侧多余斜杠后的路径。
 */
function _fsuTrimRightSlash(path) {
    var result = _fsuNormalizePath(path);

    while (result.length > 0 && (result.charAt(result.length - 1) === "\\" || result.charAt(result.length - 1) === "/") && !_fsuIsRootPath(result)) {
        result = result.substring(0, result.length - 1);
    }

    return result;
}

/**
 * 去掉路径左侧斜杠，用于拼接路径片段。
 *
 * @private
 * @param {string} path 要处理的路径片段。
 * @returns {string} 去掉左侧斜杠后的路径片段。
 */
function _fsuTrimLeftSlash(path) {
    var result = _fsuNormalizePath(path);

    while (result.length > 0 && (result.charAt(0) === "\\" || result.charAt(0) === "/")) {
        result = result.substring(1);
    }

    return result;
}

/**
 * 拼接两个路径片段。
 *
 * @private
 * @param {string} left 左侧路径片段。
 * @param {string} right 右侧路径片段。
 * @returns {string} 拼接后的路径。
 */
function _fsuJoinTwoPathParts(left, right) {
    var leftPart = _fsuTrimRightSlash(left);
    var rightPart = _fsuTrimLeftSlash(right);

    if (leftPart === "") {
        return rightPart;
    }

    if (rightPart === "") {
        return leftPart;
    }

    if (leftPart.charAt(leftPart.length - 1) === "\\") {
        return leftPart + rightPart;
    }

    return leftPart + "\\" + rightPart;
}

/**
 * 判断路径是否为根路径。
 *
 * @private
 * @param {string} path 要判断的路径。
 * @returns {boolean} 路径为盘符根目录、反斜杠根目录或 UNC 根路径时返回 true，否则返回 false。
 */
function _fsuIsRootPath(path) {
    var targetPath = _fsuNormalizePath(path);

    if (targetPath === "\\") {
        return true;
    }

    if (/^[A-Za-z]:\\?$/.test(targetPath)) {
        return true;
    }

    return /^\\\\[^\\]+\\[^\\]+\\?$/.test(targetPath);
}

/**
 * 判断 GetAttr 返回的属性值是否包含目录标记。
 *
 * @private
 * @param {number} attr GetAttr 返回的属性值。
 * @returns {boolean} 包含目录标记时返回 true，否则返回 false。
 */
function _fsuIsDirectoryAttr(attr) {
    return (Number(attr) & 16) === 16; // vbDirectory
}

/**
 * 使用 WPS/JSA 的 GetAttr 判断路径是否为文件。
 *
 * @private
 * @param {string} path 要判断的路径。
 * @returns {boolean} 路径存在且不是文件夹时返回 true，否则返回 false。
 */
function _fsuIsFilePath(path) {
    try {
        var attr = GetAttr(path);
        return !_fsuIsDirectoryAttr(attr);
    } catch (error) {
        return false;
    }
}

/**
 * 使用 WPS/JSA 的 GetAttr 判断路径是否为文件夹。
 *
 * @private
 * @param {string} path 要判断的路径。
 * @returns {boolean} 路径存在且是文件夹时返回 true，否则返回 false。
 */
function _fsuIsFolderPath(path) {
    try {
        var attr = GetAttr(path);
        return _fsuIsDirectoryAttr(attr);
    } catch (error) {
        return false;
    }
}

/**
 * 把空值、字符串或数组统一转换为字符串数组。
 *
 * @private
 * @param {*} value 要转换的值。
 * @returns {string[]} 字符串数组；空值返回空数组。
 */
function _fsuToStringArray(value) {
    if (!value) {
        return [];
    }

    if (Object.prototype.toString.call(value) !== "[object Array]") {
        value = [value];
    }

    var result = [];
    for (var i = 0; i < value.length; i++) {
        if (value[i] !== null && value[i] !== undefined && String(value[i]) !== "") {
            result.push(String(value[i]));
        }
    }

    return result;
}

/**
 * 把过滤条件统一转换为小写字符串数组，并移除扩展名前导点号。
 *
 * @private
 * @param {*} value 要转换的过滤条件。
 * @returns {string[]} 小写字符串数组。
 */
function _fsuToLowerStringArray(value) {
    var items = _fsuToStringArray(value);
    var result = [];

    for (var i = 0; i < items.length; i++) {
        var item = items[i].toLowerCase();
        if (item.charAt(0) === ".") {
            item = item.substring(1);
        }
        result.push(item);
    }

    return result;
}

/**
 * 获取文件扩展名，不包含点号。
 *
 * @private
 * @param {string} fileName 文件名或路径。
 * @returns {string} 文件扩展名；没有扩展名时返回空字符串。
 */
function _fsuGetFileExtend(fileName) {
    var index = fileName.lastIndexOf(".");

    if (index < 0 || index === fileName.length - 1) {
        return "";
    }

    return fileName.substring(index + 1);
}

/**
 * 判断文件名或文件夹名是否命中名称过滤条件。
 *
 * @private
 * @param {string} fileName 文件名或文件夹名。
 * @param {string[]} nameFilters 名称关键字列表。
 * @returns {boolean} 未设置过滤条件或命中任一关键字时返回 true，否则返回 false。
 */
function _fsuNameMatched(fileName, nameFilters) {
    if (!nameFilters || nameFilters.length === 0) {
        return true;
    }

    for (var i = 0; i < nameFilters.length; i++) {
        if (fileName.indexOf(nameFilters[i]) >= 0) {
            return true;
        }
    }

    return false;
}

/**
 * 判断文件名或文件夹名是否命中忽略过滤条件。
 *
 * @private
 * @param {string} fileName 文件名或文件夹名。
 * @param {string[]} ignoreFilters 小写忽略关键字列表。
 * @returns {boolean} 命中任一忽略关键字时返回 true，否则返回 false。
 */
function _fsuIgnoreMatched(fileName, ignoreFilters) {
    if (!ignoreFilters || ignoreFilters.length === 0) {
        return false;
    }

    var lowerName = String(fileName).toLowerCase();
    for (var i = 0; i < ignoreFilters.length; i++) {
        if (lowerName.indexOf(ignoreFilters[i]) >= 0) {
            return true;
        }
    }

    return false;
}

/**
 * 判断扩展名是否命中扩展名过滤条件。
 *
 * @private
 * @param {string} extend 文件扩展名。
 * @param {string[]} extendFilters 小写扩展名过滤列表。
 * @returns {boolean} 未设置过滤条件或扩展名命中时返回 true，否则返回 false。
 */
function _fsuExtendMatched(extend, extendFilters) {
    if (!extendFilters || extendFilters.length === 0) {
        return true;
    }

    var lowerExtend = String(extend).toLowerCase();

    for (var i = 0; i < extendFilters.length; i++) {
        if (lowerExtend === extendFilters[i]) {
            return true;
        }
    }

    return false;
}
