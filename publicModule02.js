/**
 * Open the Windows folder picker and return the selected folder path.
 * Returns an empty string when the user cancels or the picker fails.
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
        // Fallback only: avoid depending on this path when WPS API is available.
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
 * Return true when the target file exists.
 */
function fileExists(path) {
    if (!path) {
        return false;
    }

    var targetPath = String(path);

    try {
        var attr = GetAttr(targetPath);
        return !_pm02IsDirectoryAttr(attr);
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
 * Return true when the target folder exists.
 */
function folderExists(path) {
    if (!path) {
        return false;
    }

    var targetPath = _pm02TrimRightSlash(String(path));

    try {
        var attr = GetAttr(targetPath);
        return _pm02IsDirectoryAttr(attr);
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
 * Return all direct files under a folder, filtered by filename keywords
 * and file extensions.
 *
 * Example:
 * getFilesByPath("D:\\test", ["公司"],["~$"], ["xlsx", "xls", "xlsm"])
 *
 * Return format:
 * {
 *   "demo.xlsx": { fileName: "demo.xlsx", path: "D:\\test\\demo.xlsx", extend: "xlsx" }
 * }
 */
function getFilesByPath(path, filterName, filterIgnore, filterExtend) {
    var result = {};

    if (!folderExists(path)) {
        return result;
    }

    var basePath = _pm02TrimRightSlash(String(path));
    var nameFilters = _pm02ToStringArray(filterName);
    var ignoreFilters = _pm02ToLowerStringArray(filterIgnore);
    var extendFilters = _pm02ToLowerStringArray(filterExtend);

    try {
        var searchPattern = basePath + "\\*";
        var fileName = Dir(searchPattern);

        while (fileName !== "") {
            var filePath = basePath + "\\" + fileName;
            var extend = _pm02GetFileExtend(fileName);
            var isIgnored = _pm02IgnoreMatched(fileName, ignoreFilters);

            if (!isIgnored && _pm02IsFilePath(filePath)) {
                if (_pm02NameMatched(fileName, nameFilters) && _pm02ExtendMatched(extend, extendFilters)) {
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
        return _pm02GetFilesByPathFallback(basePath, nameFilters, ignoreFilters, extendFilters);
    }

    return result;
}

function _pm02GetFilesByPathFallback(basePath, nameFilters, ignoreFilters, extendFilters) {
    var result = {};

    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        var folder = fso.GetFolder(basePath);
        var files = new Enumerator(folder.Files);

        for (; !files.atEnd(); files.moveNext()) {
            var file = files.item();
            var fileName = String(file.Name);
            var filePath = String(file.Path);
            var extend = _pm02GetFileExtend(fileName);

            if (!_pm02IgnoreMatched(fileName, ignoreFilters) && _pm02NameMatched(fileName, nameFilters) && _pm02ExtendMatched(extend, extendFilters)) {
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

function _pm02TrimRightSlash(path) {
    var result = String(path || "");

    while (result.length > 0 && (result.charAt(result.length - 1) === "\\" || result.charAt(result.length - 1) === "/")) {
        result = result.substring(0, result.length - 1);
    }

    return result;
}

function _pm02IsDirectoryAttr(attr) {
    return (Number(attr) & 16) === 16; // vbDirectory
}

function _pm02IsFilePath(path) {
    try {
        var attr = GetAttr(path);
        return !_pm02IsDirectoryAttr(attr);
    } catch (error) {
        return false;
    }
}

function _pm02ToStringArray(value) {
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

function _pm02ToLowerStringArray(value) {
    var items = _pm02ToStringArray(value);
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

function _pm02GetFileExtend(fileName) {
    var index = fileName.lastIndexOf(".");

    if (index < 0 || index === fileName.length - 1) {
        return "";
    }

    return fileName.substring(index + 1);
}

function _pm02NameMatched(fileName, nameFilters) {
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

function _pm02IgnoreMatched(fileName, ignoreFilters) {
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

function _pm02ExtendMatched(extend, extendFilters) {
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
