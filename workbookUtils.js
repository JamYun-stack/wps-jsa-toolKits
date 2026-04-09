/**
 * workbookUtils.js
 * WPS JSA 宏工具库：工作簿管理工具。
 */

/**
 * 获取可用的 Application 对象。
 *
 * @param {Object} [app] 可选的应用对象。
 * @returns {Object|null} 可用的应用对象，失败时返回 null。
 */
function getApp(app) {
    if (app) {
        return app;
    }
    try {
        if (typeof Application !== "undefined" && Application) {
            return Application;
        }
    } catch (e1) {}
    try {
        if (typeof EtApplication !== "undefined" && EtApplication) {
            return EtApplication;
        }
    } catch (e2) {}
    return null;
}

/**
 * 获取当前活动工作簿。
 *
 * @param {Object} [app] 可选的应用对象。
 * @returns {Object|null} 当前活动工作簿对象，失败时返回 null。
 */
function getActiveWorkbook(app) {
    var realApp = getApp(app);
    if (!realApp) {
        return null;
    }
    try {
        return realApp.ActiveWorkbook;
    } catch (e) {
        return null;
    }
}

/**
 * 打开指定路径的工作簿。
 *
 * @param {string} path 工作簿完整路径。
 * @param {boolean} [readOnly] 是否只读打开。
 * @param {Object} [app] 可选的应用对象。
 * @returns {Object|null} 打开的工作簿对象，失败时返回 null。
 */
function openWorkbook(path, readOnly, app) {
    var realApp = getApp(app);
    if (!realApp || !path) {
        return null;
    }
    try {
        return realApp.Workbooks.Open(path, undefined, !!readOnly);
    } catch (e1) {
        try {
            return realApp.Workbooks.Open(path);
        } catch (e2) {
            return null;
        }
    }
}

/**
 * 将工作簿另存为指定路径。
 *
 * @param {Object} workbook 工作簿对象。
 * @param {string} savePath 目标完整路径。
 * @param {number} [fileFormat] 文件格式常量（可选）。
 * @param {boolean} [overwrite] 目标存在时是否覆盖。
 * @returns {boolean} 是否保存成功。
 */
function saveWorkbookAs(workbook, savePath, fileFormat, overwrite) {
    if (!workbook || !savePath) {
        return false;
    }
    if (_wbFileExists(savePath)) {
        if (!overwrite) {
            return false;
        }
        _wbTryDeleteFile(savePath);
    }
    try {
        if (typeof fileFormat === "number") {
            workbook.SaveAs(savePath, fileFormat);
        } else {
            workbook.SaveAs(savePath);
        }
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 安全关闭工作簿。
 *
 * @param {Object} workbook 工作簿对象。
 * @param {boolean} [saveChanges] 关闭前是否保存。
 * @returns {boolean} 是否关闭成功。
 */
function closeWorkbookSafe(workbook, saveChanges) {
    if (!workbook) {
        return false;
    }
    try {
        workbook.Close(!!saveChanges);
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 将指定工作表复制到新工作簿。
 *
 * @param {Object} worksheet 源工作表对象。
 * @param {string} [newWorkbookName] 新工作簿名称（仅写入 Title）。
 * @param {Object} [app] 可选的应用对象。
 * @returns {Object|null} 新工作簿对象，失败时返回 null。
 */
function copyWorksheetToNewWorkbook(worksheet, newWorkbookName, app) {
    var realApp = getApp(app);
    if (!worksheet || !realApp) {
        return null;
    }
    try {
        worksheet.Copy();
        var newBook = realApp.ActiveWorkbook;
        if (newBook && newWorkbookName) {
            try {
                newBook.Title = newWorkbookName;
            } catch (eTitle) {}
        }
        return newBook || null;
    } catch (e) {
        return null;
    }
}

/**
 * 设置并返回 DisplayAlerts。
 *
 * @param {Object} [app] 可选的应用对象。
 * @param {boolean} value 目标值。
 * @returns {boolean|null} 修改前的值，失败时返回 null。
 */
function setDisplayAlerts(app, value) {
    var realApp = getApp(app);
    if (!realApp) {
        return null;
    }
    try {
        var oldValue = !!realApp.DisplayAlerts;
        realApp.DisplayAlerts = !!value;
        return oldValue;
    } catch (e) {
        return null;
    }
}

/**
 * 判断文件是否存在。
 *
 * @param {string} path 文件完整路径。
 * @returns {boolean} 是否存在。
 */
function _wbFileExists(path) {
    if (!path) {
        return false;
    }
    try {
        return !!Dir(path);
    } catch (e1) {}
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        return fso.FileExists(path);
    } catch (e2) {}
    return false;
}

/**
 * 尝试删除文件。
 *
 * @param {string} path 文件完整路径。
 * @returns {boolean} 是否删除成功。
 */
function _wbTryDeleteFile(path) {
    if (!path) {
        return false;
    }
    try {
        Kill(path);
        return true;
    } catch (e1) {}
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (fso.FileExists(path)) {
            fso.DeleteFile(path, true);
        }
        return true;
    } catch (e2) {}
    return false;
}
