/**
 * rangeUtils.js
 * WPS JSA 宏工具库：单元格与区域读写工具。
 */

/**
 * 获取区域对象。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number|string} row 起始行号（1 基）或地址字符串（如 A1:B2）。
 * @param {number} [col] 起始列号（1 基）。
 * @param {number} [rowCount] 行数。
 * @param {number} [colCount] 列数。
 * @returns {Object|null} 区域对象，失败时返回 null。
 */
function getRange(worksheet, row, col, rowCount, colCount) {
    if (!worksheet) {
        return null;
    }
    try {
        if (typeof row === "string") {
            return worksheet.Range(row);
        }
        var rg = worksheet.Cells(row, col);
        if (typeof rowCount === "number" && typeof colCount === "number") {
            rg = rg.Resize(rowCount, colCount);
        }
        return rg;
    } catch (e) {
        return null;
    }
}

/**
 * 读取单元格值（优先 Value2）。
 *
 * @param {Object} rangeOrWorksheet 区域对象，或工作表对象。
 * @param {number} [row] 当第一个参数为工作表时使用的行号。
 * @param {number} [col] 当第一个参数为工作表时使用的列号。
 * @returns {*} 读取到的值，失败时返回 null。
 */
function readCell(rangeOrWorksheet, row, col) {
    var cell = _ruResolveCell(rangeOrWorksheet, row, col);
    if (!cell) {
        return null;
    }
    try {
        return cell.Value2;
    } catch (e1) {}
    try {
        return cell.Value;
    } catch (e2) {}
    return null;
}

/**
 * 写入单元格值（优先写 Value2）。
 *
 * @param {Object} rangeOrWorksheet 区域对象，或工作表对象。
 * @param {number} [row] 当第一个参数为工作表时使用的行号。
 * @param {number} [col] 当第一个参数为工作表时使用的列号。
 * @param {*} value 要写入的值。
 * @returns {boolean} 是否写入成功。
 */
function writeCell(rangeOrWorksheet, row, col, value) {
    var cell = _ruResolveCell(rangeOrWorksheet, row, col);
    if (!cell) {
        return false;
    }
    try {
        cell.Value2 = value;
        return true;
    } catch (e1) {}
    try {
        cell.Value = value;
        return true;
    } catch (e2) {}
    return false;
}

/**
 * 将区域读取为二维数组。
 *
 * @param {Object} range 区域对象。
 * @returns {Array} 二维数组，失败时返回空数组。
 */
function readMatrix(range) {
    if (!range) {
        return [];
    }
    var value = null;
    try {
        value = range.Value2;
    } catch (e1) {
        try {
            value = range.Value;
        } catch (e2) {
            return [];
        }
    }
    return _ruNormalizeMatrix(value);
}

/**
 * 将二维数组写入到目标区域。
 *
 * @param {Object} targetRangeOrWorksheet 目标区域，或工作表对象。
 * @param {number} [row] 当第一个参数为工作表时使用的起始行号。
 * @param {number} [col] 当第一个参数为工作表时使用的起始列号。
 * @param {Array} matrix 二维数组。
 * @returns {boolean} 是否写入成功。
 */
function writeMatrix(targetRangeOrWorksheet, row, col, matrix) {
    if (!matrix || !matrix.length) {
        return false;
    }
    var target = null;
    if (_ruIsRange(targetRangeOrWorksheet)) {
        target = targetRangeOrWorksheet;
    } else if (targetRangeOrWorksheet && typeof row === "number" && typeof col === "number") {
        target = targetRangeOrWorksheet.Cells(row, col).Resize(matrix.length, matrix[0].length);
    }
    if (!target) {
        return false;
    }
    try {
        target.Value2 = matrix;
        return true;
    } catch (e1) {
        try {
            target.Value = matrix;
            return true;
        } catch (e2) {
            return false;
        }
    }
}

/**
 * 写入 R1C1 公式。
 *
 * @param {Object} targetRangeOrWorksheet 目标区域，或工作表对象。
 * @param {number} [row] 当第一个参数为工作表时使用的行号。
 * @param {number} [col] 当第一个参数为工作表时使用的列号。
 * @param {string} formulaR1C1 R1C1 公式字符串。
 * @returns {boolean} 是否写入成功。
 */
function writeFormulaR1C1(targetRangeOrWorksheet, row, col, formulaR1C1) {
    var cell = _ruResolveCell(targetRangeOrWorksheet, row, col);
    if (!cell || !formulaR1C1) {
        return false;
    }
    try {
        cell.FormulaR1C1 = formulaR1C1;
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 清空区域内容。
 *
 * @param {Object} range 区域对象。
 * @returns {boolean} 是否清空成功。
 */
function clearRange(range) {
    if (!range) {
        return false;
    }
    try {
        range.ClearContents();
        return true;
    } catch (e1) {}
    try {
        range.Clear();
        return true;
    } catch (e2) {}
    return false;
}

/**
 * 以某区域为起点调整大小。
 *
 * @param {Object} range 区域对象。
 * @param {number} rowCount 行数。
 * @param {number} colCount 列数。
 * @returns {Object|null} 新区域对象，失败时返回 null。
 */
function resizeFrom(range, rowCount, colCount) {
    if (!range) {
        return null;
    }
    try {
        return range.Resize(rowCount, colCount);
    } catch (e) {
        return null;
    }
}

/**
 * 以某区域为起点偏移并可选调整大小。
 *
 * @param {Object} range 区域对象。
 * @param {number} rowOffset 行偏移。
 * @param {number} colOffset 列偏移。
 * @param {number} [rowCount] 可选行数。
 * @param {number} [colCount] 可选列数。
 * @returns {Object|null} 偏移后的区域对象，失败时返回 null。
 */
function offsetFrom(range, rowOffset, colOffset, rowCount, colCount) {
    if (!range) {
        return null;
    }
    try {
        var rg = range.Offset(rowOffset, colOffset);
        if (typeof rowCount === "number" && typeof colCount === "number") {
            rg = rg.Resize(rowCount, colCount);
        }
        return rg;
    } catch (e) {
        return null;
    }
}

/**
 * 在指定单元格位置冻结窗格。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} row 冻结分隔行（通常为标题下一行）。
 * @param {number} col 冻结分隔列（通常为关键列右侧一列）。
 * @param {Object} [app] 可选的应用对象。
 * @returns {boolean} 是否设置成功。
 */
function freezePanesAt(worksheet, row, col, app) {
    if (!worksheet || !row || !col) {
        return false;
    }
    var realApp = app;
    if (!realApp) {
        try {
            realApp = Application;
        } catch (e) {
            realApp = null;
        }
    }
    if (!realApp) {
        return false;
    }
    try {
        worksheet.Activate();
        worksheet.Cells(row, col).Select();
        var win = realApp.ActiveWindow;
        win.FreezePanes = false;
        win.SplitRow = row - 1;
        win.SplitColumn = col - 1;
        win.FreezePanes = true;
        return true;
    } catch (e2) {
        return false;
    }
}

/**
 * 解析为单元格对象。
 *
 * @param {Object} rangeOrWorksheet 区域或工作表对象。
 * @param {number} [row] 行号。
 * @param {number} [col] 列号。
 * @returns {Object|null} 单元格对象，失败时返回 null。
 */
function _ruResolveCell(rangeOrWorksheet, row, col) {
    if (!rangeOrWorksheet) {
        return null;
    }
    if (_ruIsRange(rangeOrWorksheet)) {
        return rangeOrWorksheet;
    }
    if (typeof row === "number" && typeof col === "number") {
        try {
            return rangeOrWorksheet.Cells(row, col);
        } catch (e) {
            return null;
        }
    }
    return null;
}

/**
 * 判断对象是否近似 Range。
 *
 * @param {Object} obj 目标对象。
 * @returns {boolean} 是否为区域对象。
 */
function _ruIsRange(obj) {
    if (!obj) {
        return false;
    }
    try {
        return obj.Cells !== undefined && obj.Rows !== undefined && obj.Columns !== undefined;
    } catch (e) {
        return false;
    }
}

/**
 * 归一化 Value2 返回值为二维数组。
 *
 * @param {*} value Value2 返回值。
 * @returns {Array} 二维数组。
 */
function _ruNormalizeMatrix(value) {
    if (value === null || value === undefined) {
        return [];
    }
    if (Object.prototype.toString.call(value) === "[object Array]") {
        if (!value.length) {
            return [];
        }
        if (Object.prototype.toString.call(value[0]) === "[object Array]") {
            return value;
        }
        return [value];
    }
    return [[value]];
}
