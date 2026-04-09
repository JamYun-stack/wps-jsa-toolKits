/**
 * dataImportUtils.js
 * WPS JSA 宏工具库：外部数据读取与对象化工具。
 */

/**
 * 读取工作表区域为二维数组。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} [startRow] 起始行，默认 1。
 * @param {number} [startCol] 起始列，默认 1。
 * @param {number} [endRow] 结束行，默认 UsedRange 末行。
 * @param {number} [endCol] 结束列，默认 UsedRange 末列。
 * @returns {Array} 二维数组。
 */
function readWorksheetMatrix(worksheet, startRow, startCol, endRow, endCol) {
    if (!worksheet) {
        return [];
    }
    var sr = typeof startRow === "number" && startRow > 0 ? startRow : 1;
    var sc = typeof startCol === "number" && startCol > 0 ? startCol : 1;
    var er = typeof endRow === "number" && endRow > 0 ? endRow : _diGetLastRow(worksheet);
    var ec = typeof endCol === "number" && endCol > 0 ? endCol : _diGetLastCol(worksheet);
    if (er < sr || ec < sc) {
        return [];
    }
    try {
        var range = worksheet.Range(worksheet.Cells(sr, sc), worksheet.Cells(er, ec));
        var value = range.Value2;
        if (value === null || value === undefined) {
            return [];
        }
        if (Object.prototype.toString.call(value) === "[object Array]") {
            if (value.length && Object.prototype.toString.call(value[0]) === "[object Array]") {
                return value;
            }
            return [value];
        }
        return [[value]];
    } catch (e) {
        return [];
    }
}

/**
 * 从工作簿第一张表按“字段 -> 表头名”映射读取数据。
 *
 * @param {Object} workbook 工作簿对象。
 * @param {Object} headerMap 字段到表头名映射。
 * @param {number} [headerRow] 表头行号，默认 1。
 * @param {number} [startRow] 数据起始行，默认 `headerRow + 1`。
 * @returns {Array} 对象数组。
 */
function readFirstSheetByHeaderMap(workbook, headerMap, headerRow, startRow) {
    if (!workbook || !headerMap) {
        return [];
    }
    var ws = null;
    try {
        ws = workbook.Worksheets(1);
    } catch (e) {
        return [];
    }
    var hr = typeof headerRow === "number" && headerRow > 0 ? headerRow : 1;
    var sr = typeof startRow === "number" && startRow > 0 ? startRow : hr + 1;
    var endRow = _diGetLastRow(ws);
    var endCol = _diGetLastCol(ws);
    if (endRow < sr || endCol < 1) {
        return [];
    }
    var headers = readWorksheetMatrix(ws, hr, 1, hr, endCol);
    if (!headers.length) {
        return [];
    }
    var headerIndex = _diBuildHeaderIndex(headers[0]);
    var rows = readWorksheetMatrix(ws, sr, 1, endRow, endCol);
    var mapKeys = _diObjectKeys(headerMap);
    var out = [];
    var i;
    for (i = 0; i < rows.length; i++) {
        var row = rows[i];
        var item = {};
        var j;
        for (j = 0; j < mapKeys.length; j++) {
            var field = mapKeys[j];
            var headName = String(headerMap[field]);
            var idx = headerIndex[headName];
            item[field] = idx >= 0 ? row[idx] : null;
        }
        out.push(item);
    }
    return out;
}

/**
 * 将对象数组按指定键字段索引为对象。
 *
 * @param {Array} rows 对象数组。
 * @param {string} keyField 键字段名。
 * @returns {Object} 索引对象。
 * 返回格式：
 * {
 *   "A001": { id: "A001", name: "苹果" },
 *   "A002": { id: "A002", name: "香蕉" }
 * }
 */
function indexRowsByKey(rows, keyField) {
    var map = {};
    if (!rows || !rows.length || !keyField) {
        return map;
    }
    var i;
    for (i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (!row) {
            continue;
        }
        var k = row[keyField];
        if (k === null || k === undefined || k === "") {
            continue;
        }
        map[String(k)] = row;
    }
    return map;
}

/**
 * 将工作表读取为对象数组（首行标题）。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} [headerRow] 表头行号，默认 1。
 * @param {number} [startRow] 数据起始行，默认 `headerRow + 1`。
 * @param {number} [endRow] 数据结束行，默认 UsedRange 末行。
 * @param {number} [endCol] 数据结束列，默认 UsedRange 末列。
 * @returns {Array} 对象数组。
 */
function readRowsAsObjects(worksheet, headerRow, startRow, endRow, endCol) {
    if (!worksheet) {
        return [];
    }
    var hr = typeof headerRow === "number" && headerRow > 0 ? headerRow : 1;
    var sr = typeof startRow === "number" && startRow > 0 ? startRow : hr + 1;
    var er = typeof endRow === "number" && endRow > 0 ? endRow : _diGetLastRow(worksheet);
    var ec = typeof endCol === "number" && endCol > 0 ? endCol : _diGetLastCol(worksheet);
    if (er < sr || ec < 1) {
        return [];
    }
    var headerMatrix = readWorksheetMatrix(worksheet, hr, 1, hr, ec);
    if (!headerMatrix.length) {
        return [];
    }
    var headers = headerMatrix[0];
    var rows = readWorksheetMatrix(worksheet, sr, 1, er, ec);
    var out = [];
    var i;
    for (i = 0; i < rows.length; i++) {
        var row = rows[i];
        var obj = {};
        var c;
        for (c = 0; c < headers.length; c++) {
            var key = String(headers[c] === null || headers[c] === undefined ? "" : headers[c]);
            if (!key) {
                continue;
            }
            obj[key] = row[c];
        }
        out.push(obj);
    }
    return out;
}

/**
 * 构建表头索引。
 *
 * @param {Array} headers 表头数组。
 * @returns {Object} 索引对象。
 */
function _diBuildHeaderIndex(headers) {
    var map = {};
    var i;
    for (i = 0; i < headers.length; i++) {
        var k = String(headers[i] === null || headers[i] === undefined ? "" : headers[i]);
        if (!k) {
            continue;
        }
        if (map[k] === undefined) {
            map[k] = i;
        }
    }
    return map;
}

/**
 * 获取对象键数组。
 *
 * @param {Object} obj 对象。
 * @returns {Array} 键数组。
 */
function _diObjectKeys(obj) {
    var keys = [];
    var k;
    for (k in obj) {
        if (Object.prototype.hasOwnProperty.call(obj, k)) {
            keys.push(k);
        }
    }
    return keys;
}

/**
 * 获取最后行。
 *
 * @param {Object} ws 工作表对象。
 * @returns {number} 最后行。
 */
function _diGetLastRow(ws) {
    try {
        var xlUp = -4162;
        return ws.Cells(ws.Rows.Count, 1).End(xlUp).Row;
    } catch (e) {
        try {
            return ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
        } catch (e2) {
            return 1;
        }
    }
}

/**
 * 获取最后列。
 *
 * @param {Object} ws 工作表对象。
 * @returns {number} 最后列。
 */
function _diGetLastCol(ws) {
    try {
        var xlToLeft = -4159;
        return ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column;
    } catch (e) {
        try {
            return ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1;
        } catch (e2) {
            return 1;
        }
    }
}
