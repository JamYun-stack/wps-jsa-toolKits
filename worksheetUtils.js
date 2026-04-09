/**
 * worksheetUtils.js
 * WPS JSA 宏工具库：工作表管理工具。
 */

/**
 * 获取工作表对象（按名称或索引）。
 *
 * @param {Object} workbook 工作簿对象。
 * @param {string|number} sheetNameOrIndex 工作表名称或 1 基索引。
 * @returns {Object|null} 工作表对象，失败时返回 null。
 */
function getWorksheet(workbook, sheetNameOrIndex) {
    if (!workbook || sheetNameOrIndex === null || sheetNameOrIndex === undefined) {
        return null;
    }
    try {
        return workbook.Worksheets(sheetNameOrIndex);
    } catch (e1) {}
    try {
        return workbook.Sheets(sheetNameOrIndex);
    } catch (e2) {}
    return null;
}

/**
 * 确保工作簿存在指定名称的工作表，不存在则新建。
 *
 * @param {Object} workbook 工作簿对象。
 * @param {string} sheetName 工作表名称。
 * @param {number} [position] 新建时插入位置（1 基）。
 * @returns {Object|null} 工作表对象，失败时返回 null。
 */
function ensureWorksheet(workbook, sheetName, position) {
    if (!workbook || !sheetName) {
        return null;
    }
    var ws = getWorksheet(workbook, sheetName);
    if (ws) {
        return ws;
    }
    try {
        if (typeof position === "number" && position >= 1 && position <= workbook.Worksheets.Count) {
            ws = workbook.Worksheets.Add(null, workbook.Worksheets(position));
        } else {
            ws = workbook.Worksheets.Add();
        }
        ws.Name = sheetName;
        return ws;
    } catch (e) {
        return null;
    }
}

/**
 * 获取工作簿中的全部工作表名称。
 *
 * @param {Object} workbook 工作簿对象。
 * @returns {Array} 名称数组，失败时返回空数组。
 */
function listWorksheetNames(workbook) {
    var list = [];
    if (!workbook) {
        return list;
    }
    try {
        var count = workbook.Worksheets.Count;
        var i;
        for (i = 1; i <= count; i++) {
            list.push(String(workbook.Worksheets(i).Name));
        }
    } catch (e) {}
    return list;
}

/**
 * 获取 UsedRange 的边界信息。
 *
 * @param {Object} worksheet 工作表对象。
 * @returns {Object|null} 边界信息对象，失败时返回 null。
 * 返回格式：
 * {
 *   row: 1,
 *   col: 1,
 *   rows: 10,
 *   cols: 5,
 *   lastRow: 10,
 *   lastCol: 5,
 *   address: "$A$1:$E$10"
 * }
 */
function getUsedRangeBounds(worksheet) {
    if (!worksheet) {
        return null;
    }
    try {
        var used = worksheet.UsedRange;
        var row = used.Row;
        var col = used.Column;
        var rows = used.Rows.Count;
        var cols = used.Columns.Count;
        return {
            row: row,
            col: col,
            rows: rows,
            cols: cols,
            lastRow: row + rows - 1,
            lastCol: col + cols - 1,
            address: String(used.Address)
        };
    } catch (e) {
        return null;
    }
}

/**
 * 查找指定列的最后一个非空行。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} [columnIndex] 列号（1 基），默认 1。
 * @param {number} [startRow] 最小返回行号，默认 1。
 * @returns {number} 最后一行行号，失败时返回 0。
 */
function findLastRow(worksheet, columnIndex, startRow) {
    if (!worksheet) {
        return 0;
    }
    var col = typeof columnIndex === "number" && columnIndex > 0 ? columnIndex : 1;
    var minRow = typeof startRow === "number" && startRow > 0 ? startRow : 1;
    try {
        var xlUp = -4162;
        var row = worksheet.Cells(worksheet.Rows.Count, col).End(xlUp).Row;
        if (row < minRow) {
            row = minRow;
        }
        return row;
    } catch (e) {
        var bounds = getUsedRangeBounds(worksheet);
        if (!bounds) {
            return 0;
        }
        if (bounds.lastRow < minRow) {
            return minRow;
        }
        return bounds.lastRow;
    }
}

/**
 * 查找指定行的最后一个非空列。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} [rowIndex] 行号（1 基），默认 1。
 * @param {number} [startColumn] 最小返回列号，默认 1。
 * @returns {number} 最后一列列号，失败时返回 0。
 */
function findLastColumn(worksheet, rowIndex, startColumn) {
    if (!worksheet) {
        return 0;
    }
    var row = typeof rowIndex === "number" && rowIndex > 0 ? rowIndex : 1;
    var minCol = typeof startColumn === "number" && startColumn > 0 ? startColumn : 1;
    try {
        var xlToLeft = -4159;
        var col = worksheet.Cells(row, worksheet.Columns.Count).End(xlToLeft).Column;
        if (col < minCol) {
            col = minCol;
        }
        return col;
    } catch (e) {
        var bounds = getUsedRangeBounds(worksheet);
        if (!bounds) {
            return 0;
        }
        if (bounds.lastCol < minCol) {
            return minCol;
        }
        return bounds.lastCol;
    }
}
