/**
 * tableQueryUtils.js
 * WPS JSA 宏工具库：表对象（ListObject）与查询表工具。
 */

/**
 * 创建表对象。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {Object|string} range 区域对象或地址字符串。
 * @param {string} [tableName] 表名称。
 * @param {boolean} [hasHeaders] 是否包含表头，默认 true。
 * @returns {Object|null} 表对象，失败返回 null。
 */
function createTable(worksheet, range, tableName, hasHeaders) {
    if (!worksheet || !range) {
        return null;
    }
    var rg = null;
    try {
        if (typeof range === "string") {
            rg = worksheet.Range(range);
        } else {
            rg = range;
        }
    } catch (e1) {
        return null;
    }
    try {
        var xlSrcRange = 1;
        var xlYes = 1;
        var xlNo = 2;
        var headerFlag = hasHeaders === false ? xlNo : xlYes;
        var table = worksheet.ListObjects.Add(xlSrcRange, rg, undefined, headerFlag);
        if (tableName) {
            try {
                table.Name = tableName;
            } catch (eName) {}
        }
        return table;
    } catch (e2) {
        return null;
    }
}

/**
 * 按名称获取表对象。
 *
 * @param {Object} workbook 工作簿对象。
 * @param {string} tableName 表名称。
 * @returns {Object|null} 表对象，未找到返回 null。
 */
function getTableByName(workbook, tableName) {
    if (!workbook || !tableName) {
        return null;
    }
    try {
        var wsCount = workbook.Worksheets.Count;
        var i;
        for (i = 1; i <= wsCount; i++) {
            var ws = workbook.Worksheets(i);
            var listCount = ws.ListObjects.Count;
            var j;
            for (j = 1; j <= listCount; j++) {
                var tb = ws.ListObjects(j);
                if (String(tb.Name) === String(tableName)) {
                    return tb;
                }
            }
        }
    } catch (e) {}
    return null;
}

/**
 * 追加一行到表对象中。
 *
 * @param {Object} table 表对象。
 * @param {Array} rowValues 行值数组。
 * @returns {boolean} 是否追加成功。
 */
function appendTableRow(table, rowValues) {
    if (!table || !rowValues) {
        return false;
    }
    try {
        var newRow = table.ListRows.Add();
        var c;
        for (c = 0; c < rowValues.length; c++) {
            newRow.Range.Cells(1, c + 1).Value2 = rowValues[c];
        }
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 向表对象新增列并可选填充默认值。
 *
 * @param {Object} table 表对象。
 * @param {string} columnName 列名。
 * @param {*} [defaultValue] 默认值。
 * @returns {boolean} 是否新增成功。
 */
function addTableColumn(table, columnName, defaultValue) {
    if (!table || !columnName) {
        return false;
    }
    try {
        var col = table.ListColumns.Add();
        col.Name = columnName;
        if (defaultValue !== undefined) {
            try {
                if (table.DataBodyRange) {
                    col.DataBodyRange.Value2 = defaultValue;
                }
            } catch (eSet) {}
        }
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 对表对象按列排序。
 *
 * @param {Object} table 表对象。
 * @param {string|number} keyColumnName 列名或 1 基列号。
 * @param {boolean} [ascending] 是否升序，默认 true。
 * @param {boolean} [hasHeader] 是否有表头，默认 true。
 * @returns {boolean} 是否排序成功。
 */
function sortTable(table, keyColumnName, ascending, hasHeader) {
    if (!table || !keyColumnName) {
        return false;
    }
    try {
        var idx = 0;
        if (typeof keyColumnName === "number") {
            idx = keyColumnName;
        } else {
            idx = table.ListColumns(String(keyColumnName)).Index;
        }
        var range = table.Range;
        var matrix = _tquReadRangeMatrix(range);
        if (!matrix.length) {
            return false;
        }
        var withHeader = hasHeader !== false;
        var header = null;
        var body = matrix;
        if (withHeader) {
            header = body.shift();
        }
        var asc = ascending !== false;
        body.sort(function(a, b) {
            var va = a[idx - 1];
            var vb = b[idx - 1];
            if (va === vb) {
                return 0;
            }
            if (va === null || va === undefined) {
                return asc ? -1 : 1;
            }
            if (vb === null || vb === undefined) {
                return asc ? 1 : -1;
            }
            if (va > vb) {
                return asc ? 1 : -1;
            }
            return asc ? -1 : 1;
        });
        if (withHeader && header) {
            body.unshift(header);
        }
        range.Value2 = body;
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 刷新查询表。
 *
 * @param {Object} tableOrWorksheet 表对象或工作表对象。
 * @param {string} [tableName] 当第一个参数为工作表时使用的表名。
 * @returns {boolean} 是否刷新成功。
 */
function refreshQueryTable(tableOrWorksheet, tableName) {
    if (!tableOrWorksheet) {
        return false;
    }
    var table = null;
    if (tableOrWorksheet.ListColumns !== undefined && tableOrWorksheet.ListRows !== undefined) {
        table = tableOrWorksheet;
    } else if (tableName) {
        try {
            table = tableOrWorksheet.ListObjects(tableName);
        } catch (e1) {
            table = null;
        }
    }
    if (!table) {
        return false;
    }
    try {
        if (table.QueryTable) {
            table.QueryTable.Refresh(false);
            return true;
        }
    } catch (e2) {}
    try {
        table.Refresh();
        return true;
    } catch (e3) {
        return false;
    }
}

/**
 * 读取区域为二维数组。
 *
 * @param {Object} range 区域对象。
 * @returns {Array} 二维数组。
 */
function _tquReadRangeMatrix(range) {
    try {
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
