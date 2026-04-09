/**
 * sortFilterUtils.js
 * WPS JSA 宏工具库：排序、筛选、去重工具。
 */

/**
 * 对二维数组按指定列排序。
 *
 * @param {Array} matrix 二维数组。
 * @param {number} columnIndex 排序列索引（0 基）。
 * @param {boolean} [ascending] 是否升序，默认 true。
 * @param {boolean} [hasHeader] 是否包含表头行，默认 false。
 * @returns {Array} 排序后的新二维数组。
 */
function sort2DByColumn(matrix, columnIndex, ascending, hasHeader) {
    if (!matrix || !matrix.length) {
        return [];
    }
    var asc = ascending !== false;
    var header = hasHeader === true;
    var copy = _sfuCloneMatrix(matrix);
    var headRow = null;
    if (header) {
        headRow = copy.shift();
    }
    copy.sort(function(a, b) {
        var va = a && a.length > columnIndex ? a[columnIndex] : null;
        var vb = b && b.length > columnIndex ? b[columnIndex] : null;
        return _sfuCompareValue(va, vb, asc);
    });
    if (header) {
        copy.unshift(headRow);
    }
    return copy;
}

/**
 * 对区域数据按指定列排序。
 *
 * @param {Object} range 区域对象。
 * @param {number} columnIndex 排序列索引（1 基）。
 * @param {boolean} [ascending] 是否升序，默认 true。
 * @param {boolean} [hasHeader] 是否包含表头行，默认 false。
 * @returns {boolean} 是否排序成功。
 */
function sortRangeByColumn(range, columnIndex, ascending, hasHeader) {
    if (!range || !columnIndex) {
        return false;
    }
    var matrix = _sfuReadMatrix(range);
    if (!matrix.length) {
        return false;
    }
    var sorted = sort2DByColumn(matrix, columnIndex - 1, ascending, hasHeader);
    return _sfuWriteMatrix(range, sorted);
}

/**
 * 对区域应用自动筛选。
 *
 * @param {Object} range 区域对象。
 * @param {number} fieldIndex 字段索引（1 基）。
 * @param {*} [criteria1] 条件 1。
 * @param {number} [operator] 操作符常量。
 * @param {*} [criteria2] 条件 2。
 * @returns {boolean} 是否应用成功。
 */
function applyAutoFilter(range, fieldIndex, criteria1, operator, criteria2) {
    if (!range || !fieldIndex) {
        return false;
    }
    try {
        if (criteria1 === undefined) {
            range.AutoFilter(fieldIndex);
        } else if (operator === undefined) {
            range.AutoFilter(fieldIndex, criteria1);
        } else if (criteria2 === undefined) {
            range.AutoFilter(fieldIndex, criteria1, operator);
        } else {
            range.AutoFilter(fieldIndex, criteria1, operator, criteria2);
        }
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 清除工作表上的自动筛选状态。
 *
 * @param {Object} worksheet 工作表对象。
 * @returns {boolean} 是否清除成功。
 */
function clearAutoFilter(worksheet) {
    if (!worksheet) {
        return false;
    }
    try {
        if (worksheet.FilterMode) {
            worksheet.ShowAllData();
        }
    } catch (e1) {}
    try {
        if (worksheet.AutoFilterMode) {
            worksheet.AutoFilterMode = false;
        }
        return true;
    } catch (e2) {
        return false;
    }
}

/**
 * 对二维数组按指定键列去重。
 *
 * @param {Array} matrix 二维数组。
 * @param {Array|number} keyColumns 键列索引（0 基）或索引数组。
 * @param {boolean} [keepFirst] 是否保留首条，默认 true。
 * @returns {Array} 去重后的新二维数组。
 */
function dedupeMatrix(matrix, keyColumns, keepFirst) {
    if (!matrix || !matrix.length) {
        return [];
    }
    var keys = _sfuNormalizeColumns(keyColumns);
    if (!keys.length) {
        return _sfuCloneMatrix(matrix);
    }
    var first = keepFirst !== false;
    var map = {};
    var out = [];
    var i;
    for (i = 0; i < matrix.length; i++) {
        var row = matrix[i];
        var k = _sfuBuildKey(row, keys);
        if (!map[k]) {
            map[k] = out.length + 1;
            out.push(_sfuCloneRow(row));
        } else if (!first) {
            out[map[k] - 1] = _sfuCloneRow(row);
        }
    }
    return out;
}

/**
 * 克隆二维数组。
 *
 * @param {Array} matrix 二维数组。
 * @returns {Array} 新二维数组。
 */
function _sfuCloneMatrix(matrix) {
    var out = [];
    var i;
    for (i = 0; i < matrix.length; i++) {
        out.push(_sfuCloneRow(matrix[i]));
    }
    return out;
}

/**
 * 克隆行数组。
 *
 * @param {Array} row 行数组。
 * @returns {Array} 新行数组。
 */
function _sfuCloneRow(row) {
    if (!row || !row.length) {
        return [];
    }
    var out = [];
    var i;
    for (i = 0; i < row.length; i++) {
        out.push(row[i]);
    }
    return out;
}

/**
 * 比较两个值。
 *
 * @param {*} a 值 a。
 * @param {*} b 值 b。
 * @param {boolean} asc 是否升序。
 * @returns {number} 比较结果。
 */
function _sfuCompareValue(a, b, asc) {
    var v1 = a;
    var v2 = b;
    if (v1 === null || v1 === undefined) {
        v1 = "";
    }
    if (v2 === null || v2 === undefined) {
        v2 = "";
    }
    var n1 = Number(v1);
    var n2 = Number(v2);
    var isNum1 = !isNaN(n1) && String(v1) !== "";
    var isNum2 = !isNaN(n2) && String(v2) !== "";
    var cmp = 0;
    if (isNum1 && isNum2) {
        cmp = n1 === n2 ? 0 : (n1 > n2 ? 1 : -1);
    } else {
        var s1 = String(v1).toLowerCase();
        var s2 = String(v2).toLowerCase();
        cmp = s1 === s2 ? 0 : (s1 > s2 ? 1 : -1);
    }
    return asc ? cmp : -cmp;
}

/**
 * 标准化列索引。
 *
 * @param {Array|number} keyColumns 列索引输入。
 * @returns {Array} 列索引数组。
 */
function _sfuNormalizeColumns(keyColumns) {
    if (Object.prototype.toString.call(keyColumns) === "[object Array]") {
        return keyColumns;
    }
    if (typeof keyColumns === "number") {
        return [keyColumns];
    }
    return [];
}

/**
 * 构建行键。
 *
 * @param {Array} row 行数组。
 * @param {Array} keyColumns 键列数组。
 * @returns {string} 键字符串。
 */
function _sfuBuildKey(row, keyColumns) {
    var parts = [];
    var i;
    for (i = 0; i < keyColumns.length; i++) {
        var idx = keyColumns[i];
        parts.push(String(row && row.length > idx ? row[idx] : ""));
    }
    return parts.join("\u0001");
}

/**
 * 读取范围为二维数组。
 *
 * @param {Object} range 区域对象。
 * @returns {Array} 二维数组。
 */
function _sfuReadMatrix(range) {
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

/**
 * 写入二维数组到范围。
 *
 * @param {Object} range 区域对象。
 * @param {Array} matrix 二维数组。
 * @returns {boolean} 是否成功。
 */
function _sfuWriteMatrix(range, matrix) {
    try {
        range.Value2 = matrix;
        return true;
    } catch (e1) {
        try {
            range.Value = matrix;
            return true;
        } catch (e2) {
            return false;
        }
    }
}
