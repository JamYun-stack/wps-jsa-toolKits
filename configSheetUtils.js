/**
 * configSheetUtils.js
 * WPS JSA 宏工具库：配置页读取工具。
 */

/**
 * 读取单个键值配置。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {string} key 要匹配的键名。
 * @param {number} [keyCol] 键列（1 基），默认 1。
 * @param {number} [valueCol] 值列（1 基），默认 2。
 * @param {number} [startRow] 起始行，默认 1。
 * @param {number} [endRow] 结束行，默认最后有效行。
 * @param {*} [defaultValue] 未命中时默认值。
 * @returns {*} 配置值；未命中返回 defaultValue。
 */
function readConfigValue(worksheet, key, keyCol, valueCol, startRow, endRow, defaultValue) {
    if (!worksheet || !key) {
        return defaultValue;
    }
    var kc = typeof keyCol === "number" && keyCol > 0 ? keyCol : 1;
    var vc = typeof valueCol === "number" && valueCol > 0 ? valueCol : 2;
    var sr = typeof startRow === "number" && startRow > 0 ? startRow : 1;
    var er = typeof endRow === "number" && endRow > 0 ? endRow : _csuGetLastRow(worksheet, kc);
    var i;
    for (i = sr; i <= er; i++) {
        var k = _csuCellValue(worksheet, i, kc);
        if (String(k) === String(key)) {
            return _csuCellValue(worksheet, i, vc);
        }
    }
    return defaultValue;
}

/**
 * 读取键值配置区域为对象。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} [keyCol] 键列，默认 1。
 * @param {number} [valueCol] 值列，默认 2。
 * @param {number} [startRow] 起始行，默认 1。
 * @param {number} [endRow] 结束行，默认最后有效行。
 * @returns {Object} 键值对象。
 * 返回格式：
 * {
 *   "inputPath": "D:\\input",
 *   "outputPath": "D:\\output"
 * }
 */
function readKeyValueConfig(worksheet, keyCol, valueCol, startRow, endRow) {
    var obj = {};
    if (!worksheet) {
        return obj;
    }
    var kc = typeof keyCol === "number" && keyCol > 0 ? keyCol : 1;
    var vc = typeof valueCol === "number" && valueCol > 0 ? valueCol : 2;
    var sr = typeof startRow === "number" && startRow > 0 ? startRow : 1;
    var er = typeof endRow === "number" && endRow > 0 ? endRow : _csuGetLastRow(worksheet, kc);
    var i;
    for (i = sr; i <= er; i++) {
        var k = _csuCellValue(worksheet, i, kc);
        if (k === null || k === undefined || String(k) === "") {
            continue;
        }
        obj[String(k)] = _csuCellValue(worksheet, i, vc);
    }
    return obj;
}

/**
 * 读取字段映射配置（字段名 -> 表头名）。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} [fieldCol] 字段名列，默认 1。
 * @param {number} [headerCol] 表头名列，默认 2。
 * @param {number} [startRow] 起始行，默认 2。
 * @param {number} [endRow] 结束行，默认最后有效行。
 * @returns {Object} 映射对象。
 * 返回格式：
 * {
 *   "shopName": "店铺名",
 *   "amount": "销售额"
 * }
 */
function readFieldMapConfig(worksheet, fieldCol, headerCol, startRow, endRow) {
    var map = {};
    if (!worksheet) {
        return map;
    }
    var fc = typeof fieldCol === "number" && fieldCol > 0 ? fieldCol : 1;
    var hc = typeof headerCol === "number" && headerCol > 0 ? headerCol : 2;
    var sr = typeof startRow === "number" && startRow > 0 ? startRow : 2;
    var er = typeof endRow === "number" && endRow > 0 ? endRow : _csuGetLastRow(worksheet, fc);
    var i;
    for (i = sr; i <= er; i++) {
        var field = _csuCellValue(worksheet, i, fc);
        if (field === null || field === undefined || String(field) === "") {
            continue;
        }
        map[String(field)] = _csuCellValue(worksheet, i, hc);
    }
    return map;
}

/**
 * 读取路径配置对象。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} [startRow] 起始行，默认 1。
 * @param {number} [endRow] 结束行，默认最后有效行。
 * @returns {Object} 路径配置对象。
 * 返回格式：
 * {
 *   inputPath: "D:\\input",
 *   outputPath: "D:\\output",
 *   templatePath: "D:\\tpl\\demo.xlsx"
 * }
 */
function readPathConfig(worksheet, startRow, endRow) {
    var raw = readKeyValueConfig(worksheet, 1, 2, startRow, endRow);
    var out = {
        inputPath: null,
        outputPath: null,
        templatePath: null
    };
    out.inputPath = _csuPickValue(raw, ["inputPath", "输入路径", "源路径", "sourcePath"]);
    out.outputPath = _csuPickValue(raw, ["outputPath", "输出路径", "targetPath"]);
    out.templatePath = _csuPickValue(raw, ["templatePath", "模板路径", "template"]);
    return out;
}

/**
 * 读取店铺分类配置列表。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} [startRow] 起始行，默认 2。
 * @param {number} [endRow] 结束行，默认最后有效行。
 * @param {number} [shopCol] 店铺列，默认 1。
 * @param {number} [categoryCol] 分类列，默认 2。
 * @param {number} [valueCol] 值列，默认 3。
 * @returns {Array} 配置数组。
 * 返回格式：
 * [
 *   { shop: "店铺A", category: "零食-100薯片", value: "Y" }
 * ]
 */
function readShopCategoryConfig(worksheet, startRow, endRow, shopCol, categoryCol, valueCol) {
    var out = [];
    if (!worksheet) {
        return out;
    }
    var sr = typeof startRow === "number" && startRow > 0 ? startRow : 2;
    var sc = typeof shopCol === "number" && shopCol > 0 ? shopCol : 1;
    var cc = typeof categoryCol === "number" && categoryCol > 0 ? categoryCol : 2;
    var vc = typeof valueCol === "number" && valueCol > 0 ? valueCol : 3;
    var er = typeof endRow === "number" && endRow > 0 ? endRow : _csuGetLastRow(worksheet, sc);
    var i;
    for (i = sr; i <= er; i++) {
        var shop = _csuCellValue(worksheet, i, sc);
        var category = _csuCellValue(worksheet, i, cc);
        var value = _csuCellValue(worksheet, i, vc);
        if ((shop === null || shop === undefined || String(shop) === "") &&
            (category === null || category === undefined || String(category) === "")) {
            continue;
        }
        out.push({
            shop: shop,
            category: category,
            value: value
        });
    }
    return out;
}

/**
 * 读取单元格值（优先 Value2）。
 *
 * @param {Object} ws 工作表对象。
 * @param {number} row 行号。
 * @param {number} col 列号。
 * @returns {*} 值。
 */
function _csuCellValue(ws, row, col) {
    try {
        return ws.Cells(row, col).Value2;
    } catch (e1) {
        try {
            return ws.Cells(row, col).Value;
        } catch (e2) {
            return null;
        }
    }
}

/**
 * 选取对象中的第一个存在值。
 *
 * @param {Object} obj 对象。
 * @param {Array} keys 键名数组。
 * @returns {*} 命中的值，未命中返回 null。
 */
function _csuPickValue(obj, keys) {
    var i;
    for (i = 0; i < keys.length; i++) {
        var k = keys[i];
        if (Object.prototype.hasOwnProperty.call(obj, k)) {
            return obj[k];
        }
    }
    return null;
}

/**
 * 获取最后有效行。
 *
 * @param {Object} ws 工作表对象。
 * @param {number} col 列号。
 * @returns {number} 最后行。
 */
function _csuGetLastRow(ws, col) {
    var c = typeof col === "number" && col > 0 ? col : 1;
    try {
        var xlUp = -4162;
        return ws.Cells(ws.Rows.Count, c).End(xlUp).Row;
    } catch (e) {
        try {
            return ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
        } catch (e2) {
            return 1;
        }
    }
}
