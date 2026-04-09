/**
 * formatUtils.js
 * WPS JSA 宏工具库：格式设置工具。
 */

/**
 * 复制行格式。
 *
 * @param {Object} sourceRowRange 源行区域。
 * @param {Object} targetRowRange 目标行区域。
 * @returns {boolean} 是否复制成功。
 */
function copyRowFormats(sourceRowRange, targetRowRange) {
    return copyRangeFormats(sourceRowRange, targetRowRange);
}

/**
 * 复制区域格式。
 *
 * @param {Object} sourceRange 源区域。
 * @param {Object} targetRange 目标区域。
 * @returns {boolean} 是否复制成功。
 */
function copyRangeFormats(sourceRange, targetRange) {
    if (!sourceRange || !targetRange) {
        return false;
    }
    try {
        sourceRange.Copy();
        var xlPasteFormats = -4122;
        targetRange.PasteSpecial(xlPasteFormats);
        try {
            Application.CutCopyMode = false;
        } catch (e0) {}
        return true;
    } catch (e1) {}
    try {
        sourceRange.Copy(targetRange);
        return true;
    } catch (e2) {
        return false;
    }
}

/**
 * 设置数字格式。
 *
 * @param {Object} range 区域对象。
 * @param {string} formatText 数字格式字符串。
 * @returns {boolean} 是否设置成功。
 */
function setNumberFormat(range, formatText) {
    if (!range || !formatText) {
        return false;
    }
    try {
        range.NumberFormat = formatText;
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 设置字体样式。
 *
 * @param {Object} range 区域对象。
 * @param {string} [fontName] 字体名称。
 * @param {number} [fontSize] 字号。
 * @param {boolean} [bold] 是否加粗。
 * @param {boolean} [italic] 是否斜体。
 * @param {number} [color] 字体颜色（RGB 整数）。
 * @returns {boolean} 是否设置成功。
 */
function setFontStyle(range, fontName, fontSize, bold, italic, color) {
    if (!range) {
        return false;
    }
    try {
        if (fontName !== undefined && fontName !== null && fontName !== "") {
            range.Font.Name = fontName;
        }
        if (typeof fontSize === "number" && fontSize > 0) {
            range.Font.Size = fontSize;
        }
        if (typeof bold === "boolean") {
            range.Font.Bold = bold;
        }
        if (typeof italic === "boolean") {
            range.Font.Italic = italic;
        }
        if (typeof color === "number") {
            range.Font.Color = color;
        }
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 设置填充色。
 *
 * @param {Object} range 区域对象。
 * @param {number} color RGB 整数颜色。
 * @returns {boolean} 是否设置成功。
 */
function setFillColor(range, color) {
    if (!range || typeof color !== "number") {
        return false;
    }
    try {
        range.Interior.Color = color;
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 设置边框样式。
 *
 * @param {Object} range 区域对象。
 * @param {number} [lineStyle] 线型常量。
 * @param {number} [weight] 粗细常量。
 * @param {number} [color] RGB 整数颜色。
 * @returns {boolean} 是否设置成功。
 */
function setBorderStyle(range, lineStyle, weight, color) {
    if (!range) {
        return false;
    }
    try {
        var borders = range.Borders;
        if (typeof lineStyle === "number") {
            borders.LineStyle = lineStyle;
        }
        if (typeof weight === "number") {
            borders.Weight = weight;
        }
        if (typeof color === "number") {
            borders.Color = color;
        }
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 自动调整列宽。
 *
 * @param {Object} rangeOrWorksheet 区域对象或工作表对象。
 * @returns {boolean} 是否设置成功。
 */
function autoFitColumns(rangeOrWorksheet) {
    if (!rangeOrWorksheet) {
        return false;
    }
    try {
        if (rangeOrWorksheet.Columns && rangeOrWorksheet.Columns.AutoFit) {
            rangeOrWorksheet.Columns.AutoFit();
            return true;
        }
    } catch (e1) {}
    try {
        rangeOrWorksheet.UsedRange.Columns.AutoFit();
        return true;
    } catch (e2) {
        return false;
    }
}

/**
 * 自动调整行高。
 *
 * @param {Object} rangeOrWorksheet 区域对象或工作表对象。
 * @returns {boolean} 是否设置成功。
 */
function autoFitRows(rangeOrWorksheet) {
    if (!rangeOrWorksheet) {
        return false;
    }
    try {
        if (rangeOrWorksheet.Rows && rangeOrWorksheet.Rows.AutoFit) {
            rangeOrWorksheet.Rows.AutoFit();
            return true;
        }
    } catch (e1) {}
    try {
        rangeOrWorksheet.UsedRange.Rows.AutoFit();
        return true;
    } catch (e2) {
        return false;
    }
}
