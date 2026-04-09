/**
 * shapeUtils.js
 * WPS JSA 宏工具库：形状与图片工具。
 */

/**
 * 添加形状。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {number} type 形状类型常量。
 * @param {number} left 左位置。
 * @param {number} top 上位置。
 * @param {number} width 宽度。
 * @param {number} height 高度。
 * @param {string} [text] 形状文字。
 * @returns {Object|null} 形状对象，失败返回 null。
 */
function addShape(worksheet, type, left, top, width, height, text) {
    if (!worksheet) {
        return null;
    }
    try {
        var shp = worksheet.Shapes.AddShape(type, left, top, width, height);
        if (text !== undefined && text !== null) {
            try {
                shp.TextFrame.Characters().Text = String(text);
            } catch (eText1) {
                try {
                    shp.TextFrame2.TextRange.Text = String(text);
                } catch (eText2) {}
            }
        }
        return shp;
    } catch (e) {
        return null;
    }
}

/**
 * 移动形状位置。
 *
 * @param {Object} shape 形状对象。
 * @param {number} left 左位置。
 * @param {number} top 上位置。
 * @returns {boolean} 是否移动成功。
 */
function moveShape(shape, left, top) {
    if (!shape) {
        return false;
    }
    try {
        shape.Left = left;
        shape.Top = top;
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 调整形状大小。
 *
 * @param {Object} shape 形状对象。
 * @param {number} width 宽度。
 * @param {number} height 高度。
 * @param {boolean} [lockAspectRatio] 是否锁定纵横比。
 * @returns {boolean} 是否调整成功。
 */
function resizeShape(shape, width, height, lockAspectRatio) {
    if (!shape) {
        return false;
    }
    try {
        if (typeof lockAspectRatio === "boolean") {
            shape.LockAspectRatio = lockAspectRatio ? 1 : 0;
        }
        if (typeof width === "number" && width > 0) {
            shape.Width = width;
        }
        if (typeof height === "number" && height > 0) {
            shape.Height = height;
        }
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 按名称分组形状。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {Array} shapeNames 形状名称数组。
 * @returns {Object|null} 分组对象，失败返回 null。
 */
function groupShapes(worksheet, shapeNames) {
    if (!worksheet || !shapeNames || !shapeNames.length) {
        return null;
    }
    try {
        var names = _shuBuildNamesArray(shapeNames);
        return worksheet.Shapes.Range(names).Group();
    } catch (e) {
        return null;
    }
}

/**
 * 从文件插入图片。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {string} imagePath 图片路径。
 * @param {number} left 左位置。
 * @param {number} top 上位置。
 * @param {number} [width] 宽度。
 * @param {number} [height] 高度。
 * @param {boolean} [linkToFile] 是否链接到文件。
 * @param {boolean} [saveWithDocument] 是否随文档保存。
 * @returns {Object|null} 图片形状对象，失败返回 null。
 */
function insertImageFromFile(worksheet, imagePath, left, top, width, height, linkToFile, saveWithDocument) {
    if (!worksheet || !imagePath) {
        return null;
    }
    var w = typeof width === "number" ? width : -1;
    var h = typeof height === "number" ? height : -1;
    var link = linkToFile === true ? 1 : 0;
    var save = saveWithDocument === false ? 0 : 1;
    try {
        return worksheet.Shapes.AddPicture(imagePath, link, save, left, top, w, h);
    } catch (e1) {}
    try {
        var shp = worksheet.Pictures().Insert(imagePath);
        shp.Left = left;
        shp.Top = top;
        if (w > 0) {
            shp.Width = w;
        }
        if (h > 0) {
            shp.Height = h;
        }
        return shp;
    } catch (e2) {
        return null;
    }
}

/**
 * 转换 JS 数组到可用于 Shapes.Range 的数组。
 *
 * @param {Array} names 名称数组。
 * @returns {Array} 名称数组。
 */
function _shuBuildNamesArray(names) {
    var out = [];
    var i;
    for (i = 0; i < names.length; i++) {
        out.push(String(names[i]));
    }
    return out;
}
