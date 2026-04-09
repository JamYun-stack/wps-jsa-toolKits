/**
 * 拆分分类字符串。
 * 默认把第一个 `-` 左侧视为组名，右侧开头的连续数字视为排序号，其余部分视为名称。
 *
 * @param {string} category 原始分类字符串。
 * @param {string} [separator] 分类分隔符；默认是 `"-"`。
 * @returns {{raw: string, group: string, order: number|string, name: string}} 分类拆分结果对象。
 *
 * 返回格式：
 * {
 *   raw: "零食-12薯片",
 *   group: "零食",
 *   order: 12,
 *   name: "薯片"
 * }
 */
function splitCategory(category, separator) {
    var text = category === null || category === undefined ? "" : String(category);
    var splitText = separator === undefined ? "-" : String(separator);
    var index = text.indexOf(splitText);
    var result = {
        raw: text,
        group: "",
        order: "",
        name: text
    };
    var rest = "";
    var match = null;

    if (index <= 0) {
        return result;
    }

    result.group = text.substring(0, index);
    rest = text.substring(index + splitText.length);
    match = /^(\d+)(.*)$/.exec(rest);

    if (match) {
        result.order = Number(match[1]);
        result.name = match[2];
    } else {
        result.name = rest;
    }

    return result;
}

/**
 * 根据组名、排序号和名称拼接分类字符串。
 *
 * @param {string} group 分类组名。
 * @param {number|string} order 排序号；为空时不拼接数字部分。
 * @param {string} name 分类名称。
 * @param {string} [separator] 分类分隔符；默认是 `"-"`。
 * @returns {string} 拼接后的分类字符串。
 */
function joinCategoryParts(group, order, name, separator) {
    var splitText = separator === undefined ? "-" : String(separator);
    var groupText = group === null || group === undefined ? "" : String(group);
    var orderText = order === null || order === undefined || order === "" ? "" : String(order);
    var nameText = name === null || name === undefined ? "" : String(name);

    if (groupText === "") {
        return orderText + nameText;
    }

    return groupText + splitText + orderText + nameText;
}

/**
 * 生成用于排序和比较的分类键。
 *
 * @param {string|Object} category 分类字符串，或 `splitCategory` 的返回对象。
 * @param {string} [separator] 分类分隔符；默认是 `"-"`。
 * @returns {string} 统一格式的分类键字符串。
 */
function buildCategoryKey(category, separator) {
    var info = _cuNormalizeCategoryInfo(category, separator);
    var orderText = info.order === "" ? "" : _cuPadNumber(info.order, 8);

    return info.group + "||" + orderText + "||" + info.name;
}

/**
 * 比较两个分类值的顺序。
 *
 * @param {string|Object} left 左侧分类值。
 * @param {string|Object} right 右侧分类值。
 * @param {string} [separator] 分类分隔符；默认是 `"-"`。
 * @returns {number} 左侧小于右侧返回 `-1`，相等返回 `0`，大于返回 `1`。
 */
function compareCategoryKey(left, right, separator) {
    var leftInfo = _cuNormalizeCategoryInfo(left, separator);
    var rightInfo = _cuNormalizeCategoryInfo(right, separator);

    if (leftInfo.group < rightInfo.group) {
        return -1;
    }

    if (leftInfo.group > rightInfo.group) {
        return 1;
    }

    if (leftInfo.order === "" && rightInfo.order !== "") {
        return -1;
    }

    if (leftInfo.order !== "" && rightInfo.order === "") {
        return 1;
    }

    if (leftInfo.order < rightInfo.order) {
        return -1;
    }

    if (leftInfo.order > rightInfo.order) {
        return 1;
    }

    if (leftInfo.name < rightInfo.name) {
        return -1;
    }

    if (leftInfo.name > rightInfo.name) {
        return 1;
    }

    return 0;
}

/**
 * 对分类数组或包含分类字段的对象数组排序。
 *
 * @param {Array} items 原始数组。
 * @param {string|Function} [getCategory] 获取分类值的方式；不传时默认数组元素本身就是分类字符串。
 * @param {string} [separator] 分类分隔符；默认是 `"-"`。
 * @param {boolean} [desc] 是否倒序；为 `true` 时倒序。
 * @returns {Array} 排序后的新数组；不会修改原数组。
 */
function sortCategoryItems(items, getCategory, separator, desc) {
    var result = items instanceof Array ? items.slice(0) : [];
    var splitText = separator === undefined ? "-" : String(separator);

    result.sort(function (left, right) {
        var leftCategory = _cuGetCategoryValue(left, getCategory);
        var rightCategory = _cuGetCategoryValue(right, getCategory);
        var compared = compareCategoryKey(leftCategory, rightCategory, splitText);

        return desc === true ? compared * -1 : compared;
    });

    return result;
}

/**
 * 获取分类值。
 *
 * @private
 * @param {*} item 当前元素。
 * @param {string|Function} getter 取值规则。
 * @returns {*} 分类值。
 */
function _cuGetCategoryValue(item, getter) {
    if (typeof getter === "function") {
        return getter(item);
    }

    if (typeof getter === "string" && getter !== "") {
        if (item && typeof item === "object" && getter in item) {
            return item[getter];
        }

        return "";
    }

    return item;
}

/**
 * 统一分类信息对象。
 *
 * @private
 * @param {string|Object} category 分类字符串或分类对象。
 * @param {string} [separator] 分隔符。
 * @returns {{raw: string, group: string, order: number|string, name: string}} 统一后的分类信息对象。
 */
function _cuNormalizeCategoryInfo(category, separator) {
    if (category && typeof category === "object" && "group" in category && "name" in category) {
        return {
            raw: category.raw === undefined ? joinCategoryParts(category.group, category.order, category.name, separator) : String(category.raw),
            group: category.group === undefined || category.group === null ? "" : String(category.group),
            order: category.order === undefined || category.order === null ? "" : category.order,
            name: category.name === undefined || category.name === null ? "" : String(category.name)
        };
    }

    return splitCategory(category, separator);
}

/**
 * 左侧补零。
 *
 * @private
 * @param {number|string} value 原始数字。
 * @param {number} size 目标位数。
 * @returns {string} 补零后的字符串。
 */
function _cuPadNumber(value, size) {
    var numberValue = parseInt(value, 10);
    var text = isNaN(numberValue) ? "" : String(Math.abs(numberValue));

    while (text.length < size) {
        text = "0" + text;
    }

    return text;
}
