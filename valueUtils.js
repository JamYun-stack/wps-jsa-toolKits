/**
 * 判断值是否为空白。
 *
 * @param {*} value 要判断的值。
 * @returns {boolean} `null`、`undefined`、空字符串或仅包含空白字符的字符串返回 `true`，否则返回 `false`。
 */
function isBlank(value) {
    if (value === null || value === undefined) {
        return true;
    }

    if (typeof value === "string") {
        return _vuTrim(value) === "";
    }

    return false;
}

/**
 * 把值转换为数字。
 *
 * @param {*} value 要转换的值。
 * @param {number} [defaultValue] 转换失败时返回的默认值；默认是 `0`。
 * @returns {number} 转换成功时返回数字，失败时返回默认值。
 */
function toNumber(value, defaultValue) {
    var fallback = defaultValue === undefined ? 0 : Number(defaultValue);

    if (!_vuIsFiniteNumber(fallback)) {
        fallback = 0;
    }

    if (isBlank(value)) {
        return fallback;
    }

    var numberValue = Number(value);
    return _vuIsFiniteNumber(numberValue) ? numberValue : fallback;
}

/**
 * 把值安全转换为字符串。
 *
 * @param {*} value 要转换的值。
 * @param {string} [defaultValue] 值为空时返回的默认字符串；默认是空字符串。
 * @returns {string} 转换后的字符串。
 */
function toStringSafe(value, defaultValue) {
    if (value === null || value === undefined) {
        return defaultValue === undefined ? "" : String(defaultValue);
    }

    return String(value);
}

/**
 * 把值转换为布尔值。
 *
 * @param {*} value 要转换的值。
 * @param {boolean} [defaultValue] 转换失败时返回的默认值；默认是 `false`。
 * @returns {boolean} 转换后的布尔值。
 */
function toBoolean(value, defaultValue) {
    var fallback = defaultValue === true;

    if (typeof value === "boolean") {
        return value;
    }

    if (value === null || value === undefined) {
        return fallback;
    }

    if (typeof value === "number") {
        return value !== 0;
    }

    var text = _vuTrim(String(value)).toLowerCase();

    if (text === "") {
        return fallback;
    }

    if (text === "true" || text === "1" || text === "yes" || text === "y" || text === "on") {
        return true;
    }

    if (text === "false" || text === "0" || text === "no" || text === "n" || text === "off") {
        return false;
    }

    return fallback;
}

/**
 * 按项目中常用规则把值转成数字。
 *
 * @param {*} value 要转换的值。
 * @returns {number} 可转成数字时返回对应数字，否则返回 `0`。
 */
function toVal(value) {
    return toNumber(value, 0);
}

/**
 * 判断值是否可以视为 WPS/Excel 日期序列值。
 *
 * @param {*} value 要判断的值。
 * @returns {boolean} 可转换为有限数字时返回 `true`，否则返回 `false`。
 */
function isDateSerial(value) {
    if (value === null || value === undefined || value === "") {
        return false;
    }

    return _vuIsFiniteNumber(Number(value));
}

/**
 * 把 WPS/Excel 1900 日期系统序列值转换为 JavaScript Date。
 *
 * @param {number|string} serial 日期序列值，允许带小数部分表示时分秒。
 * @returns {Date|null} 转换成功时返回 `Date` 对象，失败时返回 `null`。
 */
function serialToDate(serial) {
    var serialValue = Number(serial);

    if (!_vuIsFiniteNumber(serialValue)) {
        return null;
    }

    var dayValue = Math.floor(serialValue);
    var timeFraction = serialValue - dayValue;

    if (timeFraction < 0) {
        timeFraction = 0;
    }

    if (dayValue > 59) {
        dayValue = dayValue - 1;
    }

    var baseDate = new Date(1899, 11, 31, 0, 0, 0, 0);
    var result = new Date(baseDate.getTime() + dayValue * 86400000 + Math.round(timeFraction * 86400000));

    if (isNaN(result.getTime())) {
        return null;
    }

    return result;
}

/**
 * 把日期值转换为 WPS/Excel 1900 日期系统序列值。
 *
 * @param {*} value `Date`、日期序列值或常见日期字符串。
 * @returns {number} 转换成功时返回日期序列值，失败时返回 `NaN`。
 */
function dateToSerial(value) {
    var dateValue = parseDateInput(value);

    if (!dateValue) {
        return NaN;
    }

    var baseDate = new Date(1899, 11, 31, 0, 0, 0, 0);
    var diffMilliseconds = dateValue.getTime() - baseDate.getTime();
    var serialValue = diffMilliseconds / 86400000;

    if (dateValue.getTime() >= new Date(1900, 2, 1, 0, 0, 0, 0).getTime()) {
        serialValue = serialValue + 1;
    }

    return Number(serialValue.toFixed(10));
}

/**
 * 把输入值解析为 JavaScript Date。
 * 支持 `Date`、WPS/Excel 日期序列值，以及常见日期字符串格式。
 *
 * @param {*} value 要解析的值。
 * @returns {Date|null} 解析成功时返回 `Date` 对象，失败时返回 `null`。
 */
function parseDateInput(value) {
    if (value === null || value === undefined || value === "") {
        return null;
    }

    if (value instanceof Date) {
        if (isNaN(value.getTime())) {
            return null;
        }

        return new Date(value.getTime());
    }

    if (typeof value === "number") {
        return serialToDate(value);
    }

    if (typeof value === "string") {
        var text = _vuTrim(value);

        if (text === "") {
            return null;
        }

        var parsedFromText = _vuParseStringDate(text);
        if (parsedFromText) {
            return parsedFromText;
        }

        if (/^-?\d+(\.\d+)?$/.test(text)) {
            return serialToDate(Number(text));
        }
    }

    return null;
}

/**
 * 把日期值格式化为日期字符串。
 *
 * @param {*} value `Date`、日期序列值或常见日期字符串。
 * @param {string} [pattern] 输出格式，例如 `"yyyy-MM-dd"`；默认是 `"yyyy-MM-dd"`。
 * @returns {string} 格式化后的字符串；解析失败时返回空字符串。
 */
function formatDate(value, pattern) {
    var dateValue = parseDateInput(value);

    if (!dateValue) {
        return "";
    }

    return _vuFormatDateByPattern(dateValue, pattern || "yyyy-MM-dd");
}

/**
 * 把日期值格式化为日期时间字符串。
 *
 * @param {*} value `Date`、日期序列值或常见日期字符串。
 * @param {string} [pattern] 输出格式，例如 `"yyyy-MM-dd HH:mm:ss"`；默认是 `"yyyy-MM-dd HH:mm:ss"`。
 * @returns {string} 格式化后的字符串；解析失败时返回空字符串。
 */
function formatDateTime(value, pattern) {
    var dateValue = parseDateInput(value);

    if (!dateValue) {
        return "";
    }

    return _vuFormatDateByPattern(dateValue, pattern || "yyyy-MM-dd HH:mm:ss");
}

/**
 * 获取指定日期所在天的开始时刻。
 *
 * @param {*} value `Date`、日期序列值或常见日期字符串。
 * @returns {Date|null} 当天 `00:00:00` 的时间对象；解析失败时返回 `null`。
 */
function startOfDay(value) {
    var dateValue = parseDateInput(value);

    if (!dateValue) {
        return null;
    }

    return new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate(), 0, 0, 0, 0);
}

/**
 * 获取指定日期所在天的结束时刻。
 *
 * @param {*} value `Date`、日期序列值或常见日期字符串。
 * @returns {Date|null} 当天 `23:59:59` 的时间对象；解析失败时返回 `null`。
 */
function endOfDay(value) {
    var dateValue = parseDateInput(value);

    if (!dateValue) {
        return null;
    }

    return new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate(), 23, 59, 59, 999);
}

/**
 * 获取指定日期所在月的第一天开始时刻。
 *
 * @param {*} value `Date`、日期序列值或常见日期字符串。
 * @returns {Date|null} 当月第一天 `00:00:00` 的时间对象；解析失败时返回 `null`。
 */
function startOfMonth(value) {
    var dateValue = parseDateInput(value);

    if (!dateValue) {
        return null;
    }

    return new Date(dateValue.getFullYear(), dateValue.getMonth(), 1, 0, 0, 0, 0);
}

/**
 * 获取指定日期所在月的最后一天结束时刻。
 *
 * @param {*} value `Date`、日期序列值或常见日期字符串。
 * @returns {Date|null} 当月最后一天 `23:59:59` 的时间对象；解析失败时返回 `null`。
 */
function endOfMonth(value) {
    var dateValue = parseDateInput(value);

    if (!dateValue) {
        return null;
    }

    return new Date(dateValue.getFullYear(), dateValue.getMonth() + 1, 0, 23, 59, 59, 999);
}

/**
 * 计算两个日期之间相差的天数。
 *
 * @param {*} startValue 开始日期值。
 * @param {*} endValue 结束日期值。
 * @returns {number} 相差的天数；任一值无法解析时返回 `NaN`。
 */
function diffDays(startValue, endValue) {
    var startDate = parseDateInput(startValue);
    var endDate = parseDateInput(endValue);

    if (!startDate || !endDate) {
        return NaN;
    }

    return Number(((endDate.getTime() - startDate.getTime()) / 86400000).toFixed(10));
}

/**
 * 计算两个日期之间相差的秒数。
 *
 * @param {*} startValue 开始日期值。
 * @param {*} endValue 结束日期值。
 * @returns {number} 相差的秒数；任一值无法解析时返回 `NaN`。
 */
function diffSeconds(startValue, endValue) {
    var startDate = parseDateInput(startValue);
    var endDate = parseDateInput(endValue);

    if (!startDate || !endDate) {
        return NaN;
    }

    return Number(((endDate.getTime() - startDate.getTime()) / 1000).toFixed(3));
}

/**
 * 判断数字是否为有限值。
 *
 * @private
 * @param {*} value 要判断的值。
 * @returns {boolean} 是有限数字时返回 `true`，否则返回 `false`。
 */
function _vuIsFiniteNumber(value) {
    return typeof value === "number" && isFinite(value);
}

/**
 * 去掉字符串首尾空白。
 *
 * @private
 * @param {*} value 要处理的值。
 * @returns {string} 去空白后的字符串。
 */
function _vuTrim(value) {
    return String(value).replace(/^\s+|\s+$/g, "");
}

/**
 * 按常见格式解析字符串日期。
 *
 * @private
 * @param {string} text 日期字符串。
 * @returns {Date|null} 解析成功返回日期对象，否则返回 `null`。
 */
function _vuParseStringDate(text) {
    var normalizedText = String(text).replace("T", " ");
    var match = null;

    match = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:\s+(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?)?$/.exec(normalizedText);
    if (match) {
        return _vuBuildDate(match[1], match[2], match[3], match[4], match[5], match[6]);
    }

    match = /^(\d{4})(\d{2})(\d{2})(?:\s*(\d{2})(\d{2})?(\d{2})?)?$/.exec(normalizedText);
    if (match) {
        return _vuBuildDate(match[1], match[2], match[3], match[4], match[5], match[6]);
    }

    match = /^(\d{4})年(\d{1,2})月(\d{1,2})日(?:\s+(\d{1,2})时?(?::|：)?(\d{1,2})?分?(?::|：)?(\d{1,2})?秒?)?$/.exec(normalizedText);
    if (match) {
        return _vuBuildDate(match[1], match[2], match[3], match[4], match[5], match[6]);
    }

    var timestamp = Date.parse(normalizedText);
    if (!isNaN(timestamp)) {
        return new Date(timestamp);
    }

    return null;
}

/**
 * 构造日期对象。
 *
 * @private
 * @param {*} year 年。
 * @param {*} month 月。
 * @param {*} day 日。
 * @param {*} hour 时。
 * @param {*} minute 分。
 * @param {*} second 秒。
 * @returns {Date|null} 构造成功返回日期对象，否则返回 `null`。
 */
function _vuBuildDate(year, month, day, hour, minute, second) {
    var y = Number(year);
    var m = Number(month);
    var d = Number(day);
    var h = hour === undefined || hour === "" ? 0 : Number(hour);
    var mi = minute === undefined || minute === "" ? 0 : Number(minute);
    var s = second === undefined || second === "" ? 0 : Number(second);

    if (!_vuIsFiniteNumber(y) || !_vuIsFiniteNumber(m) || !_vuIsFiniteNumber(d)) {
        return null;
    }

    if (!_vuIsFiniteNumber(h)) {
        h = 0;
    }

    if (!_vuIsFiniteNumber(mi)) {
        mi = 0;
    }

    if (!_vuIsFiniteNumber(s)) {
        s = 0;
    }

    var result = new Date(y, m - 1, d, h, mi, s, 0);

    if (isNaN(result.getTime())) {
        return null;
    }

    return result;
}

/**
 * 按模式格式化日期。
 *
 * @private
 * @param {Date} dateValue 日期对象。
 * @param {string} pattern 格式模板。
 * @returns {string} 格式化后的字符串。
 */
function _vuFormatDateByPattern(dateValue, pattern) {
    var result = String(pattern || "yyyy-MM-dd");
    var year = String(dateValue.getFullYear());
    var month = _vuPadNumber(dateValue.getMonth() + 1, 2);
    var day = _vuPadNumber(dateValue.getDate(), 2);
    var hour = _vuPadNumber(dateValue.getHours(), 2);
    var minute = _vuPadNumber(dateValue.getMinutes(), 2);
    var second = _vuPadNumber(dateValue.getSeconds(), 2);

    result = result.replace(/yyyy/g, year);
    result = result.replace(/MM/g, month);
    result = result.replace(/dd/g, day);
    result = result.replace(/HH/g, hour);
    result = result.replace(/mm/g, minute);
    result = result.replace(/ss/g, second);

    return result;
}

/**
 * 左侧补零。
 *
 * @private
 * @param {number} value 原始数字。
 * @param {number} size 目标位数。
 * @returns {string} 补零后的字符串。
 */
function _vuPadNumber(value, size) {
    var text = String(Math.abs(parseInt(value, 10)));

    while (text.length < size) {
        text = "0" + text;
    }

    return text;
}
