/**
 * 安全读取对象中的嵌套属性。
 * `path` 支持 `"a.b.c"`、`["a", "b", "c"]`，也支持 `safeGet(obj, "a", "b", "c")` 这种多段调用方式。
 *
 * @param {*} obj 要读取的对象。
 * @param {string|string[]} path 路径字符串或路径数组。
 * @param {*} [defaultValue] 路径不存在时返回的默认值。
 * @returns {*} 路径存在时返回对应值，否则返回默认值。
 */
function safeGet(obj, path, defaultValue) {
    var keys = [];
    var fallback = defaultValue;
    var i = 0;

    if (arguments.length > 3) {
        for (i = 1; i < arguments.length; i++) {
            keys.push(arguments[i]);
        }
        fallback = undefined;
    } else {
        keys = _ouToPathArray(path);
    }

    var current = obj;

    for (i = 0; i < keys.length; i++) {
        if (current === null || current === undefined) {
            return fallback;
        }

        if (typeof current !== "object" && typeof current !== "function") {
            return fallback;
        }

        if (!(keys[i] in current)) {
            return fallback;
        }

        current = current[keys[i]];
    }

    return current === undefined ? fallback : current;
}

/**
 * 安全写入对象中的嵌套属性。
 *
 * @param {Object} obj 要写入的对象。
 * @param {string|string[]} path 路径字符串或路径数组。
 * @param {*} value 要设置的值。
 * @param {boolean} [createMissing] 路径中不存在的中间节点是否自动创建；默认是 `true`。
 * @returns {boolean} 写入成功时返回 `true`，否则返回 `false`。
 */
function safeSet(obj, path, value, createMissing) {
    if (!obj || (typeof obj !== "object" && typeof obj !== "function")) {
        return false;
    }

    var keys = _ouToPathArray(path);
    if (keys.length === 0) {
        return false;
    }

    var shouldCreate = createMissing !== false;
    var current = obj;
    var i = 0;

    for (i = 0; i < keys.length - 1; i++) {
        if (current[keys[i]] === undefined || current[keys[i]] === null) {
            if (!shouldCreate) {
                return false;
            }

            current[keys[i]] = {};
        }

        if (typeof current[keys[i]] !== "object" && typeof current[keys[i]] !== "function") {
            if (!shouldCreate) {
                return false;
            }

            current[keys[i]] = {};
        }

        current = current[keys[i]];
    }

    current[keys[keys.length - 1]] = value;
    return true;
}

/**
 * 判断对象中是否存在指定路径。
 *
 * @param {*} obj 要检查的对象。
 * @param {string|string[]} path 路径字符串或路径数组。
 * @returns {boolean} 路径存在时返回 `true`，否则返回 `false`。
 */
function hasPath(obj, path) {
    var keys = _ouToPathArray(path);
    var current = obj;
    var i = 0;

    if (keys.length === 0) {
        return false;
    }

    for (i = 0; i < keys.length; i++) {
        if (current === null || current === undefined) {
            return false;
        }

        if (typeof current !== "object" && typeof current !== "function") {
            return false;
        }

        if (!(keys[i] in current)) {
            return false;
        }

        current = current[keys[i]];
    }

    return true;
}

/**
 * 从对象中挑选指定键，返回新对象。
 *
 * @param {Object} obj 原始对象。
 * @param {string|string[]} keys 要保留的键或键数组。
 * @returns {Object} 只包含指定键的新对象。
 */
function pickKeys(obj, keys) {
    var result = {};
    var keyList = _ouToSimpleArray(keys);
    var i = 0;

    if (!obj || typeof obj !== "object") {
        return result;
    }

    for (i = 0; i < keyList.length; i++) {
        if (keyList[i] in obj) {
            result[keyList[i]] = obj[keyList[i]];
        }
    }

    return result;
}

/**
 * 从对象中排除指定键，返回新对象。
 *
 * @param {Object} obj 原始对象。
 * @param {string|string[]} keys 要排除的键或键数组。
 * @returns {Object} 去掉指定键后的新对象。
 */
function omitKeys(obj, keys) {
    var result = {};
    var omitMap = {};
    var keyList = _ouToSimpleArray(keys);
    var key = "";
    var i = 0;

    if (!obj || typeof obj !== "object") {
        return result;
    }

    for (i = 0; i < keyList.length; i++) {
        omitMap[keyList[i]] = true;
    }

    for (key in obj) {
        if (obj.hasOwnProperty(key) && !omitMap[key]) {
            result[key] = obj[key];
        }
    }

    return result;
}

/**
 * 对对象或数组做浅拷贝。
 *
 * @param {*} value 要拷贝的值。
 * @returns {*} 对象返回浅拷贝对象，数组返回浅拷贝数组，其它值直接原样返回。
 */
function shallowClone(value) {
    var key = "";
    var result = null;

    if (value instanceof Array) {
        return value.slice(0);
    }

    if (!value || typeof value !== "object") {
        return value;
    }

    if (value instanceof Date) {
        return new Date(value.getTime());
    }

    result = {};
    for (key in value) {
        if (value.hasOwnProperty(key)) {
            result[key] = value[key];
        }
    }

    return result;
}

/**
 * 对数组做拷贝。
 *
 * @param {*} arrayValue 要拷贝的数组。
 * @param {boolean} [deep] 是否深拷贝数组元素；为 `true` 时会递归拷贝元素。
 * @returns {Array} 拷贝后的新数组；传入的值不是数组时返回空数组。
 */
function cloneArray(arrayValue, deep) {
    var result = [];
    var i = 0;

    if (!(arrayValue instanceof Array)) {
        return result;
    }

    for (i = 0; i < arrayValue.length; i++) {
        result.push(deep === true ? deepClone(arrayValue[i]) : arrayValue[i]);
    }

    return result;
}

/**
 * 对普通对象、数组和 Date 做深拷贝。
 * 不处理循环引用，也不会拷贝 WPS 对象、ActiveX 对象等特殊运行时对象。
 *
 * @param {*} value 要拷贝的值。
 * @returns {*} 深拷贝后的值。
 */
function deepClone(value) {
    var result = null;
    var key = "";
    var i = 0;

    if (value === null || value === undefined) {
        return value;
    }

    if (value instanceof Date) {
        return new Date(value.getTime());
    }

    if (value instanceof Array) {
        result = [];
        for (i = 0; i < value.length; i++) {
            result.push(deepClone(value[i]));
        }
        return result;
    }

    if (_ouIsPlainObject(value)) {
        result = {};
        for (key in value) {
            if (value.hasOwnProperty(key)) {
                result[key] = deepClone(value[key]);
            }
        }
        return result;
    }

    return value;
}

/**
 * 把路径值统一转换为路径数组。
 *
 * @private
 * @param {string|string[]} path 路径字符串或路径数组。
 * @returns {string[]} 路径数组。
 */
function _ouToPathArray(path) {
    var result = [];
    var i = 0;

    if (path instanceof Array) {
        for (i = 0; i < path.length; i++) {
            if (path[i] !== "" && path[i] !== null && path[i] !== undefined) {
                result.push(String(path[i]));
            }
        }
        return result;
    }

    if (path === null || path === undefined || path === "") {
        return result;
    }

    var text = String(path);
    var parts = text.split(".");

    for (i = 0; i < parts.length; i++) {
        if (parts[i] !== "") {
            result.push(parts[i]);
        }
    }

    return result;
}

/**
 * 把单值或数组统一转成简单数组。
 *
 * @private
 * @param {string|string[]} value 单值或数组。
 * @returns {string[]} 统一后的数组。
 */
function _ouToSimpleArray(value) {
    var result = [];
    var i = 0;

    if (value instanceof Array) {
        for (i = 0; i < value.length; i++) {
            result.push(String(value[i]));
        }
        return result;
    }

    if (value === null || value === undefined || value === "") {
        return result;
    }

    result.push(String(value));
    return result;
}

/**
 * 判断值是否为普通对象。
 *
 * @private
 * @param {*} value 要判断的值。
 * @returns {boolean} 是普通对象时返回 `true`，否则返回 `false`。
 */
function _ouIsPlainObject(value) {
    return !!value && Object.prototype.toString.call(value) === "[object Object]";
}
