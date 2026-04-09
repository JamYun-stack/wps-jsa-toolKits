/**
 * 显示消息框。
 *
 * @param {string} message 消息内容。
 * @param {string} [title] 对话框标题。
 * @param {string|number} [buttons] 按钮类型；可传 `"ok"`、`"yesno"` 或直接传数值样式。
 * @param {string|number} [icon] 图标类型；可传 `"info"`、`"warn"`、`"error"`、`"question"` 或直接传数值样式。
 * @returns {number|string} `MsgBox` 返回值；无法显示时返回空字符串。
 */
function showMessage(message, title, buttons, icon) {
    try {
        var style = _ruResolveButtons(buttons) + _ruResolveIcon(icon);
        return MsgBox(String(message === null || message === undefined ? "" : message), style, String(title || ""));
    } catch (error) {
        return "";
    }
}

/**
 * 显示确认对话框。
 *
 * @param {string} message 提示内容。
 * @param {string} [title] 对话框标题。
 * @param {boolean} [defaultValue] 对话框不可用时返回的默认值；默认是 `false`。
 * @returns {boolean} 用户点击“是”返回 `true`，否则返回 `false`。
 */
function confirmDialog(message, title, defaultValue) {
    try {
        var result = MsgBox(
            String(message === null || message === undefined ? "" : message),
            _ruResolveButtons("yesno") + _ruResolveIcon("question"),
            String(title || "")
        );
        return result === 6 || result === -1;
    } catch (error) {
        return defaultValue === true;
    }
}

/**
 * 显示文本输入框。
 *
 * @param {string} message 提示内容。
 * @param {string} [title] 对话框标题。
 * @param {string} [defaultValue] 默认输入值。
 * @returns {string} 用户输入的文本；取消或失败时返回空字符串或默认值。
 */
function promptText(message, title, defaultValue) {
    try {
        return String(InputBox(
            String(message === null || message === undefined ? "" : message),
            String(title || ""),
            String(defaultValue === undefined ? "" : defaultValue)
        ));
    } catch (error) {
        return defaultValue === undefined ? "" : String(defaultValue);
    }
}

/**
 * 输出调试信息。
 *
 * @param {*} message 要输出的内容。
 * @returns {boolean} 输出成功返回 `true`，否则返回 `false`。
 */
function debugPrint(message) {
    try {
        if (typeof Debug !== "undefined" && Debug && typeof Debug.Print === "function") {
            Debug.Print(String(message === null || message === undefined ? "" : message));
            return true;
        }
    } catch (error) {
    }

    try {
        if (typeof console !== "undefined" && console && typeof console.log === "function") {
            console.log(String(message === null || message === undefined ? "" : message));
            return true;
        }
    } catch (fallbackError) {
    }

    return false;
}

/**
 * 让出 UI 线程，给 WPS 界面处理机会。
 *
 * @returns {boolean} 调用成功返回 `true`，否则返回 `false`。
 */
function yieldUi() {
    try {
        DoEvents();
        return true;
    } catch (error) {
        return false;
    }
}

/**
 * 解析按钮样式。
 *
 * @private
 * @param {string|number} buttons 按钮类型。
 * @returns {number} 按钮样式数字。
 */
function _ruResolveButtons(buttons) {
    if (typeof buttons === "number") {
        return buttons;
    }

    if (buttons === "yesno") {
        return 4;
    }

    return 0;
}

/**
 * 解析图标样式。
 *
 * @private
 * @param {string|number} icon 图标类型。
 * @returns {number} 图标样式数字。
 */
function _ruResolveIcon(icon) {
    if (typeof icon === "number") {
        return icon;
    }

    if (icon === "question") {
        return 32;
    }

    if (icon === "warn") {
        return 48;
    }

    if (icon === "info") {
        return 64;
    }

    if (icon === "error") {
        return 16;
    }

    return 0;
}
