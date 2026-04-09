# runtimeUtils.js 运行时辅助模块

`runtimeUtils.js` 用于封装 WPS JSA 宏运行时里常见的对话框、调试输出和 UI 让出操作。

## 使用约定

- 优先使用 WPS JSA 宏里已有的 `MsgBox`、`InputBox`、`Debug.Print`、`DoEvents`。
- 这些函数主要解决“统一调用方式”和“失败时不直接让宏崩掉”的问题。

## 公开函数

- `showMessage(message, title, buttons, icon)`
- `confirmDialog(message, title, defaultValue)`
- `promptText(message, title, defaultValue)`
- `debugPrint(message)`
- `yieldUi()`

### showMessage(message, title, buttons, icon)

作用：显示消息框。

参数：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `message` | `string` | 消息内容。 |
| `title` | `string` | 对话框标题。 |
| `buttons` | `string` 或 `number` | 按钮类型，例如 `"ok"`、`"yesno"`。 |
| `icon` | `string` 或 `number` | 图标类型，例如 `"info"`、`"warn"`、`"error"`、`"question"`。 |

返回值：`number` 或 `string`。通常返回 `MsgBox` 的按钮结果。

示例代码：

```js
showMessage("处理完成", "提示", "ok", "info");
```

示例代码完成的目的：在宏执行结束后给出统一的提示消息。

### confirmDialog(message, title, defaultValue)

作用：显示确认对话框。

参数：`message` 为提示内容；`title` 为标题；`defaultValue` 为对话框不可用时的默认结果。

返回值：`boolean`。

示例代码：

```js
if (confirmDialog("是否继续导出？", "确认", false)) {
    // continue
}
```

示例代码完成的目的：在执行覆盖、删除等操作前征求用户确认。

### promptText(message, title, defaultValue)

作用：显示文本输入框并返回用户输入。

参数：`message` 为提示内容；`title` 为标题；`defaultValue` 为默认文本。

返回值：`string`。

示例代码：

```js
var shopName = promptText("请输入店铺名称", "输入", "");
```

示例代码完成的目的：在宏执行中临时收集简单文本参数。

### debugPrint(message)

作用：输出调试信息。

参数：`message` 为要输出的内容。

返回值：`boolean`。

示例代码：

```js
debugPrint("开始处理第 1 个工作表");
```

示例代码完成的目的：在宏调试阶段输出关键步骤日志。

### yieldUi()

作用：让出 UI 线程，给 WPS 界面处理机会。

参数：无。

返回值：`boolean`。

示例代码：

```js
for (var i = 0; i < rows.length; i++) {
    // 处理逻辑
    if (i % 100 === 0) {
        yieldUi();
    }
}
```

示例代码完成的目的：长循环中适当让出 UI，避免界面长时间无响应。

## 依赖说明

- 依赖 WPS JSA 宏运行时里的 `MsgBox`、`InputBox`、`Debug.Print`、`DoEvents`。
- 在 Node 静态检查环境中无法真实执行这些函数，但语法可正常校验。
