/**
 * pivotChartUtils.js
 * WPS JSA 宏工具库：透视表与图表工具。
 */

/**
 * 创建透视表。
 *
 * @param {Object} sourceRange 源数据区域。
 * @param {Object} destinationRange 目标放置起始区域。
 * @param {string} [pivotName] 透视表名称。
 * @param {string} [cacheName] 缓存名（可选）。
 * @returns {Object|null} 透视表对象，失败返回 null。
 */
function createPivotTable(sourceRange, destinationRange, pivotName, cacheName) {
    if (!sourceRange || !destinationRange) {
        return null;
    }
    try {
        var wb = sourceRange.Worksheet.Parent;
        var xlDatabase = 1;
        var cache = wb.PivotCaches().Create(xlDatabase, sourceRange);
        if (cacheName) {
            try {
                cache.WorkbookConnection.Name = cacheName;
            } catch (eName) {}
        }
        var pName = pivotName || "PivotTable1";
        return cache.CreatePivotTable(destinationRange, pName);
    } catch (e) {
        return null;
    }
}

/**
 * 刷新透视表。
 *
 * @param {Object} pivotTable 透视表对象。
 * @returns {boolean} 是否刷新成功。
 */
function refreshPivotTable(pivotTable) {
    if (!pivotTable) {
        return false;
    }
    try {
        pivotTable.RefreshTable();
        return true;
    } catch (e1) {}
    try {
        pivotTable.PivotCache().Refresh();
        return true;
    } catch (e2) {
        return false;
    }
}

/**
 * 设置透视字段方向与位置。
 *
 * @param {Object} pivotTable 透视表对象。
 * @param {string} fieldName 字段名。
 * @param {number} orientation 方向常量（如 1 行、2 列、4 数据、3 页）。
 * @param {number} [position] 位置（1 基）。
 * @returns {boolean} 是否设置成功。
 */
function setPivotField(pivotTable, fieldName, orientation, position) {
    if (!pivotTable || !fieldName || !orientation) {
        return false;
    }
    try {
        var pf = pivotTable.PivotFields(fieldName);
        pf.Orientation = orientation;
        if (typeof position === "number" && position > 0) {
            pf.Position = position;
        }
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 创建图表对象。
 *
 * @param {Object} worksheet 工作表对象。
 * @param {Object} sourceRange 数据区域。
 * @param {number} [chartType] 图表类型常量。
 * @param {number} [left] 左位置。
 * @param {number} [top] 上位置。
 * @param {number} [width] 宽度。
 * @param {number} [height] 高度。
 * @returns {Object|null} 图表对象，失败返回 null。
 */
function createChart(worksheet, sourceRange, chartType, left, top, width, height) {
    if (!worksheet || !sourceRange) {
        return null;
    }
    var l = typeof left === "number" ? left : 100;
    var t = typeof top === "number" ? top : 80;
    var w = typeof width === "number" ? width : 520;
    var h = typeof height === "number" ? height : 320;
    try {
        var chartObj = worksheet.ChartObjects().Add(l, t, w, h);
        chartObj.Chart.SetSourceData(sourceRange);
        if (typeof chartType === "number") {
            chartObj.Chart.ChartType = chartType;
        }
        return chartObj;
    } catch (e) {
        return null;
    }
}

/**
 * 设置图表类型。
 *
 * @param {Object} chartObjectOrChart ChartObject 或 Chart。
 * @param {number} chartType 图表类型常量。
 * @returns {boolean} 是否设置成功。
 */
function setChartType(chartObjectOrChart, chartType) {
    var chart = _pcuGetChart(chartObjectOrChart);
    if (!chart || typeof chartType !== "number") {
        return false;
    }
    try {
        chart.ChartType = chartType;
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 设置图表标题。
 *
 * @param {Object} chartObjectOrChart ChartObject 或 Chart。
 * @param {string} title 图表标题。
 * @returns {boolean} 是否设置成功。
 */
function setChartTitle(chartObjectOrChart, title) {
    var chart = _pcuGetChart(chartObjectOrChart);
    if (!chart) {
        return false;
    }
    try {
        chart.HasTitle = true;
        chart.ChartTitle.Text = title || "";
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * 获取 Chart 对象。
 *
 * @param {Object} chartObjectOrChart ChartObject 或 Chart。
 * @returns {Object|null} Chart 对象。
 */
function _pcuGetChart(chartObjectOrChart) {
    if (!chartObjectOrChart) {
        return null;
    }
    try {
        if (chartObjectOrChart.Chart) {
            return chartObjectOrChart.Chart;
        }
    } catch (e1) {}
    try {
        if (chartObjectOrChart.ChartTitle !== undefined) {
            return chartObjectOrChart;
        }
    } catch (e2) {}
    return null;
}
