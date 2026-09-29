function 同步油耗到汇总() {
    var srcWb = ActiveWorkbook;
    var dstPath = "C:/Users/Jubin/Desktop/源数据/油耗分析汇总.xlsm";
    var srcSheets = ["干线油耗", "华北油耗", "华东油耗", "华南油耗"];

    // 工具:是不是有效数字
    function isNum(v) {
        return typeof v === "number" && !isNaN(v);
    }

    // 值归一化: 用于去重比对(数字统一精度, 字符串去首尾空白)
    function norm(v) {
        if (v === null || v === undefined) return "";
        if (typeof v === "number") return String(Math.round(v * 1e6) / 1e6);
        return String(v).replace(/\s+/g, " ").trim();
    }

    // 数字键: 兼容数字/字符串数字, 统一 4 位小数精度
    function numKey(v) {
        if (v === null || v === undefined || v === "") return "";
        if (typeof v === "number") return String(Math.round(v * 1e4) / 1e4);
        var f = parseFloat(v);
        return isNaN(f) ? String(v).trim() : String(Math.round(f * 1e4) / 1e4);
    }

    // 日期键统一到"天": 序列数 -> YYYY-MM-DD; 字符串 -> 提取年月日
    // (写入目标后再读回, 日期的 类型/精度 会漂移, 必须按天比对)
    function dayKey(v) {
        if (v === null || v === undefined || v === "") return "";
        if (typeof v === "number") {
            var d = new Date(Date.UTC(1899, 11, 30));
            d.setUTCDate(d.getUTCDate() + Math.floor(v));
            return d.getUTCFullYear() + "-" + ("0" + (d.getUTCMonth() + 1)).slice(-2) + "-" + ("0" + d.getUTCDate()).slice(-2);
        }
        var m = String(v).match(/(\d{4})[-\/年](\d{1,2})[-\/月](\d{1,2})/);
        if (m) return m[1] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[3]).slice(-2);
        return String(v).trim();
    }

    // 确定性找列内最后一个非空行: 从 UsedRange 尾部向上分块扫描,
    // 不受筛选/隐藏行影响(End(xlUp) 在筛选后会定位到错误行导致覆盖数据)
    function lastDataRow(ws, col) {
        var ur = ws.UsedRange;
        var end = ur.Row + ur.Rows.Count - 1;
        if (end > ws.Rows.Count) end = ws.Rows.Count;
        while (end >= 1) {
            var start = Math.max(1, end - 1999);
            var vals = ws.Range(ws.Cells(start, col), ws.Cells(end, col)).Value2;
            var n = end - start + 1;
            for (var i = n - 1; i >= 0; i--) {
                var v = (n === 1) ? vals : vals[i][0];
                if (v !== null && v !== undefined && v !== "") return start + i;
            }
            end = start - 1;
        }
        return 1;
    }

    // 工具:该行是否是未加油/未加满油的占位结果
    function isNoFuelRow(ws, row) {
        for (var c = 4; c <= 10; c++) { // D~J:司机、加油及油耗数据
            var v = ws.Cells(row, c).Value2;
            if (typeof v !== "string") continue;
            if (v.indexOf("未加油") >= 0 || v.indexOf("未加满油") >= 0) return true;
        }
        return false;
    }

    // ---------- 1. 收集 4 张源表里达标率为数字的行 ----------
    var allRows = [];
    var sheetCounts = {};
    for (var i = 0; i < srcSheets.length; i++) {
        var name = srcSheets[i];
        var ws = srcWb.Sheets(name);
        if (!ws) continue;

        var lastRow = lastDataRow(ws, 2); // B 列确定性末行(源表筛选时也不漏行)
        if (lastRow < 7) continue;
        sheetCounts[name] = 0;

        for (var r = 7; r <= lastRow; r++) {
            var rate = ws.Cells(r, 14).Value2; // N 列 = 达标率
            if (!isNum(rate)) continue;
            if (isNoFuelRow(ws, r)) continue;

            var rowData = [];
            for (var c = 2; c <= 19; c++) { // B~S 共 18 列
                rowData.push(ws.Cells(r, c).Value2);
            }
            allRows.push(rowData);
            sheetCounts[name]++;
        }
    }

    if (allRows.length === 0) {
        return "无可同步数据:4 张源表的达标率列都没有数字结果";
    }

    // ---------- 2. 打开目标文件(已开着就直接复用) ----------
    var dstWb = null;
    for (var i = 1; i <= Workbooks.Count; i++) {
        try {
            if (Workbooks.Item(i).FullName === dstPath) {
                dstWb = Workbooks.Item(i);
                break;
            }
        } catch (e) { /* ignore */ }
    }
    if (!dstWb) dstWb = Workbooks.Open(dstPath);

    var dstWs = dstWb.Sheets("Sheet1");
    if (!dstWs) return "目标文件没有 Sheet1";

    // ---------- 2.5 目标若处于筛选状态, 先显示全部行(保留筛选条件) ----------
    // 否则追加会落在筛选视图"后面", 覆盖被隐藏的正常数据行
    var filterMsg = "";
    try {
        if (dstWs.FilterMode) { dstWs.ShowAllData(); filterMsg = "(已自动显示全部行)"; }
    } catch (e) { /* ignore */ }

    // ---------- 3. 确定性找 Sheet1 末尾(C 列锚) + 建已有行键集合用于去重 ----------
    var dstLastRow = lastDataRow(dstWs, 3);
    if (dstLastRow < 1) dstLastRow = 1;
    var startRow = dstLastRow + 1;

    function keyOf(ws, row, fromCol, toCol) {
        var parts = [];
        for (var c = fromCol; c <= toCol; c++) parts.push(norm(ws.Cells(row, c).Value2));
        return parts.join("|").replace(/\|+$/, "");
    }

    // 去重键(语义键, 抗"写入后读回"的类型漂移): 车牌 + 时间A(天) + 时间B(天) + 金额 + 升数
    // 源表布局: row[0]=B车牌 [3]=E时间A [4]=F时间B [5]=G金额 [6]=H升数
    function rowKeySrc(row) {
        return norm(row[0]) + "|" + dayKey(row[3]) + "|" + dayKey(row[4]) + "|" + numKey(row[5]) + "|" + numKey(row[6]);
    }
    // 目标行: C=3车牌 F=6时间A G=7时间B H=8金额 I=9升数
    function rowKeyDst(r) {
        return norm(dstWs.Cells(r, 3).Value2) + "|" + dayKey(dstWs.Cells(r, 6).Value2) + "|" + dayKey(dstWs.Cells(r, 7).Value2) + "|" + numKey(dstWs.Cells(r, 8).Value2) + "|" + numKey(dstWs.Cells(r, 9).Value2);
    }
    var existKeys = {};
    for (var r = 2; r <= dstLastRow; r++) {
        var k = rowKeyDst(r);
        if (k.replace(/\|/g, "")) existKeys[k] = true;
    }

    // ---------- 4. 写入 C~T, 逐行去重(重复只跳过该行, 后续照常), 补 A 列序号 ----------
    Application.ScreenUpdating = false;
    var writeRow = startRow, written = 0, skipped = 0;
    try {
        for (var i = 0; i < allRows.length; i++) {
            var row = allRows[i];
            var key = rowKeySrc(row);
            if (key.replace(/\|/g, "") && existKeys[key]) { skipped++; continue; }
            for (var c = 0; c < row.length; c++) {
                dstWs.Cells(writeRow, 3 + c).Value2 = row[c];
            }
            dstWs.Cells(writeRow, 1).Value2 = writeRow - 1; // 序号
            existKeys[key] = true;
            writeRow++; written++;
        }
        dstWb.Save();
    } finally {
        Application.ScreenUpdating = true;
    }

    // ---------- 5. 在目标上跑"根据车牌号填充区域" ----------
    // 注意: 后台隐藏实例(Visible=false, 如 python COM 调用)跨簿 JSA Run 会卡死,
    // 只在用户可见的前台实例里执行; 后台模式由调用方(python)自行补做区域填充
    var runMsg = "区域填充: ";
    var runErr = null;
    if (!Application.Visible) {
        runMsg = "后台模式, 区域填充由调用方处理";
    } else {
        try {
            Application.Run("'" + dstPath + "'!根据车牌号填充区域");
            runMsg += "OK";
        } catch (e1) {
            try {
                // 容错:如果宏名少了"号"
                Application.Run("'" + dstPath + "'!根据车牌填充区域");
                runMsg += "OK(根据车牌填充区域)";
            } catch (e2) {
                runErr = (e1.message || e1) + " / " + (e2.message || e2);
                runMsg += "失败 — 目标文件没有根据车牌号填充区域宏,请先把上一个 JSA 注入";
            }
        }
    }

    // ---------- 6. 汇总 ----------
    var msg = "同步完成! 新写入 " + written + " 行(跳过重复 " + skipped + " 行),从第 " + startRow + " 行起" + filterMsg + "\n";
    for (var k in sheetCounts) msg += "  " + k + ": " + sheetCounts[k] + " 行\n";
    msg += runMsg;
    if (runErr) msg += "\n  详情: " + runErr;
    return msg;
}
