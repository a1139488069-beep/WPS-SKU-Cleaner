// Loading never scans or edits workbooks. Auto mode is session-only and opt-in.
var SKU_ADDIN = { loaded: false, registered: false, errors: [] };
function skuAddinError(err) {
    if (SKU_ADDIN.errors.length < 10) SKU_ADDIN.errors.push(String(err.message || err));
}
function SkuAddinLoad(ribbonUI) {
    try {
        if (!window.Application && typeof wps !== "undefined" && wps.EtApplication) window.Application = wps.EtApplication();
        if (!window.Application) throw new Error("WPS没有提供Application接口。");
        SKU_ADDIN.loaded = true;
        skuRegister(); SKU_ADDIN.registered = true;
    } catch (err) { skuAddinError(err); }
    return true;
}
function SkuAddinEnableAuto() { SKU_EnableAuto(); }
function SkuAddinDisableAuto() { SKU_DisableAuto(); }
function SkuAddinStop() { skuStop(false); }
function SkuAddinStatus() {
    alert("SKU工具 v1.0.5\n宿主代码加载: " + (SKU_ADDIN.loaded ? "成功" : "失败") +
        "\n事件注册: " + (SKU_ADDIN.registered ? "成功" : "失败") +
        "\n自动处理: " + (SKU_RUN.enabled ? "开启（仅绑定文件）" : "暂停") +
        (SKU_RUN.book ? "\n绑定文件: " + SKU_RUN.book.Name : "") +
        "\n任务: " + (SKU_RUN.current || SKU_RUN.pending.length ? "分批处理中" : "空闲") +
        "\n待处理任务: " + SKU_RUN.pending.length +
        (SKU_STATE.last ? "\n\n最近一次/当前进度:\n" + skuReport(SKU_STATE.last) : "\n尚未处理数据。") +
        (SKU_RUN.error ? "\n\n处理错误:\n" + SKU_RUN.error : "") +
        (SKU_ADDIN.errors.length ? "\n\n加载错误:\n" + SKU_ADDIN.errors.join("\n") : ""));
}
function SkuAddinSelfTest() {
    if (SKU_RUN.current || SKU_RUN.pending.length) { alert("请先等待当前任务完成或点击停止处理，再运行自测。"); return; }
    var app = Application, oldEvents = app.EnableEvents, oldAuto = SKU_RUN.enabled;
    var passed = false, message = "";
    try {
        SKU_RUN.enabled = false;
        app.EnableEvents = false;
        var book = app.Workbooks.Add(), sh = book.Worksheets.Item(1);
        sh.Name = "SKU工具自测";
        sh.Range("A1:C8").NumberFormat = "@";
        sh.Cells.Item(1, 1).Value2 = "公司sku";
        sh.Cells.Item(1, 2).Value2 = "平台sku";
        sh.Cells.Item(1, 3).Value2 = "其他列";
        sh.Cells.Item(2, 1).Value2 = "'2714230032211";
        sh.Cells.Item(3, 1).Value2 = "２７１４２３００３２２１１";
        sh.Cells.Item(4, 1).Value2 = " 'AB C-１２３";
        sh.Cells.Item(5, 1).Value2 = "1000000000000001";
        sh.Cells.Item(6, 1).NumberFormat = "General";
        sh.Cells.Item(6, 1).Formula = "=1+2";
        sh.Cells.Item(7, 1).Value2 = " =1+2 ";
        sh.Cells.Item(8, 1).Value2 = "1E12";
        sh.Cells.Item(2, 2).Value2 = "000123";
        sh.Cells.Item(2, 3).Value2 = "000123";
        var stats = skuExecute(book, sh, null);
        passed = typeof sh.Cells.Item(2, 1).Value2 === "number" && sh.Cells.Item(2, 1).Value2 === 2714230032211 &&
            String(sh.Cells.Item(2, 1).NumberFormat) === "0" && sh.Cells.Item(3, 1).Value2 === 2714230032211 &&
            sh.Cells.Item(2, 2).Value2 === 123 && sh.Cells.Item(4, 1).Value2 === "ABC-123" && String(sh.Cells.Item(4, 1).NumberFormat) === "0" &&
            sh.Cells.Item(5, 1).Value2 === "1000000000000001" && sh.Cells.Item(6, 1).HasFormula &&
            sh.Cells.Item(7, 1).Value2 === "=1+2" && !sh.Cells.Item(7, 1).HasFormula && sh.Cells.Item(8, 1).Value2 === "1E12" &&
            sh.Cells.Item(2, 3).Value2 === "000123" && stats.failed === 0 && stats.batches > 0;
        message = skuReport(stats);
    } catch (err) { skuAddinError(err); message = String(err.message || err); }
    finally { app.EnableEvents = oldEvents; SKU_RUN.enabled = oldAuto; }
    alert((passed ? "WPS批量接口自测通过。" : "WPS批量接口自测未通过，请查看兼容处理数量及错误。") + "\n已新建未保存的测试工作簿，现有文件未用于自测。\n\n" + message);
}
function SkuAddinHelp() {
    alert("SKU工具 v1.0.5\n\n打开文件和切换工作表不再扫描。每次启动自动处理默认暂停。\n先点击处理当前表；可开启当前文件的输入/粘贴自动处理，不扫描其他文件。\n普通数据按1000行一块批量读写，异常区域隔离处理。可用2万行测速查看实际用时；可查看状态/停止处理；停止保留已完成修改，不自动保存。\n表头仅检查前20行、前256列。含SKU文字、大小写不限。\n每个任务最多20万个SKU单元格，超出请用处理所选区域分次处理。\n纯数字整数转数值，自定义格式0；前导零会去掉。\n非纯数字也清理空格/隐藏字符/前置单引号，按文本存储并设为0格式。\n超过15位有效数字、公式保留；跳过合并单元格、错误及受保护表。\n空单元格不写格式。表头单独变化不扫描整表；已有数据用按钮处理。\n安装后先保存文件，完全退出WPS再重新打开，才能卸载内存中的旧版。");
}

function SkuAddinBenchmark() {
    if (SKU_RUN.current || SKU_RUN.pending.length) { alert("请先等待当前任务完成或点击停止处理，再测速。"); return; }
    var app = Application, oldEvents, oldScreen, captured = false, book, sh;
    try {
        oldEvents = app.EnableEvents; oldScreen = app.ScreenUpdating; captured = true;
        app.EnableEvents = false; app.ScreenUpdating = false;
        book = app.Workbooks.Add(); sh = book.Worksheets.Item(1); sh.Name = "SKU两万行测速";
        sh.Range("A1:B1").Value2 = [["公司sku", "平台sku"]];
        var values = [];
        for (var r = 0; r < 20000; r++) values.push(["000123", " 'AB C-１２３ "]);
        sh.Range("A2:B20001").NumberFormat = "@";
        sh.Range("A2:B20001").Value2 = values;
    } catch (err) { alert("无法创建测速表: " + String(err.message || err)); return; }
    finally { if (captured) { try { app.ScreenUpdating = oldScreen; } finally { app.EnableEvents = oldEvents; } } }
    try {
        skuQueue(book, [sh], null, true);
        SKU_RUN.pending[SKU_RUN.pending.length - 1].benchmark = true;
    } catch (err) { alert("无法开始测速: " + String(err.message || err)); }
}
