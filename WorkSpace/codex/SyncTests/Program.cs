using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using ClosedXML.Excel;
using ExcelToJson;

var runner = typeof(ConversionConfig).Assembly.GetType("ExcelToJson.AutomationRunner");
if (runner == null) { Console.WriteLine("FAIL: 尚无 SVN 更新、解析并同步入口"); return 1; }
var copyProperty = typeof(SheetRuleDto).GetProperty("CopyToTarget");
if (copyProperty == null) { Console.WriteLine("FAIL: 工作表缺少默认关闭的复制开关"); return 1; }
Check(!(bool)copyProperty.GetValue(new SheetRuleDto())!, "new sheet copy defaults off");
Check(!JsonSerializer.Deserialize<SheetRuleDto>("{}")!.CopyToTarget, "old rules copy defaults off");
if (args.Contains("--real-data"))
{
    var project = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../../"));
    var actual = RulesStore.LoadOrCreate(Path.Combine(project, "publish", "rules.json"));
    var warnings = (List<string>)runner.GetMethod("ValidateInputs", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [actual, actual.ExcelDirectory!, new HashSet<string>(StringComparer.OrdinalIgnoreCase)])!;
    foreach (var warning in warnings) Console.WriteLine("WARNING " + warning);
    var destination = Path.Combine(project, "codex", "real-data-check", Guid.NewGuid().ToString("N"));
    var result = ExcelConverter.Run(actual.ExcelDirectory!, destination, actual);
    if (!result.Success) throw new Exception(result.Error);
    var jsonFiles = Directory.GetFiles(destination, "*.json");
    foreach (var file in jsonFiles) { using var json = JsonDocument.Parse(File.ReadAllText(file)); }
    Console.WriteLine($"PASS real workbook / existing rules: {jsonFiles.Length} JSON files; output only in {destination}");
    return 0;
}
var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../fixtures-sync", Guid.NewGuid().ToString("N")));
Directory.CreateDirectory(root);
var svn = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../tools/svn/bin/svn.exe"));
var admin = Path.Combine(Path.GetDirectoryName(svn)!, "svnadmin.exe");
var repo = Path.Combine(root, "repo");
var wcRoot = Path.Combine(root, "本地 表格");
var otherRoot = Path.Combine(root, "另一份");
var wc = Path.Combine(wcRoot, "中文项目");
var other = Path.Combine(otherRoot, "中文项目");
var output = Path.Combine(root, "output");
var target = Path.Combine(root, "Cocos 资源");
Directory.CreateDirectory(output); Directory.CreateDirectory(target);
try
{
    await Cmd(admin, "create", repo);
    var url = new Uri(repo + Path.DirectorySeparatorChar).AbsoluteUri;
    await Cmd(svn, "mkdir", url + Uri.EscapeDataString("中文项目"), "-m", "Chinese versioned directory");
    await Cmd(svn, "checkout", url, wcRoot);
    WriteWorkbook(wc, 1);
    await Cmd(svn, "add", Path.Combine(wc, "input.xlsx"));
    await Cmd(svn, "commit", wc, "-m", "initial fixture");
    await Cmd(svn, "checkout", url, otherRoot);
    WriteWorkbook(other, 2);
    await Cmd(svn, "commit", other, "-m", "new Excel");
    var config = Config();
    File.WriteAllText(Path.Combine(output, "stale.json"), "[999]");
    File.WriteAllText(Path.Combine(target, "Data.json"), "[0]");
    File.WriteAllText(Path.Combine(target, "Data.json.meta"), "preserved UUID");
    File.WriteAllText(Path.Combine(target, "Unrelated.json"), "[123]");
    // Default-off still exports, and must not touch or even require a Cocos folder.
    config.TargetDirectory = Path.Combine(root, "not-created-target");
    var defaultResult = await AutomationRunner.RunAsync(config);
    Check(defaultResult.FileCount == 0, "default-off must report zero copied files");
    Check(File.Exists(Path.Combine(output, "Data.json")), "default-off must still generate JSON");
    Check(!Directory.Exists(config.TargetDirectory), "default-off must not create target");
    Check(File.ReadAllText(Path.Combine(target, "Data.json")) == "[0]", "default-off preserves existing target");
    config.TargetDirectory = target;
    config.FileRules["input.xlsx"].Sheets["Data"].CopyToTarget = true;
    config.FileRules["input.xlsx"].Sheets["Extra"] = new() { RowStart = 2, ColStart = "A", ColEnd = "A", OutputName = "Extra.json" };
    File.WriteAllText(Path.Combine(target, "Extra.json"), "[77]");
    var enabledResult = await AutomationRunner.RunAsync(config);
    Check(enabledResult.FileCount == 1, "only opted-in JSON is counted");
    Check(File.Exists(Path.Combine(output, "Extra.json")), "opted-out sheet still generates");
    Check(File.ReadAllText(Path.Combine(target, "Extra.json")) == "[77]", "opted-out sheet is not copied");
    config.FileRules["input.xlsx"].Sheets.Remove("Extra");
    var published = File.ReadAllText(Path.Combine(target, "Data.json"));
    using (var json = JsonDocument.Parse(published))
        Check(json.RootElement[0].GetProperty("id").GetInt32() == 2, "必须先更新再读取 Excel");
    Check(!File.Exists(Path.Combine(target, "stale.json")), "不得同步旧输出");
    Check(File.ReadAllText(Path.Combine(target, "Data.json.meta")) == "preserved UUID", "不得修改 .meta");
    Check(File.ReadAllText(Path.Combine(target, "Unrelated.json")) == "[123]", "不得删除其他 JSON");
    Check(File.ReadAllText(Path.Combine(output, "Data.json")) == published, "同时更新输出目录");
    Console.WriteLine("PASS real SVN update → fresh Excel → JSON copy; stale files and metadata preserved");

    config.ExcelFiles.Add("missing.xlsx");
    await Fails(config, "缺失工作簿"); config.ExcelFiles.Remove("missing.xlsx");
    config.FileRules["input.xlsx"].Sheets["MissingSheet"] = new() { OutputName = "Missing.json" };
    File.WriteAllText(Path.Combine(target, "Missing.json"), "[456]");
    var missingResult = await AutomationRunner.RunAsync(config);
    Check(missingResult.Warnings.Count == 1, "obsolete sheet must be visibly reported");
    Check(File.ReadAllText(Path.Combine(target, "Missing.json")) == "[456]", "missing sheet preserves old JSON");
    config.FileRules["input.xlsx"].Sheets.Remove("MissingSheet");
    config.FileRules["input.xlsx"].Sheets["Extra"] = new() { OutputName = "Data.json" };
    await Fails(config, "重名输出"); config.FileRules["input.xlsx"].Sheets.Remove("Extra");
    var sheet = config.FileRules["input.xlsx"].Sheets["Data"];
    sheet.OutputName = "../escape.json"; await Fails(config, "越界输出");
    sheet.OutputName = "Data.json.meta"; await Fails(config, "非 JSON 输出");
    sheet.OutputName = "Data.json";
    sheet.Enabled = false; await Fails(config, "空转换"); sheet.Enabled = true;
    Set(config, "SvnDirectory", target); await Fails(config, "非 SVN 目录"); Set(config, "SvnDirectory", wc);
    using (var canceled = new CancellationTokenSource())
    {
        canceled.Cancel(); await Fails(config, "取消", canceled.Token);
    }
    var configFile = Path.Combine(root, "rules.json"); RulesStore.Save(config, configFile);
    var reloaded = RulesStore.LoadOrCreate(configFile);
    Check((string?)typeof(ConversionConfig).GetProperty("TargetDirectory")!.GetValue(reloaded) == target, "配置必须持久化");
    Check(reloaded.FileRules["input.xlsx"].Sheets["Data"].CopyToTarget, "copy opt-in must persist");
    Console.WriteLine("PASS validation, cancellation and settings persistence");

    var subConfig = Config();
    subConfig.FileRules["input.xlsx"].Sheets["Data"].SubTables = new()
    {
        ["left"] = new() { RowStart = 2, ColStart = "A", ColEnd = "A", OutputName = "PartA.json" },
        ["right"] = new() { RowStart = 2, ColStart = "A", ColEnd = "A", OutputName = "PartB.json" }
    };
    var subsOff = await AutomationRunner.RunAsync(subConfig);
    Check(subsOff.FileCount == 0 && File.Exists(Path.Combine(output, "PartA.json")) && !File.Exists(Path.Combine(target, "PartA.json")), "subtable output obeys default-off");
    subConfig.FileRules["input.xlsx"].Sheets["Data"].CopyToTarget = true;
    var subsOn = await AutomationRunner.RunAsync(subConfig);
    Check(subsOn.FileCount == 2 && File.Exists(Path.Combine(target, "PartA.json")) && File.Exists(Path.Combine(target, "PartB.json")), "sheet opt-in copies all subtable outputs");
    Console.WriteLine("PASS per-sheet copy defaults off, selective copy and subtable inheritance");

    // Binary Excel conflicts must stop even if svn update returns zero.
    WriteWorkbook(wc, 3); WriteWorkbook(other, 4);
    await Cmd(svn, "commit", other, "-m", "conflicting Excel");
    await Fails(config, "SVN 冲突");
    Check((await Cmd(svn, "status", "--xml", wc)).Contains("conflicted"), "fixture must produce actual SVN conflict");
    Console.WriteLine("PASS actual SVN conflict prevents publishing");
    await Cmd(svn, "revert", "-R", wc);
    await Cmd(svn, "update", wc);

    File.WriteAllText(Path.Combine(wc, "input.xlsx"), "not an Excel file");
    await Fails(config, "解析失败");
    Console.WriteLine("PASS parse failure preserves previous Cocos output");
    await Cmd(svn, "revert", "-R", wc);

    // A later file denied replacement must roll back earlier replacements.
    config.FileRules["input.xlsx"].Sheets["Extra"] = new() { CopyToTarget = true, RowStart = 2, ColStart = "A", ColEnd = "A", CustomHeaders = new() { ["A"] = "id" }, OutputName = "Extra.json" };
    File.WriteAllText(Path.Combine(target, "Extra.json"), "[88]");
    using (var locked = new FileStream(Path.Combine(target, "Extra.json"), FileMode.Open, FileAccess.Read, FileShare.Read))
        await Fails(config, "复制失败回滚");
    Check(File.ReadAllText(Path.Combine(target, "Extra.json")) == "[88]", "locked target preserved");
    Check(File.ReadAllText(Path.Combine(output, "Data.json")) == published, "output rollback");
    Console.WriteLine("PASS copy failure rolls back earlier files");
    Console.WriteLine("ALL SYNC TESTS PASSED");
    return 0;

    async Task Fails(ConversionConfig c, string label, CancellationToken token = default)
    {
        var failed = false;
        try { await Run(c, token); } catch { failed = true; }
        Check(failed, label + " 必须失败");
        Check(File.ReadAllText(Path.Combine(target, "Data.json")) == published, label + " 不得覆盖目标");
    }
}
catch (Exception ex) { Console.WriteLine($"FAIL {ex}"); return 1; }

ConversionConfig Config()
{
    var c = new ConversionConfig { ExcelDirectory = wc, OutputDirectory = output, ExcelFiles = ["input.xlsx"], FileRules = new()
    { ["input.xlsx"] = new() { Sheets = new() { ["Data"] = new() { RowStart = 2, ColStart = "A", ColEnd = "A", CustomHeaders = new() { ["A"] = "id" }, OutputName = "Data.json" } } } } };
    Set(c, "SvnDirectory", wc); Set(c, "TargetDirectory", target); Set(c, "SvnExecutable", svn);
    return c;
}
void Set(ConversionConfig c, string name, string value) => typeof(ConversionConfig).GetProperty(name)!.SetValue(c, value);
async Task Run(ConversionConfig c, CancellationToken token = default)
{
    try { await (Task)runner!.GetMethod("RunAsync")!.Invoke(null, [c, null, token])!; }
    catch (TargetInvocationException ex) { throw ex.InnerException!; }
}
static void Check(bool value, string message) { if (!value) throw new Exception(message); }
static void WriteWorkbook(string dir, int value)
{
    using var wb = new XLWorkbook();
    foreach (var name in new[] { "Data", "Extra" })
    { var ws = wb.AddWorksheet(name); ws.Cell(1, 1).Value = "id"; ws.Cell(2, 1).Value = value; }
    wb.SaveAs(Path.Combine(dir, "input.xlsx"));
}
static async Task<string> Cmd(string exe, params string[] args)
{
    var si = new ProcessStartInfo(exe) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
    // Fixture operations use native Unicode working directories and relative targets.
    if (Path.GetFileName(exe) == "svn.exe" && args[0] is "add" or "commit" or "revert" or "update" or "status")
    {
        var index = Array.FindLastIndex(args, p => Path.IsPathFullyQualified(p));
        if (index > 0)
        {
            var targetPath = args[index];
            si.WorkingDirectory = Directory.Exists(targetPath) ? targetPath : Path.GetDirectoryName(targetPath)!;
            args[index] = Directory.Exists(targetPath) ? "." : Path.GetFileName(targetPath);
        }
    }
    foreach (var arg in args) si.ArgumentList.Add(arg);
    using var p = Process.Start(si)!;
    var stdout = p.StandardOutput.ReadToEndAsync(); var stderr = p.StandardError.ReadToEndAsync();
    await p.WaitForExitAsync(); var result = await stdout; var error = await stderr;
    if (p.ExitCode != 0) throw new Exception($"{Path.GetFileName(exe)} {args[0]} failed: {error}");
    return result;
}
