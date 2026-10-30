using ClosedXML.Excel;
using ExcelToJson;
using System.Text.Json;

var root = Path.Combine(AppContext.BaseDirectory, "fixtures", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
var failures = 0;
void Check(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}
void Test(string name, Action test)
{
    try { test(); Console.WriteLine($"PASS {name}"); }
    catch (Exception ex) { failures++; Console.WriteLine($"FAIL {name}: {ex.Message}"); }
}
ConversionConfig Config(SheetRuleDto sheet) => new()
{
    ExcelFiles = ["input.xlsx"],
    FileRules = new() { ["input.xlsx"] = new() { Sheets = new() { ["Data"] = sheet } } }
};
void Workbook(Action<IXLWorksheet> populate)
{
    using var wb = new XLWorkbook();
    populate(wb.AddWorksheet("Data"));
    wb.SaveAs(Path.Combine(root, "input.xlsx"));
}
JsonElement RunConversion(ConversionConfig config)
{
    var result = ExcelConverter.Run(root, root, config);
    Check(result.Success, result.Error ?? "Conversion failed");
    using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "input_Data.json")));
    return json.RootElement.Clone();
}

Test("automatic scan grows beyond old end and stops at blank", () =>
{
    Workbook(ws => { ws.Cell("B2").Value = 1; ws.Cell("B3").Value = 2; ws.Cell("B5").Value = 99; });
    var sheet = new SheetRuleDto { RowStart = 2, RowEnd = 2, ColStart = "B", ColEnd = "B" };
    Check(RunConversion(Config(sheet)).GetArrayLength() == 2, "Expected both contiguous rows, excluding data after gap");
    Check(sheet.RowEnd == 3, "RowEnd must become 3");
});
Test("default headers survive saving and repeated conversion", () =>
{
    Workbook(ws => { ws.Cell("A1").Value = "id"; ws.Cell("A2").Value = 7; ws.Cell("A4").Value = 99; });
    var config = Config(new SheetRuleDto());
    Check(RunConversion(config).GetArrayLength() == 1, "Generic sheet must stop at blank row");
    var path = Path.Combine(root, "rules.json");
    RulesStore.Save(config, path);
    var reloaded = RulesStore.LoadOrCreate(path);
    Check(reloaded.FileRules["input.xlsx"].Sheets["Data"].RowEnd == 2, "Detected row must persist");
    Check(RunConversion(reloaded)[0].GetProperty("id").GetInt32() == 7, "First row must remain the header");
});
Test("empty generic table replaces previous output with empty data", () =>
{
    Workbook(ws => { ws.Cell("A1").Value = "id"; ws.Cell("A2").Value = 7; });
    var config = Config(new SheetRuleDto());
    Check(RunConversion(config).GetArrayLength() == 1, "Initial conversion must have data");
    Workbook(ws => { ws.Cell("A1").Value = "id"; });
    Check(RunConversion(config).GetArrayLength() == 0, "Empty table must not retain stale JSON records");
    Check(config.FileRules["input.xlsx"].Sheets["Data"].RowEnd == 1, "Header-only sheet ends before first data row");
});
Test("manual mode keeps fixed end and reads past blank rows", () =>
{
    Workbook(ws => { ws.Cell("A2").Value = 1; ws.Cell("A4").Value = 2; ws.Cell("A5").Value = 3; });
    var sheet = new SheetRuleDto { RowStart = 2, RowEnd = 4, ColStart = "A", ColEnd = "A" };
    var config = Config(sheet);
    config.AutoDetectRows = false;
    Check(RunConversion(config).GetArrayLength() == 2, "Manual range must include data after the gap, stopping at 4");
    Check(sheet.RowEnd == 4, "Manual end must not change");
    var path = Path.Combine(root, "manual-rules.json");
    RulesStore.Save(config, path);
    Check(!RulesStore.LoadOrCreate(path).AutoDetectRows, "Manual switch must persist");
});
Test("blank detection ignores excluded and outside columns and precedes computed fields", () =>
{
    Workbook(ws =>
    {
        ws.Cell("B2").Value = 0; ws.Cell("C2").Value = false;
        ws.Cell("C3").Value = 2; // One empty cell is not an empty row.
        ws.Cell("B4").Value = "  "; ws.Cell("C4").FormulaA1 = "\"\"";
        ws.Cell("D4").Value = "excluded note"; ws.Cell("F4").Value = "outside region";
        ws.Cell("B5").Value = 99;
    });
    var sheet = new SheetRuleDto
    {
        RowStart = 2, RowEnd = 50, ColStart = "B", ColEnd = "D", ExcludeCols = ["D"],
        RemoveEmpty = false, ComputedColumns = new() { ["constant"] = "1" }
    };
    Check(RunConversion(Config(sheet)).GetArrayLength() == 2, "Whitespace/formula blank must stop before computed fields");
    Check(sheet.RowEnd == 3, "Shrinking table must update end to 3");
});
Test("subtables discover their own ends", () =>
{
    Workbook(ws =>
    {
        ws.Cell("B2").Value = 1; ws.Cell("B3").Value = 2; ws.Cell("B5").Value = 99;
        ws.Cell("E2").Value = 10; ws.Cell("E3").Value = 20; ws.Cell("E4").Value = 30;
    });
    var first = new SubTableRuleDto { RowStart = 2, RowEnd = 2, ColStart = "B", ColEnd = "B" };
    var second = new SubTableRuleDto { RowStart = 2, RowEnd = 20, ColStart = "E", ColEnd = "E" };
    var config = Config(new SheetRuleDto { SubTables = new() { ["first"] = first, ["second"] = second } });
    var json = RunConversion(config);
    Check(json.GetProperty("first").GetArrayLength() == 2, "First region must stop at its own blank");
    Check(json.GetProperty("second").GetArrayLength() == 3, "Neighbor region must continue");
    Check(first.RowEnd == 3 && second.RowEnd == 4, "Both subtable ends must update");
    var path = Path.Combine(root, "sub-rules.json");
    RulesStore.Save(config, path);
    var subs = RulesStore.LoadOrCreate(path).FileRules["input.xlsx"].Sheets["Data"].SubTables;
    Check(subs["first"].RowEnd == 3 && subs["second"].RowEnd == 4, "Subtable ends must persist");
});
Test("blank starting row produces no phantom computed records", () =>
{
    Workbook(ws => { ws.Cell("B1").Value = "header"; ws.Cell("B5").Value = 99; });
    var sheet = new SheetRuleDto
    {
        RowStart = 2, RowEnd = 10, ColStart = "B", ColEnd = "B", ComputedColumns = new() { ["constant"] = "1" }
    };
    Check(RunConversion(Config(sheet)).GetArrayLength() == 0, "Blank first row means no records");
    Check(sheet.RowEnd == 1, "Empty range is represented by start minus one");
});
Test("table without trailing blank stops at worksheet end", () =>
{
    Workbook(ws => { ws.Cell("A3").Value = 1; ws.Cell("A4").Value = 2; });
    var sheet = new SheetRuleDto { RowStart = 3 };
    Check(RunConversion(Config(sheet)).GetArrayLength() == 2 && sheet.RowEnd == 4, "Must include last used row");
});
Test("scan-only updates respect disabled sheets and old rules default to auto", () =>
{
    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("Data");
    ws.Cell("A2").Value = 1;
    var sheet = new SheetRuleDto { Enabled = false, RowStart = 2, RowEnd = 10 };
    ExcelConverter.UpdateRowEnds(ws, sheet);
    Check(sheet.RowEnd == 10, "Disabled sheet must remain untouched");
    sheet.Enabled = true;
    ExcelConverter.UpdateRowEnds(ws, sheet);
    Check(sheet.RowEnd == 2, "Scan-only must find actual end");
    var path = Path.Combine(root, "legacy-rules.json");
    File.WriteAllText(path, "{}");
    Check(RulesStore.LoadOrCreate(path).AutoDetectRows, "Old rules without switch must default to automatic");
});
Environment.ExitCode = failures == 0 ? 0 : 1;
