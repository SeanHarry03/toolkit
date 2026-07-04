using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using ClosedXML.Excel;

// ====================================================================
//  Excel → JSON 转换工具 (C# 版)
//  所有转换规则与 convert.js 保持一致
// ====================================================================

// ==================== 路径配置 ====================

// Excel 文件所在目录：先尝试当前目录，再向上查找（兼容 VS 调试和发布场景）
string excelDir = ResolveExcelDirectory();
string outputDir = Path.Combine(excelDir, "output");
Directory.CreateDirectory(outputDir);

Console.WriteLine($"Excel 目录: {excelDir}");
Console.WriteLine($"输出目录: {outputDir}");
Console.WriteLine();

// ==================== 排除的 Sheet ====================

var excludeSheets = new HashSet<string>
{
    "任务列表",
    "文本配置",
    "动效需求",
    "称号_商城_结算",
    "人扮演鬼模式",
    "整体",
};

// ==================== 默认规则 ====================

var defaultRule = new FileRule
{
    KeyColumn = null,
    ArrayFields = [],
    RemoveEmpty = true,
    ExcludeCols = [],
    OutputNames = [],
};

// ==================== 自定义文件规则 ====================

var fileRules = new Dictionary<string, FileRule>
{
    ["策划案.xlsx"] = new FileRule
    {
        Sheets = new Dictionary<string, SheetRule>
        {
            ["道具配置表"] = new SheetRule
            {
                // 1-indexed 行范围（对应 Excel 左侧物理行号）
                RowRange = new RowRangeDef(18, 177),
                // 列范围
                ColRange = new ColRangeDef("B", "AE"),
                // 跳过的列
                ExcludeCols = ["G"],
                // 自定义表头映射（列字母 → 字段名）
                CustomHeaders = new Dictionary<string, string>
                {
                    ["B"] = "id",
                    ["C"] = "itemType",
                    ["D"] = "itemLevel",
                    ["E"] = "maxLevel",
                    ["F"] = "itemName",
                    ["H"] = "itemHp",
                    ["I"] = "upgradeGold",
                    ["J"] = "upgradeEnergy",
                    ["K"] = "upgradeRequire",
                    ["L"] = "removeGold",
                    ["M"] = "removeEnergy",
                    ["N"] = "attack",
                    ["O"] = "attackRange",
                    ["P"] = "buildGoldCost",
                    ["Q"] = "buildEnergyCost",
                    ["R"] = "goldPerSec",
                    ["S"] = "energyPerSec",
                    ["T"] = "itemDescription",
                    ["U"] = "describeValues",
                    ["V"] = "triggerProbability",
                    ["W"] = "triggerProbability",
                    ["X"] = "maxCount",
                    ["Y"] = "activeDuration",
                    ["Z"] = "skillCd",
                    ["AA"] = "skillEffect",
                    ["AB"] = "skillValue",
                    ["AC"] = "skillDuration",
                    ["AD"] = "useOnce",
                    ["AE"] = "spiritKillExp",
                },
                KeyColumn = "id",
                ArrayFields = ["describeValues"],
                RemoveEmpty = true,
                OutputName = "TowerConfig.json"
            },
            ["猛鬼配置表"] = new SheetRule()
            {
                RowRange = new RowRangeDef(15, 254),
                // 列范围
                ColRange = new ColRangeDef("D", "N"),
                CustomHeaders = new Dictionary<string, string>
                {
                    ["D"] = "spiritType",
                    ["E"] = "spiritName",
                    ["F"] = "level",
                    ["G"] = "levelUpExp",
                    ["H"] = "expPerSec",
                    ["I"] = "hp",
                    ["J"] = "attack",
                    ["K"] = "moveSpeed",
                    ["L"] = "attackSpeed",
                    ["M"] = "damageToExpRate",
                    ["N"] = "skillList",
                },
                KeyColumn = "spiritId",
                ArrayFields = ["skillList"],
                RemoveEmpty = true,
                OutputName = "EnemyConfig.json",
                //计算得出，Excel原本不存在的列
                ComputedColumns = new Dictionary<string, Func<IReadOnlyDictionary<string, object?>, object?>>
                {
                    ["spiritId"] = row =>
                    {
                        var spiritType = GetDouble(row, "spiritType");
                        var level = GetDouble(row, "level");
                        return spiritType * 100 + level;
                    }
                }
            },
            ["猛鬼技能效果"] = new SheetRule()
            {
                // 1-indexed 行范围（对应 Excel 左侧物理行号）
                RowRange = new RowRangeDef(19, 109),
                // 列范围
                ColRange = new ColRangeDef("E", "N"),
                CustomHeaders = new Dictionary<string, string>
                {
                    ["E"] = "skillId",
                    ["F"] = "skillType",
                    ["G"] = "skillLevel",
                    ["H"] = "skillName",
                    ["I"] = "skillProbability",
                    ["J"] = "skillDuration",
                    ["K"] = "skillCd",
                    ["L"] = "skillEffect",
                    ["M"] = "skillValue",
                    ["N"] = "triggerCondition",
                },
                KeyColumn = "skillId",
                RemoveEmpty = true,
                OutputName = "EnemySkillConfig.json",
            },
            ["默认道具和祈福签配置"] = new SheetRule()
            {
                OutputName = "默认道具和祈福签配置.json",
                SubTables = new Dictionary<string, SubTableRule>
                {
                    ["601"] = new SubTableRule
                    {
                        RowRange = new RowRangeDef(50, 80),
                        ColRange = new ColRangeDef("C", "D"),
                        CustomHeaders = new Dictionary<string, string>
                        {
                            ["C"] = "id",
                            ["D"] = "probability",
                        },
                        KeyColumn = "id",
                    },
                    ["701"] = new SubTableRule
                    {
                        RowRange = new RowRangeDef(50, 81),
                        ColRange = new ColRangeDef("G", "H"),
                        CustomHeaders = new Dictionary<string, string>
                        {
                            ["G"] = "id",
                            ["H"] = "probability",
                        },
                        KeyColumn = "id",
                    },
                    ["801"] = new SubTableRule
                    {
                        RowRange = new RowRangeDef(50, 90),
                        ColRange = new ColRangeDef("L", "M"),
                        CustomHeaders = new Dictionary<string, string>
                        {
                            ["L"] = "id",
                            ["M"] = "probability",
                        },
                        KeyColumn = "id",
                    },
                }
            },
            ["挑战模式"] = new SheetRule()
            {
                OutputName = "LevelConfig.json",
                // 1-indexed 行范围（对应 Excel 左侧物理行号）
                RowRange = new RowRangeDef(14, 28),
                // 列范围
                ColRange = new ColRangeDef("H", "M"),
                CustomHeaders = new Dictionary<string, string>
                {
                    ["H"] = "level",
                    ["I"] = "hpFactor",
                    ["J"] = "attackFactor",
                    ["K"] = "moveFactor",
                    ["L"] = "attackSpeed",
                    ["M"] = "waveList",
                },
                KeyColumn = "level",
                RemoveEmpty = true,
                ArrayFields = ["waveList"],
                
            }
        }
    },
};

// ==================== 要转换的 Excel 文件 ====================

var excelFiles = new List<string>
{
    "策划案.xlsx",
};

// ==================== 辅助方法 ====================

// 解析 Excel 文件目录
static string ResolveExcelDirectory()
{
    // 候选目录：当前目录、以及向上查找
    var candidates = new List<string> { Directory.GetCurrentDirectory() };

    // 从当前目录向上查找包含 .xlsx 文件的目录
    string? dir = Directory.GetCurrentDirectory();
    for (int i = 0; i < 5; i++)
    {
        dir = Path.GetDirectoryName(dir);
        if (dir != null)
            candidates.Add(dir);
    }

    foreach (var candidate in candidates)
    {
        if (Directory.EnumerateFiles(candidate, "*.xlsx").Any())
            return candidate;
    }

    // 兜底：当前目录
    return Directory.GetCurrentDirectory();
}

// 字母列名转 0-based 列索引（A → 0, B → 1, Z → 25, AA → 26）
static int ColLetterToIndex(string col)
{
    int index = 0;
    foreach (char c in col)
    {
        index = index * 26 + (c - 'A' + 1);
    }

    return index - 1;
}

// 0-based 列索引转字母列名（0 → A, 1 → B, 25 → Z, 26 → AA）
static string IndexToColLetter(int index)
{
    string letter = "";
    int temp = index;
    while (temp >= 0)
    {
        letter = (char)('A' + (temp % 26)) + letter;
        temp = temp / 26 - 1;
    }

    return letter;
}

// 修复 Excel 浮点数精度误差
// 例：40960.0000000001 → 40960, 799.999999999999 → 800
static double FixFloat(double value)
{
    // 如果离整数非常近（1e-9 以内），直接取整
    double rounded = Math.Round(value);
    if (Math.Abs(value - rounded) < 1e-9)
        return rounded;

    // 否则清理到 9 位小数，去除浮点噪声
    return Math.Round(value, 9);
}

// 从行数据中安全获取 double 值（用于计算列）
static double GetDouble(IReadOnlyDictionary<string, object?> row, string key)
{
    if (row.TryGetValue(key, out var val) && val is double d)
        return d;
    return 0;
}

// 从单元格获取原始值
static object? GetCellValue(IXLCell cell)
{
    if (cell.IsEmpty() || cell.Value.IsBlank)
        return null;

    if (cell.Value.IsNumber)
        return FixFloat(cell.Value.GetNumber());

    if (cell.Value.IsBoolean)
        return cell.Value.GetBoolean();

    // 文本或其他类型
    return cell.Value.GetText()?.Replace("\\n", "\n");
}

// 处理单行数据：去空、转数组
static Dictionary<string, object?> ProcessRow(
    Dictionary<string, object?> row,
    List<string>? arrayFields,
    bool removeEmpty)
{
    var processed = new Dictionary<string, object?>();

    foreach (var (key, value) in row)
    {
        // 跳过空值
        if (removeEmpty && (value == null || (value is string s && string.IsNullOrEmpty(s))))
            continue;

        // 需要转数组的字段
        if (arrayFields != null && arrayFields.Contains(key))
        {
            if (value is string str && str.Contains(','))
            {
                // 逗号分隔 → 数组，尝试转数字（去除尾逗号产生的空段）
                processed[key] = str.Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(v =>
                    {
                        var trimmed = v.Trim();
                        if (trimmed.Length == 0)
                            return null;
                        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double num))
                            return (object)FixFloat(num);
                        return trimmed;
                    })
                    .Where(v => v != null)!
                    .ToList();
            }
            else
            {
                // 单个值也包成数组
                var list = new List<object?>();
                if (value != null)
                {
                    if (value is double dVal)
                        list.Add(FixFloat(dVal));
                    else if (double.TryParse(value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture,
                                 out double num))
                        list.Add(FixFloat(num));
                    else
                        list.Add(value);
                }

                processed[key] = list;
            }
        }
        else
        {
            processed[key] = value;
        }
    }

    return processed;
}

// 按 key 字段包裹数组为字典
static Dictionary<string, Dictionary<string, object?>> WrapWithKey(
    List<Dictionary<string, object?>> rows,
    string keyColumn)
{
    var result = new Dictionary<string, Dictionary<string, object?>>();
    foreach (var row in rows)
    {
        if (row.TryGetValue(keyColumn, out var keyValue) && keyValue != null)
        {
            result[keyValue.ToString()!] = row;
        }
    }

    return result;
}

// 处理一个子表区域，返回按规则转换后的数据（数组或 key 包裹的字典）
static object ProcessRegion(
    IXLWorksheet worksheet,
    SubTableRule rule)
{
    int startRow = rule.RowRange?.Start ?? 1;
    int endRow = rule.RowRange?.End ?? worksheet.LastRowUsed()?.RowNumber() ?? 1;
    int startCol = rule.ColRange.HasValue
        ? ColLetterToIndex(rule.ColRange.Value.StartCol) + 1
        : 1;
    int endCol = rule.ColRange.HasValue
        ? ColLetterToIndex(rule.ColRange.Value.EndCol) + 1
        : worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

    var rows = new List<Dictionary<string, object?>>();

    for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++)
    {
        var row = worksheet.Row(rowIdx);
        var newRow = new Dictionary<string, object?>();

        for (int colIdx = startCol; colIdx <= endCol; colIdx++)
        {
            string colLetter = IndexToColLetter(colIdx - 1);
            if (rule.ExcludeCols.Contains(colLetter))
                continue;

            var cell = row.Cell(colIdx);
            var val = GetCellValue(cell);

            string key;
            if (rule.CustomHeaders != null &&
                rule.CustomHeaders.TryGetValue(colLetter, out var customKey))
                key = customKey;
            else
                key = colLetter;

            newRow[key] = val;
        }

        // 计算列
        if (rule.ComputedColumns != null)
        {
            foreach (var (compKey, compFunc) in rule.ComputedColumns)
            {
                var computed = compFunc(newRow);
                if (computed is double cd)
                    newRow[compKey] = FixFloat(cd);
                else if (computed != null)
                    newRow[compKey] = computed;
            }
        }

        var cleanRow = ProcessRow(newRow, rule.ArrayFields, rule.RemoveEmpty);
        if (cleanRow.Count > 0)
            rows.Add(cleanRow);
    }

    // 是否按 key 包裹
    if (!string.IsNullOrEmpty(rule.KeyColumn))
        return WrapWithKey(rows, rule.KeyColumn!);

    return rows;
}

// ==================== JSON 序列化选项 ====================

var jsonOptions = new JsonSerializerOptions
{
    WriteIndented = true,
    // 不转义中文等 Unicode 字符
    Encoder = JavaScriptEncoder.Create(
        UnicodeRanges.BasicLatin,
        UnicodeRanges.CjkUnifiedIdeographs,
        UnicodeRanges.CjkSymbolsandPunctuation,
        UnicodeRanges.HalfwidthandFullwidthForms,
        UnicodeRanges.CjkCompatibility
    ),
};

// ==================== 主逻辑 ====================

foreach (var fileName in excelFiles)
{
    var filePath = Path.Combine(excelDir, fileName);

    if (!File.Exists(filePath))
    {
        Console.WriteLine($"[跳过] 文件不存在: {fileName}");
        continue;
    }

    Console.WriteLine($"[转换] {fileName} ...");

    // 获取当前文件的规则
    fileRules.TryGetValue(fileName, out var fileRule);

    using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    using var workbook = new XLWorkbook(fileStream);

    foreach (var worksheet in workbook.Worksheets)
    {
        var sheetName = worksheet.Name;

        if (excludeSheets.Contains(sheetName))
        {
            Console.WriteLine($"  [排除] {sheetName}");
            continue;
        }

        // 检查是否有该 sheet 的特定配置
        SheetRule? sheetRule = null;
        if (fileRule != null && fileRule.Sheets != null)
        {
            fileRule.Sheets.TryGetValue(sheetName, out sheetRule);
        }

        List<Dictionary<string, object?>> processedRows;
        FileRule finalRule;

        if (sheetRule != null)
        {
            // ======== 多子表模式：一个 Sheet 内多个独立区域 ========
            if (sheetRule.SubTables != null)
            {
                var combined = new Dictionary<string, object>();
                int totalCount = 0;
                foreach (var (subKey, subRule) in sheetRule.SubTables)
                {
                    var regionData = ProcessRegion(worksheet, subRule);
                    combined[subKey] = regionData;
                    totalCount += regionData is Dictionary<string, Dictionary<string, object?>> d
                        ? d.Count
                        : ((System.Collections.IList)regionData).Count;
                }

                string subOutName = sheetRule.OutputName
                                    ?? $"{Path.GetFileNameWithoutExtension(fileName)}_{sheetName}.json";

                var subOutPath = Path.Combine(outputDir, subOutName);
                var subJson = JsonSerializer.Serialize(combined, jsonOptions);
                File.WriteAllText(subOutPath, subJson, System.Text.Encoding.UTF8);
                Console.WriteLine($"  -> {subOutName} ({combined.Count} 个区域, {totalCount} 条数据)");
                continue;
            }

            // ======== 单区域模式：走特定 sheet 裁剪 + 自定义表头逻辑 ========
            finalRule = sheetRule.ToFileRule();

            // 1. 行范围（1-indexed，对应 Excel 物理行号）
            int startRow = sheetRule.RowRange?.Start ?? 1;
            int endRow = sheetRule.RowRange?.End ?? worksheet.LastRowUsed()?.RowNumber() ?? 1;

            // 2. 列范围
            int startCol = sheetRule.ColRange.HasValue
                ? ColLetterToIndex(sheetRule.ColRange.Value.StartCol) + 1 // 转 1-based
                : 1;
            int endCol = sheetRule.ColRange.HasValue
                ? ColLetterToIndex(sheetRule.ColRange.Value.EndCol) + 1
                : worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

            processedRows = [];

            for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++)
            {
                var row = worksheet.Row(rowIdx);
                var newRow = new Dictionary<string, object?>();

                for (int colIdx = startCol; colIdx <= endCol; colIdx++)
                {
                    string colLetter = IndexToColLetter(colIdx - 1); // 转 0-based 字母

                    // 跳过的列
                    if (sheetRule.ExcludeCols.Contains(colLetter))
                        continue;

                    var cell = row.Cell(colIdx);
                    var val = GetCellValue(cell);

                    // 自定义表头映射；未指定则用列字母
                    string key;
                    if (sheetRule.CustomHeaders != null &&
                        sheetRule.CustomHeaders.TryGetValue(colLetter, out var customKey))
                    {
                        key = customKey;
                    }
                    else
                    {
                        key = colLetter;
                    }

                    newRow[key] = val;
                }

                // 计算列：从已有列推导出新列
                if (sheetRule.ComputedColumns != null)
                {
                    foreach (var (compKey, compFunc) in sheetRule.ComputedColumns)
                    {
                        var computed = compFunc(newRow);
                        if (computed != null)
                            newRow[compKey] = computed;
                    }
                }

                // 应用类型转换（去空、转数组等）
                var cleanRow = ProcessRow(newRow, sheetRule.ArrayFields, sheetRule.RemoveEmpty);
                if (cleanRow.Count > 0)
                {
                    processedRows.Add(cleanRow);
                }
            }
        }
        else
        {
            // ======== 默认通用逻辑：第一行是表头 ========
            finalRule = fileRule ?? defaultRule;

            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            var lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

            if (lastRow < 2) continue; // 至少需要表头 + 一行数据

            // 读取第一行作为表头
            var headerRow = worksheet.Row(1);
            var colToKeyMap = new Dictionary<string, string>();
            for (int colIdx = 1; colIdx <= lastCol; colIdx++)
            {
                string colLetter = IndexToColLetter(colIdx - 1);
                var headerCell = headerRow.Cell(colIdx);
                colToKeyMap[colLetter] = headerCell.IsEmpty()
                    ? colLetter
                    : (headerCell.Value.GetText() ?? colLetter);
            }

            processedRows = [];

            // 从第二行开始读取数据
            for (int rowIdx = 2; rowIdx <= lastRow; rowIdx++)
            {
                var row = worksheet.Row(rowIdx);
                var newRow = new Dictionary<string, object?>();

                for (int colIdx = 1; colIdx <= lastCol; colIdx++)
                {
                    string colLetter = IndexToColLetter(colIdx - 1);

                    // 跳过的列
                    if (finalRule.ExcludeCols.Contains(colLetter))
                        continue;

                    var cell = row.Cell(colIdx);
                    var val = GetCellValue(cell);

                    string key = colToKeyMap.TryGetValue(colLetter, out var hdr)
                        ? hdr
                        : colLetter;

                    newRow[key] = val;
                }

                var cleanRow = ProcessRow(newRow,
                    finalRule.ArrayFields,
                    finalRule.RemoveEmpty);
                processedRows.Add(cleanRow);
            }
        }

        // 是否按 key 字段包裹外层
        object outputData;
        if (!string.IsNullOrEmpty(finalRule.KeyColumn))
        {
            outputData = WrapWithKey(processedRows, finalRule.KeyColumn!);
        }
        else
        {
            outputData = processedRows;
        }

        // 确定输出文件名：sheet 规则名 > 文件规则映射名 > 默认名
        string outputFileName =
            sheetRule?.OutputName
            ?? (finalRule.OutputNames.TryGetValue(sheetName, out var mappedName) ? mappedName : null)
            ?? $"{Path.GetFileNameWithoutExtension(fileName)}_{sheetName}.json";

        // 写入 JSON
        var outputPath = Path.Combine(outputDir, outputFileName);
        var jsonContent = JsonSerializer.Serialize(outputData, jsonOptions);
        File.WriteAllText(outputPath, jsonContent, System.Text.Encoding.UTF8);

        int count = outputData is Dictionary<string, Dictionary<string, object?>> dict
            ? dict.Count
            : ((System.Collections.IList)outputData).Count;

        Console.WriteLine($"  -> {outputFileName} ({count} 条数据)");
    }

    Console.WriteLine();
}

Console.WriteLine($"✅ 全部完成! JSON 文件已输出到: {outputDir}");

// ==================== 规则类型定义 ====================

/// <summary>列范围</summary>
internal record struct ColRangeDef(string StartCol, string EndCol);

/// <summary>行范围（1-indexed，对应 Excel 物理行号）</summary>
internal record struct RowRangeDef(int Start, int End);

/// <summary>子表规则（一个 Sheet 内多个独立区域的配置）</summary>
internal class SubTableRule
{
    public RowRangeDef? RowRange { get; set; }
    public ColRangeDef? ColRange { get; set; }
    public List<string> ExcludeCols { get; set; } = [];
    public Dictionary<string, string>? CustomHeaders { get; set; }
    public string? KeyColumn { get; set; }
    public List<string> ArrayFields { get; set; } = [];
    public bool RemoveEmpty { get; set; } = true;
    public Dictionary<string, Func<IReadOnlyDictionary<string, object?>, object?>>? ComputedColumns { get; set; }
}

/// <summary>Sheet 级规则</summary>
internal class SheetRule
{
    public RowRangeDef? RowRange { get; set; }
    public ColRangeDef? ColRange { get; set; }
    public List<string> ExcludeCols { get; set; } = [];
    public Dictionary<string, string>? CustomHeaders { get; set; }
    public string? KeyColumn { get; set; }
    public List<string> ArrayFields { get; set; } = [];
    public bool RemoveEmpty { get; set; } = true;
    public string? OutputName { get; set; }
    public Dictionary<string, Func<IReadOnlyDictionary<string, object?>, object?>>? ComputedColumns { get; set; }
    public Dictionary<string, SubTableRule>? SubTables { get; set; }

    public FileRule ToFileRule()
    {
        return new FileRule
        {
            KeyColumn = KeyColumn,
            ArrayFields = ArrayFields,
            RemoveEmpty = RemoveEmpty,
            ExcludeCols = ExcludeCols,
            OutputNames = OutputName != null
                ? new Dictionary<string, string> { ["*"] = OutputName }
                : [],
        };
    }
}

/// <summary>文件级规则</summary>
internal class FileRule
{
    public string? KeyColumn { get; set; }
    public List<string> ArrayFields { get; set; } = [];
    public bool RemoveEmpty { get; set; } = true;
    public List<string> ExcludeCols { get; set; } = [];
    public Dictionary<string, string> OutputNames { get; set; } = [];
    public Dictionary<string, SheetRule>? Sheets { get; set; }
}