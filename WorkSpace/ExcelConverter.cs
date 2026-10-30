using System.Globalization;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using ClosedXML.Excel;

namespace ExcelToJson;

/// <summary>
/// Excel → JSON 转换工具（规则与原 convert.js / Program.cs 一致）
/// </summary>
public static class ExcelConverter
{
    public sealed class ConvertResult
    {
        public bool Success { get; set; }
        public string OutputDir { get; init; } = "";
        public List<string> Messages { get; } = new();
        public string? Error { get; set; }
    }

    public static string ResolveExcelDirectory(string? startDir = null)
    {
        var candidates = new List<string>();
        string? dir = startDir ?? Directory.GetCurrentDirectory();
        candidates.Add(dir);
        for (int i = 0; i < 5; i++)
        {
            dir = Path.GetDirectoryName(dir);
            if (dir != null)
                candidates.Add(dir);
        }

        foreach (var candidate in candidates)
        {
            if (Directory.Exists(candidate) && Directory.EnumerateFiles(candidate, "*.xlsx").Any())
                return candidate;
        }

        return startDir ?? Directory.GetCurrentDirectory();
    }

    public static ConvertResult Run(
        string excelDir,
        string outputDir,
        ConversionConfig config,
        IProgress<string>? progress = null)
    {
        var result = new ConvertResult { OutputDir = outputDir };
        void Log(string msg)
        {
            result.Messages.Add(msg);
            progress?.Report(msg);
        }

        try
        {
            Directory.CreateDirectory(outputDir);
            Log($"Excel 目录: {excelDir}");
            Log($"输出目录: {outputDir}");
            Log("");

            var excludeSheets = new HashSet<string>(config.ExcludeSheets ?? [], StringComparer.Ordinal);
            var excelFiles = config.ExcelFiles?.Count > 0
                ? config.ExcelFiles
                : [];
            var fileRulesDto = config.FileRules ?? new Dictionary<string, FileRuleDto>();

            var defaultRule = new FileRule
            {
                KeyColumn = null,
                ArrayFields = [],
                RemoveEmpty = true,
                ExcludeCols = [],
                OutputNames = [],
            };

            var jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.Create(
                    UnicodeRanges.BasicLatin,
                    UnicodeRanges.CjkUnifiedIdeographs,
                    UnicodeRanges.CjkSymbolsandPunctuation,
                    UnicodeRanges.HalfwidthandFullwidthForms,
                    UnicodeRanges.CjkCompatibility
                ),
            };

            foreach (var fileName in excelFiles)
            {
                var filePath = Path.Combine(excelDir, fileName);

                if (!File.Exists(filePath))
                {
                    Log($"[跳过] 文件不存在: {fileName}");
                    continue;
                }

                Log($"[转换] {fileName} ...");

                fileRulesDto.TryGetValue(fileName, out var fileRuleDto);
                var sheetDtos = fileRuleDto?.Sheets ?? new Dictionary<string, SheetRuleDto>();

                using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workbook = new XLWorkbook(fileStream);

                foreach (var worksheet in workbook.Worksheets)
                {
                    var sheetName = worksheet.Name;
                    if (excludeSheets.Contains(sheetName))
                    {
                        Log($"  [排除] {sheetName}");
                        continue;
                    }

                    if (!sheetDtos.TryGetValue(sheetName, out var sheetDto) || !sheetDto.Enabled)
                    {
                        Log($"  [跳过未勾选] {sheetName}");
                        continue;
                    }

                    if (config.AutoDetectRows)
                        UpdateRowEnds(worksheet, sheetDto, Log);

                    SheetRule? sheetRule = sheetDto.HasCustomRegion
                        ? RuleMapper.ToSheetRule(sheetDto, Log)
                        : null;

                    List<Dictionary<string, object?>> processedRows;
                    FileRule finalRule;

                    if (sheetRule != null)
                    {
                        // ======== 多子表模式：一个 Sheet 内多个独立区域 ========
                        if (sheetRule.SubTables != null)
                        {
                            // 按 OutputName 分组：相同 OutputName 的子表合并输出到同一个文件
                            var outputGroups = new Dictionary<string, Dictionary<string, object>>();
                            var outputCounts = new Dictionary<string, int>();
                            var directOutputs = new Dictionary<string, object>();
                            var directOutputCounts = new Dictionary<string, int>();

                            foreach (var (subKey, subRule) in sheetRule.SubTables)
                            {
                                var regionData = ProcessRegion(worksheet, subRule);
                                int dataCount = regionData is Dictionary<string, Dictionary<string, object?>> d
                                    ? d.Count
                                    : ((System.Collections.IList)regionData).Count;

                                // 确定输出文件名
                                string outName = subRule.OutputName
                                                 ?? sheetRule.OutputName
                                                 ?? $"{Path.GetFileNameWithoutExtension(fileName)}_{sheetName}.json";

                                if (subRule.WrapWithSubKey)
                                {
                                    if (directOutputs.ContainsKey(outName))
                                    {
                                        throw new InvalidOperationException(
                                            $"不包裹输出的子表不能和其他子表输出到同一个文件: {outName}");
                                    }

                                    if (!outputGroups.ContainsKey(outName))
                                    {
                                        outputGroups[outName] = new Dictionary<string, object>();
                                        outputCounts[outName] = 0;
                                    }

                                    // 用子表 key（如 "601"）作为外层包裹
                                    outputGroups[outName][subKey] = regionData;
                                    outputCounts[outName] += dataCount;
                                }
                                else
                                {
                                    // 不包裹：直接将数据平铺合并
                                    if (regionData is Dictionary<string, Dictionary<string, object?>> dictData)
                                    {
                                        if (directOutputs.ContainsKey(outName))
                                        {
                                            throw new InvalidOperationException(
                                                $"不包裹输出的子表不能和其他子表输出到同一个文件: {outName}");
                                        }

                                        if (!outputGroups.ContainsKey(outName))
                                        {
                                            outputGroups[outName] = new Dictionary<string, object>();
                                            outputCounts[outName] = 0;
                                        }

                                        foreach (var (k, v) in dictData)
                                            outputGroups[outName][k] = v!;

                                        outputCounts[outName] += dataCount;
                                    }
                                    else
                                    {
                                        if (outputGroups.ContainsKey(outName) || directOutputs.ContainsKey(outName))
                                        {
                                            throw new InvalidOperationException(
                                                $"不包裹输出的列表子表不能和其他子表输出到同一个文件: {outName}");
                                        }

                                        // 列表数据不再使用 subKey 包裹，直接作为 JSON 根节点输出
                                        directOutputs[outName] = regionData;
                                        directOutputCounts[outName] = dataCount;
                                    }
                                }
                            }

                            // 输出每个分组文件
                            foreach (var (outName, groupData) in outputGroups)
                            {
                                var outPath = Path.Combine(outputDir, outName);
                                var json = JsonSerializer.Serialize(groupData, jsonOptions);
                                File.WriteAllText(outPath, json, System.Text.Encoding.UTF8);
                                Log($"  -> {outName} ({groupData.Count} 个区域, {outputCounts[outName]} 条数据)");
                            }

                            foreach (var (outName, directData) in directOutputs)
                            {
                                var outPath = Path.Combine(outputDir, outName);
                                var json = JsonSerializer.Serialize(directData, jsonOptions);
                                File.WriteAllText(outPath, json, System.Text.Encoding.UTF8);
                                Log($"  -> {outName} ({directOutputCounts[outName]} 条数据)");
                            }

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
                        var resolvedRound = ResolveRoundFields(sheetRule.RoundFields, sheetRule.CustomHeaders);

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
                            var cleanRow = ProcessRow(newRow, sheetRule.ArrayFields, sheetRule.RemoveEmpty, resolvedRound);
                            if (cleanRow.Count > 0)
                            {
                                processedRows.Add(cleanRow);
                            }
                        }
                    }
                    else
                    {
                        // ======== 默认通用逻辑：第一行是表头 ========
                        finalRule = defaultRule;

                        var lastRow = sheetDto.RowEnd ?? worksheet.LastRowUsed()?.RowNumber() ?? 0;
                        var lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

                        // 自动扫描得到空区域时仍写出空数组，清除上次导出的旧数据。
                        if (lastRow < 2 && !config.AutoDetectRows) continue;

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
                        var resolvedRound = ResolveRoundFields(finalRule.RoundFields, colToKeyMap);

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
                                finalRule.RemoveEmpty,
                                resolvedRound);
                            processedRows.Add(cleanRow);
                        }
                    }

                    // 是否按 key 字段包裹外层；未指定 key 时可配置输出数组或自然序号对象
                    var noKeyOutputMode = sheetRule?.NoKeyOutputMode ?? finalRule.NoKeyOutputMode;
                    object outputData = BuildRowsOutput(processedRows, finalRule.KeyColumn, noKeyOutputMode);

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

                    Log($"  -> {outputFileName} ({count} 条数据)");
                }

                Log("");
            }

            Log($"✅ 全部完成! JSON 文件已输出到: {outputDir}");

            result.Success = true;
            return result;
        }
        catch (Exception ex)
        {
            Log($"转换失败: {ex.Message}");
            result.Success = false;
            result.Error = ex.ToString();
            return result;
        }
    }

    /// <summary>扫描已启用工作表的实际结束行；扫描和转换共用同一规则。</summary>
    public static void UpdateRowEnds(IXLWorksheet worksheet, SheetRuleDto sheet, Action<string>? log = null)
    {
        if (!sheet.Enabled) return;

        if (sheet.SubTables.Count > 0)
        {
            foreach (var (key, sub) in sheet.SubTables)
            {
                var rule = RuleMapper.ToSubTableRule(sub);
                sub.RowEnd = FindDataEndRow(worksheet, sub.RowStart ?? 1, rule.ColRange, rule.ExcludeCols);
                log?.Invoke($"  [自动行数] {worksheet.Name}/{key}: 结束行 {sub.RowEnd}");
            }
        }
        else
        {
            bool useHeaders = !sheet.HasCustomRegion;
            var rule = RuleMapper.ToSheetRule(sheet)!;
            int startRow = useHeaders ? 2 : sheet.RowStart ?? 1;
            sheet.RowEnd = FindDataEndRow(worksheet, startRow, rule.ColRange, rule.ExcludeCols);
            sheet.UseFirstRowAsHeaders = useHeaders;
            log?.Invoke($"  [自动行数] {worksheet.Name}: 结束行 {sheet.RowEnd}");
        }
    }

    private static int FindDataEndRow(IXLWorksheet worksheet, int startRow,
        ColRangeDef? colRange, List<string> excludeCols)
    {
        int startCol = colRange.HasValue ? ColLetterToIndex(colRange.Value.StartCol) + 1 : 1;
        int endCol = colRange.HasValue ? ColLetterToIndex(colRange.Value.EndCol) + 1
            : worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        if (startRow < 1 || startRow > 1048576 || startCol < 1 || endCol > 16384 || startCol > endCol)
            throw new InvalidOperationException($"{worksheet.Name}: 起始行或列范围无效。");

        var excluded = new HashSet<string>(excludeCols, StringComparer.OrdinalIgnoreCase);
        var columns = Enumerable.Range(startCol, endCol - startCol + 1)
            .Where(col => !excluded.Contains(IndexToColLetter(col - 1))).ToArray();
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int endRow = startRow - 1; // 起始行即为空时，表示该区域没有数据。
        for (int row = startRow; row <= lastRow; row++)
        {
            bool hasData = false;
            foreach (int col in columns)
            {
                var value = worksheet.Cell(row, col).Value;
                if (!value.IsBlank && !(value.IsText && string.IsNullOrWhiteSpace(value.GetText())))
                {
                    hasData = true;
                    break;
                }
            }
            if (!hasData) break;
            endRow = row;
        }
        return endRow;
    }

    private static int ColLetterToIndex(string col)
    {
        int index = 0;
        foreach (char c in col)
        {
            index = index * 26 + (c - 'A' + 1);
        }

        return index - 1;
    }

    // 0-based 列索引转字母列名（0 → A, 1 → B, 25 → Z, 26 → AA）
    private static string IndexToColLetter(int index)
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
    private static double FixFloat(double value)
    {
        // 如果离整数非常近（1e-9 以内），直接取整
        double rounded = Math.Round(value);
        if (Math.Abs(value - rounded) < 1e-9)
            return rounded;

        // 否则清理到 9 位小数，去除浮点噪声
        return Math.Round(value, 9);
    }

    // 从行数据中安全获取 double 值（用于计算列）
    private static double GetDouble(IReadOnlyDictionary<string, object?> row, string key)
        => ExpressionEvaluator.GetDouble(row, key);

    // 从单元格获取原始值
    private static object? GetCellValue(IXLCell cell)
    {
        if (cell.IsEmpty() || cell.Value.IsBlank)
            return null;

        if (cell.Value.IsNumber)
            return FixFloat(cell.Value.GetNumber());

        if (cell.Value.IsBoolean)
            return cell.Value.GetBoolean();

        // 文本或其他类型：尝试转数字，转换不了就保留原文本
        var text = cell.Value.GetText()?.Replace("\\n", "\n");
        if (text != null && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double num))
            return FixFloat(num);
        return text;
    }

    // 解析取整字段映射，支持配置列字母或属性键名，或者用 "*" 取整所有列
    private static HashSet<string> ResolveRoundFields(List<string>? roundFields, Dictionary<string, string>? customHeaders)
    {
        var resolved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (roundFields == null)
            return resolved;

        foreach (var field in roundFields)
        {
            resolved.Add(field);
            if (customHeaders != null && customHeaders.TryGetValue(field, out var mappedHeader))
            {
                resolved.Add(mappedHeader);
            }
        }

        return resolved;
    }

    // 处理单行数据：去空、转数组、数值取整
    private static Dictionary<string, object?> ProcessRow(
        Dictionary<string, object?> row,
        List<string>? arrayFields,
        bool removeEmpty,
        HashSet<string>? resolvedRoundFields = null)
    {
        var processed = new Dictionary<string, object?>();
        bool shouldRoundAll = resolvedRoundFields != null && resolvedRoundFields.Contains("*");

        foreach (var (key, value) in row)
        {
            // 跳过空值
            if (removeEmpty && (value == null || (value is string s && string.IsNullOrEmpty(s))))
                continue;

            bool shouldRound = shouldRoundAll || (resolvedRoundFields != null && resolvedRoundFields.Contains(key));

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
                            {
                                var fixedNum = FixFloat(num);
                                return shouldRound ? (object)Convert.ToInt64(Math.Round(fixedNum)) : (object)fixedNum;
                            }

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
                        {
                            list.Add(shouldRound ? Convert.ToInt64(Math.Round(FixFloat(dVal))) : FixFloat(dVal));
                        }
                        else if (double.TryParse(value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture,
                                     out double num))
                        {
                            list.Add(shouldRound ? Convert.ToInt64(Math.Round(FixFloat(num))) : FixFloat(num));
                        }
                        else
                        {
                            list.Add(value);
                        }
                    }

                    processed[key] = list;
                }
            }
            else
            {
                if (shouldRound && value is double dVal)
                {
                    processed[key] = Convert.ToInt64(Math.Round(dVal));
                }
                else
                {
                    processed[key] = value;
                }
            }
        }

        return processed;
    }

    // 按 key 字段包裹数组为字典
    private static Dictionary<string, Dictionary<string, object?>> WrapWithKey(
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

    // 按自然序号 1、2、3 包裹数组为字典
    private static Dictionary<string, Dictionary<string, object?>> WrapWithRowIndex(
        List<Dictionary<string, object?>> rows)
    {
        var result = new Dictionary<string, Dictionary<string, object?>>();
        for (int i = 0; i < rows.Count; i++)
        {
            result[(i + 1).ToString(CultureInfo.InvariantCulture)] = rows[i];
        }

        return result;
    }

    // 根据 KeyColumn 和无 KeyColumn 输出规则构造最终数据
    private static object BuildRowsOutput(
        List<Dictionary<string, object?>> rows,
        string? keyColumn,
        NoKeyOutputMode noKeyOutputMode)
    {
        if (!string.IsNullOrEmpty(keyColumn))
            return WrapWithKey(rows, keyColumn!);

        return noKeyOutputMode == NoKeyOutputMode.RowIndexObject
            ? WrapWithRowIndex(rows)
            : rows;
    }

    // 处理一个子表区域，返回按规则转换后的数据（数组或 key 包裹的字典）
    private static object ProcessRegion(
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

            var resolvedRound = ResolveRoundFields(rule.RoundFields, rule.CustomHeaders);
            var cleanRow = ProcessRow(newRow, rule.ArrayFields, rule.RemoveEmpty, resolvedRound);
            if (cleanRow.Count > 0)
                rows.Add(cleanRow);
        }

        return BuildRowsOutput(rows, rule.KeyColumn, rule.NoKeyOutputMode);
    }
}

/// <summary>列范围</summary>
public record struct ColRangeDef(string StartCol, string EndCol);

/// <summary>行范围（1-indexed，对应 Excel 物理行号）</summary>
public record struct RowRangeDef(int Start, int End);

/// <summary>未指定 KeyColumn 时的输出模式</summary>
public enum NoKeyOutputMode
{
    /// <summary>输出数组：[{}, {}]</summary>
    Array,

    /// <summary>输出自然序号对象：{"1": {}, "2": {}}</summary>
    RowIndexObject,
}

/// <summary>子表规则（一个 Sheet 内多个独立区域的配置）</summary>
public class SubTableRule
{
    public RowRangeDef? RowRange { get; set; }
    public ColRangeDef? ColRange { get; set; }
    public List<string> ExcludeCols { get; set; } = [];
    public Dictionary<string, string>? CustomHeaders { get; set; }
    public string? KeyColumn { get; set; }
    public List<string> ArrayFields { get; set; } = [];
    public bool RemoveEmpty { get; set; } = true;
    public Dictionary<string, Func<IReadOnlyDictionary<string, object?>, object?>>? ComputedColumns { get; set; }
    public List<string> RoundFields { get; set; } = [];
    public NoKeyOutputMode NoKeyOutputMode { get; set; } = NoKeyOutputMode.Array;

    /// <summary>输出文件名（相同 OutputName 的子表合并到同一个文件）</summary>
    public string? OutputName { get; set; }

    /// <summary>是否用子表 key 作为外层包裹（默认 true）。设为 false 则数据直接平铺</summary>
    public bool WrapWithSubKey { get; set; } = true;
}

/// <summary>Sheet 级规则</summary>
public class SheetRule
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
    public List<string> RoundFields { get; set; } = [];
    public NoKeyOutputMode NoKeyOutputMode { get; set; } = NoKeyOutputMode.Array;

    public FileRule ToFileRule()
    {
        return new FileRule
        {
            KeyColumn = KeyColumn,
            ArrayFields = ArrayFields,
            RemoveEmpty = RemoveEmpty,
            ExcludeCols = ExcludeCols,
            RoundFields = RoundFields,
            NoKeyOutputMode = NoKeyOutputMode,
            OutputNames = OutputName != null
                ? new Dictionary<string, string> { ["*"] = OutputName }
                : [],
        };
    }
}

/// <summary>文件级规则</summary>
public class FileRule
{
    public string? KeyColumn { get; set; }
    public List<string> ArrayFields { get; set; } = [];
    public bool RemoveEmpty { get; set; } = true;
    public List<string> ExcludeCols { get; set; } = [];
    public Dictionary<string, string> OutputNames { get; set; } = [];
    public Dictionary<string, SheetRule>? Sheets { get; set; }
    public List<string> RoundFields { get; set; } = [];
    public NoKeyOutputMode NoKeyOutputMode { get; set; } = NoKeyOutputMode.Array;
}
