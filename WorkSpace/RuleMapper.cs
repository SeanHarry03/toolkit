namespace ExcelToJson;

/// <summary>DTO → 运行时 SheetRule / SubTableRule</summary>
public static class RuleMapper
{
    public static SheetRule? ToSheetRule(SheetRuleDto? dto, Action<string>? log = null)
    {
        if (dto == null)
            return null;

        var rule = new SheetRule
        {
            ExcludeCols = dto.ExcludeCols?.ToList() ?? [],
            CustomHeaders = dto.CustomHeaders.Count > 0
                ? new Dictionary<string, string>(dto.CustomHeaders, StringComparer.OrdinalIgnoreCase)
                : null,
            KeyColumn = string.IsNullOrWhiteSpace(dto.KeyColumn) ? null : dto.KeyColumn,
            ArrayFields = dto.ArrayFields?.ToList() ?? [],
            RemoveEmpty = dto.RemoveEmpty,
            OutputName = string.IsNullOrWhiteSpace(dto.OutputName) ? null : dto.OutputName,
            RoundFields = dto.RoundFields?.ToList() ?? [],
            NoKeyOutputMode = dto.NoKeyOutputMode,
            ComputedColumns = CompileComputed(dto.ComputedColumns, log),
        };

        if (dto.RowStart.HasValue || dto.RowEnd.HasValue)
        {
            rule.RowRange = new RowRangeDef(dto.RowStart ?? 1, dto.RowEnd ?? dto.RowStart ?? 1);
        }

        if (!string.IsNullOrWhiteSpace(dto.ColStart) && !string.IsNullOrWhiteSpace(dto.ColEnd))
        {
            rule.ColRange = new ColRangeDef(dto.ColStart.Trim().ToUpperInvariant(), dto.ColEnd.Trim().ToUpperInvariant());
        }

        if (dto.SubTables.Count > 0)
        {
            rule.SubTables = new Dictionary<string, SubTableRule>(StringComparer.Ordinal);
            foreach (var (key, subDto) in dto.SubTables)
                rule.SubTables[key] = ToSubTableRule(subDto, log);
        }

        return rule;
    }

    public static SubTableRule ToSubTableRule(SubTableRuleDto dto, Action<string>? log = null)
    {
        var rule = new SubTableRule
        {
            ExcludeCols = dto.ExcludeCols?.ToList() ?? [],
            CustomHeaders = dto.CustomHeaders.Count > 0
                ? new Dictionary<string, string>(dto.CustomHeaders, StringComparer.OrdinalIgnoreCase)
                : null,
            KeyColumn = string.IsNullOrWhiteSpace(dto.KeyColumn) ? null : dto.KeyColumn,
            ArrayFields = dto.ArrayFields?.ToList() ?? [],
            RemoveEmpty = dto.RemoveEmpty,
            OutputName = string.IsNullOrWhiteSpace(dto.OutputName) ? null : dto.OutputName,
            RoundFields = dto.RoundFields?.ToList() ?? [],
            NoKeyOutputMode = dto.NoKeyOutputMode,
            WrapWithSubKey = dto.WrapWithSubKey,
            ComputedColumns = CompileComputed(dto.ComputedColumns, log),
        };

        if (dto.RowStart.HasValue || dto.RowEnd.HasValue)
            rule.RowRange = new RowRangeDef(dto.RowStart ?? 1, dto.RowEnd ?? dto.RowStart ?? 1);

        if (!string.IsNullOrWhiteSpace(dto.ColStart) && !string.IsNullOrWhiteSpace(dto.ColEnd))
            rule.ColRange = new ColRangeDef(dto.ColStart.Trim().ToUpperInvariant(), dto.ColEnd.Trim().ToUpperInvariant());

        return rule;
    }

    private static Dictionary<string, Func<IReadOnlyDictionary<string, object?>, object?>>? CompileComputed(
        Dictionary<string, string>? expressions,
        Action<string>? log)
    {
        if (expressions == null || expressions.Count == 0)
            return null;

        var result = new Dictionary<string, Func<IReadOnlyDictionary<string, object?>, object?>>(StringComparer.Ordinal);
        foreach (var (name, expr) in expressions)
        {
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(expr))
                continue;

            if (ExpressionEvaluator.TryCompile(expr, out var func, out var error) && func != null)
            {
                result[name] = func;
            }
            else
            {
                log?.Invoke($"  [警告] 计算列 '{name}' 表达式无效: {error} ({expr})");
            }
        }

        return result.Count > 0 ? result : null;
    }
}
