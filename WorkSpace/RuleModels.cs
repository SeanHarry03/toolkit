using System.Text.Json.Serialization;

namespace ExcelToJson;

/// <summary>完整转换配置（持久化为 rules.json）</summary>
public sealed class ConversionConfig
{
    /// <summary>忽略旧结束行，向下读取到当前区域的第一个空行。</summary>
    public bool AutoDetectRows { get; set; } = true;
    public string? ExcelDirectory { get; set; }
    public string? OutputDirectory { get; set; }
    public List<string> ExcludeSheets { get; set; } = [];
    public List<string> ExcelFiles { get; set; } = [];
    public Dictionary<string, FileRuleDto> FileRules { get; set; } = new(StringComparer.Ordinal);
}

public sealed class FileRuleDto
{
    public Dictionary<string, SheetRuleDto> Sheets { get; set; } = new(StringComparer.Ordinal);
}

/// <summary>Sheet 级可编辑规则</summary>
public sealed class SheetRuleDto
{
    /// <summary>是否导出该 Sheet</summary>
    public bool Enabled { get; set; } = true;

    public int? RowStart { get; set; }
    public int? RowEnd { get; set; }
    /// <summary>普通工作表回写结束行后仍使用首行表头；其他自定义区域设置优先。</summary>
    public bool UseFirstRowAsHeaders { get; set; }
    public string? ColStart { get; set; }
    public string? ColEnd { get; set; }
    public List<string> ExcludeCols { get; set; } = [];
    public Dictionary<string, string> CustomHeaders { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public string? KeyColumn { get; set; }
    public List<string> ArrayFields { get; set; } = [];
    public bool RemoveEmpty { get; set; } = true;
    public string? OutputName { get; set; }
    public List<string> RoundFields { get; set; } = [];
    public NoKeyOutputMode NoKeyOutputMode { get; set; } = NoKeyOutputMode.Array;

    /// <summary>字段名 → 表达式，如 spiritType * 100 + level</summary>
    public Dictionary<string, string> ComputedColumns { get; set; } = new(StringComparer.Ordinal);

    public Dictionary<string, SubTableRuleDto> SubTables { get; set; } = new(StringComparer.Ordinal);

    [JsonIgnore]
    public bool HasCustomRegion =>
        RowStart.HasValue || (RowEnd.HasValue && !UseFirstRowAsHeaders) ||
        !string.IsNullOrWhiteSpace(ColStart) || !string.IsNullOrWhiteSpace(ColEnd) ||
        CustomHeaders.Count > 0 || ComputedColumns.Count > 0 || SubTables.Count > 0 ||
        !string.IsNullOrWhiteSpace(OutputName) || !string.IsNullOrWhiteSpace(KeyColumn) ||
        ExcludeCols.Count > 0 || ArrayFields.Count > 0 || RoundFields.Count > 0;
}

public sealed class SubTableRuleDto
{
    public int? RowStart { get; set; }
    public int? RowEnd { get; set; }
    public string? ColStart { get; set; }
    public string? ColEnd { get; set; }
    public List<string> ExcludeCols { get; set; } = [];
    public Dictionary<string, string> CustomHeaders { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public string? KeyColumn { get; set; }
    public List<string> ArrayFields { get; set; } = [];
    public bool RemoveEmpty { get; set; } = true;
    public string? OutputName { get; set; }
    public List<string> RoundFields { get; set; } = [];
    public NoKeyOutputMode NoKeyOutputMode { get; set; } = NoKeyOutputMode.Array;
    public Dictionary<string, string> ComputedColumns { get; set; } = new(StringComparer.Ordinal);
    public bool WrapWithSubKey { get; set; } = true;
}

/// <summary>UI 绑定用：字典键值对</summary>
public sealed class DictEntry
{
    public string Key { get; set; } = "";
    public string Value { get; set; } = "";
}
