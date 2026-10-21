using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;

namespace ExcelToJson;

public static class RulesStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.Create(
            UnicodeRanges.BasicLatin,
            UnicodeRanges.CjkUnifiedIdeographs,
            UnicodeRanges.CjkSymbolsandPunctuation,
            UnicodeRanges.HalfwidthandFullwidthForms,
            UnicodeRanges.CjkCompatibility),
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    public static string DefaultPath =>
        Path.Combine(AppContext.BaseDirectory, "rules.json");

    public static ConversionConfig LoadOrCreate(string? path = null)
    {
        path ??= DefaultPath;
        if (File.Exists(path))
        {
            var json = File.ReadAllText(path, System.Text.Encoding.UTF8);
            var loaded = JsonSerializer.Deserialize<ConversionConfig>(json, JsonOptions);
            if (loaded != null)
            {
                Normalize(loaded);
                return loaded;
            }
        }

        var created = BuiltInRules.Create();
        Save(created, path);
        return created;
    }

    public static void Save(ConversionConfig config, string? path = null)
    {
        path ??= DefaultPath;
        Normalize(config);
        var json = JsonSerializer.Serialize(config, JsonOptions);
        File.WriteAllText(path, json, System.Text.Encoding.UTF8);
    }

    public static ConversionConfig ResetToDefault(string? path = null)
    {
        var created = BuiltInRules.Create();
        Save(created, path);
        return created;
    }

    private static void Normalize(ConversionConfig config)
    {
        config.ExcludeSheets ??= [];
        config.ExcelFiles ??= [];
        config.FileRules ??= new Dictionary<string, FileRuleDto>(StringComparer.Ordinal);

        foreach (var file in config.FileRules.Values)
        {
            file.Sheets ??= new Dictionary<string, SheetRuleDto>(StringComparer.Ordinal);
            foreach (var sheet in file.Sheets.Values)
            {
                sheet.ExcludeCols ??= [];
                sheet.CustomHeaders ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                sheet.ArrayFields ??= [];
                sheet.RoundFields ??= [];
                sheet.ComputedColumns ??= new Dictionary<string, string>(StringComparer.Ordinal);
                sheet.SubTables ??= new Dictionary<string, SubTableRuleDto>(StringComparer.Ordinal);
                foreach (var sub in sheet.SubTables.Values)
                {
                    sub.ExcludeCols ??= [];
                    sub.CustomHeaders ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    sub.ArrayFields ??= [];
                    sub.RoundFields ??= [];
                    sub.ComputedColumns ??= new Dictionary<string, string>(StringComparer.Ordinal);
                }
            }
        }
    }
}
