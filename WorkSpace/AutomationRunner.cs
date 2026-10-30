using System.IO;
using System.Text.Json;
using ClosedXML.Excel;

namespace ExcelToJson;

public sealed record SyncResult(int FileCount, string BackupDirectory, IReadOnlyList<string> Warnings);

public static class AutomationRunner
{
    public static async Task<SyncResult> RunAsync(ConversionConfig config, IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var svnDir = DirectoryPath(config.SvnDirectory, "SVN 目录", mustExist: true);
        var excelDir = DirectoryPath(config.ExcelDirectory, "Excel 目录", mustExist: true);
        var outputDir = DirectoryPath(config.OutputDirectory, "输出目录", mustExist: false);
        if (!Within(excelDir, svnDir)) throw new InvalidOperationException("Excel 目录必须位于 SVN 更新目录内。");
        var svn = new SvnClient(SvnClient.FindExecutable(config.SvnExecutable));
        progress?.Report("[1/3] 更新 SVN...");
        await svn.UpdateAsync(svnDir, progress, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        progress?.Report("[2/3] 检查工作簿与输出规则...");
        var copyOutputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var warnings = ValidateInputs(config, excelDir, copyOutputs);
        foreach (var warning in warnings) progress?.Report("[警告] " + warning);
        string? targetDir = null;
        if (copyOutputs.Count > 0)
        {
            targetDir = DirectoryPath(config.TargetDirectory, "Cocos 目录", mustExist: true);
            if (Within(targetDir, excelDir) || Within(excelDir, targetDir) || Within(targetDir, outputDir) || Within(outputDir, targetDir))
                throw new InvalidOperationException("Cocos 目录不能与 Excel 或输出目录重叠。");
        }
        Directory.CreateDirectory(outputDir);
        var runDir = Path.Combine(outputDir, ".sync-runs", DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N"));
        var generated = Path.Combine(runDir, "generated");
        Directory.CreateDirectory(generated);
        var conversion = ExcelConverter.Run(excelDir, generated, config, progress);
        if (!conversion.Success) throw new InvalidOperationException("Excel 解析失败，未复制到 Cocos。\n" + conversion.Error);
        cancellationToken.ThrowIfCancellationRequested();
        var files = Directory.GetFiles(generated, "*.json").OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase).ToArray();
        if (files.Length == 0) throw new InvalidOperationException("本次未生成 JSON，请检查工作表勾选和解析范围；未覆盖任何文件。");
        foreach (var file in files)
        {
            using var reader = new StreamReader(file);
            using var json = JsonDocument.Parse(reader.ReadToEnd());
        }
        var copyCount = files.Count(file => copyOutputs.Contains(Path.GetFileName(file)));
        progress?.Report($"[3/3] 更新 {files.Length} 个输出 JSON，复制 {copyCount} 个到 Cocos...");
        Publish(files, outputDir, targetDir, copyOutputs, runDir, progress, cancellationToken);
        progress?.Report(copyCount == 0
            ? $"生成完成：本次没有需要复制的 JSON，未复制到 Cocos。备份：{runDir}"
            : $"同步完成：{targetDir}（复制 {copyCount} 个 JSON）。备份：{runDir}");
        return new SyncResult(copyCount, runDir, warnings);
    }

    private static List<string> ValidateInputs(ConversionConfig config, string excelDir, ISet<string> copyOutputs)
    {
        if (config.ExcelFiles.Count == 0) throw new InvalidOperationException("未选择任何 Excel 文件。");
        var outputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var warnings = new List<string>();
        var selected = 0;
        foreach (var file in config.ExcelFiles)
        {
            var path = Path.GetFullPath(Path.Combine(excelDir, file));
            if (!Within(path, excelDir) || !File.Exists(path)) throw new InvalidOperationException($"Excel 文件不存在或超出 Excel 目录：{file}");
            if (!config.FileRules.TryGetValue(file, out var rule)) continue;
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var book = new XLWorkbook(stream);
            foreach (var (name, sheet) in rule.Sheets.Where(s => s.Value.Enabled && !config.ExcludeSheets.Contains(s.Key)))
            {
                // Preserve the converter's existing behavior for obsolete sheet rules.
                // Only fresh generated files are published, so their old target JSON stays intact.
                if (!book.TryGetWorksheet(name, out _))
                {
                    warnings.Add($"跳过不存在的工作表：{file} / {name}；保留目标中对应的旧 JSON。");
                    continue;
                }
                selected++;
                var names = sheet.SubTables.Count == 0
                    ? new[] { sheet.OutputName ?? $"{Path.GetFileNameWithoutExtension(file)}_{name}.json" }
                    : sheet.SubTables.Select(s => s.Value.OutputName ?? sheet.OutputName ?? $"{Path.GetFileNameWithoutExtension(file)}_{name}.json").Distinct(StringComparer.Ordinal).ToArray();
                foreach (var output in names)
                {
                    if (string.IsNullOrWhiteSpace(output) || Path.GetFileName(output) != output ||
                        output.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                        !output.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException($"输出文件名必须是当前目录内的 JSON 文件名：{output}");
                    if (!outputs.Add(output)) throw new InvalidOperationException($"多个工作表输出同名文件，可能丢失数据：{output}");
                    if (sheet.CopyToTarget) copyOutputs.Add(output);
                }
            }
        }
        if (selected == 0) throw new InvalidOperationException("没有勾选可转换的工作表。");
        return warnings;
    }

    private sealed record CopyEntry(string Source, string Destination, string? Backup);

    private static void Publish(string[] files, string outputDir, string? targetDir, ISet<string> copyOutputs, string runDir,
        IProgress<string>? progress, CancellationToken token)
    {
        var backups = Path.Combine(runDir, "backups");
        Directory.CreateDirectory(backups);
        var entries = new List<CopyEntry>();
        foreach (var directory in new[] { outputDir, targetDir }.OfType<string>())
        foreach (var file in files)
        {
            if (directory == targetDir && !copyOutputs.Contains(Path.GetFileName(file))) continue;
            token.ThrowIfCancellationRequested();
            var destination = Path.Combine(directory, Path.GetFileName(file));
            if (Directory.Exists(destination)) throw new IOException($"目标是文件夹，无法覆盖：{destination}");
            string? backup = null;
            if (File.Exists(destination))
            {
                if ((File.GetAttributes(destination) & (FileAttributes.ReadOnly | FileAttributes.ReparsePoint)) != 0)
                    throw new IOException($"目标只读或为链接，无法覆盖：{destination}");
                backup = Path.Combine(backups, entries.Count + ".json");
                File.Copy(destination, backup);
            }
            entries.Add(new(file, destination, backup));
        }
        File.WriteAllText(Path.Combine(runDir, "manifest.json"), JsonSerializer.Serialize(entries, new JsonSerializerOptions { WriteIndented = true }));
        var written = new List<CopyEntry>();
        try
        {
            foreach (var entry in entries)
            {
                token.ThrowIfCancellationRequested();
                ReplaceFile(entry.Source, entry.Destination);
                written.Add(entry);
                progress?.Report("已复制：" + entry.Destination);
            }
        }
        catch (Exception error)
        {
            var rollbackErrors = new List<string>();
            foreach (var entry in written.AsEnumerable().Reverse())
            {
                try
                {
                    if (entry.Backup == null) File.Delete(entry.Destination);
                    else ReplaceFile(entry.Backup, entry.Destination);
                }
                catch (Exception ex) { rollbackErrors.Add(entry.Destination + "：" + ex.Message); }
            }
            var detail = rollbackErrors.Count == 0 ? "已回滚本次复制。" : "以下文件回滚失败：\n" + string.Join("\n", rollbackErrors);
            throw new IOException($"同步未完成：{error.Message}\n{detail}\n恢复备份：{runDir}", error);
        }
    }

    private static void ReplaceFile(string source, string destination)
    {
        var temporary = Path.Combine(Path.GetDirectoryName(destination)!, ".codex-sync-" + Guid.NewGuid().ToString("N") + ".tmp");
        try { File.Copy(source, temporary); File.Move(temporary, destination, overwrite: true); }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    private static string DirectoryPath(string? value, string label, bool mustExist)
    {
        if (string.IsNullOrWhiteSpace(value) || !Path.IsPathFullyQualified(value)) throw new InvalidOperationException($"请填写{label}的完整路径。");
        var path = Path.TrimEndingDirectorySeparator(Path.GetFullPath(value));
        if (mustExist && !Directory.Exists(path)) throw new DirectoryNotFoundException($"{label}不存在：{path}");
        return path;
    }

    private static bool Within(string path, string directory) => path.Equals(directory, StringComparison.OrdinalIgnoreCase)
        || path.StartsWith(Path.TrimEndingDirectorySeparator(directory) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
}
