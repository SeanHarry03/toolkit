using System.IO;
using System.Text.Json;

namespace ExcelToJson;

/// <summary>GameTools 文件协议；不改变 GUI 规则或默认复制开关。</summary>
public static class HeadlessRunner
{
    public sealed class Request
    {
        public string Action { get; set; } = "";
        public string SvnDirectory { get; set; } = "";
        public string TargetDirectory { get; set; } = "";
        public string OutputDirectory { get; set; } = "";
        public List<string> SelectedFiles { get; set; } = [];
    }

    private sealed class LogProgress(List<string> log) : IProgress<string>
    {
        public void Report(string value) { lock (log) log.Add(value); }
    }

    public static async Task<int> RunAsync(string[] args)
    {
        if (args.Length != 3) return 2;
        var log = new List<string>();
        object result;
        var exitCode = 0;
        try
        {
            var request = JsonSerializer.Deserialize<Request>(File.ReadAllText(args[1]),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? throw new IOException("请求为空。");
            if (request.Action != "scan" && request.Action != "sync") throw new IOException("未知操作。");
            if (!File.Exists(RulesStore.DefaultPath)) throw new IOException("缺少 rules.json，请先打开转换工具保存规则。");
            if (request.SelectedFiles == null) throw new IOException("缺少文件选择。");
            foreach (var name in request.SelectedFiles)
                if (string.IsNullOrWhiteSpace(name) || Path.GetFileName(name) != name || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || !name.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                    throw new IOException("非法 JSON 文件名：" + name);
            if (request.Action == "sync" && request.SelectedFiles.Count == 0) throw new IOException("请先扫描并勾选要复制的 JSON。");
            var config = RulesStore.LoadOrCreate();
            config.SvnDirectory = request.SvnDirectory;
            config.ExcelDirectory = request.SvnDirectory;
            config.OutputDirectory = request.OutputDirectory;
            config.TargetDirectory = request.TargetDirectory;
            var selected = request.Action == "scan" ? Array.Empty<string>() : request.SelectedFiles.ToArray();
            var sync = await AutomationRunner.RunSelectedAsync(config, new LogProgress(log), selectedFiles: selected, updateSvn: request.Action == "sync");
            var files = Directory.GetFiles(Path.Combine(sync.BackupDirectory, "generated"), "*.json")
                .Select(Path.GetFileName).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToArray();
            result = new { success = true, files, copiedCount = sync.FileCount, backupDirectory = sync.BackupDirectory, warnings = sync.Warnings, log };
        }
        catch (Exception error)
        {
            exitCode = 1;
            result = new { success = false, error = error.Message, log };
        }
        try { File.WriteAllText(args[2], JsonSerializer.Serialize(result)); }
        catch { return 3; }
        return exitCode;
    }
}
