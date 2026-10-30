using System.Diagnostics;
using System.IO;
using System.Text;
using System.Xml.Linq;

namespace ExcelToJson;

public sealed class SvnClient(string executable)
{
    public static string FindExecutable(string? configured)
    {
        if (!string.IsNullOrWhiteSpace(configured))
        {
            if (Path.IsPathFullyQualified(configured) && File.Exists(configured)) return configured;
            throw new FileNotFoundException("配置的 SVN 程序不存在，请选择 svn.exe。", configured);
        }
        var candidates = new List<string>
        {
            Path.Combine(AppContext.BaseDirectory, "tools", "svn", "bin", "svn.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "TortoiseSVN", "bin", "svn.exe")
        };
        candidates.AddRange((Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator)
            .Where(p => !string.IsNullOrWhiteSpace(p)).Select(p => Path.Combine(p.Trim('"'), "svn.exe")));
        return candidates.FirstOrDefault(File.Exists)
            ?? throw new FileNotFoundException("未找到 svn.exe。请保留工具附带的 tools 文件夹，或在设置中选择 SVN 程序。");
    }

    public async Task UpdateAsync(string directory, IProgress<string>? progress, CancellationToken token)
    {
        // A child of a working-copy root is valid; do not require a .svn inside this folder.
        await ExecuteAsync(["info", "--xml", "--non-interactive", "--", directory + "@"], null, token);
        await CheckConflictsAsync(directory, token);
        await ExecuteAsync(["update", "--non-interactive", "--accept", "postpone", "--", directory + "@"], progress, token);
        // svn update may return success while recording text, property or tree conflicts.
        await CheckConflictsAsync(directory, token);
    }

    private async Task CheckConflictsAsync(string directory, CancellationToken token)
    {
        var xml = await ExecuteAsync(["status", "--xml", "--non-interactive", "--", directory + "@"], null, token);
        var conflicts = XDocument.Parse(xml).Descendants("entry").Where(e =>
        {
            var status = e.Element("wc-status");
            return (string?)status?.Attribute("item") == "conflicted"
                || (string?)status?.Attribute("props") == "conflicted"
                || (string?)status?.Attribute("tree-conflicted") == "true";
        }).Select(e => (string?)e.Attribute("path")).ToArray();
        if (conflicts.Length > 0)
            throw new InvalidOperationException("SVN 存在冲突，请先在 TortoiseSVN 中解决：\n" + string.Join("\n", conflicts));
    }

    private async Task<string> ExecuteAsync(string[] args, IProgress<string>? progress, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var start = new ProcessStartInfo(executable)
        {
            UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardOutput = true, RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
        };
        foreach (var arg in args) start.ArgumentList.Add(arg);
        using var process = Process.Start(start) ?? throw new IOException("无法启动 SVN。");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(token);
        timeout.CancelAfter(TimeSpan.FromMinutes(10));
        var stdout = ReadAsync(process.StandardOutput, progress);
        var stderr = ReadAsync(process.StandardError, progress);
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException)
        {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync(CancellationToken.None);
            await Task.WhenAll(stdout, stderr);
            if (token.IsCancellationRequested) throw;
            throw new TimeoutException("SVN 超过 10 分钟未完成。请检查网络；必要时先在 TortoiseSVN 中清理工作副本。");
        }
        var output = await stdout;
        var error = await stderr;
        if (process.ExitCode != 0)
            throw new InvalidOperationException($"SVN {args[0]} 失败（{process.ExitCode}）：\n{error}\n如需登录，请先使用 TortoiseSVN 更新一次并保存凭据。");
        return output;
    }

    private static async Task<string> ReadAsync(StreamReader reader, IProgress<string>? progress)
    {
        var text = new StringBuilder();
        while (await reader.ReadLineAsync() is { } line)
        { text.AppendLine(line); progress?.Report(line); }
        return text.ToString();
    }
}
