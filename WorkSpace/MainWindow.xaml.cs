using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace ExcelToJson;

public partial class MainWindow : Window
{
    private ConversionConfig _config = null!;
    private bool _isRunning;
    private bool _loadingUi;
    private string? _currentFile;
    private SheetRuleDto? _currentSheet;
    private string? _currentSubKey;
    private SubTableRuleDto? _currentSub;

    private readonly ObservableCollection<SheetListItem> _sheetItems = [];
    private readonly ObservableCollection<DictEntry> _headers = [];
    private readonly ObservableCollection<DictEntry> _computed = [];
    private readonly ObservableCollection<DictEntry> _subHeaders = [];
    private readonly ObservableCollection<DictEntry> _subComputed = [];
    private readonly ObservableCollection<string> _subKeys = [];

    public MainWindow()
    {
        InitializeComponent();
        HeadersGrid.ItemsSource = _headers;
        ComputedGrid.ItemsSource = _computed;
        SubHeadersGrid.ItemsSource = _subHeaders;
        SubComputedGrid.ItemsSource = _subComputed;
        SheetList.ItemsSource = _sheetItems;
        SubTableList.ItemsSource = _subKeys;

        _config = RulesStore.LoadOrCreate();
        LoadDirectories();
        RefreshFileCombo();
        AppendLog($"已加载规则: {RulesStore.DefaultPath}");
        AppendLog("勾选 Sheet、编辑规则后可「保存规则」或直接「开始转换」。");
    }

    private void RefreshFileCombo(string? selectFile = null)
    {
        _loadingUi = true;
        FileCombo.Items.Clear();
        foreach (var f in _config.ExcelFiles)
            FileCombo.Items.Add(f);

        if (FileCombo.Items.Count > 0)
        {
            var target = selectFile ?? _config.ExcelFiles.FirstOrDefault();
            FileCombo.SelectedItem = target != null && FileCombo.Items.Contains(target)
                ? target
                : FileCombo.Items[0];
        }

        _loadingUi = false;
        LoadSheetsForSelectedFile();
    }

    private void LoadSheetsForSelectedFile()
    {
        CommitCurrentEditors();
        _sheetItems.Clear();
        _currentFile = FileCombo.SelectedItem as string;
        ClearRuleEditor();

        if (_currentFile == null)
            return;

        if (!_config.FileRules.TryGetValue(_currentFile, out var fileRule))
        {
            fileRule = new FileRuleDto();
            _config.FileRules[_currentFile] = fileRule;
        }

        foreach (var (name, sheet) in fileRule.Sheets.OrderBy(kv => kv.Key, StringComparer.Ordinal))
        {
            _sheetItems.Add(new SheetListItem(name, sheet));
        }
    }

    private void FileCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loadingUi) return;
        LoadSheetsForSelectedFile();
    }

    private void SheetList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loadingUi) return;
        CommitCurrentEditors();
        if (SheetList.SelectedItem is SheetListItem item)
            LoadSheetEditor(item.Rule, item.Name);
        else
            ClearRuleEditor();
    }

    private void LoadSheetEditor(SheetRuleDto sheet, string name)
    {
        _loadingUi = true;
        _currentSheet = sheet;
        RuleEditorPanel.IsEnabled = true;
        RuleTitle.Text = $"规则: {name}";

        OutputNameBox.Text = sheet.OutputName ?? "";
        KeyColumnBox.Text = sheet.KeyColumn ?? "";
        RowStartBox.Text = sheet.RowStart?.ToString() ?? "";
        RowEndBox.Text = sheet.RowEnd?.ToString() ?? "";
        ColStartBox.Text = sheet.ColStart ?? "";
        ColEndBox.Text = sheet.ColEnd ?? "";
        ExcludeColsBox.Text = string.Join(",", sheet.ExcludeCols);
        ArrayFieldsBox.Text = string.Join(",", sheet.ArrayFields);
        RoundFieldsBox.Text = string.Join(",", sheet.RoundFields);
        RemoveEmptyBox.IsChecked = sheet.RemoveEmpty;
        SelectMode(NoKeyModeBox, sheet.NoKeyOutputMode);

        _headers.Clear();
        foreach (var (k, v) in sheet.CustomHeaders)
            _headers.Add(new DictEntry { Key = k, Value = v });

        _computed.Clear();
        foreach (var (k, v) in sheet.ComputedColumns)
            _computed.Add(new DictEntry { Key = k, Value = v });

        _subKeys.Clear();
        foreach (var key in sheet.SubTables.Keys.OrderBy(k => k, StringComparer.Ordinal))
            _subKeys.Add(key);

        ClearSubEditor();
        _loadingUi = false;
    }

    private void ClearRuleEditor()
    {
        _currentSheet = null;
        RuleEditorPanel.IsEnabled = false;
        RuleTitle.Text = "选择左侧 Sheet 以编辑规则";
        _headers.Clear();
        _computed.Clear();
        _subKeys.Clear();
        ClearSubEditor();
    }

    private void ClearSubEditor()
    {
        _currentSubKey = null;
        _currentSub = null;
        SubTableEditor.IsEnabled = false;
        SubKeyBox.Text = "";
        SubOutputBox.Text = "";
        SubRowStartBox.Text = "";
        SubRowEndBox.Text = "";
        SubColStartBox.Text = "";
        SubColEndBox.Text = "";
        SubExcludeBox.Text = "";
        SubArrayBox.Text = "";
        SubRoundBox.Text = "";
        SubWrapBox.IsChecked = true;
        SubRemoveEmptyBox.IsChecked = true;
        SelectMode(SubNoKeyModeBox, NoKeyOutputMode.Array);
        _subHeaders.Clear();
        _subComputed.Clear();
        SubTableList.SelectedIndex = -1;
    }

    private void SubTableList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loadingUi || _currentSheet == null) return;
        CommitSubEditor();
        if (SubTableList.SelectedItem is string key &&
            _currentSheet.SubTables.TryGetValue(key, out var sub))
        {
            LoadSubEditor(key, sub);
        }
        else
        {
            _currentSubKey = null;
            _currentSub = null;
            SubTableEditor.IsEnabled = false;
        }
    }

    private void LoadSubEditor(string key, SubTableRuleDto sub)
    {
        _loadingUi = true;
        _currentSubKey = key;
        _currentSub = sub;
        SubTableEditor.IsEnabled = true;
        SubKeyBox.Text = key;
        SubOutputBox.Text = sub.OutputName ?? "";
        SubRowStartBox.Text = sub.RowStart?.ToString() ?? "";
        SubRowEndBox.Text = sub.RowEnd?.ToString() ?? "";
        SubColStartBox.Text = sub.ColStart ?? "";
        SubColEndBox.Text = sub.ColEnd ?? "";
        SubExcludeBox.Text = string.Join(",", sub.ExcludeCols);
        SubArrayBox.Text = string.Join(",", sub.ArrayFields);
        SubRoundBox.Text = string.Join(",", sub.RoundFields);
        SubWrapBox.IsChecked = sub.WrapWithSubKey;
        SubRemoveEmptyBox.IsChecked = sub.RemoveEmpty;
        SelectMode(SubNoKeyModeBox, sub.NoKeyOutputMode);

        _subHeaders.Clear();
        foreach (var (k, v) in sub.CustomHeaders)
            _subHeaders.Add(new DictEntry { Key = k, Value = v });

        _subComputed.Clear();
        foreach (var (k, v) in sub.ComputedColumns)
            _subComputed.Add(new DictEntry { Key = k, Value = v });

        _loadingUi = false;
    }

    private void CommitCurrentEditors()
    {
        CommitSheetEditor();
        CommitSubEditor();
    }

    private void CommitSheetEditor()
    {
        if (_currentSheet == null || _loadingUi) return;

        _currentSheet.OutputName = NullIfEmpty(OutputNameBox.Text);
        _currentSheet.KeyColumn = NullIfEmpty(KeyColumnBox.Text);
        _currentSheet.RowStart = ParseInt(RowStartBox.Text);
        _currentSheet.RowEnd = ParseInt(RowEndBox.Text);
        _currentSheet.ColStart = NullIfEmpty(ColStartBox.Text)?.ToUpperInvariant();
        _currentSheet.ColEnd = NullIfEmpty(ColEndBox.Text)?.ToUpperInvariant();
        _currentSheet.ExcludeCols = SplitCsv(ExcludeColsBox.Text);
        _currentSheet.ArrayFields = SplitCsv(ArrayFieldsBox.Text);
        _currentSheet.RoundFields = SplitCsv(RoundFieldsBox.Text);
        _currentSheet.RemoveEmpty = RemoveEmptyBox.IsChecked == true;
        _currentSheet.NoKeyOutputMode = ReadMode(NoKeyModeBox);
        _currentSheet.CustomHeaders = ToDict(_headers, ignoreCase: true);
        _currentSheet.ComputedColumns = ToDict(_computed, ignoreCase: false);
    }

    private void CommitSubEditor()
    {
        if (_currentSub == null || _currentSheet == null || _currentSubKey == null || _loadingUi)
            return;

        _currentSub.OutputName = NullIfEmpty(SubOutputBox.Text);
        _currentSub.RowStart = ParseInt(SubRowStartBox.Text);
        _currentSub.RowEnd = ParseInt(SubRowEndBox.Text);
        _currentSub.ColStart = NullIfEmpty(SubColStartBox.Text)?.ToUpperInvariant();
        _currentSub.ColEnd = NullIfEmpty(SubColEndBox.Text)?.ToUpperInvariant();
        _currentSub.ExcludeCols = SplitCsv(SubExcludeBox.Text);
        _currentSub.ArrayFields = SplitCsv(SubArrayBox.Text);
        _currentSub.RoundFields = SplitCsv(SubRoundBox.Text);
        _currentSub.WrapWithSubKey = SubWrapBox.IsChecked == true;
        _currentSub.RemoveEmpty = SubRemoveEmptyBox.IsChecked == true;
        _currentSub.NoKeyOutputMode = ReadMode(SubNoKeyModeBox);
        _currentSub.CustomHeaders = ToDict(_subHeaders, ignoreCase: true);
        _currentSub.ComputedColumns = ToDict(_subComputed, ignoreCase: false);

        var newKey = NullIfEmpty(SubKeyBox.Text) ?? _currentSubKey;
        if (newKey != _currentSubKey)
        {
            if (_currentSheet.SubTables.ContainsKey(newKey))
            {
                MessageBox.Show($"子表 Key 已存在: {newKey}", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                SubKeyBox.Text = _currentSubKey;
                return;
            }

            _currentSheet.SubTables.Remove(_currentSubKey);
            _currentSheet.SubTables[newKey] = _currentSub;
            var idx = _subKeys.IndexOf(_currentSubKey);
            if (idx >= 0) _subKeys[idx] = newKey;
            _currentSubKey = newKey;
        }
    }

    private void SubKeyBox_LostFocus(object sender, RoutedEventArgs e) => CommitSubEditor();

    private void BrowseExcelDir_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog { Title = "选择 Excel 文件所在目录" };
        if (!string.IsNullOrWhiteSpace(ExcelDirBox.Text) && Directory.Exists(ExcelDirBox.Text))
            dialog.InitialDirectory = ExcelDirBox.Text;

        if (dialog.ShowDialog() == true)
        {
            ExcelDirBox.Text = dialog.FolderName;
            if (string.IsNullOrWhiteSpace(OutputDirBox.Text) ||
                OutputDirBox.Text.EndsWith($"{Path.DirectorySeparatorChar}output", StringComparison.OrdinalIgnoreCase))
            {
                OutputDirBox.Text = Path.Combine(dialog.FolderName, "output");
            }
        }
    }

    private void BrowseOutputDir_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog { Title = "选择 JSON 输出目录" };
        if (!string.IsNullOrWhiteSpace(OutputDirBox.Text) && Directory.Exists(OutputDirBox.Text))
            dialog.InitialDirectory = OutputDirBox.Text;

        if (dialog.ShowDialog() == true)
            OutputDirBox.Text = dialog.FolderName;
    }

    private void AddFile_Click(object sender, RoutedEventArgs e)
    {
        var name = Prompt("添加 Excel 文件名", "文件名（含扩展名）:", "新表.xlsx");
        if (string.IsNullOrWhiteSpace(name)) return;
        name = name.Trim();
        if (!_config.ExcelFiles.Contains(name))
            _config.ExcelFiles.Add(name);
        if (!_config.FileRules.ContainsKey(name))
            _config.FileRules[name] = new FileRuleDto();
        RefreshFileCombo(name);
    }

    private void RemoveFile_Click(object sender, RoutedEventArgs e)
    {
        if (FileCombo.SelectedItem is not string name) return;
        if (MessageBox.Show($"删除文件配置「{name}」？", "确认", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
            return;
        _config.ExcelFiles.Remove(name);
        _config.FileRules.Remove(name);
        RefreshFileCombo();
    }

    private void AddSheet_Click(object sender, RoutedEventArgs e)
    {
        if (_currentFile == null) return;
        var name = Prompt("添加 Sheet", "Sheet 名称:", "新Sheet");
        if (string.IsNullOrWhiteSpace(name)) return;
        name = name.Trim();
        var fileRule = _config.FileRules[_currentFile];
        if (fileRule.Sheets.ContainsKey(name))
        {
            MessageBox.Show("Sheet 已存在", "提示");
            return;
        }

        var sheet = new SheetRuleDto { Enabled = true };
        fileRule.Sheets[name] = sheet;
        var item = new SheetListItem(name, sheet);
        _sheetItems.Add(item);
        SheetList.SelectedItem = item;
    }

    private void RemoveSheet_Click(object sender, RoutedEventArgs e)
    {
        if (_currentFile == null || SheetList.SelectedItem is not SheetListItem item) return;
        CommitCurrentEditors();
        _config.FileRules[_currentFile].Sheets.Remove(item.Name);
        _sheetItems.Remove(item);
        ClearRuleEditor();
    }

    private void AddSubTable_Click(object sender, RoutedEventArgs e)
    {
        if (_currentSheet == null) return;
        CommitSheetEditor();
        var key = Prompt("添加子表", "子表 Key:", "sub1");
        if (string.IsNullOrWhiteSpace(key)) return;
        key = key.Trim();
        if (_currentSheet.SubTables.ContainsKey(key))
        {
            MessageBox.Show("子表 Key 已存在", "提示");
            return;
        }

        var sub = new SubTableRuleDto();
        _currentSheet.SubTables[key] = sub;
        _subKeys.Add(key);
        SubTableList.SelectedItem = key;
    }

    private void RemoveSubTable_Click(object sender, RoutedEventArgs e)
    {
        if (_currentSheet == null || SubTableList.SelectedItem is not string key) return;
        _currentSheet.SubTables.Remove(key);
        _subKeys.Remove(key);
        ClearSubEditor();
    }

    private void ScanSheets_Click(object sender, RoutedEventArgs e)
    {
        if (_currentFile == null)
        {
            MessageBox.Show("请先选择 Excel 文件", "提示");
            return;
        }

        var path = Path.Combine(ExcelDirBox.Text.Trim(), _currentFile);
        if (!File.Exists(path))
        {
            MessageBox.Show($"文件不存在: {path}", "提示");
            return;
        }

        CommitCurrentEditors();
        SyncEnabledFromList();
        SaveDirectories();
        var fileRule = _config.FileRules[_currentFile];
        var added = 0;
        var selectedSheet = (SheetList.SelectedItem as SheetListItem)?.Name;
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = new XLWorkbook(fs);
            foreach (var ws in wb.Worksheets)
            {
                var name = ws.Name;
                if (_config.ExcludeSheets.Contains(name))
                    continue;
                if (!fileRule.Sheets.TryGetValue(name, out var sheet))
                {
                    sheet = new SheetRuleDto { Enabled = true };
                    fileRule.Sheets[name] = sheet;
                    added++;
                }
                if (_config.AutoDetectRows)
                    ExcelConverter.UpdateRowEnds(ws, sheet, AppendLog);
            }

            // 先刷新编辑框，避免切换列表时将旧行数重新提交到规则。
            RefreshDetectedRowEnds();
            LoadSheetsForSelectedFile();
            SheetList.SelectedItem = _sheetItems.FirstOrDefault(item => item.Name == selectedSheet);
            RulesStore.Save(_config);
            AppendLog($"扫描完成，新增 {added} 个 Sheet，规则已保存: {RulesStore.DefaultPath}");
            StatusText.Text = "扫描完成，规则已保存";
        }
        catch (Exception ex)
        {
            RefreshDetectedRowEnds();
            StatusText.Text = "扫描或保存失败";
            AppendLog($"扫描或保存失败: {ex.Message}");
        }
    }

    private void SaveRules_Click(object sender, RoutedEventArgs e)
    {
        CommitCurrentEditors();
        SyncEnabledFromList();
        SaveDirectories();
        RulesStore.Save(_config);
        AppendLog($"规则已保存: {RulesStore.DefaultPath}");
        StatusText.Text = "规则已保存";
    }

    private void ResetRules_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show("恢复内置默认规则并覆盖 rules.json？", "确认",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            return;

        _config = RulesStore.ResetToDefault();
        LoadDirectories();
        RefreshFileCombo();
        AppendLog("已恢复默认规则。");
        StatusText.Text = "已恢复默认";
    }

    private async void Convert_Click(object sender, RoutedEventArgs e)
    {
        if (_isRunning) return;

        CommitCurrentEditors();
        SyncEnabledFromList();

        var excelDir = ExcelDirBox.Text.Trim();
        var outputDir = OutputDirBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(excelDir) || !Directory.Exists(excelDir))
        {
            MessageBox.Show("请选择有效的 Excel 目录。", "提示");
            return;
        }

        if (string.IsNullOrWhiteSpace(outputDir))
        {
            MessageBox.Show("请指定输出目录。", "提示");
            return;
        }

        SaveDirectories();
        RulesStore.Save(_config);

        _isRunning = true;
        ConvertButton.IsEnabled = false;
        StatusText.Text = "转换中...";
        AppendLog("");
        AppendLog("======== 开始转换 ========");

        var progress = new Progress<string>(AppendLog);
        var config = _config;

        try
        {
            var result = await Task.Run(() => ExcelConverter.Run(excelDir, outputDir, config, progress));
            RefreshDetectedRowEnds();
            if (result.Success)
            {
                try
                {
                    RulesStore.Save(config);
                    AppendLog($"最新行数和规则已保存: {RulesStore.DefaultPath}");
                }
                catch (Exception ex)
                {
                    StatusText.Text = "转换完成，规则保存失败";
                    AppendLog($"JSON 已生成，但规则保存失败: {ex.Message}");
                    return;
                }
            }
            StatusText.Text = result.Success ? "转换完成" : "转换失败";
            if (!result.Success && !string.IsNullOrEmpty(result.Error))
                AppendLog(result.Error);
        }
        catch (Exception ex)
        {
            StatusText.Text = "转换失败";
            AppendLog($"未处理异常: {ex}");
            MessageBox.Show(ex.Message, "转换失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _isRunning = false;
            ConvertButton.IsEnabled = true;
        }
    }

    private void OpenOutput_Click(object sender, RoutedEventArgs e)
    {
        var outputDir = OutputDirBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputDir)) return;
        Directory.CreateDirectory(outputDir);
        Process.Start(new ProcessStartInfo { FileName = outputDir, UseShellExecute = true });
    }

    private void ClearLog_Click(object sender, RoutedEventArgs e) => LogBox.Clear();

    private void SyncEnabledFromList()
    {
        foreach (var item in _sheetItems)
            item.Rule.Enabled = item.Enabled;
    }

    private void LoadDirectories()
    {
        AutoDetectRowsBox.IsChecked = _config.AutoDetectRows;
        var excelDir = string.IsNullOrWhiteSpace(_config.ExcelDirectory)
            ? ExcelConverter.ResolveExcelDirectory()
            : _config.ExcelDirectory!;
        ExcelDirBox.Text = excelDir;
        OutputDirBox.Text = string.IsNullOrWhiteSpace(_config.OutputDirectory)
            ? Path.Combine(excelDir, "output")
            : _config.OutputDirectory!;
    }

    private void SaveDirectories()
    {
        _config.AutoDetectRows = AutoDetectRowsBox.IsChecked == true;
        _config.ExcelDirectory = ExcelDirBox.Text.Trim();
        _config.OutputDirectory = OutputDirBox.Text.Trim();
    }

    private void RefreshDetectedRowEnds()
    {
        if (_currentSheet != null)
            RowEndBox.Text = _currentSheet.RowEnd?.ToString() ?? "";
        if (_currentSub != null)
            SubRowEndBox.Text = _currentSub.RowEnd?.ToString() ?? "";
    }

    private void AppendLog(string message)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.Invoke(() => AppendLog(message));
            return;
        }

        if (LogBox.Text.Length > 0)
            LogBox.AppendText(Environment.NewLine);
        LogBox.AppendText(message);
        LogBox.ScrollToEnd();
    }

    private static string? NullIfEmpty(string? s)
        => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    private static int? ParseInt(string? s)
        => int.TryParse(s?.Trim(), out var n) ? n : null;

    private static List<string> SplitCsv(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return [];
        return text.Split([',', '，', ';', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();
    }

    private static Dictionary<string, string> ToDict(IEnumerable<DictEntry> entries, bool ignoreCase)
    {
        var cmp = ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
        var d = new Dictionary<string, string>(cmp);
        foreach (var e in entries)
        {
            if (string.IsNullOrWhiteSpace(e.Key)) continue;
            d[e.Key.Trim()] = e.Value?.Trim() ?? "";
        }

        return d;
    }

    private static void SelectMode(ComboBox box, NoKeyOutputMode mode)
    {
        foreach (ComboBoxItem item in box.Items)
        {
            if (item.Tag?.ToString() == mode.ToString())
            {
                box.SelectedItem = item;
                return;
            }
        }

        box.SelectedIndex = 0;
    }

    private static NoKeyOutputMode ReadMode(ComboBox box)
    {
        if (box.SelectedItem is ComboBoxItem item &&
            Enum.TryParse<NoKeyOutputMode>(item.Tag?.ToString(), out var mode))
            return mode;
        return NoKeyOutputMode.Array;
    }

    private static string? Prompt(string title, string label, string defaultValue)
    {
        var win = new Window
        {
            Title = title,
            Width = 360,
            Height = 160,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
        };
        var panel = new StackPanel { Margin = new Thickness(16) };
        panel.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0, 0, 0, 8) });
        var box = new TextBox { Text = defaultValue, Height = 28 };
        panel.Children.Add(box);
        var ok = new Button { Content = "确定", Width = 80, Height = 28, Margin = new Thickness(0, 12, 0, 0), IsDefault = true };
        string? result = null;
        ok.Click += (_, _) => { result = box.Text; win.DialogResult = true; };
        panel.Children.Add(ok);
        win.Content = panel;
        win.Owner = Application.Current.MainWindow;
        return win.ShowDialog() == true ? result : null;
    }
}

public sealed class SheetListItem : INotifyPropertyChanged
{
    private bool _enabled;

    public SheetListItem(string name, SheetRuleDto rule)
    {
        Name = name;
        Rule = rule;
        _enabled = rule.Enabled;
    }

    public string Name { get; }
    public SheetRuleDto Rule { get; }

    public bool Enabled
    {
        get => _enabled;
        set
        {
            if (_enabled == value) return;
            _enabled = value;
            Rule.Enabled = value;
            OnPropertyChanged();
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
