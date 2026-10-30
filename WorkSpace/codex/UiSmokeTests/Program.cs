using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClosedXML.Excel;
using ExcelToJson;

internal static class Program
{
    [STAThread]
    private static int Main()
    {
        var app = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        app.Resources.MergedDictionaries.Add(new ResourceDictionary
        {
            Source = new Uri("/ExcelToJson;component/Theme.xaml", UriKind.Relative)
        });
        SynchronizationContext.SetSynchronizationContext(new DispatcherSynchronizationContext(Dispatcher.CurrentDispatcher));
        MainWindow? window = null;
        try
        {
            var root = Path.Combine(AppContext.BaseDirectory, "fixtures");
            Directory.CreateDirectory(root);
            WriteWorkbook(root, 3);
            var config = new ConversionConfig
            {
                ExcelDirectory = root, OutputDirectory = Path.Combine(root, "output"),
                ExcelFiles = ["input.xlsx"], FileRules = new()
                {
                    ["input.xlsx"] = new() { Sheets = new()
                    {
                        ["Data"] = new() { RowStart = 2, RowEnd = 2, ColStart = "A", ColEnd = "A" },
                        ["Other"] = new() { Enabled = false }
                    } }
                }
            };
            // DefaultPath is isolated in this test executable's own output directory.
            RulesStore.Save(config);
            window = new MainWindow();
            var syncButton = (Button)window.FindName("SyncButton");
            Check(syncButton is not null, "Sync action must be available");
            ((TextBox)window.FindName("SvnDirBox")).Text = root;
            ((TextBox)window.FindName("TargetDirBox")).Text = Path.Combine(root, "cocos");
            var list = (ListBox)window.FindName("SheetList");
            list.SelectedIndex = 0;
            var copyToggle = (CheckBox)window.FindName("CopyToTargetBox");
            Check(copyToggle.IsChecked == false, "Copy must default off");
            var otherTab = (TabItem)window.FindName("OtherSettingsTab");
            otherTab.IsSelected = true;
            Check(ReferenceEquals(copyToggle.Parent, otherTab.Content) && otherTab.Header.ToString() == "其他配置", "Copy belongs in Other Settings");
            copyToggle.IsChecked = true;
            list.SelectedIndex = 1;
            Check(copyToggle.IsChecked == false, "Copy settings must be independent per sheet");
            list.SelectedIndex = 0;
            Check(copyToggle.IsChecked == true, "Switching sheets must retain copy selection");
            var toggle = (CheckBox)window.FindName("AutoDetectRowsBox");
            var endBox = (TextBox)window.FindName("RowEndBox");
            Pump();
            Check(toggle.IsChecked == true && endBox.IsReadOnly, "Default automatic toggle / readonly binding failed");
            Invoke(window, "ScanSheets_Click");
            Check(ReadEnd() == 3 && endBox.Text == "3", "Scan must save actual end and refresh editor");
            Invoke(window, "SaveRules_Click");
            Check(RulesStore.LoadOrCreate().TargetDirectory == Path.Combine(root, "cocos"), "Sync target must persist");
            Check(ReadEnd() == 3, "Editor must not overwrite detected end with stale text");
            Check(RulesStore.LoadOrCreate().FileRules["input.xlsx"].Sheets["Data"].CopyToTarget, "Copy selection must save");
            WriteWorkbook(root, 4);
            ((Button)window.FindName("ConvertButton")).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Pump();
            Check(!syncButton!.IsEnabled, "Sync action must be disabled while converting");
            Check(((TabControl)window.FindName("RuleTabs")).Visibility == Visibility.Visible,
                "Running must keep selected rules visible instead of showing the empty state");
            Check(((TabControl)window.FindName("WorkspaceTabs")).IsEnabled, "Workspace navigation must remain available while running");
            ((TabItem)window.FindName("LogPage")).IsSelected = true;
            Check(((TextBox)window.FindName("LogBox")).IsEnabled, "Logs must remain readable while running");
            var deadline = DateTime.UtcNow.AddSeconds(30);
            while ((bool)typeof(MainWindow).GetField("_isRunning", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(window)!)
            {
                if (DateTime.UtcNow > deadline) throw new Exception("Conversion timeout");
                Pump();
                Thread.Sleep(10);
            }
            Check(((TextBlock)window.FindName("StatusText")).Text == "转换完成", "Conversion status failed");
            Check(ReadEnd() == 4 && endBox.Text == "4", "Conversion must persist increased end and refresh editor");
            toggle.IsChecked = false;
            Pump();
            Check(!endBox.IsReadOnly, "Manual toggle must unlock end row");
            endBox.Text = "3";
            Invoke(window, "SaveRules_Click");
            Check(!RulesStore.LoadOrCreate().AutoDetectRows && ReadEnd() == 3, "Manual switch/end must persist");
            window.Close();
            window = new MainWindow();
            Check(((TextBox)window.FindName("TargetDirBox")).Text == Path.Combine(root, "cocos"), "Reopened UI must restore sync destination");
            ((ListBox)window.FindName("SheetList")).SelectedIndex = 0;
            Check(((CheckBox)window.FindName("CopyToTargetBox")).IsChecked == true, "Reopen must restore copy selection");
            ((TabItem)window.FindName("OtherSettingsTab")).IsSelected = true;
            ((CheckBox)window.FindName("CopyToTargetBox")).IsChecked = false;
            Invoke(window, "SaveRules_Click");
            Check(!RulesStore.LoadOrCreate().FileRules["input.xlsx"].Sheets["Data"].CopyToTarget, "Turning copy off must persist");
            Check(((CheckBox)window.FindName("AutoDetectRowsBox")).IsChecked == false &&
                ((TextBox)window.FindName("RowEndBox")).Text == "3", "Reopened UI must restore settings");
            if (Environment.GetCommandLineArgs().Contains("--preview-real"))
            {
                var sourceRules = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../../publish/rules.json"));
                var previewConfig = RulesStore.LoadOrCreate(sourceRules);
                window.Close();
                // Save only to the test executable's isolated configuration, never the user's rules.
                RulesStore.Save(previewConfig);
                window = new MainWindow();
                var sheets = (ListBox)window.FindName("SheetList");
                sheets.SelectedItem = sheets.Items.Cast<object>().FirstOrDefault(item =>
                    item.GetType().GetProperty("Name")?.GetValue(item)?.ToString() == "猛鬼技能效果") ?? sheets.Items.Cast<object>().FirstOrDefault();
            }
            // Render each workspace at normal and minimum sizes without opening a desktop window.
            var content = (FrameworkElement)window.Content;
            ((TabControl)window.FindName("RuleTabs")).SelectedIndex = 0;
            foreach (var size in new[] { new Size(1240, 780), new Size(1000, 640) })
            foreach (var page in new[] { "EditorPage", "DirectoriesPage", "LogPage" })
            {
                ((TabItem)window.FindName(page)).IsSelected = true;
                content.Measure(size); content.Arrange(new Rect(new Point(), size)); content.UpdateLayout();
                var status = (TextBlock)window.FindName("StatusText");
                Check(status.TransformToAncestor(content).Transform(new Point(0, status.ActualHeight)).Y <= size.Height,
                    "Status must stay inside the window");
                Check(((ListBox)window.FindName("SheetList")).ActualHeight > 180, "Persistent sheet navigation must have usable height");
                if (page == "EditorPage" && size.Height == 780)
                    Check(((Grid)window.FindName("RuleWorkspace")).ActualHeight > 500, "Rules should use the full workspace height");
                var bitmap = new RenderTargetBitmap((int)size.Width, (int)size.Height, 96, 96, PixelFormats.Pbgra32);
                bitmap.Render(content);
                var png = new PngBitmapEncoder(); png.Frames.Add(BitmapFrame.Create(bitmap));
                using var image = File.Create(Path.Combine(root, $"layout-{page}-{size.Width}.png")); png.Save(image);
            }
            Console.WriteLine("PASS WPF scan, convert, save, editor refresh, toggle and reopen");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FAIL {ex}");
            return 1;
        }
        finally { window?.Close(); app.Shutdown(); }
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    private static int? ReadEnd() => RulesStore.LoadOrCreate().FileRules["input.xlsx"].Sheets["Data"].RowEnd;

    private static void Invoke(MainWindow window, string name) => typeof(MainWindow)
        .GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance)!
        .Invoke(window, [window, new RoutedEventArgs()]);

    private static void Pump()
    {
        var frame = new DispatcherFrame();
        Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() => frame.Continue = false));
        Dispatcher.PushFrame(frame);
    }

    private static void WriteWorkbook(string root, int end)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        for (int row = 2; row <= end; row++) ws.Cell(row, 1).Value = row;
        wb.SaveAs(Path.Combine(root, "input.xlsx"));
    }
}
