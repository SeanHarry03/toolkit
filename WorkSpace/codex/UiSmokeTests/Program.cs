using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
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
                        ["Data"] = new() { RowStart = 2, RowEnd = 2, ColStart = "A", ColEnd = "A" }
                    } }
                }
            };
            // DefaultPath is isolated in this test executable's own output directory.
            RulesStore.Save(config);
            window = new MainWindow();
            var list = (ListBox)window.FindName("SheetList");
            list.SelectedIndex = 0;
            var toggle = (CheckBox)window.FindName("AutoDetectRowsBox");
            var endBox = (TextBox)window.FindName("RowEndBox");
            Pump();
            Check(toggle.IsChecked == true && endBox.IsReadOnly, "Default automatic toggle / readonly binding failed");
            Invoke(window, "ScanSheets_Click");
            Check(ReadEnd() == 3 && endBox.Text == "3", "Scan must save actual end and refresh editor");
            Invoke(window, "SaveRules_Click");
            Check(ReadEnd() == 3, "Editor must not overwrite detected end with stale text");
            WriteWorkbook(root, 4);
            ((Button)window.FindName("ConvertButton")).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
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
            ((ListBox)window.FindName("SheetList")).SelectedIndex = 0;
            Check(((CheckBox)window.FindName("AutoDetectRowsBox")).IsChecked == false &&
                ((TextBox)window.FindName("RowEndBox")).Text == "3", "Reopened UI must restore settings");
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
