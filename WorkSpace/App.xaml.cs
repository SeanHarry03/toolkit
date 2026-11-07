using System.Windows;

namespace ExcelToJson;

public partial class App : Application
{
    protected override async void OnStartup(StartupEventArgs e)
    {
        if (e.Args.Length > 0 && e.Args[0] == "--gametools")
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            var code = await HeadlessRunner.RunAsync(e.Args);
            Shutdown(code);
            return;
        }
        base.OnStartup(e);
        new MainWindow().Show();
    }
}
