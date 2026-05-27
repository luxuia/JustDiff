using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

namespace ExcelMerge {
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application {

        private void Application_Startup(object sender, StartupEventArgs e) {
            DispatcherUnhandledException += (s, args) =>
            {
                var log = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "crash.log");
                File.WriteAllText(log, $"{DateTime.Now}\n{args.Exception}");
                MessageBox.Show(args.Exception.ToString(), "Crash", MessageBoxButton.OK, MessageBoxImage.Error);
                args.Handled = true;
            };
            Entrance.ProcessInput(sender, e);
        }

        private void Application_Exit(object sender, ExitEventArgs e) {
            Entrance.Window_Closing(null, null);
        }
    }
}
