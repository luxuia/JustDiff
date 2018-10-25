using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelMerge {
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application {
        private void Application_Startup(object sender, StartupEventArgs e) {
            var win = new MainWindow();
            win.Show();
 
            if (e.Args.Length > 1) {
                win.Diff(e.Args[0], e.Args[1]);
            } else {

            }
        }
    }
}
