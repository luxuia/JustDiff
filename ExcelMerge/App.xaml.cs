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
            var win = new MainWindow();
            win.Show();

            //File.WriteAllLines("test", e.Args);
            if (e.Args.Length > 1) {
                if (e.Args[0] == "-difflist") {
                    string[] input = new string[e.Args.Length - 1];
                    Array.Copy(e.Args, 1, input, 0, e.Args.Length - 1);
                    win.DiffList(input);
                }
                else {
                    win.Diff(e.Args[0], e.Args[1]);
                }
            } else {

            }
        }
    }
}
