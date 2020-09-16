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
        const string Scheme = "xlsmerge://";

        private void Application_Startup(object sender, StartupEventArgs e) {
            var win = new MainWindow();
            win.Show();

            //File.WriteAllLines(@"F:\x19_trunk_edit\pc_daily\package\x19_pc\test", e.Args);
            if (e.Args.Length > 1) {
                if (e.Args[0] == "-difflist") {
                    string[] input = new string[e.Args.Length - 1];
                    Array.Copy(e.Args, 1, input, 0, e.Args.Length - 1);
                    win.DiffList(input);
                }
                else {
                    win.Diff(e.Args[0], e.Args[1]);
                }
            } else if (e.Args.Length == 1) {
                var url = e.Args[0];
                if (url.StartsWith(Scheme)) {
                    url = url.Substring(Scheme.Length);
                    var revisionidx = url.LastIndexOf("?revision=");
                    if (revisionidx > 0) {
                        var rev = url.Substring(revisionidx + 10);
                        var irev = long.Parse(rev);
                        win.DiffUri(irev-1, irev, new Uri(url));
                    }
                }
            }
        }
    }
}
