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

        void getUrl(string path, out string url, out long rev)
        {
            rev = 0;
            url = string.Empty;
            var revisionidx = path.LastIndexOf("?revision=");
            revisionidx = revisionidx >= 0 ? revisionidx + "?revision=".Length : -1;
            if (revisionidx < 0)
            {
                revisionidx = path.LastIndexOf("?r=");
                revisionidx = revisionidx >= 0 ? revisionidx + "?r=".Length : -1;
            }
            //xlsmerge://http://m2.svn.ejoy.com/M2/x19/editor/config/resource/C场景传送.xlsx/?r=614129
            if (revisionidx > 0)
            {
                var srev = path.Substring(revisionidx);
                rev = long.Parse(srev);
            }
            url = path;
        }

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
                    int cmpidx = url.LastIndexOf("&cmp=");
                    if (cmpidx > 0)
                    {
                        //xlsmerge://http://m2.svn.ejoy.com/M2/branch/cn/cn_20220721/editor/config/resource/G公会战.xlsx/?r=616108&cmp=http://m2.svn.ejoy.com/M2/x19/editor/config/resource/G公会战.xlsx/?r=611111


                        string path1 = url.Substring(0, cmpidx);

                        string fileurl = string.Empty;
                        long rev = 0;
                        getUrl(path1, out fileurl, out rev);

                        string path2 = url.Substring(cmpidx + "&cmp=".Length);
                        string cmpfileurl = string.Empty;
                        long cmprev = 0;
                        getUrl(path2, out cmpfileurl, out cmprev);

                        win.DiffUri(rev, new Uri(url), cmprev, new Uri(cmpfileurl));
                    }
                    else
                    {
                        //xlsmerge://http://m2.svn.ejoy.com/M2/x19/editor/config/resource/C场景传送.xlsx/?r=614129
                        string fileurl = string.Empty;
                        long rev = 0;
                        getUrl(url, out fileurl, out rev);
                        if (!string.IsNullOrEmpty(fileurl))
                        {
                            win.DiffUri(rev - 1, rev, new Uri(url));
                        }
                    }
                }
            }
        }
    }
}
