using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using SharpSvn;
using SharpSvn.UI;
using System.Collections.ObjectModel;
using NetDiff;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace ExcelMerge {
    /// <summary>
    /// DirectoryDifferWindow.xaml 的交互逻辑
    /// </summary>
    public partial class DirectoryDifferWindow : Window {

        public static DirectoryDifferWindow instance;

        public class FileInfo {
            public string root;
            public HashSet<string> files = new HashSet<string>();
        }

        public Dictionary<string, FileInfo> fileList;

        public FileInfo src {
            get {
                return fileList["src"];
            }
        }
        public FileInfo dst {
            get {
                return fileList["dst"];
            }
        }

        public List<DiffResult<string>> results;

        public List<RevisionRange> resultRevisions;



        public DirectoryDifferWindow() {
            InitializeComponent();

            Clear();
            instance = this;
        }

        public void Clear() {
            fileList = new Dictionary<string, FileInfo>();

            fileList["src"] = new FileInfo();
            fileList["dst"] = new FileInfo();

            results = new List<DiffResult<string>>();
            resultRevisions = new List<RevisionRange>();
        }

        public void OnSetDirs(string[] dirs, string tag = "src") {
            if (dirs != null && dirs.Any()) {
                OnFileLoaded(dirs[0], tag, FileOpenType.Drag);

                if (dirs.Length > 1) {
                    OnFileLoaded(dirs[1], tag == "src" ? "dst" :"src", FileOpenType.Drag);
                }
            }
        }

        // load进来单个文件的情况
        public void OnFileLoaded(string file, string tag, FileOpenType type) {
            file = file.Replace("\\", "/");

            SrcDir.Content = file;

            fileList[tag].root = file;
            List<string> list = null;
            if (File.Exists(file) && Util.CheckIsXLS(file)) {
                list = new List<string>() { file };
            } else if (Directory.Exists(file)) {
                var alist = Directory.GetFiles(file, @"*.xls*", SearchOption.AllDirectories);
                list = alist.ToList();
            }
            list = list.Select(a => { return a.Replace('\\', '/').Substring(file.Length); }).ToList();
            fileList[tag].files = new HashSet<string>(list);

            if (src.files.Count> 0 && dst.files.Count> 0) {
                DoDiff();
            }
        }

        public void DoDiff() {
            var total = src.files.Union(dst.files).ToList();
            total.Sort();

            results.Clear();
            resultRevisions.Clear();

            var md5Hash = MD5.Create();

            foreach (var file in total) {
                var res = new DiffResult<string>(file, file, DiffStatus.Equal);
                var consrc = src.files.Contains(file);
                var condst = dst.files.Contains(file);
                if (consrc && !condst) {
                    res.Obj2 = null;
                    res.Status = DiffStatus.Deleted;
                } else if (!consrc && condst) { 
                    res.Obj1 = null;
                    res.Status = DiffStatus.Inserted;
                }

                if (res.Status == DiffStatus.Equal) {
                    // check md5 first
                    //var hash1 = md5Hash.ComputeHash(File.OpenRead(src.root + file));
                    //var hash2 = md5Hash.ComputeHash(File.OpenRead(dst.root + file));
                    if (!DiffUtil.FilesAreEqual(src.root + file, dst.root+file)) {
                        res.Status = DiffStatus.Modified;
                    }
                }

                results.Add(res);
            }

            DstDataGrid.refreshData();
            SrcDataGrid.refreshData();
        }

        public void OnGridScrollChanged(string tag, ScrollChangedEventArgs e) {
            ScrollViewer view = null;
            if (tag == "src") {
                view = Util.GetVisualChild<ScrollViewer>(DstDataGrid);
            }
            else if (tag == "dst") {
                view = Util.GetVisualChild<ScrollViewer>(SrcDataGrid);
            }
            view.ScrollToVerticalOffset(e.VerticalOffset);
            view.ScrollToHorizontalOffset(e.HorizontalOffset);
        }

        public void OnSelectGridRow(string tag, int rowid) {
            if (tag == "src") {
                DstDataGrid.FileGrid.SelectedIndex = rowid;
            }
            else {
                SrcDataGrid.FileGrid.SelectedIndex = rowid;
            }
            var res = results[rowid];
            if (res.Status == DiffStatus.Modified) {
                if (resultRevisions.Count < results.Count) {
                    MainWindow.instance.Diff(src.root + res.Obj1, dst.root + res.Obj2);
                } else {
                    var range = resultRevisions[rowid];
                    MainWindow.instance.OnFileLoaded(src.root + range.file, "src", FileOpenType.Drag);
                    MainWindow.instance.Diff(range.min, range.max);
                }
            }
        }

        public class RevisionRange {
            public long min, max;
            public string file;
        }

        private void DoVersionDiff_Click(object sender, RoutedEventArgs e) {
            if (string.IsNullOrEmpty(FilterCommits.Text)) {
                TextTip.Content = "需要填写单号，多个单号空格隔开";
                return;
            }
            if (src.files.Count <=0) {
                TextTip.Content = "拖目标文件夹进来,或目标文件夹下没有xls";
                return;
            }

            Collection<SvnLogEventArgs> logitems;

            DateTime startDateTime = DateTime.Now.AddDays(-60);
            DateTime endDateTime = DateTime.Now;
            var svnRange = new SvnRevisionRange(new SvnRevision(startDateTime), new SvnRevision(endDateTime));

            List<SvnRevisionCombo> revisions = new List<SvnRevisionCombo>();

            var files = new Dictionary<string, RevisionRange>();

            var filter = FilterCommits.Text.Split(" #".ToArray(), StringSplitOptions.RemoveEmptyEntries);
            
            var sfilter = string.Join("|", filter);
            var regfilter = new Regex(sfilter);

            using (SvnClient client = new SvnClient()) {
                client.Authentication.SslServerTrustHandlers += delegate (object _sender, SharpSvn.Security.SvnSslServerTrustEventArgs _e) {
                    _e.AcceptedFailures = _e.Failures;
                    _e.Save = true; // Save acceptance to authentication store
                };

                if (client.GetUriFromWorkingCopy(src.root) != null) {

                    SvnInfoEventArgs info;
                    client.GetInfo(src.root, out info);
                    var uri = info.Uri;

                    var rootPath = info.Path;

                    client.GetLog(uri, new SvnLogArgs(svnRange), out logitems);

                    foreach (var logentry in logitems) {
                        var author = logentry.Author;
                        var message = logentry.LogMessage;
                        var date = logentry.Time;
                        var revision = logentry.Revision;

                        if (regfilter.IsMatch(message)) {
                            foreach( var filepath in logentry.ChangedPaths) {
                                var path = filepath.Path;

                                RevisionRange minmax = null;
                                if (!files.TryGetValue(path, out minmax)) {
                                    minmax = new RevisionRange(){ min = revision, max = revision, file = path};
                                    files[path] = minmax;
                                }
                                if (revision > minmax.max) minmax.max= revision;
                                if (revision < minmax.min) minmax.min = revision;
                            }
                        }
                    }
                }
            }

            results.Clear();
            resultRevisions.Clear();

            foreach (var file in files.Keys) {
                var range = files[file];
                var res = new DiffResult<string>(file + "-" + range.min, file + "-" + range.max, DiffStatus.Modified);

                results.Add(res);
                resultRevisions.Add(range);
            }

            DstDataGrid.refreshData();
            SrcDataGrid.refreshData();
        }

    }
}
