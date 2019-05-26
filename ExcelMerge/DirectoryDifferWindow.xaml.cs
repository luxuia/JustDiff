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
        }



        // load进来单个文件的情况
        public void OnFileLoaded(string file, string tag, FileOpenType type) {
            file = file.Replace("\\", "/");

            fileList[tag].root = file;
            List<string> list = null;
            if (File.Exists(file) && Util.CheckIsXLS(file)) {
                list = new List<string>() { file };
            } else if (Directory.Exists(file)) {
                var alist = Directory.GetFiles(file, @"*.xls", SearchOption.AllDirectories);
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
                    var hash1 = md5Hash.ComputeHash(File.OpenRead(src.root + file));
                    var hash2 = md5Hash.ComputeHash(File.OpenRead(dst.root + file));
                    if (!hash1.SequenceEqual(hash2)) {
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
        }
    }
}
