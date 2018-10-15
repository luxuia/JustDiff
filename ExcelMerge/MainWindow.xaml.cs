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
using System.Windows.Navigation;
using System.Windows.Shapes;
using NPOI.SS.UserModel;
using SharpSvn;
using SharpSvn.UI;
using System.Collections.ObjectModel;

namespace ExcelMerge {


    public class WorkBookWrap {
        public IWorkbook book;
        public int sheet;
        public string file;
        public string filename;
    }

    public enum FileOpenType {
        Drag,
        Menu,
        Prog, //因为diff等形式从程序内部打开的
    }

    public class CurrentStatus {
        public string diffPath1;
        public string diffPath2;
        public int diffRevision1;
        public int diffRevision2;
        public IWorkbook book1;
        public IWorkbook book2;

        public List<NetDiff.DiffResult<string>> diffHead;
        public List<NetDiff.DiffResult<string>> diffFistColumn;


    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        public static MainWindow instance;

        public Dictionary<string, WorkBookWrap> books = new Dictionary<string, WorkBookWrap>();

        public CurrentStatus status = new CurrentStatus();

        public MainWindow() {
            InitializeComponent();

            instance = this;
        }

        public void DataGrid_SelectedCellsChanged(object sender, SelectionChangedEventArgs e) {

        }

        public class SheetNameCombo {
            public string Name { get; set; }
            public int ID { get; set; }
        }

        public class SvnRevisionCombo {
            public string Revision { get; set; }
            public int ID { get; set; }
        }

        public void OnFileLoaded(string file, string tag, FileOpenType type) {

            var wb = WorkbookFactory.Create(file);

            if (type == FileOpenType.Drag || type == FileOpenType.Menu) {
                books[tag] = new WorkBookWrap() { book = wb, sheet = 0, file = file, filename = System.IO.Path.GetFileName(file) };

                UpdateSVNRevision(file, tag);
            }

            if (tag == "src") {
                SrcFilePath.Content = file;
                List<SheetNameCombo> list = new List<SheetNameCombo>();
                for (int i = 0; i < wb.NumberOfSheets; ++i) {
                    list.Add(new SheetNameCombo() { Name = wb.GetSheetName(i), ID = i });
                }
                SrcFileSheetsCombo.ItemsSource = list;
                SrcFileSheetsCombo.SelectedValue = 0;
            }
            else if (tag == "dst") {
                DstFilePath.Content = file;
                List<SheetNameCombo> list = new List<SheetNameCombo>();
                for (int i = 0; i < wb.NumberOfSheets; ++i) {
                    list.Add(new SheetNameCombo() { Name = wb.GetSheetName(i), ID = i });
                }
                DstFileSheetsCombo.ItemsSource = list;
                DstFileSheetsCombo.SelectedValue = 0;
            }
        }

        void UpdateSVNRevision(string file, string tag) {

            if (tag == "src") {
                Collection<SvnLogEventArgs> logitems;

                DateTime startDateTime = DateTime.Now.AddDays(-30);
                DateTime endDateTime = DateTime.Now;
                var svnRange = new SvnRevisionRange(new SvnRevision(startDateTime), new SvnRevision(endDateTime));

                List<SvnRevisionCombo> revisions = new List<SvnRevisionCombo>();

                using (SvnClient client = new SvnClient()) {
                    client.Authentication.SslServerTrustHandlers += delegate (object sender, SharpSvn.Security.SvnSslServerTrustEventArgs e) {
                        e.AcceptedFailures = e.Failures;
                        e.Save = true; // Save acceptance to authentication store
                    };

                    SvnInfoEventArgs info;
                    client.GetInfo(file, out info);
                    var uri = info.Uri;

                    client.GetLog(uri, new SvnLogArgs(svnRange), out logitems);

                    foreach (var logentry in logitems) {
                        var author = logentry.Author;
                        var message = logentry.LogMessage;
                        var date = logentry.Time;

                        revisions.Add(new SvnRevisionCombo() { Revision = string.Format("{0}[{1}]", author, message), ID = (int)logentry.Revision });
                    }
                }
                SVNRevisionCombo.ItemsSource = revisions;
            }
        }

        void Diff(int revision, int revisionto) {
            using (SvnClient client = new SvnClient()) {
                var srcBook = books["src"];
                string file = srcBook.file;
                SvnInfoEventArgs info;
                client.GetInfo(file, out info);
                var uri = info.Uri;

                var tempDir = System.IO.Path.GetTempPath();
                var filename = srcBook.filename;

                var file1 = tempDir + revision + "_" + filename;
                var checkoutArgs = new SvnWriteArgs() { Revision = revision };
                using (var fs = System.IO.File.Create(file1)) {
                    client.Write(uri, fs, checkoutArgs);
                }
                var file2 = tempDir + revisionto + "_" + filename;
                var checkoutArgs2 = new SvnWriteArgs() { Revision = revisionto };
                using (var fs = System.IO.File.Create(file2)) {
                    client.Write(uri, fs, checkoutArgs2);
                }

                // diff的时候就是和自己diff
                // 这种情况直接把file2带入好像也没问题
                books["dst"] = new WorkBookWrap() { book = srcBook.book, file = srcBook.file, filename = srcBook.file, sheet = srcBook.sheet };

                status.diffPath1 = file1;
                status.diffPath2 = file2;
                status.diffRevision1 = revision;
                status.diffRevision2 = revisionto;
                status.book1 = WorkbookFactory.Create(file1);
                status.book2 = WorkbookFactory.Create(file2);

                OnFileLoaded(file1, "src", FileOpenType.Prog);
                OnFileLoaded(file2, "dst", FileOpenType.Prog);
            }
        }

        string GetCellValue(ICell cell) {
            var str = string.Empty;
            switch (cell.CellType) {
                case CellType.Blank:
                    str = cell.StringCellValue;
                    break;
                case CellType.Boolean:
                    str = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Error:
                    str = cell.ErrorCellValue.ToString();
                    break;
                case CellType.Formula:
                    str = cell.CellFormula;
                    break;
                case CellType.Numeric:
                    str = cell.NumericCellValue.ToString();
                    break;
                case CellType.String:
                    str = cell.RichStringCellValue.ToString();
                    break;
            }
            return str.Replace('(', '-').Replace(')', '-');
        }

        List<string> GetHeaderStrList(ISheet sheet) {
            List<string> header = new List<string>();

            var row0 = sheet.GetRow(0);
            var row1 = sheet.GetRow(1);
            var row2 = sheet.GetRow(2);

            for (int i = 0; i <row0.Cells.Count;++i) {
                header.Add(string.Concat(GetCellValue(row0.Cells[i]), ":", GetCellValue(row1.Cells[i]),":",GetCellValue(row2.Cells[i])));
            }
            return header;
        }

        void OnSheetChanged() {
            // TODO这里检查diff标记，清理book1

            if (status.book1 != null && status.book2 != null) {
                var sheet1 = status.book1.GetSheetAt(books["src"].sheet);
                var sheet2 = status.book2.GetSheetAt(books["dst"].sheet);

                var head1 = GetHeaderStrList(sheet1);
                var head2 = GetHeaderStrList(sheet2);

                var diff = NetDiff.DiffUtil.Diff(head1, head2);
                var optimized = NetDiff.DiffUtil.OptimizeCaseDeletedFirst(diff);


                status.diffHead = optimized.ToList();
            }

        }

        private void DstFileSheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (e.AddedItems.Count > 0) {
                var selection = e.AddedItems[0] as SheetNameCombo;
                books["dst"].sheet = selection.ID;

                DstDataGrid.RefreshData();

                if (books.ContainsKey("src") && books["src"].sheet != selection.ID) {
                    SrcFileSheetsCombo.SelectedValue = selection.ID;
                }

                OnSheetChanged();
            }
        }

        private void SrcFileSheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (e.AddedItems.Count > 0) {
                var selection = e.AddedItems[0] as SheetNameCombo;
                books["src"].sheet = selection.ID;

                SrcDataGrid.RefreshData();

                if (books.ContainsKey("dst") && books["dst"].sheet != selection.ID) {
                    DstFileSheetsCombo.SelectedValue = selection.ID;
                }

                OnSheetChanged();
            }
        }

        private void SVNResivionionList_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            var selection = e.AddedItems[0] as SvnRevisionCombo;
            SVNRevisionCombo.Width = selection.Revision.Length*10;

            Diff(selection.ID - 1, selection.ID);
        }
    }


}
