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
using NetDiff;
using string2int = System.Collections.Generic.KeyValuePair<string, int>;

namespace ExcelMerge {
 
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        public static MainWindow instance;

        public Dictionary<string, WorkBookWrap> books = new Dictionary<string, WorkBookWrap>();

        public Dictionary<int, SheetDiffStatus> sheetsDiff = new Dictionary<int, SheetDiffStatus>();

        public List<DiffResult<SheetNameCombo>> diffSheetName;


        public string SrcFile;
        public string DstFile;
        

        public MainWindow() {
            InitializeComponent();

            instance = this;
        }

        public void DataGrid_SelectedCellsChanged(object sender, SelectionChangedEventArgs e) {

        }

        
        public void OnFileLoaded(string file, string tag, FileOpenType type, int sheet = 0) {

            var wb = WorkbookFactory.Create(file);

            books[tag] = new WorkBookWrap() { book = wb, sheet = sheet, file = file, filename = System.IO.Path.GetFileName(file) };

            if (type == FileOpenType.Drag || type == FileOpenType.Menu) {
                if (tag == "src")
                    SrcFile = file;
                else
                    DstFile = file;
                UpdateSVNRevision(file, tag);
            }

            if (tag == "src") {
                SrcFilePath.Content = file;
                List<ComboBoxItem> list = new List<ComboBoxItem>();
                SrcFileSheetsCombo.Items.Clear();
                for (int i = 0; i < wb.NumberOfSheets; ++i) {
                    var item = new ComboBoxItem();
                    item.Content = new SheetNameCombo() { Name = wb.GetSheetName(i), ID = i };
                    SrcFileSheetsCombo.Items.Add(item);
                    list.Add(item);
                }
                SrcFileSheetsCombo.SelectedItem = list[0];
            }
            else if (tag == "dst") {
                DstFilePath.Content = file;
                List<ComboBoxItem> list = new List<ComboBoxItem>();
                DstFileSheetsCombo.Items.Clear();
                for (int i = 0; i < wb.NumberOfSheets; ++i) {
                    var item = new ComboBoxItem();
                    item.Content = new SheetNameCombo() { Name = wb.GetSheetName(i), ID = i };
                    DstFileSheetsCombo.Items.Add(item);
                    list.Add(item);
                }
                DstFileSheetsCombo.SelectedItem = list[0];
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

        List<ComboBoxItem> SetupSheetCombo(IWorkbook wb ) {
            List<ComboBoxItem> list = new List<ComboBoxItem>();
            SrcFileSheetsCombo.Items.Clear();
            for (int i = 0; i < wb.NumberOfSheets; ++i) {
                var item = new ComboBoxItem();
                item.Content = new SheetNameCombo() { Name = wb.GetSheetName(i), ID = i };
                SrcFileSheetsCombo.Items.Add(item);
                list.Add(item);
            }
            return list;
        }

        WorkBookWrap InitWorkWrap(string file) {
            var wb = new WorkBookWrap() {
                book = WorkbookFactory.Create(file),
                file = file,
                filename = System.IO.Path.GetFileName(file)
            };

            wb.sheetCombo = new List<ComboBoxItem>();
            var list = new List<SheetNameCombo>();
            for (int i = 0; i < wb.book.NumberOfSheets; ++i) {
                list.Add(new SheetNameCombo() { Name = wb.book.GetSheetName(i), ID = i });
            }
            list.Sort((a, b) => { return a.Name.CompareTo(b.Name); });

            wb.sheetName = list;

            wb.ComboIdToItemIdx = new Dictionary<int, int>();

            list.ForEach((a) => { var item = new ComboBoxItem(); item.Content = a; wb.sheetCombo.Add(item); });

            for (int i = 0; i < list.Count;++i) {
                wb.ComboIdToItemIdx[list[i].ID] = i;
            }

            return wb;
        }

        SheetDiffStatus DiffSheet(ISheet src, ISheet dst) {
            var status = new SheetDiffStatus();

            bool changed = false;

            var head1 = GetHeaderStrList(src);
            var head2 = GetHeaderStrList(dst);

            var diff = NetDiff.DiffUtil.Diff(head1, head2);
            var optimized = NetDiff.DiffUtil.OptimizeCaseDeletedFirst(diff);
            optimized = DiffUtil.OptimizeCaseInsertedFirst(diff);

            changed = changed || optimized.Any(a => a.Status != DiffStatus.Equal);

            status.diffHead = optimized.ToList();

            status.columnCount = status.diffHead.Count;

            status.diffFistColumn = GetIDDiffList(src, dst);

            changed = changed || status.diffFistColumn.Any(a => a.Status != DiffStatus.Equal);

            status.diffSheet = new List<List<DiffResult<string>>>();
            status.rowID2DiffMap1 = new Dictionary<int, int>();
            status.rowID2DiffMap2 = new Dictionary<int, int>();

            foreach (var diffkv in status.diffFistColumn) {
                var rowid1 = diffkv.Obj1.Value;
                var rowid2 = diffkv.Obj2.Value;
                if (diffkv.Obj1.Key == null) {
                    rowid1 = -1;
                }
                if (diffkv.Obj2.Key == null) {
                    rowid2 = -1;
                }

                status.rowID2DiffMap1[rowid1] = status.diffSheet.Count;
                status.rowID2DiffMap2[rowid2] = status.diffSheet.Count;

                var diffrow = DiffSheetRow(src, rowid1, dst, rowid2, status);

                status.diffSheet.Add(diffrow);


                changed = changed || diffrow.Any(a => a.Status != DiffStatus.Equal);

                if (changed) {
                    changed = true;
                }
            }

            status.changed = changed;

            return status;
        }

        void Diff(int revision, int revisionto) {
            using (SvnClient client = new SvnClient()) {
                string file = SrcFile;
                SvnInfoEventArgs info;
                client.GetInfo(file, out info);
                var uri = info.Uri;

                var tempDir = System.IO.Path.GetTempPath();
                var filename = System.IO.Path.GetFileName(SrcFile);

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

                int old_sheet = 0;
                if (books.ContainsKey("src")) {
                    old_sheet = books["src"].sheet;
                }

                var src = InitWorkWrap(file1);
                src.sheet = old_sheet;
                var dst = InitWorkWrap(file2);
                dst.sheet = old_sheet;

                var option = new DiffOption<SheetNameCombo>();
                option.EqualityComparer = new SheetNameComboComparer();
                var result = DiffUtil.Diff(src.sheetName, dst.sheetName, option);
                diffSheetName = DiffUtil.OptimizeCaseDeletedFirst(result).ToList();

                books["src"] = src;
                books["dst"] = dst;

                for (int i = 0; i < diffSheetName.Count;++i) {
                    var sheetname = diffSheetName[i];
                    // 只有sheet名字一样的可以diff， 先这么处理
                    if (sheetname.Status == DiffStatus.Equal) {
                        var sheet1 = sheetname.Obj1.ID;
                        var sheet2 = sheetname.Obj2.ID;
                        sheetsDiff[i] = DiffSheet(src.book.GetSheetAt(sheet1), dst.book.GetSheetAt(sheet2));
                    }
                }

                // refresh ui
                SrcFilePath.Content = file1;
                DstFilePath.Content = file2;

                SrcFileSheetsCombo.Items.Clear();
                foreach (var item in books["src"].sheetCombo) {

                    int index = diffSheetName.FindIndex(a=> a.Obj1!=null && a.Obj1.ID == (item.Content as SheetNameCombo).ID);
                    SolidColorBrush color = null;
                    DiffStatus status = diffSheetName[index].Status;
                    if (status != DiffStatus.Equal) {
                        color = Util.GetColorByDiffStatus(status);
                    } else {
                        color = Util.GetColorByDiffStatus(sheetsDiff[index].changed ? DiffStatus.Modified : DiffStatus.Equal);
                    }

                    if (color!=null) {
                        item.Background = color;
                    }

                    SrcFileSheetsCombo.Items.Add(item);
                }
                SrcFileSheetsCombo.SelectedItem = books["src"].sheetCombo[old_sheet];

                DstFileSheetsCombo.Items.Clear();
                foreach (var item in books["dst"].sheetCombo) {

                    int index = diffSheetName.FindIndex(a => a.Obj2 != null && a.Obj2.ID == (item.Content as SheetNameCombo).ID);
                    SolidColorBrush color = null;
                    DiffStatus status = diffSheetName[index].Status;
                    if (status != DiffStatus.Equal) {
                        color = Util.GetColorByDiffStatus(status);
                    }
                    else {
                        color = Util.GetColorByDiffStatus(sheetsDiff[index].changed ? DiffStatus.Modified : DiffStatus.Equal);
                    }

                    if (color != null) {
                        item.Background = color;
                    }

                    DstFileSheetsCombo.Items.Add(item);
                }
                DstFileSheetsCombo.SelectedItem = books["dst"].sheetCombo[old_sheet];

                DstDataGrid.RefreshData();
                SrcDataGrid.RefreshData();
            }
        }

  
        List<string> GetHeaderStrList(ISheet sheet) {
            List<string> header = new List<string>();

            var row0 = sheet.GetRow(0);
            var row1 = sheet.GetRow(1);
            var row2 = sheet.GetRow(2);

            for (int i = 0; i <row0.Cells.Count;++i) {
                header.Add(string.Concat(Util.GetCellValue(row0.Cells[i]), ":", Util.GetCellValue(row1.Cells[i]),":", Util.GetCellValue(row2.Cells[i])));
            }
            return header;
        }

        // 把第一列认为是id列，检查增删, <value, 行id>
        List<DiffResult<string2int>> GetIDDiffList(ISheet sheet1, ISheet sheet2) {
            var list1 = new List<string2int>();
            var list2 = new List<string2int>();

            for (int i =3; ; i++) {
                var row = sheet1.GetRow(i);
                if (row == null) break;
                list1.Add(new string2int(Util.GetCellValue(row.Cells[0]), i));
            }
           list1.Sort(delegate (string2int a, string2int b) {
                return a.Key.CompareTo(b.Key);
            });
            for (int i = 3; ; i++) {
                var row = sheet2.GetRow(i);
                if (row == null) break;
                list2.Add(new string2int(Util.GetCellValue(row.Cells[0]), i));
            }
            list2.Sort(delegate (string2int a, string2int b) {
                return a.Key.CompareTo(b.Key);
            });

            var option = new DiffOption<string2int>();
            option.EqualityComparer = new SheetIDComparer();
            var result = DiffUtil.Diff(list1, list2, option);
            var optimize = DiffUtil.OptimizeCaseDeletedFirst(result);

            return optimize.ToList();
        }

        List<DiffResult<string>> DiffSheetRow(ISheet sheet1, int row1, ISheet sheet2, int row2, SheetDiffStatus status) {
            var list1 = new List<string>();
            var list2 = new List<string>();

            if (sheet1.GetRow(row1)!=null) {
                var row = sheet1.GetRow(row1);
                for (int i =0; i < status.columnCount;++i) { 
                    list1.Add(Util.GetCellValue(row.GetCell(i)));
                }
            }

            if (sheet2.GetRow(row2) != null) {
                var row = sheet2.GetRow(row2);
                for (int i = 0; i < status.columnCount; ++i) {
                    list2.Add(Util.GetCellValue(row.GetCell(i)));
                }
            }
            var diff = DiffUtil.Diff(list1, list2);
            var optimized = DiffUtil.OptimizeCaseDeletedFirst(diff);
            optimized = DiffUtil.OptimizeCaseInsertedFirst(optimized);

            return optimized.ToList();
        }

        void OnSheetChanged() {
            // TODO这里检查diff标记，清理book1


        }

        private void DstFileSheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (e.AddedItems.Count > 0) {
                var selection = (e.AddedItems[0] as ComboBoxItem).Content as SheetNameCombo;
                books["dst"].sheet = selection.ID;
                
                if (books.ContainsKey("src") && books["src"].sheet != selection.ID) {
                    var idx = books["src"].ComboIdToItemIdx[selection.ID];
                    SrcFileSheetsCombo.SelectedItem = books["src"].sheetCombo[idx];
                } else {
                    OnSheetChanged();
                }
                DstDataGrid.RefreshData();
            }
        }

        private void SrcFileSheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (e.AddedItems.Count > 0) {
                var selection = (e.AddedItems[0] as ComboBoxItem).Content as SheetNameCombo;
                books["src"].sheet = selection.ID;
   
                if (books.ContainsKey("dst") && books["dst"].sheet != selection.ID) {
                    var idx = books["dst"].ComboIdToItemIdx[selection.ID];
                    DstFileSheetsCombo.SelectedItem = books["dst"].sheetCombo[idx];
                } else {
                    OnSheetChanged();
                }

                SrcDataGrid.RefreshData();
            }
        }

        private void SVNResivionionList_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            var selection = e.AddedItems[0] as SvnRevisionCombo;
            SVNRevisionCombo.Width = selection.Revision.Length*10;

            Diff(selection.ID - 1, selection.ID);
        }

        public void OnGridScrollChanged(string tag, ScrollChangedEventArgs e) {
            ScrollViewer view = null;
            if (tag == "src") {
                view = Util.GetVisualChild<ScrollViewer>(DstDataGrid);
            } else if (tag == "dst") {
                view = Util.GetVisualChild<ScrollViewer>(SrcDataGrid);
            }
            view.ScrollToVerticalOffset(e.VerticalOffset);
            view.ScrollToHorizontalOffset(e.HorizontalOffset);
        }
    }


}
