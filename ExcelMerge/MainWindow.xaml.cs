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
using NPOI.SS.Util;
using SharpSvn;
using SharpSvn.UI;
using System.Collections.ObjectModel;
using NetDiff;
using string2int = System.Collections.Generic.KeyValuePair<string, int>;
using System.IO;
using Newtonsoft.Json;

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

        public List<string> _tempFiles = new List<string>();


        public Mode mode = Mode.Diff;

        public class Config {
            public List<string> NoHeadPaths = new List<string>();
        }

        static string ConfigPath = "config.json";

        public Config config;

        public MainWindow() {
            InitializeComponent();

            if (System.ComponentModel.DesignerProperties.GetIsInDesignMode(this)) {
                return;
            }


            var path = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, ConfigPath);
            if (!File.Exists(path)) {
                config = new Config();
                File.WriteAllText(path, JsonConvert.SerializeObject(config));
            } else {
                config = JsonConvert.DeserializeObject<Config>(File.ReadAllText(path));
            }
            instance = this;
        }

        public void DataGrid_SelectedCellsChanged(object sender, SelectionChangedEventArgs e) {

        }

        // load进来单个文件的情况
        public void OnFileLoaded(string file, string tag, FileOpenType type, int sheet = 0) {
            file = file.Replace("\\", "/");
            foreach (var reg in config.NoHeadPaths) {
                if (System.Text.RegularExpressions.Regex.Match(file, reg).Length > 0) {
                    ProcessHeader.IsChecked = false;
                }
            }

            var wb = Util.GetWorkBook(file);

            books[tag] = new WorkBookWrap() { book = wb, sheet = sheet, file = file, filename = System.IO.Path.GetFileName(file) };

            if (type == FileOpenType.Drag) {
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

        internal void CopyCellsValue(string v, string otherTag, IList<DataGridCellInfo> selectCells) {
            var srcSheet = books[v].GetCurSheet();
            var dstSheet = books[otherTag].GetCurSheet();

            foreach (var cell in selectCells) {
                var rowdata = cell.Item as ExcelData;
                var column = cell.Column.DisplayIndex;
                var rowid = rowdata.rowId;

                var row = srcSheet.GetRow(rowid);

                Util.CopyCell(row.GetCell(column), dstSheet.GetRow(rowid).GetCell(column));

                var diffidx = rowdata.diffIdx;

                OnCellEdited(otherTag, rowid, column, CellEditMode.Self);
                OnCellEdited(v, rowid, column, CellEditMode.OtherSide);
            }

            RefreshCurSheet();
        }

        internal void SetCellValue(string v, ICell targetCell) {
            targetCell.SetCellValue( v);
        }

        void UpdateSVNRevision(string file, string tag) {
            if (tag == "src") {
                Collection<SvnLogEventArgs> logitems;

                DateTime startDateTime = DateTime.Now.AddDays(-60);
                DateTime endDateTime = DateTime.Now;
                var svnRange = new SvnRevisionRange(new SvnRevision(startDateTime), new SvnRevision(endDateTime));

                List<SvnRevisionCombo> revisions = new List<SvnRevisionCombo>();

                using (SvnClient client = new SvnClient()) {
                    client.Authentication.SslServerTrustHandlers += delegate (object sender, SharpSvn.Security.SvnSslServerTrustEventArgs e) {
                        e.AcceptedFailures = e.Failures;
                        e.Save = true; // Save acceptance to authentication store
                    };

                    if (client.GetUriFromWorkingCopy(file) != null) {

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
                        revisions.Sort((a, b) => {
                            return b.ID - a.ID;
                        });
                    }
                }
                SVNRevisionCombo.ItemsSource = revisions;
            }
        }

        WorkBookWrap InitWorkWrap(string file) {
            var wb = new WorkBookWrap() {
                book = Util.GetWorkBook(file),
                file = file,
                filename = System.IO.Path.GetFileName(file)
            };

            wb.sheetCombo = new List<ComboBoxItem>();
            var list = new List<SheetNameCombo>();
            for (int i = 0; i < wb.book.NumberOfSheets; ++i) {
                list.Add(new SheetNameCombo() { Name = wb.book.GetSheetName(i), ID = i });
            }
            list.Sort((a, b) => { return a.Name.CompareTo(b.Name); });

            wb.sheetNameCombos = list;

            wb.ItemID2ComboIdx = new Dictionary<int, int>();

            list.ForEach((a) => { var item = new ComboBoxItem(); item.Content = a; wb.sheetCombo.Add(item); });

            for (int i = 0; i < list.Count;++i) {
                wb.ItemID2ComboIdx[list[i].ID] = i;
            }

            wb.SheetValideRow = new Dictionary<string, int>();

            return wb;
        }

        int[] getColumn2Diff(List<DiffResult<string>> diff, bool from, int count) {
            int idx = 0;
            var ret = new int[diff.Count];
            for (int i = 0; i < diff.Count; ++i) {
                ret[idx] = i;
                if (from) {
                    if (diff[i].Status != DiffStatus.Inserted) {
                        idx++;
                    }
                } else {
                    if (diff[i].Status != DiffStatus.Deleted) {
                        idx++;
                    }
                }
            }
            return ret;
        }

        SheetDiffStatus DiffSheet(ISheet src, ISheet dst, SheetDiffStatus status = null) {
            status = status??new SheetDiffStatus();

            bool changed = false;

            var head1 = GetHeaderStrList(src);
            var head2 = GetHeaderStrList(dst);
            if (head1 == null || head2 == null) return null;

            var diff = NetDiff.DiffUtil.Diff(head1, head2);
            var optimized = diff.ToList();// NetDiff.DiffUtil.OptimizeCaseDeletedFirst(diff);
            //optimized = DiffUtil.OptimizeCaseDeletedFirst(diff);

            changed = changed || optimized.Any(a => a.Status != DiffStatus.Equal);

            status.diffHead = optimized.ToList();
            status.column2diff1 = new Dictionary<int, int[]>();
            status.column2diff2 = new Dictionary<int, int[]>();
            status.column2diff1[0] = getColumn2Diff(status.diffHead, true, head1.Count);
            status.column2diff2[0] = getColumn2Diff(status.diffHead, false, head2.Count);

            status.columnCount1 =  head1.Count;
            status.columnCount2 = head2.Count;
            
            status.diffFistColumn = GetIDDiffList(src, dst, 1);

            changed = changed || status.diffFistColumn.Any(a => a.Status != DiffStatus.Equal);

            status.diffSheet = new List<List<DiffResult<string>>>();
            status.rowID2DiffMap1 = new Dictionary<int, int>();
            status.rowID2DiffMap2 = new Dictionary<int, int>();
            status.Diff2RowID1 = new Dictionary<int, int>();
            status.Diff2RowID2 = new Dictionary<int, int>();
            status.RowEdited1 = status.RowEdited1?? new Dictionary<int, Dictionary<int, CellEditMode>>();
            status.RowEdited2 = status.RowEdited2?? new Dictionary<int, Dictionary<int, CellEditMode>>();

            foreach (var diffkv in status.diffFistColumn) {
                var rowid1 = diffkv.Obj1.Value;
                var rowid2 = diffkv.Obj2.Value;
                if (diffkv.Obj1.Key == null) {
                    // 创建新行，方便比较
                    rowid1 = -1;
                }
                if (diffkv.Obj2.Key == null) {
                    rowid2 = -1;
                }
                var diffrow = DiffSheetRow(src, rowid1, dst, rowid2, status);

                if (diffkv.Obj1.Key == null) {
                    // 创建新行，方便比较,放在后面是为了保证diff的时候是new,delete的形式，而不是modify
                    rowid1 =  books["src"].SheetValideRow[src.SheetName] + 1;
                    src.CreateRow(rowid1);
                }
                if (diffkv.Obj2.Key == null) {
                    rowid2 = books["dst"].SheetValideRow[dst.SheetName] + 1;
                    dst.CreateRow(rowid2);
                }
                status.column2diff1[rowid1] = getColumn2Diff(diffrow, true, status.columnCount1);
                status.column2diff2[rowid2] = getColumn2Diff(diffrow, false, status.columnCount2);

                int diffIdx = status.diffSheet.Count;

                status.rowID2DiffMap1[rowid1] = diffIdx;
                status.rowID2DiffMap2[rowid2] = diffIdx;

                status.Diff2RowID1[diffIdx] = rowid1;
                status.Diff2RowID2[diffIdx] = rowid2;

                if (!status.RowEdited1.ContainsKey(rowid1)) {
                    status.RowEdited1[rowid1] = new Dictionary<int, CellEditMode>();
                }
                if (!status.RowEdited2.ContainsKey(rowid2)) {
                    status.RowEdited2[rowid2] = new Dictionary<int, CellEditMode>();
                }

                status.diffSheet.Add(diffrow);
                
                changed = changed || diffrow.Any(a => a.Status != DiffStatus.Equal);

                if (changed) {
                    changed = true;
                }
            }

            status.changed = changed;

            return status;
        }
        

        public void Diff(string file1, string file2, bool resetInitFile = true) {
            if (string.IsNullOrEmpty(file1) || string.IsNullOrEmpty(file2)) return;

            if (resetInitFile) {
                SrcFile = file1;
                DstFile = file2;
            }

            string oldsheetName = null;
            if (books.ContainsKey("src")) {
                oldsheetName = books["src"].sheetname;
            }

            var src = InitWorkWrap(file1);
            var dst = InitWorkWrap(file2);


            var option = new DiffOption<SheetNameCombo>();
            option.EqualityComparer = new SheetNameComboComparer();
            var result = DiffUtil.Diff(src.sheetNameCombos, dst.sheetNameCombos, option);
            diffSheetName = result.ToList();// DiffUtil.OptimizeCaseDeletedFirst(result).ToList();

            books["src"] = src;
            books["dst"] = dst;
            var srcSheetID = -1;
            var dstSheetID = -1;

            for (int i = 0; i < diffSheetName.Count; ++i) {
                var sheetname = diffSheetName[i];

                // 只有sheet名字一样的可以diff， 先这么处理
                if (sheetname.Status == DiffStatus.Equal) {
                    var sheet1 = sheetname.Obj1.ID;
                    var sheet2 = sheetname.Obj2.ID;
                    
                    sheetsDiff[i] = DiffSheet(src.book.GetSheetAt(sheet1), dst.book.GetSheetAt(sheet2));

                    if (sheetsDiff[i] != null) {
                        oldsheetName = sheetname.Obj1.Name;
                        var sheetidx = 0;
                        if (!string.IsNullOrEmpty(oldsheetName)) {
                            sheetidx = src.book.GetSheetIndex(oldsheetName);
                        }
                        if (sheetsDiff[i].changed || srcSheetID == -1) {
                            src.sheet = sheetidx;
                            srcSheetID = sheetidx;
                        }

                        if (!string.IsNullOrEmpty(oldsheetName)) {
                            sheetidx = dst.book.GetSheetIndex(oldsheetName);
                        }
                        if (sheetsDiff[i].changed || dstSheetID == -1) {
                            dst.sheet = sheetidx;
                            dstSheetID = sheetidx;
                        }
                    }
                }
            }

            // refresh ui
            SrcFilePath.Content = file1;
            DstFilePath.Content = file2;

            SrcFileSheetsCombo.Items.Clear();
            foreach (var item in src.sheetCombo) {

                int index = diffSheetName.FindIndex(a => a.Obj1 != null && a.Obj1.ID == (item.Content as SheetNameCombo).ID);
                SolidColorBrush color = null;
                DiffStatus status = diffSheetName[index].Status;
                if (status != DiffStatus.Equal) {
                    color = Util.GetColorByDiffStatus(status);
                }
                else {
                    color = Util.GetColorByDiffStatus(sheetsDiff.ContainsKey(index) && sheetsDiff[index]!=null && sheetsDiff[index].changed ? DiffStatus.Modified : DiffStatus.Equal);
                }

                if (color != null) {
                    item.Background = color;
                }

                SrcFileSheetsCombo.Items.Add(item);
            }
            var comboidx = src.ItemID2ComboIdx[src.sheet];
            SrcFileSheetsCombo.SelectedItem = src.sheetCombo[comboidx];

            DstFileSheetsCombo.Items.Clear();
            foreach (var item in dst.sheetCombo) {

                int index = diffSheetName.FindIndex(a => a.Obj2 != null && a.Obj2.ID == (item.Content as SheetNameCombo).ID);
                SolidColorBrush color = null;
                DiffStatus status = diffSheetName[index].Status;
                if (status != DiffStatus.Equal) {
                    color = Util.GetColorByDiffStatus(status);
                }
                else {
                    color = Util.GetColorByDiffStatus(sheetsDiff.ContainsKey(index) && sheetsDiff[index] != null && sheetsDiff[index].changed ? DiffStatus.Modified : DiffStatus.Equal);
                }

                if (color != null) {
                    item.Background = color;
                }

                DstFileSheetsCombo.Items.Add(item);
            }
            comboidx = dst.ItemID2ComboIdx[dst.sheet];
            DstFileSheetsCombo.SelectedItem = dst.sheetCombo[comboidx];

            DstDataGrid.RefreshData();
            SrcDataGrid.RefreshData();
        }

        public int DiffStartIdx() {
            // 首三行一起作为key
            return ProcessHeader.IsChecked == true ? 3 : 0;
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

                _tempFiles.Add(file1);
                _tempFiles.Add(file2);
                Diff(file1, file2, false);
            }
        }
    
        public void RefreshCurSheet() {
            Dispatcher.BeginInvoke(new Action(ReDiff));
        }

        void ReDiff() {
            var src_sheet = books["src"].sheet;
            int index = diffSheetName.FindIndex(a => a.Obj1 != null && a.Obj1.ID == src_sheet);
            
            DiffSheet(books["src"].GetCurSheet(), books["dst"].GetCurSheet(), sheetsDiff[index]);
  
            DstDataGrid.RefreshData();
            SrcDataGrid.RefreshData();
        }

        public void OnCellEdited(string tag, int rowid, int columnid, CellEditMode mode) {
            Dictionary<int, Dictionary<int, CellEditMode>> edited;
            if (tag == "src") {
                var src_sheet = books["src"].sheet;
                int index = diffSheetName.FindIndex(a => a.Obj1 != null && a.Obj1.ID == src_sheet);

                edited = sheetsDiff[index].RowEdited1;
            } else {
                var src_sheet = books["dst"].sheet;
                int index = diffSheetName.FindIndex(a => a.Obj2 != null && a.Obj2.ID == src_sheet);

                edited = sheetsDiff[index].RowEdited2;
            }
            edited[rowid][columnid] = mode;
        }
  
        List<string> GetHeaderStrList(ISheet sheet) {
            List<string> header = new List<string>();

            if (ProcessHeader.IsChecked == true) {
                var list = new List<IRow>();
                for (int i = 0; i < DiffStartIdx(); ++i) {
                    var row = sheet.GetRow(i);
                    if (row == null) return null;
                    list.Add(row);
                }
                
                for (int i = 0; i < list[0].Cells.Count; ++i) {
                    var str = "";
                    for (int j = 0; j < DiffStartIdx(); ++j) {
                        var cell_s = Util.GetCellValue(list[j].GetCell(i));
                        if (j == 0 && string.IsNullOrWhiteSpace(cell_s)) {
                            return header;
                        }
                        str = str + (j > 0 ? ":" + cell_s : cell_s);
                    }
                   
                    header.Add(str);
                }
            } else {
                var row0 = sheet.GetRow(0);
                if (row0 == null ) return null;

                for (int i = 0; i < row0.Cells.Count; ++i) {
                    var s1 = Util.GetCellValue(row0.GetCell(i));
                    if (string.IsNullOrWhiteSpace(s1)) {
                        return header;
                    }
                    header.Add((i+1).ToString());
                }
            }
            return header;
        }

        // 把第一列认为是id列，检查增删, <value, 行id>
        List<DiffResult<string2int>> GetIDDiffList(ISheet sheet1, ISheet sheet2, int checkCellCount, bool addRowID = false) {
            var list1 = new List<string2int>();
            var list2 = new List<string2int>();

            var nameHash = new HashSet<string>();

            var startIdx = DiffStartIdx();

            // 尝试找一个id不会重复的前几列的值作为key
            for (int i = startIdx; ; i++) {
                var row = sheet1.GetRow(i);
                if (row == null || !Util.CheckValideRow(row)) {
                    books["src"].SheetValideRow[sheet1.SheetName] = i;
                    break;
                };
 
                var val = "";
                for (var j = 0; j < checkCellCount; ++j) {
                    val += Util.GetCellValue(row.GetCell(j));
                }
                var hash_val = val;
                if (addRowID) {
                    hash_val = hash_val + "." + i;
                }
                if (nameHash.Contains(hash_val)) {
                    if (checkCellCount < 6) {
                        return GetIDDiffList(sheet1, sheet2, checkCellCount + 1, addRowID);
                    } else {
                        // 已经找不到能作为key的了。把id和行号连一块
                        return GetIDDiffList(sheet1, sheet2, 1, true);
                    }
                } 

                nameHash.Add(hash_val);

                list1.Add(new string2int(val, i));
            }
           list1.Sort(delegate (string2int a, string2int b) {
               var cmp = a.Key.CompareTo(b.Key);
               if (cmp == 0) {
                   return a.Value.CompareTo(b.Value);
               }
               return cmp;
           });
            nameHash.Clear();
            for (int i = startIdx; ; i++) {
                var row = sheet2.GetRow(i);
                if (row == null || !Util.CheckValideRow(row)) {
                    books["dst"].SheetValideRow[sheet2.SheetName] = i;
                    break;
                }
                var val = "";
                for (var j = 0; j < checkCellCount; ++j) {
                    val += Util.GetCellValue(row.GetCell(j));
                }
                var hash_val = val;
                if (addRowID) {
                    hash_val = hash_val + "." + i;
                }
                if (nameHash.Contains(hash_val)) {
                    if (checkCellCount < 6) {
                        return GetIDDiffList(sheet1, sheet2, checkCellCount + 1, addRowID);
                    }
                    else {
                        // 已经找不到能作为key的了。把id和行号连一块
                        return GetIDDiffList(sheet1, sheet2, 1, true);
                    }
                }
                nameHash.Add(hash_val);

                list2.Add(new string2int(val, i));
            }
            list2.Sort(delegate (string2int a, string2int b) {
                var cmp = a.Key.CompareTo(b.Key);
                if (cmp == 0) {
                    return a.Value.CompareTo(b.Value);
                }
                return cmp;
            });

            var option = new DiffOption<string2int>();
            option.Optimize = false;
            option.EqualityComparer = new SheetIDComparer();
            var result = DiffUtil.Diff(list1, list2, option);
            var optimize = result.ToList();// DiffUtil.OptimizeCaseDeletedFirst(result);

            return optimize.ToList();
        }

        List<DiffResult<string>> DiffSheetRow(ISheet sheet1, int row1, ISheet sheet2, int row2, SheetDiffStatus status) {
            var list1 = new List<string>();
            var list2 = new List<string>();

            if (sheet1.GetRow(row1)!=null) {
                var row = sheet1.GetRow(row1);
                for (int i =0; i < status.columnCount1;++i) { 
                    list1.Add(Util.GetCellValue(row.GetCell(i)));
                }
            }

            if (sheet2.GetRow(row2) != null) {
                var row = sheet2.GetRow(row2);
                for (int i = 0; i < status.columnCount2; ++i) {
                    list2.Add(Util.GetCellValue(row.GetCell(i)));
                }
            }
            var diff = DiffUtil.Diff(list1, list2);
            var optimized = diff.ToList();// DiffUtil.OptimizeCaseDeletedFirst(diff);
            //optimized = DiffUtil.OptimizeCaseInsertedFirst(optimized);

            return optimized.ToList();
        }

        void OnSheetChanged() {
            // TODO这里检查diff标记，清理book1


        }

        private void DstFileSheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (e.AddedItems.Count > 0) {
                var selection = (e.AddedItems[0] as ComboBoxItem).Content as SheetNameCombo;
                books["dst"].sheet = selection.ID;
                

                if (books.ContainsKey("src") && books["src"].sheetname != books["dst"].sheetname) {
                    var idx = books["src"].GetComboIDBySheetName(books["dst"].sheetname);
                    if (idx >= 0) {
                        SrcFileSheetsCombo.SelectedItem = books["src"].sheetCombo[idx];
                    }
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
   
                if (books.ContainsKey("dst") && books["src"].sheetname != books["dst"].sheetname) {
                    var idx = books["dst"].GetComboIDBySheetName(books["src"].sheetname);
                    if (idx >= 0) {
                        DstFileSheetsCombo.SelectedItem = books["dst"].sheetCombo[idx];
                    }
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

        public void OnSelectGridRow(string tag, int rowid) {
            if (tag == "src") {
                DstDataGrid.ExcelGrid.SelectedIndex = rowid;
            }
            else{
                SrcDataGrid.ExcelGrid.SelectedIndex = rowid;
            }
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e) {
            if ((sender as RadioButton).Content as string == Mode.Diff.ToString()) {
                mode = Mode.Diff;
            } else {
                mode = Mode.Merge;
            }
        }

        private void DoDiff_Click(object sender, RoutedEventArgs e) {
            Diff(SrcFile, DstFile);
        }

        private void ApplyChange_Click(object sender, RoutedEventArgs e) {
            var oldfile = books["dst"].file;
            var filepath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(oldfile), System.IO.Path.GetFileNameWithoutExtension(oldfile) + "_apply.xls");
            System.IO.File.Copy(oldfile, filepath, true);

            using (var dstfile = File.Open(filepath, FileMode.OpenOrCreate, FileAccess.Write)) {

                books["dst"].book.Write(dstfile);

                dstfile.Flush();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
            foreach (var file in _tempFiles) {
                if (File.Exists(file)) {
                    File.Delete(file);
                }
            }
        }

        private void SimpleHeader_Checked(object sender, RoutedEventArgs e) {
            Diff(SrcFile, DstFile);
        }
    }


}
