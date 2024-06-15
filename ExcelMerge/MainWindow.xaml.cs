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
using System.Collections.ObjectModel;
using NetDiff;
using string2int = System.Collections.Generic.KeyValuePair<string, int>;
using System.IO;
using Newtonsoft.Json;
using Microsoft.Win32;
using NPOI.Util;

namespace ExcelMerge {

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        public static MainWindow instance;
        //"dst","src"
        public Dictionary<string, WorkBookWrap> books = new Dictionary<string, WorkBookWrap>();

        public Dictionary<string, SheetDiffStatus> sheetsDiff = new Dictionary<string, SheetDiffStatus>();
        
        public List<DiffResult<SheetNameCombo>> diffSheetName;

        public Dictionary<string, Dictionary<int, ExcelData>> excelGridData = new Dictionary<string, Dictionary<int, ExcelData>>();

        public DirectoryDifferWindow dirWindow;

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

            //dirWindow = new DirectoryDifferWindow();
            //dirWindow.Show();
            if (config.NoHead) {
                ProcessHeader.IsChecked = false;
            }

            try {
                var key = Registry.ClassesRoot.CreateSubKey(@"xlsmerge");
                key = Registry.ClassesRoot.OpenSubKey("xlsmerge", true);
                key.SetValue("URL Protocol", "");
                key.SetValue(null, "URL:xlsmerge");
                key.Close();
                key = Registry.ClassesRoot.CreateSubKey(@"xlsmerge\shell\open\command");
                key = Registry.ClassesRoot.OpenSubKey(@"xlsmerge\shell\open\command", true);
                var dir = System.AppDomain.CurrentDomain.BaseDirectory;
                dir.Replace("/", "\\");
                dir = string.Format("\"{0}ExcelMerge.exe\" \"%1\"", dir);
                key.SetValue(null, dir);
                key.Close();
            } catch {

            } finally {
                var key = Registry.ClassesRoot.OpenSubKey(@"xlsmerge\shell\open\command");
                if (key != null) {
                    Title = "ExcelMerge " + "[已绑定]";
                }
            }

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        private void HandleEsc(object sender, KeyEventArgs e) {
            if (e.Key == Key.Escape)
                Close();
        }

        public void ShowDirectoryWindow(string[] dirs, string tag) {
            dirWindow = dirWindow ?? new DirectoryDifferWindow();
            dirWindow.Show();

            dirWindow.OnSetDirs(dirs, tag);
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

                            revisions.Add(new SvnRevisionCombo() { Revision = string.Format("{0}[{1}]", author, message), ID = logentry.Revision });
                        }
                        revisions.Sort((a, b) => {
                            return (int)(b.ID - a.ID);
                        });
                    }
                }
                SVNRevisionCombo.ItemsSource = revisions;
            }
        }


        int[] getColumn2Diff(List<DiffResult<string>> diff, bool from) {
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

        // 对比两个sheet
        SheetDiffStatus DiffSheet(ISheet src, ISheet dst, SheetDiffStatus status = null) {
            status = status??new SheetDiffStatus() { sortKey = config.DefaultKeyID };

            bool changed = false;

            var srcwrap = books["src"];
            var dstwrap = books["dst"];

            var head1 = GetHeaderStrList(srcwrap, src);
            var head2 = GetHeaderStrList(dstwrap, dst);
            if (head1 == null || head2 == null) return null;

            var diff = NetDiff.DiffUtil.Diff(head1, head2);
            //var optimized = diff.ToList();// NetDiff.DiffUtil.OptimizeCaseDeletedFirst(diff);
            var optimized = DiffUtil.OptimizeCaseDeletedFirst(diff);

            changed = changed || optimized.Any(a => a.Status != DiffStatus.Equal);

            var diffhead = optimized.ToList();
            status.diffHead = new SheetRowDiff() { diffcells = diffhead };

            status.column2diff1[0] = getColumn2Diff(diffhead, true);
            status.column2diff2[0] = getColumn2Diff(diffhead, false);

            srcwrap.SheetValideColumn[src.SheetName] = head1.Count;
            dstwrap.SheetValideColumn[dst.SheetName] = head2.Count;
            
            status.diffFistColumn = GetIDDiffList(src, dst, 1, false, status.sortKey);

            changed = changed || status.diffFistColumn.Any(a => a.Status != DiffStatus.Equal);

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
                int maxLineCount = 0;
                var diffrow = DiffSheetRow(src, rowid1, dst, rowid2, status, out maxLineCount);

                if (diffkv.Obj1.Key == null) {
                    // 创建新行，方便比较,放在后面是为了保证diff的时候是new,delete的形式，而不是modify
                    rowid1 =  books["src"].SheetValideRow[src.SheetName];
                    //src.CreateRow(rowid1);
                }
                if (diffkv.Obj2.Key == null) {
                    rowid2 = books["dst"].SheetValideRow[dst.SheetName];
                    //dst.CreateRow(rowid2);
                }
                status.column2diff1[rowid1] = getColumn2Diff(diffrow, true);
                status.column2diff2[rowid2] = getColumn2Diff(diffrow, false);

                int diffIdx = status.diffSheet.Count;
                status.DiffMaxLineCount[diffIdx] = maxLineCount;

                status.rowID2DiffMap1[rowid1] = diffIdx;
                status.rowID2DiffMap2[rowid2] = diffIdx;

                status.Diff2RowID1[diffIdx] = rowid1;
                status.Diff2RowID2[diffIdx] = rowid2;

                var rowdiff = new SheetRowDiff();
                rowdiff.diffcells = diffrow;

                rowdiff.changed = diffrow.Any(a => a.Status != DiffStatus.Equal);
                if (rowdiff.changed) {
                    rowdiff.diffcell_details = new List<List<DiffResult<char>>>();
                    foreach (var cell in diffrow) {
                        if (cell.Status == DiffStatus.Modified) {
                            var cell_diff = NetDiff.DiffUtil.Diff(cell.Obj1, cell.Obj2);
                            //var optimized = diff.ToList();// NetDiff.DiffUtil.OptimizeCaseDeletedFirst(diff);
                            var opt_cell_diff = DiffUtil.OptimizeCaseDeletedFirst(cell_diff);

                            rowdiff.diffcell_details.Add(opt_cell_diff.ToList());
                        } else {
                            rowdiff.diffcell_details.Add(null);
                        }
                    }
                }
                status.diffSheet.Add(rowdiff);
                
                changed = changed || rowdiff.changed;
            }

            status.changed = changed;

            return status;
        }
        
        public void Refresh() {
            var file1 = Entrance.SrcFile;
            var file2 = Entrance.DstFile;

            if (string.IsNullOrEmpty(file1) || string.IsNullOrEmpty(file2)) return;


            string oldsheetName = null;
            if (books.ContainsKey("src")) {
                oldsheetName = books["src"].sheetname;
            }

            var src = new WorkBookWrap(file1, config);
            var dst = new WorkBookWrap(file2, config);

            var option = new DiffOption<SheetNameCombo>();
            option.EqualityComparer = new SheetNameComboComparer();
            var result = DiffUtil.Diff(src.sheetNameCombos, dst.sheetNameCombos, option);
            //diffSheetName = result.ToList();//
            diffSheetName = DiffUtil.OptimizeCaseDeletedFirst(result).ToList();
            books["src"] = src;
            books["dst"] = dst;
            var srcSheetID = -1;
            var dstSheetID = -1;

            for (int i = 0; i < diffSheetName.Count; ++i) {
                var sheetname = diffSheetName[i];
                var name = sheetname.Obj1 == null ? sheetname.Obj2.Name : sheetname.Obj1.Name;

                // 只有sheet名字一样的可以diff， 先这么处理
                if (sheetname.Status == DiffStatus.Equal) {
                    var sheet1 = sheetname.Obj1.ID;
                    var sheet2 = sheetname.Obj2.ID;
                    
                    sheetsDiff[name] = DiffSheet(src.book.GetSheetAt(sheet1), dst.book.GetSheetAt(sheet2));

                    if (sheetsDiff[name] != null) {
                        oldsheetName = sheetname.Obj1.Name;
                        var sheetidx = 0;
                        if (!string.IsNullOrEmpty(oldsheetName)) {
                            sheetidx = src.book.GetSheetIndex(oldsheetName);
                        }
                        if (sheetsDiff[name].changed || srcSheetID == -1) {
                            src.sheet = sheetidx;
                            srcSheetID = sheetidx;
                        }

                        if (!string.IsNullOrEmpty(oldsheetName)) {
                            sheetidx = dst.book.GetSheetIndex(oldsheetName);
                        }
                        if (sheetsDiff[name].changed || dstSheetID == -1) {
                            dst.sheet = sheetidx;
                            dstSheetID = sheetidx;
                        }
                    }
                } else
                {
                    // 新增sheet，直接完整显示
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
                    var name = diffSheetName[index].Obj1.Name;
                    color = Util.GetColorByDiffStatus(sheetsDiff.ContainsKey(name) && sheetsDiff[name] !=null && sheetsDiff[name].changed ? DiffStatus.Modified : DiffStatus.Equal);
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
                    var name = diffSheetName[index].Obj1.Name;
                    color = Util.GetColorByDiffStatus(sheetsDiff.ContainsKey(name) && sheetsDiff[name] != null && sheetsDiff[name].changed ? DiffStatus.Modified : DiffStatus.Equal);
                }

                if (color != null) {
                    item.Background = color;
                }

                DstFileSheetsCombo.Items.Add(item);
            }
            comboidx = dst.ItemID2ComboIdx[dst.sheet];
            DstFileSheetsCombo.SelectedItem = dst.sheetCombo[comboidx];

            //DstDataGrid.RefreshData();
            //SrcDataGrid.RefreshData();

            //OnSheetChanged();
        }

        public int DiffStartIdx(int emptyline) {
            // 首三行一起作为key
            return ProcessHeader.IsChecked == true ? config.HeadCount+ emptyline : emptyline;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Entrance.Window_Closing(this, e);
        }

        bool CheckIfVersionExists(SvnClient client, Uri uri, long revision, string sheet_name, string id, string header, string value)
        {
            var file = Entrance.GetVersionFile(client, uri, revision);

            var wrap = new WorkBookWrap(file, config);

            var startpoint = wrap.SheetStartPoint[sheet_name];
            var rowStart = startpoint.Item1;
            var columnStart = startpoint.Item2;

            var headers = wrap.SheetHeaders[sheet_name];
            var header_idx = headers.IndexOf(header);

            var ids = wrap.SheetIDs[sheet_name];
            var id_idx = ids.IndexOf(id);

            if (header_idx >= 0 && id_idx >= 0)
            {
                var sheet = wrap.book.GetSheet(sheet_name);
                var row = sheet.GetRow(id_idx + rowStart);
                var val = Util.GetCellValue( row.GetCell(header_idx + columnStart));
                if (val == value)
                {
                    return true;
                }
            }

            return false;
        }

        public void FindCellEdit(WorkBookWrap wrap, int rowid, int col_id)
        {
  
            using (SvnClient client = new SvnClient())
            {
                SvnInfoEventArgs info;
                client.GetInfo(wrap.file, out info);
                var uri = info.Uri;
                var sheetname = wrap.sheetname;

                var startpoint = wrap.SheetStartPoint[sheetname];
                var startrow = startpoint.Item1;
                var startcol = startpoint.Item2;

                string col_name = "", value = "" ;
                string id = wrap.SheetIDs[sheetname][rowid-startrow];
                if (col_id >= 0)
                {
                    col_name = wrap.SheetHeaders[sheetname][col_id];
                }

                var fileversion = new Collection<SvnFileVersionEventArgs>();
                client.GetFileVersions(uri, new SvnFileVersionsArgs() { Start = 0L }, out fileversion);

                var left = 0;
                var right = fileversion.Count()-1;

                long  revision = 0;
                while (left < right)
                {
                    var mid = (left+right) /2;
 
                    var exist = CheckIfVersionExists(client, uri, fileversion[mid].Revision, sheetname, id, col_name, value);
                    if (!exist)
                    {
                        if (mid == right-1)
                        {
                            revision = right;
                            break;
                        }
                        right = mid;
                    } else
                    {
                        left = mid;
                    }
                }
                if (revision > 0)
                {
                    Entrance.DiffUri(revision - 1, revision, uri);
                }
            }
        }

        public void RefreshCurSheet() {
            Dispatcher.BeginInvoke(new Action(ReDiffCurSheet));
        }

        void ReDiffCurSheet() {
            var src_sheet = books["src"].sheetname;
            
            DiffSheet(books["src"].GetCurSheet(), books["dst"].GetCurSheet(), sheetsDiff[src_sheet]);
  
            DstDataGrid.RefreshData();
            SrcDataGrid.RefreshData();
        }

        public void ReDiffFile() {
            Refresh();
        }
  
        List<string> GetHeaderStrList(WorkBookWrap wrap, ISheet sheet) {
            List<string> header = new List<string>();

            var startpoint = wrap.SheetStartPoint[sheet.SheetName];
            var startrow = startpoint.Item1;
            var startcol = startpoint.Item2;
            if (ProcessHeader.IsChecked == true) {
                var list = new List<IRow>();
                for (int i = startrow; i < DiffStartIdx(startrow); ++i) {
                    var row = sheet.GetRow(i);
                    if (row == null) continue;
                    list.Add(row);
                }
                
                if (list.Count == 0)
                {
                    return null;
                }
                for (int i = startcol; i < list[0].Cells.Count; ++i) {
                    var str = "";
                    for (int j = 0; j < list.Count; ++j) {
                        var cell_s = Util.GetCellValue(list[j].GetCell(i));
                   
                        str = str + (j > 0 ? ":" + cell_s : cell_s);
                    }
                    if (string.IsNullOrWhiteSpace(str))
                    {
                        return header;
                    }
                    header.Add(str);
                }
            } else {
                var row0 = sheet.GetRow(startrow);
                if (row0 == null ) return null;

                for (int i = startcol; i < row0.Cells.Count; ++i) {
                    var s1 = Util.GetCellValue(row0.GetCell(i));
                    // 起码有两列
                    if (string.IsNullOrWhiteSpace(s1) && i > 1) {
                        return header;
                    }
                    header.Add((i+1).ToString());
                }
            }
            return header;
        }
        enum SearchStatus { Succ, NextCol, Fail };
        // 把第一列认为是id列，检查增删, <value, 行id>
        List<DiffResult<string2int>> GetIDDiffList(ISheet sheet1, ISheet sheet2, int checkCellCount, bool addRowID = false, int startCheckCell=0) {
 
     
            bool allNum = checkCellCount==1;

            Func<WorkBookWrap, ISheet, List<string2int>, SearchStatus> search = (WorkBookWrap wrap, ISheet sheet, List<string2int> list) => {
                var nameHash = new HashSet<string>();
                var startrow = wrap.SheetStartPoint[sheet.SheetName].Item1;
                var startIdx = DiffStartIdx(startrow);
                int ignoreEmptyLine = config.EmptyLine;
                // 尝试找一个id不会重复的前几列的值作为key
                for (int i = startIdx; ; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null || !Util.CheckValideRow(row))
                    {
                        break;
                    };

                    var val = "";
                    for (var j = startCheckCell; j < startCheckCell + checkCellCount; ++j)
                    {
                        if (row.GetCell(j) == null || row.GetCell(j).CellType != CellType.Numeric)
                        {
                            allNum = false;
                        }
                        val += Util.GetCellValue(row.GetCell(j));
                    }
                    var hash_val = val;
                    if (addRowID)
                    {
                        hash_val = hash_val + "." + i;
                    }
                    if (nameHash.Contains(hash_val))
                    {
                        if (checkCellCount < 6)
                        {
                            return SearchStatus.NextCol;
                        }
                        else
                        {
                            // 已经找不到能作为key的了。把id和行号连一块
                            return SearchStatus.Fail;
                        }
                    }

                    nameHash.Add(hash_val);

                    list.Add(new string2int(val, i));
                }

                return SearchStatus.Succ;
            };
            var list1 = new List<string2int>();
            var list2 = new List<string2int>();

            var searchstatus1 = search(books["src"], sheet1, list1);
            switch (searchstatus1)
            {
                case SearchStatus.Fail:
                    return GetIDDiffList(sheet1, sheet2, 1, true, startCheckCell);
                case SearchStatus.NextCol:
                    return GetIDDiffList(sheet1, sheet2, checkCellCount + 1, addRowID, startCheckCell);
                default:
                    break;
            }
            var searchstatus2 = search(books["dst"], sheet2, list2);
            switch (searchstatus2)
            {
                case SearchStatus.Fail:
                    return GetIDDiffList(sheet1, sheet2, 1, true, startCheckCell);
                case SearchStatus.NextCol:
                    return GetIDDiffList(sheet1, sheet2, checkCellCount + 1, addRowID, startCheckCell);
                default:
                    break;
            }
            Comparison<string2int> sortfunc = (string2int a, string2int b) =>
            {
                int cmp = 0;
                if (allNum)
                {
                    cmp = Double.Parse(a.Key).CompareTo(Double.Parse(b.Key));
                }
                else
                {
                    cmp = a.Key.CompareTo(b.Key);
                }

                if (cmp == 0)
                {
                    return a.Value.CompareTo(b.Value);
                }
                return cmp;
            };

            list1.Sort(sortfunc);
            list2.Sort(sortfunc);

            var option = new DiffOption<string2int>();
            option.EqualityComparer = new SheetIDComparer();
            var result = DiffUtil.Diff(list1, list2, option);
            //var optimize = result.ToList();// 
            // id列不应该把delete/add优化成modify
           // var optimize = DiffUtil.OptimizeCaseDeletedFirst(result);
            return result.ToList();
        }

        List<DiffResult<string>> DiffSheetRow(ISheet sheet1, int row1, ISheet sheet2, int row2, SheetDiffStatus status, out int maxLineCount) {
            var list1 = new List<string>();
            var list2 = new List<string>();

            maxLineCount = 0;
            if (sheet1.GetRow(row1)!=null) {
                var row = sheet1.GetRow(row1);
                var columnCount = books["src"].SheetValideColumn[sheet1.SheetName];
                var columnstart = books["src"].SheetStartPoint[sheet1.SheetName].Item2;
                for (int i = 0; i < columnCount; ++i) {
                    var value = Util.GetCellValue(row.GetCell(i+columnstart));
                    maxLineCount = Math.Max(maxLineCount, value.Count((c) => { return c == '\n'; }) + 1);

                    list1.Add(value);
                }
            }

            if (sheet2.GetRow(row2) != null) {
                var row = sheet2.GetRow(row2);
                var columnCount = books["dst"].SheetValideColumn[sheet2.SheetName];
                var columnstart = books["dst"].SheetStartPoint[sheet2.SheetName].Item2;
                for (int i = 0; i < columnCount; ++i) {
                    var value = Util.GetCellValue(row.GetCell(i+columnstart));
                    maxLineCount = Math.Max(maxLineCount, value.Count((c) => { return c == '\n'; }) + 1);
                    list2.Add(value);
                }
            }
            var diff = DiffUtil.Diff(list1, list2);
            //var optimized = diff.ToList();// DiffUtil.OptimizeCaseDeletedFirst(diff);
            var optimized = DiffUtil.OptimizeCaseDeletedFirst(diff);
            optimized = DiffUtil.OptimizeCaseInsertedFirst(optimized);
            var tlist = optimized.ToList();
            optimized = DiffUtil.OptimizeShift(tlist, false);
            optimized = DiffUtil.OptimizeShift(optimized, true);

            return optimized.ToList();
        }

        void OnSheetChanged() {
            List<SheetSortKeyCombo> keys = new List<SheetSortKeyCombo>();

            var wrap = books["src"];
            var sheet = wrap.GetCurSheet();
            var src_sheet = wrap.sheetname;
            if (!sheetsDiff.ContainsKey(src_sheet)) return;

            var sheetdata = sheetsDiff[src_sheet];

            var startpoint = wrap.SheetStartPoint[sheet.SheetName];
            var RowStart = startpoint.Item1;
            var columnStart = startpoint.Item2;
            var columnCount = wrap.SheetValideColumn[sheet.SheetName];

            var list = new List<string>();
            if (ProcessHeader.IsChecked == true) {
                int namekey = config.KeyLineID - 1+ RowStart;
                if (sheet.GetRow(namekey) != null) {
                    var row = sheet.GetRow(namekey);
                    for (int i = columnStart; i < columnCount; ++i) {
                        list.Add(Util.GetCellValue(row.GetCell(i)));
                    }
                }
            }
            else {
                for (int i = columnStart; i < columnCount; ++i) {
                    list.Add((i+1).ToString());
                }
            }

            for (var idx = 0; idx < list.Count; ++idx) {
                keys.Add(new SheetSortKeyCombo() { ColumnName = list[idx], ID = idx });
            }
            SortKeyCombo.ItemsSource = keys;
            SortKeyCombo.SelectedIndex = config.DefaultKeyID;
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
                } 

                DstDataGrid.RefreshData();
                OnSheetChanged();
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
                }

                SrcDataGrid.RefreshData();
                OnSheetChanged();
            }
        }

        private void SVNResivionionList_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            var selection = e.AddedItems[0] as SvnRevisionCombo;
            SVNRevisionCombo.Width = Math.Min(selection.Revision.Length*10, 440);

            Entrance.Diff(selection.ID - 1, selection.ID);
        }

        public void OnGridScrollChanged(string tag, ScrollChangedEventArgs e) {
            ScrollViewer view = null;
            if (tag == "src") {
                view = Util.GetVisualChild<ScrollViewer>(DstDataGrid);
            } else if (tag == "dst") {
                view = Util.GetVisualChild<ScrollViewer>(SrcDataGrid);
            }
            if (e.VerticalChange != 0)
                view.ScrollToVerticalOffset(e.VerticalOffset);
            else if (e.HorizontalChange != 0)
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

        private void DoDiff_Click(object sender, RoutedEventArgs e) {
            Refresh();
        }

        private void SimpleHeader_Checked(object sender, RoutedEventArgs e) {
            Refresh();
        }

        private void SVNVersionBtn_Click(object sender, RoutedEventArgs e) {
            
        }

        private void SortKeyCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {

            var src_sheet = books["src"].sheetname;
            var sheetdata = sheetsDiff[src_sheet];

            if (e.AddedItems.Count > 0) {
                if (e.AddedItems[0] is SheetSortKeyCombo sortkey && sheetdata.sortKey != sortkey.ID) {
                    sheetdata.sortKey = sortkey.ID;

                    ReDiffCurSheet();
                }
            }
        }
    }


}
