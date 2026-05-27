using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using NetDiff;
using System.IO;
using System.Windows.Input;
using System.Net;
using System.Windows.Shapes;
// NPOI 已替换为基于 MiniExcel 的自定义只读抽象（参见 ExcelModel.cs）。

namespace ExcelMerge {
    /// <summary>
    /// Interaction logic for ExcelGridControl.xaml
    /// </summary>
    public partial class ExcelGridControl : UserControl {


        public bool isSrc {
            get { return Tag as string == "src"; }
        }

        public string selfTag {
            get {
                return Tag as string;
            }
        }

        public string otherTag {
            get {
                return isSrc ? "dst" : "src";
            }
        }

        public ExcelGridControl() {
            InitializeComponent();

            var data = new ObservableCollection<ExcelData>();

            ExcelGrid.DataContext = data;

            ExcelGrid.CellEditEnding += ExcelGrid_CellEditEnding;

            ExcelGrid.CommandBindings.Add(new CommandBinding(ApplicationCommands.Copy, Menu_Save));
            ExcelGrid.InputBindings.Add(new InputBinding(ApplicationCommands.Copy, ApplicationCommands.Copy.InputGestures[0]));

            DiffMarkerCanvas.SizeChanged += (s, e) => DrawMarkers();
        }

        const double RowHeaderWidthPx = 50;

        public bool editing = false;
        private void ExcelGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e) {
            var selectCells = ExcelGrid.SelectedCells;
            editing = false;
            if (e.EditAction == DataGridEditAction.Commit) {
                var data = e.EditingElement.DataContext as ExcelData;
                var el = e.EditingElement as TextBox;
                if (data.data.ContainsKey(e.Column.SortMemberPath)) {
                    var celldata = data.data[e.Column.SortMemberPath];

                    //MainWindow.instance.SetCellValue(el.Text, celldata.cell);
                }
            }
        }
        
        private void ExcelGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e) {
            var tag = sender;
            var ee = e;
            editing = true;
        }

        private void Menu_Save(object sender, ExecutedRoutedEventArgs e) {
            var selectCells = ExcelGrid.SelectedCells;
            var srcSheet = MainWindow.instance.books[selfTag].GetCurSheet();

            var ret = "<TABLE><TR>";
            var last_rowid = -1;
            foreach (var cell in selectCells) {
                var rowdata = cell.Item as ExcelData;
                var column = cell.Column.DisplayIndex;
                var rowid = rowdata.rowId;

                var row = srcSheet.GetRow(rowid);

                if (last_rowid >= 0) { 
                    if (last_rowid != rowid) {
                        ret += "</TR><TR>";
                    } else {
                        ret += "";
                    }
                }

                ret += String.Format("<TD>{0}</TD>", WebUtility.HtmlEncode( Util.GetCellValue(row.GetCell(column))).Replace("\n", "<br style=\"mso-data-placement:same-cell; \" />"));

                last_rowid = rowid;
            }
            ret += "</TR></TABLE>";

            Clipboard.SetDataObject(ret);

        }


        DependencyProperty GetDependencyPropertyByName(Type dependencyObjectType, string dpName) {
            DependencyProperty dp = null;

            var fieldInfo = dependencyObjectType.GetField(dpName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.FlattenHierarchy);
            if (fieldInfo != null) {
                dp = fieldInfo.GetValue(null) as DependencyProperty;
            }

            return dp;
        }

        public class ConverterParamter {
            public int columnID;
            public string coloumnName;
        }

        public void RefreshData() {
            var tag = Tag as string;
            var wrap = MainWindow.instance.books[tag];
            var wb = wrap.book;
            var sheet = wb.GetSheetAt(wrap.sheet);

            var issrc = isSrc;

            ExcelGrid.Columns.Clear();

            //var datas = new ObservableCollection<ExcelData>();
            var datas = new List<ExcelData>();
            var data_maps = new Dictionary<int, ExcelData>();
            MainWindow.instance.excelGridData[tag] = data_maps;

            if (MainWindow.instance.diffSheetName != null) {
                var columns = ExcelGrid.Columns;

                // 不把diff结果转换为原来的顺序。因为隐藏相同行后，转换没有意义

                var sheetname = wrap.sheetname;

                if (!MainWindow.instance.sheetsDiff.ContainsKey(sheetname)) { 
                    ExcelGrid.DataContext = datas;
                    return;
                }
                var status = MainWindow.instance.sheetsDiff[sheetname];

                // 没有比较数据的sheet
                if (status == null) return;

                // header不会空
                var columnCount = wrap.SheetValideColumn[sheet.SheetName];
                var startpoint = wrap.SheetStartPoint[sheet.SheetName];
                var startrow = startpoint.Item1;
                var startcol = startpoint.Item2;

                var headerStr = new string[columnCount];

                var needChangeHead = status.isConfigDiff;
                if (needChangeHead) {
                    var headershow = sheet.GetRow(MainWindow.instance.config.ShowLineID-1 + startrow);
                    var headerkey = sheet.GetRow(MainWindow.instance.config.KeyLineID-1 +startrow);

                    if (headershow == null || headerkey == null) return;

                    int linecount = 0;
                    for (int i = 0; i < columnCount; ++i) {
                        var cellshow = headershow.GetCell(i + startcol);
                        var cellkey = headerkey.GetCell(i + startcol);

                        var strshow = Util.GetCellValue(cellshow);
                        var strkey = Util.GetCellValue(cellkey);
                        if (string.IsNullOrWhiteSpace(strshow)) {
                            strshow = strkey;
                        }

                        if (string.IsNullOrWhiteSpace(strkey)) {
                            columnCount = i;
                            break;
                        }
                        // 第二行+第三行，合起来作为key
                        var encodestr = System.Uri.EscapeDataString(strkey) + "_" + i;// + System.Uri.EscapeDataString(str);
                        linecount = Math.Max(linecount, strshow.Count((c) => { return c == '\n'; })+1);

                        var tc = new DataGridTemplateColumn();
                        tc.Header = strshow;
                        tc.Width = new DataGridLength(Math.Max(100, strshow.Length * 8 + 20));
                        tc.CanUserResize = true;
                        tc.CellTemplateSelector = new CellTemplateSelector(encodestr, i, tag);
                        tc.CellEditingTemplateSelector = new CellTemplateSelector(encodestr, i, tag);

                        columns.Add(tc);

                        headerStr[i] = encodestr;
                    }
                    ExcelGrid.ColumnHeaderHeight = linecount * 25;
                }
                else {
                    //AddPrefixRowID();

                    for (int i = 0; i < columnCount; ++i) {
                        var str = (i + 1).ToString();

                        var tc = new DataGridTemplateColumn();
                        tc.Header = Util.NumberToExcelColumnId(i+1);
                        tc.Width = new DataGridLength(100);
                        tc.CanUserResize = true;
                        tc.CellTemplateSelector = new CellTemplateSelector(str, i, tag);
                        tc.CellEditingTemplateSelector = new CellTemplateSelector(str, i, tag);

                        columns.Add(tc);

                        headerStr[i] = str;
                    }
                    ExcelGrid.ColumnHeaderHeight = 25;
                }

                if (needChangeHead) {
                    // 头
                    for (int j = startrow; j < MainWindow.instance.DiffStartIdx(startrow); j++) {
                        var row = sheet.GetRow(j);
                        if (row == null || !Util.CheckValideRow(row)) break;

                        var data = new ExcelData();
                        data.rowId = row.RowNum;
                        data.tag = Tag as string;

                        data.column2diff = issrc ? status.column2diff1[0] : status.column2diff2[0];
                        data.diffstatus = status.diffHead;

                        for (int i = 0; i < columnCount; ++i) {
                            var cell = row.GetCell(i+startcol);
                            var value = Util.GetCellValue(cell);
                            data.data[headerStr[i]] = new CellData() { value = value,  cell = cell };
                        }
                        if (!status.DiffMaxLineCount.TryGetValue(j, out data.maxLineCount)) {
                            data.maxLineCount = 1;
                        }
   
                        datas.Add(data);
                        data_maps[data.rowId] = data;
                    }
                }


                for (int j = 0; j< status.diffSheet.Count; j++) {
                    int rowid = issrc ? status.Diff2RowID1[j] : status.Diff2RowID2[j];

                    // 修改过，或者是
                    if ( status.diffSheet[j].changed) {
       
                        var row = sheet.GetRow(rowid);

                        var data = new ExcelData();
                        data.rowId = rowid;
                        data.tag = Tag as string;
                        data.diffstatus = status.diffSheet[j];
                        data.column2diff = issrc ? status.column2diff1[rowid] : status.column2diff2[rowid];

                        data.data["rowid"] = new CellData() { value = (rowid+1).ToString() };
                        for (int i = 0; i < columnCount; ++i) {
                            var cell = row != null ? row.GetCell(i+startcol):null;
                            var value = Util.GetCellValue(cell);
                            data.data[headerStr[i]] = new CellData() { value = value, cell = cell};
                        }
                        if (!status.DiffMaxLineCount.TryGetValue(j, out data.maxLineCount)) {
                            data.maxLineCount = 1;
                        }

                        datas.Add(data);
                        data_maps[data.rowId] = data;
                    }
                }
            }
            ExcelGrid.ItemsSource = datas;
            UpdateDiffMarkers(datas);

            CtxMenu.Items.Clear();
        }

        List<int> _diffColumnPositions;
        int _totalColumns;

        void UpdateDiffMarkers(List<ExcelData> datas) {
            _diffColumnPositions = new List<int>();
            _totalColumns = ExcelGrid.Columns.Count;

            var tag = Tag as string;
            if (!MainWindow.instance.books.ContainsKey(tag)) {
                DrawMarkers();
                return;
            }
            var wrap = MainWindow.instance.books[tag];
            var sheetname = wrap.sheetname;

            if (!MainWindow.instance.sheetsDiff.ContainsKey(sheetname)) {
                DrawMarkers();
                return;
            }
            var status = MainWindow.instance.sheetsDiff[sheetname];
            if (status == null || _totalColumns <= 0) {
                DrawMarkers();
                return;
            }

            var changedColumns = new HashSet<int>();
            for (int j = 0; j < status.diffSheet.Count; j++) {
                var rowDiff = status.diffSheet[j];
                if (!rowDiff.changed || rowDiff.diffcells == null) continue;

                bool allInserted = true, allDeleted = true;
                foreach (var cell in rowDiff.diffcells) {
                    if (cell.Status != NetDiff.DiffStatus.Inserted) allInserted = false;
                    if (cell.Status != NetDiff.DiffStatus.Deleted) allDeleted = false;
                    if (!allInserted && !allDeleted) break;
                }
                if (allInserted || allDeleted) continue;

                int cellCount = Math.Min(rowDiff.diffcells.Count, _totalColumns);
                for (int col = 0; col < cellCount; col++) {
                    if (rowDiff.diffcells[col].Status != NetDiff.DiffStatus.Equal)
                        changedColumns.Add(col);
                }
            }
            _diffColumnPositions = changedColumns.OrderBy(x => x).ToList();
            DrawMarkers();
        }

        bool GetMarkerLayout(out double leftOffset, out double w) {
            leftOffset = RowHeaderWidthPx;
            double rightOffset = SystemParameters.VerticalScrollBarWidth;
            w = DiffMarkerCanvas.ActualWidth - leftOffset - rightOffset;
            return w > 0;
        }

        void DrawMarkers() {
            DiffMarkerCanvas.Children.Clear();
            if (_totalColumns <= 0 || _diffColumnPositions == null || _diffColumnPositions.Count == 0) return;
            if (!GetMarkerLayout(out double leftOffset, out double w)) return;

            double markerW = Math.Max(2, Math.Min(4, w / _totalColumns));
            foreach (var col in _diffColumnPositions) {
                double x = leftOffset + (double)col / _totalColumns * w;
                var rect = new Rectangle {
                    Width = markerW,
                    Height = 8,
                    Fill = Brushes.Red,
                    Opacity = 0.7
                };
                Canvas.SetLeft(rect, x);
                DiffMarkerCanvas.Children.Add(rect);
            }
        }

        void DiffMarkerCanvas_Click(object sender, MouseButtonEventArgs e) {
            if (_totalColumns <= 0 || ExcelGrid.Columns.Count == 0 || ExcelGrid.Items.Count == 0) return;
            if (!GetMarkerLayout(out double leftOffset, out double w)) return;

            double x = e.GetPosition(DiffMarkerCanvas).X - leftOffset;
            double ratio = Math.Max(0, Math.Min(1, x / w));
            int targetCol = (int)(ratio * _totalColumns);
            targetCol = Math.Max(0, Math.Min(targetCol, ExcelGrid.Columns.Count - 1));

            var anchor = ExcelGrid.SelectedItem ?? ExcelGrid.Items[0];
            ExcelGrid.ScrollIntoView(anchor, ExcelGrid.Columns[targetCol]);
        }

        private void ExcelGrid_Drop(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                Entrance.OnDragFile(files, isSrc);
            }
        }

        private void ExcelGrid_LoadingRow(object sender, DataGridRowEventArgs e) {
            var row = e.Row;
            var index = row.GetIndex();

            if (row.Item is ExcelData item) { 
                row.Header = (item.rowId+1).ToString();
                row.Height = item.maxLineCount * 15+5;
            }
            //row.Header = ite
        }

        private void ExcelGrid_ScrollChanged(object sender, ScrollChangedEventArgs e) {
            var tag = sender;

            if (MainWindow.instance != null && !editing)
                MainWindow.instance.OnGridScrollChanged(Tag as string, e);
        }

        private void ExcelGrid_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (e.AddedItems.Count > 0) {
                if (e.AddedItems[0] is ExcelData row) {
                    // 新行 NewRowItem 类
                    //MainWindow.instance.OnSelectGridRow(Tag as string, row.rowId);
                }
            }
        }

        private void ExcelGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e) {
            if (e.AddedCells.Count > 0) {
                var cells = e.AddedCells[0];
            }
        }

    }
    
}
