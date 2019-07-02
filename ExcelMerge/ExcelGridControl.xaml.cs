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
        }

        private void ExcelGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e) {
            var selectCells = ExcelGrid.SelectedCells;
            if (e.EditAction == DataGridEditAction.Commit) {
                var data = e.EditingElement.DataContext as ExcelData;
                var el = e.EditingElement as TextBox;
                if (data.data.ContainsKey(e.Column.SortMemberPath)) {
                    var celldata = data.data[e.Column.SortMemberPath];

                    //MainWindow.instance.SetCellValue(el.Text, celldata.cell);
                }
            }
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

                if (last_rowid > 0) { 
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

            Clipboard.SetText(ret);

        }

        private void Menu_CopyToSide(object sender, RoutedEventArgs e) {
            var selectCells = ExcelGrid.SelectedCells;
            //var text = Clipboard.GetText(TextDataFormat.Html);
            MainWindow.instance.CopyCellsValue(Tag as string, otherTag, selectCells);
        }

        private void ExcelGridResized(object sender, SizeChangedEventArgs e) {

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

        void AddPrefixRowID() {
            var columns = ExcelGrid.Columns;

            {
                var column = new DataGridTextColumn();

                column.Binding = new Binding("rowid");
                column.Header = "行";

                Style aStyle = new Style(typeof(TextBlock));

                var abinding = new Binding() { Converter = new ConvertToBackground(), ConverterParameter = new ConverterParamter() { columnID = -1, coloumnName = "行" } };

                aStyle.Setters.Add(new Setter(TextBlock.BackgroundProperty, abinding));

                column.ElementStyle = aStyle;

                columns.Add(column);
            }
        }

        public void RefreshData() {
            var tag = Tag as string;
            var wrap = MainWindow.instance.books[tag];
            var wb = wrap.book;
            var sheet = wb.GetSheetAt(wrap.sheet);

            var issrc = isSrc;

            ExcelGrid.Columns.Clear();

            var datas = new ObservableCollection<ExcelData>();

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
                var headerStr = new string[columnCount];

                var needChangeHead = MainWindow.instance.ProcessHeader.IsChecked == true;
                if (needChangeHead) {
                    var header = sheet.GetRow(MainWindow.instance.config.ShowLineID-1);

                    var headerkey = sheet.GetRow(1);
                    if (header == null || headerkey == null) return;

                    // header不会空
                    //                     columnCount = header.Cells.Count;
                    //                     headerStr = new string[columnCount];
                    //AddPrefixRowID();

                    for (int i = 0; i < columnCount; ++i) {
                        var cell = header.GetCell(i);
                        var cellkey = headerkey.GetCell(i);

                        var str = Util.GetCellValue(cell);
                        var strkey = Util.GetCellValue(cellkey);

                        if (string.IsNullOrWhiteSpace(str) || string.IsNullOrWhiteSpace(strkey)) {
                            columnCount = i;
                            break;
                        }
                        // 第二行+第三行，合起来作为key
                        var encodestr = System.Uri.EscapeDataString(strkey) + "_" + i;// + System.Uri.EscapeDataString(str);

                        var column = new DataGridTextColumn();

                        column.Binding = new Binding(encodestr);
                        column.Header = str;

                        Style aStyle = new Style(typeof(TextBlock));

                        var abinding = new Binding() { Converter = new ConvertToBackground(), ConverterParameter = new ConverterParamter() { columnID = i, coloumnName = str } };

                        aStyle.Setters.Add(new Setter(TextBlock.BackgroundProperty, abinding));

                        column.ElementStyle = aStyle;

                        columns.Add(column);

                        headerStr[i] = encodestr;
                    }
                }
                else {

                    //AddPrefixRowID();

                    for (int i = 0; i < columnCount; ++i) {
                        var str = (i + 1).ToString();
                        // 新建一列
                        var column = new DataGridTextColumn();
                        column.Binding = new Binding(str);
                        column.Header = str;

                        Style aStyle = new Style(typeof(TextBlock));
                        // 传下去的参数，当渲染格子的时候，只知道行id，需要通过这里传参数知道列id
                        var abinding = new Binding() { Converter = new ConvertToBackground(), ConverterParameter = new ConverterParamter() { columnID = i, coloumnName = str } };

                        aStyle.Setters.Add(new Setter(TextBlock.BackgroundProperty, abinding));

                        column.ElementStyle = aStyle;

                        columns.Add(column);

                        headerStr[i] = str;
                    }
                }

                if (needChangeHead) {
                    // 头
                    for (int j = 0; j < MainWindow.instance.DiffStartIdx(); j++) {
                        var row = sheet.GetRow(j);
                        if (row == null || !Util.CheckValideRow(row)) break;

                        var data = new ExcelData();
                        data.rowId = row.RowNum;
                        data.tag = Tag as string;
                        data.diffIdx = j;

                        data.column2diff = issrc ? status.column2diff1[0] : status.column2diff2[0];
                        data.diffstatus = status.diffHead;

                        for (int i = 0; i < columnCount; ++i) {
                            var cell = row.GetCell(i);
                            data.data[headerStr[i]] = new CellData() { value = Util.GetCellValue(cell), cell = cell };
                        }

                        datas.Add(data);
                    }
                }

                Dictionary<int, Dictionary<int, CellEditMode>> edited = issrc ? status.RowEdited1 : status.RowEdited2;

                for (int j = 0; j< status.diffSheet.Count; j++) {
                    int rowid = issrc ? status.Diff2RowID1[j] : status.Diff2RowID2[j];

                    // 修改过，或者是
                    if (edited[rowid].Count > 0 || status.diffSheet[j].Any(a => a.Status != DiffStatus.Equal)) {
       

                        var row = sheet.GetRow(rowid);

                        var data = new ExcelData();
                        data.rowId = rowid;
                        data.tag = Tag as string;
                        data.diffstatus = status.diffSheet[j];
                        data.diffIdx = j;
                        data.CellEdited = edited[rowid];
                        data.column2diff = issrc ? status.column2diff1[rowid] : status.column2diff2[rowid];

                        data.data["rowid"] = new CellData() { value = (rowid+1).ToString() };
                        for (int i = 0; i < columnCount; ++i) {
                            var cell = row != null ? row.GetCell(i):null;
                            data.data[headerStr[i]] = new CellData() { value = Util.GetCellValue(cell), cell = cell};
                        }

                        datas.Add(data);
                    }
                }
            }

            ExcelGrid.DataContext = datas;

            CtxMenu.Items.Clear();
            var item = new MenuItem();
            item.Header = "行复制到" + (issrc ? "右侧" : "左侧");
            item.Click += Menu_CopyToSide;
            CtxMenu.Items.Add(item);
        }

        public void HandleFileOpen(string file, FileOpenType type, string tag) {
            var wb = Util.GetWorkBook(file);

            if (wb != null) {
                var window = MainWindow.instance;

                window.books.Clear();
                window.OnFileLoaded(file, tag, type);
                RefreshData();
            }
        }

        private void ExcelGrid_Drop(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                
                if (files != null && files.Any()) {
                    var file = files[0];
                    if (Directory.Exists(file)) {
                        MainWindow.instance.ShowDirectoryWindow(files, Tag as string);
                    }
                    else {
                        HandleFileOpen(files[0], FileOpenType.Drag, Tag as string);

                        if (files.Length > 1) {
                            HandleFileOpen(files[1], FileOpenType.Drag, otherTag);
                        }
                    }
                }
            }
        }

        private void ExcelGrid_LoadingRow(object sender, DataGridRowEventArgs e) {
            var row = e.Row;
            var index = row.GetIndex();
            var item = row.Item as ExcelData;

            if (item != null) { 
                row.Header = (item.rowId+1).ToString();
            }
            //row.Header = ite
        }

        private void ExcelGrid_ScrollChanged(object sender, ScrollChangedEventArgs e) {
            var tag = sender;

            if (MainWindow.instance != null)
                MainWindow.instance.OnGridScrollChanged(Tag as string, e);
        }

        private void ExcelGrid_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (e.AddedItems.Count > 0) {
                // chang selected row
                var row = e.AddedItems[0] as ExcelData;
                if (row != null) {
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

    class ConvertToBackground : IValueConverter {

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
            var param = (ExcelGridControl.ConverterParamter)parameter;
            if (value is ExcelData) {
                var rowdata = (ExcelData)value;
                var rowid = rowdata.rowId;
                var coloumnid = param.columnID;

                if (rowdata.diffstatus != null && rowdata.diffstatus.Count > coloumnid && coloumnid >= 0) {
                    var diffid = rowdata.column2diff[coloumnid];

                    DiffStatus status = rowdata.diffstatus[diffid].Status;

                    switch (status) {
                        case DiffStatus.Modified:
                            return Brushes.Yellow;
                        case DiffStatus.Deleted:
                            // 列增删的时候不好处理，不显示影响的格子
                            if (rowdata.tag == "src")
                                return Brushes.Gray;
                            break;
                        case DiffStatus.Inserted:
                            // 列增删的时候不好处理，不显示影响的格子
                            if (rowdata.tag == "dst")
                                return Brushes.LightGreen;
                            break;
                        default:
                            if (rowdata.CellEdited != null && rowdata.CellEdited.ContainsKey(diffid) && rowdata.CellEdited[diffid] == CellEditMode.Self) {
                                // 单元格修改
                                return new SolidColorBrush(Color.FromRgb(160,238,225));
                            }
                            return DependencyProperty.UnsetValue;
                    }
                }
            } 
            return DependencyProperty.UnsetValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
            throw new NotImplementedException();
        }
    }
}
