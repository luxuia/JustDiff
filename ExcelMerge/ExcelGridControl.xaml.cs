using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Dynamic;
using NetDiff;
using System.IO;
using System.Windows.Media;

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

        private void Menu_CopyToSide(object sender, RoutedEventArgs e) {
            var selectCells = ExcelGrid.SelectedCells;

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
        public void RefreshData() {
            var tag = Tag as string;
            var wrap = MainWindow.instance.books[tag];
            var wb = wrap.book;
            var sheet = wb.GetSheetAt(wrap.sheet);

            ExcelGrid.Columns.Clear();
            var columns = ExcelGrid.Columns;
            var header = sheet.GetRow(2);
            var headerkey = sheet.GetRow(1);
            if (header == null || headerkey == null) return;

            // header不会空
            var columnCount = header.Cells.Count;
            var headerStr = new string[columnCount];
            for (int i = 0; i < columnCount; ++i) {
                var cell = header.GetCell(i);
                var cellkey = headerkey.GetCell(i);
                
                var str = Util.GetCellValue(cell);
                var strkey = Util.GetCellValue(cellkey);
  
                if (string.IsNullOrWhiteSpace(str) || string.IsNullOrWhiteSpace(strkey) ) {
                    columnCount = i;
                    break;
                }
                var encodestr = System.Uri.EscapeDataString(strkey);

                var column = new DataGridTextColumn();

                column.Binding = new Binding(encodestr);// { Converter = new ConvertToBackground() };
                column.Header = str;

                Style aStyle = new Style(typeof(TextBlock));
                //var abinding = new MultiBinding() { Converter = new ConvertToBackground() };
                //abinding.Bindings.Add(new Binding(str) { ConverterParameter = "test" });
                //abinding.Bindings.Add(new Binding() { RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor) });
                //abinding.Bindings.Add(new Binding());
                var abinding = new Binding() { Converter = new ConvertToBackground(), ConverterParameter = new ConverterParamter() { columnID = i, coloumnName = str } };

                //abinding.RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor);
                aStyle.Setters.Add(new Setter(TextBlock.BackgroundProperty, abinding));

                column.ElementStyle = aStyle;

                columns.Add(column);

                headerStr[i] = encodestr;
            }

            var datas = new ObservableCollection<ExcelData>();

            int MAX_RANGE_COUNT = 3;

            int changedAnchorCount = 0;
            var rangeData = new List<ExcelData>();
            /*
            if (MainWindow.instance.diffSheetName != null) {
                int sheetDiffidx = MainWindow.instance.diffSheetName.FindIndex(a => tag == "src" ? a.Obj1!=null&& a.Obj1.ID == wrap.sheet : a.Obj2!=null&& a.Obj2.ID == wrap.sheet);

                var status = MainWindow.instance.sheetsDiff[sheetDiffidx];

                for (int j = 0; ; j++) {
                    var row = sheet.GetRow(j);
                    if (row == null || !Util.CheckValideRow(row)) break;

                    var data = new ExcelData();
                    data.idx = row.RowNum;
                    data.tag = Tag as string;

                    var rowid2DiffMap = status.rowID2DiffMap1;
                    if (tag == "dst") {
                        rowid2DiffMap = status.rowID2DiffMap2;
                    }
                    
                    if (j < 3) {
                        data.diffstatus = status.diffHead;
                        changedAnchorCount = 1;
                    }else {

                        if (rowid2DiffMap.ContainsKey(j+1)) {

                        }

                        data.diffstatus = status.diffSheet[rowid2DiffMap[j]];

                        var changed = data.diffstatus.Any((a) => a.Status != DiffStatus.Equal);

                        if (changed) {
                            changedAnchorCount = MAX_RANGE_COUNT;
                        }
                        else {

                        }
                    }
        
                    for (int i = 0; i < columnCount; ++i) {
                        var cell = row.GetCell(i);
                        data.data[headerStr[i]] = Util.GetCellValue(cell);
                    }
                    if (changedAnchorCount > 0) {

                        foreach (var i in rangeData) {
                            datas.Add(i);
                        }
                        rangeData.Clear();

                        datas.Add(data);
                        changedAnchorCount--;
                    }
                    else {
                        if (rangeData.Count > 2) {
                            rangeData.RemoveAt(0);
                        }
                        rangeData.Add(data);
                    }
                }
            }
            */
            // 不把diff结果转换为原来的顺序。因为隐藏相同行后，转换没有意义
            if (MainWindow.instance.diffSheetName != null) {
                int sheetDiffidx = MainWindow.instance.diffSheetName.FindIndex(a => tag == "src" ? a.Obj1 != null && a.Obj1.ID == wrap.sheet : a.Obj2 != null && a.Obj2.ID == wrap.sheet);

                var status = MainWindow.instance.sheetsDiff[sheetDiffidx];

                // 头
                for (int j = 0; j<3; j++) {
                    var row = sheet.GetRow(j);
                    if (row == null || !Util.CheckValideRow(row)) break;

                    var data = new ExcelData();
                    data.rowId = row.RowNum;
                    data.tag = Tag as string;
                    data.diffIdx = j;

                    var rowid2DiffMap = status.rowID2DiffMap1;
                    if (tag == "dst") {
                        rowid2DiffMap = status.rowID2DiffMap2;
                    }
                    data.diffstatus = status.diffHead;

                    for (int i = 0; i < columnCount; ++i) {
                        var cell = row.GetCell(i);
                        data.data[headerStr[i]] = new CellData() { value = Util.GetCellValue(cell), cell = cell };
                    }

                    datas.Add(data);
                }

                Dictionary<int, Dictionary<int, CellEditMode>> edited;
                if (selfTag == "src") {
                    edited = status.RowEdited1;
                } else {
                    edited = status.RowEdited2;
                }
                for (int j = 0; j< status.diffSheet.Count; j++) {
                    int rowid = status.Diff2RowID1[j];
                    if (tag == "dst") {
                        rowid = status.Diff2RowID2[j];
                    }

                    // 修改过，或者是
                    if (edited[rowid].Count > 0 || status.diffSheet[j].Any(a => a.Status != DiffStatus.Equal)) {
       

                        var row = sheet.GetRow(rowid);

                        var data = new ExcelData();
                        data.rowId = rowid;
                        data.tag = Tag as string;
                        data.diffstatus = status.diffSheet[j];
                        data.diffIdx = j;
                        data.CellEdited = edited[rowid];

                        for (int i = 0; i < columnCount; ++i) {
                            var cell = row.GetCell(i);
                            data.data[headerStr[i]] = new CellData() { value = Util.GetCellValue(cell), cell = cell};
                        }

                        datas.Add(data);
                    }
                }
            }

            ExcelGrid.DataContext = datas;

            CtxMenu.Items.Clear();
            var item = new MenuItem();
            item.Header = "复制到" + (isSrc ? "右侧" : "左侧");
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
                    HandleFileOpen(files[0], FileOpenType.Drag, Tag as string);

                    if (files.Length > 1) {
                        HandleFileOpen(files[1], FileOpenType.Drag, otherTag);
                    }
                }
            }
        }

        private void ExcelGrid_LoadingRow(object sender, DataGridRowEventArgs e) {
            var row = e.Row;
            var item = row.Item;
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
                    MainWindow.instance.OnSelectGridRow(Tag as string, row.rowId);
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

                if (rowdata.diffstatus != null && rowdata.diffstatus.Count > coloumnid) {
                    DiffStatus status = rowdata.diffstatus[coloumnid].Status;

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
                            if (rowdata.CellEdited != null && rowdata.CellEdited.ContainsKey(coloumnid) && rowdata.CellEdited[coloumnid] == CellEditMode.Self) {
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
