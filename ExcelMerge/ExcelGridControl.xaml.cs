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

namespace ExcelMerge {
    /// <summary>
    /// Interaction logic for ExcelGridControl.xaml
    /// </summary>
    public partial class ExcelGridControl : UserControl {
        public class ExcelData :DynamicObject {
            public Dictionary<string, string> data = new Dictionary<string, string>();
            public int idx;
            public string tag;

            public List<DiffResult<string>> diffstatus;
            public Dictionary<int, int> RowID2DiffMap;

            public override bool TryGetMember(GetMemberBinder binder, out object result) {
                string ret = null;
                if (data.TryGetValue(binder.Name, out ret)) {
                    result = ret;

                    return true;
                }
                result = ret;
                return false;
            }

            public override bool TrySetMember(SetMemberBinder binder, object value) {
                data[binder.Name] = value.ToString();

                return true;
            }
        }

        public ExcelGridControl() {
            InitializeComponent();

            var data = new ObservableCollection<ExcelData>();

            ExcelGrid.DataContext = data;
        }

        public void RefreshView() {
            ExcelGrid.Items.Refresh();
        }

        private void selectionCommandClick(object sender, RoutedEventArgs e) {

        }

        private void ExcelGridResized(object sender, SizeChangedEventArgs e) {

        }

        private void hscroll_Scroll(object sender, System.Windows.Controls.Primitives.ScrollEventArgs e) {

        }

        private void vscroll_Scroll(object sender, System.Windows.Controls.Primitives.ScrollEventArgs e) {

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
            // header不会空
            var headerStr = new string[header.Cells.Count];
            for (int i = 0; i < header.Cells.Count; ++i) {
                var cell = header.Cells[i];
                var column = new DataGridTextColumn();
                var str = Util.GetCellValue(cell);

                column.Binding = new Binding(str);// { Converter = new ConvertToBackground() };
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

                headerStr[i] = str;
            }

            var datas = new ObservableCollection<ExcelData>();

            if (MainWindow.instance.diffSheetName != null) {
                int sheetDiffidx = MainWindow.instance.diffSheetName.FindIndex(a => tag == "src" ? a.Obj1.ID == wrap.sheet : a.Obj2.ID == wrap.sheet);

                var status = MainWindow.instance.sheetsDiff[sheetDiffidx];

                for (int j = 0; ; j++) {
                    var row = sheet.GetRow(j);
                    if (row == null) break;

                    var data = new ExcelData();
                    data.idx = row.RowNum;
                    data.tag = Tag as string;

                    if (tag == "src") {
                        data.RowID2DiffMap = status.rowID2DiffMap1;
                    }
                    else {
                        data.RowID2DiffMap = status.rowID2DiffMap2;
                    }

                    if (j < 3)
                        data.diffstatus = status.diffHead;
                    else {

                        data.diffstatus = status.diffSheet[data.RowID2DiffMap[j]];
                    }

                    for (int i = 0; i < status.columnCount; ++i) {
                        var cell = row.GetCell(i);
                        data.data[headerStr[i]] = Util.GetCellValue(cell);
                    }
                    datas.Add(data);
                }
            }
            ExcelGrid.DataContext = datas;
        }

        public void HandleFileOpen(string file, FileOpenType type) {
            var wb = WorkbookFactory.Create(file);
            
            if (wb!=null)
                MainWindow.instance.OnFileLoaded(file, Tag as string, type);
        }

        private void ExcelGrid_Drop(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];

                if (files != null && files.Any()) {
                    HandleFileOpen(files[0], FileOpenType.Drag);
                }
            }
        }

        private void ExcelGrid_LoadingRow(object sender, DataGridRowEventArgs e) {
            var row = e.Row;
            var item = row.Item;
        }

        private void ExcelGrid_ScrollChanged(object sender, ScrollChangedEventArgs e) {
            var tag = sender;

            MainWindow.instance.OnGridScrollChanged(Tag as string, e);
        }
    }

    class ConvertToBackground : IValueConverter {

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
            var param = (ExcelGridControl.ConverterParamter)parameter;
            if (value is ExcelGridControl.ExcelData) {
                var rowdata = (ExcelGridControl.ExcelData)value;
                var rowid = rowdata.idx;
                var coloumnid = param.columnID;

                if (rowdata.diffstatus != null) {
                    DiffStatus status = rowdata.diffstatus[coloumnid].Status;

                    switch (status) {
                        case DiffStatus.Modified:
                            return Brushes.Yellow;
                        case DiffStatus.Deleted:
                            return Brushes.Gray;
                        case DiffStatus.Inserted:
                            return Brushes.LightGreen;
                        default:
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
