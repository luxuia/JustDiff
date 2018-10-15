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

namespace ExcelMerge {
    /// <summary>
    /// Interaction logic for ExcelGridControl.xaml
    /// </summary>
    public partial class ExcelGridControl : UserControl {
        public class ExcelData :DynamicObject {
            public Dictionary<string, string> data = new Dictionary<string, string>();
            public int idx;

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

        DependencyProperty GetDependencyPropertyByName(Type dependencyObjectType, string dpName) {
            DependencyProperty dp = null;

            var fieldInfo = dependencyObjectType.GetField(dpName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.FlattenHierarchy);
            if (fieldInfo != null) {
                dp = fieldInfo.GetValue(null) as DependencyProperty;
            }

            return dp;
        }

        public void RefreshData() {
            var wrap = MainWindow.instance.books[Tag as string];
            var wb = wrap.book;
            var sheet = wb.GetSheetAt(wrap.sheet);

            ExcelGrid.Columns.Clear();
            var columns = ExcelGrid.Columns;
            var header = sheet.GetRow(2);
            var headerStr = new string[header.Cells.Count];
            for (int i = 0; i < header.Cells.Count; ++i) {
                var cell = header.Cells[i];
                var column = new DataGridTextColumn();
                var str = GetCellValue(cell);

                column.Binding = new Binding(str);// { Converter = new ConvertToBackground() };
                column.Header = str;

                Style aStyle = new Style(typeof(TextBlock));
                //var abinding = new MultiBinding() { Converter = new ConvertToBackground() };
                //abinding.Bindings.Add(new Binding(str) { ConverterParameter = "test" });
                //abinding.Bindings.Add(new Binding() { RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor) });
                //abinding.Bindings.Add(new Binding());
                var abinding = new Binding() { Converter = new ConvertToBackground(), ConverterParameter = str };

                //abinding.RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor);
                aStyle.Setters.Add(new Setter(TextBlock.BackgroundProperty, abinding));

                column.ElementStyle = aStyle;

                columns.Add(column);

                headerStr[i] = str;
            }

            var datas = new ObservableCollection<ExcelData>();

            var ER = sheet.GetRowEnumerator();
            while (ER.MoveNext()) {
                var row = (IRow)ER.Current;
                var data = new ExcelData();
                data.idx = row.RowNum;
                for (int i = 0; i < row.Cells.Count; ++i) {
                    var cell = row.Cells[i];
                    data.data[headerStr[i]] = GetCellValue(cell);
                }
                datas.Add(data);
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
    }

    class ConvertToBackground : IValueConverter {

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
            string input = value as string;
            switch (input) {
                case "New":
                    return Brushes.LightGreen;
                case "Changed":
                    return Brushes.Yellow;
                case "Not Found":
                    return Brushes.Tomato;
                default:
                    return DependencyProperty.UnsetValue;
            }

        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
            throw new NotImplementedException();
        }
    }
}
