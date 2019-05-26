using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using NPOI.SS.UserModel;
using NetDiff;
using System.IO;

namespace ExcelMerge {
    /// <summary>
    /// DirectoryGridControl.xaml 的交互逻辑
    /// </summary>
    public partial class DirectoryGridControl : UserControl {

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



        public DirectoryGridControl() {
            InitializeComponent();

            var data = new ObservableCollection<ExcelData>();

            FileGrid.DataContext = data;
        }



        public void HandleDirOpen(string file, FileOpenType type, string tag) {
            DirectoryDifferWindow.instance.OnFileLoaded(file, tag, type);


        }

        private void FileGrid_Drop(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];

                if (files != null && files.Any()) {
                    HandleDirOpen(files[0], FileOpenType.Drag, Tag as string);

                    if (files.Length > 1) {
                        HandleDirOpen(files[1], FileOpenType.Drag, otherTag);
                    }
                }

            }
        }


        private void FileGrid_ScrollChanged(object sender, ScrollChangedEventArgs e) {
            var tag = sender;

            if (DirectoryDifferWindow.instance != null)
                DirectoryDifferWindow.instance.OnGridScrollChanged(Tag as string, e);
        }

        private void FileGrid_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (e.AddedItems.Count > 0) {
                // chang selected row
                var row = e.AddedItems[0] as ExcelData;
                if (row != null) {
                    // 新行 NewRowItem 类
                    DirectoryDifferWindow.instance.OnSelectGridRow(Tag as string, row.rowId);
                }
            }
        }

        internal void refreshData() {
            var tag = Tag as string;

            var issrc = isSrc;

            FileGrid.Columns.Clear();

            var datas = new ObservableCollection<ExcelData>();

            if (DirectoryDifferWindow.instance.results != null) {
                var columns = FileGrid.Columns;

                var headstr = "filename";
                var column = new DataGridTextColumn();
                column.Binding = new Binding(headstr);
                column.Header = headstr;

                Style aStyle = new Style(typeof(TextBlock));
                // 传下去的参数，当渲染格子的时候，只知道行id，需要通过这里传参数知道列id
                var abinding = new Binding() { Converter = new FileConvertToBackground(), ConverterParameter = new ConverterParamter() { columnID = 1, coloumnName = headstr} };

                aStyle.Setters.Add(new Setter(TextBlock.BackgroundProperty, abinding));

                column.ElementStyle = aStyle;

                columns.Add(column);



                var results = DirectoryDifferWindow.instance.results;

                for (int j = 0; j < results.Count; j++) {

                    var res = results[j];
                    if (res.Status != DiffStatus.Equal) {

                        var data = new ExcelData();
                        data.tag = Tag as string;
                        data.diffIdx = j;
                        data.diffstatus = results;
                        data.rowId = j;

                        var path = res.Obj1 == null ? res.Obj2 : res.Obj1;
                        var filename = Path.GetFileName(path); 

                        data.data["filename"] = new CellData() { value = filename };

                        datas.Add(data);
                    }
                }
            }

            FileGrid.DataContext = datas;

            CtxMenu.Items.Clear();
            var item = new MenuItem();
            item.Header = "复制到" + (issrc ? "右侧" : "左侧");
            item.Click += Menu_CopyToSide;
            CtxMenu.Items.Add(item);
        }

        private void Menu_CopyToSide(object sender, RoutedEventArgs e) {
            var selectCells = FileGrid.SelectedCells;

            MainWindow.instance.CopyCellsValue(Tag as string, otherTag, selectCells);
        }

        public class ConverterParamter {
            public int columnID;
            public string coloumnName;
        }

    }



    class FileConvertToBackground : IValueConverter {

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
            var param = (DirectoryGridControl.ConverterParamter)parameter;
            if (value is ExcelData) {
                var rowdata = (ExcelData)value;
                var rowid = rowdata.rowId;
                var coloumnid = param.columnID;

                DiffStatus status = rowdata.diffstatus[rowid].Status;

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
                            return DependencyProperty.UnsetValue;
                    }
            }
            return DependencyProperty.UnsetValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
            throw new NotImplementedException();
        }
    }
}
