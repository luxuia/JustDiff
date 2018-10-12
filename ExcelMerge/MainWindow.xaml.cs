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

namespace ExcelMerge {


    public class WorkBookWrap {
        public IWorkbook book;
        public int sheet;
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {

        public static MainWindow instance;

        public Dictionary<string, WorkBookWrap> books = new Dictionary<string, WorkBookWrap>();

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
        public void OnFileLoaded(string file, string tag) {
            var wb = WorkbookFactory.Create(file);

            books[tag] = new WorkBookWrap() { book = wb, sheet = 0 };

            if (tag == "src") {
                SrcFilePath.Content = file;
                List<SheetNameCombo> list = new List<SheetNameCombo>();
                for (int i = 0; i < wb.NumberOfSheets; ++i) {
                    list.Add( new SheetNameCombo() { Name = wb.GetSheetName(i), ID = i });

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

        private void DstFileSheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            var selection = e.AddedItems[0] as SheetNameCombo;
            books["dst"].sheet = selection.ID;
            DstDataGrid.RefreshData();
        }

        private void SrcFileSheetsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            var selection = e.AddedItems[0] as SheetNameCombo;
            books["src"].sheet = selection.ID;
            SrcDataGrid.RefreshData();
        }
    }


}
