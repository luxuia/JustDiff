using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using string2int = System.Collections.Generic.KeyValuePair<string, int>;
using NPOI.SS.UserModel;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using NetDiff;

namespace ExcelMerge {
    class Util {
        public static T GetVisualChild<T>(DependencyObject parent) where T : Visual {
            T child = null;

            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++) {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null) {
                    child = GetVisualChild<T>(v);
                }
                if (child != null) {
                    return child;
                }
            }
            return null;
        }


        public static string GetCellValue(ICell cell) {
            var str = string.Empty;
            if (cell == null) return str;

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

        public static SolidColorBrush GetColorByDiffStatus(DiffStatus status) {
            switch (status) {
                case DiffStatus.Deleted:
                    return Brushes.Gray;
                case DiffStatus.Inserted:
                    return Brushes.Green;
                case DiffStatus.Modified:
                    return Brushes.Yellow; 
            }
            return null;
        }
    }


    public class WorkBookWrap {
        public IWorkbook book;
        public int sheet;
        public string file;
        public string filename;

        public int reversion;

        public List<ComboBoxItem> sheetCombo;
        public List<SheetNameCombo> sheetName;

        public Dictionary<int, int> ComboIdToItemIdx;
    }

    public enum FileOpenType {
        Drag,
        Menu,
        Prog, //因为diff等形式从程序内部打开的
    }

    public class SheetDiffStatus {
        public int columnCount;
        public List<DiffResult<string>> diffHead;
        public List<DiffResult<string2int>> diffFistColumn;

        public List<List<DiffResult<string>>> diffSheet;

        public Dictionary<int, int> rowID2DiffMap1;
        public Dictionary<int, int> rowID2DiffMap2;

        public bool changed;
    }


    public class SheetNameCombo {
        public string Name { get; set; }
        public int ID { get; set; }

        public override string ToString() {
            return Name;
        }
    }

    public class SvnRevisionCombo {
        public string Revision { get; set; }
        public int ID { get; set; }
    }


    public class SheetIDComparer : IEqualityComparer<string2int> {
        public bool Equals(string2int a, string2int b) {
            return a.Key == b.Key;
        }

        public int GetHashCode(string2int a) {
            return a.Key.GetHashCode();
        }
    }

    public class SheetNameComboComparer : IEqualityComparer<SheetNameCombo> {
        public bool Equals(SheetNameCombo a, SheetNameCombo b) {
            return a.Name == b.Name;
        }

        public int GetHashCode(SheetNameCombo a) {
            return a.Name.GetHashCode();
        }
    }

}
