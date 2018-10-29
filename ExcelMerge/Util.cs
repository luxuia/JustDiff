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
using System.Dynamic;

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

        public static bool CheckValideRow(IRow row) {
            var str = string.Empty;
            for (int i = 0; i < 5; i++) {
                str += GetCellValue(row.Cells[0]);
            }
            return !string.IsNullOrWhiteSpace(str);
        }

        public static bool CopyCell(ICell oldCell, ICell newCell) {
            if (oldCell.CellStyle != null) {
                // apply style from old cell to new cell 
                // 不是一个xls，没法直接拷贝cellstyle
                //newCell.CellStyle = oldCell.CellStyle;
            }

            // If there is a cell comment, copy
            if (oldCell.CellComment != null) {
                newCell.CellComment = oldCell.CellComment;
            }

            // If there is a cell hyperlink, copy
            if (oldCell.Hyperlink != null) {
                newCell.Hyperlink = oldCell.Hyperlink;
            }

            // Set the cell data type
            newCell.SetCellType(oldCell.CellType);

            // Set the cell data value
            switch (oldCell.CellType) {
                case CellType.Blank:
                    newCell.SetCellValue(oldCell.StringCellValue);
                    break;
                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    newCell.SetCellFormula(oldCell.CellFormula);
                    break;
                case CellType.Numeric:
                    newCell.SetCellValue(oldCell.NumericCellValue);
                    break;
                case CellType.String:
                    newCell.SetCellValue(oldCell.RichStringCellValue);
                    break;
            }

            return true;
        }
    }


    public class WorkBookWrap {
        public IWorkbook book;
        private int _sheet;
        public int sheet {
            set {
                _sheet = Math.Max( Math.Min( value, book.NumberOfSheets-1), 0);
                sheetname = book.GetSheetName(_sheet);
            }get { return _sheet; }
        }

        public string sheetname;
        public string file;
        public string filename;

        public int reversion;

        public List<ComboBoxItem> sheetCombo;
        public List<SheetNameCombo> sheetNameCombos;

        public Dictionary<int, int> ItemID2ComboIdx;

        public ISheet GetCurSheet() {
            return book.GetSheetAt(sheet);
        }
    }

    public enum FileOpenType {
        Drag,
        Menu,
        Prog, //因为diff等形式从程序内部打开的
    }

    public enum Mode {
        Diff,
        Merge,
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

    public class ExcelData : DynamicObject {
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

}
