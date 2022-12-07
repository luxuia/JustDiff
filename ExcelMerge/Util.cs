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
using System.IO;
using System.Windows.Documents;

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

        public static bool CheckIsXLS(string path) {
            path = path.ToLower();
            return path.EndsWith(".xls") || path.EndsWith(".xlsx");
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
                    if (cell.CachedFormulaResultType == CellType.Numeric) {
                        str = cell.NumericCellValue.ToString();
                    }
                    else if (cell.CachedFormulaResultType == CellType.String) {
                        str = cell.StringCellValue.ToString();
                    }
                    else {
                        str = "(公式)";
                        //str = cell.CellFormula;
                    }
                    break;
                case CellType.Numeric:
                    str = cell.NumericCellValue.ToString();
                    break;
                case CellType.String:
                    str = cell.RichStringCellValue.ToString();
                    break;
            }
            return str;
            //return '[' + str + ']';
            //return str.Replace('(', '-').Replace(')', '-').Replace("/", "-");
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

        public static IWorkbook GetWorkBook(string file) {
            IWorkbook book = null;
            try
            {
                using (var s = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    book = WorkbookFactory.Create(s);
                }
            }
            finally
            {
                if (book == null)
                {
                    using (var s = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        book = WorkbookFactory.Create(s);
                    }
                }
            }
            return book;
        }

        public static bool CheckValideRow(IRow row) {
            var str = string.Empty;
            for (int i = 0; i < 5; i++) {
                str += GetCellValue(row.GetCell(i));
            }
            return !string.IsNullOrWhiteSpace(str);
        }

        public static bool CopyCell(ICell oldCell, ICell newCell) {
            if (oldCell == null || newCell == null) return false;
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
                    //newCell.SetCellFormula(oldCell.CellFormula);
                    //break;
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
    public class Config
    {
        public List<string> NoHeadPaths = new List<string>();

        public int HeadCount = 3;

        public int ShowLineID = 3;

        public int KeyLineID = 2;

        public bool NoHead = false;

        public int DefaultKeyID = 0;

        public int EmptyLine = 0;
    }

    public class WorkBookWrap {
        public IWorkbook book;

        private int _sheet;
        // 当前的显示的sheet
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

        public List<ComboBoxItem> sheetCombo = new List<ComboBoxItem>();
        public List<SheetNameCombo> sheetNameCombos = new List<SheetNameCombo>();

        public Dictionary<int, int> ItemID2ComboIdx = new Dictionary<int, int>();

        public Dictionary<string, int> SheetValideRow = new Dictionary<string, int>();
        public Dictionary<string, int> SheetValideColumn = new Dictionary<string, int>();

        public Dictionary<string, Tuple<int, int>> SheetStartPoint = new Dictionary<string, Tuple<int, int>>();

        public Dictionary<string, List<string>> SheetHeaders = new Dictionary<string, List<string>>();

        public Dictionary<string, List<string>> SheetIDs = new Dictionary<string, List<string>>();

        public WorkBookWrap(string file, Config cfg)
        {
            book = Util.GetWorkBook(file);
            this.file = file;
            filename = System.IO.Path.GetFileName(file);

            for (int i = 0; i < book.NumberOfSheets; ++i)
            {
                sheetNameCombos.Add(new SheetNameCombo() { Name = book.GetSheetName(i), ID = i });
            }
            sheetNameCombos.Sort((a, b) => { return a.Name.CompareTo(b.Name); });

            sheetNameCombos.ForEach((a) => { var item = new ComboBoxItem(); item.Content = a; sheetCombo.Add(item); });

            for (int i = 0; i < sheetNameCombos.Count; ++i)
            {
                ItemID2ComboIdx[sheetNameCombos[i].ID] = i;
            }

            CalValideRow(cfg);

            CalValideHeader(cfg);

            CalValideIds(cfg);
        }

        void CalValideRow(Config cfg)
        {

            for (int sheeti = 0; sheeti < book.NumberOfSheets; ++sheeti)
            {
                var sheet = book.GetSheetAt(sheeti);
                int startrow = 0, startcol = 0;
                for (int i = 0; i < 20; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row != null && Util.CheckValideRow(row))
                    {
                        startrow = i;
                        break;
                    };
                }
                var firstrow = sheet.GetRow(startrow);
                for (int i = 0; i < 5; i++)
                {
                    var cell = firstrow != null ? firstrow.GetCell(i) : null;
                    if (cell != null && !string.IsNullOrEmpty(Util.GetCellValue(cell)))
                    {
                        startcol = i;
                        break;
                    };
                }
                SheetStartPoint[sheet.SheetName] = new Tuple<int, int>(startrow, startcol);

                for (int i = startrow; ; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null || !Util.CheckValideRow(row))
                    {
                        if (cfg.EmptyLine-- > 0)
                        {
                            continue;
                        }
                        else
                        {
                            SheetValideRow[sheet.SheetName] = i;
                            break;
                        }
                    };
                }
            }
        }

        void CalValideHeader(Config cfg)
        {
            for (int sheeti = 0; sheeti < book.NumberOfSheets; ++sheeti)
            {
                var sheet = book.GetSheetAt(sheeti);
                List<string> header = new List<string>();

                var startpoint = SheetStartPoint[sheet.SheetName];
                var startrow = startpoint.Item1;
                var startcol = startpoint.Item2;
                var headrow = startrow + cfg.KeyLineID;

                var row = sheet.GetRow(headrow);
                if (row != null)
                {
                    for (int i = startcol; i < row.Cells.Count; ++i)
                    {
                        var val = Util.GetCellValue(row.GetCell(i));
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            header.Add(val);
                        } else
                        {
                            break;
                        }
                    }

                    SheetHeaders[sheet.SheetName] = header;
                }
            }
        }

        List<string> GetIdList(ISheet sheet, int startrow, int totalrow, int startcol, int colcount)
        {
            HashSet<string> idmap = new HashSet<string>();
            List<string> header = new List<string>();

            for (int i = startrow; i < totalrow; i++)
            {
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    string value = "";
                    for (var j = startcol; j < startcol + colcount; ++j)
                    {
                        value += "|" + Util.GetCellValue(row.GetCell(j));
                    }
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        header.Add(value);
                        if (idmap.Contains(value) && colcount <5)
                        {
                            return GetIdList(sheet, startrow, totalrow, startcol, colcount+1);
                        }
                        idmap.Add(value);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            return header;
        }

        void CalValideIds(Config cfg)
        {
            for (int sheeti = 0; sheeti < book.NumberOfSheets; ++sheeti)
            {
                var sheet = book.GetSheetAt(sheeti);

                var startpoint = SheetStartPoint[sheet.SheetName];
                var startrow = startpoint.Item1;
                var startcol = startpoint.Item2;
                var totalrow = SheetValideRow[sheet.SheetName];
                var idcol = startcol + cfg.DefaultKeyID;

                SheetIDs[sheet.SheetName] = GetIdList(sheet, startrow, totalrow, startcol, 1);
            }
        }

        public ISheet GetCurSheet() {
            return book.GetSheetAt(sheet);
        }

        public string GetSheetNameByComboID(int index) {
            return index < sheetNameCombos.Count ? sheetNameCombos[index].Name : null;
        }

        public int GetComboIDBySheetName(string name) {
            return sheetNameCombos.FindIndex((a) => a.Name == name);
        }
    }

    public enum FileOpenType {
        Drag,
        Prog, //因为diff等形式从程序内部打开的
    }

    public enum Mode {
        Diff,
        Merge,
    }

    public enum CellEditMode {
        Self, // 自己修改
        OtherSide, // 另一边的格子修改
    }

    public class SheetRowDiff {
        public List<DiffResult<string>> diffcells;
        public List<List<DiffResult<char>>> diffcell_details;
        public bool changed = false;
    }

    public class SheetDiffStatus {
        public int columnCount1;
        public int columnCount2;
        public SheetRowDiff diffHead;

        public int sortKey;

        public List<DiffResult<string2int>> diffFistColumn;

        public List<SheetRowDiff> diffSheet;

        public Dictionary<int, int> rowID2DiffMap1;
        public Dictionary<int, int> rowID2DiffMap2;

        public Dictionary<int, int> Diff2RowID1;
        public Dictionary<int, int> Diff2RowID2;

        public Dictionary<int, int> DiffMaxLineCount;

        public Dictionary<int, Dictionary<int, CellEditMode>> RowEdited1;
        public Dictionary<int, Dictionary<int, CellEditMode>> RowEdited2;

        public HashSet<int> ignoreRow1;
        public HashSet<int> ignoreRow2;

        // 
        public Dictionary<int, int[]> column2diff1;
        public Dictionary<int, int[]> column2diff2;

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
        public long ID { get; set; }
    }

    public class SheetSortKeyCombo {
        public string ColumnName { get; set; }
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

    public class CellData {
        public string value;
        public ICell cell;
    }

    public class ExcelData : DynamicObject {
        public Dictionary<string, CellData> data = new Dictionary<string, CellData>();
        public int rowId;
        public string tag;

        public int maxLineCount=1;


        public SheetRowDiff diffstatus;
        public Dictionary<int, int> RowID2DiffMap;
        public Dictionary<int, CellEditMode> CellEdited;

        public int[] column2diff;
  
        public override bool TryGetMember(GetMemberBinder binder, out object result) {
            CellData ret;
            if (data.TryGetValue(binder.Name, out ret)) {
                //result = "<Bold> " + ret.value + "<Bold/>";
                result = ret.value;
                //var block = new TextBlock();
                //result = new Bold(new Run(ret.value));
               // result = block;
                return true;
            }
            result = "";
            return false;
        }

        public override bool TrySetMember(SetMemberBinder binder, object value) {
            var ret = data[binder.Name];
            ret.value = value as string;
            ret.cell.SetCellValue(ret.value);

            MainWindow.instance.OnCellEdited(tag, rowId, ret.cell.ColumnIndex, CellEditMode.Self);
            MainWindow.instance.RefreshCurSheet();
  
            return true;
        }
    }

}
