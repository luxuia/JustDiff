using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using MiniExcelLibs;

namespace ExcelMerge
{
    /// <summary>
    /// 单元格类型。保留 NPOI 风格的枚举名，便于外部代码最小改动。
    /// </summary>
    public enum CellType
    {
        Unknown = -1,
        Numeric = 0,
        String = 1,
        Formula = 2,
        Blank = 3,
        Boolean = 4,
        Error = 5,
    }

    public interface ICell
    {
        CellType CellType { get; }
        string StringCellValue { get; }
        double NumericCellValue { get; }
        bool BooleanCellValue { get; }

        /// <summary>
        /// 归一化后的字符串值，已经在构建阶段完成转换以保证读取路径的极致性能。
        /// </summary>
        string DisplayValue { get; }
    }

    public interface IRow
    {
        int RowNum { get; }
        IReadOnlyList<ICell> Cells { get; }
        ICell GetCell(int columnIndex);
    }

    public interface ISheet
    {
        string SheetName { get; }
        int RowCount { get; }
        IRow GetRow(int rowIndex);
    }

    public interface IWorkbook : IDisposable
    {
        int NumberOfSheets { get; }
        string GetSheetName(int index);
        ISheet GetSheetAt(int index);
        ISheet GetSheet(string name);
        int GetSheetIndex(string name);
    }

    // ---------------- MiniExcel 只读实现 ----------------

    /// <summary>
    /// 基于 MiniExcel 的只读 Cell 实现。
    /// 所有字段在构建期一次性填好，读取路径是纯字段访问，没有反射/分支开销。
    /// </summary>
    internal sealed class MiniExcelCell : ICell
    {
        public static readonly MiniExcelCell Blank = new MiniExcelCell(CellType.Blank, string.Empty, 0d, false);

        public CellType CellType { get; }
        public string StringCellValue { get; }
        public double NumericCellValue { get; }
        public bool BooleanCellValue { get; }
        public string DisplayValue { get; }

        private MiniExcelCell(CellType type, string str, double num, bool boolean)
        {
            CellType = type;
            StringCellValue = str ?? string.Empty;
            NumericCellValue = num;
            BooleanCellValue = boolean;
            DisplayValue = type == CellType.Numeric
                ? num.ToString(CultureInfo.InvariantCulture)
                : (type == CellType.Boolean ? boolean.ToString() : StringCellValue);
        }

        public static ICell Create(object raw)
        {
            if (raw == null) return Blank;

            switch (raw)
            {
                case string s:
                    return new MiniExcelCell(CellType.String, s, 0d, false);
                case bool b:
                    return new MiniExcelCell(CellType.Boolean, string.Empty, 0d, b);
                case double d:
                    return new MiniExcelCell(CellType.Numeric, string.Empty, d, false);
                case float f:
                    return new MiniExcelCell(CellType.Numeric, string.Empty, f, false);
                case decimal dec:
                    return new MiniExcelCell(CellType.Numeric, string.Empty, (double)dec, false);
                case long l:
                    return new MiniExcelCell(CellType.Numeric, string.Empty, l, false);
                case int i:
                    return new MiniExcelCell(CellType.Numeric, string.Empty, i, false);
                case short sh:
                    return new MiniExcelCell(CellType.Numeric, string.Empty, sh, false);
                case byte by:
                    return new MiniExcelCell(CellType.Numeric, string.Empty, by, false);
                case DateTime dt:
                    // 与原 NPOI 行为保持一致：日期以字符串展示（NPOI Numeric 日期走 ToString 实际也是 OA 浮点）
                    return new MiniExcelCell(CellType.String, dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), 0d, false);
                default:
                    return new MiniExcelCell(CellType.String, raw.ToString() ?? string.Empty, 0d, false);
            }
        }
    }

    internal sealed class MiniExcelRow : IRow
    {
        private readonly ICell[] _cells;

        public int RowNum { get; }
        public IReadOnlyList<ICell> Cells => _cells;

        public MiniExcelRow(int rowNum, ICell[] cells)
        {
            RowNum = rowNum;
            _cells = cells;
        }

        public ICell GetCell(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= _cells.Length) return null;
            return _cells[columnIndex];
        }
    }

    internal sealed class MiniExcelSheet : ISheet
    {
        private readonly IRow[] _rows;

        public string SheetName { get; }
        public int RowCount => _rows.Length;

        public MiniExcelSheet(string name, IRow[] rows)
        {
            SheetName = name;
            _rows = rows;
        }

        public IRow GetRow(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= _rows.Length) return null;
            return _rows[rowIndex];
        }
    }

    /// <summary>
    /// 基于 MiniExcel 的只读 Workbook。
    /// 打开时一次性把所有 sheet 流式读入内存并立刻释放文件句柄，
    /// 后续全部访问都是数组下标访问，保证读取性能最高。
    /// </summary>
    internal sealed class MiniExcelWorkbook : IWorkbook
    {
        private readonly ISheet[] _sheets;
        private readonly Dictionary<string, int> _name2Index;

        public int NumberOfSheets => _sheets.Length;

        private MiniExcelWorkbook(ISheet[] sheets)
        {
            _sheets = sheets;
            _name2Index = new Dictionary<string, int>(sheets.Length, StringComparer.Ordinal);
            for (int i = 0; i < sheets.Length; i++)
            {
                _name2Index[sheets[i].SheetName] = i;
            }
        }

        public static MiniExcelWorkbook Load(string file)
        {
            // FileShare.Read 允许其它进程继续只读访问，保持与原 NPOI 实现一致的打开语义。
            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read);

            var sheetNames = MiniExcel.GetSheetNames(fs);
            var sheets = new ISheet[sheetNames.Count];

            for (int si = 0; si < sheetNames.Count; si++)
            {
                // 每个 sheet 需要把流指针归零；MiniExcel 的 Query 支持按 sheetName 读取。
                fs.Position = 0;

                var rowsRaw = new List<ICell[]>();
                int maxCols = 0;

                // useHeaderRow:false -> 返回 IDictionary<string, object>，key 为 "A","B","C"...
                foreach (IDictionary<string, object> rawRow in MiniExcel.Query(fs, useHeaderRow: false, sheetName: sheetNames[si]))
                {
                    var cols = rawRow.Count;
                    if (cols > maxCols) maxCols = cols;

                    var cells = new ICell[cols];
                    int ci = 0;
                    foreach (var kv in rawRow)
                    {
                        cells[ci++] = MiniExcelCell.Create(kv.Value);
                    }
                    rowsRaw.Add(cells);
                }

                var rowArr = new IRow[rowsRaw.Count];
                for (int ri = 0; ri < rowsRaw.Count; ri++)
                {
                    var src = rowsRaw[ri];
                    if (src.Length < maxCols)
                    {
                        // 补齐到当前 sheet 最大列宽，避免下游 GetCell 时越界返回 null 造成分支开销。
                        var padded = new ICell[maxCols];
                        Array.Copy(src, padded, src.Length);
                        for (int k = src.Length; k < maxCols; k++) padded[k] = MiniExcelCell.Blank;
                        rowArr[ri] = new MiniExcelRow(ri, padded);
                    }
                    else
                    {
                        rowArr[ri] = new MiniExcelRow(ri, src);
                    }
                }

                sheets[si] = new MiniExcelSheet(sheetNames[si], rowArr);
            }

            return new MiniExcelWorkbook(sheets);
        }

        public string GetSheetName(int index) => _sheets[index].SheetName;

        public ISheet GetSheetAt(int index) => _sheets[index];

        public ISheet GetSheet(string name)
        {
            return _name2Index.TryGetValue(name, out var idx) ? _sheets[idx] : null;
        }

        public int GetSheetIndex(string name)
        {
            return _name2Index.TryGetValue(name, out var idx) ? idx : -1;
        }

        public void Dispose()
        {
            // 数据已全部驻留内存，无外部句柄需要释放。
        }
    }
}
