using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelMerge;
using NetDiff;
using string2int = System.Collections.Generic.KeyValuePair<string, int>;

class TestExcelDiff
{
    [STAThread]
    static void Main(string[] args)
    {
        if (args.Length < 2) { Console.WriteLine("Usage: TestExcelDiff src.xlsx dst.xlsx"); return; }

        var cfg = new Config();
        var src = new WorkBookWrap(args[0], cfg);
        var dst = new WorkBookWrap(args[1], cfg);

        Console.WriteLine($"src: {src.book.NumberOfSheets} sheets, dst: {dst.book.NumberOfSheets} sheets");

        var option = new DiffOption<SheetNameCombo>();
        option.EqualityComparer = new SheetNameComboComparer();
        var diffSheetName = DiffUtil.OptimizeCaseDeletedFirst(
            DiffUtil.Diff(src.sheetNameCombos, dst.sheetNameCombos, option)).ToList();

        foreach (var d in diffSheetName)
        {
            var name = d.Obj1?.Name ?? d.Obj2?.Name;
            if (d.Status != DiffStatus.Equal)
            {
                Console.WriteLine($"  SHEET [{name}] status={d.Status}");
                continue;
            }

            var s1 = src.book.GetSheetAt(d.Obj1.ID);
            var s2 = dst.book.GetSheetAt(d.Obj2.ID);

            // Check row counts
            var srcRows = src.SheetValideRow.ContainsKey(name) ? src.SheetValideRow[name] : -1;
            var dstRows = dst.SheetValideRow.ContainsKey(name) ? dst.SheetValideRow[name] : -1;
            var srcCols = s1.RowCount > 0 ? s1.GetRow(0)?.Cells.Count ?? 0 : 0;
            var dstCols = s2.RowCount > 0 ? s2.GetRow(0)?.Cells.Count ?? 0 : 0;

            // Quick cell-by-cell comparison
            int diffCells = 0;
            string firstDiff = null;
            int maxRows = Math.Max(s1.RowCount, s2.RowCount);
            int maxCols = Math.Max(srcCols, dstCols);
            for (int r = 0; r < maxRows && diffCells < 5; r++)
            {
                var r1 = s1.GetRow(r);
                var r2 = s2.GetRow(r);
                for (int c = 0; c < maxCols && diffCells < 5; c++)
                {
                    var v1 = r1?.GetCell(c)?.DisplayValue ?? "";
                    var v2 = r2?.GetCell(c)?.DisplayValue ?? "";
                    if (v1 != v2)
                    {
                        diffCells++;
                        if (firstDiff == null)
                            firstDiff = $"row={r} col={c} src=[{Trunc(v1)}] dst=[{Trunc(v2)}]";
                    }
                }
            }

            var realChanged = diffCells > 0;

            // Check what GetHeaderStrList would produce
            var srcHead = GetHeaders(src, s1, cfg);
            var dstHead = GetHeaders(dst, s2, cfg);

            Console.WriteLine($"  SHEET [{name}] srcRows={s1.RowCount} dstRows={s2.RowCount} srcCols={srcCols} dstCols={dstCols} headCols={srcHead}/{dstHead} validRow={srcRows}/{dstRows} diffCells={diffCells} realChanged={realChanged}");
            if (firstDiff != null)
                Console.WriteLine($"    first diff: {firstDiff}");
            if (srcCols != dstCols && diffCells == 0)
            {
                int minC = Math.Min(srcCols, dstCols);
                int maxC = Math.Max(srcCols, dstCols);
                var wider = srcCols > dstCols ? s1 : s2;
                var tag = srcCols > dstCols ? "src" : "dst";
                Console.WriteLine($"    TRAILING cols [{minC}..{maxC-1}] in {tag}:");
                for (int c = minC; c < Math.Min(maxC, minC + 5); c++)
                {
                    for (int r = 0; r < Math.Min(wider.RowCount, 3); r++)
                    {
                        var cell = wider.GetRow(r)?.GetCell(c);
                        var ct = cell?.CellType.ToString() ?? "null";
                        var dv = cell?.DisplayValue ?? "(null)";
                        Console.WriteLine($"      col={c} row={r} type={ct} display=[{dv}]");
                    }
                }
            }
        }
    }

    static int GetHeaders(WorkBookWrap wrap, ISheet sheet, Config cfg)
    {
        var sp = wrap.SheetStartPoint[sheet.SheetName];
        int startrow = sp.Item1, startcol = sp.Item2;
        int headEnd = cfg.HeadCount + startrow;
        var rows = new List<IRow>();
        for (int i = startrow; i < headEnd; i++)
        {
            var r = sheet.GetRow(i);
            if (r != null) rows.Add(r);
        }
        if (rows.Count == 0) return -1;
        int count = 0;
        bool foundEmpty = false;
        for (int i = startcol; i < rows[0].Cells.Count; i++)
        {
            var str = "";
            for (int j = 0; j < rows.Count; j++)
            {
                var v = rows[j].GetCell(i)?.DisplayValue ?? "";
                str += (j > 0 ? ":" : "") + v;
            }
            if (string.IsNullOrWhiteSpace(str))
            {
                if (!foundEmpty)
                {
                    foundEmpty = true;
                    Console.WriteLine($"    HEAD [{sheet.SheetName}] empty at col={i}, but cells.count={rows[0].Cells.Count}");
                    // print next few cols
                    for (int k = i; k < Math.Min(i + 5, rows[0].Cells.Count); k++)
                    {
                        var sv = "";
                        for (int j = 0; j < rows.Count; j++)
                        {
                            var v = rows[j].GetCell(k)?.DisplayValue ?? "";
                            sv += (j > 0 ? "|" : "") + v;
                        }
                        Console.WriteLine($"      col[{k}] = [{sv}]");
                    }
                }
                break;
            }
            count++;
        }
        return count;
    }

    static string Trunc(string s) => s.Length > 40 ? s.Substring(0, 40) + "..." : s;
}
