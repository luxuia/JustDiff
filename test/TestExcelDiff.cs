using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        var sw = Stopwatch.StartNew();
        var src = new WorkBookWrap(args[0], cfg);
        var t1 = sw.ElapsedMilliseconds;
        var dst = new WorkBookWrap(args[1], cfg);
        var t2 = sw.ElapsedMilliseconds;
        Console.WriteLine($"[PERF] Load src: {t1}ms, Load dst: {t2 - t1}ms, total: {t2}ms");
        Console.WriteLine($"src: {src.book.NumberOfSheets} sheets, dst: {dst.book.NumberOfSheets} sheets");

        var option = new DiffOption<SheetNameCombo>();
        option.EqualityComparer = new SheetNameComboComparer();
        var diffSheetName = DiffUtil.OptimizeCaseDeletedFirst(
            DiffUtil.Diff(src.sheetNameCombos, dst.sheetNameCombos, option)).ToList();

        long totalDiffTime = 0;
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

            var srcCols = s1.RowCount > 0 ? s1.GetRow(0)?.Cells.Count ?? 0 : 0;
            var dstCols = s2.RowCount > 0 ? s2.GetRow(0)?.Cells.Count ?? 0 : 0;

            // Measure per-sheet diff time
            sw.Restart();

            // Simulate what DiffConfigSheet does:
            // 1. GetHeaderStrList
            var head1 = GetHeaders(src, s1, cfg);
            var head2 = GetHeaders(dst, s2, cfg);
            var tHead = sw.ElapsedMilliseconds;

            // 2. Header diff
            var headList1 = GetHeaderStrList(src, s1, cfg);
            var headList2 = GetHeaderStrList(dst, s2, cfg);
            List<DiffResult<string>> headDiff = null;
            if (headList1 != null && headList2 != null)
            {
                headDiff = DiffUtil.OptimizeCaseDeletedFirst(
                    DiffUtil.Diff(headList1, headList2)).ToList();
            }
            var tHeadDiff = sw.ElapsedMilliseconds;

            // 3. ID diff (row matching)
            var idList1 = GetIDList(src, s1, cfg);
            var idList2 = GetIDList(dst, s2, cfg);
            var idOption = new DiffOption<string2int>();
            idOption.EqualityComparer = new SheetIDComparer();
            var idDiff = DiffUtil.Diff(idList1, idList2, idOption).ToList();
            var tIDDiff = sw.ElapsedMilliseconds;

            // 4. Row-by-row diff
            int changedRows = 0;
            int colCount = headList1?.Count ?? srcCols;
            foreach (var idResult in idDiff)
            {
                if (idResult.Obj1.Key != null && idResult.Obj2.Key != null)
                {
                    // Both exist, compare cells
                    var r1 = s1.GetRow(idResult.Obj1.Value);
                    var r2 = s2.GetRow(idResult.Obj2.Value);
                    bool rowChanged = false;
                    for (int c = 0; c < colCount; c++)
                    {
                        var v1 = r1?.GetCell(c)?.DisplayValue ?? "";
                        var v2 = r2?.GetCell(c)?.DisplayValue ?? "";
                        if (v1 != v2) { rowChanged = true; break; }
                    }
                    if (rowChanged) changedRows++;
                }
                else
                {
                    changedRows++;
                }
            }
            var tRowDiff = sw.ElapsedMilliseconds;
            totalDiffTime += tRowDiff;

            bool headerChanged = headDiff != null && headDiff.Any(a => a.Status != DiffStatus.Equal);
            bool idChanged = idDiff.Any(a => a.Status != DiffStatus.Equal);

            Console.WriteLine($"  SHEET [{name}] rows={s1.RowCount}/{s2.RowCount} cols={srcCols}/{dstCols} idRows={idList1.Count}/{idList2.Count} headChanged={headerChanged} idChanged={idChanged} changedRows={changedRows}");
            Console.WriteLine($"    [TIME] header={tHead}ms headDiff={tHeadDiff - tHead}ms idDiff={tIDDiff - tHeadDiff}ms rowDiff={tRowDiff - tIDDiff}ms total={tRowDiff}ms");
        }
        Console.WriteLine($"[PERF] Total diff time: {totalDiffTime}ms");
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
        for (int i = startcol; i < rows[0].Cells.Count; i++)
        {
            var str = "";
            for (int j = 0; j < rows.Count; j++)
            {
                var v = rows[j].GetCell(i)?.DisplayValue ?? "";
                str += (j > 0 ? ":" : "") + v;
            }
            if (string.IsNullOrWhiteSpace(str)) break;
            count++;
        }
        return count;
    }

    static List<string> GetHeaderStrList(WorkBookWrap wrap, ISheet sheet, Config cfg)
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
        if (rows.Count == 0) return null;
        var header = new List<string>();
        for (int i = startcol; i < rows[0].Cells.Count; i++)
        {
            var str = "";
            for (int j = 0; j < rows.Count; j++)
            {
                var v = rows[j].GetCell(i)?.DisplayValue ?? "";
                str += (j > 0 ? ":" + v : v);
            }
            if (string.IsNullOrWhiteSpace(str)) break;
            header.Add(str);
        }
        return header;
    }

    static List<string2int> GetIDList(WorkBookWrap wrap, ISheet sheet, Config cfg)
    {
        var sp = wrap.SheetStartPoint[sheet.SheetName];
        int startrow = sp.Item1;
        int startIdx = cfg.HeadCount + startrow;
        var validRow = wrap.SheetValideRow.ContainsKey(sheet.SheetName) ? wrap.SheetValideRow[sheet.SheetName] : sheet.RowCount;
        var list = new List<string2int>();
        for (int i = startIdx; i < validRow; i++)
        {
            var row = sheet.GetRow(i);
            if (row == null) break;
            var val = row.GetCell(0)?.DisplayValue ?? "";
            list.Add(new string2int(val, i));
        }
        return list;
    }

    static string Trunc(string s) => s.Length > 40 ? s.Substring(0, 40) + "..." : s;
}
