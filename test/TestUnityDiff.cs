using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ExcelMerge;

class TestUnityDiff
{
    static int passed = 0;
    static int failed = 0;
    static List<string> failures = new List<string>();

    static void Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "--textdiff")
        {
            TestTextDiffPerf();
            return;
        }

        if (args.Length > 0 && args[0] == "--parse-dir")
        {
            TestParseDir(args.Length > 1 ? args[1] : ".");
            return;
        }

        var testDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "test", "prefab_tests");
        testDir = Path.GetFullPath(testDir);

        if (!Directory.Exists(testDir))
        {
            Console.WriteLine($"Test directory not found: {testDir}");
            return;
        }

        var srcFiles = Directory.GetFiles(testDir, "*_src.prefab").OrderBy(f => f).ToArray();
        Console.WriteLine($"Found {srcFiles.Length} test pairs in {testDir}");
        Console.WriteLine(new string('=', 60));

        foreach (var srcFile in srcFiles)
        {
            var baseName = Path.GetFileName(srcFile).Replace("_src.prefab", "");
            var dstFile = Path.Combine(testDir, baseName + "_dst.prefab");

            if (!File.Exists(dstFile))
            {
                Console.WriteLine($"SKIP: {baseName} (no dst file)");
                continue;
            }

            RunTest(baseName, srcFile, dstFile);
        }

        Console.WriteLine(new string('=', 60));
        Console.WriteLine($"RESULTS: {passed} passed, {failed} failed, {passed + failed} total");
        if (failures.Count > 0)
        {
            Console.WriteLine("\nFAILURES:");
            foreach (var f in failures) Console.WriteLine($"  {f}");
        }
    }

    static void RunTest(string name, string srcFile, string dstFile)
    {
        Console.Write($"TEST: {name,-45} ");
        var sw = Stopwatch.StartNew();

        try
        {
            var srcRoots = UnityDiffEngine.ParseFile(srcFile);
            var dstRoots = UnityDiffEngine.ParseFile(dstFile);

            var diffNodes = UnityDiffEngine.DiffRoots(srcRoots, dstRoots);

            sw.Stop();

            int totalNodes = CountNodes(diffNodes);
            int changedNodes = CountChanged(diffNodes);
            int totalProps = CountProperties(diffNodes);

            Console.WriteLine($"OK  {sw.ElapsedMilliseconds,5}ms  roots:{srcRoots.Count}/{dstRoots.Count}  nodes:{totalNodes}  changed:{changedNodes}  props:{totalProps}");
            passed++;
        }
        catch (Exception ex)
        {
            sw.Stop();
            Console.WriteLine($"FAIL  {sw.ElapsedMilliseconds,5}ms  {ex.GetType().Name}: {ex.Message}");
            failures.Add($"{name}: {ex.GetType().Name}: {ex.Message}");
            failed++;
        }
    }

    static int CountNodes(List<UnityDiffNode> nodes)
    {
        int count = 0;
        foreach (var n in nodes)
        {
            count++;
            if (n.Children != null) count += CountNodes(n.Children);
        }
        return count;
    }

    static int CountChanged(List<UnityDiffNode> nodes)
    {
        int count = 0;
        foreach (var n in nodes)
        {
            if (n.HasChanges) count++;
            if (n.Children != null) count += CountChanged(n.Children);
        }
        return count;
    }

    static int CountProperties(List<UnityDiffNode> nodes)
    {
        int count = 0;
        foreach (var n in nodes)
        {
            if (n.ChangedProperties != null) count += n.ChangedProperties.Count;
            if (n.Children != null) count += CountProperties(n.Children);
        }
        return count;
    }

    static void TestParseDir(string dir)
    {
        var files = Directory.GetFiles(dir, "*.unity", SearchOption.AllDirectories)
            .Concat(Directory.GetFiles(dir, "*.prefab", SearchOption.AllDirectories))
            .OrderBy(f => f).ToArray();

        Console.WriteLine($"Parse test: {files.Length} files in {dir}");
        Console.WriteLine(new string('=', 80));

        int ok = 0, fail = 0;
        foreach (var file in files)
        {
            var name = file.Substring(dir.Length).TrimStart('\\', '/');
            var sw = Stopwatch.StartNew();
            try
            {
                var roots = UnityDiffEngine.ParseFile(file);
                sw.Stop();
                Console.WriteLine($"OK   {sw.ElapsedMilliseconds,5}ms  roots:{roots.Count,4}  {name}");
                ok++;
            }
            catch (Exception ex)
            {
                sw.Stop();
                Console.WriteLine($"FAIL {sw.ElapsedMilliseconds,5}ms  {ex.GetType().Name}: {ex.Message}  {name}");
                fail++;
            }
        }

        Console.WriteLine(new string('=', 80));
        Console.WriteLine($"RESULTS: {ok} ok, {fail} fail, {ok + fail} total");
    }

    static void TestTextDiffPerf()
    {
        var testDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "test", "prefab_tests");
        testDir = Path.GetFullPath(testDir);

        var srcFiles = Directory.GetFiles(testDir, "*_src.prefab").OrderBy(f => f).ToArray();
        Console.WriteLine($"Text diff performance test ({srcFiles.Length} files)");
        Console.WriteLine(new string('=', 80));

        foreach (var srcFile in srcFiles)
        {
            var baseName = Path.GetFileName(srcFile).Replace("_src.prefab", "");
            var dstFile = Path.Combine(testDir, baseName + "_dst.prefab");
            if (!File.Exists(dstFile)) continue;

            var srcLines = File.ReadAllLines(srcFile).Length;
            var dstLines = File.ReadAllLines(dstFile).Length;

            var sw = Stopwatch.StartNew();
            var result = UnityDiffEngine.DiffText(srcFile, dstFile);
            sw.Stop();

            Console.WriteLine($"{baseName,-40} {sw.ElapsedMilliseconds,5}ms  src:{srcLines,6} dst:{dstLines,6} result:{result.Count,6}");
        }
    }
}
