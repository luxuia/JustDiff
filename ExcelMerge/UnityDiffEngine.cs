using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NetDiff;

namespace ExcelMerge
{
    public class UnityDiffEngine
    {
        public static HashSet<string> IgnoredPropertyKeys = new HashSet<string>
        {
            "m_RootOrder",
        };

        public static bool IgnorePositionChanges = true;
        public static bool IgnoreFileIdChanges = true;
        public static bool OnlyNodeChanges = true;

        static bool ShouldIgnoreProperty(string key, string srcVal, string dstVal)
        {
            if (IgnorePositionChanges)
            {
                if (IgnoredPropertyKeys.Contains(key)) return true;
                if (key.StartsWith("m_Children[")) return true;
                if (key.StartsWith("m_Component[")) return true;
            }
            if (IgnoreFileIdChanges)
            {
                if (key.EndsWith(".fileID") || key == "fileID") return true;
                if (key.EndsWith(".guid") || key == "guid") return true;
            }
            return false;
        }

        public static List<UnityNodeData> ParseFile(string filePath)
        {
            var scene = ParseUnityYaml.ParseYaml(filePath);
            var roots = new List<UnityNodeData>();
            foreach (var go in scene.roots)
            {
                roots.Add(ConvertToNodeData(go, ""));
            }
            return roots;
        }

        static UnityNodeData ConvertToNodeData(UnityGameObject go, string parentPath)
        {
            var node = new UnityNodeData();
            node.Name = go.name ?? "(unnamed)";
            node.Path = string.IsNullOrEmpty(parentPath) ? node.Name : parentPath + "/" + node.Name;
            node.Type = go.comps.Count > 0 ? "GameObject" : "PrefabInstance";
            node.Components = new List<string>(go.comps);

            if (go.data != null)
            {
                FlattenProperties(go.data, "", node.Properties);
            }

            foreach (var child in go.childs)
            {
                node.Children.Add(ConvertToNodeData(child, node.Path));
            }
            node.Children.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.Ordinal));

            return node;
        }

        static void FlattenProperties(object obj, string prefix, Dictionary<string, string> result)
        {
            if (obj == null)
            {
                if (!string.IsNullOrEmpty(prefix))
                    result[prefix] = "";
                return;
            }

            if (obj is IDictionary dict)
            {
                foreach (DictionaryEntry entry in dict)
                {
                    var keyStr = entry.Key?.ToString() ?? "";
                    var key = string.IsNullOrEmpty(prefix) ? keyStr : prefix + "." + keyStr;
                    var val = entry.Value;

                    if (val is IDictionary || val is IList)
                    {
                        FlattenProperties(val, key, result);
                    }
                    else
                    {
                        result[key] = val?.ToString() ?? "";
                    }
                }
            }
            else if (obj is IList list)
            {
                for (int i = 0; i < list.Count; i++)
                {
                    FlattenProperties(list[i], prefix + "[" + i + "]", result);
                }
            }
            else
            {
                result[string.IsNullOrEmpty(prefix) ? "_value" : prefix] = obj.ToString() ?? "";
            }
        }

        public static List<UnityDiffNode> DiffRoots(List<UnityNodeData> srcRoots, List<UnityNodeData> dstRoots, IProgress<double> progress = null)
        {
            return DiffNodeLists(srcRoots, dstRoots, progress, 0);
        }

        static List<UnityDiffNode> DiffNodeLists(List<UnityNodeData> srcList, List<UnityNodeData> dstList, IProgress<double> progress = null, int depth = 0)
        {
            var option = new DiffOption<UnityNodeData>();
            option.EqualityComparer = new UnityNodeComparer();
            var diffResults = DiffUtil.Diff(srcList, dstList, option);
            var optimized = DiffUtil.OptimizeCaseDeletedFirst(diffResults).ToList();

            var result = new List<UnityDiffNode>();
            bool reportProgress = depth == 0 && progress != null && optimized.Count > 0;
            int done = 0;
            foreach (var dr in optimized)
            {
                var node = new UnityDiffNode();
                node.SrcNode = dr.Obj1;
                node.DstNode = dr.Obj2;

                switch (dr.Status)
                {
                    case DiffStatus.Equal:
                        node.ChangedProperties = OnlyNodeChanges
                            ? new List<PropertyDiff>()
                            : DiffProperties(dr.Obj1.Properties, dr.Obj2.Properties);
                        node.Children = DiffNodeLists(dr.Obj1.Children, dr.Obj2.Children);
                        if (OnlyNodeChanges)
                            node.Status = DiffStatus.Equal;
                        else
                            node.Status = (node.ChangedProperties.Count > 0 || node.Children.Any(c => c.HasChanges))
                                ? DiffStatus.Modified : DiffStatus.Equal;
                        break;
                    case DiffStatus.Modified:
                        node.ChangedProperties = OnlyNodeChanges
                            ? new List<PropertyDiff>()
                            : DiffProperties(dr.Obj1.Properties, dr.Obj2.Properties);
                        node.Children = DiffNodeLists(dr.Obj1.Children, dr.Obj2.Children);
                        node.Status = DiffStatus.Modified;
                        break;
                    case DiffStatus.Deleted:
                        node.Status = DiffStatus.Deleted;
                        break;
                    case DiffStatus.Inserted:
                        node.Status = DiffStatus.Inserted;
                        break;
                }
                result.Add(node);
                if (reportProgress)
                {
                    done++;
                    progress.Report((double)done / optimized.Count);
                }
            }
            return result;
        }

        static List<PropertyDiff> DiffProperties(Dictionary<string, string> src, Dictionary<string, string> dst)
        {
            var result = new List<PropertyDiff>();
            var allKeys = new HashSet<string>(src.Keys);
            allKeys.UnionWith(dst.Keys);

            foreach (var key in allKeys.OrderBy(k => k))
            {
                var hasSrc = src.TryGetValue(key, out var srcVal);
                var hasDst = dst.TryGetValue(key, out var dstVal);

                if (ShouldIgnoreProperty(key, srcVal, dstVal)) continue;

                if (hasSrc && hasDst)
                {
                    if (srcVal != dstVal)
                    {
                        result.Add(new PropertyDiff { Key = key, SrcValue = srcVal, DstValue = dstVal, Status = DiffStatus.Modified });
                    }
                }
                else if (hasSrc)
                {
                    result.Add(new PropertyDiff { Key = key, SrcValue = srcVal, DstValue = null, Status = DiffStatus.Deleted });
                }
                else
                {
                    result.Add(new PropertyDiff { Key = key, SrcValue = null, DstValue = dstVal, Status = DiffStatus.Inserted });
                }
            }
            return result;
        }

        public static List<DiffResult<string>> DiffText(string file1, string file2, IProgress<double> progress = null)
        {
            var lines1 = File.Exists(file1) ? File.ReadAllLines(file1) : Array.Empty<string>();
            var lines2 = File.Exists(file2) ? File.ReadAllLines(file2) : Array.Empty<string>();

            var chunks1 = SplitYamlChunks(lines1);
            var chunks2 = SplitYamlChunks(lines2);

            if (chunks1.Count <= 1 && chunks2.Count <= 1)
                return DiffTextDirect(lines1.ToList(), lines2.ToList(), progress);

            var chunkDiff = DiffUtil.Diff(chunks1, chunks2, new DiffOption<YamlChunk> { EqualityComparer = new YamlChunkComparer() });
            var optimized = DiffUtil.OptimizeCaseDeletedFirst(chunkDiff).ToList();

            var result = new List<DiffResult<string>>();
            int done = 0;

            foreach (var cd in optimized)
            {
                switch (cd.Status)
                {
                    case DiffStatus.Equal:
                        var innerDiff = DiffUtil.Diff(cd.Obj1.Lines, cd.Obj2.Lines);
                        result.AddRange(DiffUtil.OptimizeCaseDeletedFirst(innerDiff));
                        break;
                    case DiffStatus.Modified:
                        var modDiff = DiffUtil.Diff(cd.Obj1.Lines, cd.Obj2.Lines);
                        result.AddRange(DiffUtil.OptimizeCaseDeletedFirst(modDiff));
                        break;
                    case DiffStatus.Deleted:
                        foreach (var line in cd.Obj1.Lines)
                            result.Add(new DiffResult<string>(line, null, DiffStatus.Deleted));
                        break;
                    case DiffStatus.Inserted:
                        foreach (var line in cd.Obj2.Lines)
                            result.Add(new DiffResult<string>(null, line, DiffStatus.Inserted));
                        break;
                }
                done++;
                progress?.Report((double)done / optimized.Count);
            }

            return result;
        }

        static List<DiffResult<string>> DiffTextDirect(List<string> lines1, List<string> lines2, IProgress<double> progress)
        {
            const int chunkSize = 2000;
            if (lines1.Count <= chunkSize && lines2.Count <= chunkSize)
            {
                progress?.Report(0.5);
                var diff = DiffUtil.Diff(lines1, lines2);
                var r = DiffUtil.OptimizeCaseDeletedFirst(diff).ToList();
                progress?.Report(1.0);
                return r;
            }

            progress?.Report(0.1);
            var result = new List<DiffResult<string>>();
            int maxLines = Math.Max(lines1.Count, lines2.Count);
            for (int offset = 0; offset < maxLines; offset += chunkSize)
            {
                var c1 = lines1.Skip(offset).Take(chunkSize).ToList();
                var c2 = lines2.Skip(offset).Take(chunkSize).ToList();
                if (c1.Count == 0 && c2.Count == 0) break;
                var diff = DiffUtil.Diff(c1, c2);
                result.AddRange(DiffUtil.OptimizeCaseDeletedFirst(diff));
                progress?.Report((double)(offset + chunkSize) / maxLines);
            }
            return result;
        }

        class YamlChunk
        {
            public string Header;
            public List<string> Lines = new List<string>();
        }

        class YamlChunkComparer : IEqualityComparer<YamlChunk>
        {
            public bool Equals(YamlChunk a, YamlChunk b)
            {
                if (a == null && b == null) return true;
                if (a == null || b == null) return false;
                return a.Header == b.Header;
            }
            public int GetHashCode(YamlChunk c) => c?.Header?.GetHashCode() ?? 0;
        }

        static List<YamlChunk> SplitYamlChunks(string[] lines)
        {
            var chunks = new List<YamlChunk>();
            YamlChunk current = null;

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                if (line.StartsWith("--- !u!"))
                {
                    current = new YamlChunk();
                    current.Header = line;
                    if (i + 1 < lines.Length)
                    {
                        current.Header = lines[i + 1].TrimEnd(':') + " " + line;
                    }
                    current.Lines.Add(line);
                    chunks.Add(current);
                }
                else if (current != null)
                {
                    current.Lines.Add(line);
                }
                else
                {
                    if (chunks.Count == 0)
                    {
                        current = new YamlChunk { Header = "__preamble__" };
                        chunks.Add(current);
                    }
                    chunks[0].Lines.Add(line);
                }
            }
            return chunks;
        }
    }
}
