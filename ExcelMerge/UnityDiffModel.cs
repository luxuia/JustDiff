using System.Collections.Generic;
using NetDiff;

namespace ExcelMerge
{
    public class UnityNodeData
    {
        public string Path;
        public string Name;
        public string Type;
        public List<string> Components = new List<string>();
        public Dictionary<string, string> Properties = new Dictionary<string, string>();
        public List<UnityNodeData> Children = new List<UnityNodeData>();

        public override string ToString()
        {
            return Name;
        }
    }

    public class UnityNodeComparer : IEqualityComparer<UnityNodeData>
    {
        public bool Equals(UnityNodeData a, UnityNodeData b)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;
            return a.Name == b.Name;
        }

        public int GetHashCode(UnityNodeData a)
        {
            return a?.Name?.GetHashCode() ?? 0;
        }
    }

    public class PropertyDiff
    {
        public string Key;
        public string SrcValue;
        public string DstValue;
        public DiffStatus Status;
    }

    public class UnityDiffNode
    {
        public DiffStatus Status;
        public UnityNodeData SrcNode;
        public UnityNodeData DstNode;
        public List<PropertyDiff> ChangedProperties = new List<PropertyDiff>();
        public List<UnityDiffNode> Children = new List<UnityDiffNode>();

        public string DisplayName
        {
            get
            {
                var node = SrcNode ?? DstNode;
                return node?.Name ?? "(unknown)";
            }
        }

        public string DisplayComponents
        {
            get
            {
                var node = SrcNode ?? DstNode;
                if (node == null || node.Components.Count == 0) return "";
                return "[" + string.Join(", ", node.Components) + "]";
            }
        }

        public bool HasChanges
        {
            get
            {
                if (Status != DiffStatus.Equal) return true;
                foreach (var child in Children)
                {
                    if (child.HasChanges) return true;
                }
                return false;
            }
        }
    }
}
