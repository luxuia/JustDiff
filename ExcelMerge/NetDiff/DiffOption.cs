using System.Collections.Generic;

namespace NetDiff
{
    public class DiffOption<T>
    {
        public IEqualityComparer<T> EqualityComparer { get; set; }
        public int Limit { get; set; }

        // 优化模式，会多搜索一个分支，大数据容易崩溃
        public bool Optimize = true;
    }
}
