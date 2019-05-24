using System;
using System.Collections.Generic;
using System.Linq;

namespace NetDiff
{
    internal enum Direction
    {
        Right,
        Bottom,
        Diagonal,
    }

    internal struct Point : IEquatable<Point>
    {
        public int X { get; }
        public int Y { get; }

        public Point(int x, int y)
        {
            X = x;
            Y = y;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Point))
                return false;

            return Equals((Point)obj);
        }

        public override int GetHashCode()
        {
            var hash = 17;
            hash = hash * 23 + X.GetHashCode();
            hash = hash * 23 + Y.GetHashCode();

            return hash;
        }

        public bool Equals(Point other)
        {
            return X == other.X && Y == other.Y;
        }

        public override string ToString()
        {
            return $"X:{X} Y:{Y}";
        }
    }

    internal class Node
    {
        public Point Point { get; set; }
        public Node Parent { get; set; }
        public int Dis;

        public Node(Point point, int dis = 0)
        {
            Point = point;
            Dis = dis;
        }

        public override string ToString()
        {
            return $"X:{Point.X} Y:{Point.Y}";
        }
    }

    internal class EditGraph<T>
    {
        private T[] seq1;
        private T[] seq2;
        private DiffOption<T> option;
        private List<Node> heads;
        private Point endpoint;
        private int[] farthestPoints;
        private int offset;
        private bool isEnd;

        private Dictionary<int, Dictionary<int, Node>> visited;

        public EditGraph(
            IEnumerable<T> seq1, IEnumerable<T> seq2)
        {
            this.seq1 = seq1.ToArray();
            this.seq2 = seq2.ToArray();
            endpoint = new Point(this.seq1.Length, this.seq2.Length);
            offset = this.seq2.Length;
        }

        public List<Point> CalculatePath(DiffOption<T> option)
        {
            if (!seq1.Any())
                return Enumerable.Range(0, seq2.Length + 1).Select(i => new Point(0, i)).ToList();

            if (!seq2.Any())
                return Enumerable.Range(0, seq1.Length + 1).Select(i => new Point(i, 0)).ToList();

            this.option = option;

            //return DoNewCalcuatePath();

            BeginCalculatePath();

            while (Next()) { }

            return EndCalculatePath();
        }



        private List<Point> DoNewCalcuatePath() {
            var len1 = seq1.Length;
            var len2 = seq2.Length;
            int[,] maps = new int[len1+ 1,len2 + 1];
            for (var i = 0; i < len1 + 1; i++) maps[i, 0] = i;
            for (var i = 0; i < len2 + 1; i++) maps[0, i] = i;

            for (var i=1; i <len1+1; i++) {
                for (var j = 1; j < len2+1; j++) {
                    var equal = option.EqualityComparer != null
                        ? option.EqualityComparer.Equals(seq1[i - 1], (seq2[j - 1]))
                        : seq1[i - 1].Equals(seq2[j - 1]);
                    if (equal) {
                        maps[i, j] = maps[i - 1, j - 1];
                    } else {
                        maps[i, j] = Math.Min(maps[i - 1, j] + 1, Math.Min(maps[i, j - 1] + 1, maps[i - 1, j - 1] + 1));
                    }
                }
            }

            var waypoints = new List<Point>();
            int x = len1;
            int y = len2;

            do {
                waypoints.Add(new Point(x, y));
                int dis = maps[x, y];
                int left = x > 0 ? maps[x - 1, y] : -1;
                int up = y > 0 ? maps[x, y - 1] : -1;
                int dig = x > 0 && y > 0 ? maps[x - 1, y - 1] : -1;
                if (dig == dis - 1) {

                    x--; y--;
                }
                else if (up == dis - 1) {
                    y--;
                }
                else if (left == dis - 1) {
                    x--;
                }
                else if (dig == dis) {
                    // 这里要和上面的区分开，低优先级
                    x--; y--;
                }
                else {
                    throw new OverflowException();
                }
            } while (x > 0 || y > 0);
            waypoints.Add(new Point(0, 0));

            waypoints.Reverse();
            return waypoints;
        }

        private void Initialize()
        {
            farthestPoints = new int[seq1.Length + seq2.Length + 1];
            heads = new List<Node>();
            visited = new Dictionary<int, Dictionary<int, Node>>();
        }

        private void BeginCalculatePath()
        {
            Initialize();

            heads.Add(new Node(new Point(0, 0)));

            Snake();
        }

        private List<Point> EndCalculatePath()
        {
            var wayponit = new List<Point>();

            var current = heads.Where(h => h.Point.Equals(endpoint)).FirstOrDefault();
            while (current != null)
            {
                wayponit.Add(current.Point);

                current = current.Parent;
            }

            wayponit.Reverse();

            return wayponit;
        }

        private bool Next()
        {
            if (isEnd)
                return false;

            UpdateHeads();

            return true;
        }

        private void UpdateHeads()
        {
            if (option.Limit > 0 && heads.Count > option.Limit)
            {
                var tmp = heads.First();
                heads.Clear();

                heads.Add(tmp);
            }

            var updated = new List<Node>();

            foreach (var head in heads)
            {
                Node rightHead;
                if (TryCreateHead(head, Direction.Right, out rightHead))
                {
                    updated.Add(rightHead);
                }

                Node bottomHead;
                if (TryCreateHead(head, Direction.Bottom, out bottomHead))
                {
                    updated.Add(bottomHead);
                }

                //if (option.Optimize) {
                //    var diag = GetPoint(head.Point, Direction.Diagonal);
                //    if (InRange(diag)) {
                //        var newHead = new Node(diag);
                //        newHead.Parent = head;
                //
                //        isEnd |= newHead.Point.Equals(endpoint);
                //
                //        updated.Add(newHead);
                //    }
                //}
            }

            heads = updated;

            Snake();
        }

        private void Snake()
        {
            var tmp = new List<Node>();
            foreach (var h in heads)
            {
                var newHead = Snake(h);

                if (newHead != null)
                    tmp.Add(newHead);
                else
                    tmp.Add(h);
            }

            heads = tmp;
        }

        private Node Snake(Node head)
        {
            Node newHead = null;
            while (true)
            {
                Node tmp;
                if (TryCreateHead(newHead ?? head, Direction.Diagonal, out tmp))
                    newHead = tmp;
                else
                    break;
            }

            return newHead;
        }

        private bool TryCreateHead(Node head, Direction direction, out Node newHead)
        {
            newHead = null;
            var newPoint = GetPoint(head.Point, direction);

            if (!CanCreateHead(head.Point, direction, newPoint))
                return false;

            newHead = new Node(newPoint);
            newHead.Parent = head;

            isEnd |= newHead.Point.Equals(endpoint);

            Dictionary<int, Node> lines = null;
            if (!visited.TryGetValue(newPoint.X, out lines)) {
                visited[newPoint.X] = new Dictionary<int, Node>();
            }
            visited[newPoint.X][newPoint.Y] = newHead;

            return true;
        }

        private bool CanCreateHead(Point currentPoint, Direction direction, Point nextPoint)
        {
            if (!InRange(nextPoint))
                return false;

            if (direction == Direction.Diagonal)
            {
                var equal = option.EqualityComparer != null
                    ? option.EqualityComparer.Equals(seq1[nextPoint.X - 1], (seq2[nextPoint.Y - 1]))
                    : seq1[nextPoint.X - 1].Equals(seq2[nextPoint.Y - 1]);

                if (!equal)
                    return false;
            }

            if (visited.ContainsKey(nextPoint.X) && visited[nextPoint.X].ContainsKey(nextPoint.Y)) {
                return false;
            }
            return true;
            //return UpdateFarthestPoint(nextPoint);
        }

        private Point GetPoint(Point currentPoint, Direction direction)
        {
            switch (direction)
            {
                case Direction.Right:
                    return new Point(currentPoint.X + 1, currentPoint.Y);
                case Direction.Bottom:
                    return new Point(currentPoint.X, currentPoint.Y + 1);
                case Direction.Diagonal:
                    return new Point(currentPoint.X + 1, currentPoint.Y + 1);
            }

            throw new ArgumentException();
        }

        private bool InRange(Point point)
        {
            return point.X >= 0 && point.Y >= 0 && point.X <= endpoint.X && point.Y <= endpoint.Y;
        }

        private bool UpdateFarthestPoint(Point point)
        {
            var k = point.X - point.Y;
            var y = farthestPoints[k + offset];

            if (point.Y <= y)
                return false;

            farthestPoints[k + offset] = point.Y;

            return true;
        }
    }
}

