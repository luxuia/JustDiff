﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace NetDiff
{
    public class DiffUtil
    {
        public static IEnumerable<DiffResult<T>> Diff<T>(IEnumerable<T> seq1, IEnumerable<T> seq2)
        {
            return Diff(seq1, seq2, new DiffOption<T>());
        }

        public static IEnumerable<DiffResult<T>> Diff<T>(IEnumerable<T> seq1, IEnumerable<T> seq2, DiffOption<T> option)
        {
            if (seq1 == null || seq2 == null || (!seq1.Any() && !seq2.Any()))
                return Enumerable.Empty<DiffResult<T>>();

            var editGrap = new EditGraph<T>(seq1, seq2);
            var waypoints = editGrap.CalculatePath(option);

            return MakeResults<T>(waypoints, seq1, seq2, option);
        }

        public static IEnumerable<T> CreateSrc<T>(IEnumerable<DiffResult<T>> diffResults)
        {
            return diffResults.Where(r => r.Status != DiffStatus.Inserted).Select(r => r.Obj1);
        }

        public static IEnumerable<T> CreateDst<T>(IEnumerable<DiffResult<T>> diffResults)
        {
            return diffResults.Where(r => r.Status != DiffStatus.Deleted).Select(r => r.Obj2);
        }

        public static IEnumerable<DiffResult<T>> OptimizeCaseDeletedFirst<T>(IEnumerable<DiffResult<T>> diffResults)
        {
            return Optimize(diffResults, true);
        }

        public static IEnumerable<DiffResult<T>> OptimizeCaseInsertedFirst<T>(IEnumerable<DiffResult<T>> diffResults)
        {
            return Optimize(diffResults, false);
        }

        private static IEnumerable<DiffResult<T>> Optimize<T>(IEnumerable<DiffResult<T>> diffResults, bool deleteFirst = true)
        {
            var currentStatus = deleteFirst ? DiffStatus.Deleted : DiffStatus.Inserted;
            var nextStatus = deleteFirst ? DiffStatus.Inserted : DiffStatus.Deleted;

            var queue = new Queue<DiffResult<T>>(diffResults);
            var list = diffResults.ToList();

            int j = 0;
            int optCount = 0;
            for (var i = 0; i < list.Count;) {
                j = i + 1;
                optCount = 0;
                
                if (list[i].Status == currentStatus) {
                    while (j < list.Count && list[j].Status == currentStatus ) {
                        j++;
                    }
                    if ( j<list.Count && list[j].Status == nextStatus) {
                        while ( optCount < (j-i) && j+optCount < list.Count && list[j+optCount].Status == nextStatus) {
                            optCount++;
                        }
                    }
                }
                while (i < (j-optCount) && i < list.Count) {
                    yield return list[i];
                    i++;
                }
                while (i< j && i < list.Count) {
                    var obj1 = deleteFirst ? list[i].Obj1 : list[i+optCount].Obj1;
                    var obj2 = deleteFirst ? list[i+optCount].Obj2 : list[i].Obj2;
                    yield return new DiffResult<T>(obj1, obj2, DiffStatus.Modified);
                    i++;
                }
                i += optCount;
            }

            /*
            while (queue.Any())
            {
                var result = queue.Dequeue();
                if (result.Status == currentStatus)
                {
                    if (queue.Any() && queue.Peek().Status == nextStatus)
                    {
                        var obj1 = deleteFirst ? result.Obj1 : queue.Dequeue().Obj1;
                        var obj2 = deleteFirst ? queue.Dequeue().Obj2 : result.Obj2;
                        yield return new DiffResult<T>(obj1, obj2, DiffStatus.Modified);
                    }
                    else
                        yield return result;

                    continue;
                }

                yield return result;
            }
            */
        }

        /*  A    0   0   nil
        *  nil  0   0   0 
        *  del          ins
        *  0   0   0 
        *  modify        
        *  
        *  nil 0   0   0
        *  A   0   0   nil
        *  ins         del
        *  
        *  A   0   0
        *  modify
        *  
        */
        public static IEnumerable<DiffResult<T>> OptimizeShift<T>(IEnumerable<DiffResult<T>> diffResults, bool deleteFirst = true) {
            var currentStatus = deleteFirst ? DiffStatus.Deleted : DiffStatus.Inserted;
            var nextStatus = deleteFirst ? DiffStatus.Inserted : DiffStatus.Deleted;

            var ret_list = new List<DiffResult<T>>();
            //var queue = new Queue<DiffResult<T>>(diffResults);
            var list = diffResults.ToList();

            int j = 0;
            int optCount = 0;
            for (var i = 0; i < list.Count;) {
                j = i + 1;
                optCount = 0;
  
                if (list[i].Status == currentStatus) {
                    while (j < list.Count && list[j].Status == currentStatus) {
                        j++;
                    }
                    // equal or modify
                    while (j + optCount < list.Count && list[j + optCount].Status != nextStatus) {
                        optCount++;
                    }

                    if (j + optCount < list.Count) {
                        // 只处理删1增1的情况
                        while (i < j - 1) {
                            ret_list.Add(list[i]);
                            i++;
                        }

                        int oldi = i;
                        int diffcount = 0;
                        var test_list = new List<DiffResult<T>>();
                        while (i < j + optCount) {
                            var obj1 = deleteFirst ? list[i].Obj1 : list[i + 1].Obj1;
                            var obj2 = deleteFirst ? list[i + 1].Obj2 : list[i].Obj2;
                            if (obj1 == null) {
                                obj1 = (T)(object)string.Empty;
                            }
                            if (obj2 == null) {
                                obj2 = (T)(object)string.Empty;
                            }
                            var status = obj2.Equals(obj1) ? DiffStatus.Equal : DiffStatus.Modified;
                            if (status == DiffStatus.Modified) {
                                diffcount++;
                                // 超过一个修改，认为不应该优化，回退修改
                                if (diffcount > 1) {
                                    break;
                                }
                            }

                            test_list.Add(new DiffResult<T>(obj1, obj2, status));
                            i++;
                        }
                        if (diffcount <= 1) {
                            ret_list.AddRange(test_list);
                            //跳过最后一个优化掉的insert
                            i += 1;
                        } else {
                            i = oldi;
                        }

                    }
                }

                while (i < j) {
                    ret_list.Add(list[i]);
                    i++;
                }
            }
            return ret_list;
        }


        private static IEnumerable<DiffResult<T>> MakeResults<T>(IEnumerable<Point> waypoints, IEnumerable<T> seq1, IEnumerable<T> seq2, DiffOption<T> option)
        {
            var array1 = seq1.ToArray();
            var array2 = seq2.ToArray();

            foreach (var pair in waypoints.MakePairsWithNext())
            {
                var status = GetStatus(pair.Item1, pair.Item2, ref array1, ref array2, option);
                T obj1 = default(T);
                T obj2 = default(T);
                switch (status)
                {
                    case DiffStatus.Equal:
                    case DiffStatus.Modified:
                        obj1 = array1[pair.Item2.X - 1];
                        obj2 = array2[pair.Item2.Y - 1];
                        break;
                    case DiffStatus.Inserted:
                        obj2 = array2[pair.Item2.Y - 1];
                        break;
                    case DiffStatus.Deleted:
                        obj1 = array1[pair.Item2.X - 1];
                        break;
                }

                yield return new DiffResult<T>(obj1, obj2, status);
            }
        }

        private static DiffStatus GetStatus<T>(Point current, Point prev, ref T[] array1, ref T[] array2, DiffOption<T> option)
        {
            if (current.X != prev.X && current.Y != prev.Y) {
                var equal = option.EqualityComparer != null ? option.EqualityComparer.Equals(array1[prev.X - 1], (array2[prev.Y - 1])) : array1[prev.X - 1].Equals(array2[prev.Y - 1]);

                if (equal) {
                    return DiffStatus.Equal;
                }
                else {
                    return DiffStatus.Modified;
                }
            }
            else if (current.X != prev.X)
                return DiffStatus.Deleted;
            else if (current.Y != prev.Y)
                return DiffStatus.Inserted;
            else
                throw new Exception();
        }
        const int BYTES_TO_READ = sizeof(Int64);

        public static bool FilesAreEqual(string path1, string path2) {
            var first = new FileInfo(path1);
            var second = new FileInfo(path2);

            if (first.Length != second.Length)
                return false;

            if (string.Equals(first.FullName, second.FullName, StringComparison.OrdinalIgnoreCase))
                return true;

            int iterations = (int)Math.Ceiling((double)first.Length / BYTES_TO_READ);

            using (FileStream fs1 = first.OpenRead())
            using (FileStream fs2 = second.OpenRead()) {
                byte[] one = new byte[BYTES_TO_READ];
                byte[] two = new byte[BYTES_TO_READ];

                for (int i = 0; i < iterations; i++) {
                    fs1.Read(one, 0, BYTES_TO_READ);
                    fs2.Read(two, 0, BYTES_TO_READ);

                    if (BitConverter.ToInt64(one, 0) != BitConverter.ToInt64(two, 0))
                        return false;
                }
            }

            return true;
        }


        public static IEnumerable<DiffResult<T>> Order<T>(IEnumerable<DiffResult<T>> results, DiffOrderType orderType)
        {
            var resultArray = results.ToArray();

            for (int i = 0; i < resultArray.Length; i++)
            {
                if (resultArray[i].Status == DiffStatus.Deleted)
                {
                    while (i - 1 >= 0)
                    {
                        if (resultArray[i - 1].Status == DiffStatus.Equal && resultArray[i].Obj1.Equals(resultArray[i - 1].Obj1))
                        {
                            var tmp = resultArray[i];
                            resultArray[i] = resultArray[i - 1];
                            resultArray[i - 1] = tmp;

                            i--;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }

            var resultQueue = new Queue<DiffResult<T>>(resultArray);
            var additionQueue = new Queue<DiffResult<T>>();
            var deletionQueue = new Queue<DiffResult<T>>();

            while (resultQueue.Any())
            {
                if (resultQueue.Peek().Status == DiffStatus.Equal)
                {
                    yield return resultQueue.Dequeue();
                    continue;
                }

                while (resultQueue.Any() && resultQueue.Peek().Status != DiffStatus.Equal)
                {
                    while (resultQueue.Any() && resultQueue.Peek().Status == DiffStatus.Inserted)
                    {
                        additionQueue.Enqueue(resultQueue.Dequeue());
                    }

                    while (resultQueue.Any() && resultQueue.Peek().Status == DiffStatus.Deleted)
                    {
                        deletionQueue.Enqueue(resultQueue.Dequeue());
                    }
                }

                var latestReturenStatus = DiffStatus.Equal;
                while (true)
                {
                    if (additionQueue.Any() && !deletionQueue.Any())
                    {
                        yield return additionQueue.Dequeue();
                    }
                    else if (!additionQueue.Any() && deletionQueue.Any())
                    {
                        yield return deletionQueue.Dequeue();
                    }
                    else if (additionQueue.Any() && deletionQueue.Any())
                    {
                        switch (orderType)
                        {
                            case DiffOrderType.GreedyDeleteFirst:
                                yield return deletionQueue.Dequeue();
                                latestReturenStatus = DiffStatus.Deleted;
                                break;
                            case DiffOrderType.GreedyInsertFirst:
                                yield return additionQueue.Dequeue();
                                latestReturenStatus = DiffStatus.Inserted;
                                break;
                            case DiffOrderType.LazyDeleteFirst:
                                if (latestReturenStatus != DiffStatus.Deleted)
                                {
                                    yield return deletionQueue.Dequeue();
                                    latestReturenStatus = DiffStatus.Deleted;
                                }
                                else
                                {
                                    yield return additionQueue.Dequeue();
                                    latestReturenStatus = DiffStatus.Inserted;
                                }
                                break;
                            case DiffOrderType.LazyInsertFirst:
                                if (latestReturenStatus != DiffStatus.Inserted)
                                {
                                    yield return additionQueue.Dequeue();
                                    latestReturenStatus = DiffStatus.Inserted;
                                }
                                else
                                {
                                    yield return deletionQueue.Dequeue();
                                    latestReturenStatus = DiffStatus.Deleted;
                                }
                                break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }
    }
}
