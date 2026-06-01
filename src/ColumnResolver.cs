using System;
using System.Collections.Generic;

namespace SiClearing
{
    public class ColumnResolver
    {
        public int Resolve(string[,] data, object colSpec)
        {
            int colCount = data.GetLength(1);

            if (colSpec is double d)
            {
                int idx = (int)d - 1;
                return (idx >= 0 && idx < colCount) ? idx : -1;
            }

            if (colSpec is string name)
            {
                name = name.Trim();
                for (int c = 0; c < colCount; c++)
                    if (data[0, c].IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                        return c;
                return -1;
            }

            return -1;
        }

        private IEnumerable<object> Flatten(IEnumerable<object> specs)
        {
            foreach (var s in specs)
            {
                if (s is object[,] grid)
                {
                    int rows = grid.GetLength(0), cols = grid.GetLength(1);
                    for (int r = 0; r < rows; r++)
                        for (int c = 0; c < cols; c++)
                            yield return grid[r, c];
                }
                else if (s is object[] arr)
                {
                    foreach (var item in arr) yield return item;
                }
                else
                {
                    yield return s;
                }
            }
        }

        /// <summary>
        /// Resolves a list of column specs.
        /// If specs is empty/missing and defaultCols is non-empty, uses defaultCols (semicolon-separated names).
        /// Falls back to all columns only when both are empty.
        /// </summary>
        public int[] ResolveList(string[,] data, object[] specs, string defaultCols = "")
        {
            int colCount = data.GetLength(1);

            bool specsEmpty = specs == null || specs.Length == 0;

            if (specsEmpty)
            {
                if (!string.IsNullOrWhiteSpace(defaultCols))
                    return ResolveDefaultCols(data, defaultCols);
                return AllColumns(colCount);
            }

            var result = new List<int>();
            foreach (var spec in Flatten(specs))
            {
                if (spec == null
                    || spec is ExcelDna.Integration.ExcelMissing
                    || spec is ExcelDna.Integration.ExcelEmpty)
                    continue;

                int idx = Resolve(data, spec);
                if (idx >= 0) result.Add(idx);
            }

            if (result.Count == 0)
            {
                if (!string.IsNullOrWhiteSpace(defaultCols))
                    return ResolveDefaultCols(data, defaultCols);
                return AllColumns(colCount);
            }

            return result.ToArray();
        }

        public int[] ResolveFromOptional(string[,] data, object? colsArg, string defaultCols = "")
        {
            if (colsArg == null
                || colsArg is ExcelDna.Integration.ExcelMissing
                || colsArg is ExcelDna.Integration.ExcelEmpty)
            {
                if (!string.IsNullOrWhiteSpace(defaultCols))
                    return ResolveDefaultCols(data, defaultCols);
                return AllColumns(data.GetLength(1));
            }

            object[] wrapped = colsArg is object[] arr ? arr : new object[] { colsArg };
            return ResolveList(data, wrapped, defaultCols);
        }

        private int[] ResolveDefaultCols(string[,] data, string defaultCols)
        {
            var names = defaultCols.Split(';');
            var result = new List<int>();
            foreach (var name in names)
            {
                int idx = Resolve(data, name.Trim());
                if (idx >= 0) result.Add(idx);
            }
            return result.Count > 0 ? result.ToArray() : AllColumns(data.GetLength(1));
        }

        private static int[] AllColumns(int count)
        {
            var all = new int[count];
            for (int i = 0; i < count; i++) all[i] = i;
            return all;
        }
    }
}
