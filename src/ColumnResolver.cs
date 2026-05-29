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

        public int[] ResolveList(string[,] data, object[] specs)
        {
            int colCount = data.GetLength(1);

            if (specs == null || specs.Length == 0)
            {
                var all = new int[colCount];
                for (int i = 0; i < colCount; i++) all[i] = i;
                return all;
            }

            var result = new List<int>();
            foreach (var spec in specs)
            {
                if (spec is object[] arr)
                {
                    foreach (var item in arr)
                    {
                        int idx = Resolve(data, item);
                        if (idx >= 0) result.Add(idx);
                    }
                }
                else if (spec != null && !(spec is ExcelDna.Integration.ExcelMissing) && !(spec is ExcelDna.Integration.ExcelEmpty))
                {
                    int idx = Resolve(data, spec);
                    if (idx >= 0) result.Add(idx);
                }
            }

            if (result.Count == 0)
            {
                var all = new int[colCount];
                for (int i = 0; i < colCount; i++) all[i] = i;
                return all;
            }

            return result.ToArray();
        }
    }
}
