using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;

namespace SiClearing
{
    public static class UDFs
    {
        private static readonly ColumnResolver Resolver = new ColumnResolver();

        [ExcelFunction(Name = "SiClearingHeaders",
            Description = "Restituisce indice e nome di ogni colonna del CSV SiClearing.")]
        public static object SiClearingHeaders()
        {
            var data = AddIn.Cache.GetOrLoad(AddIn.Settings.SaveFolder);
            if (data == null || data.GetLength(0) == 0)
                return ExcelError.ExcelErrorNA;

            int colCount = data.GetLength(1);
            var result = new object[colCount, 2];
            for (int c = 0; c < colCount; c++)
            {
                result[c, 0] = c + 1;
                result[c, 1] = data[0, c];
            }
            return result;
        }

        [ExcelFunction(Name = "SiClearingGetISIN",
            Description = "Restituisce righe del CSV per il codice ISIN specificato.")]
        public static object SiClearingGetISIN(
            [ExcelArgument(Name = "isin", Description = "Codice ISIN da cercare")] string isin,
            [ExcelArgument(Name = "cols", Description = "Nomi colonna o indici 1-based (opzionale)")] params object[] cols)
        {
            var data = AddIn.Cache.GetOrLoad(AddIn.Settings.SaveFolder);
            if (data == null || data.GetLength(0) == 0)
                return ExcelError.ExcelErrorNA;

            int isinCol = Resolver.Resolve(data, "ISIN Code");
            if (isinCol < 0) isinCol = Resolver.Resolve(data, "ISIN");
            if (isinCol < 0) return ExcelError.ExcelErrorValue;

            int[] colIdxs = Resolver.ResolveList(data, cols);

            var rows = new List<int>();
            int rowCount = data.GetLength(0);
            for (int r = 1; r < rowCount; r++)
                if (data[r, isinCol].Equals(isin.Trim(), StringComparison.OrdinalIgnoreCase))
                    rows.Add(r);

            if (rows.Count == 0) return ExcelError.ExcelErrorNA;

            return BuildResult(data, rows, colIdxs);
        }

        [ExcelFunction(Name = "SiClearingBuyInAlert",
            Description = "Restituisce ISIN con Buy-In Alert Date uguale alla data specificata.")]
        public static object SiClearingBuyInAlert(
            [ExcelArgument(Name = "alertDate", Description = "Data alert (numero seriale Excel o dd/mm/yyyy)")] object alertDate,
            [ExcelArgument(Name = "cols", Description = "Nomi colonna o indici (opzionale)")] object? cols = null)
        {
            var data = AddIn.Cache.GetOrLoad(AddIn.Settings.SaveFolder);
            if (data == null || data.GetLength(0) == 0)
                return ExcelError.ExcelErrorNA;

            DateTime targetDate;
            if (alertDate is double serial)
                targetDate = DateTime.FromOADate(serial).Date;
            else if (alertDate is string ds)
            {
                if (!DateTime.TryParseExact(ds, "dd/MM/yyyy", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out targetDate))
                    return ExcelError.ExcelErrorValue;
            }
            else return ExcelError.ExcelErrorValue;

            int buyInCol = Resolver.Resolve(data, "Buy-In Alert");
            if (buyInCol < 0) buyInCol = Resolver.Resolve(data, "BuyIn");
            if (buyInCol < 0) return ExcelError.ExcelErrorValue;

            int isinCol = Resolver.Resolve(data, "ISIN Code");
            if (isinCol < 0) isinCol = Resolver.Resolve(data, "ISIN");
            if (isinCol < 0) return ExcelError.ExcelErrorValue;

            int[] colIdxs = Resolver.ResolveFromOptional(data, cols);

            var seenIsin = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var rows = new List<int>();
            int rowCount = data.GetLength(0);

            for (int r = 1; r < rowCount; r++)
            {
                string raw = data[r, buyInCol].Trim();
                if (raw == "--" || string.IsNullOrEmpty(raw)) continue;

                if (!DateTime.TryParseExact(raw, "dd/MM/yyyy", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out DateTime dt)) continue;

                if (dt.Date != targetDate) continue;

                string isin = data[r, isinCol];
                if (!seenIsin.Add(isin)) continue;

                rows.Add(r);
            }

            if (rows.Count == 0) return ExcelError.ExcelErrorNA;

            return BuildResult(data, rows, colIdxs);
        }

        private static object[,] BuildResult(string[,] data, List<int> rows, int[] colIdxs)
        {
            var result = new object[rows.Count + 1, colIdxs.Length];
            for (int c = 0; c < colIdxs.Length; c++)
                result[0, c] = data[0, colIdxs[c]];
            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < colIdxs.Length; c++)
                    result[r + 1, c] = data[rows[r], colIdxs[c]];
            return result;
        }
    }
}
