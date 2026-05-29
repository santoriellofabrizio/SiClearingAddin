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
            if (colIdxs.Length == 0) return ExcelError.ExcelErrorValue;

            var rows = new List<int>();
            int rowCount = data.GetLength(0);
            for (int r = 1; r < rowCount; r++)
                if (data[r, isinCol].Equals(isin.Trim(), StringComparison.OrdinalIgnoreCase))
                    rows.Add(r);

            if (rows.Count == 0) return ExcelError.ExcelErrorNA;

            var result = new object[rows.Count + 1, colIdxs.Length];
            for (int c = 0; c < colIdxs.Length; c++)
                result[0, c] = data[0, colIdxs[c]];
            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < colIdxs.Length; c++)
                    result[r + 1, c] = data[rows[r], colIdxs[c]];

            return result;
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

            object[] colSpecs = cols == null || cols is ExcelMissing || cols is ExcelEmpty
                ? Array.Empty<object>()
                : cols is object[] arr ? arr : new object[] { cols };

            int[] colIdxs = Resolver.ResolveList(data, colSpecs);

            // collect distinct ISINs matching the date
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

            var result = new object[rows.Count + 1, colIdxs.Length];
            for (int c = 0; c < colIdxs.Length; c++)
                result[0, c] = data[0, colIdxs[c]];
            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < colIdxs.Length; c++)
                    result[r + 1, c] = data[rows[r], colIdxs[c]];

            return result;
        }

        [ExcelFunction(Name = "SiClearingLiveSaldo", IsVolatile = true,
            Description = "Combina sospesi storici con trades live per calcolare il saldo netto.")]
        public static object SiClearingLiveSaldo(
            [ExcelArgument(Name = "market", Description = "Mercato (es. ETFP)")] string market,
            [ExcelArgument(Name = "portfolio", Description = "Portfolio (es. ETF_EQUITY)")] string portfolio,
            [ExcelArgument(Name = "isin", Description = "ISIN specifico (opzionale)")] object? isin = null)
        {
            var data = AddIn.Cache.GetOrLoad(AddIn.Settings.SaveFolder);
            bool hasData = data != null && data.GetLength(0) > 0;

            // build historical balances
            var historical = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            if (hasData)
            {
                int isinCol = Resolver.Resolve(data!, "ISIN Code");
                if (isinCol < 0) isinCol = Resolver.Resolve(data!, "ISIN");
                int segnoCol = Resolver.Resolve(data!, "Segno");
                int qtaCol = Resolver.Resolve(data!, "Quantita'");

                if (isinCol >= 0 && segnoCol >= 0 && qtaCol >= 0)
                {
                    int rowCount = data!.GetLength(0);
                    for (int r = 1; r < rowCount; r++)
                    {
                        string code = data[r, isinCol].Trim();
                        string segno = data[r, segnoCol].Trim().ToUpperInvariant();
                        string qtaRaw = data[r, qtaCol].Trim();

                        if (string.IsNullOrEmpty(code) || string.IsNullOrEmpty(qtaRaw)) continue;

                        double qty;
                        try { qty = CsvCache.ParseItalianNumber(qtaRaw); }
                        catch { continue; }

                        double signed = segno == "A" ? qty : -qty;

                        if (!historical.TryGetValue(code, out double cur))
                            historical[code] = signed;
                        else
                            historical[code] = cur + signed;
                    }
                }
            }

            // get live trades
            var live = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            try
            {
                string formula =
                    $"=LET(market,\"{market}\",portfolio,\"{portfolio}\"," +
                    $"tab,dumaGetTableRecords(\"Trade\",market,TRUE)," +
                    $"op,rdsSlice(tab,,\"clientid\")," +
                    $"side,rdsSlice(tab,,\"side\")," +
                    $"qty,rdsSlice(tab,,\"tradeqty\")*IF(side=\"buy\",1,-1)," +
                    $"isin,rdsSlice(tab,,\"instrument.isincode\")," +
                    $"final_tab,FILTER(rdsHStack(isin,qty),op=portfolio)," +
                    $"rdsGroupBy(final_tab,1,\"sum\",2))";

                var raw = ExcelDnaUtil.Application is Microsoft.Office.Interop.Excel.Application app
                    ? app.Evaluate(formula)
                    : null;

                if (raw is object[,] table)
                {
                    int rows = table.GetLength(0);
                    for (int r = 0; r < rows; r++)
                    {
                        var isinVal = table[r, 0];
                        var qtyVal = table[r, 1];
                        if (isinVal is string s && qtyVal is double q)
                            live[s] = q;
                    }
                }
            }
            catch { /* RTD not available — proceed with historical only */ }

            string? isinFilter = null;
            if (isin != null && isin is string si && !string.IsNullOrWhiteSpace(si))
                isinFilter = si.Trim();
            else if (isin is ExcelMissing || isin is ExcelEmpty || isin == null)
                isinFilter = null;

            if (isinFilter != null)
            {
                double hist = historical.TryGetValue(isinFilter, out double h) ? h : 0;
                double lv = live.TryGetValue(isinFilter, out double l) ? l : 0;
                return hist + lv;
            }

            // all ISINs — union
            var allIsins = new HashSet<string>(historical.Keys, StringComparer.OrdinalIgnoreCase);
            foreach (var k in live.Keys) allIsins.Add(k);

            if (allIsins.Count == 0) return ExcelError.ExcelErrorNA;

            var result = new object[allIsins.Count + 1, 4];
            result[0, 0] = "ISIN";
            result[0, 1] = "Saldo Storico";
            result[0, 2] = "Saldo Live";
            result[0, 3] = "Saldo Totale";

            int idx = 1;
            foreach (var code in allIsins)
            {
                double h = historical.TryGetValue(code, out double hv) ? hv : 0;
                double l = live.TryGetValue(code, out double lv) ? lv : 0;
                result[idx, 0] = code;
                result[idx, 1] = h;
                result[idx, 2] = l;
                result[idx, 3] = h + l;
                idx++;
            }
            return result;
        }
    }
}
