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
            if (!TryParseDate(alertDate, out targetDate))
                return ExcelError.ExcelErrorValue;

            int buyInCol = Resolver.Resolve(data, "Buy-In Alert");
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
                if (!TryParseDateString(raw, out DateTime dt)) continue;
                if (dt.Date != targetDate) continue;

                string isin = data[r, isinCol];
                if (!seenIsin.Add(isin)) continue;
                rows.Add(r);
            }

            if (rows.Count == 0) return ExcelError.ExcelErrorNA;
            return BuildResult(data, rows, colIdxs);
        }

        [ExcelFunction(Name = "SiClearingBuyInSaldi",
            Description = "GroupBy ISIN delle quantità per una Buy-In Alert Date, con filtri su Tipo Conto e Controparte.")]
        public static object SiClearingBuyInSaldi(
            [ExcelArgument(Name = "alertDate",    Description = "Data Buy-In Alert (seriale Excel o dd/mm/yyyy)")] object alertDate,
            [ExcelArgument(Name = "saldiLive",    Description = "Range 2 colonne {ISIN, qty} con saldi live (opzionale)")] object? saldiLive = null,
            [ExcelArgument(Name = "tipoConto",    Description = "Filtro Tipo Conto: stringa o range verticale (opzionale)")] object? tipoConto = null,
            [ExcelArgument(Name = "controparte",  Description = "Filtro Descrizione Controparte: stringa o range verticale (opzionale)")] object? controparte = null)
        {
            var data = AddIn.Cache.GetOrLoad(AddIn.Settings.SaveFolder);
            if (data == null || data.GetLength(0) == 0)
                return ExcelError.ExcelErrorNA;

            if (!TryParseDate(alertDate, out DateTime targetDate))
                return ExcelError.ExcelErrorValue;

            // resolve required columns
            int isinCol   = ResolveFirst(data, "ISIN Code", "ISIN");
            int buyInCol  = ResolveFirst(data, "Buy-In Alert");
            int segnoCol  = ResolveFirst(data, "Segno");
            int qtaCol    = ResolveFirst(data, "Quantita'");
            int tipoCol   = ResolveFirst(data, "Tipo Conto");
            int contrCol  = ResolveFirst(data, "Controparte");   // partial match

            if (isinCol < 0 || buyInCol < 0 || segnoCol < 0 || qtaCol < 0)
            {
                Logger.Log("SiClearingBuyInSaldi: colonne obbligatorie non trovate nel CSV.");
                return ExcelError.ExcelErrorValue;
            }

            // build filter sets
            var tipoFilter  = BuildFilterSet(tipoConto);
            var contrFilter = BuildFilterSet(controparte);

            // step 1 — collect ISINs with matching Buy-In Alert date
            var targetIsins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int rowCount = data.GetLength(0);

            for (int r = 1; r < rowCount; r++)
            {
                string raw = data[r, buyInCol].Trim();
                if (raw == "--" || string.IsNullOrEmpty(raw)) continue;
                if (!TryParseDateString(raw, out DateTime dt)) continue;
                if (dt.Date != targetDate) continue;
                targetIsins.Add(data[r, isinCol].Trim());
            }

            if (targetIsins.Count == 0)
            {
                Logger.Log($"SiClearingBuyInSaldi: nessun ISIN con Buy-In Alert {targetDate:dd/MM/yyyy}.");
                return ExcelError.ExcelErrorNA;
            }

            // step 2 — filter rows + groupby ISIN (sum signed qty)
            var saldoCsv = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            for (int r = 1; r < rowCount; r++)
            {
                string isin = data[r, isinCol].Trim();
                if (!targetIsins.Contains(isin)) continue;

                if (tipoFilter != null && tipoCol >= 0 &&
                    !tipoFilter.Contains(data[r, tipoCol].Trim()))
                    continue;

                if (contrFilter != null && contrCol >= 0 &&
                    !contrFilter.Contains(data[r, contrCol].Trim()))
                    continue;

                string qtaRaw = data[r, qtaCol].Trim();
                if (string.IsNullOrEmpty(qtaRaw)) continue;

                double qty;
                try { qty = CsvCache.ParseItalianNumber(qtaRaw); }
                catch { continue; }

                string segno = data[r, segnoCol].Trim().ToUpperInvariant();
                double signed = segno == "A" ? qty : -qty;

                if (!saldoCsv.TryGetValue(isin, out double cur))
                    saldoCsv[isin] = signed;
                else
                    saldoCsv[isin] = cur + signed;
            }

            if (saldoCsv.Count == 0)
                return ExcelError.ExcelErrorNA;

            // step 3 — parse live balances if provided
            var livemap = ParseLiveSaldi(saldiLive);

            bool hasLive = livemap != null;

            // build output
            int cols = hasLive ? 4 : 2;
            var result = new object[saldoCsv.Count + 1, cols];
            result[0, 0] = "ISIN";
            result[0, 1] = "Saldo CSV";
            if (hasLive) { result[0, 2] = "Saldo Live"; result[0, 3] = "Saldo Totale"; }

            int idx = 1;
            foreach (var kv in saldoCsv)
            {
                double live = (hasLive && livemap!.TryGetValue(kv.Key, out double lv)) ? lv : 0;
                result[idx, 0] = kv.Key;
                result[idx, 1] = kv.Value;
                if (hasLive) { result[idx, 2] = live; result[idx, 3] = kv.Value + live; }
                idx++;
            }

            Logger.Log($"SiClearingBuyInSaldi {targetDate:dd/MM/yyyy}: {saldoCsv.Count} ISIN.");
            return result;
        }

        // ── helpers ──────────────────────────────────────────────────────────

        private static int ResolveFirst(string[,] data, params string[] names)
        {
            foreach (var n in names)
            {
                int idx = Resolver.Resolve(data, n);
                if (idx >= 0) return idx;
            }
            return -1;
        }

        /// <summary>
        /// Builds a case-insensitive set of allowed values from a scalar or Excel range.
        /// Returns null if the argument is missing/empty (= no filter).
        /// </summary>
        private static HashSet<string>? BuildFilterSet(object? arg)
        {
            if (arg == null || arg is ExcelMissing || arg is ExcelEmpty)
                return null;

            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (arg is object[,] grid)
            {
                int rows = grid.GetLength(0), cols = grid.GetLength(1);
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                    {
                        var v = grid[r, c]?.ToString()?.Trim();
                        if (!string.IsNullOrEmpty(v)) set.Add(v!);
                    }
            }
            else if (arg is object[] arr)
            {
                foreach (var item in arr)
                {
                    var v = item?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(v)) set.Add(v!);
                }
            }
            else
            {
                var v = arg.ToString()?.Trim();
                if (!string.IsNullOrEmpty(v)) set.Add(v!);
            }

            return set.Count > 0 ? set : null;
        }

        /// <summary>
        /// Parses a 2-column range {ISIN, qty} into a dictionary.
        /// Returns null if arg is missing.
        /// </summary>
        private static Dictionary<string, double>? ParseLiveSaldi(object? arg)
        {
            if (arg == null || arg is ExcelMissing || arg is ExcelEmpty)
                return null;

            if (!(arg is object[,] grid)) return null;

            int rows = grid.GetLength(0), cols = grid.GetLength(1);
            if (cols < 2) return null;

            var map = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            for (int r = 0; r < rows; r++)
            {
                string? isin = grid[r, 0]?.ToString()?.Trim();
                if (string.IsNullOrEmpty(isin)) continue;
                if (grid[r, 1] is double qty)
                    map[isin!] = map.TryGetValue(isin!, out double cur) ? cur + qty : qty;
            }
            return map.Count > 0 ? map : null;
        }

        private static bool TryParseDate(object arg, out DateTime result)
        {
            if (arg is double serial)
            {
                result = DateTime.FromOADate(serial).Date;
                return true;
            }
            if (arg is string s)
                return TryParseDateString(s, out result);

            result = default;
            return false;
        }

        private static bool TryParseDateString(string s, out DateTime result) =>
            DateTime.TryParseExact(s.Trim(), "dd/MM/yyyy",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out result);

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
