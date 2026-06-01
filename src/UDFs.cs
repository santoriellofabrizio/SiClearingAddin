using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace SiClearing
{
    public static class UDFs
    {
        private static readonly ColumnResolver Resolver = new ColumnResolver();

        // ── global filter state (resolved once per UDF call) ─────────────────

        private struct GlobalFilter
        {
            public int MercatoCol;
            public int TipoContoCol;
            public HashSet<string>? Mercati;
            public HashSet<string>? TipiConto;

            public bool Pass(string[,] data, int row)
            {
                if (Mercati != null && MercatoCol >= 0 &&
                    !Mercati.Contains(data[row, MercatoCol].Trim()))
                    return false;
                if (TipiConto != null && TipoContoCol >= 0 &&
                    !TipiConto.Contains(data[row, TipoContoCol].Trim()))
                    return false;
                return true;
            }
        }

        private static GlobalFilter BuildGlobalFilter(string[,] data)
        {
            var f = new GlobalFilter
            {
                MercatoCol   = Resolver.Resolve(data, "Mercato"),
                TipoContoCol = Resolver.Resolve(data, "Tipo Conto"),
                Mercati      = ParseFilterSetting(AddIn.Settings.MercatoFilter),
                TipiConto    = ParseFilterSetting(AddIn.Settings.TipoContoFilter),
            };
            return f;
        }

        private static HashSet<string>? ParseFilterSetting(string setting)
        {
            if (string.IsNullOrWhiteSpace(setting)) return null;
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var v in setting.Split(';'))
            {
                var t = v.Trim();
                if (!string.IsNullOrEmpty(t)) set.Add(t);
            }
            return set.Count > 0 ? set : null;
        }

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

            int[] colIdxs = Resolver.ResolveList(data, cols, AddIn.Settings.DefaultColumns);
            var gf = BuildGlobalFilter(data);

            var rows = new List<int>();
            int rowCount = data.GetLength(0);
            for (int r = 1; r < rowCount; r++)
                if (data[r, isinCol].Equals(isin.Trim(), StringComparison.OrdinalIgnoreCase) && gf.Pass(data, r))
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

            int[] colIdxs = Resolver.ResolveFromOptional(data, cols, AddIn.Settings.DefaultColumns);
            var gf = BuildGlobalFilter(data);

            var seenIsin = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var rows = new List<int>();
            int rowCount = data.GetLength(0);

            for (int r = 1; r < rowCount; r++)
            {
                if (!gf.Pass(data, r)) continue;
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
            [ExcelArgument(Name = "alertDate",   Description = "Data Buy-In Alert (seriale Excel o dd/mm/yyyy)")] object alertDate,
            [ExcelArgument(Name = "saldiLive",   Description = "Range 2 colonne {ISIN, qty} con saldi live (opzionale)")] object? saldiLive = null,
            [ExcelArgument(Name = "tipoConto",   Description = "Filtro Tipo Conto: stringa o range verticale (opzionale)")] object? tipoConto = null,
            [ExcelArgument(Name = "controparte", Description = "Filtro Descrizione Controparte: stringa o range verticale (opzionale)")] object? controparte = null)
        {
            var data = AddIn.Cache.GetOrLoad(AddIn.Settings.SaveFolder);
            if (data == null || data.GetLength(0) == 0)
                return ExcelError.ExcelErrorNA;

            if (!TryParseDate(alertDate, out DateTime targetDate))
                return ExcelError.ExcelErrorValue;

            int isinCol  = ResolveFirst(data, "ISIN Code", "ISIN");
            int buyInCol = ResolveFirst(data, "Buy-In Alert");
            int segnoCol = ResolveFirst(data, "Segno");
            int qtaCol   = ResolveFirst(data, "Quantita'");
            int tipoCol  = ResolveFirst(data, "Tipo Conto");
            int contrCol = ResolveFirst(data, "Controparte");

            if (isinCol < 0 || buyInCol < 0 || segnoCol < 0 || qtaCol < 0)
            {
                Logger.Log("SiClearingBuyInSaldi: colonne obbligatorie non trovate nel CSV.");
                return ExcelError.ExcelErrorValue;
            }

            var tipoFilter  = BuildFilterSet(tipoConto);
            var contrFilter = BuildFilterSet(controparte);
            var gf = BuildGlobalFilter(data);

            var targetIsins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int rowCount = data.GetLength(0);

            for (int r = 1; r < rowCount; r++)
            {
                if (!gf.Pass(data, r)) continue;
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

            var saldoCsv = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            for (int r = 1; r < rowCount; r++)
            {
                if (!gf.Pass(data, r)) continue;
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

                saldoCsv[isin] = saldoCsv.TryGetValue(isin, out double cur) ? cur + signed : signed;
            }

            if (saldoCsv.Count == 0)
                return ExcelError.ExcelErrorNA;

            var livemap = ParseLiveSaldi(saldiLive);
            bool hasLive = livemap != null;

            int outCols = hasLive ? 4 : 2;
            var result = new object[saldoCsv.Count + 1, outCols];
            result[0, 0] = "ISIN";
            result[0, 1] = "Saldo CSV";
            if (hasLive) { result[0, 2] = "Saldo Live"; result[0, 3] = "Saldo Totale"; }

            int idx = 1;
            foreach (var kv in saldoCsv)
            {
                double lv = (hasLive && livemap!.TryGetValue(kv.Key, out double lvv)) ? lvv : 0;
                result[idx, 0] = kv.Key;
                result[idx, 1] = kv.Value;
                if (hasLive) { result[idx, 2] = lv; result[idx, 3] = kv.Value + lv; }
                idx++;
            }

            Logger.Log($"SiClearingBuyInSaldi {targetDate:dd/MM/yyyy}: {saldoCsv.Count} ISIN.");
            return result;
        }

        [ExcelFunction(Name = "SiClearingGetTodayBalance", IsVolatile = true,
            Description = "Saldo live per ISIN da dumaGetTableRecords su uno o più mercati (Buy-Sell qty).")]
        public static object SiClearingGetTodayBalance(
            [ExcelArgument(Name = "markets", Description = "Stringa o range verticale di market ID")] object markets,
            [ExcelArgument(Name = "isin",    Description = "ISIN specifico → scalare; omesso → array ISIN/Saldo")] object? isin = null)
        {
            var marketList = ExtractStringList(markets);
            if (marketList.Count == 0)
                return ExcelError.ExcelErrorValue;

            var balance = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            var xl = ExcelDnaUtil.Application as Excel.Application;
            if (xl == null) return ExcelError.ExcelErrorValue;

            foreach (var market in marketList)
            {
                object raw;
                try
                {
                    raw = xl.Evaluate($"dumaGetTableRecords(\"Trade\",\"{market}\",TRUE)");
                }
                catch (Exception ex)
                {
                    Logger.Log($"[GetTodayBalance] Duma non disponibile per mercato {market} — assicurarsi che Duma sia avviato e connesso. ({ex.Message})");
                    continue;
                }

                if (!(raw is object[,] table))
                {
                    Logger.Log($"[GetTodayBalance] Duma non disponibile per mercato {market} — assicurarsi che Duma sia avviato e connesso.");
                    continue;
                }

                // COM arrays from Application.Evaluate can be 1-based — use GetLowerBound/GetUpperBound
                int r0 = table.GetLowerBound(0), c0 = table.GetLowerBound(1);
                int rowCount = table.GetUpperBound(0) - r0 + 1;
                int colCount = table.GetUpperBound(1) - c0 + 1;

                if (rowCount < 2)
                {
                    Logger.Log($"[GetTodayBalance] Mercato {market}: nessun trade ricevuto.");
                    continue;
                }

                // find column indices from header row
                int colIsin = -1, colQty = -1, colSide = -1;
                for (int c = 0; c < colCount; c++)
                {
                    string hdr = table[r0, c0 + c]?.ToString() ?? "";
                    if (hdr.Equals("instrument.isincode", StringComparison.OrdinalIgnoreCase)) colIsin = c;
                    else if (hdr.Equals("tradeqty", StringComparison.OrdinalIgnoreCase)) colQty = c;
                    else if (hdr.Equals("side", StringComparison.OrdinalIgnoreCase)) colSide = c;
                }

                if (colIsin < 0 || colQty < 0 || colSide < 0)
                {
                    Logger.Log($"[GetTodayBalance] Mercato {market}: colonne instrument.isincode/tradeqty/side non trovate nell'header Duma.");
                    continue;
                }

                int tradeCount = 0;
                for (int r = 1; r < rowCount; r++)
                {
                    string isinVal = table[r0 + r, c0 + colIsin]?.ToString()?.Trim() ?? "";
                    string sideVal = table[r0 + r, c0 + colSide]?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(isinVal)) continue;

                    double qty;
                    var qtyRaw = table[r0 + r, c0 + colQty];
                    if (qtyRaw is double qd) qty = qd;
                    else if (!double.TryParse(qtyRaw?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out qty)) continue;

                    double signed = sideVal.Equals("Buy", StringComparison.OrdinalIgnoreCase) ? qty : -qty;
                    balance[isinVal] = balance.TryGetValue(isinVal, out double cur) ? cur + signed : signed;
                    tradeCount++;
                }

                Logger.Log($"[GetTodayBalance] Mercato {market}: {tradeCount} trade ricevuti.");
            }

            if (balance.Count == 0)
                return ExcelError.ExcelErrorNA;

            Logger.Log($"[GetTodayBalance] Totale ISIN: {balance.Count}");

            // single ISIN → scalar
            string? isinFilter = null;
            if (isin != null && !(isin is ExcelMissing) && !(isin is ExcelEmpty))
                isinFilter = isin.ToString()?.Trim();

            if (!string.IsNullOrEmpty(isinFilter))
                return balance.TryGetValue(isinFilter!, out double v) ? (object)v : 0.0;

            // all ISINs → array
            var result = new object[balance.Count + 1, 2];
            result[0, 0] = "ISIN";
            result[0, 1] = "Saldo";
            int idx = 1;
            foreach (var kv in balance)
            {
                result[idx, 0] = kv.Key;
                result[idx, 1] = kv.Value;
                idx++;
            }
            return result;
        }

        // ── helpers ──────────────────────────────────────────────────────────

        private static List<string> ExtractStringList(object arg)
        {
            var list = new List<string>();
            if (arg == null || arg is ExcelMissing || arg is ExcelEmpty) return list;

            if (arg is object[,] grid)
            {
                int rows = grid.GetLength(0), cols = grid.GetLength(1);
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                    {
                        var v = grid[r, c]?.ToString()?.Trim();
                        if (!string.IsNullOrEmpty(v)) list.Add(v!);
                    }
            }
            else if (arg is object[] arr)
            {
                foreach (var item in arr)
                {
                    var v = item?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(v)) list.Add(v!);
                }
            }
            else
            {
                var v = arg.ToString()?.Trim();
                if (!string.IsNullOrEmpty(v)) list.Add(v!);
            }

            return list;
        }

        private static int ResolveFirst(string[,] data, params string[] names)
        {
            foreach (var n in names)
            {
                int idx = Resolver.Resolve(data, n);
                if (idx >= 0) return idx;
            }
            return -1;
        }

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

        private static Dictionary<string, double>? ParseLiveSaldi(object? arg)
        {
            if (arg == null || arg is ExcelMissing || arg is ExcelEmpty) return null;
            if (!(arg is object[,] grid)) return null;

            int rows = grid.GetLength(0), cols = grid.GetLength(1);
            if (cols < 2) return null;

            var map = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            for (int r = 0; r < rows; r++)
            {
                string? isinV = grid[r, 0]?.ToString()?.Trim();
                if (string.IsNullOrEmpty(isinV)) continue;
                if (grid[r, 1] is double qty)
                    map[isinV!] = map.TryGetValue(isinV!, out double cur) ? cur + qty : qty;
            }
            return map.Count > 0 ? map : null;
        }

        private static bool TryParseDate(object arg, out DateTime result)
        {
            if (arg is double serial) { result = DateTime.FromOADate(serial).Date; return true; }
            if (arg is string s) return TryParseDateString(s, out result);
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
