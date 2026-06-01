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
            [ExcelArgument(Name = "isin",      Description = "Codice ISIN da cercare")] string isin,
            [ExcelArgument(Name = "cols",      Description = "Colonne da restituire: nomi o indici 1-based (opzionale)")] object? cols = null,
            [ExcelArgument(Name = "tipoConto", Description = "Filtro Tipo Conto: stringa o range (opzionale)")] object? tipoConto = null,
            [ExcelArgument(Name = "mercato",   Description = "Filtro Mercato: stringa o range (opzionale)")] object? mercato = null)
        {
            var data = AddIn.Cache.GetOrLoad(AddIn.Settings.SaveFolder);
            if (data == null || data.GetLength(0) == 0)
                return ExcelError.ExcelErrorNA;

            int isinCol = Resolver.Resolve(data, "ISIN Code");
            if (isinCol < 0) isinCol = Resolver.Resolve(data, "ISIN");
            if (isinCol < 0) return ExcelError.ExcelErrorValue;

            int[] colIdxs = Resolver.ResolveFromOptional(data, cols, AddIn.Settings.DefaultColumns);
            var gf = BuildGlobalFilter(data);
            var tipoFilter = BuildFilterSet(tipoConto);
            var mercatoFilter = BuildFilterSet(mercato);
            int tipoCol = Resolver.Resolve(data, "Tipo Conto");
            int mercatoCol = Resolver.Resolve(data, "Mercato");

            var rows = new List<int>();
            int rowCount = data.GetLength(0);
            for (int r = 1; r < rowCount; r++)
            {
                if (!data[r, isinCol].Equals(isin.Trim(), StringComparison.OrdinalIgnoreCase)) continue;
                if (!gf.Pass(data, r)) continue;
                if (tipoFilter != null && tipoCol >= 0 && !tipoFilter.Contains(data[r, tipoCol].Trim())) continue;
                if (mercatoFilter != null && mercatoCol >= 0 && !mercatoFilter.Contains(data[r, mercatoCol].Trim())) continue;
                rows.Add(r);
            }

            if (rows.Count == 0) return ExcelError.ExcelErrorNA;
            return BuildResult(data, rows, colIdxs);
        }

        [ExcelFunction(Name = "SiClearingBuyInAlert",
            Description = "Restituisce ISIN con Buy-In Alert Date uguale alla data specificata.")]
        public static object SiClearingBuyInAlert(
            [ExcelArgument(Name = "alertDate", Description = "Data alert (numero seriale Excel o dd/mm/yyyy)")] object alertDate,
            [ExcelArgument(Name = "cols",      Description = "Colonne da restituire (opzionale)")] object? cols = null,
            [ExcelArgument(Name = "isin",      Description = "Filtro ISIN: stringa o range (opzionale)")] object? isin = null,
            [ExcelArgument(Name = "tipoConto", Description = "Filtro Tipo Conto: stringa o range (opzionale)")] object? tipoConto = null,
            [ExcelArgument(Name = "mercato",   Description = "Filtro Mercato: stringa o range (opzionale)")] object? mercato = null)
        {
            var data = AddIn.Cache.GetOrLoad(AddIn.Settings.SaveFolder);
            if (data == null || data.GetLength(0) == 0)
                return ExcelError.ExcelErrorNA;

            if (!TryParseDate(alertDate, out DateTime targetDate))
                return ExcelError.ExcelErrorValue;

            int buyInCol = Resolver.Resolve(data, "Buy-In Alert");
            if (buyInCol < 0) return ExcelError.ExcelErrorValue;

            int isinCol = Resolver.Resolve(data, "ISIN Code");
            if (isinCol < 0) isinCol = Resolver.Resolve(data, "ISIN");
            if (isinCol < 0) return ExcelError.ExcelErrorValue;

            int[] colIdxs = Resolver.ResolveFromOptional(data, cols, AddIn.Settings.DefaultColumns);
            var gf = BuildGlobalFilter(data);
            var isinFilter  = BuildFilterSet(isin);
            var tipoFilter  = BuildFilterSet(tipoConto);
            var mercatoFilter = BuildFilterSet(mercato);
            int tipoCol    = Resolver.Resolve(data, "Tipo Conto");
            int mercatoCol = Resolver.Resolve(data, "Mercato");

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

                string isinVal = data[r, isinCol].Trim();
                if (isinFilter != null && !isinFilter.Contains(isinVal)) continue;
                if (tipoFilter != null && tipoCol >= 0 && !tipoFilter.Contains(data[r, tipoCol].Trim())) continue;
                if (mercatoFilter != null && mercatoCol >= 0 && !mercatoFilter.Contains(data[r, mercatoCol].Trim())) continue;

                if (!seenIsin.Add(isinVal)) continue;
                rows.Add(r);
            }

            if (rows.Count == 0) return ExcelError.ExcelErrorNA;
            return BuildResult(data, rows, colIdxs);
        }

        [ExcelFunction(Name = "SiClearingSumUp", IsVolatile = true,
            Description = "Greedy settlement: quanti scoperti si chiudono e quante shares mancano.")]
        public static object SiClearingSumUp(
            [ExcelArgument(Name = "alertDate", Description = "Data buy-in alert (seriale Excel o dd/MM/yyyy)")] object alertDate,
            [ExcelArgument(Name = "markets",   Description = "Market ID o range verticale per Duma")] object markets,
            [ExcelArgument(Name = "tipoConto", Description = "Filtro Tipo Conto (opzionale)")] object? tipoConto = null,
            [ExcelArgument(Name = "mercato",   Description = "Filtro colonna Mercato CSV (opzionale)")] object? mercato = null,
            [ExcelArgument(Name = "isin",      Description = "Filtro ISIN (opzionale)")] object? isin = null)
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
            int contrCol = ResolveFirst(data, "Descrizione Controparte", "Controparte");
            int tipoCol  = ResolveFirst(data, "Tipo Conto");
            int mercatoCol = ResolveFirst(data, "Mercato");

            if (isinCol < 0 || buyInCol < 0 || segnoCol < 0 || qtaCol < 0)
            {
                Logger.Log("SiClearingSumUp: colonne obbligatorie non trovate nel CSV.");
                return ExcelError.ExcelErrorValue;
            }

            var gf = BuildGlobalFilter(data);
            var tipoFilter    = BuildFilterSet(tipoConto);
            var mercatoFilter = BuildFilterSet(mercato);
            var isinFilter    = BuildFilterSet(isin);

            // collect target ISINs from buy-in alert date
            var targetIsins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int rowCount = data.GetLength(0);
            for (int r = 1; r < rowCount; r++)
            {
                if (!gf.Pass(data, r)) continue;
                string raw = data[r, buyInCol].Trim();
                if (raw == "--" || string.IsNullOrEmpty(raw)) continue;
                if (!TryParseDateString(raw, out DateTime dt)) continue;
                if (dt.Date != targetDate) continue;
                string isinVal = data[r, isinCol].Trim();
                if (isinFilter != null && !isinFilter.Contains(isinVal)) continue;
                targetIsins.Add(isinVal);
            }

            if (targetIsins.Count == 0)
            {
                Logger.Log($"SiClearingSumUp: nessun ISIN con Buy-In Alert {targetDate:dd/MM/yyyy}.");
                return ExcelError.ExcelErrorNA;
            }

            // per ISIN: list of individual scoperto qty (Segno=D), split by CCG vs other
            var scopertiCcg   = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);
            var scopertiOther = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);
            // aggregated incoming (Segno=A)
            var incomingCcg = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            var incomingAll = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            foreach (var i in targetIsins)
            {
                scopertiCcg[i]   = new List<double>();
                scopertiOther[i] = new List<double>();
            }

            for (int r = 1; r < rowCount; r++)
            {
                if (!gf.Pass(data, r)) continue;
                string isinVal = data[r, isinCol].Trim();
                if (!targetIsins.Contains(isinVal)) continue;

                if (tipoFilter != null && tipoCol >= 0 && !tipoFilter.Contains(data[r, tipoCol].Trim())) continue;
                if (mercatoFilter != null && mercatoCol >= 0 && !mercatoFilter.Contains(data[r, mercatoCol].Trim())) continue;

                string qtaRaw = data[r, qtaCol].Trim();
                if (string.IsNullOrEmpty(qtaRaw)) continue;
                double qty;
                try { qty = CsvCache.ParseItalianNumber(qtaRaw); }
                catch { continue; }

                string segno = data[r, segnoCol].Trim().ToUpperInvariant();
                bool isCcg = contrCol >= 0 && data[r, contrCol].IndexOf("CCG", StringComparison.OrdinalIgnoreCase) >= 0;

                if (segno == "D")
                {
                    if (isCcg) scopertiCcg[isinVal].Add(qty);
                    else        scopertiOther[isinVal].Add(qty);
                }
                else if (segno == "A")
                {
                    if (!incomingAll.ContainsKey(isinVal)) incomingAll[isinVal] = 0;
                    incomingAll[isinVal] += qty;
                    if (isCcg)
                    {
                        if (!incomingCcg.ContainsKey(isinVal)) incomingCcg[isinVal] = 0;
                        incomingCcg[isinVal] += qty;
                    }
                }
            }

            // get Duma live net per ISIN
            var dumaNet = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            var xl = ExcelDnaUtil.Application as Excel.Application;
            if (xl != null)
            {
                var marketList = ExtractStringList(markets);
                foreach (var market in marketList)
                {
                    object raw;
                    try { raw = xl.Evaluate($"dumaGetTableRecords(\"Trade\",\"{market}\",TRUE)"); }
                    catch (Exception ex)
                    {
                        Logger.Log($"[SumUp] Duma non disponibile per mercato {market} — assicurarsi che Duma sia avviato e connesso. ({ex.Message})");
                        continue;
                    }
                    if (!(raw is object[,] table))
                    {
                        Logger.Log($"[SumUp] Duma non disponibile per mercato {market} — assicurarsi che Duma sia avviato e connesso.");
                        continue;
                    }

                    int dr0 = table.GetLowerBound(0), dc0 = table.GetLowerBound(1);
                    int dRows = table.GetUpperBound(0) - dr0 + 1;
                    int dCols = table.GetUpperBound(1) - dc0 + 1;
                    if (dRows < 2) continue;

                    int colDIsin = -1, colDQty = -1, colDSide = -1;
                    for (int c = 0; c < dCols; c++)
                    {
                        string hdr = table[dr0, dc0 + c]?.ToString() ?? "";
                        if (hdr.Equals("instrument.isincode", StringComparison.OrdinalIgnoreCase)) colDIsin = c;
                        else if (hdr.Equals("tradeqty", StringComparison.OrdinalIgnoreCase)) colDQty = c;
                        else if (hdr.Equals("side", StringComparison.OrdinalIgnoreCase)) colDSide = c;
                    }
                    if (colDIsin < 0 || colDQty < 0 || colDSide < 0) continue;

                    for (int r = 1; r < dRows; r++)
                    {
                        string isinVal = table[dr0 + r, dc0 + colDIsin]?.ToString()?.Trim() ?? "";
                        if (!targetIsins.Contains(isinVal)) continue;
                        string sideVal = table[dr0 + r, dc0 + colDSide]?.ToString()?.Trim() ?? "";
                        double qty;
                        var qRaw = table[dr0 + r, dc0 + colDQty];
                        if (qRaw is double qd) qty = qd;
                        else if (!double.TryParse(qRaw?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out qty)) continue;
                        double signed = sideVal.Equals("Buy", StringComparison.OrdinalIgnoreCase) ? qty : -qty;
                        dumaNet[isinVal] = dumaNet.TryGetValue(isinVal, out double cur) ? cur + signed : signed;
                    }
                }
            }
            else
            {
                Logger.Log("[SumUp] Impossibile accedere a Excel — Duma non interrogato.");
            }

            // output: ISIN | Residuo CCG | Residuo Tutte CP
            var isinList = new List<string>(targetIsins);
            isinList.Sort(StringComparer.OrdinalIgnoreCase);
            var result = new object[isinList.Count + 1, 3];
            result[0, 0] = "ISIN";
            result[0, 1] = "Residuo CCG";
            result[0, 2] = "Residuo Tutte CP";

            for (int i = 0; i < isinList.Count; i++)
            {
                string key = isinList[i];
                var ccgList   = scopertiCcg[key];
                var otherList = scopertiOther[key];
                ccgList.Sort();
                otherList.Sort();

                double dNet = dumaNet.TryGetValue(key, out double dn) ? dn : 0.0;
                incomingCcg.TryGetValue(key, out double incCcg);
                incomingAll.TryGetValue(key, out double incAll);

                GreedyClose(ccgList, otherList, incCcg + Math.Max(0, dNet),
                    out _, out _, out double resCcg);
                GreedyClose(ccgList, otherList, incAll + Math.Max(0, dNet),
                    out _, out _, out double resAll);

                result[i + 1, 0] = key;
                result[i + 1, 1] = resCcg;
                result[i + 1, 2] = resAll;
            }

            Logger.Log($"[SumUp] {targetDate:dd/MM/yyyy}: {isinList.Count} ISIN elaborati.");
            return result;
        }

        private static void GreedyClose(List<double> ccg, List<double> other, double available,
            out int closedN, out double closedQty, out double residuo)
        {
            closedN = 0; closedQty = 0;
            foreach (var s in ccg)
                if (available >= s) { available -= s; closedN++; closedQty += s; }
            foreach (var s in other)
                if (available >= s) { available -= s; closedN++; closedQty += s; }
            double total = 0;
            foreach (var s in ccg)   total += s;
            foreach (var s in other) total += s;
            residuo = total - closedQty;
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
