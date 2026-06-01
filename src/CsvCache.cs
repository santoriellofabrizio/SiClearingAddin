using System;
using System.Globalization;
using System.IO;
using System.Threading;

namespace SiClearing
{
    public class CsvCache
    {
        private readonly ReaderWriterLockSlim _lock = new ReaderWriterLockSlim();
        private string[,]? _data;

        public bool IsLoaded
        {
            get
            {
                _lock.EnterReadLock();
                try { return _data != null; }
                finally { _lock.ExitReadLock(); }
            }
        }

        public string[,]? Data
        {
            get
            {
                _lock.EnterReadLock();
                try { return _data; }
                finally { _lock.ExitReadLock(); }
            }
        }

        public void Invalidate()
        {
            _lock.EnterWriteLock();
            try { _data = null; }
            finally { _lock.ExitWriteLock(); }
        }

        public void LoadFile(string path)
        {
            if (!File.Exists(path)) return;
            var loaded = ParseCsv(path);
            _lock.EnterWriteLock();
            try { _data = loaded; }
            finally { _lock.ExitWriteLock(); }
        }

        public void LoadLatest(string saveFolder)
        {
            if (!Directory.Exists(saveFolder))
                return;

            var files = Directory.GetFiles(saveFolder, "*.csv");
            if (files.Length == 0)
                return;

            // pick most recently modified
            string latest = files[0];
            DateTime latestTime = File.GetLastWriteTime(latest);
            foreach (var f in files)
            {
                var t = File.GetLastWriteTime(f);
                if (t > latestTime) { latestTime = t; latest = f; }
            }

            var loaded = ParseCsv(latest);

            _lock.EnterWriteLock();
            try { _data = loaded; }
            finally { _lock.ExitWriteLock(); }
        }

        public string[,]? GetOrLoad(string saveFolder)
        {
            _lock.EnterReadLock();
            try
            {
                if (_data != null) return _data;
            }
            finally { _lock.ExitReadLock(); }

            LoadLatest(saveFolder);
            return Data;
        }

        private static string[,] ParseCsv(string path)
        {
            var lines = File.ReadAllLines(path, System.Text.Encoding.UTF8);
            if (lines.Length == 0)
                return new string[0, 0];

            // detect separator
            char sep = lines[0].Contains(";") ? ';' : ',';

            var rows = new System.Collections.Generic.List<string[]>();
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                rows.Add(SplitLine(line, sep));
            }

            if (rows.Count == 0) return new string[0, 0];

            int colCount = rows[0].Length;
            var data = new string[rows.Count, colCount];
            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < colCount; c++)
                    data[r, c] = c < rows[r].Length ? rows[r][c].Trim('"').Trim() : "";

            return data;
        }

        private static string[] SplitLine(string line, char sep)
        {
            // simple split — no quoted-field multiline support needed for this CSV
            return line.Split(sep);
        }

        public static double ParseItalianNumber(string s)
        {
            s = s.Trim().Replace(" ", "");
            if (s.Contains(","))
                return double.Parse(s.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);

            int dotIdx = s.LastIndexOf('.');
            if (dotIdx >= 0 && s.Length - dotIdx - 1 == 3)
                return double.Parse(s.Replace(".", ""), CultureInfo.InvariantCulture);

            return double.Parse(s, CultureInfo.InvariantCulture);
        }
    }
}
