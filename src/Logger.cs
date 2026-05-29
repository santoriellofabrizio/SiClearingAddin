using System;
using System.Collections.Generic;

namespace SiClearing
{
    public static class Logger
    {
        private static readonly object _lock = new object();
        private static readonly List<string> _entries = new List<string>();

        public static event Action<string>? EntryAdded;

        public static IReadOnlyList<string> Entries
        {
            get { lock (_lock) { return _entries.ToArray(); } }
        }

        public static void Log(string message)
        {
            string entry = $"[{DateTime.Now:HH:mm:ss}] {message}";
            lock (_lock) { _entries.Add(entry); }
            try { EntryAdded?.Invoke(entry); } catch { }
        }

        public static void Clear()
        {
            lock (_lock) { _entries.Clear(); }
        }
    }
}
