// MD Doc Link - Excel key <-> model item name matching (no Navisworks refs, unit tested)
// Author: MD (Mayank Dholariya)
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DocLink.Core
{
    public enum MatchMode
    {
        /// <summary>Whole name equals the key.</summary>
        Exact = 0,
        /// <summary>Name contains the key as a whole part: the characters either side are not letters/digits.</summary>
        ContainsWhole = 1,
        /// <summary>Name contains the key anywhere.</summary>
        Contains = 2,
        /// <summary>Name starts with the key.</summary>
        StartsWith = 3
    }

    public sealed class MatchOptions
    {
        public MatchMode Mode = MatchMode.Exact;
        /// <summary>Drop a leading '/' (E3D / PDMS names).</summary>
        public bool StripLeadingSlash = true;
        /// <summary>Drop a leading type word before the name, e.g. "PIPE /150-P-1001-..." -> "/150-P-1001-...".</summary>
        public bool StripTypePrefix = true;
        /// <summary>Compare letters and digits only ("150-P-1001" == "150 P 1001" == "150P1001").</summary>
        public bool IgnoreSeparators = false;

        public string Describe()
        {
            return Mode + (StripLeadingSlash ? ", strip '/'" : "") + (StripTypePrefix ? ", strip type word" : "")
                + (IgnoreSeparators ? ", ignore separators" : "");
        }
    }

    public static class KeyNormalizer
    {
        private static readonly Regex TypePrefix = new Regex(@"^[A-Za-z]{2,12}\s+(?=/)", RegexOptions.CultureInvariant);

        /// <summary>Case-insensitive, trimmed, single-spaced form used for comparing keys and names.</summary>
        public static string Normalize(string raw, MatchOptions o)
        {
            if (raw == null) return "";
            string s = Names.Clean(raw);
            if (s.Length == 0) return "";
            if (o == null) o = new MatchOptions();
            if (o.StripTypePrefix) s = TypePrefix.Replace(s, "");
            if (o.StripLeadingSlash) s = s.TrimStart('/').Trim();
            s = s.ToUpperInvariant();
            if (o.IgnoreSeparators)
            {
                var sb = new StringBuilder(s.Length);
                foreach (char c in s) if (char.IsLetterOrDigit(c)) sb.Append(c);
                s = sb.ToString();
            }
            return s;
        }

        /// <summary>True when an already-normalized name matches an already-normalized key.</summary>
        public static bool IsMatch(string normName, string normKey, MatchMode mode)
        {
            if (string.IsNullOrEmpty(normName) || string.IsNullOrEmpty(normKey)) return false;
            switch (mode)
            {
                case MatchMode.Exact: return string.Equals(normName, normKey, StringComparison.Ordinal);
                case MatchMode.StartsWith: return normName.StartsWith(normKey, StringComparison.Ordinal);
                case MatchMode.Contains: return normName.IndexOf(normKey, StringComparison.Ordinal) >= 0;
                case MatchMode.ContainsWhole:
                    int from = 0;
                    while (from <= normName.Length - normKey.Length)
                    {
                        int i = normName.IndexOf(normKey, from, StringComparison.Ordinal);
                        if (i < 0) return false;
                        int end = i + normKey.Length;
                        bool leftOk = i == 0 || !char.IsLetterOrDigit(normName[i - 1]) || !char.IsLetterOrDigit(normKey[0]);
                        bool rightOk = end == normName.Length || !char.IsLetterOrDigit(normName[end]) || !char.IsLetterOrDigit(normKey[normKey.Length - 1]);
                        if (leftOk && rightOk) return true;
                        from = i + 1;
                    }
                    return false;
            }
            return false;
        }
    }

    /// <summary>
    /// Index of model items by normalized name. Exact mode is a hash lookup; other modes scan the distinct
    /// names once per key (distinct names are far fewer than items in a plant model).
    /// </summary>
    public sealed class KeyIndex<T> where T : class
    {
        private readonly MatchOptions _o;
        private readonly Dictionary<string, List<T>> _byName = new Dictionary<string, List<T>>(StringComparer.Ordinal);
        private List<string> _names;

        public KeyIndex(MatchOptions options) { _o = options ?? new MatchOptions(); }

        public int ItemCount { get; private set; }
        public int DistinctNames { get { return _byName.Count; } }
        public MatchOptions Options { get { return _o; } }

        public void Add(string rawName, T item)
        {
            string n = KeyNormalizer.Normalize(rawName, _o);
            if (n.Length == 0) return;
            List<T> list;
            if (!_byName.TryGetValue(n, out list)) { list = new List<T>(1); _byName[n] = list; }
            list.Add(item);
            ItemCount++;
            _names = null;
        }

        public List<T> Find(string rawKey)
        {
            var result = new List<T>();
            string k = KeyNormalizer.Normalize(rawKey, _o);
            if (k.Length == 0) return result;
            if (_o.Mode == MatchMode.Exact)
            {
                List<T> list;
                if (_byName.TryGetValue(k, out list)) result.AddRange(list);
                return result;
            }
            if (_names == null) _names = new List<string>(_byName.Keys);
            foreach (string n in _names)
                if (KeyNormalizer.IsMatch(n, k, _o.Mode)) result.AddRange(_byName[n]);
            return result;
        }

        /// <summary>
        /// Keep only the top-most matches: drop an item when one of its ancestors is also matched
        /// (a pipe and its branches both named "/150-P-1001-..." -> link only the pipe).
        /// </summary>
        public static List<T> TopMost(IList<T> items, Func<T, T> parent, IEqualityComparer<T> cmp = null)
        {
            cmp = cmp ?? EqualityComparer<T>.Default;
            var set = new HashSet<T>(items, cmp);
            var keep = new List<T>();
            var seen = new HashSet<T>(cmp);
            foreach (T it in items)
            {
                if (!seen.Add(it)) continue;
                bool covered = false;
                T p = parent(it);
                int guard = 0;
                while (p != null && guard++ < 512)
                {
                    if (set.Contains(p)) { covered = true; break; }
                    p = parent(p);
                }
                if (!covered) keep.Add(it);
            }
            return keep;
        }
    }

    public static class TextUtil
    {
        public static string Plural(int n, string one, string many)
        {
            return n.ToString(CultureInfo.InvariantCulture) + " " + (n == 1 ? one : many);
        }
    }
}
