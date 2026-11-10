// MD Doc Link - turn Excel document cells (paths, folders, URLs) and a doc folder + pattern into link targets
// Author: MD (Mayank Dholariya)
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace DocLink.Core
{
    public interface IFileSystem
    {
        bool FileExists(string path);
        bool DirectoryExists(string path);
        IEnumerable<string> EnumerateFiles(string folder, bool recursive);
    }

    public sealed class RealFileSystem : IFileSystem
    {
        public bool FileExists(string path) { try { return File.Exists(path); } catch { return false; } }
        public bool DirectoryExists(string path) { try { return Directory.Exists(path); } catch { return false; } }
        public IEnumerable<string> EnumerateFiles(string folder, bool recursive)
        {
            try { return Directory.GetFiles(folder, "*", recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly); }
            catch { return new string[0]; }
        }
    }

    public enum DocKind { File, Folder, Url }

    public sealed class ResolvedDoc
    {
        public string Target;
        public DocKind Kind;
        public bool Exists;
        /// <summary>Where it came from: column name or "Doc folder".</summary>
        public string Source;
        public string DisplayName
        {
            get
            {
                if (Kind == DocKind.Url) return Target;
                string t = Target.TrimEnd('\\', '/');
                int i = t.LastIndexOfAny(new[] { '\\', '/' });
                string n = i >= 0 ? t.Substring(i + 1) : t;
                return n.Length == 0 ? Target : n;
            }
        }
    }

    public sealed class DocOptions
    {
        /// <summary>Relative paths in Excel are resolved against this (default: the Excel file's folder).</summary>
        public string BaseFolder;
        /// <summary>Optional folder of documents searched by key (e.g. stress iso PDFs named by line number).</summary>
        public string SearchFolder;
        /// <summary>File name pattern for the doc folder search; {KEY} is replaced by the row key. * and ? wildcards.</summary>
        public string Pattern = "*{KEY}*.pdf";
        public bool Recursive = true;
        /// <summary>When a document cell holds a folder, look in it for files matching the pattern; otherwise link the folder.</summary>
        public bool SearchCellFolders = true;
        /// <summary>Search the doc folder only for rows whose document cells are empty (true) or always (false).</summary>
        public bool FolderOnlyWhenCellsEmpty = true;
        /// <summary>Compare file names on letters/digits only (ISO-1001A == ISO1001A).</summary>
        public bool LooseFileNames = false;
    }

    public sealed class DocResolver
    {
        private static readonly Regex UrlScheme = new Regex(@"^[A-Za-z][A-Za-z0-9+.\-]{1,}:(//|[^\\/])", RegexOptions.CultureInvariant);
        private readonly IFileSystem _fs;
        private readonly DocOptions _o;
        private readonly Dictionary<string, List<string>> _listing = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        public DocResolver(DocOptions options, IFileSystem fs)
        {
            _o = options ?? new DocOptions();
            _fs = fs ?? new RealFileSystem();
        }

        public DocOptions Options { get { return _o; } }

        /// <summary>Split a cell holding several documents (";" or line breaks).</summary>
        public static List<string> SplitCell(string cell)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(cell)) return list;
            foreach (string part in cell.Split(new[] { ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string p = part.Trim().Trim('"').Trim();
                if (p.Length > 0) list.Add(p);
            }
            return list;
        }

        public static bool LooksLikeUrl(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();
            if (s.StartsWith("www.", StringComparison.OrdinalIgnoreCase)) return true;
            if (s.Length >= 2 && s[1] == ':' && char.IsLetter(s[0])) return false; // C:\ drive path
            return UrlScheme.IsMatch(s);
        }

        /// <summary>Resolve one row: its document cells (already split by column) plus the doc folder search.</summary>
        public List<ResolvedDoc> Resolve(string key, IList<KeyValuePair<string, string>> docCells)
        {
            var docs = new List<ResolvedDoc>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool anyCell = false;
            if (docCells != null)
            {
                foreach (KeyValuePair<string, string> cell in docCells)
                {
                    foreach (string entry in SplitCell(cell.Value))
                    {
                        anyCell = true;
                        foreach (ResolvedDoc d in ResolveEntry(entry, key, cell.Key))
                            if (seen.Add(d.Target)) docs.Add(d);
                    }
                }
            }
            if (!string.IsNullOrWhiteSpace(_o.SearchFolder) && (!anyCell || !_o.FolderOnlyWhenCellsEmpty))
            {
                foreach (string f in FindInFolder(_o.SearchFolder, key))
                    if (seen.Add(f)) docs.Add(new ResolvedDoc { Target = f, Kind = DocKind.File, Exists = true, Source = "Doc folder" });
            }
            return docs;
        }

        public IEnumerable<ResolvedDoc> ResolveEntry(string entry, string key, string source)
        {
            if (LooksLikeUrl(entry))
            {
                string u = entry.StartsWith("www.", StringComparison.OrdinalIgnoreCase) ? "https://" + entry : entry;
                if (u.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        string local = new Uri(u).LocalPath;
                        return new[] { new ResolvedDoc { Target = local, Kind = DocKind.File, Exists = _fs.FileExists(local), Source = source } };
                    }
                    catch { }
                }
                return new[] { new ResolvedDoc { Target = u, Kind = DocKind.Url, Exists = true, Source = source } };
            }
            string path = MakeAbsolute(entry);
            if (_fs.FileExists(path))
                return new[] { new ResolvedDoc { Target = path, Kind = DocKind.File, Exists = true, Source = source } };
            if (_fs.DirectoryExists(path))
            {
                if (_o.SearchCellFolders && !string.IsNullOrWhiteSpace(key))
                {
                    var found = new List<ResolvedDoc>();
                    foreach (string f in FindInFolder(path, key))
                        found.Add(new ResolvedDoc { Target = f, Kind = DocKind.File, Exists = true, Source = source });
                    if (found.Count > 0) return found;
                }
                return new[] { new ResolvedDoc { Target = path, Kind = DocKind.Folder, Exists = true, Source = source } };
            }
            return new[] { new ResolvedDoc { Target = path, Kind = DocKind.File, Exists = false, Source = source } };
        }

        public string MakeAbsolute(string p)
        {
            p = p.Trim();
            bool rooted = p.StartsWith("\\\\", StringComparison.Ordinal) || (p.Length >= 2 && p[1] == ':') || p.StartsWith("/", StringComparison.Ordinal);
            if (rooted || string.IsNullOrWhiteSpace(_o.BaseFolder)) return p;
            string b = _o.BaseFolder.TrimEnd('\\', '/');
            char sep = b.IndexOf('\\') >= 0 || (b.Length >= 2 && b[1] == ':') ? '\\' : '/';
            string rel = p.Replace(sep == '\\' ? '/' : '\\', sep);
            while (rel.StartsWith("." + sep, StringComparison.Ordinal)) rel = rel.Substring(2);
            while (rel.StartsWith(".." + sep, StringComparison.Ordinal))
            {
                int cut = b.LastIndexOf(sep);
                if (cut <= 0) break;
                b = b.Substring(0, cut);
                rel = rel.Substring(3);
            }
            return b + sep + rel;
        }

        /// <summary>Files in folder whose names match the pattern for this key, sorted by name.</summary>
        public List<string> FindInFolder(string folder, string key)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(folder) || string.IsNullOrWhiteSpace(key)) return result;
            List<string> files;
            if (!_listing.TryGetValue(folder, out files))
            {
                files = new List<string>(_fs.DirectoryExists(folder) ? _fs.EnumerateFiles(folder, _o.Recursive) : new string[0]);
                files.Sort(StringComparer.OrdinalIgnoreCase);
                _listing[folder] = files;
            }
            Regex rx = PatternRegex(_o.Pattern, key, _o.LooseFileNames);
            foreach (string f in files)
            {
                string name = FileNameOf(f);
                if (_o.LooseFileNames) name = LooseName(name);
                if (rx.IsMatch(name)) result.Add(f);
            }
            return result;
        }

        private static string FileNameOf(string path)
        {
            int i = Math.Max(path.LastIndexOf('\\'), path.LastIndexOf('/'));
            return i >= 0 ? path.Substring(i + 1) : path;
        }

        /// <summary>Letters/digits of the file stem, upper-cased, plus the original extension.</summary>
        public static string LooseName(string fileName)
        {
            int dot = fileName.LastIndexOf('.');
            string stem = dot > 0 ? fileName.Substring(0, dot) : fileName;
            string ext = dot > 0 ? fileName.Substring(dot) : "";
            var sb = new StringBuilder(stem.Length);
            foreach (char c in stem) if (char.IsLetterOrDigit(c)) sb.Append(char.ToUpperInvariant(c));
            return sb.ToString() + ext;
        }

        /// <summary>
        /// Build a whole-file-name regex from a glob pattern with {KEY}. The key is taken literally, except a
        /// leading '/' is dropped and characters not allowed in file names (/ \ : * ? " &lt; &gt; |) match any one char.
        /// </summary>
        public static Regex PatternRegex(string pattern, string key, bool loose)
        {
            if (string.IsNullOrWhiteSpace(pattern)) pattern = "*{KEY}*";
            string k = (key ?? "").Trim().TrimStart('/');
            var keyRx = new StringBuilder();
            if (loose)
            {
                foreach (char c in k) if (char.IsLetterOrDigit(c)) keyRx.Append(Regex.Escape(char.ToUpperInvariant(c).ToString()));
            }
            else
            {
                foreach (char c in k)
                {
                    if ("/\\:*?\"<>|".IndexOf(c) >= 0) keyRx.Append('.');
                    else keyRx.Append(Regex.Escape(c.ToString()));
                }
            }
            var sb = new StringBuilder("^");
            int i = 0;
            bool hadKey = false;
            while (i < pattern.Length)
            {
                if (string.Compare(pattern, i, "{KEY}", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    sb.Append(keyRx);
                    hadKey = true;
                    i += 5;
                    continue;
                }
                char c = pattern[i++];
                if (c == '*') sb.Append(".*");
                else if (c == '?') sb.Append('.');
                else if (loose && !char.IsLetterOrDigit(c) && c != '.') continue; // loose names have no separators
                else sb.Append(Regex.Escape(c.ToString()));
            }
            if (!hadKey) sb.Insert(1, ".*" + keyRx + ".*");
            sb.Append('$');
            return new Regex(sb.ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }
    }
}
