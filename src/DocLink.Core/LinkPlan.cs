// MD Doc Link - per-Excel-row plan: key, documents, link label, status
// Author: MD (Mayank Dholariya)
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DocLink.Core
{
    public enum RowStatus { Matched, NotFound, DocMissing, NoDocument, NoKey, Pending }

    public sealed class PlanRow
    {
        public int ExcelRow;
        public string Key = "";
        public string LinkLabel = "";
        public readonly List<ResolvedDoc> Docs = new List<ResolvedDoc>();
        public int MatchedItems;
        public RowStatus Status = RowStatus.Pending;

        public bool HasKey { get { return Key.Trim().Length > 0; } }
        public int MissingDocs { get { return Docs.Count(d => !d.Exists); } }
        public IEnumerable<ResolvedDoc> UsableDocs(bool includeMissing)
        {
            return Docs.Where(d => includeMissing || d.Exists);
        }

        public string DocsText
        {
            get { return string.Join("; ", Docs.Select(d => d.Target).ToArray()); }
        }

        public static string StatusText(RowStatus s)
        {
            switch (s)
            {
                case RowStatus.Matched: return "Matched";
                case RowStatus.NotFound: return "Not found in model";
                case RowStatus.DocMissing: return "Doc file missing";
                case RowStatus.NoDocument: return "No document";
                case RowStatus.NoKey: return "No key";
                default: return "Not checked";
            }
        }

        /// <summary>Status once the model has been searched. Key/doc problems win over model problems.</summary>
        public void Finish(int matchedItems)
        {
            MatchedItems = matchedItems;
            if (!HasKey) Status = RowStatus.NoKey;
            else if (Docs.Count == 0) Status = RowStatus.NoDocument;
            else if (matchedItems == 0) Status = RowStatus.NotFound;
            else if (MissingDocs > 0) Status = RowStatus.DocMissing;
            else Status = RowStatus.Matched;
        }

        /// <summary>Navisworks link names for this row's docs ("MD | label" or "MD | file.pdf").</summary>
        public List<KeyValuePair<string, string>> LinkPairs(bool includeMissing)
        {
            var list = new List<KeyValuePair<string, string>>();
            List<ResolvedDoc> docs = UsableDocs(includeMissing).ToList();
            for (int i = 0; i < docs.Count; i++)
            {
                string label;
                if (LinkLabel.Trim().Length > 0)
                    label = docs.Count == 1 ? LinkLabel : LinkLabel + " (" + (i + 1).ToString(CultureInfo.InvariantCulture) + ")";
                else
                    label = docs[i].DisplayName;
                list.Add(new KeyValuePair<string, string>(Names.MdLinkName(label), docs[i].Target));
            }
            return list;
        }
    }

    public sealed class PlanSettings
    {
        public string KeyColumn;
        public List<string> DocColumns = new List<string>();
        public string LabelColumn;
    }

    public static class LinkPlanner
    {
        public static List<PlanRow> Build(DataTable table, PlanSettings s, DocResolver resolver)
        {
            var rows = new List<PlanRow>();
            if (table == null || s == null || string.IsNullOrEmpty(s.KeyColumn)) return rows;
            foreach (TableRow r in table.Rows)
            {
                var p = new PlanRow
                {
                    ExcelRow = r.ExcelRow,
                    Key = Names.Clean(r.Get(s.KeyColumn)),
                    LinkLabel = string.IsNullOrEmpty(s.LabelColumn) ? "" : Names.Clean(r.Get(s.LabelColumn))
                };
                var cells = new List<KeyValuePair<string, string>>();
                foreach (string c in s.DocColumns)
                    cells.Add(new KeyValuePair<string, string>(c, r.GetLinkOrText(c)));
                if (p.HasKey) p.Docs.AddRange(resolver.Resolve(p.Key, cells));
                rows.Add(p);
            }
            return rows;
        }

        /// <summary>Merge links from many rows onto each item, de-duplicated by target (first label wins).</summary>
        public static Dictionary<TItem, List<KeyValuePair<string, string>>> MergePerItem<TItem>(
            IEnumerable<KeyValuePair<PlanRow, IList<TItem>>> rowsWithItems, bool includeMissing, IEqualityComparer<TItem> cmp = null)
        {
            var map = new Dictionary<TItem, List<KeyValuePair<string, string>>>(cmp ?? EqualityComparer<TItem>.Default);
            foreach (KeyValuePair<PlanRow, IList<TItem>> kv in rowsWithItems)
            {
                List<KeyValuePair<string, string>> pairs = kv.Key.LinkPairs(includeMissing);
                if (pairs.Count == 0) continue;
                foreach (TItem it in kv.Value)
                {
                    List<KeyValuePair<string, string>> list;
                    if (!map.TryGetValue(it, out list)) { list = new List<KeyValuePair<string, string>>(); map[it] = list; }
                    foreach (KeyValuePair<string, string> p in pairs)
                        if (!list.Any(x => string.Equals(x.Value, p.Value, StringComparison.OrdinalIgnoreCase))) list.Add(p);
                }
            }
            return map;
        }

        /// <summary>
        /// Final link list for an item. Replace mode drops existing MD links; append keeps them.
        /// Non-MD links (from the model or added by hand) are always kept. De-duplicated by target.
        /// </summary>
        public static List<KeyValuePair<string, string>> Combine(IEnumerable<KeyValuePair<string, string>> existing,
            IEnumerable<KeyValuePair<string, string>> added, bool replaceMdLinks)
        {
            var result = new List<KeyValuePair<string, string>>();
            if (existing != null)
                foreach (KeyValuePair<string, string> e in existing)
                {
                    if (replaceMdLinks && Names.IsMdLinkName(e.Key)) continue;
                    result.Add(e);
                }
            if (added != null)
                foreach (KeyValuePair<string, string> a in added)
                    if (!result.Any(x => string.Equals(x.Value, a.Value, StringComparison.OrdinalIgnoreCase))) result.Add(a);
            return result;
        }

        public static DataTable ReportTable(IEnumerable<PlanRow> rows, string sheetName)
        {
            var t = new DataTable { SheetName = sheetName ?? "DocLink Report" };
            foreach (string c in new[] { "Excel Row", "Key", "Link Name", "Documents", "Missing Docs", "Matched Items", "Status" })
                t.EnsureColumn(c);
            foreach (PlanRow p in rows)
            {
                TableRow r = t.AddRow();
                r.Set("Excel Row", p.ExcelRow.ToString(CultureInfo.InvariantCulture));
                r.Set("Key", p.Key);
                r.Set("Link Name", p.LinkLabel);
                r.Set("Documents", p.DocsText);
                r.Set("Missing Docs", string.Join("; ", p.Docs.Where(d => !d.Exists).Select(d => d.Target).ToArray()));
                r.Set("Matched Items", p.MatchedItems.ToString(CultureInfo.InvariantCulture));
                r.Set("Status", PlanRow.StatusText(p.Status));
            }
            return t;
        }
    }
}
