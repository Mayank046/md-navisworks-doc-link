// MD Doc Link - Navisworks side: scope, key index, hyperlinks (COM API), selection sets
// Author: MD (Mayank Dholariya)
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocLink.Core;
using Nw = Autodesk.Navisworks.Api;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace MDDocLink
{
    /// <summary>Which value of a model item is compared with the Excel key.</summary>
    internal sealed class PropRef
    {
        public string Category; // null => Item display name
        public string Property;
        public bool IsItemName { get { return string.IsNullOrEmpty(Category); } }
        public static readonly PropRef ItemName = new PropRef();
        public override string ToString() { return IsItemName ? Names.ItemNameProperty : Category + " > " + Property; }

        public string Read(Nw.ModelItem item)
        {
            try
            {
                if (item == null) return null;
                if (IsItemName) return item.DisplayName;
                Nw.DataProperty p = item.PropertyCategories.FindPropertyByDisplayName(Category, Property);
                return p == null ? null : NavisLinks.VariantToString(p.Value);
            }
            catch { return null; }
        }
    }

    /// <summary>One hyperlink as read from / written to Navisworks.</summary>
    internal sealed class NwLink
    {
        public string Name = "";
        public string Url = "";
        public object Com; // existing COM object (kept so category / attachment points survive)
    }

    internal sealed class ApplyStats
    {
        public int ItemsLinked, LinksWritten, ItemsMerged, Batches, Errors;
        public readonly List<string> Messages = new List<string>();
        public void Note(string m) { if (Messages.Count < 12) Messages.Add(m); }
    }

    internal static class NavisLinks
    {
        public const string HyperlinkCategoryUser = "Hyperlink";
        public const string HyperlinkCategoryInternal = "LcOaURLCategoryHyperlink";

        // ---------- scope / index ----------

        public static IEnumerable<Nw.ModelItem> EnumerateScope(Nw.Document doc, bool selectionOnly)
        {
            if (doc == null) yield break;
            if (selectionOnly)
            {
                var seen = new HashSet<Nw.ModelItem>();
                foreach (Nw.ModelItem it in doc.CurrentSelection.SelectedItems)
                {
                    if (it == null) continue;
                    foreach (Nw.ModelItem d in it.DescendantsAndSelf)
                        if (d != null && seen.Add(d)) yield return d;
                }
                yield break;
            }
            foreach (Nw.Model model in doc.Models)
            {
                if (model == null) continue;
                Nw.ModelItem root = null;
                try { root = model.RootItem; } catch { }
                if (root == null) continue;
                foreach (Nw.ModelItem it in root.DescendantsAndSelf)
                    if (it != null) yield return it;
            }
        }

        public static KeyIndex<Nw.ModelItem> BuildIndex(Nw.Document doc, bool selectionOnly, PropRef prop,
            MatchOptions options, Func<double, bool> progress, out int walked)
        {
            var idx = new KeyIndex<Nw.ModelItem>(options);
            walked = 0;
            foreach (Nw.ModelItem item in EnumerateScope(doc, selectionOnly))
            {
                walked++;
                string v = prop.Read(item);
                if (!string.IsNullOrEmpty(v)) idx.Add(v, item);
                if ((walked & 0x7FF) == 0 && progress != null && !progress(Math.Min(0.9, walked / 300000.0)))
                    throw new OperationCanceledException();
            }
            return idx;
        }

        public static List<Nw.ModelItem> FindTopMost(KeyIndex<Nw.ModelItem> idx, string key, bool topMostOnly)
        {
            List<Nw.ModelItem> hits = idx.Find(key);
            if (!topMostOnly || hits.Count < 2) return hits;
            return KeyIndex<Nw.ModelItem>.TopMost(hits, SafeParent);
        }

        private static Nw.ModelItem SafeParent(Nw.ModelItem m)
        {
            try { return m.Parent; } catch { return null; }
        }

        /// <summary>Category &gt; property names on the first selected item (for the match-property list).</summary>
        public static List<PropRef> PropertiesOf(Nw.ModelItem item)
        {
            var list = new List<PropRef>();
            if (item == null) return list;
            try
            {
                foreach (Nw.PropertyCategory cat in item.PropertyCategories)
                {
                    if (cat == null || string.IsNullOrEmpty(cat.DisplayName)) continue;
                    foreach (Nw.DataProperty p in cat.Properties)
                    {
                        if (p == null || string.IsNullOrEmpty(p.DisplayName)) continue;
                        if (list.Any(x => x.Category == cat.DisplayName && x.Property == p.DisplayName)) continue;
                        list.Add(new PropRef { Category = cat.DisplayName, Property = p.DisplayName });
                    }
                }
            }
            catch (Exception ex) { Diagnostics.Log("PropertiesOf: " + ex.Message); }
            return list;
        }

        // ---------- hyperlinks (COM) ----------

        private static ComApi.InwOpState10 State { get { return ComBridge.State; } }

        public static List<NwLink> ReadLinks(Nw.ModelItem item)
        {
            var list = new List<NwLink>();
            try
            {
                ComApi.InwOaPath path = ComBridge.ToInwOaPath(item);
                ComApi.InwURLOverride ov = State.GetOverrideURL(path, true);
                if (ov == null) return list;
                ComApi.InwURLColl urls = ov.URLs();
                if (urls == null) return list;
                foreach (object o in urls)
                {
                    var u = o as ComApi.InwURL;
                    if (u == null) continue;
                    string url = u.URL ?? "";
                    if (url.Length == 0) continue;
                    list.Add(new NwLink { Name = u.name ?? "", Url = url, Com = o });
                }
            }
            catch (Exception ex) { Diagnostics.Log("ReadLinks: " + ex.Message); }
            return list;
        }

        private static ComApi.InwURL2 NewUrl(string name, string url)
        {
            var u = (ComApi.InwURL2)State.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwURL, null, null);
            u.name = name;
            u.URL = url;
            try { u.SetCategory(HyperlinkCategoryUser, HyperlinkCategoryInternal); } catch { }
            return u;
        }

        private static object CopyExisting(NwLink l)
        {
            try
            {
                var u2 = l.Com as ComApi.InwURL2;
                if (u2 != null) return u2.Copy();
                var u = l.Com as ComApi.InwURL;
                if (u != null) return u.Copy();
            }
            catch { }
            ComApi.InwURL2 n = NewUrl(l.Name, l.Url);
            try
            {
                var old = l.Com as ComApi.InwURL2;
                if (old != null && !string.IsNullOrEmpty(old.CategoryInternalName))
                    n.SetCategory(old.CategoryUserName, old.CategoryInternalName);
            }
            catch { }
            return n;
        }

        /// <summary>Set exactly these links on every item in the collection (one COM call).</summary>
        private static void SetLinks(Nw.ModelItemCollection items, IList<NwLink> links)
        {
            ComApi.InwOpSelection sel = ComBridge.ToInwOpSelection(items);
            if (links.Count == 0) { State.OverrideURLReset(sel); return; }
            var ov = (ComApi.InwURLOverride)State.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwURLOverride, null, null);
            ComApi.InwURLColl coll = ov.URLs();
            foreach (NwLink l in links)
                coll.Add(l.Com != null ? CopyExisting(l) : NewUrl(l.Name, l.Url));
            State.SetOverrideURL(sel, ov);
        }

        private static void EnableUrlDisplay()
        {
            try { if (!State.URLsEnabled) State.URLsEnabled = true; } catch { }
        }

        /// <summary>
        /// Write links. perItem: target item -> new MD links. Existing non-MD links are kept;
        /// existing MD links are dropped when replaceMd is true.
        /// </summary>
        public static ApplyStats Apply(Dictionary<Nw.ModelItem, List<KeyValuePair<string, string>>> perItem,
            bool replaceMd, Func<double, bool> progress)
        {
            var stats = new ApplyStats();
            // Items with no current links can share one SetOverrideURL per identical link list.
            var batches = new Dictionary<string, KeyValuePair<List<NwLink>, Nw.ModelItemCollection>>(StringComparer.Ordinal);
            int done = 0, total = perItem.Count;
            foreach (KeyValuePair<Nw.ModelItem, List<KeyValuePair<string, string>>> kv in perItem)
            {
                done++;
                try
                {
                    List<NwLink> existing = ReadLinks(kv.Key);
                    var final = new List<NwLink>();
                    foreach (NwLink e in existing)
                    {
                        if (replaceMd && Names.IsMdLinkName(e.Name)) continue;
                        if (final.Any(x => SameUrl(x.Url, e.Url))) continue;
                        final.Add(e);
                    }
                    int before = final.Count;
                    foreach (KeyValuePair<string, string> a in kv.Value)
                        if (!final.Any(x => SameUrl(x.Url, a.Value)))
                            final.Add(new NwLink { Name = a.Key, Url = a.Value });
                    int added = final.Count - before;

                    if (existing.Count == 0)
                    {
                        string sig = string.Join("\n", final.Select(l => l.Name + "\t" + l.Url).ToArray());
                        KeyValuePair<List<NwLink>, Nw.ModelItemCollection> b;
                        if (!batches.TryGetValue(sig, out b))
                        {
                            b = new KeyValuePair<List<NwLink>, Nw.ModelItemCollection>(final, new Nw.ModelItemCollection());
                            batches[sig] = b;
                        }
                        b.Value.Add(kv.Key);
                    }
                    else
                    {
                        var one = new Nw.ModelItemCollection();
                        one.Add(kv.Key);
                        SetLinks(one, final);
                        stats.ItemsMerged++;
                    }
                    stats.ItemsLinked++;
                    stats.LinksWritten += added;
                }
                catch (Exception ex)
                {
                    stats.Errors++;
                    stats.Note(SafeName(kv.Key) + ": " + ex.Message);
                    Diagnostics.Log("Apply item: " + ex);
                }
                if ((done & 0x3F) == 0 && progress != null && !progress(0.1 + 0.6 * done / Math.Max(1, total)))
                {
                    stats.Note("Cancelled after " + done + " items.");
                    break;
                }
            }
            int bi = 0;
            foreach (KeyValuePair<List<NwLink>, Nw.ModelItemCollection> b in batches.Values)
            {
                bi++;
                try { SetLinks(b.Value, b.Key); stats.Batches++; }
                catch (Exception ex)
                {
                    stats.Errors++;
                    stats.Note("Batch of " + b.Value.Count + " items: " + ex.Message);
                    Diagnostics.Log("Apply batch: " + ex);
                }
                if ((bi & 0x1F) == 0 && progress != null) progress(0.7 + 0.25 * bi / Math.Max(1, batches.Count));
            }
            EnableUrlDisplay();
            Diagnostics.Log(string.Format(CultureInfo.InvariantCulture,
                "Apply items={0} links={1} merged={2} batches={3} errors={4} replace={5}",
                stats.ItemsLinked, stats.LinksWritten, stats.ItemsMerged, stats.Batches, stats.Errors, replaceMd));
            return stats;
        }

        /// <summary>Remove links named "MD | ..." from the items; other links stay.</summary>
        public static int RemoveMdLinks(IEnumerable<Nw.ModelItem> items, Func<double, bool> progress, out int linksRemoved, out int errors)
        {
            int itemsChanged = 0, walked = 0;
            linksRemoved = 0; errors = 0;
            foreach (Nw.ModelItem item in items)
            {
                walked++;
                try
                {
                    List<NwLink> links = ReadLinks(item);
                    int md = links.Count(l => Names.IsMdLinkName(l.Name));
                    if (md > 0)
                    {
                        var one = new Nw.ModelItemCollection();
                        one.Add(item);
                        SetLinks(one, links.Where(l => !Names.IsMdLinkName(l.Name)).ToList());
                        itemsChanged++;
                        linksRemoved += md;
                    }
                }
                catch (Exception ex) { errors++; Diagnostics.Log("RemoveMdLinks: " + ex.Message); }
                if ((walked & 0x3FF) == 0 && progress != null && !progress(Math.Min(0.95, walked / 300000.0)))
                    break;
            }
            return itemsChanged;
        }

        /// <summary>Links on the item and its ancestors (closest first), de-duplicated by URL.</summary>
        public static List<KeyValuePair<string, NwLink>> LinksForOpen(Nw.ModelItem item)
        {
            var list = new List<KeyValuePair<string, NwLink>>();
            if (item == null) return list;
            Nw.ModelItem cur = item;
            int depth = 0;
            while (cur != null && depth < 64)
            {
                foreach (NwLink l in ReadLinks(cur))
                    if (!list.Any(x => SameUrl(x.Value.Url, l.Url)))
                        list.Add(new KeyValuePair<string, NwLink>(SafeName(cur), l));
                cur = SafeParent(cur);
                depth++;
            }
            return list;
        }

        private static bool SameUrl(string a, string b)
        {
            return string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
        }

        public static string SafeName(Nw.ModelItem m)
        {
            try { return m.DisplayName ?? ""; } catch { return ""; }
        }

        // ---------- selection sets ----------

        public static void ReplaceSet(Nw.Document doc, string setName, Nw.ModelItemCollection items)
        {
            Nw.FolderItem folder = EnsureFolder(doc, Names.SetFolder);
            RemoveChildNamed(doc, folder, setName);
            if (items == null || items.Count == 0) return;
            var set = new Nw.SelectionSet(items);
            set.DisplayName = setName;
            doc.SelectionSets.InsertCopy(folder, folder.Children.Count, set);
        }

        private static Nw.FolderItem EnsureFolder(Nw.Document doc, string folderName)
        {
            Nw.FolderItem existing = FindFolder(doc.SelectionSets.RootItem, folderName);
            if (existing != null) return existing;
            var folder = new Nw.FolderItem();
            folder.DisplayName = folderName;
            doc.SelectionSets.InsertCopy(doc.SelectionSets.RootItem, doc.SelectionSets.RootItem.Children.Count, folder);
            existing = FindFolder(doc.SelectionSets.RootItem, folderName);
            if (existing == null) throw new InvalidOperationException("Could not create selection-set folder '" + folderName + "'.");
            return existing;
        }

        private static Nw.FolderItem FindFolder(Nw.GroupItem parent, string name)
        {
            if (parent == null) return null;
            foreach (Nw.SavedItem child in parent.Children)
            {
                var f = child as Nw.FolderItem;
                if (f != null && string.Equals(f.DisplayName, name, StringComparison.OrdinalIgnoreCase)) return f;
            }
            return null;
        }

        private static void RemoveChildNamed(Nw.Document doc, Nw.FolderItem folder, string name)
        {
            try
            {
                Nw.SavedItem victim = null;
                foreach (Nw.SavedItem child in folder.Children)
                    if (child != null && !(child is Nw.FolderItem) && string.Equals(child.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                    { victim = child; break; }
                if (victim != null) doc.SelectionSets.Remove(victim);
            }
            catch (Exception ex) { Diagnostics.Log("RemoveChildNamed: " + ex.Message); }
        }

        // ---------- values ----------

        public static string VariantToString(Nw.VariantData v)
        {
            if (v == null) return null;
            try
            {
                if (v.IsDisplayString) return v.ToDisplayString();
                if (v.IsIdentifierString) return v.ToIdentifierString();
                if (v.IsInt32) return v.ToInt32().ToString(CultureInfo.InvariantCulture);
                if (v.IsDouble) return v.ToDouble().ToString(CultureInfo.InvariantCulture);
                if (v.IsNamedConstant)
                {
                    Nw.NamedConstant nc = v.ToNamedConstant();
                    if (nc != null) return nc.DisplayName ?? nc.Name ?? "";
                }
                return v.ToDisplayString();
            }
            catch
            {
                try { return v.ToString(); } catch { return null; }
            }
        }
    }
}
