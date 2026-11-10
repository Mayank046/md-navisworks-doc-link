// MD Doc Link - minimal .xlsx reader (any worksheet, header-row detect, shared + inline strings)
// Author: MD (Mayank Dholariya)
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace DocLink.Core
{
    public static class XlsxReader
    {
        public static DataTable Load(string path)
        {
            return Load(path, null, 0);
        }

        /// <summary>Load one worksheet. sheetName null = first sheet. headerRow = 1-based Excel row, 0 = auto-detect.</summary>
        public static DataTable Load(string path, string sheetName, int headerRow)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("Excel file not found.", path);

            if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                return LoadCsv(path);

            // FileShare.ReadWrite so a workbook that is open in Excel can still be read.
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                string[] sst = LoadSharedStrings(zip);
                List<KeyValuePair<string, string>> sheets = ReadSheetList(zip);
                string sheetPath = null;
                string chosenName = null;
                foreach (KeyValuePair<string, string> kv in sheets)
                {
                    if (sheetName == null || string.Equals(kv.Key, sheetName, StringComparison.OrdinalIgnoreCase))
                    { sheetPath = kv.Value; chosenName = kv.Key; break; }
                }
                if (sheetPath == null)
                {
                    if (sheetName != null) throw new InvalidOperationException("Sheet '" + sheetName + "' not found.");
                    sheetPath = "xl/worksheets/sheet1.xml";
                }
                ZipArchiveEntry sheet = zip.GetEntry(sheetPath);
                if (sheet == null) throw new InvalidOperationException("Workbook has no worksheet.");
                Dictionary<string, string> sheetRels = ReadExternalRels(zip, sheetPath);
                Func<string, string> relTarget = id => { string v; return sheetRels.TryGetValue(id, out v) ? v : null; };
                using (Stream stream = sheet.Open())
                {
                    DataTable t = ParseSheet(stream, sst, headerRow, relTarget);
                    if (chosenName != null) t.SheetName = chosenName;
                    return t;
                }
            }
        }

        /// <summary>Worksheet names in workbook order.</summary>
        public static List<string> ListSheets(string path)
        {
            var names = new List<string>();
            if (path != null && path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                names.Add(Path.GetFileNameWithoutExtension(path));
                return names;
            }
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                foreach (KeyValuePair<string, string> kv in ReadSheetList(zip)) names.Add(kv.Key);
            return names;
        }

        /// <summary>Hyperlink relationship id -> target for one worksheet (xl/worksheets/_rels/sheetN.xml.rels).</summary>
        private static Dictionary<string, string> ReadExternalRels(ZipArchive zip, string sheetPath)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            int slash = sheetPath.LastIndexOf('/');
            string relPath = (slash >= 0 ? sheetPath.Substring(0, slash + 1) : "") + "_rels/" + sheetPath.Substring(slash + 1) + ".rels";
            ZipArchiveEntry e = zip.GetEntry(relPath);
            if (e == null) return map;
            using (Stream stream = e.Open())
            using (XmlReader xr = XmlReader.Create(stream))
            {
                while (xr.Read())
                {
                    if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "Relationship") continue;
                    string type = xr.GetAttribute("Type") ?? "";
                    if (!type.EndsWith("/hyperlink", StringComparison.OrdinalIgnoreCase)) continue;
                    string id = xr.GetAttribute("Id") ?? "";
                    string target = xr.GetAttribute("Target") ?? "";
                    if (id.Length > 0 && target.Length > 0) map[id] = Uri.UnescapeDataString(target);
                }
            }
            return map;
        }

        /// <summary>(sheet name, zip entry path) for worksheets (chartsheets skipped).</summary>
        private static List<KeyValuePair<string, string>> ReadSheetList(ZipArchive zip)
        {
            var relTargets = new Dictionary<string, string>(StringComparer.Ordinal);
            ZipArchiveEntry rels = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (rels != null)
            {
                using (Stream stream = rels.Open())
                using (XmlReader xr = XmlReader.Create(stream))
                {
                    while (xr.Read())
                    {
                        if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "Relationship") continue;
                        string type = xr.GetAttribute("Type") ?? "";
                        string target = (xr.GetAttribute("Target") ?? "").Replace('\\', '/');
                        string id = xr.GetAttribute("Id") ?? "";
                        if (!type.EndsWith("/worksheet", StringComparison.OrdinalIgnoreCase) || target.Length == 0) continue;
                        relTargets[id] = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
                    }
                }
            }
            var list = new List<KeyValuePair<string, string>>();
            ZipArchiveEntry wb = zip.GetEntry("xl/workbook.xml");
            if (wb != null)
            {
                using (Stream stream = wb.Open())
                using (XmlReader xr = XmlReader.Create(stream))
                {
                    while (xr.Read())
                    {
                        if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "sheet") continue;
                        string name = xr.GetAttribute("name") ?? "";
                        string rid = null;
                        if (xr.MoveToFirstAttribute())
                        {
                            do { if (xr.LocalName == "id" && xr.Prefix.Length > 0) rid = xr.Value; } while (xr.MoveToNextAttribute());
                            xr.MoveToElement();
                        }
                        string target;
                        if (rid != null && relTargets.TryGetValue(rid, out target))
                            list.Add(new KeyValuePair<string, string>(name, target));
                    }
                }
            }
            return list;
        }

        public static DataTable LoadCsv(string path)
        {
            var table = new DataTable();
            table.SheetName = Path.GetFileNameWithoutExtension(path);
            string[] lines = File.ReadAllLines(path, Encoding.UTF8);
            if (lines.Length == 0) return table;
            List<string> header = ParseCsvLine(lines[0]);
            foreach (string h in header) table.EnsureColumn(h);
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                List<string> cells = ParseCsvLine(lines[i]);
                TableRow row = table.AddRow();
                for (int c = 0; c < table.Columns.Count; c++)
                    row.Set(table.Columns[c], c < cells.Count ? cells[c] : "");
            }
            return table;
        }

        private static List<string> ParseCsvLine(string line)
        {
            var cells = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                        else inQuotes = false;
                    }
                    else sb.Append(ch);
                }
                else
                {
                    if (ch == '"') inQuotes = true;
                    else if (ch == ',') { cells.Add(sb.ToString()); sb.Clear(); }
                    else sb.Append(ch);
                }
            }
            cells.Add(sb.ToString());
            return cells;
        }

        private static string[] LoadSharedStrings(ZipArchive zip)
        {
            ZipArchiveEntry e = zip.GetEntry("xl/sharedStrings.xml");
            if (e == null) return new string[0];
            var list = new List<string>();
            using (Stream stream = e.Open())
            using (XmlReader xr = XmlReader.Create(stream, new XmlReaderSettings { IgnoreComments = true, IgnoreWhitespace = true }))
            {
                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "si")
                        list.Add(ReadSi(xr));
                }
            }
            return list.ToArray();
        }

        private static string ReadSi(XmlReader xr)
        {
            // xr is positioned on <si>. ReadElementContentAsString on child <t> lands on
            // the parent EndElement, so we must break without an extra Read().
            var sb = new StringBuilder();
            if (xr.IsEmptyElement) return "";
            int depth = xr.Depth;
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.EndElement && xr.Depth <= depth)
                    break;
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "t")
                {
                    sb.Append(ReadElementText(xr));
                    if (xr.NodeType == XmlNodeType.EndElement && xr.Depth <= depth)
                        break;
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Reads text of the current element without skipping the parent's EndElement.
        /// XmlReader.ReadElementContentAsString lands on the parent EndElement when this
        /// element was the last child — callers must not Read() past that.
        /// </summary>
        private static string ReadElementText(XmlReader xr)
        {
            if (xr.IsEmptyElement) return "";
            var sb = new StringBuilder();
            int depth = xr.Depth;
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth)
                    break;
                if (xr.NodeType == XmlNodeType.Text
                    || xr.NodeType == XmlNodeType.SignificantWhitespace
                    || xr.NodeType == XmlNodeType.CDATA
                    || xr.NodeType == XmlNodeType.Whitespace)
                    sb.Append(xr.Value);
            }
            return sb.ToString();
        }

        private static DataTable ParseSheet(Stream stream, string[] sst, int requestedHeaderRow)
        {
            return ParseSheet(stream, sst, requestedHeaderRow, null);
        }

        /// <param name="relTarget">Maps a hyperlink r:id to its external target (sheet rels); may be null.</param>
        private static DataTable ParseSheet(Stream stream, string[] sst, int requestedHeaderRow, Func<string, string> relTarget)
        {
            var links = new Dictionary<long, string>();
            var table = new DataTable();
            // Map: rowIndex -> (colIndex -> value)
            var grid = new SortedDictionary<int, Dictionary<int, string>>();
            int maxCol = -1;

            using (XmlReader xr = XmlReader.Create(stream, new XmlReaderSettings { IgnoreComments = true }))
            {
                while (xr.Read())
                {
                    if (xr.NodeType != XmlNodeType.Element) continue;
                    if (xr.LocalName == "hyperlink")
                    {
                        string href = xr.GetAttribute("ref") ?? "";
                        string rid = null;
                        if (xr.MoveToFirstAttribute())
                        {
                            do { if (xr.LocalName == "id" && xr.Prefix.Length > 0) rid = xr.Value; } while (xr.MoveToNextAttribute());
                            xr.MoveToElement();
                        }
                        string target = rid != null && relTarget != null ? relTarget(rid) : null;
                        int hc, hr;
                        int colon = href.IndexOf(':');
                        if (!string.IsNullOrEmpty(target) && TryParseCellRef(colon > 0 ? href.Substring(0, colon) : href, out hc, out hr))
                            links[((long)hr << 20) | (uint)hc] = target;
                        continue;
                    }
                    if (xr.LocalName != "c") continue;
                    string r = xr.GetAttribute("r") ?? "";
                    string t = xr.GetAttribute("t") ?? "";
                    int row, col;
                    if (!TryParseCellRef(r, out col, out row)) continue;
                    string raw = ReadCellValue(xr);
                    string text = DecodeCell(t, raw, sst);
                    Dictionary<int, string> rowMap;
                    if (!grid.TryGetValue(row, out rowMap))
                    {
                        rowMap = new Dictionary<int, string>();
                        grid[row] = rowMap;
                    }
                    rowMap[col] = text;
                    if (col > maxCol) maxCol = col;
                }
            }

            if (grid.Count == 0) return table;

            int headerRow = requestedHeaderRow > 0 ? requestedHeaderRow : DetectHeaderRow(grid);
            Dictionary<int, string> headerCells;
            if (!grid.TryGetValue(headerRow, out headerCells)) headerCells = new Dictionary<int, string>();
            table.HeaderRow = headerRow;
            var colIndexToName = new Dictionary<int, string>();
            for (int c = 0; c <= maxCol; c++)
            {
                string name;
                if (!headerCells.TryGetValue(c, out name) || string.IsNullOrWhiteSpace(name))
                    name = "Column" + XlsxWriter.ColumnLetter(c);
                name = System.Text.RegularExpressions.Regex.Replace(name, @"\s+", " ").Trim();
                // de-dupe
                string unique = name;
                int n = 2;
                while (table.Columns.Exists(x => string.Equals(x, unique, StringComparison.OrdinalIgnoreCase)))
                {
                    unique = name + "_" + n;
                    n++;
                }
                table.EnsureColumn(unique);
                colIndexToName[c] = unique;
            }

            foreach (KeyValuePair<int, Dictionary<int, string>> kv in grid)
            {
                if (kv.Key <= headerRow) continue;
                bool any = false;
                foreach (string v in kv.Value.Values) if (!string.IsNullOrWhiteSpace(v)) { any = true; break; }
                if (!any) continue;
                TableRow row = table.AddRow();
                row.ExcelRow = kv.Key;
                foreach (KeyValuePair<int, string> cell in kv.Value)
                {
                    string colName;
                    if (colIndexToName.TryGetValue(cell.Key, out colName))
                    {
                        row.Set(colName, cell.Value ?? "");
                        string target;
                        if (links.TryGetValue(((long)kv.Key << 20) | (uint)cell.Key, out target))
                            row.Links[colName] = target;
                    }
                }
            }
            return table;
        }

        /// <summary>
        /// Header = the row (within the first 30 used rows) with the most non-empty, non-numeric cells.
        /// Title rows above a header and unit rows below it have fewer text cells. Ties keep the earlier row.
        /// </summary>
        public static int DetectHeaderRow(SortedDictionary<int, Dictionary<int, string>> grid)
        {
            int best = -1, bestScore = -1, seen = 0;
            foreach (KeyValuePair<int, Dictionary<int, string>> kv in grid)
            {
                int score = 0;
                foreach (string v in kv.Value.Values)
                {
                    if (string.IsNullOrWhiteSpace(v)) continue;
                    double d;
                    if (double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) continue;
                    score++;
                }
                if (score > bestScore) { best = kv.Key; bestScore = score; }
                if (++seen >= 30) break;
            }
            return best;
        }

        private static string ReadCellValue(XmlReader xr)
        {
            // xr is positioned on <c>. Do not use ReadElementContentAsString on <v>/<t>
            // — it lands on </c> and the next Read() would skip into the following cell.
            if (xr.IsEmptyElement) return "";
            string v = "";
            string isText = null;
            int depth = xr.Depth;
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.EndElement && xr.Depth <= depth)
                    break;
                if (xr.NodeType != XmlNodeType.Element) continue;
                if (xr.LocalName == "v")
                    v = ReadElementText(xr);
                else if (xr.LocalName == "t" && xr.Depth > depth)
                    isText = (isText ?? "") + ReadElementText(xr);
                else if (xr.LocalName == "is" && !xr.IsEmptyElement)
                {
                    int d2 = xr.Depth;
                    while (xr.Read())
                    {
                        if (xr.NodeType == XmlNodeType.EndElement && xr.Depth <= d2) break;
                        if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "t")
                            isText = (isText ?? "") + ReadElementText(xr);
                    }
                }
                if (xr.NodeType == XmlNodeType.EndElement && xr.Depth <= depth)
                    break;
            }
            return isText ?? v ?? "";
        }

        private static string DecodeCell(string t, string raw, string[] sst)
        {
            if (string.Equals(t, "s", StringComparison.OrdinalIgnoreCase))
            {
                int idx;
                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx)
                    && idx >= 0 && idx < sst.Length)
                    return sst[idx];
                return raw ?? "";
            }
            if (string.Equals(t, "b", StringComparison.OrdinalIgnoreCase))
                return raw == "1" ? "TRUE" : "FALSE";
            if (string.Equals(t, "inlineStr", StringComparison.OrdinalIgnoreCase))
                return raw ?? "";
            return raw ?? "";
        }

        private static bool TryParseCellRef(string r, out int col0, out int row1)
        {
            col0 = 0; row1 = 0;
            if (string.IsNullOrEmpty(r)) return false;
            int i = 0;
            int col = 0;
            while (i < r.Length && r[i] >= 'A' && r[i] <= 'Z')
            {
                col = col * 26 + (r[i] - 'A' + 1);
                i++;
            }
            if (col == 0) return false;
            int row;
            if (!int.TryParse(r.Substring(i), NumberStyles.Integer, CultureInfo.InvariantCulture, out row) || row < 1)
                return false;
            col0 = col - 1;
            row1 = row;
            return true;
        }
    }
}
