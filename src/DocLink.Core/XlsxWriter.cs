// MD Doc Link - minimal .xlsx writer (no Excel / ClosedXML needed)
// Author: MD (Mayank Dholariya)
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace DocLink.Core
{
    public static class XlsxWriter
    {
        public static void Save(string path, DataTable table, string author)
        {
            if (table == null) throw new ArgumentNullException("table");
            if (File.Exists(path)) File.Delete(path);
            using (var fs = new FileStream(path, FileMode.CreateNew))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                var sst = new SharedStrings();
                string sheetXml = BuildSheet(table, sst);
                Put(zip, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                    "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
                    "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/></Types>");
                Put(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/></Relationships>");
                Put(zip, "docProps/core.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                    "<dc:creator>" + X(author ?? Names.Author) + "</dc:creator><dc:title>" + X(Names.Product) + "</dc:title>" +
                    "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture) + "</dcterms:created></cp:coreProperties>");
                string sheetName = SafeSheetName(table.SheetName);
                Put(zip, "xl/workbook.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<sheets><sheet name=\"" + X(sheetName) + "\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                Put(zip, "xl/_rels/workbook.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
                    "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>");
                Put(zip, "xl/styles.xml", Styles());
                Put(zip, "xl/sharedStrings.xml", sst.ToXml());
                Put(zip, "xl/worksheets/sheet1.xml", sheetXml);
            }
        }

        private static string BuildSheet(DataTable table, SharedStrings sst)
        {
            int cols = table.Columns.Count;
            int rows = table.Rows.Count + 1;
            var sb = new StringBuilder(Math.Max(4096, rows * Math.Max(1, cols) * 24));
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>");
            if (cols > 0)
            {
                sb.Append("<cols>");
                for (int c = 0; c < cols; c++)
                {
                    double w = c == 0 ? 38 : (c <= 2 ? 22 : 18);
                    sb.Append("<col min=\"").Append(c + 1).Append("\" max=\"").Append(c + 1)
                      .Append("\" width=\"").Append(w.ToString(CultureInfo.InvariantCulture)).Append("\" customWidth=\"1\"/>");
                }
                sb.Append("</cols>");
            }
            sb.Append("<sheetData>");
            sb.Append("<row r=\"1\">");
            for (int c = 0; c < cols; c++)
            {
                int si = sst.Add(table.Columns[c] ?? "");
                sb.Append("<c r=\"").Append(CellRef(c, 0)).Append("\" t=\"s\" s=\"1\"><v>").Append(si).Append("</v></c>");
            }
            sb.Append("</row>");
            for (int r = 0; r < table.Rows.Count; r++)
            {
                TableRow row = table.Rows[r];
                sb.Append("<row r=\"").Append(r + 2).Append("\">");
                for (int c = 0; c < cols; c++)
                {
                    string val = row.Get(table.Columns[c]);
                    int si = sst.Add(val ?? "");
                    sb.Append("<c r=\"").Append(CellRef(c, r + 1)).Append("\" t=\"s\"><v>").Append(si).Append("</v></c>");
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData>");
            if (cols > 0 && rows > 0)
                sb.Append("<autoFilter ref=\"A1:").Append(CellRef(cols - 1, rows - 1)).Append("\"/>");
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        private static string Styles()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<fonts count=\"2\"><font><sz val=\"10\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"10\"/><name val=\"Calibri\"/></font></fonts>" +
                "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>" +
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFCCE8EC\"/></patternFill></fill></fills>" +
                "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                "<cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
                "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/></cellXfs>" +
                "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";
        }

        private sealed class SharedStrings
        {
            private readonly Dictionary<string, int> _map = new Dictionary<string, int>(StringComparer.Ordinal);
            private readonly List<string> _list = new List<string>();
            public int Add(string s)
            {
                if (s == null) s = "";
                int i;
                if (_map.TryGetValue(s, out i)) return i;
                i = _list.Count;
                _list.Add(s);
                _map[s] = i;
                return i;
            }
            public string ToXml()
            {
                var sb = new StringBuilder(Math.Max(256, _list.Count * 32));
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"")
                  .Append(_list.Count).Append("\" uniqueCount=\"").Append(_list.Count).Append("\">");
                foreach (string s in _list)
                    sb.Append("<si><t xml:space=\"preserve\">").Append(X(s)).Append("</t></si>");
                sb.Append("</sst>");
                return sb.ToString();
            }
        }

        private static string CellRef(int col0, int row0)
        {
            return ColumnLetter(col0) + (row0 + 1).ToString(CultureInfo.InvariantCulture);
        }

        public static string ColumnLetter(int col0)
        {
            int n = col0 + 1;
            var sb = new StringBuilder();
            while (n > 0)
            {
                int rem = (n - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                n = (n - 1) / 26;
            }
            return sb.ToString();
        }

        private static string SafeSheetName(string n)
        {
            n = string.IsNullOrWhiteSpace(n) ? Names.SheetName : n.Trim();
            foreach (char c in "[]:*?/\\") n = n.Replace(c, '_');
            return n.Length > 31 ? n.Substring(0, 31) : n;
        }

        private static void Put(ZipArchive zip, string name, string content)
        {
            var e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var w = new StreamWriter(e.Open(), new UTF8Encoding(false)))
                w.Write(content);
        }

        private static string X(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }
    }
}
