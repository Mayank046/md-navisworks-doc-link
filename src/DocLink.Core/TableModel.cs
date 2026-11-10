// MD Doc Link - in-memory table read from / written to Excel
// Author: MD (Mayank Dholariya)
using System;
using System.Collections.Generic;

namespace DocLink.Core
{
    public sealed class TableRow
    {
        /// <summary>1-based Excel row number (0 when not read from Excel).</summary>
        public int ExcelRow;
        public readonly Dictionary<string, string> Cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>Excel hyperlink target per column, when the cell carries a hyperlink.</summary>
        public readonly Dictionary<string, string> Links = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>Hyperlink target if the cell has one, else the cell text.</summary>
        public string GetLinkOrText(string col)
        {
            string v;
            if (col != null && Links.TryGetValue(col, out v) && !string.IsNullOrWhiteSpace(v)) return v;
            return Get(col);
        }
        public string Get(string col)
        {
            if (col == null) return "";
            string v;
            return Cells.TryGetValue(col, out v) && v != null ? v : "";
        }
        public void Set(string col, string value)
        {
            if (string.IsNullOrEmpty(col)) return;
            Cells[col] = value ?? "";
        }
    }

    public sealed class DataTable
    {
        public string SheetName = Names.SheetName;
        /// <summary>1-based Excel row used as header (0 for CSV / generated tables).</summary>
        public int HeaderRow;
        public readonly List<string> Columns = new List<string>();
        public readonly List<TableRow> Rows = new List<TableRow>();

        public void EnsureColumn(string name)
        {
            if (string.IsNullOrEmpty(name)) return;
            foreach (string c in Columns)
                if (string.Equals(c, name, StringComparison.OrdinalIgnoreCase)) return;
            Columns.Add(name);
        }

        public TableRow AddRow()
        {
            var r = new TableRow();
            Rows.Add(r);
            return r;
        }
    }
}
