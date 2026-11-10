// MD Doc Link - docked panel UI
// Author: MD (Mayank Dholariya)
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocLink.Core;
using Nw = Autodesk.Navisworks.Api;

namespace MDDocLink
{
    internal static class Theme
    {
        internal static readonly Color Surface = Color.FromArgb(243, 244, 246);
        internal static readonly Color Card = Color.White;
        internal static readonly Color CardBorder = Color.FromArgb(222, 225, 230);
        internal static readonly Color Text = Color.FromArgb(32, 33, 36);
        internal static readonly Color Subtle = Color.FromArgb(96, 100, 108);
        internal static readonly Color Accent = Color.FromArgb(28, 96, 160);
        internal static Font Ui(float size, FontStyle style = FontStyle.Regular) { return new Font("Segoe UI", size, style); }
        internal static Font Semibold(float size)
        {
            try
            {
                using (var probe = new Font("Segoe UI Semibold", size))
                    if (string.Equals(probe.Name, "Segoe UI Semibold", StringComparison.OrdinalIgnoreCase))
                        return new Font("Segoe UI Semibold", size);
            }
            catch { }
            return new Font("Segoe UI", size, FontStyle.Bold);
        }
    }

    internal sealed class Card : Panel
    {
        private readonly string _title;
        private readonly Font _titleFont = Theme.Semibold(9.75F);
        internal Card(string title, Control content)
        {
            _title = title;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            BackColor = Theme.Surface;
            Dock = DockStyle.Top;
            Margin = new Padding(0, 0, 0, 8);
            Padding = new Padding(12, 34, 12, 12);
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            content.Dock = DockStyle.Top;
            content.BackColor = Theme.Card;
            Controls.Add(content);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            using (GraphicsPath path = Rounded(r, 6))
            using (var fill = new SolidBrush(Theme.Card))
            using (var pen = new Pen(Theme.CardBorder))
            {
                g.FillPath(fill, path);
                g.DrawPath(pen, path);
            }
            TextRenderer.DrawText(g, _title, _titleFont, new Point(12, 10), Theme.Text, TextFormatFlags.NoPadding);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) _titleFont.Dispose();
            base.Dispose(disposing);
        }
        internal static GraphicsPath Rounded(Rectangle r, int radius)
        {
            int d = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    public sealed class DocLinkPaneControl : UserControl
    {
        private const string NoneLabel = "(none - use file name)";
        private const int FieldWidth = 330;

        // Excel
        private readonly TextBox _txtFile = new TextBox();
        private readonly ComboBox _cmbSheet = new ComboBox();
        private readonly NumericUpDown _numHeader = new NumericUpDown();
        private readonly Label _lblLoaded = new Label();
        // Columns
        private readonly ComboBox _cmbKey = new ComboBox();
        private readonly ComboBox _cmbLabel = new ComboBox();
        private readonly CheckedListBox _lstDocCols = new CheckedListBox();
        // Doc folder
        private readonly TextBox _txtDocFolder = new TextBox();
        private readonly TextBox _txtPattern = new TextBox();
        private readonly CheckBox _chkSubfolders = new CheckBox();
        private readonly CheckBox _chkFolderOnlyEmpty = new CheckBox();
        private readonly CheckBox _chkLoose = new CheckBox();
        // Match
        private readonly ComboBox _cmbProp = new ComboBox();
        private readonly ComboBox _cmbMode = new ComboBox();
        private readonly CheckBox _chkTypePrefix = new CheckBox();
        private readonly CheckBox _chkIgnoreSep = new CheckBox();
        private readonly CheckBox _chkTopMost = new CheckBox();
        private readonly RadioButton _rbEntire = new RadioButton();
        private readonly RadioButton _rbSelection = new RadioButton();
        // Preview
        private readonly DataGridView _grid = new DataGridView();
        private readonly Label _lblSummary = new Label();
        // Apply
        private readonly RadioButton _rbReplace = new RadioButton();
        private readonly RadioButton _rbAppend = new RadioButton();
        private readonly CheckBox _chkSets = new CheckBox();
        private readonly CheckBox _chkSkipMissing = new CheckBox();
        // Status
        private readonly TextBox _txtStatus = new TextBox();
        private readonly ProgressBar _progress = new ProgressBar();

        private DataTable _table;
        private List<PlanRow> _plan;
        private Dictionary<PlanRow, List<Nw.ModelItem>> _rowItems;
        private Nw.Document _planDoc;
        private bool _cancel;

        public DocLinkPaneControl()
        {
            Dock = DockStyle.Fill;
            BackColor = Theme.Surface;
            AutoScroll = true;
            Padding = new Padding(10);

            var root = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, BackColor = Theme.Surface };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var header = new Panel { Height = 52, Dock = DockStyle.Top, BackColor = Theme.Accent };
            header.Controls.Add(new Label { Text = Names.Product, ForeColor = Color.White, Font = Theme.Semibold(12F), AutoSize = true, Location = new Point(12, 8) });
            header.Controls.Add(new Label { Text = "Excel list → document hyperlinks on model items", ForeColor = Color.FromArgb(220, 232, 245), Font = Theme.Ui(8.25F), AutoSize = true, Location = new Point(12, 30) });
            root.Controls.Add(header);

            // 1. Excel
            var excel = Flow();
            _txtFile.Width = FieldWidth; _txtFile.ReadOnly = true; _txtFile.Font = Theme.Ui(8.5F); _txtFile.BackColor = Color.White;
            var btnBrowse = MakeButton("Choose Excel file…", true);
            btnBrowse.Click += (s, e) => BrowseExcel();
            Combo(_cmbSheet);
            _cmbSheet.SelectedIndexChanged += (s, e) => { if (_cmbSheet.Focused) LoadSheet(); };
            _numHeader.Minimum = 0; _numHeader.Maximum = 1000; _numHeader.Width = 60; _numHeader.Font = Theme.Ui(9F);
            var headerRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            headerRow.Controls.Add(Hint("Header row (0 = auto):", 6));
            headerRow.Controls.Add(_numHeader);
            var btnReload = MakeButton("Reload", false);
            btnReload.MinimumSize = new Size(80, 26);
            btnReload.Margin = new Padding(8, 0, 0, 0);
            btnReload.Click += (s, e) => LoadSheet();
            headerRow.Controls.Add(btnReload);
            _lblLoaded.AutoSize = true; _lblLoaded.Font = Theme.Ui(8F); _lblLoaded.ForeColor = Theme.Subtle; _lblLoaded.MaximumSize = new Size(FieldWidth, 0);
            excel.Controls.Add(btnBrowse);
            excel.Controls.Add(_txtFile);
            excel.Controls.Add(Hint("Sheet:", 4));
            excel.Controls.Add(_cmbSheet);
            excel.Controls.Add(headerRow);
            excel.Controls.Add(_lblLoaded);
            root.Controls.Add(new Card("1. Excel", excel));

            // 2. Columns
            var cols = Flow();
            Combo(_cmbKey); Combo(_cmbLabel);
            _lstDocCols.Width = FieldWidth; _lstDocCols.Height = 96; _lstDocCols.CheckOnClick = true; _lstDocCols.Font = Theme.Ui(8.5F);
            cols.Controls.Add(Hint("KEY column (line no. / tag that matches the model):", 0));
            cols.Controls.Add(_cmbKey);
            cols.Controls.Add(Hint("DOCUMENT column(s) - file path, folder or URL (; for several):", 6));
            cols.Controls.Add(_lstDocCols);
            cols.Controls.Add(Hint("LINK NAME column (optional):", 6));
            cols.Controls.Add(_cmbLabel);
            root.Controls.Add(new Card("2. Columns", cols));

            // 3. Doc folder
            var folder = Flow();
            _txtDocFolder.Width = FieldWidth - 90; _txtDocFolder.Font = Theme.Ui(8.5F);
            var folderRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            folderRow.Controls.Add(_txtDocFolder);
            var btnFolder = MakeButton("Browse…", false);
            btnFolder.MinimumSize = new Size(80, 24); btnFolder.Margin = new Padding(6, 0, 0, 0);
            btnFolder.Click += (s, e) => BrowseFolder();
            folderRow.Controls.Add(btnFolder);
            _txtPattern.Width = 160; _txtPattern.Text = "*{KEY}*.pdf"; _txtPattern.Font = Theme.Ui(9F);
            var patRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            patRow.Controls.Add(Hint("File name pattern:", 6));
            patRow.Controls.Add(_txtPattern);
            Check(_chkSubfolders, "Include subfolders", true);
            Check(_chkFolderOnlyEmpty, "Use folder only when the row's document cells are empty", true);
            Check(_chkLoose, "Loose file names (ignore - _ . and spaces)", false);
            folder.Controls.Add(Hint("Optional. Find files for each KEY in this folder:", 0));
            folder.Controls.Add(folderRow);
            folder.Controls.Add(patRow);
            folder.Controls.Add(_chkSubfolders);
            folder.Controls.Add(_chkFolderOnlyEmpty);
            folder.Controls.Add(_chkLoose);
            folder.Controls.Add(Hint("{KEY} = the row's key, * = anything. E.g. *{KEY}*.pdf or ISO_{KEY}_*.pdf", 2));
            root.Controls.Add(new Card("3. Document folder (optional)", folder));

            // 4. Match
            var match = Flow();
            Combo(_cmbProp);
            _cmbProp.Items.Add(PropRef.ItemName);
            _cmbProp.SelectedIndex = 0;
            var btnProps = MakeButton("Read properties from selected item", false);
            btnProps.Click += (s, e) => ReadProperties();
            Combo(_cmbMode);
            _cmbMode.Items.AddRange(new object[] {
                "Exact (after cleaning)", "Contains key (whole part)", "Contains key (anywhere)", "Starts with key" });
            _cmbMode.SelectedIndex = 0;
            Check(_chkTypePrefix, "Ignore type word before '/' (e.g. PIPE /0-A-1)", true);
            Check(_chkIgnoreSep, "Ignore separators (- _ . / spaces)", false);
            Check(_chkTopMost, "Link top-most match only (skip children of a matched item)", true);
            _rbEntire.Text = "Whole model"; _rbEntire.AutoSize = true; _rbEntire.Checked = true; _rbEntire.Font = Theme.Ui(9F);
            _rbSelection.Text = "Current selection (and its children)"; _rbSelection.AutoSize = true; _rbSelection.Font = Theme.Ui(9F);
            var scope = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = true, MaximumSize = new Size(FieldWidth + 20, 0) };
            scope.Controls.Add(_rbEntire);
            scope.Controls.Add(_rbSelection);
            match.Controls.Add(Hint("Model property to compare:", 0));
            match.Controls.Add(_cmbProp);
            match.Controls.Add(btnProps);
            match.Controls.Add(Hint("Match mode (case-insensitive, leading '/' ignored):", 8));
            match.Controls.Add(_cmbMode);
            match.Controls.Add(_chkTypePrefix);
            match.Controls.Add(_chkIgnoreSep);
            match.Controls.Add(_chkTopMost);
            match.Controls.Add(Hint("Search in:", 6));
            match.Controls.Add(scope);
            root.Controls.Add(new Card("4. Match to model", match));

            // 5. Preview
            var prev = Flow();
            var btnPreview = MakeButton("Preview matches", true);
            btnPreview.Click += (s, e) => Preview(true);
            SetupGrid();
            _lblSummary.AutoSize = true; _lblSummary.Font = Theme.Ui(8.5F); _lblSummary.ForeColor = Theme.Text; _lblSummary.MaximumSize = new Size(FieldWidth, 0);
            prev.Controls.Add(btnPreview);
            prev.Controls.Add(_lblSummary);
            prev.Controls.Add(_grid);
            prev.Controls.Add(Hint("Double-click a row to select its items in the model.", 2));
            var btnReport = MakeButton("Export unmatched report…", false);
            btnReport.Click += (s, e) => ExportReport();
            prev.Controls.Add(btnReport);
            root.Controls.Add(new Card("5. Preview", prev));

            // 6. Apply
            var apply = Flow();
            _rbReplace.Text = "Replace my earlier MD links"; _rbReplace.AutoSize = true; _rbReplace.Checked = true; _rbReplace.Font = Theme.Ui(9F);
            _rbAppend.Text = "Append (keep earlier MD links)"; _rbAppend.AutoSize = true; _rbAppend.Font = Theme.Ui(9F);
            Check(_chkSkipMissing, "Skip documents that are not found on disk", true);
            Check(_chkSets, "Create selection sets (folder \"" + Names.SetFolder + "\")", true);
            var btnApply = MakeButton("Add links to model", true);
            btnApply.Click += (s, e) => ApplyLinks();
            apply.Controls.Add(_rbReplace);
            apply.Controls.Add(_rbAppend);
            apply.Controls.Add(_chkSkipMissing);
            apply.Controls.Add(_chkSets);
            apply.Controls.Add(btnApply);
            apply.Controls.Add(Hint("Links are named \"" + Names.LinkPrefix + "…\". Other hyperlinks are kept.\nSave the NWF / NWD afterwards - links are kept only in the saved file.", 4));
            root.Controls.Add(new Card("6. Add links", apply));

            // 7. Tools
            var tools = Flow();
            var btnOpen = MakeButton("Open linked docs (selected item)", false);
            btnOpen.Click += (s, e) => OpenLinkedDocs();
            var btnRemove = MakeButton("Remove MD links…", false);
            btnRemove.Click += (s, e) => RemoveLinks();
            tools.Controls.Add(btnOpen);
            tools.Controls.Add(btnRemove);
            tools.Controls.Add(Hint("Remove works on the selection (with children), or the whole model when nothing is selected. Only \"" + Names.LinkPrefix + "\" links are removed.", 4));
            root.Controls.Add(new Card("7. Tools", tools));

            // Status
            var statusBody = new Panel { Height = 130, Dock = DockStyle.Top };
            _progress.Dock = DockStyle.Top; _progress.Height = 14; _progress.Style = ProgressBarStyle.Continuous;
            _txtStatus.Multiline = true; _txtStatus.ReadOnly = true; _txtStatus.ScrollBars = ScrollBars.Vertical;
            _txtStatus.Dock = DockStyle.Fill; _txtStatus.Font = Theme.Ui(8.25F); _txtStatus.BackColor = Color.White;
            _txtStatus.Text = "Ready. Choose an Excel file to start. (Esc cancels a long run.)";
            statusBody.Controls.Add(_txtStatus);
            statusBody.Controls.Add(_progress);
            root.Controls.Add(new Card("Status", statusBody));

            root.Controls.Add(new Label { Text = Names.Author + "  ·  v" + Names.Version, Font = Theme.Ui(8F), ForeColor = Theme.Subtle, AutoSize = true, Padding = new Padding(4, 4, 4, 12) });
            Controls.Add(root);
        }

        // ---------- small UI helpers ----------

        private static FlowLayoutPanel Flow()
        {
            return new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Dock = DockStyle.Top };
        }

        private static Label Hint(string text, int top)
        {
            return new Label { Text = text, AutoSize = true, Font = Theme.Ui(8.25F), ForeColor = Theme.Subtle, MaximumSize = new Size(FieldWidth, 0), Margin = new Padding(0, top, 0, 2) };
        }

        private static void Combo(ComboBox c)
        {
            c.DropDownStyle = ComboBoxStyle.DropDownList;
            c.Width = FieldWidth;
            c.Font = Theme.Ui(9F);
            c.DropDownWidth = 480;
            c.MaxDropDownItems = 20;
        }

        private static void Check(CheckBox c, string text, bool on)
        {
            c.Text = text; c.AutoSize = true; c.Checked = on; c.Font = Theme.Ui(8.75F);
            c.MaximumSize = new Size(FieldWidth + 20, 0);
        }

        private static Button MakeButton(string text, bool primary)
        {
            var b = new Button { Text = text, AutoSize = true, MinimumSize = new Size(220, 30), Font = Theme.Ui(9F), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0, 6, 0, 0) };
            b.FlatAppearance.BorderSize = 1;
            if (primary) { b.BackColor = Theme.Accent; b.ForeColor = Color.White; b.FlatAppearance.BorderColor = Theme.Accent; }
            else { b.BackColor = Color.White; b.ForeColor = Theme.Text; b.FlatAppearance.BorderColor = Theme.CardBorder; }
            return b;
        }

        private void SetupGrid()
        {
            _grid.Width = FieldWidth + 10;
            _grid.Height = 260;
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.AllowUserToResizeRows = false;
            _grid.RowHeadersVisible = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.BackgroundColor = Color.White;
            _grid.BorderStyle = BorderStyle.FixedSingle;
            _grid.Font = Theme.Ui(8.25F);
            _grid.RowTemplate.Height = 20;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _grid.Columns.Add("row", "Row");
            _grid.Columns.Add("key", "Key");
            _grid.Columns.Add("doc", "Document(s)");
            _grid.Columns.Add("items", "Items");
            _grid.Columns.Add("status", "Status");
            _grid.Columns[0].Width = 42;
            _grid.Columns[1].Width = 120;
            _grid.Columns[2].Width = 160;
            _grid.Columns[3].Width = 44;
            _grid.Columns[4].Width = 110;
            _grid.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) SelectRowItems(e.RowIndex); };
        }

        private Nw.Document Doc() { return Nw.Application.ActiveDocument; }

        private void SetStatus(string text)
        {
            _txtStatus.Text = text ?? "";
            Diagnostics.Log((text ?? "").Replace("\r\n", " | "));
        }

        private bool SetProgress(double fraction)
        {
            int v = (int)Math.Max(0, Math.Min(100, fraction * 100));
            if (_progress.Value != v) _progress.Value = v;
            Application.DoEvents();
            return !_cancel;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape) { _cancel = true; return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void Fail(string what, Exception ex)
        {
            Diagnostics.Log(what + ": " + ex);
            SetStatus(what + " failed: " + ex.Message + "\r\nLog: " + Diagnostics.LogPath);
            MessageBox.Show(what + " failed:\r\n\r\n" + ex.Message, Names.Product, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void Info(string msg)
        {
            MessageBox.Show(msg, Names.Product, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Busy(bool on)
        {
            Cursor = on ? Cursors.WaitCursor : Cursors.Default;
            if (on) { _cancel = false; _progress.Value = 0; }
        }

        // ---------- 1. Excel ----------

        private void BrowseExcel()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel / CSV (*.xlsx;*.xlsm;*.csv)|*.xlsx;*.xlsm;*.csv|All files (*.*)|*.*";
                dlg.Title = "Excel with keys and documents";
                if (File.Exists(_txtFile.Text)) dlg.FileName = _txtFile.Text;
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                _txtFile.Text = dlg.FileName;
            }
            try
            {
                _cmbSheet.Items.Clear();
                if (!_txtFile.Text.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    foreach (string s in XlsxReader.ListSheets(_txtFile.Text)) _cmbSheet.Items.Add(s);
                if (_cmbSheet.Items.Count > 0) _cmbSheet.SelectedIndex = 0;
                _numHeader.Value = 0;
                LoadSheet();
            }
            catch (Exception ex) { Fail("Open Excel", ex); }
        }

        private void LoadSheet()
        {
            string path = _txtFile.Text;
            if (!File.Exists(path)) { SetStatus("Choose an Excel file first."); return; }
            try
            {
                Busy(true);
                SetStatus("Reading " + Path.GetFileName(path) + "…");
                Application.DoEvents();
                if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    _table = XlsxReader.LoadCsv(path);
                else
                    _table = XlsxReader.Load(path, _cmbSheet.SelectedItem as string, (int)_numHeader.Value);
                ClearPlan();
                FillColumns();
                _lblLoaded.Text = string.Format("{0} data rows, {1} columns. Header on row {2} of \"{3}\".",
                    _table.Rows.Count, _table.Columns.Count, _table.HeaderRow, _table.SheetName);
                SetStatus("Loaded " + Path.GetFileName(path) + "\r\n" + _lblLoaded.Text +
                          "\r\nCheck the KEY and DOCUMENT columns, then Preview.");
            }
            catch (Exception ex) { Fail("Read Excel", ex); }
            finally { Busy(false); }
        }

        private void FillColumns()
        {
            string oldKey = _cmbKey.SelectedItem as string;
            _cmbKey.Items.Clear(); _cmbLabel.Items.Clear(); _lstDocCols.Items.Clear();
            _cmbLabel.Items.Add(NoneLabel);
            if (_table == null) return;
            foreach (string c in _table.Columns)
            {
                _cmbKey.Items.Add(c);
                _cmbLabel.Items.Add(c);
                bool docLike = LooksLikeDocColumn(c) || ColumnHasLinks(c);
                _lstDocCols.Items.Add(c, docLike);
            }
            int key = -1;
            if (oldKey != null) key = _cmbKey.Items.IndexOf(oldKey);
            if (key < 0) key = IndexOfColumn(new[] { "line no", "line number", "line tag", "tag", "key", "line" });
            _cmbKey.SelectedIndex = key >= 0 ? key : (_cmbKey.Items.Count > 0 ? 0 : -1);
            _cmbLabel.SelectedIndex = 0;
        }

        private int IndexOfColumn(string[] hints)
        {
            foreach (string h in hints)
                for (int i = 0; i < _cmbKey.Items.Count; i++)
                    if (_cmbKey.Items[i].ToString().IndexOf(h, StringComparison.OrdinalIgnoreCase) >= 0) return i;
            return -1;
        }

        private static bool LooksLikeDocColumn(string c)
        {
            foreach (string h in new[] { "document", "doc ", "doc.", "docs", "pdf", "drawing", "dwg", "hyperlink", "link", "file", "path", "url", "isometric" })
                if (c.IndexOf(h, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return string.Equals(c, "doc", StringComparison.OrdinalIgnoreCase) || string.Equals(c, "iso", StringComparison.OrdinalIgnoreCase);
        }

        private bool ColumnHasLinks(string c)
        {
            int n = 0;
            foreach (TableRow r in _table.Rows)
            {
                if (r.Links.ContainsKey(c)) return true;
                if (++n > 200) break;
            }
            return false;
        }

        // ---------- 3. folder ----------

        private void BrowseFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Folder that holds the documents (PDF, DWG…)";
                if (Directory.Exists(_txtDocFolder.Text)) dlg.SelectedPath = _txtDocFolder.Text;
                if (dlg.ShowDialog(this) == DialogResult.OK) _txtDocFolder.Text = dlg.SelectedPath;
            }
        }

        // ---------- 4. match ----------

        private void ReadProperties()
        {
            try
            {
                Nw.Document doc = Doc();
                if (doc == null || doc.CurrentSelection.SelectedItems.Count == 0)
                {
                    Info("Select one model item that carries the key value (e.g. a pipe), then click again.");
                    return;
                }
                Nw.ModelItem item = doc.CurrentSelection.SelectedItems.First;
                List<PropRef> props = NavisLinks.PropertiesOf(item);
                _cmbProp.Items.Clear();
                _cmbProp.Items.Add(PropRef.ItemName);
                foreach (PropRef p in props) _cmbProp.Items.Add(p);
                _cmbProp.SelectedIndex = 0;
                SetStatus("Read " + props.Count + " properties from \"" + NavisLinks.SafeName(item) + "\". Pick the one that holds the key.");
                _cmbProp.DroppedDown = true;
            }
            catch (Exception ex) { Fail("Read properties", ex); }
        }

        private MatchOptions Options()
        {
            var o = new MatchOptions
            {
                StripLeadingSlash = true,
                StripTypePrefix = _chkTypePrefix.Checked,
                IgnoreSeparators = _chkIgnoreSep.Checked
            };
            switch (_cmbMode.SelectedIndex)
            {
                case 1: o.Mode = MatchMode.ContainsWhole; break;
                case 2: o.Mode = MatchMode.Contains; break;
                case 3: o.Mode = MatchMode.StartsWith; break;
                default: o.Mode = MatchMode.Exact; break;
            }
            return o;
        }

        private DocOptions DocOpts()
        {
            string excelFolder = null;
            try { excelFolder = Path.GetDirectoryName(_txtFile.Text); } catch { }
            return new DocOptions
            {
                BaseFolder = excelFolder,
                SearchFolder = _txtDocFolder.Text.Trim().Length > 0 ? _txtDocFolder.Text.Trim() : null,
                Pattern = _txtPattern.Text.Trim().Length > 0 ? _txtPattern.Text.Trim() : "*{KEY}*",
                Recursive = _chkSubfolders.Checked,
                FolderOnlyWhenCellsEmpty = _chkFolderOnlyEmpty.Checked,
                LooseFileNames = _chkLoose.Checked,
                SearchCellFolders = true
            };
        }

        // ---------- 5. preview ----------

        private void ClearPlan()
        {
            _plan = null; _rowItems = null; _planDoc = null;
            _grid.Rows.Clear();
            _lblSummary.Text = "";
        }

        private bool Preview(bool showGrid)
        {
            Nw.Document doc = Doc();
            if (doc == null || doc.Models.Count == 0) { Info("Open a model first."); return false; }
            if (_table == null) { Info("Choose an Excel file first (step 1)."); return false; }
            string keyCol = _cmbKey.SelectedItem as string;
            if (keyCol == null) { Info("Pick the KEY column (step 2)."); return false; }
            var docCols = _lstDocCols.CheckedItems.Cast<object>().Select(o => o.ToString()).ToList();
            if (docCols.Count == 0 && _txtDocFolder.Text.Trim().Length == 0)
            {
                Info("Tick at least one DOCUMENT column (step 2), or set a document folder (step 3).");
                return false;
            }
            if (_txtDocFolder.Text.Trim().Length > 0 && !Directory.Exists(_txtDocFolder.Text.Trim()))
            {
                Info("Document folder not found:\r\n" + _txtDocFolder.Text);
                return false;
            }
            bool selectionOnly = _rbSelection.Checked;
            if (selectionOnly && doc.CurrentSelection.SelectedItems.Count == 0)
            {
                Info("Nothing is selected. Select items or choose Whole model.");
                return false;
            }
            try
            {
                Busy(true);
                ClearPlan();
                SetStatus("Checking documents…");
                var settings = new PlanSettings
                {
                    KeyColumn = keyCol,
                    DocColumns = docCols,
                    LabelColumn = (_cmbLabel.SelectedItem as string) == NoneLabel ? null : _cmbLabel.SelectedItem as string
                };
                var resolver = new DocResolver(DocOpts(), new RealFileSystem());
                List<PlanRow> plan = LinkPlanner.Build(_table, settings, resolver);
                SetProgress(0.05);

                var prop = _cmbProp.SelectedItem as PropRef ?? PropRef.ItemName;
                SetStatus("Reading \"" + prop + "\" from the model… (Esc to cancel)");
                int walked;
                KeyIndex<Nw.ModelItem> idx = NavisLinks.BuildIndex(doc, selectionOnly, prop, Options(), SetProgress, out walked);

                SetStatus("Matching " + plan.Count + " rows against " + idx.DistinctNames + " names…");
                var rowItems = new Dictionary<PlanRow, List<Nw.ModelItem>>();
                int i = 0;
                foreach (PlanRow p in plan)
                {
                    List<Nw.ModelItem> items = p.HasKey ? NavisLinks.FindTopMost(idx, p.Key, _chkTopMost.Checked) : new List<Nw.ModelItem>();
                    p.Finish(items.Count);
                    rowItems[p] = items;
                    if ((++i & 0x3F) == 0 && !SetProgress(0.9 + 0.09 * i / Math.Max(1, plan.Count))) throw new OperationCanceledException();
                }
                _plan = plan; _rowItems = rowItems; _planDoc = doc;
                if (showGrid) FillGrid();
                SetProgress(1);
                string summary = Summary(walked, idx);
                _lblSummary.Text = summary;
                SetStatus("Preview ready.\r\n" + summary + "\r\nMatch: " + prop + ", " + Options().Describe());
                return true;
            }
            catch (OperationCanceledException) { SetStatus("Cancelled."); return false; }
            catch (Exception ex) { Fail("Preview", ex); return false; }
            finally { Busy(false); }
        }

        private string Summary(int walked, KeyIndex<Nw.ModelItem> idx)
        {
            int Count(RowStatus s) { return _plan.Count(p => p.Status == s); }
            int linkedItems = _rowItems.Where(kv => kv.Key.Status == RowStatus.Matched || kv.Key.Status == RowStatus.DocMissing)
                .SelectMany(kv => kv.Value).Distinct().Count();
            return string.Format("{0} rows: {1} matched, {2} doc missing, {3} not found in model, {4} no document, {5} no key.\r\n{6} model items would get links ({7} items searched).",
                _plan.Count, Count(RowStatus.Matched), Count(RowStatus.DocMissing), Count(RowStatus.NotFound),
                Count(RowStatus.NoDocument), Count(RowStatus.NoKey), linkedItems, walked);
        }

        private void FillGrid()
        {
            _grid.SuspendLayout();
            _grid.Rows.Clear();
            var rows = new List<DataGridViewRow>();
            foreach (PlanRow p in _plan)
            {
                var r = new DataGridViewRow();
                r.CreateCells(_grid, p.ExcelRow, p.Key, p.DocsText, p.MatchedItems, PlanRow.StatusText(p.Status));
                r.Tag = p;
                Color c = StatusColor(p.Status);
                if (c != Color.Empty) r.Cells[4].Style.BackColor = c;
                r.Cells[2].ToolTipText = string.Join("\r\n", p.Docs.Select(d => (d.Exists ? "" : "[missing] ") + d.Target).ToArray());
                rows.Add(r);
            }
            _grid.Rows.AddRange(rows.ToArray());
            _grid.ResumeLayout();
        }

        private static Color StatusColor(RowStatus s)
        {
            switch (s)
            {
                case RowStatus.Matched: return Color.FromArgb(214, 240, 220);
                case RowStatus.DocMissing: return Color.FromArgb(255, 236, 200);
                case RowStatus.NotFound: return Color.FromArgb(250, 215, 215);
                case RowStatus.NoDocument:
                case RowStatus.NoKey: return Color.FromArgb(232, 232, 232);
                default: return Color.Empty;
            }
        }

        private bool PlanIsCurrent()
        {
            return _plan != null && _rowItems != null && _planDoc != null && ReferenceEquals(_planDoc, Doc());
        }

        private void SelectRowItems(int gridRow)
        {
            try
            {
                var p = _grid.Rows[gridRow].Tag as PlanRow;
                if (p == null || !PlanIsCurrent()) return;
                List<Nw.ModelItem> items;
                if (!_rowItems.TryGetValue(p, out items) || items.Count == 0)
                {
                    SetStatus("Row " + p.ExcelRow + " (" + p.Key + "): no model items. " + PlanRow.StatusText(p.Status));
                    return;
                }
                var coll = new Nw.ModelItemCollection();
                foreach (Nw.ModelItem m in items) coll.Add(m);
                Doc().CurrentSelection.CopyFrom(coll);
                SetStatus("Row " + p.ExcelRow + " (" + p.Key + "): selected " + TextUtil.Plural(items.Count, "item", "items") +
                          ".\r\n" + string.Join("\r\n", p.Docs.Select(d => (d.Exists ? "" : "[missing] ") + d.Target).ToArray()));
            }
            catch (Exception ex) { Fail("Select items", ex); }
        }

        private void ExportReport()
        {
            if (_plan == null) { Info("Run Preview first."); return; }
            try
            {
                List<PlanRow> rows = _plan.Where(p => p.Status != RowStatus.Matched).ToList();
                bool all = rows.Count == 0;
                if (all) rows = _plan;
                using (var dlg = new SaveFileDialog())
                {
                    dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                    dlg.Title = "Save unmatched report";
                    dlg.FileName = "MD_DocLink_Unmatched_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx";
                    if (dlg.ShowDialog(this) != DialogResult.OK) return;
                    XlsxWriter.Save(dlg.FileName, LinkPlanner.ReportTable(rows, all ? "All rows" : "Unmatched"), Names.Author);
                    SetStatus((all ? "All rows matched - saved the full list (" : "Saved " ) + rows.Count + " rows" + (all ? ")" : " that need attention") + ".\r\n" + dlg.FileName);
                }
            }
            catch (Exception ex) { Fail("Export report", ex); }
        }

        // ---------- 6. apply ----------

        private void ApplyLinks()
        {
            Nw.Document doc = Doc();
            if (doc == null) { Info("Open a model first."); return; }
            if (!PlanIsCurrent() && !Preview(true)) return;
            try
            {
                bool includeMissing = !_chkSkipMissing.Checked;
                var usable = _rowItems.Where(kv => kv.Value.Count > 0 &&
                    (kv.Key.Status == RowStatus.Matched || kv.Key.Status == RowStatus.DocMissing))
                    .Select(kv => new KeyValuePair<PlanRow, IList<Nw.ModelItem>>(kv.Key, kv.Value)).ToList();
                Dictionary<Nw.ModelItem, List<KeyValuePair<string, string>>> perItem = LinkPlanner.MergePerItem(usable, includeMissing);
                int links = perItem.Values.Sum(v => v.Count);
                if (perItem.Count == 0)
                {
                    Info("Nothing to link. Check the preview: rows need a model match and a document" + (includeMissing ? "." : " that exists on disk."));
                    return;
                }
                string mode = _rbReplace.Checked ? "Replace earlier \"" + Names.LinkPrefix + "\" links" : "Append to existing links";
                if (MessageBox.Show(string.Format("Add {0} to {1}?\r\n\r\n{2}. Other hyperlinks are kept.\r\n\r\nSave the NWF/NWD afterwards to keep the links.",
                        TextUtil.Plural(links, "link", "links"), TextUtil.Plural(perItem.Count, "model item", "model items"), mode),
                        Names.Product, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;

                Busy(true);
                SetStatus("Writing links… (Esc to cancel)");
                ApplyStats st = NavisLinks.Apply(perItem, _rbReplace.Checked, SetProgress);

                string sets = "";
                if (_chkSets.Checked)
                {
                    try
                    {
                        var linked = new Nw.ModelItemCollection();
                        foreach (Nw.ModelItem m in perItem.Keys) linked.Add(m);
                        var missing = new Nw.ModelItemCollection();
                        foreach (Nw.ModelItem m in _rowItems.Where(kv => kv.Key.Status == RowStatus.DocMissing).SelectMany(kv => kv.Value).Distinct())
                            missing.Add(m);
                        NavisLinks.ReplaceSet(doc, Names.SetLinked, linked);
                        NavisLinks.ReplaceSet(doc, Names.SetDocMissing, missing);
                        sets = "\r\nSelection sets updated in \"" + Names.SetFolder + "\" (Linked: " + linked.Count + ", Doc missing: " + missing.Count + ").";
                    }
                    catch (Exception ex) { sets = "\r\nSelection sets failed: " + ex.Message; Diagnostics.Log("Sets: " + ex); }
                }
                SetProgress(1);
                string msg = string.Format("Done. {0} got links ({1} new links, {2} batch writes, {3} merged with existing links).{4}{5}\r\n\r\nSave the NWF / NWD to keep the links.",
                    TextUtil.Plural(st.ItemsLinked, "item", "items"), st.LinksWritten, st.Batches, st.ItemsMerged,
                    st.Errors > 0 ? "\r\nErrors: " + st.Errors + "\r\n" + string.Join("\r\n", st.Messages.ToArray()) : "", sets);
                SetStatus(msg);
                MessageBox.Show(msg, Names.Product, MessageBoxButtons.OK, st.Errors > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (Exception ex) { Fail("Add links", ex); }
            finally { Busy(false); }
        }

        // ---------- 7. tools ----------

        private void OpenLinkedDocs()
        {
            try
            {
                Nw.Document doc = Doc();
                if (doc == null || doc.CurrentSelection.SelectedItems.Count == 0) { Info("Select a model item first."); return; }
                Nw.ModelItem item = doc.CurrentSelection.SelectedItems.First;
                List<KeyValuePair<string, NwLink>> links = NavisLinks.LinksForOpen(item);
                if (links.Count == 0) { Info("No hyperlinks on \"" + NavisLinks.SafeName(item) + "\" or its parents."); return; }
                if (links.Count == 1) { OpenTarget(links[0].Value.Url); return; }
                using (var f = new LinkPicker(links))
                    if (f.ShowDialog(this) == DialogResult.OK)
                        foreach (string u in f.Chosen) OpenTarget(u);
            }
            catch (Exception ex) { Fail("Open linked docs", ex); }
        }

        internal static void OpenTarget(string target)
        {
            string t = (target ?? "").Trim();
            if (t.Length == 0) return;
            bool url = DocResolver.LooksLikeUrl(t);
            if (!url && !File.Exists(t) && !Directory.Exists(t))
            {
                MessageBox.Show("File not found:\r\n" + t, Names.Product, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (url && t.StartsWith("www.", StringComparison.OrdinalIgnoreCase)) t = "http://" + t;
            Diagnostics.Log("Open " + t);
            Process.Start(new ProcessStartInfo(t) { UseShellExecute = true });
        }

        private void RemoveLinks()
        {
            Nw.Document doc = Doc();
            if (doc == null || doc.Models.Count == 0) { Info("Open a model first."); return; }
            bool sel = doc.CurrentSelection.SelectedItems.Count > 0;
            string where = sel ? "the selected items (and their children)" : "the WHOLE model (nothing is selected)";
            if (MessageBox.Show("Remove all \"" + Names.LinkPrefix + "…\" hyperlinks from " + where + "?\r\n\r\nOther hyperlinks stay.",
                    Names.Product, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK) return;
            try
            {
                Busy(true);
                SetStatus("Removing MD links… (Esc to cancel)");
                int removed, errors;
                int items = NavisLinks.RemoveMdLinks(NavisLinks.EnumerateScope(doc, sel), SetProgress, out removed, out errors);
                SetProgress(1);
                string msg = "Removed " + TextUtil.Plural(removed, "MD link", "MD links") + " from " + TextUtil.Plural(items, "item", "items") + "." +
                             (errors > 0 ? " Errors: " + errors + " (see log)." : "") + (_cancel ? " (cancelled part-way)" : "") +
                             "\r\nSave the NWF / NWD to keep the change.";
                SetStatus(msg);
                Info(msg);
            }
            catch (Exception ex) { Fail("Remove MD links", ex); }
            finally { Busy(false); }
        }
    }

    /// <summary>Pick one or more hyperlinks to open.</summary>
    internal sealed class LinkPicker : Form
    {
        private readonly ListView _list = new ListView();
        public readonly List<string> Chosen = new List<string>();

        public LinkPicker(List<KeyValuePair<string, NwLink>> links)
        {
            Text = Names.Product + " - open linked documents";
            Width = 640; Height = 360;
            StartPosition = FormStartPosition.CenterParent;
            Font = Theme.Ui(9F);
            MinimizeBox = false; MaximizeBox = false; ShowInTaskbar = false;
            _list.View = View.Details; _list.FullRowSelect = true; _list.CheckBoxes = true; _list.Dock = DockStyle.Fill;
            _list.Columns.Add("Name", 200); _list.Columns.Add("Target", 300); _list.Columns.Add("On item", 110);
            foreach (KeyValuePair<string, NwLink> kv in links)
            {
                var li = new ListViewItem(new[] { kv.Value.Name, kv.Value.Url, kv.Key }) { Tag = kv.Value.Url };
                _list.Items.Add(li);
            }
            _list.ItemActivate += (s, e) => { foreach (ListViewItem li in _list.SelectedItems) Chosen.Add((string)li.Tag); DialogResult = DialogResult.OK; };
            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 40, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(6) };
            var cancel = new Button { Text = "Close", DialogResult = DialogResult.Cancel, AutoSize = true };
            var open = new Button { Text = "Open ticked", AutoSize = true };
            open.Click += (s, e) =>
            {
                foreach (ListViewItem li in _list.CheckedItems) Chosen.Add((string)li.Tag);
                if (Chosen.Count == 0) foreach (ListViewItem li in _list.SelectedItems) Chosen.Add((string)li.Tag);
                if (Chosen.Count > 0) DialogResult = DialogResult.OK;
            };
            bar.Controls.Add(cancel); bar.Controls.Add(open);
            Controls.Add(_list); Controls.Add(bar);
            Controls.Add(new Label { Text = "Double-click a link to open it, or tick several and click Open ticked.", Dock = DockStyle.Top, Height = 24, Padding = new Padding(6, 6, 0, 0) });
            AcceptButton = open; CancelButton = cancel;
        }
    }
}
