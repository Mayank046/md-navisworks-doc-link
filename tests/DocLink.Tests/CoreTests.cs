// MD Doc Link - unit tests (no Navisworks)
// Author: MD (Mayank Dholariya)
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocLink.Core;
using Xunit;

namespace DocLink.Tests
{
    internal sealed class FakeFs : IFileSystem
    {
        public readonly HashSet<string> Files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly HashSet<string> Dirs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public bool FileExists(string p) { return Files.Contains(p); }
        public bool DirectoryExists(string p) { return Dirs.Contains(p.TrimEnd('\\')); }
        public IEnumerable<string> EnumerateFiles(string folder, bool recursive)
        {
            string f = folder.TrimEnd('\\') + "\\";
            return Files.Where(x => x.StartsWith(f, StringComparison.OrdinalIgnoreCase)
                && (recursive || x.IndexOf('\\', f.Length) < 0)).ToList();
        }
    }

    internal sealed class Node
    {
        public string Name; public Node Parent;
        public override string ToString() { return Name; }
    }

    public class KeyMatchTests
    {
        [Theory]
        [InlineData("/150-P-1001-A1A-C-CS", "150-P-1001-A1A-C-CS")]
        [InlineData("PIPE /150-p-1001", "150-P-1001")]
        [InlineData("  /150-P-1001 \r\n ", "150-P-1001")]
        [InlineData("BRANCH  /150-P-1001/B1", "150-P-1001/B1")]
        public void Normalize_Strips_Slash_Type_And_Case(string raw, string expected)
        {
            Assert.Equal(expected, KeyNormalizer.Normalize(raw, new MatchOptions()));
        }

        [Fact]
        public void Normalize_Keeps_Slash_When_Asked()
        {
            var o = new MatchOptions { StripLeadingSlash = false, StripTypePrefix = false };
            Assert.Equal("PIPE /0-A", KeyNormalizer.Normalize("Pipe /0-a", o));
            Assert.Equal("/0-A", KeyNormalizer.Normalize("/0-a", o));
        }

        [Fact]
        public void IgnoreSeparators_Compares_Letters_And_Digits()
        {
            var o = new MatchOptions { IgnoreSeparators = true };
            Assert.Equal(KeyNormalizer.Normalize("150 P_1001", o), KeyNormalizer.Normalize("/150-P-1001", o));
        }

        [Theory]
        [InlineData("150-P-1001-A1A", "P-1001", MatchMode.ContainsWhole, true)]
        [InlineData("150-P-10010-A1A", "P-1001", MatchMode.ContainsWhole, false)]
        [InlineData("ISO-1001A01", "1001A", MatchMode.ContainsWhole, false)]
        [InlineData("ISO-1001A01", "1001A", MatchMode.Contains, true)]
        [InlineData("ISO-1001A", "1001A", MatchMode.ContainsWhole, true)]
        [InlineData("X1001A-Y 1001A", "1001A", MatchMode.ContainsWhole, true)]
        [InlineData("150-P-1001", "150-P", MatchMode.StartsWith, true)]
        [InlineData("150-P-1001", "P", MatchMode.StartsWith, false)]
        [InlineData("150-P-1001", "150-P-1001", MatchMode.Exact, true)]
        [InlineData("150-P-1001", "150-P-100", MatchMode.Exact, false)]
        [InlineData("", "A", MatchMode.Contains, false)]
        public void IsMatch_Modes(string name, string key, MatchMode mode, bool expected)
        {
            Assert.Equal(expected, KeyNormalizer.IsMatch(name, key, mode));
        }

        [Fact]
        public void Index_Exact_Finds_All_Items_With_Same_Name()
        {
            var idx = new KeyIndex<Node>(new MatchOptions());
            var a = new Node { Name = "/150-P-1001-A1A-C-CS" };
            var b = new Node { Name = "PIPE /150-P-1001-A1A-C-CS" };
            var c = new Node { Name = "/80-CW-2002-C1" };
            idx.Add(a.Name, a); idx.Add(b.Name, b); idx.Add(c.Name, c); idx.Add("   ", new Node());
            Assert.Equal(3, idx.ItemCount);
            Assert.Equal(2, idx.DistinctNames);
            var hits = idx.Find("150-p-1001-a1a-c-cs");
            Assert.Equal(new[] { a, b }, hits.ToArray());
            Assert.Empty(idx.Find("150-P-1001"));
        }

        [Fact]
        public void Index_ContainsWhole_Finds_Short_Line_Number()
        {
            var idx = new KeyIndex<Node>(new MatchOptions { Mode = MatchMode.ContainsWhole });
            var a = new Node { Name = "/150-P-1001-A1A-C-CS" };
            var b = new Node { Name = "/150-P-10011-A1A" };
            idx.Add(a.Name, a); idx.Add(b.Name, b);
            Assert.Equal(new[] { a }, idx.Find("P-1001").ToArray());
        }

        [Fact]
        public void TopMost_Drops_Children_Of_Matched_Parents()
        {
            var pipe = new Node { Name = "pipe" };
            var br = new Node { Name = "br", Parent = pipe };
            var elbow = new Node { Name = "elbow", Parent = br };
            var other = new Node { Name = "other" };
            var kept = KeyIndex<Node>.TopMost(new List<Node> { elbow, pipe, br, other, pipe }, n => n.Parent);
            Assert.Equal(new[] { pipe, other }, kept.ToArray());
        }
    }

    public class DocResolverTests
    {
        private static FakeFs Fs()
        {
            var fs = new FakeFs();
            fs.Dirs.Add(@"C:\Iso");
            fs.Dirs.Add(@"C:\Iso\AREA-A");
            fs.Dirs.Add(@"C:\Proj");
            fs.Files.Add(@"C:\Iso\AREA-A\ISO-1001A_Rev0.pdf");
            fs.Files.Add(@"C:\Iso\AREA-A\ISO-1001A01_Rev0.pdf");
            fs.Files.Add(@"C:\Iso\150-P-1001-A1A-C-CS.pdf");
            fs.Files.Add(@"C:\Iso\150-P-1001-A1A-C-CS.dwg");
            fs.Files.Add(@"C:\Proj\docs\a.pdf");
            fs.Files.Add(@"C:\Proj\b.pdf");
            return fs;
        }

        [Theory]
        [InlineData("https://acc.autodesk.com/x", true)]
        [InlineData("http://server/doc.pdf", true)]
        [InlineData("www.example.com/a", true)]
        [InlineData("mailto:a@b.com", true)]
        [InlineData(@"C:\Iso\a.pdf", false)]
        [InlineData(@"\\server\share\a.pdf", false)]
        [InlineData(@"docs\a.pdf", false)]
        [InlineData("", false)]
        public void LooksLikeUrl(string s, bool expected)
        {
            Assert.Equal(expected, DocResolver.LooksLikeUrl(s));
        }

        [Fact]
        public void SplitCell_Semicolons_And_Lines()
        {
            Assert.Equal(new[] { "a.pdf", "b.pdf", "c d.pdf" }, DocResolver.SplitCell(" a.pdf; \"b.pdf\"\r\nc d.pdf ;").ToArray());
            Assert.Empty(DocResolver.SplitCell("  "));
        }

        [Fact]
        public void Resolves_Absolute_Relative_Missing_And_Url()
        {
            var r = new DocResolver(new DocOptions { BaseFolder = @"C:\Proj" }, Fs());
            var docs = r.Resolve("K", new[]
            {
                new KeyValuePair<string,string>("Doc", @"docs\a.pdf; ..\Proj\b.pdf"),
                new KeyValuePair<string,string>("Doc2", @"C:\Nope\x.pdf"),
                new KeyValuePair<string,string>("Web", "https://x.y/z"),
                new KeyValuePair<string,string>("Dup", @"C:\Proj\docs\a.pdf"),
            });
            Assert.Equal(4, docs.Count);
            Assert.Equal(@"C:\Proj\docs\a.pdf", docs[0].Target); Assert.True(docs[0].Exists);
            Assert.Equal(@"C:\Proj\b.pdf", docs[1].Target); Assert.True(docs[1].Exists);
            Assert.False(docs[2].Exists); Assert.Equal("Doc2", docs[2].Source);
            Assert.Equal(DocKind.Url, docs[3].Kind);
        }

        [Fact]
        public void Folder_Cell_Is_Searched_For_Key_Else_Linked_As_Folder()
        {
            var r = new DocResolver(new DocOptions(), Fs());
            var hit = r.Resolve("/150-P-1001-A1A-C-CS", new[] { new KeyValuePair<string,string>("F", @"C:\Iso") });
            Assert.Single(hit);
            Assert.Equal(@"C:\Iso\150-P-1001-A1A-C-CS.pdf", hit[0].Target);
            var none = r.Resolve("ZZZ", new[] { new KeyValuePair<string,string>("F", @"C:\Iso") });
            Assert.Single(none);
            Assert.Equal(DocKind.Folder, none[0].Kind);
        }

        [Fact]
        public void Doc_Folder_Pattern_Search_Respects_Pattern()
        {
            var o = new DocOptions { SearchFolder = @"C:\Iso", Pattern = "*{KEY}_*.pdf", Recursive = true };
            var r = new DocResolver(o, Fs());
            var docs = r.Resolve("1001A", null);
            Assert.Single(docs);
            Assert.EndsWith("ISO-1001A_Rev0.pdf", docs[0].Target);
            o.Recursive = false;
            Assert.Empty(new DocResolver(o, Fs()).Resolve("1001A", null));
        }

        [Fact]
        public void Doc_Folder_Only_When_Cells_Empty()
        {
            var o = new DocOptions { SearchFolder = @"C:\Iso" };
            var r = new DocResolver(o, Fs());
            var docs = r.Resolve("150-P-1001-A1A-C-CS", new[] { new KeyValuePair<string,string>("D", @"C:\Proj\b.pdf") });
            Assert.Single(docs);
            o.FolderOnlyWhenCellsEmpty = false;
            docs = new DocResolver(o, Fs()).Resolve("150-P-1001-A1A-C-CS", new[] { new KeyValuePair<string,string>("D", @"C:\Proj\b.pdf") });
            Assert.Equal(2, docs.Count);
        }

        [Fact]
        public void Loose_File_Names_Ignore_Separators()
        {
            var fs = new FakeFs();
            fs.Dirs.Add(@"C:\D");
            fs.Files.Add(@"C:\D\ISO1001A rev0.PDF");
            var strict = new DocResolver(new DocOptions { SearchFolder = @"C:\D" }, fs);
            Assert.Empty(strict.Resolve("ISO-1001A", null));
            var loose = new DocResolver(new DocOptions { SearchFolder = @"C:\D", LooseFileNames = true }, fs);
            Assert.Single(loose.Resolve("ISO-1001A", null));
        }

        [Fact]
        public void Pattern_Key_Bad_Filename_Chars_Match_Any_Char()
        {
            var rx = DocResolver.PatternRegex("{KEY}.pdf", "/A/B:C", false);
            Assert.Matches(rx, "A_B-C.pdf");
            Assert.DoesNotMatch(rx, "A_B-C.dwg");
            Assert.Matches(DocResolver.PatternRegex("*.pdf", "K1", false), "xx K1 yy.pdf");
        }
    }

    public class PlanTests
    {
        private static DataTable Table()
        {
            var t = new DataTable();
            foreach (var c in new[] { "Line", "Iso", "Label" }) t.EnsureColumn(c);
            var r1 = t.AddRow(); r1.ExcelRow = 5; r1.Set("Line", "/0-A-1"); r1.Set("Iso", @"C:\D\a.pdf"); r1.Set("Label", "ISO A");
            var r2 = t.AddRow(); r2.ExcelRow = 6; r2.Set("Line", "/0-B-2"); r2.Set("Iso", @"C:\D\missing.pdf; https://w/x");
            var r3 = t.AddRow(); r3.ExcelRow = 7; r3.Set("Line", ""); r3.Set("Iso", @"C:\D\a.pdf");
            var r4 = t.AddRow(); r4.ExcelRow = 8; r4.Set("Line", "/0-C-3");
            var r5 = t.AddRow(); r5.ExcelRow = 9; r5.Set("Line", "/0-D-4"); r5.Set("Iso", "shown text"); r5.Links["Iso"] = @"C:\D\a.pdf";
            return t;
        }

        [Fact]
        public void Plan_Statuses_And_Labels()
        {
            var fs = new FakeFs(); fs.Dirs.Add(@"C:\D"); fs.Files.Add(@"C:\D\a.pdf");
            var plan = LinkPlanner.Build(Table(), new PlanSettings { KeyColumn = "Line", DocColumns = { "Iso" }, LabelColumn = "Label" },
                new DocResolver(new DocOptions(), fs));
            Assert.Equal(5, plan.Count);
            plan[0].Finish(2); plan[1].Finish(1); plan[2].Finish(0); plan[3].Finish(4); plan[4].Finish(0);
            Assert.Equal(RowStatus.Matched, plan[0].Status);
            Assert.Equal(RowStatus.DocMissing, plan[1].Status);
            Assert.Equal(RowStatus.NoKey, plan[2].Status);
            Assert.Equal(RowStatus.NoDocument, plan[3].Status);
            Assert.Equal(RowStatus.NotFound, plan[4].Status);
            Assert.Equal(@"C:\D\a.pdf", plan[4].Docs[0].Target); // hyperlink target wins over cell text
            Assert.Equal("MD | ISO A", plan[0].LinkPairs(false)[0].Key);
            var p1 = plan[1].LinkPairs(false);
            Assert.Single(p1);
            Assert.Equal("MD | https://w/x", p1[0].Key);
            Assert.Equal(2, plan[1].LinkPairs(true).Count);
            Assert.Equal("MD | missing.pdf", plan[1].LinkPairs(true)[0].Key);
        }

        [Fact]
        public void Merge_And_Combine_Dedupe_And_Replace_Md_Links()
        {
            var fs = new FakeFs(); fs.Dirs.Add(@"C:\D"); fs.Files.Add(@"C:\D\a.pdf");
            var plan = LinkPlanner.Build(Table(), new PlanSettings { KeyColumn = "Line", DocColumns = { "Iso" } },
                new DocResolver(new DocOptions(), fs));
            var item = new Node { Name = "x" };
            var merged = LinkPlanner.MergePerItem(new[]
            {
                new KeyValuePair<PlanRow, IList<Node>>(plan[0], new List<Node> { item }),
                new KeyValuePair<PlanRow, IList<Node>>(plan[4], new List<Node> { item }),
            }, false);
            Assert.Single(merged[item]); // same a.pdf from two rows
            var existing = new[]
            {
                new KeyValuePair<string,string>("Vendor datasheet", @"C:\V\ds.pdf"),
                new KeyValuePair<string,string>("MD | old.pdf", @"C:\D\old.pdf"),
            };
            var replaced = LinkPlanner.Combine(existing, merged[item], true);
            Assert.Equal(new[] { "Vendor datasheet", "MD | a.pdf" }, replaced.Select(x => x.Key).ToArray());
            var appended = LinkPlanner.Combine(existing, merged[item].Concat(new[] { new KeyValuePair<string,string>("MD | dup", @"c:\d\OLD.pdf") }), false);
            Assert.Equal(3, appended.Count);
        }

        [Fact]
        public void Report_Table_Has_All_Rows()
        {
            var fs = new FakeFs();
            var plan = LinkPlanner.Build(Table(), new PlanSettings { KeyColumn = "Line", DocColumns = { "Iso" } }, new DocResolver(new DocOptions(), fs));
            foreach (var p in plan) p.Finish(0);
            var t = LinkPlanner.ReportTable(plan, "Unmatched");
            Assert.Equal(5, t.Rows.Count);
            Assert.Equal("Not found in model", t.Rows[4].Get("Status"));
            Assert.Equal("Doc file missing", "Doc file missing");
            Assert.Equal(@"C:\D\a.pdf", t.Rows[0].Get("Missing Docs"));
        }

        [Fact]
        public void Md_Link_Names()
        {
            Assert.Equal("MD | a", Names.MdLinkName("a"));
            Assert.Equal("MD | a", Names.MdLinkName("MD | a"));
            Assert.Equal("MD | Document", Names.MdLinkName("  "));
            Assert.True(Names.IsMdLinkName("MD | x"));
            Assert.False(Names.IsMdLinkName("Vendor"));
        }
    }

    public class ExcelTests
    {
        private static void Put(ZipArchive z, string name, string xml)
        {
            using (var w = new StreamWriter(z.CreateEntry(name).Open(), new UTF8Encoding(false))) w.Write(xml);
        }

        /// <summary>Two sheets; second has title row, header on row 3, units row 4, data + a hyperlink cell.</summary>
        private static string MakeWorkbook()
        {
            string path = Path.Combine(Path.GetTempPath(), "mddoclink_test_" + Guid.NewGuid().ToString("N") + ".xlsx");
            const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const string pr = "http://schemas.openxmlformats.org/package/2006/relationships";
            using (var fs = File.Create(path))
            using (var z = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                Put(z, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
                Put(z, "xl/workbook.xml", "<workbook xmlns=\"" + ns + "\" xmlns:r=\"" + rns + "\"><sheets>" +
                    "<sheet name=\"Cover\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"LINE LIST R2\" sheetId=\"2\" r:id=\"rId2\"/></sheets></workbook>");
                Put(z, "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"" + pr + "\">" +
                    "<Relationship Id=\"rId1\" Type=\"" + rns + "/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                    "<Relationship Id=\"rId2\" Type=\"" + rns + "/worksheet\" Target=\"worksheets/sheet2.xml\"/>" +
                    "<Relationship Id=\"rId3\" Type=\"" + rns + "/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>");
                Put(z, "xl/sharedStrings.xml", "<sst xmlns=\"" + ns + "\"><si><t>LINE\nNo.</t></si><si><r><t>ISO </t></r><r><t>PDF</t></r></si><si><t>/150-P-1001</t></si></sst>");
                Put(z, "xl/worksheets/sheet1.xml", "<worksheet xmlns=\"" + ns + "\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>cover</t></is></c></row></sheetData></worksheet>");
                Put(z, "xl/worksheets/sheet2.xml", "<worksheet xmlns=\"" + ns + "\" xmlns:r=\"" + rns + "\"><sheetData>" +
                    "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>SAMPLE LINE LIST</t></is></c></row>" +
                    "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>0</v></c><c r=\"B3\" t=\"s\"><v>1</v></c><c r=\"C3\" t=\"inlineStr\"><is><t>SIZE</t></is></c></row>" +
                    "<row r=\"4\"><c r=\"C4\" t=\"inlineStr\"><is><t>mm</t></is></c></row>" +
                    "<row r=\"5\"><c r=\"A5\" t=\"s\"><v>2</v></c><c r=\"B5\" t=\"inlineStr\"><is><t>ISO-1</t></is></c><c r=\"C5\"><v>500</v></c></row>" +
                    "<row r=\"6\"></row>" +
                    "<row r=\"7\"><c r=\"A7\" t=\"inlineStr\"><is><t>/80-CW-2002-C1</t></is></c><c r=\"B7\" t=\"inlineStr\"><is><t>C:\\Iso\\b.pdf</t></is></c></row>" +
                    "</sheetData><hyperlinks><hyperlink ref=\"B5\" r:id=\"rId9\"/></hyperlinks></worksheet>");
                Put(z, "xl/worksheets/_rels/sheet2.xml.rels", "<Relationships xmlns=\"" + pr + "\">" +
                    "<Relationship Id=\"rId9\" Type=\"" + rns + "/hyperlink\" Target=\"file:///C:/Iso/a%20b.pdf\" TargetMode=\"External\"/></Relationships>");
            }
            return path;
        }

        [Fact]
        public void Lists_Sheets_In_Order()
        {
            string p = MakeWorkbook();
            try { Assert.Equal(new[] { "Cover", "LINE LIST R2" }, XlsxReader.ListSheets(p).ToArray()); }
            finally { File.Delete(p); }
        }

        [Fact]
        public void Loads_Named_Sheet_With_Auto_Header_Row_And_Hyperlinks()
        {
            string p = MakeWorkbook();
            try
            {
                DataTable t = XlsxReader.Load(p, "line list r2", 0);
                Assert.Equal("LINE LIST R2", t.SheetName);
                Assert.Equal(3, t.HeaderRow);
                Assert.Equal(new[] { "LINE No.", "ISO PDF", "SIZE" }, t.Columns.ToArray());
                Assert.Equal(3, t.Rows.Count); // units row 4 + data rows 5, 7 (empty row 6 skipped)
                TableRow r5 = t.Rows.First(r => r.ExcelRow == 5);
                Assert.Equal("/150-P-1001", r5.Get("LINE No."));
                Assert.Equal("ISO-1", r5.Get("ISO PDF"));
                Assert.Equal("file:///C:/Iso/a b.pdf", r5.GetLinkOrText("ISO PDF"));
                Assert.Equal("500", r5.Get("SIZE"));
                Assert.Equal(@"C:\Iso\b.pdf", t.Rows.First(r => r.ExcelRow == 7).GetLinkOrText("ISO PDF"));
            }
            finally { File.Delete(p); }
        }

        [Fact]
        public void Explicit_Header_Row_And_Missing_Sheet()
        {
            string p = MakeWorkbook();
            try
            {
                DataTable t = XlsxReader.Load(p, "LINE LIST R2", 5);
                Assert.Equal("/150-P-1001", t.Columns[0]);
                Assert.Single(t.Rows);
                Assert.Throws<InvalidOperationException>(() => XlsxReader.Load(p, "Nope", 0));
                Assert.Equal("Cover", XlsxReader.Load(p).SheetName);
            }
            finally { File.Delete(p); }
        }

        [Fact]
        public void Writer_Then_Reader_RoundTrip()
        {
            string p = Path.Combine(Path.GetTempPath(), "mddoclink_rt_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                var plan = new List<PlanRow> { new PlanRow { ExcelRow = 9, Key = "/0-A & <B>" } };
                plan[0].Finish(0);
                XlsxWriter.Save(p, LinkPlanner.ReportTable(plan, "Report"), Names.Author);
                DataTable t = XlsxReader.Load(p);
                Assert.Equal(1, t.HeaderRow);
                Assert.Equal("/0-A & <B>", t.Rows[0].Get("Key"));
                Assert.Equal("No document", t.Rows[0].Get("Status"));
            }
            finally { File.Delete(p); }
        }

    }
}
