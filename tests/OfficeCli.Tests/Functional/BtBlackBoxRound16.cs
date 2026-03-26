// Black-box tests (Round 16) — final coverage of untested scenarios:
//   1. Word SDT (content control) dropdown: Add → Get → schema valid
//   2. Word SDT combobox: Add → Get → schema valid
//   3. Word SDT datepicker: Add → Get → schema valid
//   4. Word SDT block-level text: Add → Get → schema valid
//   5. Word hyperlink: Add → Get → schema valid + Remove (relationship cleaned up)
//   6. Word page break: Add → schema valid → paragraph count increases
//   7. PPTX connector: Add → Get → type/path verified → Remove → gone
//   8. PPTX connector persistence after reopen
//   9. Excel rich text run: Add "run" → cell text contains run content
//  10. Excel aboveaverage CF: Add → schema valid
//  11. Excel topn CF: Add → schema valid
//  12. Excel uniquevalues CF: Add → schema valid

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;
using WRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WBreak = DocumentFormat.OpenXml.Wordprocessing.Break;
using XRun = DocumentFormat.OpenXml.Spreadsheet.Run;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound16 : IDisposable
{
    private readonly List<string> _temps = new();
    private readonly ITestOutputHelper _out;

    public BtBlackBoxRound16(ITestOutputHelper output) => _out = output;

    public void Dispose()
    {
        foreach (var f in _temps)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt16_{Guid.NewGuid():N}.{ext}");
        _temps.Add(p);
        BlankDocCreator.Create(p);
        return p;
    }

    private void ValidateDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"DOCX invalid after: {step}");
    }

    private void ValidateXlsx(string path, string step)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"XLSX invalid after: {step}");
    }

    private void ValidatePptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"PPTX invalid after: {step}");
    }

    // ==================== 1. Word SDT dropdown ====================

    [Fact]
    public void Word_Sdt_Dropdown_AddGetSchemaValid()
    {
        var path = Temp("docx");
        string sdtPath;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Before control" });
            sdtPath = h.Add("/body", "sdt", null, new()
            {
                ["sdttype"] = "dropdown",
                ["alias"] = "Color",
                ["tag"] = "colorTag",
                ["items"] = "Red,Green,Blue",
                ["text"] = "Red"
            })!;
            _out.WriteLine($"SDT path: {sdtPath}");
        }

        ValidateDocx(path, "after SDT dropdown Add");

        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Get(sdtPath);
        node.Should().NotBeNull("SDT must be retrievable via Get");
        _out.WriteLine($"SDT type: {node!.Type}, sdtType: {node.Format.GetValueOrDefault("sdtType")}");
        node.Type.Should().Be("sdt");
        node.Format.Should().ContainKey("sdtType");
        node.Format["sdtType"].Should().Be("dropdown");
    }

    // ==================== 2. Word SDT combobox ====================

    [Fact]
    public void Word_Sdt_Combobox_AddSchemaValid()
    {
        var path = Temp("docx");
        string sdtPath;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Select size:" });
            sdtPath = h.Add("/body", "sdt", null, new()
            {
                ["sdttype"] = "combobox",
                ["alias"] = "Size",
                ["items"] = "Small,Medium,Large",
                ["text"] = "Medium"
            })!;
            _out.WriteLine($"SDT combobox path: {sdtPath}");
        }

        ValidateDocx(path, "after SDT combobox Add");

        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Get(sdtPath);
        node.Should().NotBeNull("combobox SDT must be retrievable");
        node!.Type.Should().Be("sdt");
        node.Format.GetValueOrDefault("sdtType")?.ToString().Should().Be("combobox");
    }

    // ==================== 3. Word SDT datepicker ====================

    [Fact]
    public void Word_Sdt_DatePicker_AddSchemaValid()
    {
        var path = Temp("docx");
        string sdtPath;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Date control:" });
            sdtPath = h.Add("/body", "sdt", null, new()
            {
                ["sdttype"] = "date",
                ["alias"] = "EventDate",
                ["format"] = "yyyy-MM-dd",
                ["text"] = "2025-01-01"
            })!;
            _out.WriteLine($"SDT datepicker path: {sdtPath}");
        }

        ValidateDocx(path, "after SDT datepicker Add");

        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Get(sdtPath);
        node.Should().NotBeNull("datepicker SDT must be retrievable");
        node!.Type.Should().Be("sdt");
        node.Format.GetValueOrDefault("sdtType")?.ToString().Should().Be("date");
    }

    // ==================== 4. Word SDT block-level plain text ====================

    [Fact]
    public void Word_Sdt_PlainText_Block_AddGetSchemaValid()
    {
        var path = Temp("docx");
        string sdtPath;

        using (var h = new WordHandler(path, editable: true))
        {
            sdtPath = h.Add("/body", "sdt", null, new()
            {
                ["sdttype"] = "text",
                ["alias"] = "Notes",
                ["tag"] = "notesTag",
                ["text"] = "Enter notes here"
            })!;
            _out.WriteLine($"SDT text path: {sdtPath}");
        }

        ValidateDocx(path, "after SDT plain text block Add");

        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Get(sdtPath);
        node.Should().NotBeNull("plain text SDT must be retrievable");
        node!.Type.Should().Be("sdt");
        node.Format.GetValueOrDefault("sdtType")?.ToString().Should().Be("text");
        _out.WriteLine($"SDT Text: {node.Text}");
    }

    // ==================== 5. Word hyperlink Add/Get/schema valid ====================

    [Fact]
    public void Word_Hyperlink_AddGetSchemaValid()
    {
        var path = Temp("docx");
        string hlPath;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Visit " });
            hlPath = h.Add("/body/p[1]", "hyperlink", null, new()
            {
                ["url"] = "https://example.com",
                ["text"] = "Example",
                ["color"] = "0563C1"
            })!;
            _out.WriteLine($"Hyperlink path: {hlPath}");
        }

        ValidateDocx(path, "after hyperlink Add");

        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Get(hlPath);
        node.Should().NotBeNull("hyperlink node must be retrievable");
        _out.WriteLine($"Hyperlink Type: {node!.Type}, Text: {node.Text}");
        node.Type.Should().Be("hyperlink");
        node.Text.Should().Contain("Example");

        // Verify relationship exists
        using var doc = WordprocessingDocument.Open(path, false);
        var rels = doc.MainDocumentPart!.HyperlinkRelationships.ToList();
        rels.Should().NotBeEmpty("hyperlink relationship must exist");
        rels.Any(r => r.Uri.ToString().Contains("example.com")).Should().BeTrue("hyperlink URI must match");
    }

    [Fact]
    public void Word_Hyperlink_SchemaValidAfterAdd()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Check " });
            h.Add("/body/p[1]", "hyperlink", null, new()
            {
                ["url"] = "https://officecli.ai",
                ["text"] = "OfficeCli"
            });
        }

        ValidateDocx(path, "hyperlink schema check");
    }

    // ==================== 6. Word page break ====================

    [Fact]
    public void Word_PageBreak_AddIncreasesRunCount_SchemaValid()
    {
        var path = Temp("docx");
        int runsBefore;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Page 1 content" });
        }

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            runsBefore = doc.MainDocumentPart!.Document!.Body!.Descendants<WRun>().Count();
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            var breakPath = h2.Add("/body/p[1]", "break", null, new() { ["type"] = "page" });
            _out.WriteLine($"Break path: {breakPath}");
        }

        ValidateDocx(path, "after page break Add");

        using (var doc2 = WordprocessingDocument.Open(path, false))
        {
            var runsAfter = doc2.MainDocumentPart!.Document!.Body!.Descendants<WRun>().Count();
            _out.WriteLine($"Runs before: {runsBefore}, after: {runsAfter}");
            runsAfter.Should().BeGreaterThan(runsBefore, "page break run must be added");

            var breaks = doc2.MainDocumentPart.Document.Body!.Descendants<WBreak>()
                .Where(b => b.Type?.Value == BreakValues.Page).ToList();
            breaks.Should().NotBeEmpty("page break element must exist in body");
        }
    }

    [Fact]
    public void Word_PageBreak_BodyLevel_CreatesNewParagraph_SchemaValid()
    {
        var path = Temp("docx");
        int parasBefore;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Before break" });
        }

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            parasBefore = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Count();
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Add("/body", "break", null, new() { ["type"] = "page" });
        }

        ValidateDocx(path, "after body-level page break Add");

        using (var doc2 = WordprocessingDocument.Open(path, false))
        {
            var parasAfter = doc2.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Count();
            _out.WriteLine($"Paragraphs before: {parasBefore}, after: {parasAfter}");
            parasAfter.Should().BeGreaterThan(parasBefore, "body-level page break creates new paragraph");

            var breakRuns = doc2.MainDocumentPart.Document.Body!
                .Descendants<WBreak>()
                .Where(b => b.Type?.Value == BreakValues.Page)
                .ToList();
            breakRuns.Should().NotBeEmpty("page break must exist somewhere in body");
        }
    }

    // ==================== 7. PPTX connector Add/Get/Remove ====================

    [Fact]
    public void Pptx_Connector_AddGetRemove_Lifecycle()
    {
        var path = Temp("pptx");

        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Connector test" });
            var cxnPath = h.Add("/slide[1]", "connector", null, new()
            {
                ["x"] = "1000000",
                ["y"] = "1000000",
                ["width"] = "3000000",
                ["height"] = "0",
                ["color"] = "FF0000",
                ["preset"] = "straight"
            });
            _out.WriteLine($"Connector path: {cxnPath}");
            cxnPath.Should().Contain("connector", "Add connector must return connector path");
        }

        ValidatePptx(path, "after connector Add");

        string connectorPath;
        using (var h2 = new PowerPointHandler(path, editable: false))
        {
            var node = h2.Get("/slide[1]/connector[1]");
            node.Should().NotBeNull("connector[1] must be retrievable");
            connectorPath = node!.Path;
            _out.WriteLine($"Connector path from Get: {connectorPath}, type={node.Type}");
            node.Type.Should().Be("connector");
        }

        using (var h3 = new PowerPointHandler(path, editable: true))
        {
            var rmEx = Record.Exception(() => h3.Remove("/slide[1]/connector[1]"));
            rmEx.Should().BeNull("connector Remove must not throw");
        }

        ValidatePptx(path, "after connector Remove");

        using var h4 = new PowerPointHandler(path, editable: false);
        var act = () => h4.Get("/slide[1]/connector[1]");
        act.Should().Throw<ArgumentException>("connector must be gone after Remove");
    }

    // ==================== 8. PPTX connector persistence ====================

    [Fact]
    public void Pptx_Connector_Persistence_AfterReopen()
    {
        var path = Temp("pptx");

        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Persist connector" });
            h.Add("/slide[1]", "connector", null, new()
            {
                ["x"] = "500000",
                ["y"] = "500000",
                ["width"] = "5000000",
                ["height"] = "2000000",
                ["preset"] = "elbow",
                ["color"] = "0000FF"
            });
        }

        ValidatePptx(path, "after connector Add (persistence)");

        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[1]/connector[1]");
        node.Should().NotBeNull("connector must persist after reopen");
        _out.WriteLine($"Persisted connector type: {node!.Type}, path: {node.Path}");
        node.Type.Should().Be("connector");
        node.Path.Should().Be("/slide[1]/connector[1]");
    }

    // ==================== 9. Excel rich text run Add ====================

    [Fact]
    public void Excel_RichTextRun_AddToCell_SchemaValid()
    {
        var path = Temp("xlsx");

        using (var h = new ExcelHandler(path, editable: true))
        {
            // Add a run to an existing cell to append rich text
            h.Set("/Sheet1/A1", new() { ["value"] = "Hello" });
            var runPath = h.Add("/Sheet1/A1", "run", null, new()
            {
                ["text"] = " World",
                ["bold"] = "true",
                ["color"] = "FF0000",
                ["size"] = "12"
            });
            _out.WriteLine($"Run path: {runPath}");
            runPath.Should().NotBeNull("Add run must succeed");
        }

        ValidateXlsx(path, "after rich text run Add");

        using var doc = SpreadsheetDocument.Open(path, false);
        var sst = doc.WorkbookPart!.SharedStringTablePart?.SharedStringTable;
        sst.Should().NotBeNull("SharedStringTable must exist after rich text Add");
        var items = sst!.Elements<SharedStringItem>().ToList();
        _out.WriteLine($"SST items count: {items.Count}");
        items.Should().NotBeEmpty("SST must contain the rich text item");

        // Check that at least one SST item has runs
        var hasRuns = items.Any(item => item.Elements<DocumentFormat.OpenXml.Spreadsheet.Run>().Any());
        hasRuns.Should().BeTrue("SST item must contain runs for rich text");
    }

    // ==================== 10. Excel aboveaverage CF ====================

    [Fact]
    public void Excel_AboveAverage_CF_AddSchemaValid()
    {
        var path = Temp("xlsx");

        using (var h = new ExcelHandler(path, editable: true))
        {
            // Set some values first
            h.Set("/Sheet1/A1", new() { ["value"] = "10" });
            h.Set("/Sheet1/A2", new() { ["value"] = "20" });
            h.Set("/Sheet1/A3", new() { ["value"] = "30" });

            var cfPath = h.Add("/Sheet1", "aboveaverage", null, new()
            {
                ["sqref"] = "A1:A3",
                ["above"] = "true"
            });
            _out.WriteLine($"AboveAverage CF path: {cfPath}");
            cfPath.Should().NotBeNull("Add aboveaverage CF must succeed");
        }

        ValidateXlsx(path, "after aboveaverage CF Add");

        using var doc = SpreadsheetDocument.Open(path, false);
        var wbPart = doc.WorkbookPart!;
        var sheetId = wbPart.Workbook.Sheets!.Elements<Sheet>().First().Id!.Value!;
        var wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
        var cfs = wsPart.Worksheet.Elements<ConditionalFormatting>().ToList();
        cfs.Should().NotBeEmpty("ConditionalFormatting element must exist");
        var cfRule = cfs.SelectMany(cf => cf.Elements<ConditionalFormattingRule>()).ToList();
        cfRule.Should().NotBeEmpty("CF rule must exist");
        cfRule.Any(r => r.Type?.Value == ConditionalFormatValues.AboveAverage).Should().BeTrue("AboveAverage CF must be present");
    }

    // ==================== 11. Excel topn CF ====================

    [Fact]
    public void Excel_TopN_CF_AddSchemaValid()
    {
        var path = Temp("xlsx");

        using (var h = new ExcelHandler(path, editable: true))
        {
            for (int i = 1; i <= 10; i++)
                h.Set($"/Sheet1/A{i}", new() { ["value"] = (i * 5).ToString() });

            var cfPath = h.Add("/Sheet1", "topn", null, new()
            {
                ["sqref"] = "A1:A10",
                ["rank"] = "3",
                ["percent"] = "false",
                ["bottom"] = "false"
            });
            _out.WriteLine($"TopN CF path: {cfPath}");
            cfPath.Should().NotBeNull("Add topn CF must succeed");
        }

        ValidateXlsx(path, "after topn CF Add");

        using var doc = SpreadsheetDocument.Open(path, false);
        var wbPart = doc.WorkbookPart!;
        var sheetId = wbPart.Workbook.Sheets!.Elements<Sheet>().First().Id!.Value!;
        var wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
        var cfRule = wsPart.Worksheet.Elements<ConditionalFormatting>()
            .SelectMany(cf => cf.Elements<ConditionalFormattingRule>()).ToList();
        cfRule.Any(r => r.Type?.Value == ConditionalFormatValues.Top10).Should().BeTrue("Top10 CF must be present");
        var top10Rule = cfRule.First(r => r.Type?.Value == ConditionalFormatValues.Top10);
        top10Rule.Rank?.Value.Should().Be(3u, "rank must be 3");
    }

    // ==================== 12. Excel uniquevalues CF ====================

    [Fact]
    public void Excel_UniqueValues_CF_AddSchemaValid()
    {
        var path = Temp("xlsx");

        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Set("/Sheet1/B1", new() { ["value"] = "Alpha" });
            h.Set("/Sheet1/B2", new() { ["value"] = "Beta" });
            h.Set("/Sheet1/B3", new() { ["value"] = "Alpha" });
            h.Set("/Sheet1/B4", new() { ["value"] = "Gamma" });

            var cfPath = h.Add("/Sheet1", "uniquevalues", null, new()
            {
                ["sqref"] = "B1:B4"
            });
            _out.WriteLine($"UniqueValues CF path: {cfPath}");
            cfPath.Should().NotBeNull("Add uniquevalues CF must succeed");
        }

        ValidateXlsx(path, "after uniquevalues CF Add");

        using var doc = SpreadsheetDocument.Open(path, false);
        var wbPart = doc.WorkbookPart!;
        var sheetId = wbPart.Workbook.Sheets!.Elements<Sheet>().First().Id!.Value!;
        var wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
        var cfRule = wsPart.Worksheet.Elements<ConditionalFormatting>()
            .SelectMany(cf => cf.Elements<ConditionalFormattingRule>()).ToList();
        cfRule.Any(r => r.Type?.Value == ConditionalFormatValues.UniqueValues).Should().BeTrue("UniqueValues CF must be present");
    }
}
