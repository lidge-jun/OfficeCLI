// Black-box tests (Round 3) targeting new functional areas:
//   - Excel: formula, merge cells, conditional formatting, pivot table, chart, data validation
//   - Word: table, list/numbering, section break, footnote, TOC, header/footer
//   - PPTX: table, chart, notes, animation, transition
//   - Cross-format: special chars (Chinese/emoji/long text)

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound3 : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly ITestOutputHelper _output;

    public BtBlackBoxRound3(ITestOutputHelper output) => _output = output;

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bt3_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            if (File.Exists(f)) File.Delete(f);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private void AssertValidDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"DOCX must be schema-valid after: {step}");
    }

    private void AssertValidPptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"PPTX must be schema-valid after: {step}");
    }

    private void AssertValidXlsx(string path, string step)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"XLSX must be schema-valid after: {step}");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 1 — Excel: Formula
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Excel_Formula_AddAndGet()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["value"] = "10" });
        handler.Add("/Sheet1", "cell", null, new() { ["address"] = "A2", ["value"] = "20" });
        handler.Add("/Sheet1", "cell", null, new() { ["address"] = "A3", ["formula"] = "=SUM(A1:A2)" });

        var node = handler.Get("/Sheet1/A3");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].ToString().Should().Contain("SUM");
    }

    [Fact]
    public void Excel_Formula_PersistsAfterReopen()
    {
        var path = CreateTemp("xlsx");

        using (var handler = new ExcelHandler(path, editable: true))
        {
            handler.Add("/Sheet1", "cell", null, new() { ["address"] = "B1", ["value"] = "5" });
            handler.Add("/Sheet1", "cell", null, new() { ["address"] = "B2", ["formula"] = "=B1*2" });
        }

        using (var handler = new ExcelHandler(path, editable: false))
        {
            var node = handler.Get("/Sheet1/B2");
            node.Format.Should().ContainKey("formula");
            node.Format["formula"].ToString().Should().Contain("B1");
        }
    }

    [Fact]
    public void Excel_Formula_SetUpdatesFormula()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["address"] = "C1", ["formula"] = "=1+1" });
        handler.Set("/Sheet1/C1", new() { ["formula"] = "=2+2" });

        var node = handler.Get("/Sheet1/C1");
        node.Format["formula"].ToString().Should().Contain("2+2");
    }

    [Fact]
    public void Excel_Formula_SchemaValid()
    {
        var path = CreateTemp("xlsx");

        using (var handler = new ExcelHandler(path, editable: true))
        {
            handler.Add("/Sheet1", "cell", null, new() { ["address"] = "D1", ["value"] = "100" });
            handler.Add("/Sheet1", "cell", null, new() { ["address"] = "D2", ["formula"] = "=D1+50" });
        }

        AssertValidXlsx(path, "formula cell");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 2 — Excel: Merge Cells
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Excel_MergeCells_SetAndGet()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["value"] = "Merged" });
        handler.Set("/Sheet1", new() { ["merge"] = "A1:C1" });

        var node = handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Excel_MergeCells_PersistsAfterReopen()
    {
        var path = CreateTemp("xlsx");

        using (var handler = new ExcelHandler(path, editable: true))
        {
            handler.Add("/Sheet1", "cell", null, new() { ["address"] = "B2", ["value"] = "Header" });
            handler.Set("/Sheet1", new() { ["merge"] = "B2:D2" });
        }

        using (var handler = new ExcelHandler(path, editable: false))
        {
            var node = handler.Get("/Sheet1/B2");
            node.Should().NotBeNull("merged cell should still be accessible");
        }
    }

    [Fact]
    public void Excel_MergeCells_SchemaValid()
    {
        var path = CreateTemp("xlsx");

        using (var handler = new ExcelHandler(path, editable: true))
        {
            handler.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["value"] = "Merged Title" });
            handler.Set("/Sheet1", new() { ["merge"] = "A1:E1" });
        }

        AssertValidXlsx(path, "merge cells");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 3 — Excel: Conditional Formatting
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Excel_ConditionalFormatting_AddHighlight()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["value"] = "50" });
        handler.Add("/Sheet1", "cell", null, new() { ["address"] = "A2", ["value"] = "150" });

        var act = () => handler.Add("/Sheet1", "cf", null, new()
        {
            ["range"] = "A1:A2",
            ["type"] = "highlight",
            ["operator"] = "greaterThan",
            ["value"] = "100",
            ["bgColor"] = "FFFF00"
        });

        act.Should().NotThrow("conditional formatting add should not throw");
    }

    [Fact]
    public void Excel_ConditionalFormatting_SchemaValid()
    {
        var path = CreateTemp("xlsx");

        using (var handler = new ExcelHandler(path, editable: true))
        {
            for (int i = 1; i <= 5; i++)
                handler.Add("/Sheet1", "cell", null, new() { ["address"] = $"A{i}", ["value"] = (i * 20).ToString() });

            handler.Add("/Sheet1", "cf", null, new()
            {
                ["range"] = "A1:A5",
                ["type"] = "colorscale"
            });
        }

        AssertValidXlsx(path, "colorscale conditional formatting");
    }

    [Fact]
    public void Excel_ConditionalFormatting_Formula()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        for (int i = 1; i <= 4; i++)
            handler.Add("/Sheet1", "cell", null, new() { ["address"] = $"B{i}", ["value"] = (i * 10).ToString() });

        var act = () => handler.Add("/Sheet1", "formulacf", null, new()
        {
            ["range"] = "B1:B4",
            ["formula"] = "$B1>25",
            ["bgColor"] = "FFA500"
        });

        act.Should().NotThrow("formula-based conditional formatting should not throw");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 4 — Excel: Data Validation
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Excel_DataValidation_ListAdd()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);

        // Excel validation requires 'sqref' (not 'range') for the cell reference
        var act = () => handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "C1:C10",
            ["type"] = "list",
            ["formula1"] = "\"Yes,No,Maybe\""
        });

        act.Should().NotThrow("list data validation should not throw");
    }

    [Fact]
    public void Excel_DataValidation_WholeNumberRange()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);

        var act = () => handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "D1:D10",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "100"
        });

        act.Should().NotThrow("whole number validation should not throw");
    }

    [Fact]
    public void Excel_DataValidation_SchemaValid()
    {
        var path = CreateTemp("xlsx");

        using (var handler = new ExcelHandler(path, editable: true))
        {
            handler.Add("/Sheet1", "validation", null, new()
            {
                ["sqref"] = "E1:E5",
                ["type"] = "list",
                ["formula1"] = "\"Alpha,Beta,Gamma\""
            });
        }

        AssertValidXlsx(path, "data validation list");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 5 — Excel: Chart
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Excel_Chart_AddBarChartNoException()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);

        // Charts require series data in format "SeriesName:val1,val2,val3" via data property
        var act = () => handler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["title"] = "Monthly Sales",
            ["data"] = "Sales:100,200,150",
            ["categories"] = "Jan,Feb,Mar"
        });

        act.Should().NotThrow("adding a bar chart should not throw");
    }

    [Fact]
    public void Excel_Chart_SchemaValid()
    {
        var path = CreateTemp("xlsx");

        using (var handler = new ExcelHandler(path, editable: true))
        {
            handler.Add("/Sheet1", "chart", null, new()
            {
                ["type"] = "line",
                ["data"] = "Revenue:10,20,30",
                ["categories"] = "Q1,Q2,Q3"
            });
        }

        AssertValidXlsx(path, "line chart");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 6 — Word: Table
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Word_Table_AddAndGet()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "3" });

        // Word tables are addressed as tbl[N] in paths
        var node = handler.Get("/body/tbl[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("table");
    }

    [Fact]
    public void Word_Table_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "4" });
        }

        AssertValidDocx(path, "Add 2x4 table");
    }

    [Fact]
    public void Word_Table_CellTextReadback()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        // Word table paths use tbl/tr/tc
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Hello" });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
        node.Text.Should().Be("Hello");
    }

    [Fact]
    public void Word_Table_PersistsAfterReopen()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "3" });
            handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Cell11" });
        }

        using (var handler = new WordHandler(path, editable: false))
        {
            var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
            node.Text.Should().Be("Cell11");
        }
    }

    [Fact]
    public void Word_Table_MultipleTablesIndexed()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Between tables" });
        handler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });

        var t1 = handler.Get("/body/tbl[1]");
        var t2 = handler.Get("/body/tbl[2]");
        t1.Should().NotBeNull();
        t2.Should().NotBeNull();
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 7 — Word: List/Numbering
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Word_List_BulletAdd()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Item 1", ["listStyle"] = "bullet" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Item 2", ["listStyle"] = "bullet" });

        // Word paragraphs use /body/p[N] path notation
        var node = handler.Get("/body/p[1]");
        node.Should().NotBeNull();
        node.Text.Should().Be("Item 1");
    }

    [Fact]
    public void Word_List_OrderedAdd()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Step 1", ["listStyle"] = "ordered" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Step 2", ["listStyle"] = "ordered" });

        // Word paragraphs use /body/p[N] path notation
        var node = handler.Get("/body/p[2]");
        node.Should().NotBeNull();
        node.Text.Should().Be("Step 2");
    }

    [Fact]
    public void Word_List_BulletSchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            for (int i = 1; i <= 3; i++)
                handler.Add("/body", "paragraph", null, new() { ["text"] = $"Bullet {i}", ["listStyle"] = "bullet" });
        }

        AssertValidDocx(path, "bullet list");
    }

    [Fact]
    public void Word_List_OrderedSchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            for (int i = 1; i <= 4; i++)
                handler.Add("/body", "paragraph", null, new() { ["text"] = $"Step {i}", ["listStyle"] = "ordered" });
        }

        AssertValidDocx(path, "ordered list");
    }

    [Fact]
    public void Word_List_PersistsAfterReopen()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Apple", ["listStyle"] = "bullet" });
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Banana", ["listStyle"] = "bullet" });
        }

        using (var handler = new WordHandler(path, editable: false))
        {
            // Word paragraphs use /body/p[N] path notation
            var p1 = handler.Get("/body/p[1]");
            p1.Text.Should().Be("Apple");
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 8 — Word: Section Break
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Word_SectionBreak_AddNextPage()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Page 1" });

        var act = () => handler.Add("/body", "section", null, new() { ["type"] = "nextPage" });

        act.Should().NotThrow("section break add should not throw");
    }

    [Fact]
    public void Word_SectionBreak_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Section 1 content" });
            handler.Add("/body", "section", null, new() { ["type"] = "nextPage" });
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Section 2 content" });
        }

        AssertValidDocx(path, "section break");
    }

    [Fact]
    public void Word_SectionBreak_Continuous()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Before" });
            handler.Add("/body", "section", null, new() { ["type"] = "continuous" });
            handler.Add("/body", "paragraph", null, new() { ["text"] = "After" });
        }

        AssertValidDocx(path, "continuous section break");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 9 — Word: Footnote
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Word_Footnote_AddNoException()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Main text." });

        // Word paragraph path uses /body/p[N]
        var act = () => handler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "This is a footnote." });

        act.Should().NotThrow("adding a footnote should not throw");
    }

    [Fact]
    public void Word_Footnote_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Text with note." });
            handler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Footnote content." });
        }

        AssertValidDocx(path, "footnote");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 10 — Word: TOC
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Word_TOC_AddNoException()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Introduction", ["style"] = "Heading1" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Chapter 1", ["style"] = "Heading1" });

        var act = () => handler.Add("/body", "toc", null, new());

        act.Should().NotThrow("adding a TOC should not throw");
    }

    [Fact]
    public void Word_TOC_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "toc", null, new());
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Overview" });
        }

        AssertValidDocx(path, "TOC");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 11 — Word: Header/Footer
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Word_Header_AddAndGet()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Body" });
        handler.Add("/", "header", null, new() { ["type"] = "default", ["text"] = "My Header" });

        var node = handler.Get("/header");
        node.Should().NotBeNull();
        node.Type.Should().Be("header");
    }

    [Fact]
    public void Word_Footer_AddAndGet()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Body" });
        handler.Add("/", "footer", null, new() { ["type"] = "default", ["text"] = "Page Footer" });

        var node = handler.Get("/footer");
        node.Should().NotBeNull();
        node.Type.Should().Be("footer");
    }

    [Fact]
    public void Word_Header_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "header", null, new() { ["type"] = "default", ["text"] = "Company Name" });
        }

        AssertValidDocx(path, "default header");
    }

    [Fact]
    public void Word_Footer_SchemaValid()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "footer", null, new() { ["type"] = "default", ["text"] = "Confidential" });
        }

        AssertValidDocx(path, "default footer");
    }

    [Fact]
    public void Word_Header_PersistsAfterReopen()
    {
        var path = CreateTemp("docx");

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Body text" });
            handler.Add("/", "header", null, new() { ["type"] = "default", ["text"] = "Persisted Header" });
        }

        using (var handler = new WordHandler(path, editable: false))
        {
            var node = handler.Get("/header");
            node.Should().NotBeNull();
            node.Text.Should().Contain("Persisted Header");
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 12 — PPTX: Table
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Pptx_Table_AddAndGet()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "3" });

        var node = handler.Get("/slide[1]/table[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("table");
    }

    [Fact]
    public void Pptx_Table_CellSetAndGet()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        // PPTX table paths use tr[N]/tc[N] notation
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "R1C1" });

        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
        node.Text.Should().Be("R1C1");
    }

    [Fact]
    public void Pptx_Table_SchemaValid()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "4" });
        }

        AssertValidPptx(path, "3x4 table");
    }

    [Fact]
    public void Pptx_Table_PersistsAfterReopen()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
            // PPTX table paths use tr[N]/tc[N] notation
            handler.Set("/slide[1]/table[1]/tr[1]/tc[2]", new() { ["text"] = "Cell12" });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[2]");
            node.Text.Should().Be("Cell12");
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 13 — PPTX: Chart
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Pptx_Chart_AddBarNoException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        // Charts require series data via data="SeriesName:v1,v2,v3" format
        var act = () => handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "bar",
            ["title"] = "Sales Chart",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        act.Should().NotThrow("adding a bar chart to PPTX should not throw");
    }

    [Fact]
    public void Pptx_Chart_AddLineNoException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var act = () => handler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "line",
            ["data"] = "Revenue:5,15,25",
            ["categories"] = "Jan,Feb,Mar"
        });

        act.Should().NotThrow("adding a line chart to PPTX should not throw");
    }

    [Fact]
    public void Pptx_Chart_SchemaValid()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "chart", null, new()
            {
                ["type"] = "pie",
                ["data"] = "Market:30,50,20",
                ["categories"] = "A,B,C"
            });
        }

        AssertValidPptx(path, "pie chart");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 14 — PPTX: Notes
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Pptx_Notes_AddAndGet()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Add("/slide[1]", "notes", null, new() { ["text"] = "Speaker notes here" });

        var node = handler.Get("/slide[1]/notes");
        node.Should().NotBeNull();
        node.Type.Should().Be("notes");
        node.Text.Should().Be("Speaker notes here");
    }

    [Fact]
    public void Pptx_Notes_PersistsAfterReopen()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "notes", null, new() { ["text"] = "Remember to demo the new feature!" });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get("/slide[1]/notes");
            node.Should().NotBeNull();
            node.Text.Should().Contain("demo");
        }
    }

    [Fact]
    public void Pptx_Notes_SchemaValid()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "notes", null, new() { ["text"] = "Slide notes content" });
        }

        AssertValidPptx(path, "notes add");
    }

    [Fact]
    public void Pptx_Notes_SetUpdatesText()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "notes", null, new() { ["text"] = "Original notes" });

        handler.Set("/slide[1]/notes", new() { ["text"] = "Updated notes" });

        var node = handler.Get("/slide[1]/notes");
        node.Text.Should().Be("Updated notes");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 15 — PPTX: Transition
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Pptx_Transition_SetFadeNoException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var act = () => handler.Set("/slide[1]", new() { ["transition"] = "fade" });

        act.Should().NotThrow("setting fade transition should not throw");
    }

    [Fact]
    public void Pptx_Transition_SetWipeLeft()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());

        var act = () => handler.Set("/slide[1]", new() { ["transition"] = "wipe-left" });

        act.Should().NotThrow("setting wipe-left transition should not throw");
    }

    [Fact]
    public void Pptx_Transition_SchemaValid()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Set("/slide[1]", new() { ["transition"] = "fade" });
        }

        AssertValidPptx(path, "fade transition");
    }

    [Fact]
    public void Pptx_Transition_PersistsAfterReopen()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Set("/slide[1]", new() { ["transition"] = "dissolve" });
        }

        // Reopen and verify slide still accessible (transition stored in slide XML)
        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get("/slide[1]");
            node.Should().NotBeNull("slide with transition must survive reopen");
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 16 — PPTX: Animation
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void Pptx_Animation_AddNoException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Animated Shape" });

        var act = () => handler.Add("/slide[1]/shape[1]", "animation", null, new()
        {
            ["effect"] = "appear",
            ["trigger"] = "onClick"
        });

        act.Should().NotThrow("adding an animation should not throw");
    }

    [Fact]
    public void Pptx_Animation_SchemaValid()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape" });
            handler.Add("/slide[1]/shape[1]", "animation", null, new()
            {
                ["effect"] = "appear"
            });
        }

        AssertValidPptx(path, "animation");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SECTION 17 — Cross-format: Special Characters
    // ═══════════════════════════════════════════════════════════════════════

    [Fact]
    public void CrossFormat_Chinese_InPptx()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "你好世界 — Hello World" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("你好世界 — Hello World");
    }

    [Fact]
    public void CrossFormat_Chinese_InDocx()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "中文段落：测试内容" });

        // Word paragraphs use /body/p[N] path notation
        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("中文段落：测试内容");
    }

    [Fact]
    public void CrossFormat_Chinese_InXlsx()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["value"] = "中文内容" });

        var node = handler.Get("/Sheet1/A1");
        node.Text.Should().Be("中文内容");
    }

    [Fact]
    public void CrossFormat_LongText_InPptx()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new());
        var longText = string.Concat(Enumerable.Repeat("Lorem ipsum dolor sit amet. ", 20));
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = longText });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be(longText);
    }

    [Fact]
    public void CrossFormat_LongText_InDocx()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        var longText = string.Concat(Enumerable.Repeat("Word document long text sample. ", 30));
        handler.Add("/body", "paragraph", null, new() { ["text"] = longText });

        // Word paragraphs use /body/p[N] path notation
        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be(longText);
    }

    [Fact]
    public void CrossFormat_SpecialXmlChars_InDocx()
    {
        var path = CreateTemp("docx");

        var act = () =>
        {
            using var handler = new WordHandler(path, editable: true);
            handler.Add("/body", "paragraph", null, new() { ["text"] = "5 < 10 & 10 > 5 \"quoted\"" });
        };

        act.Should().NotThrow("XML special chars must be escaped properly");
    }

    [Fact]
    public void CrossFormat_SpecialXmlChars_InPptx()
    {
        var path = CreateTemp("pptx");

        var act = () =>
        {
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "A & B < C > D" });
        };

        act.Should().NotThrow("XML special chars in PPTX shape text must be handled");
    }

    [Fact]
    public void CrossFormat_SpecialXmlChars_InXlsx()
    {
        var path = CreateTemp("xlsx");

        var act = () =>
        {
            using var handler = new ExcelHandler(path, editable: true);
            handler.Add("/Sheet1", "cell", null, new() { ["address"] = "A1", ["value"] = "<bold> & \"quoted\"" });
        };

        act.Should().NotThrow("XML special chars in cell value must be handled");
    }

    [Fact]
    public void CrossFormat_Unicode_PersistsInPptx()
    {
        var path = CreateTemp("pptx");
        const string text = "日本語テスト — 한국어 — Ελληνικά";

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = text });
        }

        using (var handler = new PowerPointHandler(path, editable: false))
        {
            var node = handler.Get("/slide[1]/shape[1]");
            node.Text.Should().Be(text, "multi-language unicode must survive save/reopen");
        }
    }

    [Fact]
    public void CrossFormat_MixedContentPptx_SchemaValid()
    {
        var path = CreateTemp("pptx");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new() { ["title"] = "混合内容测试" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "English & 中文 混合 text", ["fill"] = "#4472C4" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Second shape" });
        }

        AssertValidPptx(path, "mixed content slide");
    }
}
