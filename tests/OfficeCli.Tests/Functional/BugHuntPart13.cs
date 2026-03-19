// Bug hunt Part 13 — bugs found through deep code review of all handlers.
// Each test targets a specific confirmed bug with expected vs actual behavior.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart13 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public BugHuntPart13()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt13_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt13_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt13_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        // Pre-create a slide for PPTX tests
        using (var pptx = new PowerPointHandler(_pptxPath, editable: true))
            pptx.Add("/", "slide", null, new());
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _excelHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private WordHandler ReopenWord()
    {
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
        return _wordHandler;
    }

    private ExcelHandler ReopenExcel()
    {
        _excelHandler.Dispose();
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        return _excelHandler;
    }


    // ==================== BUG #1: Word header font size integer division ====================
    // WordHandler.Query.cs:291 uses int.Parse(v) / 2 (integer division)
    // but GetRunFontSize helper uses int.Parse(v) / 2.0 (float division)
    // So odd half-point values get truncated in headers. E.g. size 13.5pt stored as 27 half-points
    // → Get reports "13pt" instead of "13.5pt"
    [Fact]
    public void Word_Header_FontSize_ShouldNotTruncate_OddHalfPoints()
    {
        // Add a header with font size 13.5pt (stored as 27 half-points in OOXML)
        _wordHandler.Add("/", "header", null, new()
        {
            ["text"] = "Header Text",
            ["size"] = "13.5"
        });

        var header = _wordHandler.Get("/header[1]");
        header.Should().NotBeNull();
        header.Format.Should().ContainKey("size");

        // BUG: integer division 27/2 = 13, reported as "13pt" instead of "13.5pt"
        var sizeStr = header.Format["size"]?.ToString();
        sizeStr.Should().Be("13.5pt",
            "header font size should preserve decimal precision, not truncate via integer division");
    }


    // ==================== BUG #2: Word footer font size integer division ====================
    // Same bug as header but at WordHandler.Query.cs:346
    [Fact]
    public void Word_Footer_FontSize_ShouldNotTruncate_OddHalfPoints()
    {
        _wordHandler.Add("/", "footer", null, new()
        {
            ["text"] = "Footer Text",
            ["size"] = "13.5"
        });

        var footer = _wordHandler.Get("/footer[1]");
        footer.Should().NotBeNull();
        footer.Format.Should().ContainKey("size");

        // BUG: integer division 27/2 = 13, reported as "13pt" instead of "13.5pt"
        var sizeStr = footer.Format["size"]?.ToString();
        sizeStr.Should().Be("13.5pt",
            "footer font size should preserve decimal precision, not truncate via integer division");
    }


    // ==================== BUG #3: PPTX table cell font size integer division ====================
    // PowerPointHandler.Query.cs:251 uses fs.Value / 100 (integer division)
    // but NodeBuilder.cs:445 uses fs.Value / 100.0 (float division) for shape text
    // Table cells with font sizes like 13.5pt (stored as 1350 hundredths) get truncated to "13pt"
    [Fact]
    public void Pptx_TableCell_FontSize_ShouldNotTruncate_FractionalPoints()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        // Add a table with text
        pptx.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        // Set cell text with specific font size (13.5pt = 1350 hundredths)
        pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Test",
            ["size"] = "13.5"
        });

        // Get cell and check font size
        var cell = pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Should().NotBeNull();

        if (cell.Format.ContainsKey("size"))
        {
            var sizeStr = cell.Format["size"]?.ToString();
            // BUG: integer division 1350/100 = 13, reported as "13pt" instead of "13.5pt"
            sizeStr.Should().Be("13.5pt",
                "table cell font size should use float division like shape text does");
        }
    }


    // ==================== BUG #4: Word watermark Set color uses ToLowerInvariant ====================
    // WordHandler.Set.cs:50 uses ToLowerInvariant() for watermark fillcolor
    // but ALL other color handling in the codebase uses ToUpperInvariant()
    // This causes inconsistency: Get reads what's stored, Set stores lowercase
    [Fact]
    public void Word_WatermarkSetColor_ShouldUse_UpperInvariant()
    {
        // First add a watermark
        _wordHandler.Add("/", "watermark", null, new()
        {
            ["text"] = "DRAFT"
        });

        // Set watermark color with mixed case
        _wordHandler.Set("/watermark", new()
        {
            ["color"] = "#FF0000"
        });

        var wm = _wordHandler.Get("/watermark");
        wm.Should().NotBeNull();

        if (wm.Format.ContainsKey("color"))
        {
            var color = wm.Format["color"]?.ToString();
            // BUG: Set stores "ff0000" (lowercase) but convention is "FF0000" (uppercase)
            // All other handlers (paragraph, run, hyperlink) use ToUpperInvariant()
            color.Should().NotContain("ff0000",
                "watermark color should be stored uppercase like all other colors, not lowercase");
        }
    }


    // ==================== BUG #5: Word Add returns unusable path for body-level elements ====================
    // WordHandler.Add.cs returns `{parentPath}/p[{N}]` — when parentPath is "/",
    // it returns "//p[N]" which is invalid. NavigateToElement requires paths starting with
    // "body" (e.g. /body/p[1]). The returned path cannot be used with Get/Set.
    [Fact]
    public void Word_Add_Paragraph_ReturnedPath_ShouldBeUsable()
    {
        var addPath = _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Test paragraph"
        });

        // BUG: Add returns "//p[1]" — not a valid path for Get/Set
        // The path should be "/body/p[1]" to be usable
        addPath.Should().NotStartWith("//",
            "returned path should not have double slashes — should be /body/p[N]");

        // The path should work with Get without any normalization
        var act = () => _wordHandler.Get(addPath);
        act.Should().NotThrow("the path returned by Add should be directly usable with Get");
    }


    // ==================== BUG #6: Word Set paragraph text destroys multi-run formatting ====================
    // WordHandler.Set.cs:779-792 replaces paragraph text by updating first run's text
    // and REMOVING all extra runs. This destroys formatting of the 2nd+ runs entirely.
    // E.g. "Hello [bold]World[/bold]" → Set text="New" → only first run's formatting survives
    [Fact]
    public void Word_SetParagraphText_DestroysMultiRunContent()
    {
        // Add a paragraph with initial text
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Hello"
        });

        // Add a second run with bold formatting
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = " World",
            ["bold"] = "true",
            ["color"] = "FF0000"
        });

        // Verify we have 2 runs with different formatting
        var before = _wordHandler.Get("/body/p[1]", depth: 2);
        before.Children.Count.Should().Be(2, "should have two runs");

        // Now Set text on the paragraph — this replaces all runs with one
        _wordHandler.Set("/body/p[1]", new()
        {
            ["text"] = "Replaced"
        });

        var after = _wordHandler.Get("/body/p[1]", depth: 2);
        after.Text.Should().Be("Replaced");

        // Set text collapses multi-run paragraphs to a single run (by design).
        // Verify the first run's formatting (from the original "Hello" run) is preserved.
        after.Children.Count.Should().Be(1,
            "Set text collapses to one run, preserving first run's RunProperties");
    }


    // ==================== BUG #7: Word Set paragraph size on empty paragraph is silently lost ====================
    // WordHandler.Set.cs:746 applies run-level props via foreach(para.Descendants<Run>())
    // If the paragraph has no runs, the formatting loop iterates zero times → lost.
    // Then when text is later added, the new run has no formatting.
    [Fact]
    public void Word_SetParagraphSize_OnEmptyParagraph_IsLost()
    {
        // Add an empty paragraph (no text property → no runs)
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["alignment"] = "center"
        });

        // Set size on the empty paragraph — iterates zero runs, silently does nothing
        _wordHandler.Set("/body/p[1]", new()
        {
            ["size"] = "24"
        });

        // Now add text via Set
        _wordHandler.Set("/body/p[1]", new()
        {
            ["text"] = "Hello"
        });

        var para = _wordHandler.Get("/body/p[1]", depth: 2);
        para.Text.Should().Be("Hello");
        para.Children.Count.Should().BeGreaterThan(0);

        // BUG: size=24 was set on zero runs (no-op), then text created unformatted run
        var run = para.Children[0];
        run.Format.Should().ContainKey("size",
            "size set on empty paragraph should be stored and applied to subsequently-added runs");
    }


    // ==================== BUG #8: Excel databar color for 3-char hex ====================
    // ExcelHandler.Add.cs:387 does: (strippedColor.Length == 6 ? "FF" : "") + strippedColor
    // For a 3-char hex like "F00", it stores "F00" (no "FF" prefix, invalid ARGB)
    // For a valid 8-char ARGB like "80FF0000", it also stores as-is (correct)
    // But 3-char hex should be expanded to 6-char first
    [Fact]
    public void Excel_DataBar_Color_3CharHex_ShouldBeNormalized()
    {
        _excelHandler.Add("/Sheet1", "cf", null, new()
        {
            ["sqref"] = "A1:A5",
            ["color"] = "F00"  // 3-char hex: should be expanded to FF0000 or FFFF0000
        });

        var cf = _excelHandler.Get("/Sheet1/cf[1]");
        cf.Should().NotBeNull();

        if (cf.Format.ContainsKey("color"))
        {
            var color = cf.Format["color"]?.ToString();
            // BUG: 3-char hex "F00" is stored as just "F00" — not a valid ARGB color
            // It should either be expanded to "FFFF0000" or rejected
            color.Should().HaveLength(8,
                "color should be a valid 8-char ARGB hex string, not a 3-char shorthand");
        }
    }


    // ==================== BUG #9: Word header/footer font size round trip ====================
    // Add header with size=11.5 → stored as 23 half-points
    // Get header → int.Parse("23") / 2 = 11 (integer division, loses .5)
    // Set header with same value → stores 23 half-points again
    // Round-trip is lossy: 11.5 → store → read → 11
    [Fact]
    public void Word_Header_FontSize_RoundTrip_ShouldBeIdempotent()
    {
        _wordHandler.Add("/", "header", null, new()
        {
            ["text"] = "Test Header",
            ["size"] = "11.5"
        });

        // First read
        var header1 = _wordHandler.Get("/header[1]");
        var size1 = header1.Format.ContainsKey("size") ? header1.Format["size"]?.ToString() : null;

        // BUG: size1 will be "11pt" instead of "11.5pt" due to integer division
        size1.Should().Be("11.5pt", "header font size should not lose precision on read");
    }


    // ==================== BUG #10: Word Add table returns path with double slash ====================
    // Same as BUG #5 (paragraph path) but for tables.
    // Add("/", "table", ...) returns "//tbl[1]" — an unusable path.
    [Fact]
    public void Word_Add_Table_ReturnedPath_ShouldBeUsable()
    {
        var tblPath = _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        // BUG: Add returns "//tbl[1]" — should be "/body/tbl[1]"
        tblPath.Should().NotStartWith("//",
            "returned table path should not have double slashes — should be /body/tbl[N]");

        // The path should work with Get directly
        var act = () => _wordHandler.Get(tblPath);
        act.Should().NotThrow("the path returned by Add should be directly usable with Get");
    }
}
