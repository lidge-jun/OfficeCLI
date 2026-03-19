// Bug hunt Part 14 — more bugs found through deep code review.
// PPTX table issues, Excel conditional formatting, Word path and persistence bugs.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart14 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public BugHuntPart14()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt14_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt14_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt14_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
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


    // ==================== BUG #1: Word Set paragraph liststyle on subsequent calls doesn't update ====================
    // When a paragraph already has a list style and you Set a new liststyle,
    // the old NumberingProperties should be replaced, not duplicated.
    [Fact]
    public void Word_SetParagraph_ListStyle_ShouldNotDuplicate()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Item 1",
            ["liststyle"] = "bullet"
        });

        var before = _wordHandler.Get("/body/p[1]");
        before.Format.Should().ContainKey("listStyle");
        before.Format["listStyle"]?.ToString().Should().Be("bullet");

        // Now change to ordered list
        _wordHandler.Set("/body/p[1]", new()
        {
            ["liststyle"] = "ordered"
        });

        var after = _wordHandler.Get("/body/p[1]");
        // BUG: The listStyle may not update properly or may create duplicate
        // NumberingProperties in the paragraph
        after.Format.Should().ContainKey("listStyle");
        after.Format["listStyle"]?.ToString().Should().Be("ordered",
            "changing liststyle from bullet to ordered should update, not accumulate");
    }


    // ==================== BUG #2: Word Set font on paragraph applies to runs but new runs lack it ====================
    // When you Set font on a paragraph, it's applied to existing runs only.
    // If you then Add a new run, the new run has no font setting.
    // The paragraph should carry "default run properties" but doesn't.
    [Fact]
    public void Word_SetParagraphFont_ThenAddRun_NewRunShouldInheritFont()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Hello"
        });

        // Set font on all runs in the paragraph
        _wordHandler.Set("/body/p[1]", new()
        {
            ["font"] = "Courier New"
        });

        // Verify first run got the font
        var before = _wordHandler.Get("/body/p[1]", depth: 2);
        before.Children.Count.Should().BeGreaterThan(0);
        before.Children[0].Format.Should().ContainKey("font");
        before.Children[0].Format["font"]?.ToString().Should().Be("Courier New");

        // Add a new run to the paragraph
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = " World"
        });

        // The new run should inherit the paragraph's font
        var after = _wordHandler.Get("/body/p[1]", depth: 2);
        after.Children.Count.Should().Be(2);
        var newRun = after.Children[1];

        // BUG: The new run has no font property — it uses default font
        // The paragraph-level font set via Set doesn't become the default for new runs
        newRun.Format.Should().ContainKey("font",
            "new runs added after Set font on paragraph should inherit the paragraph font");
    }


    // ==================== BUG #4: Excel colorscale color for 8-char ARGB ====================
    // Same normalization bug as databar: (length == 6 ? "FF" : "") + color
    // For an 8-char ARGB like "80FF0000" (semi-transparent red), it's stored correctly (no prefix).
    // But for a 4-char or 5-char typo, it's stored as-is without validation.
    [Fact]
    public void Excel_ColorScale_InvalidHexLength_ShouldBeHandled()
    {
        _excelHandler.Add("/Sheet1", "colorscale", null, new()
        {
            ["sqref"] = "A1:A10",
            ["mincolor"] = "F00",   // 3-char — invalid, should be expanded or rejected
            ["maxcolor"] = "0F0"    // 3-char — invalid
        });

        var cf = _excelHandler.Get("/Sheet1/cf[1]");
        cf.Should().NotBeNull();

        // BUG: 3-char hex "F00" becomes "F00" (no "FF" prefix since length != 6)
        // This is not a valid ARGB color — should be "FFFF0000" or rejected
        if (cf.Format.ContainsKey("mincolor"))
        {
            var minColor = cf.Format["mincolor"]?.ToString();
            minColor.Should().HaveLength(8,
                "color should be a valid 8-char ARGB hex string");
        }
    }


    // ==================== BUG #5: PPTX Add shape returns path but Set modifies wrong shape ====================
    // When shapes are added/removed, the index-based paths become stale.
    // The shape ID or name should be used for stable identification.
    [Fact]
    public void Pptx_Shape_IndexPath_IsFragileAfterRemove()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shape A",
            ["fill"] = "FF0000"
        });
        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shape B",
            ["fill"] = "00FF00"
        });

        // Verify both shapes exist
        var shapeA = pptx.Get("/slide[1]/shape[1]");
        shapeA.Text.Should().Be("Shape A");
        var shapeB = pptx.Get("/slide[1]/shape[2]");
        shapeB.Text.Should().Be("Shape B");

        // Remove shape A
        pptx.Remove("/slide[1]/shape[1]");

        // Now shape B should be at index 1
        var newShape1 = pptx.Get("/slide[1]/shape[1]");
        newShape1.Text.Should().Be("Shape B",
            "after removing shape[1], former shape[2] should become the new shape[1]");
    }


    // ==================== BUG #6: Excel Set cell link doesn't validate URL ====================
    // ExcelHandler.Set.cs:548 creates URI with: new Uri(value)
    // Invalid URIs will throw. Should handle gracefully.
    [Fact]
    public void Excel_SetCell_Link_WithInvalidUri_ShouldNotCrash()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Click" });

        // Try to set an invalid link (no scheme)
        var act = () => _excelHandler.Set("/Sheet1/A1", new()
        {
            ["link"] = "not a valid url"
        });

        // BUG: new Uri("not a valid url") throws UriFormatException
        // Should either accept it as relative URI or return a clear error
        act.Should().NotThrow<UriFormatException>(
            "invalid URLs should be handled gracefully, not throw unhandled UriFormatException");
    }


    // ==================== BUG #7: PPTX Set table row height uses EMU but Query returns formatted ====================
    // PowerPointHandler.Set.cs:426 uses ParseEmu(value) for row height
    // But the table row height in Get doesn't include the height per row in the Format
    // This makes it impossible to verify Set worked without raw XML access.
    [Fact]
    public void Pptx_Table_RowHeight_SetThenGet_ShouldRoundTrip()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "1"
        });

        // Set specific row height
        pptx.Set("/slide[1]/table[1]/tr[1]", new()
        {
            ["height"] = "1cm"
        });

        // Get the table and check row height
        var table = pptx.Get("/slide[1]/table[1]", depth: 2);
        table.Should().NotBeNull();
        table.Children.Count.Should().BeGreaterThanOrEqualTo(2);

        // BUG: The table row node doesn't include height in its Format
        // so there's no way to verify the Set worked via Get
        var row1 = table.Children[0];
        row1.Format.Should().ContainKey("height",
            "table row should expose height property so Set can be verified via Get");
    }


    // ==================== BUG #8: Excel cell clear doesn't reset style ====================
    // ExcelHandler.Set.cs:531-534 clears CellValue, CellFormula, DataType
    // but does NOT clear StyleIndex. So the cell appears empty but retains
    // its background color, font, borders, etc.
    [Fact]
    public void Excel_CellClear_ShouldResetStyle()
    {
        // Set cell with value and style
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Hello",
            ["bgcolor"] = "FFFF00",
            ["font.bold"] = "true"
        });

        // Verify style was applied (Get returns fill under "fill" key)
        var before = _excelHandler.Get("/Sheet1/A1");
        (before.Format.ContainsKey("fill") || before.Format.ContainsKey("bgcolor")).Should().BeTrue(
            "fill color should be applied");

        // Clear the cell
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["clear"] = "true"
        });

        var after = _excelHandler.Get("/Sheet1/A1");

        // BUG: clear resets value/formula/type but not StyleIndex
        // The cell still shows yellow background even though content is cleared
        after.Format.Should().NotContainKey("fill",
            "clearing a cell should also reset its style/formatting, not just content");
        after.Format.Should().NotContainKey("bgcolor",
            "clearing a cell should also reset its style/formatting, not just content");
    }


    // ==================== BUG #9: Word bookmark Add returns wrong path format ====================
    // WordHandler.Add.cs:793 returns "/bookmark[{bkName}]" — a root-level path
    // But the bookmark is actually inside a paragraph, not at document root.
    // The returned path doesn't reflect the actual document hierarchy.
    [Fact]
    public void Word_Add_Bookmark_ReturnedPath_ShouldReflectHierarchy()
    {
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Some text" });

        var bkPath = _wordHandler.Add("/body/p[1]", "bookmark", null, new()
        {
            ["name"] = "mybookmark",
            ["text"] = "marked text"
        });

        // The bookmark path should include its parent paragraph
        // BUG: Returns "/bookmark[mybookmark]" instead of "/body/p[1]/bookmark[mybookmark]"
        bkPath.Should().Contain("/body/p",
            "bookmark path should include the parent paragraph in the hierarchy, " +
            "not be a flat root-level path like /bookmark[name]");
    }


    // ==================== BUG #10: PPTX shape font size inconsistency between Set and Get ====================
    // Set uses (int)(ParseFontSize(value) * 100) — truncates
    // Get uses fontSize.Value / 100.0 — precise float division
    // For size "10.5": Set stores (int)(10.5 * 100) = 1050. Get reads 1050/100.0 = "10.5pt" ✓
    // For size "10.55": Set stores (int)(10.55 * 100) = 1055. Get reads 1055/100.0 = "10.55pt" ✓
    // But for floating point edge cases: (int)(10.005 * 100) might be 1000 due to FP representation
    [Fact]
    public void Pptx_Shape_FontSize_FloatingPointPrecision()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test"
        });

        // Set a font size that may have floating point issues
        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["size"] = "10.005"
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        var sizeStr = shape.Format.ContainsKey("size") ? shape.Format["size"]?.ToString() : null;

        // (int)(10.005 * 100) could be 1000 instead of 1001 due to FP representation
        // 10.005 in IEEE 754 is actually 10.004999999999999...
        // So (int)(10.005 * 100) = (int)(1000.4999...) = 1000
        // Then Get reads 1000/100.0 = "10pt" instead of "10.01pt"
        if (sizeStr != null)
        {
            sizeStr.Should().Be("10.01pt",
                "font size 10.005 should round to nearest hundredth, not truncate via (int) cast");
        }
    }


    // ==================== BUG #11: PPTX Set shape text replaces all paragraphs ====================
    // Similar to Word bug: when Set text on a shape, all paragraphs are replaced.
    // But the shape might have multi-paragraph content (bullets, etc.) that gets destroyed.
    [Fact]
    public void Pptx_SetShape_Text_DestroysMultiParagraphContent()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line1\\nLine2\\nLine3"  // Multi-line text creates multiple paragraphs
        });

        var before = pptx.Get("/slide[1]/shape[1]", depth: 2);
        before.Should().NotBeNull();

        // Set bold on the shape (applies to all runs)
        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["bold"] = "true"
        });

        // Now replace text
        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Single line"
        });

        var after = pptx.Get("/slide[1]/shape[1]", depth: 2);
        after.Text.Should().Contain("Single line");

        // BUG: The bold formatting applied to the old paragraphs is destroyed
        // when text is replaced, because new paragraphs are created fresh
        after.Format.Should().ContainKey("bold",
            "bold formatting set before text replacement should be preserved");
    }
}
