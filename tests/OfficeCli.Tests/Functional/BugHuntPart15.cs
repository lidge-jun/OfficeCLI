// Bug hunt Part 15 — persistence, table cell, cross-handler edge cases.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart15 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public BugHuntPart15()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt15_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt15_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt15_{Guid.NewGuid():N}.pptx");
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


    // ==================== BUG #1: Word table cell Set text+font in same call order issue ====================
    // WordHandler.Set.cs:809-862 — for table cells, "text" removes old runs and creates
    // a new plain run. "font" modifies existing runs. If "text" is processed before "font",
    // the new run gets the font. But if "font" is processed first, it applies to old runs,
    // then "text" creates a new run without font.
    [Fact]
    public void Word_TableCell_SetTextAndFont_ShouldApplyBoth()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        // Set both text and font in a single call
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["font"] = "Courier New",
            ["text"] = "Hello Cell"
        });

        var cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]", depth: 2);
        cell.Text.Should().Contain("Hello Cell");

        // Check if the run has the font
        var para = cell.Children.FirstOrDefault();
        if (para?.Children.Count > 0)
        {
            var run = para.Children[0];
            // BUG: If "font" was iterated before "text", font was applied to old run,
            // then "text" created a new run without the font
            run.Format.Should().ContainKey("font",
                "setting font and text together on a table cell should apply font to the final text");
        }
    }


    // ==================== BUG #2: PPTX Add shape with font appends duplicate LatinFont ====================
    // PowerPointHandler.Add.cs:124 does rProps.Append(new Drawing.LatinFont {...})
    // But CreateTextShape may have already set a LatinFont on the RunProperties.
    // This creates duplicate LatinFont children, making the Run invalid.
    [Fact]
    public void Pptx_AddShape_WithFont_ShouldNotDuplicateLatinFont()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello",
            ["font"] = "Arial Black"
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        shape.Should().NotBeNull();
        shape.Format.Should().ContainKey("font");

        // The font should be "Arial Black", not something else due to duplicate elements
        shape.Format["font"]?.ToString().Should().Be("Arial Black");

        // Validate the document to check for schema violations from duplicates
        var errors = pptx.Validate();
        var fontErrors = errors.Where(e =>
            e.Description.Contains("LatinFont", StringComparison.OrdinalIgnoreCase) ||
            e.Description.Contains("duplicate", StringComparison.OrdinalIgnoreCase)).ToList();

        fontErrors.Should().BeEmpty(
            "shape created with font should not have duplicate LatinFont elements in RunProperties");
    }


    // ==================== BUG #4: Excel Set multiple cells in different sheets ====================
    // When setting cells in Sheet2, the handler needs to find the correct worksheet.
    // Path format is /Sheet2/A1 — but if Sheet2 doesn't exist yet, it should error clearly.
    [Fact]
    public void Excel_Set_Cell_InNonexistentSheet_ShouldGiveClearError()
    {
        var act = () => _excelHandler.Set("/NonExistentSheet/A1", new()
        {
            ["value"] = "Hello"
        });

        // Should throw a clear ArgumentException, not a NullReferenceException
        act.Should().Throw<ArgumentException>(
            "setting a cell in a non-existent sheet should give a clear error")
            .WithMessage("*not found*");
    }


    // ==================== BUG #5: Word Set table cell text creates run without SpacePreserve ====================
    // When setting text on a table cell, the new Run's Text element should have
    // Space = SpaceProcessingModeValues.Preserve to keep leading/trailing spaces.
    // Let's verify this works for text with leading spaces.
    [Fact]
    public void Word_TableCell_SetText_ShouldPreserveSpaces()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "  spaced  "
        });

        var cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        // The text should preserve leading and trailing spaces
        cell.Text.Should().Contain("  spaced  ",
            "text with leading/trailing spaces should be preserved in table cell");
    }


    // ==================== BUG #6: PPTX table cell Get doesn't include alignment ====================
    // The table cell Get handler (Query.cs:229-259) returns text, fill, font, size,
    // bold, italic, color — but NOT alignment. Setting alignment on a cell has no
    // way to verify it via Get.
    [Fact]
    public void Pptx_TableCell_Get_ShouldIncludeAlignment()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Centered",
            ["alignment"] = "center"
        });

        var cell = pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Text.Should().Contain("Centered");

        // BUG: The cell node Format doesn't include alignment
        cell.Format.Should().ContainKey("alignment",
            "table cell Get should expose alignment property so Set can be verified");
    }


    // ==================== BUG #8: Word table cell Set font on empty cell is lost ====================
    // Same pattern as paragraph: Set font iterates over existing runs.
    // Empty cell has no runs → font is silently lost.
    [Fact]
    public void Word_TableCell_SetFont_OnEmptyCell_IsLost()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        // Set font on the empty cell (default cells have empty paragraphs, no runs)
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["font"] = "Courier New"
        });

        // Then add text
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Hello"
        });

        var cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]", depth: 2);
        cell.Text.Should().Contain("Hello");

        // BUG: Font was set on zero runs (cell was empty), then text created a plain run
        var para = cell.Children.FirstOrDefault();
        if (para?.Children.Count > 0)
        {
            para.Children[0].Format.Should().ContainKey("font",
                "font set on empty cell should be applied when text is subsequently added");
        }
    }


    // ==================== BUG #10: Word ViewAsStats may crash on empty document ====================
    // ViewAsStats iterates over paragraphs, runs, etc. On a truly empty document
    // (no body or no paragraphs), this could throw NullReferenceException.
    [Fact]
    public void Word_ViewAsStats_ShouldNotCrashOnMinimalDoc()
    {
        var act = () => _wordHandler.ViewAsStats();
        act.Should().NotThrow("ViewAsStats should handle minimal/empty documents gracefully");

        var stats = _wordHandler.ViewAsStats();
        stats.Should().NotBeNullOrEmpty();
    }


    // ==================== BUG #11: PPTX shape Add with empty text creates run ====================
    // When text="" is passed to Add shape, it still creates a Run with empty Text.
    // This creates unnecessary DOM elements that could affect formatting.
    [Fact]
    public void Pptx_AddShape_EmptyText_ShouldNotCreateEmptyRun()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = ""
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        shape.Should().NotBeNull();

        // The shape text should be empty
        (shape.Text ?? "").Should().BeEmpty();

        // But more importantly, setting font size on this "empty" shape should work
        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["size"] = "24"
        });

        // Then set actual text
        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Now has text"
        });

        var updated = pptx.Get("/slide[1]/shape[1]");
        updated.Text.Should().Contain("Now has text");

        // The font size should be preserved
        updated.Format.Should().ContainKey("size",
            "font size set before text should be preserved on the shape");
    }
}
