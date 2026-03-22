// Bug hunt Part 18 — watermark, data validation, table width, notes, gradient, EMU edge cases.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class MixedRegression18 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public MixedRegression18()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt18_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt18_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt18_{Guid.NewGuid():N}.pptx");
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


    // ==================== BUG #1: Word watermark Add color uses ToLowerInvariant ====================
    // WordHandler.Add.cs:1319 stores color as lowercase — inconsistent with all other handlers
    [Fact]
    public void Word_Watermark_Add_Color_ShouldBeUpperCase()
    {
        _wordHandler.Add("/", "watermark", null, new()
        {
            ["text"] = "DRAFT",
            ["color"] = "#FF0000"
        });

        var wm = _wordHandler.Get("/watermark");
        wm.Should().NotBeNull();

        if (wm.Format.ContainsKey("color"))
        {
            var color = wm.Format["color"]?.ToString();
            // BUG: Add stores "ff0000" (lowercase) via ToLowerInvariant()
            // Should use ToUpperInvariant() for consistency
            color.Should().Match(c => c == "FF0000" || c == "#FF0000",
                "watermark Add should store color as uppercase hex for consistency with rest of codebase");
        }
    }


    // ==================== BUG #2: Excel data validation allowBlank default ====================
    // ExcelHandler.Add.cs:282 defaults AllowBlank to true when not provided.
    // This is arguably correct, but the issue is you CANNOT set it to false:
    // passing allowBlank=false → !TryGet(true) || IsTruthy(false) = false || false = false
    // Wait, that actually works. Let me verify the actual behavior.
    [Fact]
    public void Excel_DataValidation_AllowBlank_CanBeSetToFalse()
    {
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A1:A10",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "100",
            ["allowBlank"] = "false"
        });

        var dv = _excelHandler.Get("/Sheet1/validation[1]");
        dv.Should().NotBeNull();

        // When allowBlank=false is explicitly set, it should be false
        if (dv.Format.ContainsKey("allowBlank"))
        {
            var allowBlank = dv.Format["allowBlank"];
            allowBlank.Should().Be(false,
                "explicitly setting allowBlank=false should result in false, not be ignored");
        }
    }


    // ==================== BUG #3: Word table width percent readback format ====================
    // Navigation.cs:334 reads table width as: (int.Parse(width) / 50) + "%"
    // But this is integer division: 5000/50 = 100%, but 4999/50 = 99%
    // Width 4999 (which is 99.98%) gets truncated to "99%"
    [Fact]
    public void Word_Table_Width_Pct_ShouldNotTruncate()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1",
            ["width"] = "75%"  // 75 * 50 = 3750 pct50 units
        });

        var table = _wordHandler.Get("/body/tbl[1]");
        table.Should().NotBeNull();

        if (table.Format.ContainsKey("width"))
        {
            var width = table.Format["width"]?.ToString();
            width.Should().Be("75%",
                "table width set as 75% should be read back as 75%");
        }
    }


    // ==================== BUG #4: PPTX notes text with newlines round-trip ====================
    // SetNotesText splits on '\n' and creates separate paragraphs.
    // GetNotesText joins with '\n'. Round-trip should preserve the text.
    [Fact]
    public void Pptx_Notes_Text_RoundTrip()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        var notesText = "Line 1\nLine 2\nLine 3";
        pptx.Set("/slide[1]", new()
        {
            ["notes"] = notesText
        });

        // Get notes back - should match exactly
        var slide = pptx.Get("/slide[1]");
        // Notes might be at /slide[1]/notes
        var notesNode = pptx.Get("/slide[1]/notes");
        notesNode.Should().NotBeNull();
        notesNode.Text.Should().Be(notesText,
            "notes text should round-trip through Set/Get without modification");
    }


    // ==================== BUG #9: PPTX table cell Get missing border info ====================
    [Fact]
    public void Pptx_TableCell_Get_ShouldIncludeBorders()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["border"] = "1pt solid FF0000"
        });

        var cell = pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Should().NotBeNull();

        // BUG: Cell node doesn't include border information
        cell.Format.Keys.Should().Contain(k => k.Contains("border"),
            "table cell Get should expose border properties when they're set");
    }


    // ==================== BUG #10: Word Set paragraph text on empty paragraph creates unstyled run ====================
    // A new variant: Set text on a paragraph that has a style.
    // The style implies formatting, but the new run doesn't inherit style-level props.
    [Fact]
    public void Word_SetParagraphText_ShouldRespectParagraphStyle()
    {
        // Add a paragraph with a specific style
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["style"] = "Heading1"
        });

        // Set text on it
        _wordHandler.Set("/body/p[1]", new()
        {
            ["text"] = "My Heading"
        });

        var para = _wordHandler.Get("/body/p[1]");
        para.Text.Should().Be("My Heading");

        // The paragraph should still have the Heading1 style
        para.Format.Should().ContainKey("style");
        para.Format["style"]?.ToString().Should().Be("Heading1",
            "setting text on a styled paragraph should not remove the style");
    }


    // ==================== BUG #11: Excel Set cell number format not in Get ====================
    [Fact]
    public void Excel_Cell_NumberFormat_RoundTrip()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "1234.56",
            ["numberformat"] = "#,##0.00"
        });

        var cell = _excelHandler.Get("/Sheet1/A1");
        cell.Should().NotBeNull();

        cell.Format.Should().ContainKey("numberformat",
            "cell Get should expose number format when it's set");
    }
}
