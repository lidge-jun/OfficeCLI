// Bug hunt tests Part 11: Word Add run color #, bookmark rename collision,
// PPTX shape missing Get properties, Excel comment/validation edge cases.
// All bugs verified by running tests — every test in this file SHOULD FAIL.

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class WordRegression11 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public WordRegression11()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.pptx");
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

    // ===========================================================================================
    // CATEGORY A: Word Add run color missing TrimStart('#')
    // Add.cs line 295: rColor.ToUpperInvariant() — no TrimStart('#')
    // Compare with paragraph Add line 167 which correctly does TrimStart('#')
    // ===========================================================================================

    // BUG #1601: Word Add run color doesn't strip #
    [Fact]
    public void Bug1601_Word_Add_RunColor_HashNotStripped()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "Red run",
            ["color"] = "#FF0000"
        });

        var raw = _wordHandler.Raw("/document");

        // BUG: XML contains w:val="#FF0000" instead of w:val="FF0000"
        // Add.cs line 295: newRProps.Color = new Color { Val = rColor.ToUpperInvariant() }
        raw.Should().NotContain("w:val=\"#FF0000\"",
            "run color should strip # prefix, " +
            "but WordHandler.Add.cs line 295 does ToUpperInvariant() without TrimStart('#')");
    }

    // BUG #1602: Word Add run color vs paragraph Add color inconsistency
    [Fact]
    public void Bug1602_Word_Add_RunVsParagraph_ColorInconsistency()
    {
        // Paragraph Add with color (uses TrimStart on line 167)
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Para text",
            ["color"] = "#00FF00"
        });

        // Run Add with same color (missing TrimStart on line 295)
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "Run text",
            ["color"] = "#00FF00"
        });

        var raw = _wordHandler.Raw("/document");

        // Count occurrences of the color values - they should all be "00FF00" (no #)
        var hashCount = System.Text.RegularExpressions.Regex.Matches(raw, "#00FF00").Count;

        hashCount.Should().Be(0,
            "neither paragraph nor run color should contain # in XML, " +
            "but Add.cs line 295 (run) doesn't strip # while line 167 (paragraph) does");
    }

    // ===========================================================================================
    // CATEGORY B: Word bookmark rename — no collision detection
    // Set.cs line 357-358: bkStart.Name = value; — no uniqueness check
    // ===========================================================================================

    // BUG #1603: Word Set bookmark rename allows duplicate names
    [Fact]
    public void Bug1603_Word_Set_BookmarkRename_AllowsDuplicates()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "First para" });
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Second para" });

        _wordHandler.Add("/body/p[1]", "bookmark", null, new() { ["name"] = "BM_First" });
        _wordHandler.Add("/body/p[2]", "bookmark", null, new() { ["name"] = "BM_Second" });

        // Rename second bookmark to same name as first — should fail but doesn't
        var act = () => _wordHandler.Set("/bookmark[BM_Second]", new()
        {
            ["name"] = "BM_First"
        });

        // BUG: No validation — silently creates duplicate bookmark names
        act.Should().Throw<Exception>(
            "renaming a bookmark to an existing name should throw, " +
            "but Set.cs line 357-358 silently allows duplicate bookmark names");
    }

    // ===========================================================================================
    // CATEGORY C: PPTX shape Get — missing paragraph-level format keys
    // NodeBuilder.cs only reads first run's properties. Paragraph alignment,
    // line spacing, etc. are not returned.
    // ===========================================================================================

    // ===========================================================================================
    // CATEGORY D: Word Add paragraph with firstlineindent
    // Add.cs handles leftindent, rightindent, hanging but check firstlineindent
    // ===========================================================================================

    // BUG #1606: Word Add paragraph firstlineindent support
    [Fact]
    public void Bug1606_Word_Add_Paragraph_FirstLineIndentSupported()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Indented first line",
            ["firstlineindent"] = "720"
        });

        var raw = _wordHandler.Raw("/document");

        raw.Should().Contain("w:firstLine=\"720\"",
            "Word Add paragraph should support 'firstlineindent' property");
    }

    // ===========================================================================================
    // CATEGORY F: Excel Set "wrap" property on cell
    // ===========================================================================================

    // ===========================================================================================
    // CATEGORY J: FormulaParser — \text command round-trip
    // ===========================================================================================

    // BUG #1611: FormulaParser \text command round-trip
    [Fact]
    public void Bug1611_FormulaParser_Text_RoundTrip()
    {
        var omml = FormulaParser.Parse("x = \\text{hello world}");
        var latex = FormulaParser.ToLatex(omml);

        latex.Should().Contain("\\text",
            "FormulaParser should round-trip \\text command back to LaTeX");
    }

    // ===========================================================================================
    // CATEGORY K: Word Add hyperlink — missing italic support
    // Add.cs hyperlink only supports font/size, not italic
    // ===========================================================================================

    // BUG #1612: Word Add hyperlink italic not supported
    [Fact]
    public void Bug1612_Word_Add_Hyperlink_ItalicNotSupported()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });

        _wordHandler.Add("/body/p[1]", "hyperlink", null, new()
        {
            ["url"] = "https://example.com",
            ["text"] = "Italic link",
            ["italic"] = "true"
        });

        var raw = _wordHandler.Raw("/document");

        raw.Should().Contain("w:i ",
            "hyperlink Add should support italic property, " +
            "but Add.cs hyperlink section only handles font and size");
    }
}
