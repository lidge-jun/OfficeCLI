// FuzzRound3 — Extended fuzz coverage, new areas not covered in rounds 1–2.
//
// Areas:
//   F70: Word Set font size — extreme values, negative, NaN strings, empty
//   F71: Word Set spacing — invalid units, NaN, negative, overflow strings
//   F72: Word Set color — invalid hex, out-of-range, empty, null-like strings
//   F73: Excel Set border — invalid styles, empty, combinations; valid styles don't throw
//   F74: Excel Set numberformat — empty string, very long format, special chars
//   F75: Excel Set conditional format — boundary rule index, invalid type
//   F76: PPTX Add table — invalid rows/cols (0, negative, non-numeric, missing)
//   F77: PPTX Set transition — invalid type, empty, boundary duration
//   F78: PPTX Set animation — invalid effect on shape, boundary timing
//   C01: Concurrent-style: same file open/close multiple times sequentially
//   C02: Large Add batch — 50 shapes then Query (performance / no crash)
//   R01: Regression guard — previously fixed ArgumentException paths accept valid values

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound3 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz3_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) File.Delete(f); } catch { }
    }

    // ==================== F70: Word Set font size boundary values ====================

    [Theory]
    [InlineData("1")]        // minimum valid
    [InlineData("1pt")]
    [InlineData("72")]       // 1 inch
    [InlineData("72pt")]
    [InlineData("10.5pt")]   // fractional
    [InlineData("0.5")]      // sub-1 still parseable
    public void F70_Word_RunSetSize_ValidValues_DoNotThrow(string size)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]/r[1]", new() { ["size"] = size });
        act.Should().NotThrow($"Word run size '{size}' is a valid value");
    }

    [Theory]
    [InlineData("abc")]      // non-numeric
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("-Infinity")]
    [InlineData("1,5")]      // comma decimal separator
    public void F70_Word_RunSetSize_InvalidNonNumeric_ThrowsArgumentException(string size)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]/r[1]", new() { ["size"] = size });
        act.Should().Throw<ArgumentException>($"Word run size '{size}' should throw ArgumentException");
    }

    [Fact]
    public void F70_Word_RunSetSize_SpaceBeforePt_IsAccepted()
    {
        // ParseFontSize trims whitespace, so "12 pt" → "12" → valid
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]/r[1]", new() { ["size"] = "12 pt" });
        act.Should().NotThrow("'12 pt' with space before 'pt' is trimmed and accepted by ParseFontSize");
    }

    [Theory]
    [InlineData("-1")]       // negative font size
    [InlineData("-0.5pt")]
    public void F70_Word_RunSetSize_NegativeValues_ThrowOrNoThrow_NoUnhandledException(string size)
    {
        // Negative font sizes: parser may accept or reject — must not crash with unhandled exception
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        try
        {
            handler.Set("/body/p[1]/r[1]", new() { ["size"] = size });
            // If it doesn't throw, verify document is still valid by re-reading
            var node = handler.Get("/body/p[1]/r[1]");
            node.Should().NotBeNull();
        }
        catch (ArgumentException)
        {
            // acceptable
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            Assert.Fail($"Unexpected exception type {ex.GetType().Name} for size '{size}': {ex.Message}");
        }
    }

    // ==================== F71: Word Set spacing invalid units ====================

    [Theory]
    [InlineData("12pt")]
    [InlineData("0.5cm")]
    [InlineData("0.25in")]
    [InlineData("0")]
    [InlineData("0pt")]
    public void F71_Word_ParagraphSpacing_ValidValues_DoNotThrow(string spacing)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act1 = () => handler.Set("/body/p[1]", new() { ["spaceBefore"] = spacing });
        var act2 = () => handler.Set("/body/p[1]", new() { ["spaceAfter"] = spacing });
        act1.Should().NotThrow($"spaceBefore='{spacing}' is valid");
        act2.Should().NotThrow($"spaceAfter='{spacing}' is valid");
    }

    [Theory]
    [InlineData("12px")]        // px not a valid Word unit
    [InlineData("12em")]        // em not supported
    [InlineData("notaunit")]
    [InlineData("abc pt")]
    public void F71_Word_ParagraphSpacing_InvalidUnits_ThrowOrNoThrow_NoUnhandledException(string spacing)
    {
        // Invalid units: must not crash with NullReferenceException/OverflowException
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        try
        {
            handler.Set("/body/p[1]", new() { ["spaceBefore"] = spacing });
        }
        catch (ArgumentException)
        {
            // acceptable
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            Assert.Fail($"Unexpected exception {ex.GetType().Name} for spaceBefore='{spacing}': {ex.Message}");
        }
    }

    [Theory]
    [InlineData("-12pt")]       // negative spacing — should throw
    [InlineData("-1cm")]
    public void F71_Word_ParagraphSpacing_NegativeValues_ThrowArgumentException(string spacing)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]", new() { ["spaceBefore"] = spacing });
        act.Should().Throw<ArgumentException>($"Negative spaceBefore='{spacing}' should throw");
    }

    // ==================== F72: Word Set color invalid values ====================

    [Theory]
    [InlineData("FF0000")]
    [InlineData("#FF0000")]
    [InlineData("red")]
    [InlineData("rgb(255,0,0)")]
    [InlineData("F00")]          // 3-char shorthand
    public void F72_Word_RunSetColor_ValidValues_DoNotThrow(string color)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]/r[1]", new() { ["color"] = color });
        act.Should().NotThrow($"Word run color '{color}' is valid");
    }

    [Theory]
    [InlineData("ZZZZZZ")]      // invalid hex letters
    [InlineData("12345")]       // 5 hex chars (not 3 or 6)
    [InlineData("not-a-color")]
    [InlineData("#GGGGGG")]
    public void F72_Word_RunSetColor_InvalidValues_ThrowOrNoThrow_NoUnhandledException(string color)
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        try
        {
            handler.Set("/body/p[1]/r[1]", new() { ["color"] = color });
        }
        catch (ArgumentException)
        {
            // acceptable — invalid color rejected
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            Assert.Fail($"Unexpected exception {ex.GetType().Name} for color='{color}': {ex.Message}");
        }
    }

    // ==================== F73: Excel Set border valid/invalid styles ====================

    [Theory]
    [InlineData("thin")]
    [InlineData("medium")]
    [InlineData("thick")]
    [InlineData("dashed")]
    [InlineData("dotted")]
    [InlineData("double")]
    [InlineData("none")]
    [InlineData("hair")]
    [InlineData("mediumdashed")]
    public void F73_Excel_SetBorder_ValidStyles_DoNotThrow(string style)
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Set("/Sheet1/A1", new() { ["value"] = "test" });
        var act = () => handler.Set("/Sheet1/A1", new() { ["border.all"] = style });
        act.Should().NotThrow($"Excel border style '{style}' is valid");
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("bold")]
    [InlineData("1px")]
    public void F73_Excel_SetBorder_InvalidStyles_ThrowArgumentException(string style)
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Set("/Sheet1/A1", new() { ["value"] = "test" });
        var act = () => handler.Set("/Sheet1/A1", new() { ["border.all"] = style });
        act.Should().Throw<ArgumentException>($"Excel border style '{style}' should throw ArgumentException");
    }

    [Theory]
    [InlineData("THICK")]      // case-insensitive — should be accepted
    [InlineData("THIN")]
    [InlineData("DASHED")]
    public void F73_Excel_SetBorder_UppercaseStyles_AcceptedCaseInsensitive(string style)
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Set("/Sheet1/A1", new() { ["value"] = "test" });
        var act = () => handler.Set("/Sheet1/A1", new() { ["border.all"] = style });
        act.Should().NotThrow($"Excel border style '{style}' is case-insensitive and should be accepted");
    }

    [Fact]
    public void F73_Excel_SetBorder_WithColor_ValidCombination_DoNotThrow()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Set("/Sheet1/B2", new() { ["value"] = "hello" });
        var act = () => handler.Set("/Sheet1/B2", new() {
            ["border.all"] = "thin",
            ["border.color"] = "FF0000"
        });
        act.Should().NotThrow("border.all + border.color is a valid combination");
    }

    [Fact]
    public void F73_Excel_SetBorder_IndividualSides_ValidCombination_DoNotThrow()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Set("/Sheet1/C3", new() { ["value"] = "data" });
        var act = () => handler.Set("/Sheet1/C3", new() {
            ["border.left"] = "thin",
            ["border.right"] = "medium",
            ["border.top"] = "dashed",
            ["border.bottom"] = "dotted"
        });
        act.Should().NotThrow("individual border sides with different styles is valid");
    }

    // ==================== F74: Excel Set numberformat boundary values ====================

    [Theory]
    [InlineData("General")]
    [InlineData("0.00")]
    [InlineData("#,##0.00")]
    [InlineData("m/d/yy")]
    [InlineData("@")]           // text format
    [InlineData("0%")]
    [InlineData("\"text\"@")]   // format with literal string
    public void F74_Excel_SetNumberFormat_ValidFormats_DoNotThrow(string fmt)
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Set("/Sheet1/A1", new() { ["value"] = "123" });
        var act = () => handler.Set("/Sheet1/A1", new() { ["numberformat"] = fmt });
        act.Should().NotThrow($"Excel numberformat '{fmt}' should be accepted");
    }

    [Fact]
    public void F74_Excel_SetNumberFormat_EmptyString_DoesNotCrash()
    {
        // Empty format string may be accepted or rejected — must not crash with unhandled exception
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Set("/Sheet1/A1", new() { ["value"] = "123" });
        try
        {
            handler.Set("/Sheet1/A1", new() { ["numberformat"] = "" });
        }
        catch (ArgumentException) { /* acceptable */ }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            Assert.Fail($"Unexpected exception {ex.GetType().Name} for empty numberformat: {ex.Message}");
        }
    }

    [Fact]
    public void F74_Excel_SetNumberFormat_VeryLongFormat_DoesNotCrash()
    {
        // Very long format string: should not overflow or cause unhandled exception
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Set("/Sheet1/A1", new() { ["value"] = "123" });
        var longFmt = new string('#', 1000) + ".00";
        try
        {
            handler.Set("/Sheet1/A1", new() { ["numberformat"] = longFmt });
        }
        catch (ArgumentException) { /* acceptable */ }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            Assert.Fail($"Unexpected exception {ex.GetType().Name} for long numberformat: {ex.Message}");
        }
    }

    // ==================== F75: Excel Set conditional format boundary values ====================

    [Fact]
    public void F75_Excel_SetCF_OutOfRangeIndex_ThrowsArgumentException()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        // CF[999] — no conditional formats exist, should throw ArgumentException
        var act = () => handler.Set("/Sheet1/cf[999]", new() { ["type"] = "cellIs" });
        act.Should().Throw<ArgumentException>("CF index 999 is out of range");
    }

    [Fact]
    public void F75_Excel_SetCF_ZeroIndex_ThrowsArgumentException()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        var act = () => handler.Set("/Sheet1/cf[0]", new() { ["type"] = "cellIs" });
        act.Should().Throw<ArgumentException>("CF index 0 is invalid (1-based)");
    }

    [Fact]
    public void F75_Excel_SetCF_InvalidType_ThrowsArgumentException()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        // Add a CF first
        try
        {
            handler.Add("/Sheet1/A1:C3", "conditionalformat", null,
                new() { ["type"] = "cellIs", ["operator"] = "greaterThan", ["value"] = "10", ["fill"] = "FFFF00" });
            var act = () => handler.Set("/Sheet1/cf[1]", new() { ["type"] = "invalidType" });
            act.Should().Throw<ArgumentException>("invalid CF type should throw");
        }
        catch (ArgumentException)
        {
            // If Add CF is not supported, skip
        }
    }

    // ==================== F76: PPTX Add table invalid rows/cols ====================

    [Theory]
    [InlineData("0", "3")]      // zero rows
    [InlineData("3", "0")]      // zero cols
    [InlineData("-1", "3")]     // negative rows
    [InlineData("3", "-1")]     // negative cols
    public void F76_Pptx_AddTable_InvalidRowsCols_ThrowArgumentException(string rows, string cols)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        var act = () => handler.Add("/slide[1]", "table", null,
            new() { ["rows"] = rows, ["cols"] = cols });
        act.Should().Throw<ArgumentException>($"rows={rows},cols={cols} should be invalid");
    }

    [Theory]
    [InlineData("abc", "3")]    // non-numeric rows
    [InlineData("3", "xyz")]    // non-numeric cols
    public void F76_Pptx_AddTable_NonNumericRowsCols_ThrowArgumentException(string rows, string cols)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        var act = () => handler.Add("/slide[1]", "table", null,
            new() { ["rows"] = rows, ["cols"] = cols });
        act.Should().Throw<ArgumentException>($"non-numeric rows='{rows}'/cols='{cols}' should throw ArgumentException");
    }

    [Fact]
    public void F76_Pptx_AddTable_MissingRows_UsesDefaultOf3()
    {
        // rows defaults to "3" when not specified — no throw, creates 3x3 table
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        var act = () => handler.Add("/slide[1]", "table", null, new() { ["cols"] = "3" });
        act.Should().NotThrow("missing rows defaults to 3, should not throw");
    }

    [Fact]
    public void F76_Pptx_AddTable_ValidMinimum_Succeeds()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        var act = () => handler.Add("/slide[1]", "table", null,
            new() { ["rows"] = "1", ["cols"] = "1" });
        act.Should().NotThrow("1x1 table is minimum valid size");
    }

    // ==================== F77: PPTX Set transition boundary values ====================

    [Theory]
    [InlineData("fade")]
    [InlineData("cut")]
    [InlineData("dissolve")]
    [InlineData("wipe")]
    [InlineData("none")]
    public void F77_Pptx_SetTransition_ValidTypes_DoNotThrow(string transition)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        var act = () => handler.Set("/slide[1]", new() { ["transition"] = transition });
        act.Should().NotThrow($"PPTX transition '{transition}' is valid");
    }

    [Theory]
    [InlineData("invalid_transition")]
    [InlineData("fly")]
    public void F77_Pptx_SetTransition_InvalidTypes_ThrowArgumentException(string transition)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        var act = () => handler.Set("/slide[1]", new() { ["transition"] = transition });
        act.Should().Throw<ArgumentException>($"PPTX transition '{transition}' should throw ArgumentException");
    }

    [Theory]
    [InlineData("FADE")]         // case-insensitive — accepted
    [InlineData("CUT")]
    [InlineData("Dissolve")]
    public void F77_Pptx_SetTransition_UppercaseTypes_AcceptedCaseInsensitive(string transition)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        var act = () => handler.Set("/slide[1]", new() { ["transition"] = transition });
        act.Should().NotThrow($"PPTX transition '{transition}' is case-insensitive and should be accepted");
    }

    [Theory]
    [InlineData("0")]        // zero duration
    [InlineData("1000")]     // large duration in ms
    [InlineData("500")]      // normal
    public void F77_Pptx_SetTransitionDur_ValidValues_DoNotThrow(string dur)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Set("/slide[1]", new() { ["transition"] = "fade" });
        var act = () => handler.Set("/slide[1]", new() { ["transitiondur"] = dur });
        act.Should().NotThrow($"transitiondur={dur} should be valid");
    }

    [Theory]
    [InlineData("-500")]     // negative duration
    [InlineData("abc")]      // non-numeric
    [InlineData("NaN")]
    public void F77_Pptx_SetTransitionDur_InvalidValues_ThrowOrNoThrow_NoUnhandledException(string dur)
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Set("/slide[1]", new() { ["transition"] = "fade" });
        try
        {
            handler.Set("/slide[1]", new() { ["transitiondur"] = dur });
        }
        catch (ArgumentException) { /* acceptable */ }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            Assert.Fail($"Unexpected exception {ex.GetType().Name} for transitiondur='{dur}': {ex.Message}");
        }
    }

    // ==================== F78: PPTX Set animation boundary values ====================

    [Fact]
    public void F78_Pptx_AddAnimation_ValidEffect_DoNotThrow()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello", ["x"] = "100", ["y"] = "100", ["width"] = "200", ["height"] = "100" });
        var act = () => handler.Add("/slide[1]/shape[1]", "animation", null,
            new() { ["effect"] = "appear", ["trigger"] = "click" });
        act.Should().NotThrow("animation effect 'appear' with trigger 'click' is valid");
    }

    [Fact]
    public void F78_Pptx_AddAnimation_InvalidEffect_DoesNotCrashWithUnhandledException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello", ["x"] = "100", ["y"] = "100", ["width"] = "200", ["height"] = "100" });
        try
        {
            handler.Add("/slide[1]/shape[1]", "animation", null,
                new() { ["effect"] = "totally_invalid_effect", ["trigger"] = "click" });
        }
        catch (ArgumentException) { /* acceptable */ }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            Assert.Fail($"Unexpected exception {ex.GetType().Name} for invalid animation effect: {ex.Message}");
        }
    }

    // ==================== C01: Sequential open/close same file ====================

    [Fact]
    public void C01_Word_SequentialReopenSameFile_NoCorruption()
    {
        var path = CreateTemp("docx");

        // First open: add paragraph
        using (var h1 = new WordHandler(path, editable: true))
        {
            h1.Add("/body", "paragraph", null, new() { ["text"] = "First open" });
        }

        // Second open: add another paragraph, verify first still there
        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Add("/body", "paragraph", null, new() { ["text"] = "Second open" });
            var nodes = h2.Query("paragraph");
            nodes.Should().HaveCountGreaterOrEqualTo(2, "both paragraphs should be present after reopen");
        }

        // Third open: read-only, verify both
        using var h3 = new WordHandler(path, editable: false);
        var finalNodes = h3.Query("paragraph");
        finalNodes.Should().HaveCountGreaterOrEqualTo(2, "paragraphs persist across multiple opens");
    }

    [Fact]
    public void C01_Excel_SequentialReopenSameFile_NoCorruption()
    {
        var path = CreateTemp("xlsx");

        using (var h1 = new ExcelHandler(path, editable: true))
            h1.Set("/Sheet1/A1", new() { ["value"] = "round1" });

        using (var h2 = new ExcelHandler(path, editable: true))
        {
            h2.Set("/Sheet1/A2", new() { ["value"] = "round2" });
            var n1 = h2.Get("/Sheet1/A1");
            n1.Should().NotBeNull();
            n1!.Text.Should().Be("round1", "A1 value persists after reopen");
        }

        using var h3 = new ExcelHandler(path, editable: false);
        var n2 = h3.Get("/Sheet1/A2");
        n2.Should().NotBeNull();
        n2!.Text.Should().Be("round2");
    }

    [Fact]
    public void C01_Pptx_SequentialReopenSameFile_NoCorruption()
    {
        var path = CreateTemp("pptx");

        using (var h1 = new PowerPointHandler(path, editable: true))
            h1.Add("/", "slide", null, new() { ["title"] = "Slide 1" });

        using (var h2 = new PowerPointHandler(path, editable: true))
        {
            h2.Add("/", "slide", null, new() { ["title"] = "Slide 2" });
            var slides = h2.Query("/slide");
            slides.Should().HaveCountGreaterOrEqualTo(2, "slides persist across reopen");
        }

        using var h3 = new PowerPointHandler(path, editable: false);
        var finalSlides = h3.Query("/slide");
        finalSlides.Should().HaveCountGreaterOrEqualTo(2);
    }

    // ==================== C02: Large Add batch + Query performance ====================

    [Fact]
    public void C02_Pptx_Add50Shapes_QueryAllSucceeds()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Batch test" });

        for (int i = 0; i < 50; i++)
        {
            handler.Add("/slide[1]", "shape", null, new() {
                ["text"] = $"Shape {i}",
                ["x"] = (i * 10).ToString(),
                ["y"] = (i * 5).ToString(),
                ["width"] = "50",
                ["height"] = "30"
            });
        }

        var shapes = handler.Query("/slide[1]/shape");
        shapes.Should().HaveCountGreaterOrEqualTo(50, "all 50 shapes should be queryable");
    }

    [Fact]
    public void C02_Excel_Set100Cells_QueryAllSucceeds()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);

        for (int i = 1; i <= 100; i++)
            handler.Set($"/Sheet1/A{i}", new() { ["value"] = $"row{i}", ["bold"] = "true" });

        // Query rows — should not crash
        var rows = handler.Query("/Sheet1/row");
        rows.Should().HaveCountGreaterOrEqualTo(100, "all 100 rows should be present");
    }

    [Fact]
    public void C02_Word_Add30Paragraphs_QueryAllSucceeds()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);

        for (int i = 1; i <= 30; i++)
            handler.Add("/body", "paragraph", null, new() { ["text"] = $"Paragraph {i}" });

        var paras = handler.Query("paragraph");
        paras.Should().HaveCountGreaterOrEqualTo(30, "all 30 paragraphs should be queryable");
    }

    // ==================== R01: Regression guard — ArgumentException paths accept valid values ====================

    [Fact]
    public void R01_Word_SectionBreak_ValidTypes_DoNotThrow()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "para1" });

        foreach (var bt in new[] { "nextPage", "continuous", "evenPage", "oddPage" })
        {
            var act = () => handler.Set("/section[1]", new() { ["type"] = bt });
            act.Should().NotThrow($"section type='{bt}' is valid");
        }
    }

    [Fact]
    public void R01_Word_SectionBreak_InvalidType_ThrowsArgumentException()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        var act = () => handler.Set("/section[1]", new() { ["type"] = "diagonal" });
        act.Should().Throw<ArgumentException>("invalid section type 'diagonal' should throw");
    }

    [Fact]
    public void R01_Excel_BorderStyle_CaseInsensitive_UpperCaseAccepted()
    {
        // Excel border style is case-insensitive (ToLowerInvariant before switch)
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Set("/Sheet1/A1", new() { ["value"] = "test" });
        var act = () => handler.Set("/Sheet1/A1", new() { ["border.all"] = "THIN" });
        act.Should().NotThrow("Excel border style 'THIN' is case-insensitive and should be accepted");
    }

    [Fact]
    public void R01_Pptx_TableAddToNonSlide_ThrowsArgumentException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        var act = () => handler.Add("/", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        act.Should().Throw<ArgumentException>("Adding table to '/' (not a slide) should throw");
    }

    [Fact]
    public void R01_Word_Underline_DoubleValue_DoesNotThrow()
    {
        // Regression: "double" underline was broken (IsTruthy("double") = false), now fixed
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]/r[1]", new() { ["underline"] = "double" });
        act.Should().NotThrow("underline='double' should be accepted (regression guard)");

        var node = handler.Get("/body/p[1]/r[1]");
        node.Should().NotBeNull();
    }

    [Fact]
    public void R01_LineSpacing_MultiplierFormat_RoundTrip()
    {
        // Regression guard: lineSpacing input "1.5x" must come back as "1.5x" on Get
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello", ["lineSpacing"] = "1.5x" });
        var node = handler.Get("/body/p[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("lineSpacing");
        node.Format["lineSpacing"].ToString().Should().Be("1.5x", "lineSpacing round-trip: 1.5x → 1.5x");
    }

    [Fact]
    public void R01_LineSpacing_PercentFormat_RoundTrip()
    {
        // "150%" input → should read back as "1.5x"
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello", ["lineSpacing"] = "150%" });
        var node = handler.Get("/body/p[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("lineSpacing");
        node.Format["lineSpacing"].ToString().Should().Be("1.5x", "lineSpacing 150% → 1.5x canonical form");
    }

    [Fact]
    public void R01_ColorOutput_HasHashPrefix()
    {
        // Colors returned from Get must have # prefix
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        handler.Set("/body/p[1]/r[1]", new() { ["color"] = "FF0000" });
        var node = handler.Get("/body/p[1]/r[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("color");
        node.Format["color"].ToString().Should().StartWith("#", "colors returned from Get must have # prefix");
    }

    [Fact]
    public void R01_Pptx_ColorOutput_HasHashPrefix()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        // Add a shape with explicit fill color
        handler.Add("/slide[1]", "shape", null, new() {
            ["text"] = "Hello", ["fill"] = "4472C4",
            ["x"] = "100", ["y"] = "100", ["width"] = "200", ["height"] = "100"
        });
        // Find the non-title shape — query shapes and pick the last one (title is first)
        var allShapes = handler.Query("/slide[1]/shape");
        var shapePath = allShapes.Last().Path;
        var node = handler.Get(shapePath);
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("fill", "shape added with fill='4472C4' must expose fill in Format");
        node.Format["fill"].ToString().Should().StartWith("#", "PPTX fill color from Get must have # prefix");
    }
}
