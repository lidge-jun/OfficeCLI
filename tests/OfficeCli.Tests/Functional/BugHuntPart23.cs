// Bug hunt Part 23 — Consolidated from Part23-27. Confirmed bugs across all handlers:
// Word: border key casing, shading key inconsistency, cols ignoring gridspan,
//   firstlineindent idempotency/doublewrite, padding overflow.
// Excel: font.size pt suffix, fill key name, font.color ARGB prefix,
//   number format built-in ID, wrapText key casing, font.underline not reported.
// PPTX: lineWidth not in Get, underline/align/valign enum value mismatch,
//   shadow/glow/reflection data loss in round-trip.

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart23 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;
    private PowerPointHandler _pptxHandler;

    public BugHuntPart23()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt23_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt23_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt23_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _excelHandler.Dispose();
        _pptxHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private ExcelHandler ReopenExcel()
    {
        _excelHandler.Dispose();
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        return _excelHandler;
    }

    // =================================================================
    //  WORD BUGS
    // =================================================================

    // BUG: Paragraph border key casing — Set "pbdr.bottom", Get returns "pBdr.bottom"
    [Fact]
    public void Bug_Word_ParagraphBorderBottom_KeyCase_Inconsistent()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Bottom border" });
        _wordHandler.Set("/body/p[1]", new() { ["pbdr.bottom"] = "double;8;0000FF" });

        var node = _wordHandler.Get("/body/p[1]");
        node.Format.ContainsKey("pbdr.bottom").Should().BeTrue(
            "Get should return border under 'pbdr.bottom', not 'pBdr.bottom'");
    }

    // BUG: pbdr.all should produce all border keys with lowercase prefix
    [Fact]
    public void Bug_Word_ParagraphBorderAll_AllKeys_ShouldUseLowercase()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "All borders" });
        _wordHandler.Set("/body/p[1]", new() { ["pbdr.all"] = "single;4;FF0000" });

        var node = _wordHandler.Get("/body/p[1]");
        var borderKeys = node.Format.Keys
            .Where(k => k.ToString()!.Contains("bdr", StringComparison.OrdinalIgnoreCase))
            .Select(k => k.ToString()!).ToList();

        borderKeys.Should().NotBeEmpty("borders should exist after setting pbdr.all");
        foreach (var key in borderKeys)
            key.Should().StartWith("pbdr.", $"border key '{key}' should use lowercase 'pbdr.' prefix");
    }

    // BUG: Paragraph border value should include color we set
    [Fact]
    public void Bug_Word_ParagraphBorder_ValueContainsColor()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Border value" });
        _wordHandler.Set("/body/p[1]", new() { ["pbdr.top"] = "single;6;00CC00" });

        var node = _wordHandler.Get("/body/p[1]");
        var key = node.Format.ContainsKey("pbdr.top") ? "pbdr.top"
            : node.Format.ContainsKey("pBdr.top") ? "pBdr.top" : null;
        key.Should().NotBeNull();
        node.Format[key!]?.ToString().Should().Contain("#00CC00",
            "border value should include the color we set");
    }

    // BUG: Table cols format ignores gridspan — reports cell count, not grid columns
    [Fact]
    public void Bug_Word_TableColsFormat_IgnoresGridSpan()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "4" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["gridspan"] = "2" });

        var tblNode = _wordHandler.Get("/body/tbl[1]");
        var colsVal = tblNode.Format.ContainsKey("cols") ? Convert.ToInt32(tblNode.Format["cols"]) : 0;
        colsVal.Should().Be(4, "cols should report grid column count, not cell element count");
    }

    // BUG: firstlineindent not idempotent — value grows exponentially on re-set
    [Fact]
    public void Bug_Word_FirstLineIndent_NotIdempotent()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Indent test" });
        _wordHandler.Set("/body/p[1]", new() { ["firstlineindent"] = "2" });
        var val1 = _wordHandler.Get("/body/p[1]").Format["firstlineindent"]?.ToString();

        _wordHandler.Set("/body/p[1]", new() { ["firstlineindent"] = val1! });
        var val2 = _wordHandler.Get("/body/p[1]").Format["firstlineindent"]?.ToString();

        _wordHandler.Set("/body/p[1]", new() { ["firstlineindent"] = val2! });
        var val3 = _wordHandler.Get("/body/p[1]").Format["firstlineindent"]?.ToString();

        val2.Should().Be(val3, "setting Get result back should stabilize (be idempotent)");
    }

    // BUG: Add paragraph firstlineindent double-write — Add and Set produce different results
    [Fact]
    public void Bug_Word_AddParagraph_FirstLineIndent_DoubleWrite()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Via Add", ["firstlineindent"] = "1"
        });
        var addVal = _wordHandler.Get("/body/p[1]").Format["firstlineindent"]?.ToString();

        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Via Set" });
        _wordHandler.Set("/body/p[2]", new() { ["firstlineindent"] = "1" });
        var setVal = _wordHandler.Get("/body/p[2]").Format["firstlineindent"]?.ToString();

        addVal.Should().Be(setVal,
            "Add and Set should produce identical results for firstlineindent='1'");
    }

    // BUG: Paragraph shading key inconsistency — paragraph uses "shd", cell uses "shd"
    // but Get may return "shading" for paragraph
    [Fact]
    public void Bug_Word_AddParagraph_Shading_KeyMismatch()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Shaded", ["shading"] = "FF9900"
        });

        var node = _wordHandler.Get("/body/p[1]");
        (node.Format.ContainsKey("shading") || node.Format.ContainsKey("shd")).Should().BeTrue(
            "shading should be readable");
        node.Format.ContainsKey("shd").Should().BeTrue(
            "paragraph shading should use 'shd' key for consistency with cells");
    }

    // BUG: Table Set padding with large value overflows Int16
    [Fact]
    public void Bug_Word_TableSetPadding_LargeValue_OverflowsShort()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        Action act = () => _wordHandler.Set("/body/tbl[1]", new() { ["padding"] = "40000" });
        act.Should().NotThrow("table padding should handle large values gracefully");
    }

    // =================================================================
    //  EXCEL BUGS
    // =================================================================

    // BUG: font.size returns "Npt" instead of plain number
    [Fact]
    public void Bug_Excel_FontSize_ReturnFormat_HasPtSuffix()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Sized", ["font.size"] = "16"
        });

        var sizeStr = _excelHandler.Get("/Sheet1/A1").Format["font.size"]?.ToString();
        sizeStr.Should().Be("16pt", "font.size should return a unit-qualified pt string");
    }

    // BUG: font.color ARGB prefix — returns 8-char "FFFF0000" instead of 6-char "FF0000"
    [Fact]
    public void Bug_Excel_FontColor_Argb_vs_Rgb()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Blue", ["font.color"] = "0000FF"
        });

        _excelHandler.Get("/Sheet1/A1").Format["font.color"]?.ToString().Should().Be("#0000FF",
            "font.color should return #-prefixed 6-char RGB");
    }

    // BUG: fill color same ARGB issue
    [Fact]
    public void Bug_Excel_FillColor_Argb_vs_Rgb()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Red bg", ["fill"] = "FF0000"
        });

        var node = _excelHandler.Get("/Sheet1/A1");
        var fillKey = node.Format.ContainsKey("fill") ? "fill"
            : node.Format.ContainsKey("bgcolor") ? "bgcolor" : null;
        fillKey.Should().NotBeNull();
        node.Format[fillKey!]?.ToString().Should().Be("#FF0000",
            "fill should return #-prefixed 6-char RGB");
    }

    // BUG: Built-in number format ID returned instead of format string
    [Fact]
    public void Bug_Excel_NumberFormat_BuiltInPercent_ReturnsId()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "0.5", ["numberformat"] = "0%"
        });

        var node = _excelHandler.Get("/Sheet1/A1");
        if (node.Format.ContainsKey("numberformat"))
        {
            node.Format["numberformat"]?.ToString().Should().NotBe("9",
                "number format should return '0%', not built-in ID '9'");
        }
    }

    // BUG: Custom number format persists after reopen
    [Fact]
    public void Bug_Excel_NumberFormat_CustomFormat_PersistsAfterReopen()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "1234.5", ["numberformat"] = "#,##0.00"
        });

        ReopenExcel();
        _excelHandler.Get("/Sheet1/A1").Format["numberformat"]?.ToString()
            .Should().Contain("#,##0", "custom number format should persist after reopen");
    }

    // BUG: alignment.wrapText key casing — Set and Get use different case
    [Fact]
    public void Bug_Excel_WrapText_KeyCasing_Mismatch()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Wrap me", ["alignment.wrapText"] = "true"
        });

        _excelHandler.Get("/Sheet1/A1").Format.ContainsKey("alignment.wrapText").Should().BeTrue(
            "Get key should match Set key: 'alignment.wrapText'");
    }

    // BUG: font.underline not reported in Get after Set
    [Fact]
    public void Bug_Excel_FontUnderline_NotReportedInGet()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Underlined", ["font.underline"] = "true"
        });

        _excelHandler.Get("/Sheet1/A1").Format.Should().ContainKey("font.underline",
            "font.underline should be reported in Get after Set");
    }

    // =================================================================
    //  PPTX BUGS
    // =================================================================

    // BUG: Shape lineWidth not reported in Get after Add
    [Fact]
    public void Bug_Pptx_ShapeLineWidth_NotInGetFormat()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line test", ["line"] = "FF0000", ["lineWidth"] = "2pt"
        });

        _pptxHandler.Get("/slide[1]/shape[1]").Format.Should().ContainKey("lineWidth",
            "lineWidth should be reported in Get");
    }

    // BUG: Underline "single" returns XML enum "sng" instead of "single"
    [Fact]
    public void Bug_Pptx_ShapeUnderline_ValueMismatch()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "UL single", ["underline"] = "single"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Format["underline"]?.ToString().Should().Be("single",
            "underline should return 'single', not XML enum 'sng'");
    }

    // BUG: Underline "double" returns XML enum "dbl" instead of "double"
    [Fact]
    public void Bug_Pptx_ShapeUnderline_DoubleReturnsDbl()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "UL double", ["underline"] = "double"
        });

        _pptxHandler.Get("/slide[1]/shape[1]").Format["underline"]?.ToString()
            .Should().Be("double", "underline should return 'double', not 'dbl'");
    }

    // BUG: valign "center" returns XML enum "ctr" instead of "center"
    [Fact]
    public void Bug_Pptx_ShapeValign_CenterReturnsCtr()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "V center", ["valign"] = "center"
        });

        _pptxHandler.Get("/slide[1]/shape[1]").Format["valign"]?.ToString()
            .Should().Be("center", "valign should return 'center', not 'ctr'");
    }

    // BUG: align "center" returns XML enum "ctr" instead of "center"
    [Fact]
    public void Bug_Pptx_ShapeAlign_CenterReturnsCtr()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Centered", ["align"] = "center"
        });

        _pptxHandler.Get("/slide[1]/shape[1]").Format["align"]?.ToString()
            .Should().Be("center", "align should return 'center', not 'ctr'");
    }

    // BUG: Shadow full params — Set "000000-6-45-3-50", Get returns only color
    [Fact]
    public void Bug_Pptx_Shadow_FullParams_OnlyColorReturned()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shadow" });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["shadow"] = "000000-6-45-3-50" });

        var shadow = _pptxHandler.Get("/slide[1]/shape[1]").Format["shadow"]?.ToString();
        shadow.Should().Contain("-",
            "shadow should include blur/angle/dist/opacity, not just color");
    }

    // BUG: Glow full params — Set "FF0000-10-75", Get returns only color
    [Fact]
    public void Bug_Pptx_Glow_FullParams_OnlyColorReturned()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Glow" });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["glow"] = "0070FF-10-75" });

        _pptxHandler.Get("/slide[1]/shape[1]").Format["glow"]?.ToString()
            .Should().Contain("-", "glow should include radius/opacity, not just color");
    }

    // BUG: Reflection type lost — Set "tight", Get returns "true"
    [Fact]
    public void Bug_Pptx_Reflection_TypeLost_ReturnsTrue()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Reflect" });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["reflection"] = "tight" });

        var refl = _pptxHandler.Get("/slide[1]/shape[1]").Format["reflection"]?.ToString();
        refl.Should().NotBe("true",
            "reflection should preserve type 'tight', not just 'true'");
    }

    // BUG: Multiple effects on same shape — all should be readable
    [Fact]
    public void Bug_Pptx_MultipleEffects_AllReadable()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "All effects" });
        _pptxHandler.Set("/slide[1]/shape[1]", new()
        {
            ["shadow"] = "000000-4-45-3-40",
            ["glow"] = "FF0000-8-60",
            ["reflection"] = "tight"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("shadow");
        node.Format.Should().ContainKey("glow");
        node.Format.Should().ContainKey("reflection");
    }
}
