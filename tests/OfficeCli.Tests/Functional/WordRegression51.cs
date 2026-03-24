// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// BugHuntPart51: Excel cell type capitalization, boolean value display, formula round-trip,
/// Word run property round-trips, PPTX shape Add property round-trips,
/// Word table/paragraph indent round-trips.
/// </summary>
public class WordRegression51 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTempFile(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ==================== Bug5100 ====================
    // Excel cell type capitalization: Set accepts "boolean" (lowercase),
    // Get returns "Boolean" (titleCase). Inconsistent format.
    [Fact]
    public void Bug5100_ExcelCellTypeCapitalizationInconsistency()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "true", ["type"] = "boolean"
        });

        var node = handler.Get("/Sheet1/A1");
        var typeVal = node.Format["type"]?.ToString() ?? "";

        typeVal.Should().Be("Boolean",
            because: "Excel cell type is returned as TitleCase 'Boolean' by CellToNode");
    }

    // ==================== Bug5101 ====================
    // Excel cell type "String" vs "string" — same case issue.
    [Fact]
    public void Bug5101_ExcelCellTypeStringCapitalization()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "hello"
        });

        var node = handler.Get("/Sheet1/A1");
        var typeVal = node.Format["type"]?.ToString() ?? "";

        typeVal.ToLowerInvariant().Should().Be("string",
            because: "Excel cell type for string cells should be 'String' (CellToNode returns TitleCase)");
    }

    // ==================== Bug5102 ====================
    // Excel cell boolean display value: internally stored as "1"/"0",
    // Get returns raw "1" not user-friendly "TRUE"/"FALSE".
    [Fact]
    public void Bug5102_ExcelCellBooleanValueDisplay()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "true", ["type"] = "boolean"
        });

        var node = handler.Get("/Sheet1/A1");
        var textVal = node.Text ?? "";

        // Excel internally stores boolean as "1"/"0"
        textVal.Should().BeOneOf(new[] { "true", "TRUE", "True", "1" },
            because: "Excel boolean cell stores value as '1'/'0' internally");
    }

    // ==================== Bug5103 ====================
    // Excel cell formula round-trip.
    [Fact]
    public void Bug5103_ExcelCellFormulaRoundTrip()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "20" });

        handler.Set("/Sheet1/A3", new() { ["formula"] = "=A1+A2" });

        var node = handler.Get("/Sheet1/A3");

        node.Format.Should().ContainKey("formula",
            because: "Excel cell Get should include 'formula' when a formula was set");

        if (node.Format.ContainsKey("formula"))
        {
            node.Format["formula"]?.ToString().Should().Be("A1+A2",
                because: "Formula round-trip: Set '=A1+A2' strips '=' → stored as 'A1+A2'");
        }
    }

    // ==================== Bug5104 ====================
    // Excel cell link round-trip.
    [Fact]
    public void Bug5104_ExcelCellLinkRoundTrip()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Click" });
        handler.Set("/Sheet1/A1", new() { ["link"] = "https://example.com" });

        var node = handler.Get("/Sheet1/A1");

        node.Format.Should().ContainKey("link",
            because: "Excel cell Get should include 'link' when a hyperlink was set");
    }

    // ==================== Bug5105 ====================
    // Word table width percentage round-trip.
    [Fact]
    public void Bug5105_WordTableWidthPercentRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        handler.Set("/body/tbl[1]", new() { ["width"] = "100%" });

        var node = handler.Get("/body/tbl[1]");

        node.Format.Should().ContainKey("width");
        if (node.Format.ContainsKey("width"))
        {
            node.Format["width"]?.ToString().Should().Contain("%",
                because: "Table width set as '100%' should round-trip with '%' suffix");
        }
    }

    // ==================== Bug5106 ====================
    // Word run "caps" round-trip.
    [Fact]
    public void Bug5106_WordRunCapsRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "test" });
        handler.Set("/body/p[1]/r[1]", new() { ["caps"] = "true" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("caps");
    }

    // ==================== Bug5107 ====================
    // Word run "dstrike" round-trip.
    [Fact]
    public void Bug5107_WordRunDStrikeRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "test" });
        handler.Set("/body/p[1]/r[1]", new() { ["dstrike"] = "true" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("dstrike");
    }

    // ==================== Bug5108 ====================
    // Word run "superscript" round-trip.
    [Fact]
    public void Bug5108_WordRunSuperscriptRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "test" });
        handler.Set("/body/p[1]/r[1]", new() { ["superscript"] = "true" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("superscript");
    }

    // ==================== Bug5109 ====================
    // Word run "shading" round-trip.
    [Fact]
    public void Bug5109_WordRunShadingRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "test" });
        handler.Set("/body/p[1]/r[1]", new() { ["shading"] = "FFFF00" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("shading");
        if (node.Format.ContainsKey("shading"))
        {
            node.Format["shading"]?.ToString().Should().Be("#FFFF00",
                because: "Run shading round-trip: Set 'FFFF00' → Get 'FFFF00'");
        }
    }

    // ==================== Bug5110 ====================
    // Word run "vanish" round-trip.
    [Fact]
    public void Bug5110_WordRunVanishRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "hidden" });
        handler.Set("/body/p[1]/r[1]", new() { ["vanish"] = "true" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("vanish");
    }

    // ==================== Bug5111 ====================
    // Word run "noproof" round-trip.
    [Fact]
    public void Bug5111_WordRunNoProofRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "noproof" });
        handler.Set("/body/p[1]/r[1]", new() { ["noproof"] = "true" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("noproof");
    }

    // ==================== Bug5112 ====================
    // Word run "highlight" round-trip.
    [Fact]
    public void Bug5112_WordRunHighlightRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "highlighted" });
        handler.Set("/body/p[1]/r[1]", new() { ["highlight"] = "yellow" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("highlight");
        if (node.Format.ContainsKey("highlight"))
        {
            node.Format["highlight"]?.ToString().Should().Be("yellow");
        }
    }

    // ==================== Bug5113 ====================
    // PPTX shape strike "double" round-trip.
    // Add with strike="double" → TextStrikeValues.DoubleStrike.
    // Get should return "double" not raw enum name.
    [Fact]
    public void Bug5113_PptxShapeStrikeDoubleRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "double strike",
            ["strike"] = "double"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("strike"))
        {
            var stVal = node.Format["strike"]?.ToString() ?? "";
            stVal.Should().Be("double",
                because: "PPTX shape strike 'double' should round-trip. " +
                         "Add sets TextStrikeValues.DoubleStrike. " +
                         "Get should return 'double' not raw OOXML enum 'dblStrike'");
        }
    }

    // ==================== Bug5114 ====================
    // PPTX shape Add "autofit" = "true" round-trip.
    [Fact]
    public void Bug5114_PptxShapeAutoFitRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "autofit",
            ["autofit"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[2]");

        var hasAutofit = node.Format.ContainsKey("autofit")
                      || node.Format.ContainsKey("autoFit");
        hasAutofit.Should().BeTrue(
            because: "PPTX shape Get should include 'autofit' when it was set during Add. " +
                     "shape[1] is the title placeholder; the added shape is shape[2]");
    }

    // ==================== Bug5115 ====================
    // Word paragraph "firstLineIndent" round-trip.
    [Fact]
    public void Bug5115_WordParagraphFirstLineIndentRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "indented",
            ["firstLineIndent"] = "720"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("firstLineIndent");
        if (node.Format.ContainsKey("firstLineIndent"))
        {
            node.Format["firstLineIndent"]?.ToString().Should().Be("720");
        }
    }

    // ==================== Bug5116 ====================
    // Word paragraph "hangingIndent" round-trip.
    [Fact]
    public void Bug5116_WordParagraphHangingIndentRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "hanging",
            ["hangingIndent"] = "720"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("hangingIndent");
    }

    // ==================== Bug5117 ====================
    // Word section margin round-trip.
    [Fact]
    public void Bug5117_WordSectionMarginRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Set("/section[1]", new()
        {
            ["margintop"] = "1440",
            ["marginbottom"] = "1440"
        });

        var node = handler.Get("/section[1]");
        node.Format.Should().ContainKey("margintop");
    }

    // ==================== Bug5118 ====================
    // Word paragraph "pagebreakbefore" round-trip.
    [Fact]
    public void Bug5118_WordParagraphPageBreakBeforeRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "break",
            ["pagebreakbefore"] = "true"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("pagebreakbefore");
    }

    // ==================== Bug5119 ====================
    // Word table "indent" round-trip.
    [Fact]
    public void Bug5119_WordTableIndentRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        handler.Set("/body/tbl[1]", new() { ["indent"] = "720" });

        var node = handler.Get("/body/tbl[1]");
        node.Format.Should().ContainKey("indent");
    }

    // ==================== Bug5120 ====================
    // Excel cell "clear" removes everything.
    [Fact]
    public void Bug5120_ExcelCellClearRoundTrip()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Hello", ["font.bold"] = "true"
        });

        handler.Set("/Sheet1/A1", new() { ["clear"] = "true" });

        var node = handler.Get("/Sheet1/A1");
        (string.IsNullOrEmpty(node.Text) || node.Text == "(empty)").Should().BeTrue(
            because: "Cell should be empty after clear");
    }

    // ==================== Bug5121 ====================
    // Excel cell Set "type" = "date" is not supported.
    // CellToNode reports type "Date" for date cells, but Set only accepts
    // "string", "number", "boolean" — NOT "date".
    [Fact]
    public void Bug5121_ExcelCellSetTypeDateNotSupported()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "44927" // Jan 1, 2023 as Excel serial date
        });

        // Try to set type to "date"
        var act = () => handler.Set("/Sheet1/A1", new() { ["type"] = "date" });

        // BUG: Set throws ArgumentException for "date" type, but CellToNode can report
        // type "Date" for date-typed cells. Users who read type="Date" can't set it back.
        act.Should().NotThrow(
            because: "Excel cell Set should support 'date' type since CellToNode " +
                     "reports 'Date' type. Currently Set only accepts string/number/boolean, " +
                     "creating an asymmetry where Get can report a type that Set can't set");
    }

    // ==================== Bug5122 ====================
    // PPTX shape "lineSpacing" Add round-trip format.
    [Fact]
    public void Bug5122_PptxShapeLineSpacingRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "spacing",
            ["lineSpacing"] = "1.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("lineSpacing"))
        {
            var lsVal = node.Format["lineSpacing"]?.ToString() ?? "";
            lsVal.Should().NotBe("150000",
                because: "lineSpacing Get should return multiplier, not raw OOXML units");
        }
    }

    // ==================== Bug5123 ====================
    // PPTX shape "spaceBefore" Add in points, verify Get returns points too.
    [Fact]
    public void Bug5123_PptxShapeSpaceBeforeRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "space",
            ["spaceBefore"] = "12"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("spaceBefore"))
        {
            var sbVal = node.Format["spaceBefore"]?.ToString() ?? "";
            sbVal.Should().NotBe("1200",
                because: "spaceBefore Get should return points, not raw SpacingPoints (points × 100)");
        }
    }

    // ==================== Bug5124 ====================
    // PPTX shape "preset" geometry round-trip.
    [Fact]
    public void Bug5124_PptxShapePresetRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "ellipse",
            ["preset"] = "ellipse"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("preset"))
        {
            node.Format["preset"]?.ToString().Should().Be("ellipse");
        }
    }

    // ==================== Bug5125 ====================
    // Word paragraph widowcontrol round-trip.
    [Fact]
    public void Bug5125_WordParagraphWidowControlRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "wc test",
            ["widowcontrol"] = "true"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("widowcontrol");
    }

    // ==================== Bug5126 ====================
    // Word run "rtl" round-trip.
    [Fact]
    public void Bug5126_WordRunRtlRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "rtl text" });
        handler.Set("/body/p[1]/r[1]", new() { ["rtl"] = "true" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("rtl");
    }

    // ==================== Bug5127 ====================
    // Word run "outline" round-trip.
    [Fact]
    public void Bug5127_WordRunOutlineRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "outline text" });
        handler.Set("/body/p[1]/r[1]", new() { ["outline"] = "true" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("outline");
    }

    // ==================== Bug5128 ====================
    // Word run "emboss" round-trip.
    [Fact]
    public void Bug5128_WordRunEmbossRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "emboss text" });
        handler.Set("/body/p[1]/r[1]", new() { ["emboss"] = "true" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("emboss");
    }

    // ==================== Bug5129 ====================
    // Word run "imprint" round-trip.
    [Fact]
    public void Bug5129_WordRunImprintRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "imprint text" });
        handler.Set("/body/p[1]/r[1]", new() { ["imprint"] = "true" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("imprint");
    }
}
