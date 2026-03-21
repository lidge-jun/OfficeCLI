// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

// Fuzz tests for color-valued properties across PPTX, XLSX, DOCX handlers.
// Tests that invalid colors produce a clear ArgumentException and valid colors succeed.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Fuzz;

public class ColorFuzzer : IDisposable
{
    private readonly string _pptxPath;
    private readonly string _xlsxPath;
    private readonly string _docxPath;

    public ColorFuzzer()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"fuzz_color_{Guid.NewGuid():N}.pptx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"fuzz_color_{Guid.NewGuid():N}.xlsx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"fuzz_color_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_docxPath);

        // Pre-create a slide and shape for PPTX tests
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm", ["width"] = "10cm", ["height"] = "3cm" });
    }

    public void Dispose()
    {
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    // ==================== Valid colors — must succeed ====================

    public static IEnumerable<object[]> ValidColors => new[]
    {
        new object[] { "FF0000" },
        new object[] { "00FF00" },
        new object[] { "0000FF" },
        new object[] { "000000" },
        new object[] { "FFFFFF" },
        new object[] { "ff0000" },    // lowercase
        new object[] { "#FF0000" },   // hash prefix
        new object[] { "#ff0000" },   // hash + lowercase
        new object[] { "F00" },       // 3-char shorthand
        new object[] { "4472C4" },    // typical Office blue
        new object[] { "80FF0000" },  // AARRGGBB with alpha
        new object[] { "000" },       // 3-char shorthand for black (valid)
        new object[] { "FFF" },       // 3-char shorthand for white (valid)
        new object[] { "red" },       // named color
        new object[] { "blue" },      // named color
        new object[] { "Green" },     // named color (case-insensitive)
        new object[] { "rgb(255,0,0)" },  // rgb() notation
        new object[] { "rgb(0, 128, 255)" }, // rgb() with spaces
    };

    [Theory]
    [MemberData(nameof(ValidColors))]
    public void Pptx_SetShapeFill_ValidColor_Succeeds(string color)
    {
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["fill"] = color });
        act.Should().NotThrow($"'{color}' is a valid color");
    }

    [Theory]
    [MemberData(nameof(ValidColors))]
    public void Pptx_SetShapeColor_ValidColor_Succeeds(string color)
    {
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["color"] = color });
        act.Should().NotThrow($"'{color}' is a valid color");
    }

    [Theory]
    [MemberData(nameof(ValidColors))]
    public void Xlsx_SetCellFill_ValidColor_Succeeds(string color)
    {
        using var handler = new ExcelHandler(_xlsxPath, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "test" });
        var act = () => handler.Set("/Sheet1/A1", new() { ["fill"] = color });
        act.Should().NotThrow($"'{color}' is a valid color");
    }

    [Theory]
    [MemberData(nameof(ValidColors))]
    public void Docx_SetRunColor_ValidColor_Succeeds(string color)
    {
        using var handler = new WordHandler(_docxPath, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]/r[1]", new() { ["color"] = color });
        act.Should().NotThrow($"'{color}' is a valid color");
    }

    // ==================== Invalid colors — must throw ArgumentException ====================

    public static IEnumerable<object[]> InvalidColors => new[]
    {
        new object[] { "" },
        new object[] { "invalid" },
        new object[] { "GGGGGG" },
        new object[] { "12345" },       // 5 chars — invalid
        new object[] { "1234567" },     // 7 chars — invalid
        new object[] { "ZZZZZZ" },
        new object[] { "FF 00 00" },    // spaces
        new object[] { "transparent" },
        new object[] { "#GG0000" },
        new object[] { "rgb(256,0,0)" },  // out of range
        new object[] { "rgb(0,0)" },      // too few components
    };

    [Theory]
    [MemberData(nameof(InvalidColors))]
    public void Pptx_SetShapeFill_InvalidColor_ThrowsArgumentException(string color)
    {
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["fill"] = color });
        act.Should().Throw<ArgumentException>($"'{color}' is an invalid color and should be rejected");
    }

    [Theory]
    [MemberData(nameof(InvalidColors))]
    public void Pptx_SetShapeColor_InvalidColor_ThrowsArgumentException(string color)
    {
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["color"] = color });
        act.Should().Throw<ArgumentException>($"'{color}' is an invalid color and should be rejected");
    }

    [Theory]
    [MemberData(nameof(InvalidColors))]
    public void Xlsx_SetCellFill_InvalidColor_ThrowsArgumentException(string color)
    {
        using var handler = new ExcelHandler(_xlsxPath, editable: true);
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "test" });
        var act = () => handler.Set("/Sheet1/A1", new() { ["fill"] = color });
        act.Should().Throw<ArgumentException>($"'{color}' is an invalid color and should be rejected");
    }

    [Theory]
    [MemberData(nameof(InvalidColors))]
    public void Docx_SetRunColor_InvalidColor_ThrowsArgumentException(string color)
    {
        using var handler = new WordHandler(_docxPath, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]/r[1]", new() { ["color"] = color });
        act.Should().Throw<ArgumentException>($"'{color}' is an invalid color and should be rejected");
    }

    // ==================== ParseHelpers.SanitizeColorForOoxml unit tests ====================

    [Theory]
    [MemberData(nameof(ValidColors))]
    public void ParseHelpers_SanitizeColorForOoxml_ValidColor_DoesNotThrow(string color)
    {
        // Skip 8-char ARGB — it's valid but needs special test
        var act = () => ParseHelpers.SanitizeColorForOoxml(color);
        act.Should().NotThrow($"'{color}' is a valid color");
    }

    [Fact]
    public void ParseHelpers_SanitizeColorForOoxml_8CharArgb_ReturnsRgbWithAlpha()
    {
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml("80FF0000");
        rgb.Should().Be("FF0000");
        alpha.Should().NotBeNull();
        alpha.Should().BeInRange(0, 100000);
    }

    [Fact]
    public void ParseHelpers_SanitizeColorForOoxml_OpaqueArgb_ReturnsNullAlpha()
    {
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml("FFFF0000");
        rgb.Should().Be("FF0000");
        alpha.Should().BeNull();
    }

    [Theory]
    [InlineData("")]
    [InlineData("invalid")]
    [InlineData("GGGGGG")]
    [InlineData("12345")]
    [InlineData("1234567")]
    public void ParseHelpers_SanitizeColorForOoxml_InvalidColor_ThrowsArgumentException(string color)
    {
        var act = () => ParseHelpers.SanitizeColorForOoxml(color);
        act.Should().Throw<ArgumentException>($"'{color}' is invalid");
    }
}
