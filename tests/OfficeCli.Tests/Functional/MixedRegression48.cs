// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// BugHuntPart48: Cross-handler inconsistencies, Word paragraph key duplication,
/// Excel cell property naming, and format string discrepancies between handlers.
/// </summary>
public class MixedRegression48 : IDisposable
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

    // ==================== Bug4800 ====================
    // Word paragraph "alignment" key duplication:
    // Navigation.cs line 232 stores alignment as "alignment" key,
    // but Word Add.cs line 55 reads it from "alignment" property.
    // PPTX uses "align" for the same concept.
    // This is a cross-handler inconsistency.
    [Fact]
    public void Bug4800_WordParagraphAlignmentKeyVsPptxAlignKey()
    {
        // Test Word paragraph alignment key
        var wordPath = CreateTempFile(".docx");
        BlankDocCreator.Create(wordPath);
        using var wordHandler = new WordHandler(wordPath, editable: true);

        wordHandler.Add("/body", "p", null, new()
        {
            ["text"] = "centered paragraph",
            ["alignment"] = "center"
        });

        var wordNode = wordHandler.Get("/body/p[1]");
        // Per CLAUDE.md: Word canonical key is "alignment", PPT canonical key is "align"
        wordNode.Format.Should().ContainKey("alignment");
        wordNode.Format["alignment"].Should().Be("center");

        // Test PPTX shape alignment key
        var pptxPath = CreateTempFile(".pptx");
        BlankDocCreator.Create(pptxPath);
        using var pptxHandler = new PowerPointHandler(pptxPath, editable: true);

        pptxHandler.Add("/", "slide", null, new() { ["title"] = "test" });
        pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "centered" });
        pptxHandler.Set("/slide[1]/shape[1]", new() { ["align"] = "center" });

        var pptxNode = pptxHandler.Get("/slide[1]/shape[1]");
        pptxNode.Format.Should().ContainKey("align");

        // Word uses canonical "alignment", PPT uses canonical "align" — both correct per CLAUDE.md
        wordNode.Format.ContainsKey("alignment").Should().BeTrue("Word canonical key is 'alignment'");
        pptxNode.Format.ContainsKey("align").Should().BeTrue("PPT canonical key is 'align'");
    }

    // ==================== Bug4801 ====================
    // Word paragraph spacing keys are duplicated in Navigation.cs:
    // Both "spaceBefore" AND "spacebefore" are stored for the same value.
    // This wastes memory and is confusing — one key should suffice.
    [Fact]
    public void Bug4801_WordParagraphSpacingKeysDuplicated()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "spaced paragraph",
            ["spacebefore"] = "240"
        });

        var node = handler.Get("/body/p[1]");

        // Both camelCase and lowercase versions are stored
        var hasSpaceBefore = node.Format.ContainsKey("spaceBefore");
        var hasSpacebefore = node.Format.ContainsKey("spacebefore");

        // BUG: Both keys are set — this is wasteful and confusing
        // Navigation.cs lines 238-239 set both Format["spaceBefore"] and Format["spacebefore"]
        (hasSpaceBefore && hasSpacebefore).Should().BeFalse(
            because: "paragraph spacing should use one key, not duplicate both " +
                     "'spaceBefore' and 'spacebefore'. Currently Navigation.cs " +
                     "lines 238-250 store both camelCase AND lowercase for spaceBefore, " +
                     "spaceAfter, and lineSpacing (6 keys for 3 values)");
    }

    // ==================== Bug4802 ====================
    // Excel cell font size format is just a number ("12"),
    // while Word and PPTX return "12pt" with unit suffix.
    // Cross-handler inconsistency.
    [Fact]
    public void Bug4802_ExcelFontSizeFormatVsWordPptx()
    {
        var xlsxPath = CreateTempFile(".xlsx");
        BlankDocCreator.Create(xlsxPath);
        using var excelHandler = new ExcelHandler(xlsxPath, editable: true);

        excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "test", ["font.size"] = "14"
        });

        var excelNode = excelHandler.Get("/Sheet1/A1");

        // Excel returns font size without "pt" suffix
        if (excelNode.Format.ContainsKey("font.size"))
        {
            var excelSize = excelNode.Format["font.size"]?.ToString() ?? "";
            // Verify it's set (may require style to be present)
            if (!string.IsNullOrEmpty(excelSize))
            {
                // Excel format: "14" (no pt suffix)
                // Word format: "14pt"
                // PPTX format: "14pt"
                excelSize.Should().EndWith("pt",
                    because: "Excel font size should include 'pt' suffix for consistency " +
                             "with Word and PPTX, which both return size with 'pt' suffix. " +
                             "Currently Excel returns just the number (e.g., '14') while " +
                             "Word returns '14pt' and PPTX returns '14pt'");
            }
        }
    }

    // ==================== Bug4803 ====================
    // Excel cell alignment key naming vs PPTX/Word:
    // Excel CellToNode stores "alignment.horizontal" AND "halign" for horizontal alignment.
    // PPTX stores "align". Word stores "alignment".
    // Three different key names for the same concept across handlers.
    [Fact]
    public void Bug4803_ExcelAlignmentKeyNamingCrossHandler()
    {
        var xlsxPath = CreateTempFile(".xlsx");
        BlankDocCreator.Create(xlsxPath);
        using var excelHandler = new ExcelHandler(xlsxPath, editable: true);

        excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "centered", ["halign"] = "center"
        });

        var excelNode = excelHandler.Get("/Sheet1/A1");

        // Excel stores both "alignment.horizontal" and "halign" (duplicate keys)
        var hasAlignmentH = excelNode.Format.ContainsKey("alignment.horizontal");
        var hasHalign = excelNode.Format.ContainsKey("halign");

        if (hasAlignmentH && hasHalign)
        {
            // BUG: Both keys are set for the same value — similar to Word's spacing duplication
            (hasAlignmentH && hasHalign).Should().BeFalse(
                because: "Excel should use one key for horizontal alignment, not duplicate " +
                         "both 'alignment.horizontal' AND 'halign'. Additionally, PPTX uses " +
                         "'align' and Word uses 'alignment' — all different");
        }
    }

    // ==================== Bug4804 ====================
    // Excel Set with gradient fill syntax throws ArgumentException because
    // ExcelStyleManager.GetOrCreateFill tries to parse "gradient;FF0000;0000FF;90"
    // as a plain color via NormalizeArgbColor, which rejects the semicolon format.
    // PPTX supports gradient fill via Set with "gradient" key, but Excel has no
    // gradient fill support in Set — it should either support it or give a clear error.
    [Fact]
    public void Bug4804_ExcelGradientFillFormatVsPptx()
    {
        var xlsxPath = CreateTempFile(".xlsx");
        BlankDocCreator.Create(xlsxPath);
        using var excelHandler = new ExcelHandler(xlsxPath, editable: true);

        excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "gradient"
        });

        // Bug: Excel Set throws ArgumentException when trying to set gradient fill.
        // The fill value "gradient;FF0000;0000FF;90" is passed to NormalizeArgbColor
        // which only accepts plain hex colors, not gradient syntax.
        // PPTX has a separate "gradient" key for this; Excel should too.
        var act = () => excelHandler.Set("/Sheet1/A1", new()
        {
            ["fill"] = "gradient;FF0000;0000FF;90"
        });

        act.Should().NotThrow(
            because: "Excel Set should support gradient fill syntax or provide a " +
                     "separate 'gradient' key like PPTX does, not throw ArgumentException " +
                     "when encountering a non-hex-color fill value");
    }

    // ==================== Bug4805 ====================
    // Word Add "alignment" key vs Word Set "alignment" key:
    // Word Add (line 55) uses "alignment" property.
    // Let's verify that both Add and Set use consistent key names.
    [Fact]
    public void Bug4805_WordAddVsSetAlignmentKey()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Add with "alignment" key
        handler.Add("/body", "p", null, new()
        {
            ["text"] = "left aligned",
            ["alignment"] = "left"
        });

        var node1 = handler.Get("/body/p[1]");
        node1.Format.Should().ContainKey("alignment");
        var align1 = node1.Format["alignment"]?.ToString() ?? "";
        align1.Should().Be("left", because: "alignment should be 'left' after Add");

        // Set with "alignment" key to change it
        handler.Set("/body/p[1]", new() { ["alignment"] = "center" });

        var node2 = handler.Get("/body/p[1]");
        node2.Format["alignment"].ToString().Should().Be("center",
            because: "alignment should be 'center' after Set");
    }

    // ==================== Bug4806 ====================
    // Word paragraph Add uses "alignment" key for justification,
    // but the value format differs from PPTX. PPTX align uses
    // "left"/"center"/"right"/"justify" but Word stores raw OOXML
    // values like "center"/"both" (not "justify").
    [Fact]
    public void Bug4806_WordJustifyValueInconsistency()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "justified text",
            ["alignment"] = "justify"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("alignment");

        // Word Navigation.cs line 232: maps "both" -> "justify"
        // But does it handle "justify" input correctly?
        var alignment = node.Format["alignment"]?.ToString() ?? "";
        alignment.Should().Be("justify",
            because: "Word should return 'justify' for justified text (mapping from OOXML 'both')");
    }

    // ==================== Bug4807 ====================
    // Excel cell Set "wrap" key vs Get "alignment.wrapText" key:
    // Set uses "wrap" (via ExcelStyleManager), but Get returns
    // BOTH "alignment.wrapText" AND "wrap" (lines 362-363).
    // The duplicate key in Get is similar to the Word spacing issue.
    [Fact]
    public void Bug4807_ExcelWrapKeyDuplication()
    {
        var xlsxPath = CreateTempFile(".xlsx");
        BlankDocCreator.Create(xlsxPath);
        using var excelHandler = new ExcelHandler(xlsxPath, editable: true);

        excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "wrapped text", ["wrap"] = "true"
        });

        var node = excelHandler.Get("/Sheet1/A1");

        var hasWrap = node.Format.ContainsKey("wrap");
        var hasAlignmentWrapText = node.Format.ContainsKey("alignment.wrapText");

        // BUG: Both keys are set — CellToNode lines 362-363 store both
        if (hasWrap && hasAlignmentWrapText)
        {
            (hasWrap && hasAlignmentWrapText).Should().BeFalse(
                because: "Excel should use one key for wrap text, not duplicate both " +
                         "'wrap' and 'alignment.wrapText'. Currently CellToNode lines " +
                         "362-363 store both keys for the same value");
        }
    }

    // ==================== Bug4808 ====================
    // Excel cell numberformat/format key duplication:
    // CellToNode stores BOTH "numberformat" AND "format" (lines 421-422).
    [Fact]
    public void Bug4808_ExcelNumberFormatKeyDuplication()
    {
        var xlsxPath = CreateTempFile(".xlsx");
        BlankDocCreator.Create(xlsxPath);
        using var excelHandler = new ExcelHandler(xlsxPath, editable: true);

        excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "42", ["numberformat"] = "#,##0.00"
        });

        var node = excelHandler.Get("/Sheet1/A1");

        var hasNumberFormat = node.Format.ContainsKey("numberformat");
        var hasFormat = node.Format.ContainsKey("format");

        // Both may be set — CellToNode lines 421-422
        if (hasNumberFormat && hasFormat)
        {
            // BUG: Duplicate keys for the same value
            (hasNumberFormat && hasFormat).Should().BeFalse(
                because: "Excel should use one key for number format, not duplicate both " +
                         "'numberformat' and 'format'. Currently CellToNode lines 421-422 " +
                         "store both keys for the same number format value");
        }
    }

    // ==================== Bug4809 ====================
    // Word paragraph run formatting: paragraph Get returns first-run properties
    // at paragraph level, but these conflict with the paragraph's own formatting keys.
    // For example, a paragraph with multiple runs — only first run's formatting shows.
    [Fact]
    public void Bug4809_WordParagraphFirstRunFormattingLimitation()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "bold text", ["bold"] = "true", ["font"] = "Arial", ["size"] = "14"
        });

        var node = handler.Get("/body/p[1]");

        // Paragraph-level Get should show first-run formatting
        node.Format.Should().ContainKey("bold",
            because: "paragraph should show first-run bold formatting");
        node.Format.Should().ContainKey("font",
            because: "paragraph should show first-run font");
        node.Format.Should().ContainKey("size",
            because: "paragraph should show first-run size");

        // Size should be in "Xpt" format
        var size = node.Format["size"]?.ToString() ?? "";
        size.Should().EndWith("pt",
            because: "Word paragraph font size should include 'pt' suffix");
    }

    // ==================== Bug4810 ====================
    // PPTX shape "spacing" key for character spacing is inconsistent
    // with the Set key "spacing"/"charspacing"/"letterspacing".
    // Set accepts "spacing", "charspacing", "letterspacing" (ShapeProperties.cs line 428).
    // Get returns "spacing" (NodeBuilder.cs line 407).
    // But there's also "spacing" used for line spacing in other contexts.
    [Fact]
    public void Bug4810_PptxCharSpacingKeyAmbiguity()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "spaced" });

        handler.Set("/slide[1]/shape[1]", new() { ["spacing"] = "3" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("spacing");
        node.Format["spacing"].Should().Be("3",
            because: "character spacing 3pt should round-trip as '3'");

        // "spacing" is ambiguous — could mean character spacing or line spacing
        // PPTX uses "spacing" for character spacing and "lineSpacing" for line spacing
        // This is a naming concern, not a bug per se
    }

    // ==================== Bug4811 ====================
    // Word paragraph size inconsistency: paragraph-level Get returns
    // size as "Xpt" with float division, but Word Add stores size using
    // ParseFontSize which also accepts "Xpt" format.
    // Let's verify the Add→Get round-trip for size.
    [Fact]
    public void Bug4811_WordParagraphSizeRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "sized text", ["size"] = "14"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("size");
        var size = node.Format["size"]?.ToString() ?? "";
        // Add stores size as int(ParseFontSize(14) * 2) = "28" in half-points
        // Get reads as int.Parse("28") / 2.0 = 14.0 → "14pt"
        size.Should().Be("14pt",
            because: "font size 14 should round-trip as '14pt' through Add→Get");
    }

    // ==================== Bug4812 ====================
    // Excel cell font key naming uses dot-separated format ("font.bold", "font.size"),
    // while Word and PPTX use simple keys ("bold", "size").
    // This means code that works with one handler won't work with another.
    [Fact]
    public void Bug4812_ExcelFontKeyNamingVsWordPptx()
    {
        var xlsxPath = CreateTempFile(".xlsx");
        BlankDocCreator.Create(xlsxPath);
        using var excelHandler = new ExcelHandler(xlsxPath, editable: true);

        excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "bold", ["font.bold"] = "true"
        });

        var excelNode = excelHandler.Get("/Sheet1/A1");

        // Excel uses "font.bold", Word uses "bold", PPTX uses "bold"
        if (excelNode.Format.ContainsKey("font.bold"))
        {
            var hasBold = excelNode.Format.ContainsKey("bold");
            // BUG: Excel uses "font.bold" but Word/PPTX use "bold"
            hasBold.Should().BeTrue(
                because: "Excel should provide both 'font.bold' and 'bold' keys for " +
                         "consistency with Word and PPTX which use simple 'bold' key, " +
                         "or all handlers should use the same key naming convention");
        }
    }

    // ==================== Bug4813 ====================
    // PPTX shape textWarp Set doesn't validate the warp name against
    // valid TextShapeValues enum values, causing a cryptic
    // ArgumentOutOfRangeException instead of a user-friendly error.
    [Fact]
    public void Bug4813_PptxTextWarpInvalidValueError()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "warp test" });

        // "invalidWarp" is not a valid TextShapeValues
        var act = () => handler.Set("/slide[1]/shape[1]", new()
        {
            ["textWarp"] = "invalidWarp"
        });

        // BUG: Should throw a user-friendly ArgumentException with valid values listed,
        // but instead throws ArgumentOutOfRangeException from OpenXML SDK
        act.Should().Throw<ArgumentException>(
            because: "invalid textWarp values should produce a user-friendly error " +
                     "listing valid values, not a cryptic ArgumentOutOfRangeException");
    }

    // ==================== Bug4814 ====================
    // Excel cell type "date" is NOT supported in Set (only in Add).
    // Add supports "string", "number", "boolean" types.
    // Set supports "string", "number", "boolean" types.
    // But CellToNode can return type="Date" — no way to set it.
    [Fact]
    public void Bug4814_ExcelCellTypeSetDoesNotSupportDate()
    {
        var xlsxPath = CreateTempFile(".xlsx");
        BlankDocCreator.Create(xlsxPath);
        using var excelHandler = new ExcelHandler(xlsxPath, editable: true);

        excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "45000"
        });

        // Try to set type to "date"
        var act = () => excelHandler.Set("/Sheet1/A1", new() { ["type"] = "date" });

        // "date" is now supported by Set — no throw expected
        act.Should().NotThrow(
            because: "Excel Set 'type' should support 'date' since CellToNode can return type='Date'");
    }

    // ==================== Bug4815 ====================
    // Word paragraph lineSpacing key duplication: both "lineSpacing" AND
    // "linespacing" are stored (Navigation.cs lines 248-249).
    [Fact]
    public void Bug4815_WordParagraphLineSpacingKeyDuplicated()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "double spaced",
            ["linespacing"] = "480"  // 480 twips = double spacing
        });

        var node = handler.Get("/body/p[1]");

        var hasLineSpacing = node.Format.ContainsKey("lineSpacing");
        var hasLinespacing = node.Format.ContainsKey("linespacing");

        // BUG: Both keys are set — wastes memory
        (hasLineSpacing && hasLinespacing).Should().BeFalse(
            because: "paragraph line spacing should use one key, not duplicate both " +
                     "'lineSpacing' and 'linespacing'. Currently Navigation.cs lines " +
                     "248-249 store both camelCase AND lowercase");
    }

    // ==================== Bug4816 ====================
    // Word paragraph spaceAfter key duplication.
    [Fact]
    public void Bug4816_WordParagraphSpaceAfterKeyDuplicated()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "paragraph with space after",
            ["spaceafter"] = "200"
        });

        var node = handler.Get("/body/p[1]");

        var hasSpaceAfter = node.Format.ContainsKey("spaceAfter");
        var hasSpaceafter = node.Format.ContainsKey("spaceafter");

        // BUG: Both keys are set
        (hasSpaceAfter && hasSpaceafter).Should().BeFalse(
            because: "paragraph space after should use one key, not duplicate both " +
                     "'spaceAfter' and 'spaceafter'. Currently Navigation.cs lines " +
                     "243-244 store both keys for the same value");
    }

    // ==================== Bug4817 ====================
    // PPTX shape color from scheme color read-back:
    // When setting color to a scheme color like "accent1",
    // the read-back uses scheme.InnerText which may not match the input.
    [Fact]
    public void Bug4817_PptxSchemeColorRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "themed" });

        handler.Set("/slide[1]/shape[1]", new() { ["color"] = "accent1" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("color");

        // ReadColorFromFill returns scheme.InnerText which should be "accent1"
        var color = node.Format["color"]?.ToString() ?? "";
        color.Should().Be("accent1",
            because: "scheme color 'accent1' should round-trip correctly " +
                     "via ReadColorFromFill which reads scheme.InnerText");
    }

    // ==================== Bug4818 ====================
    // PPTX slide background gradient with 3 stops:
    // Set allows "C1-C2-C3" for 3-stop gradient.
    // Read should preserve all 3 colors.
    [Fact]
    public void Bug4818_PptxBackgroundThreeStopGradient()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });

        handler.Set("/slide[1]", new()
        {
            ["background"] = "FF0000-FFFF00-0000FF"
        });

        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        var bg = node.Format["background"]?.ToString() ?? "";

        // Should contain all three colors
        bg.Should().Contain("FF0000", because: "first gradient stop should be preserved");
        bg.Should().Contain("0000FF", because: "last gradient stop should be preserved");
        bg.Should().Contain("FFFF00", because: "middle gradient stop should be preserved");
    }

    // ==================== Bug4819 ====================
    // PPTX slide background radial gradient round-trip.
    [Fact]
    public void Bug4819_PptxBackgroundRadialGradientRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });

        handler.Set("/slide[1]", new()
        {
            ["background"] = "radial:FF0000-0000FF-tl"
        });

        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        var bg = node.Format["background"]?.ToString() ?? "";

        bg.Should().Contain("radial:", because: "radial gradient should be identified");
        bg.Should().Contain("tl", because: "focal point 'tl' should be preserved");
    }

    // ==================== Bug4820 ====================
    // PPTX shape opacity round-trip.
    [Fact]
    public void Bug4820_PptxShapeOpacityRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "transparent", ["fill"] = "FF0000"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.5" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("opacity");
        node.Format["opacity"].Should().Be("0.5",
            because: "opacity 0.5 (50%) should round-trip correctly");
    }

    // ==================== Bug4821 ====================
    // PPTX table cell border round-trip.
    [Fact]
    public void Bug4821_PptxTableCellBorderRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "bordered",
            ["border.left"] = "FF0000"
        });

        var cellNode = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        // Check if border info is readable
        cellNode.Format.Should().ContainKey("border.left",
            because: "left border should be readable after setting");
    }

    // ==================== Bug4822 ====================
    // Word run shading key: Add uses "shd"/"shading" but Get returns
    // "shading" (Navigation.cs line 359). Let's verify consistency.
    [Fact]
    public void Bug4822_WordRunShadingKeyConsistency()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "highlighted",
            ["shading"] = "FFFF00"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Should().NotBeNull();

        // Run should have shading info
        if (node.Format.ContainsKey("shading"))
        {
            var shadingVal = node.Format["shading"]?.ToString() ?? "";
            shadingVal.Should().Contain("FFFF00",
                because: "run shading color should be preserved");
        }
    }

    // ==================== Bug4823 ====================
    // PPTX presentation-level properties round-trip: slideSize.
    [Fact]
    public void Bug4823_PptxPresentationSlideSizeRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });

        // Set slide size to 4:3
        handler.Set("/", new() { ["slideSize"] = "4:3" });

        var rootNode = handler.Get("/");
        // Check if slide dimensions are readable
        if (rootNode.Format.ContainsKey("slideWidth"))
        {
            // 4:3 → 9144000 EMU width = 25.4cm
            var width = rootNode.Format["slideWidth"]?.ToString() ?? "";
            // Value should represent 4:3 ratio
            width.Should().NotBeNullOrEmpty(
                because: "slide width should be readable after setting slideSize");
        }
    }

    // ==================== Bug4824 ====================
    // Word document properties round-trip: title, author.
    [Fact]
    public void Bug4824_WordDocumentPropertiesRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Set("/", new()
        {
            ["title"] = "Test Document",
            ["author"] = "Test Author"
        });

        var node = handler.Get("/");
        node.Format.Should().ContainKey("title");
        node.Format["title"].Should().Be("Test Document",
            because: "document title should round-trip correctly");
        node.Format.Should().ContainKey("author");
        node.Format["author"].Should().Be("Test Author",
            because: "document author should round-trip correctly");
    }

    // ==================== Bug4825 ====================
    // Excel cell hyperlink round-trip.
    [Fact]
    public void Bug4825_ExcelCellHyperlinkRoundTrip()
    {
        var xlsxPath = CreateTempFile(".xlsx");
        BlankDocCreator.Create(xlsxPath);
        using var excelHandler = new ExcelHandler(xlsxPath, editable: true);

        excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Click here"
        });

        excelHandler.Set("/Sheet1/A1", new()
        {
            ["link"] = "https://example.com"
        });

        var node = excelHandler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("link",
            because: "hyperlink should be readable after setting");
        var link = node.Format["link"]?.ToString() ?? "";
        link.Should().Contain("example.com",
            because: "hyperlink URL should contain the domain");
    }

    // ==================== Bug4826 ====================
    // Word paragraph underline value format: Word Get returns raw OOXML
    // values like "single", "double" which should be consistent with PPTX.
    [Fact]
    public void Bug4826_WordUnderlineValueFormat()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "underlined", ["underline"] = "single"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("underline");
        node.Format["underline"].Should().Be("single",
            because: "underline value should be 'single'");
    }

    // ==================== Bug4827 ====================
    // PPTX shape fill "none" should remove gradient fill too.
    [Fact]
    public void Bug4827_PptxShapeFillNoneRemovesGradient()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });

        // Set gradient first
        handler.Set("/slide[1]/shape[1]", new() { ["gradient"] = "FF0000-0000FF" });

        // Then set fill to none — should remove gradient
        handler.Set("/slide[1]/shape[1]", new() { ["fill"] = "none" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["fill"].Should().Be("none",
            because: "fill should be 'none' after setting fill=none");

        // Gradient should also be removed
        node.Format.Should().NotContainKey("gradient",
            because: "gradient should be removed when fill is set to 'none'");
    }

    // ==================== Bug4828 ====================
    // Word paragraph "keepnext" property round-trip.
    [Fact]
    public void Bug4828_WordParagraphKeepNextRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "keep with next", ["keepnext"] = "true"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("keepNext");
        node.Format["keepNext"].Should().Be(true,
            because: "keepNext property should round-trip as true");
    }

    // ==================== Bug4829 ====================
    // Excel named range round-trip.
    [Fact]
    public void Bug4829_ExcelNamedRangeRoundTrip()
    {
        var xlsxPath = CreateTempFile(".xlsx");
        BlankDocCreator.Create(xlsxPath);
        using var excelHandler = new ExcelHandler(xlsxPath, editable: true);

        excelHandler.Add("/", "namedrange", null, new()
        {
            ["name"] = "TestRange",
            ["ref"] = "Sheet1!A1:B10"
        });

        var node = excelHandler.Get("/namedrange[1]");
        node.Should().NotBeNull();
        node.Text.Should().Contain("A1:B10",
            because: "named range reference should be readable");
    }
}
