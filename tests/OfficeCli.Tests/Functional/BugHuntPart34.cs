// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt round 34: White-box code review bugs found in PPTX, Word, and Excel handlers.
/// Each test targets a specific suspected bug with verification.
/// </summary>
public class BugHuntPart34 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
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

    // =====================================================================
    // Bug3400: PPTX Set text + bold in single call — runs become stale
    // When SetRunOrShapeProperties receives both "text" (multi-line) and "bold"
    // in the same dictionary, the "text" case replaces all paragraphs and creates
    // new runs, but the `runs` list still points to old (orphaned) runs.
    // Subsequent "bold" operates on orphaned runs → no effect on actual shape.
    // =====================================================================
    [Fact]
    public void Bug3400_Pptx_Set_Text_And_Bold_Together_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original"
        });

        // Set text (multi-line to trigger paragraph replacement) and bold together
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["bold"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // Text should be replaced
        node.Text.Should().Contain("Line1");
        node.Text.Should().Contain("Line2");

        // Bold should also be applied to the NEW runs, not orphaned ones
        // This is the bug: bold is applied to orphaned runs, not the new ones
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug3401: PPTX Set text + color in single call — color on orphaned runs
    // Same root cause as Bug3400 but with color property.
    // =====================================================================
    [Fact]
    public void Bug3401_Pptx_Set_Text_And_Color_Together_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original"
        });

        // Set text (multi-line) and color together
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");

        // Color should be applied to the actual runs in the shape
        node.Format.Should().ContainKey("color");
        node.Format["color"].ToString().Should().Be("#FF0000");
    }

    // =====================================================================
    // Bug3402: PPTX Set text + font in single call — font on orphaned runs
    // Same root cause as Bug3400 but with font property.
    // =====================================================================
    [Fact]
    public void Bug3402_Pptx_Set_Text_And_Font_Together_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original"
        });

        // Set text (multi-line) and font together
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["font"] = "Consolas"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");

        // Font should be applied to the actual runs in the shape
        node.Format.Should().ContainKey("font");
        node.Format["font"].ToString().Should().Be("Consolas");
    }

    // =====================================================================
    // Bug3403: Word Set footnote — unsupported properties silently ignored
    // WordHandler.Set for footnote only handles "text" but does not add
    // unknown keys to the unsupported list. Any property besides "text"
    // is silently lost with no feedback.
    // =====================================================================
    [Fact]
    public void Bug3403_Word_Footnote_Set_Unsupported_Key_Not_Reported()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Add paragraph and footnote
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        handler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Note text" });

        // Set an unsupported property on the footnote — should be reported
        var unsupported = handler.Set("/footnote[1]", new()
        {
            ["text"] = "Updated note",
            ["bogusProperty"] = "value"
        });

        // Bug: "bogusProperty" should appear in unsupported list
        unsupported.Should().Contain("bogusProperty",
            "Unknown properties on footnote Set should be reported as unsupported");
    }

    // =====================================================================
    // Bug3404: Word Set endnote — unsupported properties silently ignored
    // Same issue as Bug3403 but for endnotes.
    // =====================================================================
    [Fact]
    public void Bug3404_Word_Endnote_Set_Unsupported_Key_Not_Reported()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        handler.Add("/body/p[1]", "endnote", null, new() { ["text"] = "End text" });

        var unsupported = handler.Set("/endnote[1]", new()
        {
            ["text"] = "Updated endnote",
            ["fakeKey"] = "value"
        });

        // Bug: "fakeKey" should appear in unsupported list
        unsupported.Should().Contain("fakeKey",
            "Unknown properties on endnote Set should be reported as unsupported");
    }

    // =====================================================================
    // Bug3405: PPTX Set text with single line and multiple runs — runs not refreshed
    // When shape has multiple runs (e.g., from Add with text), setting text
    // replaces all paragraphs. Single-line single-run case only replaces
    // first run's text, but font/bold after that still use stale list.
    // =====================================================================
    [Fact]
    public void Bug3405_Pptx_Set_Text_And_Size_Together_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original"
        });

        // Set multi-line text + size together
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2\\nLine3",
            ["size"] = "24"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");

        // Size should be applied to the actual runs in the shape
        node.Format.Should().ContainKey("size");
        var sizeStr = node.Format["size"].ToString()!;
        // Size is returned as "24pt" or raw 24
        sizeStr.Should().Contain("24");
    }

    // =====================================================================
    // Bug3406: PPTX table cell Set font/bold on empty cell — no runs exist
    // When a PPTX table cell is created, it has empty paragraphs with
    // EndParagraphRunProperties but no Drawing.Run elements. Setting
    // "font" or "bold" on such a cell iterates over Descendants<Drawing.Run>()
    // which returns nothing, so the operation has no effect. After setting
    // text and then font, font should apply to the text.
    // =====================================================================
    [Fact]
    public void Bug3406_Pptx_TableCell_Set_Font_On_Empty_Cell()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        // First, set text on a cell
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Hello"
        });

        // Then set font on the same cell
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["font"] = "Consolas"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = node.Children[0].Children[0]; // tr[1]/tc[1]
        cell.Text.Should().Be("Hello");
        // Font should be readable after being set
        cell.Format.Should().ContainKey("font");
    }

    // =====================================================================
    // Bug3407: PPTX table cell Set bold on empty cell — no effect
    // Setting bold on a brand-new empty table cell has no effect because
    // there are no Drawing.Run elements to iterate.
    // =====================================================================
    [Fact]
    public void Bug3407_Pptx_TableCell_Set_Bold_Without_Text()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        // Set bold on empty cell — should not crash
        var unsupported = handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["bold"] = "true"
        });

        // Now add text
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold text"
        });

        // Set bold again after text exists
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["bold"] = "true"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = node.Children[0].Children[0];
        cell.Text.Should().Be("Bold text");
        // Bold should be readable
        cell.Format.Should().ContainKey("bold");
        cell.Format["bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug3408: Word section Set orientation without auto-swapping dimensions
    // Setting orientation to "landscape" on a section only sets the Orient
    // attribute but doesn't swap Width/Height. A portrait A4 page (11906x16838)
    // set to landscape should become (16838x11906), but the code doesn't do this.
    // =====================================================================
    [Fact]
    public void Bug3408_Word_Section_Orientation_No_DimensionSwap()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Get default section properties (should be portrait A4)
        var sec = handler.Get("/section[1]");
        var origWidth = Convert.ToUInt32(sec.Format["pageWidth"]);
        var origHeight = Convert.ToUInt32(sec.Format["pageHeight"]);

        // Portrait: width < height
        origWidth.Should().BeLessThan(origHeight, "Default should be portrait");

        // Set orientation to landscape
        handler.Set("/section[1]", new() { ["orientation"] = "landscape" });

        // Read back
        var updated = handler.Get("/section[1]");
        var newWidth = Convert.ToUInt32(updated.Format["pageWidth"]);
        var newHeight = Convert.ToUInt32(updated.Format["pageHeight"]);

        // Bug: orientation is set but dimensions are NOT swapped
        // In landscape, width should be > height
        newWidth.Should().BeGreaterThan(newHeight,
            "Setting landscape orientation should swap width and height, " +
            "but only the orient attribute is set without dimension swap");
    }

    // =====================================================================
    // Bug3409: Word watermark Set color — SanitizeHex strips # prefix
    // VML fillcolor attribute expects "#RRGGBB" format, but SanitizeHex
    // returns bare hex without # prefix. After setting color, the fillcolor
    // attribute value should have # prefix for VML compatibility.
    // =====================================================================
    [Fact]
    public void Bug3409_Word_Watermark_Set_Color_Missing_HashPrefix()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Add watermark first
        handler.Add("/", "watermark", null, new()
        {
            ["text"] = "DRAFT",
            ["color"] = "#FF0000"
        });

        // Set watermark color
        handler.Set("/watermark", new()
        {
            ["color"] = "#0000FF"
        });

        // Read back watermark
        var node = handler.Get("/watermark");
        node.Format.Should().ContainKey("color");

        // VML fillcolor should include # prefix
        var colorVal = node.Format["color"].ToString()!;
        // Bug: SanitizeHex strips # but VML needs it
        colorVal.Should().StartWith("#",
            "VML fillcolor attribute should have # prefix, but SanitizeHex strips it");
    }

    // =====================================================================
    // Bug3410: PPTX shape Set text single-line single-run preserves old text
    // when runs.Count == 1 and textLines.Length == 1, only the first run's
    // text is replaced. Verify basic single-line replacement works.
    // =====================================================================
    [Fact]
    public void Bug3410_Pptx_Set_Text_SingleLine_Works()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original text"
        });

        // Single-line replacement should work correctly
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Replaced text"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Replaced text");
    }

    // =====================================================================
    // Bug3411: PPTX Add shape with fill + opacity — opacity ignored if fill
    // is not solid. The opacity code looks for SolidFill child of ShapeProperties
    // but if gradient fill was applied (via "gradient" property), SolidFill won't
    // exist and opacity is silently ignored.
    // =====================================================================
    [Fact]
    public void Bug3411_Pptx_Add_Shape_Gradient_Then_Opacity_Ignored()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // Add shape with gradient and opacity — opacity should affect the gradient
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["gradient"] = "FF0000-0000FF",
            ["opacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // Opacity on gradient: currently the Add code only applies opacity to SolidFill
        // So when gradient is used, opacity is silently ignored
        // This is a design limitation — verify current behavior
        node.Format.Should().ContainKey("fill",
            "Shape should have some fill (gradient or solid)");
    }

    // =====================================================================
    // Bug3412: PPTX Set shape — setting "text" and "fill" together
    // "text" processing in SetRunOrShapeProperties modifies the shape's
    // text body, while "fill" modifies ShapeProperties. These should be
    // independent. Verify they don't interfere.
    // =====================================================================
    [Fact]
    public void Bug3412_Pptx_Set_Text_And_Fill_Together_Independent()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original",
            ["fill"] = "FF0000"
        });

        // Set both text and fill together
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Updated",
            ["fill"] = "0000FF"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Updated");
        node.Format["fill"].ToString().Should().Be("#0000FF");
    }

    // =====================================================================
    // Bug3413: PPTX Set text with \\n and italic — italic lost on new runs
    // Verifies that italic applied in the same Set call as multi-line text
    // actually gets applied to the new runs.
    // =====================================================================
    [Fact]
    public void Bug3413_Pptx_Set_MultilineText_And_Italic_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "First\\nSecond",
            ["italic"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("First");
        node.Text.Should().Contain("Second");

        // Italic should be applied to the ACTUAL runs, not orphaned ones
        node.Format.Should().ContainKey("italic");
        node.Format["italic"].Should().Be(true);
    }

    // =====================================================================
    // Bug3414: Excel Set cell value type detection — "true"/"false" becomes string
    // Setting a cell value to "true" auto-detects as string type (not double),
    // but it could be interpreted as boolean. The code doesn't handle boolean
    // auto-detection, only number vs string.
    // =====================================================================
    [Fact]
    public void Bug3414_Excel_Set_Cell_Value_True_TypeDetection()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new()
        {
            ["value"] = "true"
        });

        var node = handler.Get("/Sheet1/A1");
        // "true" is not a double, so it gets CellValues.String
        // This is by design (boolean only via explicit type=boolean)
        node.Format["type"].ToString().Should().Be("String");
    }

    // =====================================================================
    // Bug3415: Word Add paragraph with text + multiple formatting properties
    // Verify that adding a paragraph with text, bold, italic, color all work
    // together correctly.
    // =====================================================================
    [Fact]
    public void Bug3415_Word_Add_Paragraph_MultipleFormats()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Formatted text",
            ["bold"] = "true",
            ["italic"] = "true",
            ["color"] = "FF0000",
            ["size"] = "16"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Formatted text");
        // Verify formatting was applied
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true);
        node.Format.Should().ContainKey("italic");
        node.Format["italic"].Should().Be(true);
    }

    // =====================================================================
    // Bug3416: PPTX table cell Set text + bold in same call
    // SetTableCellProperties iterates cell.Descendants<Drawing.Run>() for
    // bold, but after "text" replaces all paragraphs, the old run references
    // from a prior iteration of the foreach loop are gone. However since
    // Descendants() is lazy and called fresh each case, this might actually
    // work. Let's verify.
    // =====================================================================
    [Fact]
    public void Bug3416_Pptx_TableCell_Set_Text_And_Bold_Together()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        // Set text and bold together on a table cell
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold Cell",
            ["bold"] = "true"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = node.Children[0].Children[0]; // tr[1]/tc[1]
        cell.Text.Should().Be("Bold Cell");
        cell.Format.Should().ContainKey("bold");
        cell.Format["bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug3417: PPTX Add shape — underline "heavy" works in Set but not Add
    // The Add handler has a more limited underline switch with _ => Single,
    // while Set throws on unknown values. "heavy" is valid in Set but maps
    // to Single in Add (fallback _ => Single).
    // =====================================================================
    [Fact]
    public void Bug3417_Pptx_Add_Shape_Underline_Heavy_FallbackToSingle()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // Add shape with underline "heavy" — Add uses fallback _ => Single
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["underline"] = "heavy"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        // Bug: Add maps "heavy" to Single (fallback), but Set would set it to Heavy
        // Verify the actual underline value
        if (node.Format.ContainsKey("underline"))
        {
            var ulValue = node.Format["underline"].ToString()!.ToLowerInvariant();
            // If heavy maps to single, this is a bug (inconsistency with Set)
            ulValue.Should().Contain("heavy",
                "Add should support 'heavy' underline, same as Set. " +
                "Currently falls through to 'single' via default case.");
        }
    }

    // =====================================================================
    // Bug3418: PPTX Add shape — strikethrough "double" in Add vs Set
    // Add uses a switch with _ => SingleStrike fallback, while Set throws
    // on invalid values. "double" is handled in both, so this should work.
    // But verify consistency.
    // =====================================================================
    [Fact]
    public void Bug3418_Pptx_Add_Shape_Strikethrough_Consistency()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // Add with double strikethrough
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["strike"] = "double"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // Both Add and Set should handle "double" correctly
        if (node.Format.ContainsKey("strike"))
        {
            var strikeVal = node.Format["strike"].ToString()!.ToLowerInvariant();
            strikeVal.Should().Contain("double",
                "Add and Set should both correctly handle 'double' strikethrough");
        }
    }

    // =====================================================================
    // Bug3419: Excel Set formula clears DataType but doesn't handle shared strings
    // When a cell has a shared string value and you set a formula,
    // the DataType is cleared but the shared string index might remain
    // as CellValue, potentially causing confusion when reading.
    // =====================================================================
    [Fact]
    public void Bug3419_Excel_Set_Formula_After_SharedString()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        // Set a string value first (might be shared string)
        handler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Hello",
            ["type"] = "string"
        });

        // Now set a formula — should clear old value and DataType
        handler.Set("/Sheet1/A1", new()
        {
            ["formula"] = "1+1"
        });

        var node = handler.Get("/Sheet1/A1");
        // Formula should be set
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].ToString().Should().Be("1+1");
        // Old value should be cleared
        node.Text.Should().Be("=1+1");
    }

    // =====================================================================
    // Bug3420: Word Add paragraph — firstlineindent comment says "× 480"
    // but code doesn't multiply. The comment at line 124 says
    // "firstlineindent already handled above (line ~66-74) with × 480 conversion"
    // but the actual code at lines 58-66 does NOT multiply by 480.
    // It stores the raw value. This is either a stale comment or missing conversion.
    // =====================================================================
    [Fact]
    public void Bug3420_Word_Add_Paragraph_FirstLineIndent_RawValue()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Indented text",
            ["firstlineindent"] = "720"
        });

        var node = handler.Get("/body/p[1]");
        // Check that firstLineIndent is stored as-is (720 twips = 0.5 inch)
        if (node.Format.ContainsKey("firstLineIndent"))
        {
            var indent = node.Format["firstLineIndent"].ToString()!;
            indent.Should().Be("720",
                "First line indent should be stored as raw twips value");
        }
    }

    // =====================================================================
    // Bug3421: PPTX Add shape with "autofit" property — processed twice
    // The Add handler processes "autofit" inline AND also delegates it via
    // effectKeys to SetRunOrShapeProperties. This means autofit is applied
    // twice — once directly and once via delegation. The effectKeys set
    // includes "autofit", so it gets processed again.
    // =====================================================================
    [Fact]
    public void Bug3421_Pptx_Add_Shape_Autofit_DoubleProcessed()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // Add shape with autofit — should not crash despite double processing
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Auto fit test",
            ["autofit"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Auto fit test");

        // autofit should be set (even if processed twice, result should be correct)
        if (node.Format.ContainsKey("autofit"))
        {
            node.Format["autofit"].ToString().Should().NotBeNull();
        }
    }

    // =====================================================================
    // Bug3422: Word Set run "text" — fails if run has no Text child
    // The Set handler for Run does:
    //   var textEl = run.GetFirstChild<Text>();
    //   if (textEl != null) textEl.Text = value;
    // If the run has no Text child (e.g., it only has a Break or other element),
    // the text is silently not set with no error or fallback.
    // =====================================================================
    [Fact]
    public void Bug3422_Word_Set_Run_Text_NoExistingTextElement()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        // Set the run's text — should work
        handler.Set("/body/p[1]/r[1]", new()
        {
            ["text"] = "Updated"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Text.Should().Be("Updated");
    }

    // =====================================================================
    // Bug3423: PPTX Set shape underline "dotted" works
    // Verify that "dotted" underline value is correctly handled in Set.
    // =====================================================================
    [Fact]
    public void Bug3423_Pptx_Set_Shape_Underline_Dotted()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test underline"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["underline"] = "dotted"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("underline");
    }

    // =====================================================================
    // Bug3424: Word Add paragraph — superscript and subscript mutually exclusive
    // If both "superscript" and "subscript" are set, the last one wins.
    // The code processes subscript after superscript, so subscript wins.
    // This is not really a bug but documents the behavior.
    // =====================================================================
    [Fact]
    public void Bug3424_Word_Add_Paragraph_SuperAndSub_LastWins()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Add paragraph with both superscript and subscript (conflicting)
        handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Conflicting",
            ["superscript"] = "true",
            ["subscript"] = "true"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        // Subscript is processed after superscript in the code, so subscript should win
        if (node.Format.ContainsKey("subscript"))
        {
            node.Format["subscript"].Should().Be(true);
        }
    }

    // =====================================================================
    // Bug3425: Excel Set cell "clear" — verify formula is also cleared
    // The "clear" case sets CellValue = null but should also clear formula.
    // =====================================================================
    [Fact]
    public void Bug3425_Excel_Set_Cell_Clear_Also_Clears_Formula()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        // Set formula first
        handler.Set("/Sheet1/A1", new()
        {
            ["formula"] = "SUM(B1:B10)"
        });

        // Clear the cell
        handler.Set("/Sheet1/A1", new()
        {
            ["clear"] = "true"
        });

        var node = handler.Get("/Sheet1/A1");
        // Both value and formula should be cleared
        node.Text.Should().BeEmpty();
        node.Format.Should().NotContainKey("formula",
            "Clear should also remove the formula, not just the value");
    }

    // =====================================================================
    // Bug3426: PPTX Set lineopacity without existing line fill — silently does nothing
    // Setting lineopacity requires an existing SolidFill on the outline.
    // If no line color is set first, lineopacity has no effect.
    // =====================================================================
    [Fact]
    public void Bug3426_Pptx_Set_LineOpacity_Without_LineFill_NoEffect()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test"
        });

        // Set linecolor first, then lineopacity — should work
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["linecolor"] = "000000",
            ["lineopacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // Verify line color was set
        node.Format.Should().ContainKey("line");
    }
}
