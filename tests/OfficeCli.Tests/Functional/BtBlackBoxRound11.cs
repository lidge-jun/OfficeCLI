// Black-box tests (Round 11) — key objectives:
//   1. R9 fixes: font size 10.25pt/11.25pt round-trip, watermark multi-section, Set dict immutability
//   2. Regression of all prior fixes
//   3. New territory: PPTX notes, slide transitions, Word endnotes, Excel merged cells, PPTX alignment

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound11 : IDisposable
{
    private readonly List<string> _temps = new();
    private readonly ITestOutputHelper _out;

    public BtBlackBoxRound11(ITestOutputHelper output) => _out = output;

    public void Dispose()
    {
        foreach (var f in _temps)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt11_{Guid.NewGuid():N}.{ext}");
        _temps.Add(p);
        BlankDocCreator.Create(p);
        return p;
    }

    private string TempImg()
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt11_img_{Guid.NewGuid():N}.png");
        _temps.Add(p);
        File.WriteAllBytes(p, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg=="));
        return p;
    }

    private void ValidateDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[DOCX {step}] {e.Description}");
        errors.Should().BeEmpty($"DOCX invalid after: {step}");
    }

    private void ValidatePptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[PPTX {step}] {e.Description}");
        errors.Should().BeEmpty($"PPTX invalid after: {step}");
    }

    private void ValidateXlsx(string path, string step)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[XLSX {step}] {e.Description}");
        errors.Should().BeEmpty($"XLSX invalid after: {step}");
    }

    // ==================== 1. R9 fix: font size round-trip with MidpointRounding.AwayFromZero ====================

    [Fact]
    public void Word_Add_FontSize_10pt25_RoundsTo_10pt5_AwayFromZero()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "10.25 test", ["size"] = "10.25pt" });
        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        var size = node!.Format.GetValueOrDefault("size")?.ToString();
        _out.WriteLine($"10.25pt Add → {size}");
        // 10.25 * 2 = 20.5 → AwayFromZero → 21 → 10.5pt
        size.Should().Be("10.5pt", "10.25pt rounds to 10.5pt with AwayFromZero");
    }

    [Fact]
    public void Word_Set_FontSize_11pt25_RoundsTo_11pt5()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "11.25 test", ["size"] = "12pt" });
        h.Set("/body/p[1]", new() { ["size"] = "11.25pt" });
        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        var size = node!.Format.GetValueOrDefault("size")?.ToString();
        _out.WriteLine($"11.25pt Set → {size}");
        // 11.25 * 2 = 22.5 → AwayFromZero → 23 → 11.5pt
        size.Should().Be("11.5pt", "11.25pt rounds to 11.5pt with AwayFromZero");
    }

    [Fact]
    public void Word_FontSize_Persistence_RoundTrip()
    {
        var path = Temp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Persist size", ["size"] = "10.5pt" });
        }
        using (var h2 = new WordHandler(path, editable: false))
        {
            var node = h2.Get("/body/p[1]");
            node.Should().NotBeNull();
            node!.Format["size"].ToString().Should().Be("10.5pt", "10.5pt persists after reopen");
        }
    }

    // ==================== 2. R9 fix: Set does not mutate caller's dict ====================

    [Fact]
    public void Word_Set_DoesNotMutateCallerDict()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Mutation test" });

        var original = new Dictionary<string, string>
        {
            ["bold"] = "true",
            ["italic"] = "true",
            ["size"] = "14pt",
            ["text"] = "find:Mutation=>replace:Updated"
        };
        var snapshot = original.ToDictionary(kv => kv.Key, kv => kv.Value);

        h.Set("/body/p[1]", original);

        // Caller's dict must be unchanged
        original.Should().BeEquivalentTo(snapshot,
            "Set must not mutate the caller's dictionary");
    }

    [Fact]
    public void Pptx_Set_DoesNotMutateCallerDict()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Mutation test" });
        h.Add("/slide[1]", "textbox", null, new()
        {
            ["text"] = "Test", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "5cm", ["height"] = "2cm"
        });

        var original = new Dictionary<string, string>
        {
            ["bold"] = "true",
            ["size"] = "16pt",
            ["fill"] = "#FF0000"
        };
        var snapshot = original.ToDictionary(kv => kv.Key, kv => kv.Value);

        h.Set("/slide[1]/shape[2]", original);

        original.Should().BeEquivalentTo(snapshot,
            "PPTX Set must not mutate the caller's dictionary");
    }

    // ==================== 3. R9 fix: watermark on multi-section document ====================

    [Fact]
    public void Word_Watermark_MultiSection_AppliedToAll()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);

        // Add two paragraphs (section break between them)
        h.Add("/body", "paragraph", null, new() { ["text"] = "Section 1 content" });
        h.Add("/body", "paragraph", null, new() { ["text"] = "Section 2 content" });
        // Add a section break to create multi-section doc
        h.Add("/body", "section", null, new() { ["pageWidth"] = "12240", ["pageHeight"] = "15840" });

        // Add watermark
        var ex = Record.Exception(() => h.Add("/body", "watermark", null, new() { ["text"] = "CONFIDENTIAL" }));
        ex.Should().BeNull("watermark Add should not throw on multi-section doc");

        h.Dispose();
        ValidateDocx(path, "watermark multi-section");

        // Reopen: watermark should be present (via Get "/watermark")
        using var h2 = new WordHandler(path, editable: false);
        var wmNode = h2.Get("/watermark");
        _out.WriteLine($"Watermark node text: '{wmNode?.Text}'");
        wmNode.Should().NotBeNull("watermark node returned");
        wmNode!.Text.Should().NotBe("(no watermark)", "watermark text should be CONFIDENTIAL, not placeholder");
    }

    [Fact]
    public void Word_Watermark_AddTwice_OnlyOneWatermark()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);

        h.Add("/body", "watermark", null, new() { ["text"] = "DRAFT" });
        // Adding again should replace, not duplicate
        h.Add("/body", "watermark", null, new() { ["text"] = "FINAL" });

        h.Dispose();
        ValidateDocx(path, "watermark replace");

        using var h2 = new WordHandler(path, editable: false);
        var wmNode2 = h2.Get("/watermark");
        _out.WriteLine($"Watermark text after double-add: '{wmNode2?.Text}'");
        wmNode2.Should().NotBeNull("watermark present after double Add");
        wmNode2!.Text.Should().NotBe("(no watermark)", "watermark text is set");
        // The text should be "FINAL" (the second watermark replaced the first)
        wmNode2.Text.Should().Be("FINAL", "second Add replaces the first watermark");
    }

    // ==================== 4. New: PPTX slide notes add/get/set ====================

    [Fact]
    public void Pptx_SlideNotes_AddAndGet()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Notes slide" });

        // Set notes via Set
        h.Set("/slide[1]", new() { ["notes"] = "Speaker notes for slide 1" });

        var slide = h.Get("/slide[1]");
        slide.Should().NotBeNull();
        _out.WriteLine($"Slide format keys: {string.Join(", ", slide!.Format.Keys)}");

        // Notes should appear in Format or be accessible via /slide[1]/notes
        var notesNode = h.Get("/slide[1]/notes");
        _out.WriteLine($"Notes node: {notesNode?.Text}");
        // Either the slide Format contains notes, or a notes node exists
        var hasNotes = (slide.Format.TryGetValue("notes", out var nv) && !string.IsNullOrEmpty(nv?.ToString()))
            || (notesNode != null && !string.IsNullOrEmpty(notesNode.Text));
        hasNotes.Should().BeTrue("speaker notes should be retrievable");

        h.Dispose();
        ValidatePptx(path, "slide notes");
    }

    [Fact]
    public void Pptx_SlideNotes_Persistence()
    {
        var path = Temp("pptx");
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Persist notes" });
            h.Set("/slide[1]", new() { ["notes"] = "Persisted notes text" });
        }

        ValidatePptx(path, "notes persist");

        using (var h2 = new PowerPointHandler(path, editable: false))
        {
            var notesNode = h2.Get("/slide[1]/notes");
            _out.WriteLine($"Persisted notes text: {notesNode?.Text}");
            notesNode.Should().NotBeNull("notes node persists after reopen");
        }
    }

    // ==================== 5. New: PPTX slide transition properties ====================

    [Fact]
    public void Pptx_Transition_SetAndGet()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Transition slide" });

        h.Set("/slide[1]", new() { ["transition"] = "fade", ["transitionDuration"] = "1000" });

        var node = h.Get("/slide[1]");
        node.Should().NotBeNull();
        _out.WriteLine($"Slide format keys: {string.Join(", ", node!.Format.Keys)}");

        h.Dispose();
        ValidatePptx(path, "slide transition");
    }

    // ==================== 6. New: Word endnotes add/get ====================

    [Fact]
    public void Word_Endnote_AddGet_Works()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Endnote paragraph" });

        var ex = Record.Exception(() =>
            h.Add("/body/p[1]", "endnote", null, new() { ["text"] = "Endnote content" }));
        ex.Should().BeNull("endnote Add should not throw");

        // Access endnote via Get path
        var enNode = h.Get("/endnote[1]");
        _out.WriteLine($"Endnote[1] node: type='{enNode?.Type}', text='{enNode?.Text}'");
        enNode.Should().NotBeNull("endnote[1] accessible via Get path");
        enNode!.Text.Should().Contain("Endnote content", "endnote text is set");

        h.Dispose();
        ValidateDocx(path, "Word endnote");
    }

    // ==================== 7. New: Excel merged cells ====================

    [Fact]
    public void Excel_MergeCells_SetAndGet()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        h.Set("/Sheet1/B2", new() { ["value"] = "Merged", ["merge"] = "B2:D4" });

        var node = h.Get("/Sheet1/B2");
        node.Should().NotBeNull();
        var mergedKeys = node?.Format.Keys.ToList() ?? new List<string>();
        _out.WriteLine($"Merged cell B2: text='{node?.Text}', format keys={string.Join(", ", mergedKeys)}");
        node!.Text.Should().Be("Merged", "merged cell retains value");

        h.Dispose();
        ValidateXlsx(path, "Excel merged cells");
    }

    // ==================== 8. New: PPTX text alignment Set+Get ====================

    [Fact]
    public void Pptx_TextAlign_SetAndGet_RoundTrip()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Alignment test" });
        h.Add("/slide[1]", "textbox", null, new()
        {
            ["text"] = "Centered text",
            ["x"] = "2cm", ["y"] = "3cm", ["width"] = "10cm", ["height"] = "3cm",
            ["align"] = "center"
        });

        var node = h.Get("/slide[1]/shape[2]");
        node.Should().NotBeNull();
        _out.WriteLine($"Align format: {node!.Format.GetValueOrDefault("align")}");
        node.Format.GetValueOrDefault("align")?.ToString().Should().Be("center",
            "center alignment persists after Add");

        // Change to right
        h.Set("/slide[1]/shape[2]", new() { ["align"] = "right" });
        var node2 = h.Get("/slide[1]/shape[2]");
        node2!.Format.GetValueOrDefault("align")?.ToString().Should().Be("right",
            "alignment updated to right via Set");

        h.Dispose();
        ValidatePptx(path, "text alignment round-trip");
    }

    // ==================== 9. Regression: color #-prefix canonical output ====================

    [Fact]
    public void Pptx_Shape_Fill_GetReturnsHashPrefixColor()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Color test" });
        h.Add("/slide[1]", "textbox", null, new()
        {
            ["text"] = "Colored",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "2cm",
            ["fill"] = "FF0000"  // no hash prefix on input
        });

        var node = h.Get("/slide[1]/shape[2]");
        node.Should().NotBeNull();
        var fill = node!.Format.GetValueOrDefault("fill")?.ToString();
        _out.WriteLine($"Fill returned: {fill}");
        fill.Should().Be("#FF0000", "Get should return #-prefixed color");
    }

    [Fact]
    public void Word_Run_Color_GetReturnsHashPrefixColor()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Colored text",
            ["color"] = "4472C4"  // no hash on input
        });

        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        var color = node!.Format.GetValueOrDefault("color")?.ToString();
        _out.WriteLine($"Color returned: {color}");
        color.Should().Be("#4472C4", "Word Get should return #-prefixed color");
    }

    // ==================== 10. Regression: Excel spacing canonical keys ====================

    [Fact]
    public void Excel_CellAlignment_CanonicalKeys_Returned()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new()
        {
            ["value"] = "test",
            ["alignment.horizontal"] = "center",
            ["alignment.vertical"] = "top",
            ["alignment.wrapText"] = "true"
        });

        var node = h.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        _out.WriteLine($"Excel cell format keys: {string.Join(", ", node!.Format.Keys)}");

        // Canonical keys must be present
        node.Format.Should().ContainKey("alignment.horizontal", "canonical key used");
        node.Format.Should().ContainKey("alignment.vertical", "canonical key used");
        node.Format["alignment.horizontal"].ToString().Should().Be("center");
        node.Format["alignment.vertical"].ToString().Should().Be("top");

        // Legacy aliases must NOT be present as separate keys
        var allKeys = node.Format.Keys.ToList();
        allKeys.Should().NotContain("halign", "no legacy alias");
        allKeys.Should().NotContain("wrap", "no legacy alias");
    }

    // ==================== 11. New: Word hyperlink add/get ====================

    [Fact]
    public void Word_Hyperlink_AddGet_Works()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Visit " });
        h.Add("/body/p[1]", "hyperlink", null, new()
        {
            ["text"] = "Click here",
            ["url"] = "https://example.com"
        });

        var para = h.Get("/body/p[1]");
        para.Should().NotBeNull();
        _out.WriteLine($"Para text: '{para!.Text}'");
        para.Text.Should().Contain("Click here", "hyperlink text is in paragraph");

        h.Dispose();
        ValidateDocx(path, "Word hyperlink");
    }

    // ==================== 12. New: Excel number format canonical key ====================

    [Fact]
    public void Excel_NumberFormat_CanonicalKey_Returned()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/C1", new()
        {
            ["value"] = "1234.56",
            ["numberformat"] = "#,##0.00"
        });

        var node = h.Get("/Sheet1/C1");
        node.Should().NotBeNull();
        _out.WriteLine($"Cell format keys: {string.Join(", ", node!.Format.Keys)}");
        _out.WriteLine($"numberformat = {node.Format.GetValueOrDefault("numberformat")}");

        node.Format.Should().ContainKey("numberformat", "canonical key 'numberformat' returned");
        var keyList = node.Format.Keys.ToList();
        keyList.Should().NotContain("format", "legacy alias 'format' not returned");
    }

    // ==================== 13. New: PPTX shape border (line) Set/Get ====================

    [Fact]
    public void Pptx_ShapeBorder_SetAndGet()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Border test" });
        h.Add("/slide[1]", "textbox", null, new()
        {
            ["text"] = "Bordered",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "8cm", ["height"] = "3cm",
            ["line"] = "#000000",
            ["lineWidth"] = "2pt"
        });

        var node = h.Get("/slide[1]/shape[2]");
        node.Should().NotBeNull();
        _out.WriteLine($"Shape format keys: {string.Join(", ", node!.Format.Keys)}");

        // line color should be returned with # prefix
        if (node.Format.TryGetValue("line", out var lineColor))
        {
            _out.WriteLine($"line = {lineColor}");
            lineColor?.ToString().Should().StartWith("#", "line color has # prefix");
        }

        h.Dispose();
        ValidatePptx(path, "shape border");
    }

    // ==================== 14. Regression: PPTX shadow 'drop-shadow' fix ====================

    [Fact]
    public void Pptx_Shadow_SetAndGet_DoesNotCrash()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Shadow test" });
        h.Add("/slide[1]", "textbox", null, new()
        {
            ["text"] = "Shadowed",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "8cm", ["height"] = "3cm"
        });

        var ex = Record.Exception(() =>
            h.Set("/slide[1]/shape[2]", new() { ["shadow"] = "true" }));
        ex.Should().BeNull("shadow Set should not throw");

        h.Dispose();
        ValidatePptx(path, "shape shadow");
    }

    // ==================== 15. Regression: Word spacing unit-qualified output ====================

    [Fact]
    public void Word_Spacing_GetReturns_UnitQualified()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Spaced",
            ["spaceBefore"] = "12pt",
            ["spaceAfter"] = "6pt",
            ["lineSpacing"] = "1.5x"
        });

        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        _out.WriteLine($"spaceBefore={node!.Format.GetValueOrDefault("spaceBefore")}, " +
                       $"spaceAfter={node.Format.GetValueOrDefault("spaceAfter")}, " +
                       $"lineSpacing={node.Format.GetValueOrDefault("lineSpacing")}");

        node.Format["spaceBefore"].ToString().Should().Be("12pt", "spaceBefore is unit-qualified");
        node.Format["spaceAfter"].ToString().Should().Be("6pt", "spaceAfter is unit-qualified");
        var ls = node.Format.GetValueOrDefault("lineSpacing")?.ToString();
        (ls == "1.5x" || ls == "1.5").Should().BeTrue("lineSpacing is multiplier format");
    }
}
