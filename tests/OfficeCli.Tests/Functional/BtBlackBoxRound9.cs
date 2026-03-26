// Black-box tests (Round 9) — key objectives:
//   1. R7 regression: header with picture removed → ImagePart is actually cleaned up
//   2. Comprehensive regression: combine multiple previously fixed scenarios in one doc
//   3. PPTX CopyFrom (CloneSlide): content + image parts copied, original intact
//   4. Extreme sequence: Add 30 shapes → Remove all → file valid + slide still exists
//   5. Get→Set round-trip: pass Format dict directly back to Set, verify no data loss

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound9 : IDisposable
{
    private readonly List<string> _temps = new();
    private readonly ITestOutputHelper _out;

    public BtBlackBoxRound9(ITestOutputHelper output) => _out = output;

    public void Dispose()
    {
        foreach (var f in _temps)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt9_{Guid.NewGuid():N}.{ext}");
        _temps.Add(p);
        BlankDocCreator.Create(p);
        return p;
    }

    private string TempImg()
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt9_img_{Guid.NewGuid():N}.png");
        _temps.Add(p);
        // Minimal 1x1 PNG
        File.WriteAllBytes(p, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg=="));
        return p;
    }

    private void ValidateDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"DOCX invalid after: {step}");
    }

    private void ValidatePptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"PPTX invalid after: {step}");
    }

    // ==================== 1. R7: ImagePart cleaned up after header removal ====================

    [Fact]
    public void Word_HeaderWithPicture_Remove_ImagePartCleaned()
    {
        var imgPath = TempImg();
        var docPath = Temp("docx");

        long sizeWithImage;
        using (var h = new WordHandler(docPath, editable: true))
        {
            h.Add("/body", "header", null, new() { ["text"] = "Header text" });
            h.Add("/header[1]", "picture", null, new()
            {
                ["src"] = imgPath,
                ["width"] = "3cm",
                ["height"] = "2cm"
            });

            // Verify image is embedded in header
            var headers = h.Query("header");
            headers.Should().NotBeEmpty("header must exist before remove");
        }
        sizeWithImage = new FileInfo(docPath).Length;
        _out.WriteLine($"Size with header+image: {sizeWithImage} bytes");

        // Verify ImagePart exists before removal (image is stored on MainDocumentPart)
        int imageCountBefore;
        using (var doc = WordprocessingDocument.Open(docPath, false))
        {
            imageCountBefore = doc.MainDocumentPart!.ImageParts.Count();
        }
        _out.WriteLine($"ImagePart count on mainPart before remove: {imageCountBefore}");
        imageCountBefore.Should().BeGreaterThan(0, "main part should contain image part for header image");

        // Remove header
        using (var h2 = new WordHandler(docPath, editable: true))
        {
            h2.Remove("/header[1]");
        }

        ValidateDocx(docPath, "after header+image removal");

        // Verify ImagePart is cleaned up from mainPart
        int imageCountAfter;
        using (var doc = WordprocessingDocument.Open(docPath, false))
        {
            imageCountAfter = doc.MainDocumentPart!.ImageParts.Count();
            _out.WriteLine($"ImagePart count on mainPart after remove: {imageCountAfter}");
        }

        imageCountAfter.Should().Be(0, "ImagePart should be deleted when its only-referencing header is removed");

        // File should be smaller (no orphan blob)
        var sizeAfter = new FileInfo(docPath).Length;
        _out.WriteLine($"Size after header removal: {sizeAfter} bytes");
        sizeAfter.Should().BeLessThan(sizeWithImage, "file should shrink after image-containing header is removed");
    }

    // ==================== 2. Comprehensive regression: mixed scenario in one doc ====================

    [Fact]
    public void Word_CombinedScenario_MultipleOperations_FileValid()
    {
        var imgPath = TempImg();
        var docPath = Temp("docx");

        using var h = new WordHandler(docPath, editable: true);

        // Add paragraphs with various properties
        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Bold paragraph",
            ["bold"] = "true",
            ["alignment"] = "center",
            ["spaceBefore"] = "12pt"
        });
        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Italic paragraph",
            ["italic"] = "true",
            ["spaceAfter"] = "6pt"
        });

        // Add a table
        h.Add("/body", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "3"
        });

        // Add header + footer with text
        h.Add("/body", "header", null, new() { ["text"] = "Header text" });
        h.Add("/body", "footer", null, new() { ["text"] = "Footer text" });

        // Add image in body
        h.Add("/body", "paragraph", null, new() { ["text"] = "Image paragraph" });
        h.Add("/body/p[3]", "picture", null, new()
        {
            ["src"] = imgPath,
            ["width"] = "2cm",
            ["height"] = "1cm"
        });

        // Add footnote
        h.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Footnote for bold para" });

        // Verify various elements
        h.Get("/body/p[1]")!.Format["alignment"].ToString().Should().Be("center");
        h.Get("/body/p[2]").Should().NotBeNull();
        h.Get("/header[1]").Should().NotBeNull();
        h.Get("/footer[1]").Should().NotBeNull();
        h.Query("picture").Should().NotBeEmpty();

        // Remove one paragraph
        h.Remove("/body/p[2]");

        // Remove header (should also clean image if any)
        h.Remove("/header[1]");

        // Remove footnote
        h.Remove("/footnote[1]");

        h.Dispose();
        ValidateDocx(docPath, "combined scenario");

        // Reopen + spot-check
        using var h2 = new WordHandler(docPath, editable: false);
        h2.Get("/body").Should().NotBeNull();
        h2.Query("picture").Should().NotBeEmpty("body image survives");
        h2.Query("footer").Should().NotBeEmpty("footer survives");
    }

    // ==================== 3. PPTX CloneSlide (CopyFrom) ====================

    [Fact]
    public void Pptx_CloneSlide_ContentCopied_OriginalIntact()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        // Add source slide with content
        h.Add("/", "slide", null, new() { ["title"] = "Original Slide" });
        h.Add("/slide[1]", "textbox", null, new()
        {
            ["text"] = "Clone me",
            ["x"] = "2cm",
            ["y"] = "4cm",
            ["width"] = "6cm",
            ["height"] = "2cm"
        });

        // Clone slide[1] → new slide
        var clonedPath = h.CopyFrom("/slide[1]", "/", null);
        _out.WriteLine($"Cloned to: {clonedPath}");

        clonedPath.Should().NotBeNullOrEmpty("clone should return a path");

        // Original slide still intact
        var orig = h.Get("/slide[1]");
        orig.Should().NotBeNull();

        // Cloned slide exists and has text
        var clone = h.Get(clonedPath);
        clone.Should().NotBeNull($"clone path {clonedPath} should be accessible");

        h.Dispose();
        ValidatePptx(path, "after CloneSlide");

        // Reopen: both slides present (use PresentationDocument to count)
        using var doc = PresentationDocument.Open(path, false);
        var slideCount = doc.PresentationPart!.SlideParts.Count();
        _out.WriteLine($"Slide count after reopen: {slideCount}");
        slideCount.Should().BeGreaterOrEqualTo(2, "cloned slide persists after reopen");
    }

    [Fact]
    public void Pptx_CloneSlide_WithImage_ImagePartCopied()
    {
        var imgPath = TempImg();
        var path = Temp("pptx");

        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Slide with Image" });
        h.Add("/slide[1]", "picture", null, new()
        {
            ["src"] = imgPath,
            ["x"] = "1cm",
            ["y"] = "1cm",
            ["width"] = "4cm",
            ["height"] = "3cm"
        });

        // Clone the slide that contains the image
        var clonedPath = h.CopyFrom("/slide[1]", "/", null);

        h.Dispose();
        ValidatePptx(path, "after CloneSlide with image");

        // Both slides should have image parts
        using var doc = PresentationDocument.Open(path, false);
        var slideParts = doc.PresentationPart!.SlideParts.ToList();
        slideParts.Should().HaveCountGreaterOrEqualTo(2);

        foreach (var sp in slideParts)
        {
            var imgParts = sp.ImageParts.ToList();
            _out.WriteLine($"Slide has {imgParts.Count} image parts");
            imgParts.Should().NotBeEmpty("each cloned slide should have its own image part");
        }
    }

    // ==================== 4. Add 30 shapes → Remove all → file valid ====================

    [Fact]
    public void Pptx_Add30Shapes_RemoveAll_FileValid()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Stress Test" });

        // Add 30 textboxes
        for (int i = 1; i <= 30; i++)
        {
            h.Add("/slide[1]", "textbox", null, new()
            {
                ["text"] = $"Shape {i}",
                ["x"] = $"{(i % 5) * 2}cm",
                ["y"] = $"{(i / 5) * 2}cm",
                ["width"] = "1.8cm",
                ["height"] = "1cm"
            });
        }

        var shapes = h.Query("textbox");
        _out.WriteLine($"Shapes after add: {shapes.Count}");
        // Query("textbox") includes the title placeholder, so we have 31 total (1 title + 30 textboxes)
        shapes.Should().HaveCountGreaterOrEqualTo(30, "all 30 textboxes + title should exist");

        // Remove all non-title textboxes one by one — shape[1] is title, textboxes start at shape[2]
        // Always remove shape[2] (it shifts after each removal)
        for (int i = 0; i < 30; i++)
        {
            h.Remove("/slide[1]/shape[2]");
        }

        var remaining = h.Query("textbox");
        _out.WriteLine($"Shapes after remove-all: {remaining.Count}");
        // Only the title placeholder shape[1] should remain
        remaining.Should().HaveCount(1, "only the title shape should remain after removing 30 textboxes");
        remaining[0].Type.Should().Be("title", "remaining shape is the slide title");

        h.Dispose();
        ValidatePptx(path, "after add30+remove30");

        // Slide itself should still exist
        using var h2 = new PowerPointHandler(path, editable: false);
        h2.Get("/slide[1]").Should().NotBeNull("slide survives after all shapes removed");
    }

    // ==================== 5. Get→Set round-trip: pass Format dict back ====================

    [Fact]
    public void Pptx_Shape_GetSetRoundTrip_FormatPreserved()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Round-trip" });
        h.Add("/slide[1]", "textbox", null, new()
        {
            ["text"] = "Test shape",
            ["x"] = "2cm",
            ["y"] = "3cm",
            ["width"] = "5cm",
            ["height"] = "2cm",
            ["fill"] = "#4472C4",
            ["bold"] = "true",
            ["size"] = "16pt"
        });

        // Get the shape
        var node = h.Get("/slide[1]/shape[2]");
        node.Should().NotBeNull();
        _out.WriteLine($"Format keys before Set: {string.Join(", ", node!.Format.Keys)}");

        // Extract only settable keys (exclude read-only structural keys)
        var settableKeys = new[] { "x", "y", "width", "height", "fill", "bold", "size" };
        var roundTripDict = new Dictionary<string, string>();
        foreach (var key in settableKeys)
        {
            if (node.Format.TryGetValue(key, out var val) && val != null)
                roundTripDict[key] = val.ToString()!;
        }
        _out.WriteLine($"Round-trip dict: {string.Join(", ", roundTripDict.Select(kv => $"{kv.Key}={kv.Value}"))}");

        // Pass Format values back to Set
        h.Set("/slide[1]/shape[2]", roundTripDict);

        // Get again — values should still match
        var node2 = h.Get("/slide[1]/shape[2]");
        node2.Should().NotBeNull();
        node2!.Text.Should().Be("Test shape");

        if (node2.Format.TryGetValue("fill", out var fill2))
            fill2?.ToString().Should().Be("#4472C4", "fill should survive round-trip");
        if (node2.Format.TryGetValue("bold", out var bold2))
            bold2?.ToString()?.ToLower().Should().BeOneOf("true", "1", "yes", "bold");

        h.Dispose();
        ValidatePptx(path, "after Get→Set round-trip");
    }

    [Fact]
    public void Word_Paragraph_GetSetRoundTrip_FormatPreserved()
    {
        var docPath = Temp("docx");
        using var h = new WordHandler(docPath, editable: true);

        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Round-trip paragraph",
            ["alignment"] = "center",
            ["spaceBefore"] = "12pt",
            ["spaceAfter"] = "6pt",
            ["bold"] = "true"
        });

        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        _out.WriteLine($"Word Format keys: {string.Join(", ", node!.Format.Keys)}");

        // Build round-trip dict from canonical keys
        var settable = new[] { "alignment", "spaceBefore", "spaceAfter" };
        var rtDict = new Dictionary<string, string>();
        foreach (var k in settable)
        {
            if (node.Format.TryGetValue(k, out var v) && v != null)
                rtDict[k] = v.ToString()!;
        }
        _out.WriteLine($"Round-trip dict: {string.Join(", ", rtDict.Select(kv => $"{kv.Key}={kv.Value}"))}");

        h.Set("/body/p[1]", rtDict);

        var node2 = h.Get("/body/p[1]");
        node2.Should().NotBeNull();
        node2!.Text.Should().Be("Round-trip paragraph");

        if (node2.Format.TryGetValue("alignment", out var align2))
            align2?.ToString().Should().Be("center", "alignment survives round-trip");
        if (node2.Format.TryGetValue("spaceBefore", out var sb2))
            sb2?.ToString().Should().Be("12pt", "spaceBefore survives round-trip");

        h.Dispose();
        ValidateDocx(docPath, "after Word paragraph round-trip");
    }

    [Fact]
    public void Excel_Cell_GetSetRoundTrip_FormatPreserved()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        // Set cell with formatting
        h.Set("/Sheet1/B2", new()
        {
            ["value"] = "42",
            ["bold"] = "true",
            ["fontsize"] = "14pt",
            ["fill"] = "#FFFF00",
            ["alignment.horizontal"] = "center"
        });

        var node = h.Get("/Sheet1/B2");
        node.Should().NotBeNull();
        _out.WriteLine($"Excel cell Format keys: {string.Join(", ", node!.Format.Keys)}");

        // Pass Format back — should not throw
        var rtDict = new Dictionary<string, string>();
        foreach (var kv in node.Format)
        {
            if (kv.Value != null)
                rtDict[kv.Key] = kv.Value.ToString()!;
        }

        // Set should not throw
        var ex = Record.Exception(() => h.Set("/Sheet1/B2", rtDict));
        ex.Should().BeNull("Set with Get-returned Format should not throw");

        // Value should survive
        var node2 = h.Get("/Sheet1/B2");
        node2.Should().NotBeNull();
        node2!.Text.Should().Be("42", "cell value survives round-trip");

        h.Dispose();

        using var xlDoc = SpreadsheetDocument.Open(path, false);
        var errs = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(xlDoc).ToList();
        foreach (var e in errs) _out.WriteLine($"[xlsx round-trip] {e.Description}");
        errs.Should().BeEmpty("Excel file valid after round-trip Set");
    }

    // ==================== 6. PPTX Swap slides ====================

    [Fact]
    public void Pptx_SwapSlides_OrderChanges_Persists()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Slide A" });
        h.Add("/", "slide", null, new() { ["title"] = "Slide B" });

        var before1 = h.Get("/slide[1]");
        var before2 = h.Get("/slide[2]");
        _out.WriteLine($"Before swap: slide[1]='{before1?.Text}', slide[2]='{before2?.Text}'");

        h.Swap("/slide[1]", "/slide[2]");

        var after1 = h.Get("/slide[1]");
        var after2 = h.Get("/slide[2]");
        _out.WriteLine($"After swap: slide[1]='{after1?.Text}', slide[2]='{after2?.Text}'");

        // After swap, titles should have swapped positions
        if (before1?.Text != null && after2?.Text != null)
            after2.Text.Should().Contain("A", "original slide A moved to position 2");
        if (before2?.Text != null && after1?.Text != null)
            after1.Text.Should().Contain("B", "original slide B moved to position 1");

        h.Dispose();
        ValidatePptx(path, "after swap slides");

        // Reopen: 2 slides still present
        using var pDoc = PresentationDocument.Open(path, false);
        pDoc.PresentationPart!.SlideParts.Should().HaveCount(2, "both slides survive swap+reopen");
    }

    // ==================== 7. Word CopyFrom paragraph ====================

    [Fact]
    public void Word_CopyFrom_Paragraph_CreatesNewParagraph()
    {
        var docPath = Temp("docx");
        using var h = new WordHandler(docPath, editable: true);

        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Source paragraph",
            ["bold"] = "true"
        });

        // Copy paragraph to body
        var newPath = h.CopyFrom("/body/p[1]", "/body", null);
        _out.WriteLine($"CopyFrom returned: {newPath}");

        newPath.Should().NotBeNullOrEmpty("CopyFrom should return a path");

        // Both paragraphs should exist
        var p1 = h.Get("/body/p[1]");
        var p2 = h.Get("/body/p[2]");
        p1.Should().NotBeNull();
        p2.Should().NotBeNull();
        p1!.Text.Should().Contain("Source paragraph");
        p2!.Text.Should().Contain("Source paragraph", "copy should have same text");

        h.Dispose();
        ValidateDocx(docPath, "after CopyFrom paragraph");
    }
}
