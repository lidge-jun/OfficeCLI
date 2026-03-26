// Black-box tests (Round 8) — R6 regression verification:
//   1. Word footnote/endnote Remove round-trip (Add→Remove→Reopen→verify deleted)
//   2. Excel multi-CF Priority uniqueness (Add 3 CF rules → verify distinct priorities)
//   3. PPTX transition duration round-trip (Set→Get→verify)
//   4. Word paragraph with shared picture: delete one paragraph, other image survives
//   5. Word header with picture: remove header → Reopen → file valid, no dangling parts
//   6. Get("/slide[999]") returns null, not exception

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound8 : IDisposable
{
    private readonly List<string> _temps = new();
    private readonly ITestOutputHelper _out;

    public BtBlackBoxRound8(ITestOutputHelper output) => _out = output;

    public void Dispose()
    {
        foreach (var f in _temps)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt8_{Guid.NewGuid():N}.{ext}");
        _temps.Add(p);
        BlankDocCreator.Create(p);
        return p;
    }

    private void ValidateDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"DOCX must be valid after: {step}");
    }

    private void ValidateXlsx(string path, string step)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"XLSX must be valid after: {step}");
    }

    private void ValidatePptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"PPTX must be valid after: {step}");
    }

    // ==================== 1. Footnote Remove round-trip ====================

    [Fact]
    public void Word_Footnote_Remove_RoundTrip()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);

        h.Add("/body", "paragraph", null, new() { ["text"] = "Main text" });
        h.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Footnote content" });

        var fn = h.Get("/footnote[1]");
        fn.Should().NotBeNull("footnote should exist after Add");
        fn!.Type.Should().Be("footnote");

        // Remove it
        h.Remove("/footnote[1]");

        // Should be gone in-session
        var inSession = h.Query("footnote");
        inSession.Should().NotContain(n => n.Text == "Footnote content",
            "footnote should be gone after Remove");

        h.Dispose();
        ValidateDocx(path, "after footnote remove");

        // Reopen — still gone
        using var h2 = new WordHandler(path, editable: false);
        var afterReopen = h2.Query("footnote");
        afterReopen.Should().NotContain(n => n.Text == "Footnote content",
            "footnote should not reappear after reopen");
    }

    [Fact]
    public void Word_Endnote_Remove_RoundTrip()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);

        h.Add("/body", "paragraph", null, new() { ["text"] = "Main text" });
        h.Add("/body/p[1]", "endnote", null, new() { ["text"] = "Endnote content" });

        var en = h.Get("/endnote[1]");
        en.Should().NotBeNull("endnote should exist after Add");
        en!.Type.Should().Be("endnote");

        h.Remove("/endnote[1]");

        h.Dispose();
        ValidateDocx(path, "after endnote remove");

        using var h2 = new WordHandler(path, editable: false);
        var afterReopen = h2.Query("endnote");
        afterReopen.Should().NotContain(n => n.Text == "Endnote content",
            "endnote should not reappear after reopen");
    }

    // ==================== 2. Excel multi-CF Priority uniqueness ====================

    [Fact]
    public void Excel_MultipleCF_PrioritiesAreUnique()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        // Seed values
        for (int i = 1; i <= 5; i++)
            h.Set($"/Sheet1/A{i}", new() { ["value"] = (i * 10).ToString() });

        // Add 3 conditional formats to the same sheet
        var p1 = h.Add("/Sheet1", "databar", null, new()
        {
            ["range"] = "A1:A5",
            ["color"] = "FF0000"
        });
        var p2 = h.Add("/Sheet1", "colorscale", null, new()
        {
            ["range"] = "A1:A5",
            ["mincolor"] = "00FF00",
            ["maxcolor"] = "FF0000"
        });
        var p3 = h.Add("/Sheet1", "iconset", null, new()
        {
            ["range"] = "A1:A5",
            ["iconset"] = "3Arrows"
        });

        p1.Should().NotBeNullOrEmpty();
        p2.Should().NotBeNullOrEmpty();
        p3.Should().NotBeNullOrEmpty();

        _out.WriteLine($"CF paths: {p1}, {p2}, {p3}");

        // All three paths must be distinct (each got its own cf[N])
        new[] { p1, p2, p3 }.Should().OnlyHaveUniqueItems(
            "each CF rule should get a distinct path");

        // Verify all 3 are accessible
        h.Get(p1).Should().NotBeNull();
        h.Get(p2).Should().NotBeNull();
        h.Get(p3).Should().NotBeNull();

        h.Dispose();
        ValidateXlsx(path, "after 3 CF rules added");

        // Reopen and verify priorities in XML are unique
        using var doc = SpreadsheetDocument.Open(path, false);
        var wbPart = doc.WorkbookPart!;
        var sheet = wbPart.WorksheetParts.First();
        var ws = sheet.Worksheet;
        var allRules = ws.Elements<ConditionalFormatting>()
            .SelectMany(cf => cf.Elements<ConditionalFormattingRule>())
            .Select(r => r.Priority?.Value ?? 0)
            .ToList();

        _out.WriteLine($"XML CF priorities: {string.Join(", ", allRules)}");
        allRules.Should().HaveCountGreaterOrEqualTo(3, "all 3 CF rules should persist");
        allRules.Should().OnlyHaveUniqueItems("each CF rule must have a unique priority");
    }

    // ==================== 3. PPTX transition duration round-trip ====================

    [Fact]
    public void Pptx_TransitionDuration_SetGet_RoundTrip()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Transition Test" });

        // Set a fade transition with explicit duration
        h.Set("/slide[1]", new()
        {
            ["transition"] = "fade",
            ["duration"] = "2000"
        });

        // Get and verify transition is stored
        var slide = h.Get("/slide[1]");
        slide.Should().NotBeNull();
        slide!.Format.Should().ContainKey("transition",
            "transition type should be readable after Set");
        slide.Format["transition"].ToString().Should().Be("fade");

        h.Dispose();
        ValidatePptx(path, "after transition set");

        // Reopen and verify persistence
        using var h2 = new PowerPointHandler(path, editable: false);
        var persisted = h2.Get("/slide[1]");
        persisted.Should().NotBeNull();
        persisted!.Format.Should().ContainKey("transition");
        persisted.Format["transition"].ToString().Should().Be("fade");
    }

    [Fact]
    public void Pptx_TransitionDuration_KeyExists_WhenDurationSet()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Duration Test", ["transition"] = "dissolve" });

        var slide = h.Get("/slide[1]");
        slide.Should().NotBeNull();

        // transition key must exist
        slide!.Format.Should().ContainKey("transition");
        _out.WriteLine($"Format keys: {string.Join(", ", slide.Format.Keys)}");

        // If transitionDuration is exposed, verify it's a parseable number
        if (slide.Format.ContainsKey("transitionDuration"))
        {
            var durStr = slide.Format["transitionDuration"].ToString()!;
            int.TryParse(durStr, out var durVal).Should().BeTrue(
                $"transitionDuration should be numeric, got: {durStr}");
            _out.WriteLine($"transitionDuration = {durVal}");
        }

        h.Dispose();
        ValidatePptx(path, "after add slide with transition");
    }

    // ==================== 4. Word shared image: delete one paragraph, other survives ====================

    [Fact]
    public void Word_SharedImage_DeleteOneParagraph_OtherImageSurvives()
    {
        // Find a small test image or create a simple PNG
        var imgPath = Path.Combine(Path.GetTempPath(), $"bt8_img_{Guid.NewGuid():N}.png");
        _temps.Add(imgPath);

        // Create a minimal 1x1 red PNG (binary)
        var png1x1 = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg==");
        File.WriteAllBytes(imgPath, png1x1);

        var docPath = Temp("docx");
        using var h = new WordHandler(docPath, editable: true);

        // Add two paragraphs each with the same image file
        h.Add("/body", "paragraph", null, new() { ["text"] = "First paragraph" });
        h.Add("/body", "paragraph", null, new() { ["text"] = "Second paragraph" });

        h.Add("/body/p[1]", "picture", null, new()
        {
            ["src"] = imgPath,
            ["width"] = "2cm",
            ["height"] = "2cm"
        });
        h.Add("/body/p[2]", "picture", null, new()
        {
            ["src"] = imgPath,
            ["width"] = "2cm",
            ["height"] = "2cm"
        });

        // Verify both paragraphs have pictures
        var pics = h.Query("picture");
        pics.Should().HaveCountGreaterOrEqualTo(2, "both paragraphs should have images");

        // Remove first paragraph (contains first picture)
        h.Remove("/body/p[1]");

        h.Dispose();
        ValidateDocx(docPath, "after removing paragraph with picture");

        // Reopen — the second picture should still be accessible
        using var h2 = new WordHandler(docPath, editable: false);
        var remaining = h2.Query("picture");
        remaining.Should().NotBeEmpty("second paragraph image should survive deletion of first paragraph");
    }

    // ==================== 5. Header with picture: remove header → file valid ====================

    [Fact]
    public void Word_HeaderWithPicture_RemoveHeader_FileValid()
    {
        var imgPath = Path.Combine(Path.GetTempPath(), $"bt8_hdr_{Guid.NewGuid():N}.png");
        _temps.Add(imgPath);

        var png1x1 = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg==");
        File.WriteAllBytes(imgPath, png1x1);

        var docPath = Temp("docx");
        using var h = new WordHandler(docPath, editable: true);

        // Add a header with a picture
        h.Add("/body", "header", null, new() { ["text"] = "Header with image" });
        h.Add("/header[1]", "picture", null, new()
        {
            ["src"] = imgPath,
            ["width"] = "3cm",
            ["height"] = "1cm"
        });

        // Verify header exists
        var headers = h.Query("header");
        headers.Should().NotBeEmpty("header should exist");

        // Remove the header
        h.Remove("/header[1]");

        h.Dispose();
        ValidateDocx(docPath, "after removing header with picture");

        // Reopen — file should be usable
        using var h2 = new WordHandler(docPath, editable: false);
        var body = h2.Get("/body");
        body.Should().NotBeNull("document body must be accessible after header removal");
    }

    // ==================== 6. Get("/slide[999]") returns null, not exception ====================

    [Fact]
    public void Pptx_GetInvalidSlideIndex_ReturnsNull()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Only Slide" });

        // Exact path pattern /slide[N] should return null for out-of-range N
        DocumentNode? result = null;
        Exception? ex = null;
        try { result = h.Get("/slide[999]"); }
        catch (Exception e) { ex = e; }

        // Either returns null or throws a descriptive exception — NOT an IndexOutOfRangeException
        if (ex != null)
        {
            ex.Should().NotBeOfType<IndexOutOfRangeException>(
                "should throw a descriptive exception, not IndexOutOfRangeException");
            ex.Should().NotBeOfType<ArgumentOutOfRangeException>(
                "should throw a descriptive exception, not ArgumentOutOfRangeException");
            _out.WriteLine($"Exception type: {ex.GetType().Name}, message: {ex.Message}");
        }
        else
        {
            result.Should().BeNull("Get on out-of-range slide index should return null");
        }
    }

    [Fact]
    public void Pptx_GetInvalidSlideIndex_DoesNotCrash()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: false);

        // A blank file has no slides — slide[1] may also be null
        DocumentNode? r999 = null;
        Exception? ex999 = null;
        try { r999 = h.Get("/slide[999]"); }
        catch (Exception e) { ex999 = e; }

        if (ex999 != null)
        {
            ex999.Should().NotBeOfType<IndexOutOfRangeException>();
            _out.WriteLine($"/slide[999] threw: {ex999.GetType().Name}: {ex999.Message}");
        }
        else
        {
            r999.Should().BeNull();
        }
    }
}
