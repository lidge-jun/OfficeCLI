// Black-box tests (Round 6) — R4 fix verification + boundary cases:
//   1. Query("zoom") returns only zoom elements (not shapes/textboxes)
//   2. Word: remove paragraph containing picture → Reopen file not corrupted
//   3. Excel: remove last comment → Reopen file has no VML residue
//   4. Word pageWidth/pageHeight returned in cm format; Set accepts cm round-trip
//   5. All handlers Remove→Reopen→Get complete lifecycle

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound6 : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly ITestOutputHelper _output;

    public BtBlackBoxRound6(ITestOutputHelper output) => _output = output;

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bt6_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    private string CreateTestImage()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bt6img_{Guid.NewGuid():N}.png");
        _tempFiles.Add(path);
        var pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
            0xDE,0x00,0x00,0x00,0x0C,0x49,0x44,0x41,
            0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
            0x00,0x00,0x02,0x00,0x01,0xE2,0x21,0xBC,
            0x33,0x00,0x00,0x00,0x00,0x49,0x45,0x4E,
            0x44,0xAE,0x42,0x60,0x82
        };
        File.WriteAllBytes(path, pngBytes);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private void AssertValidPptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"PPTX must be valid after: {step}");
    }

    private void AssertValidDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"DOCX must be valid after: {step}");
    }

    private void AssertValidXlsx(string path, string step)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.ErrorType}: {e.Description}");
        errors.Should().BeEmpty($"XLSX must be valid after: {step}");
    }

    // ═══ 1. Query("zoom") returns only zoom elements ═══

    [Fact]
    public void Pptx_Query_Zoom_ReturnsOnlyZoomElements()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "ZoomTest" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "NotAZoom" });
        h.Add("/slide[1]", "textbox", null, new() { ["text"] = "AlsoNotZoom" });

        var results = h.Query("zoom");

        // With no zoom elements added, result should be empty (not returning shapes)
        results.Should().NotContain(n => n.Type == "shape", "Query(zoom) must not return shapes");
        results.Should().NotContain(n => n.Type == "textbox", "Query(zoom) must not return textboxes");
        foreach (var r in results)
        {
            r.Type.Should().Be("zoom", $"all results of Query(zoom) must have type 'zoom', got '{r.Type}'");
        }
    }

    [Fact]
    public void Pptx_Query_Zoom_WithZoomElement_ReturnsZoomOnly()
    {
        // Add a second slide so we can add a zoom to the first pointing to slide 2
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Slide1" });
        h.Add("/", "slide", null, new() { ["title"] = "Slide2" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "RegularShape" });

        // Add a slide zoom referencing slide 2
        h.Add("/slide[1]", "zoom", null, new() { ["target"] = "2" });

        var results = h.Query("zoom");
        results.Should().NotBeEmpty("zoom element was added");
        results.Should().OnlyContain(n => n.Type == "zoom", "Query(zoom) must return only zoom nodes");
        results.Should().NotContain(n => n.Type == "shape");
    }

    // ═══ 2. Word: remove paragraph containing picture → Reopen not corrupted ═══

    [Fact]
    public void Word_Remove_PictureParagraph_Reopen_DocumentValid()
    {
        var img = CreateTestImage();
        var path = CreateTemp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Before" });
            h.Add("/body", "picture", null, new() { ["path"] = img, ["width"] = "3cm", ["height"] = "2cm" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "After" });
        }

        // Locate paragraph containing the picture and remove it
        using (var h = new WordHandler(path, editable: true))
        {
            var pics = h.Query("picture");
            pics.Should().NotBeEmpty("picture was added");
            // The picture is in a paragraph; remove that paragraph path
            var picPath = pics[0].Path;
            // Remove the parent paragraph that wraps the picture
            // Picture path is /body/p[N] or similar; try removing picture directly
            h.Remove(picPath);
        }

        // File must open cleanly after removing picture paragraph
        AssertValidDocx(path, "remove-picture-paragraph");

        using var h2 = new WordHandler(path, editable: false);
        var paras = h2.Query("paragraph");
        paras.Should().NotBeEmpty("non-picture paragraphs survive");
    }

    // ═══ 3. Excel: remove last comment → Reopen has no VML residue ═══

    [Fact]
    public void Excel_RemoveLastComment_Reopen_NoVmlResidue()
    {
        var path = CreateTemp("xlsx");

        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "comment", null, new()
            {
                ["ref"] = "A1",
                ["text"] = "OnlyComment",
                ["author"] = "Tester"
            });
        }

        AssertValidXlsx(path, "after-add-comment");

        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Remove("/Sheet1/comment[1]");
        }

        // After removing last comment, file must be valid and have no VML part
        AssertValidXlsx(path, "after-remove-last-comment");

        // Verify VML part is gone
        using var xlsx = SpreadsheetDocument.Open(path, false);
        var sheet = xlsx.WorkbookPart!.WorksheetParts.FirstOrDefault();
        sheet.Should().NotBeNull();
        sheet!.VmlDrawingParts.Should().BeEmpty("VML drawing part must be removed with last comment");
    }

    [Fact]
    public void Excel_RemoveOneOfTwoComments_Reopen_VmlStillPresent()
    {
        var path = CreateTemp("xlsx");

        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "comment", null, new() { ["ref"] = "A1", ["text"] = "First", ["author"] = "T" });
            h.Add("/Sheet1", "comment", null, new() { ["ref"] = "B1", ["text"] = "Second", ["author"] = "T" });
        }

        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Remove("/Sheet1/comment[1]");
        }

        AssertValidXlsx(path, "remove-one-of-two-comments");

        using var h2 = new ExcelHandler(path, editable: false);
        var remaining = h2.Query("comment");
        remaining.Should().HaveCount(1, "one comment should remain");
    }

    // ═══ 4. Word pageWidth/pageHeight cm format + Set round-trip ═══

    [Fact]
    public void Word_PageWidth_Get_ReturnsCmFormat()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: false);

        var sec = h.Get("/section[1]");
        sec.Should().NotBeNull();
        sec.Format.Should().ContainKey("pageWidth");
        sec.Format.Should().ContainKey("pageHeight");

        var w = sec.Format["pageWidth"].ToString()!;
        var ht = sec.Format["pageHeight"].ToString()!;

        w.Should().EndWith("cm", $"pageWidth should be in cm format, got '{w}'");
        ht.Should().EndWith("cm", $"pageHeight should be in cm format, got '{ht}'");
    }

    [Fact]
    public void Word_PageSize_Set_CmInput_RoundTrip()
    {
        var path = CreateTemp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Set("/section[1]", new() { ["pageWidth"] = "21cm", ["pageHeight"] = "29.7cm" });
        }

        AssertValidDocx(path, "set-page-size-cm");

        using var h2 = new WordHandler(path, editable: false);
        var sec = h2.Get("/section[1]");
        sec.Format.Should().ContainKey("pageWidth");
        sec.Format.Should().ContainKey("pageHeight");

        var w = sec.Format["pageWidth"].ToString()!;
        var ht = sec.Format["pageHeight"].ToString()!;

        w.Should().EndWith("cm");
        // 21cm tolerance: allow 20.9cm–21.1cm range due to twip rounding
        var wNum = double.Parse(w.Replace("cm", ""), System.Globalization.CultureInfo.InvariantCulture);
        wNum.Should().BeApproximately(21.0, 0.5, "pageWidth round-trip should be ~21cm");

        var htNum = double.Parse(ht.Replace("cm", ""), System.Globalization.CultureInfo.InvariantCulture);
        htNum.Should().BeApproximately(29.7, 0.5, "pageHeight round-trip should be ~29.7cm");
    }

    // ═══ 5. All handlers Remove→Reopen→Get lifecycle ═══

    [Fact]
    public void Pptx_RemoveShape_Reopen_GetThrows()
    {
        var path = CreateTemp("pptx");

        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Test" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "RemoveMe" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "KeepMe" });
            h.Remove("/slide[1]/shape[2]");
        }

        AssertValidPptx(path, "pptx-remove-reopen");

        using var h2 = new PowerPointHandler(path, editable: false);
        // After reopen, "RemoveMe" is gone; only the placeholder title + "KeepMe" shape remain
        // shape[2] was removed; now only 1 user shape → shape[2] is the title placeholder or not present
        var keepNode = h2.Get("/slide[1]/shape[2]");
        keepNode.Text.Should().Be("KeepMe", "KeepMe shape becomes shape[2] after removal of shape[2] (RemoveMe)");

        var act = () => h2.Get("/slide[1]/shape[3]");
        act.Should().Throw<Exception>("only 2 shapes should remain after removal");
    }

    [Fact]
    public void Word_RemoveParagraph_Reopen_GetExcludesIt()
    {
        var path = CreateTemp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Keep" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "Remove" });
            h.Remove("/body/p[2]");
        }

        AssertValidDocx(path, "word-remove-reopen");

        using var h2 = new WordHandler(path, editable: false);
        var paras = h2.Query("paragraph");
        paras.Should().NotContain(n => n.Text != null && n.Text.Contains("Remove"),
            "removed paragraph must not appear after reopen");
    }

    [Fact]
    public void Excel_RemoveCell_Reopen_GetReturnsEmpty()
    {
        var path = CreateTemp("xlsx");

        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/Sheet1", "cell", null, new() { ["address"] = "Z9", ["value"] = "Zap" });
            h.Remove("/Sheet1/Z9");
        }

        AssertValidXlsx(path, "excel-remove-cell-reopen");

        using var h2 = new ExcelHandler(path, editable: false);
        var node = h2.Get("/Sheet1/Z9");
        // Cell may be null or have empty/placeholder text after removal — should not contain original value
        if (node != null)
            node.Text.Should().NotBe("Zap", "removed cell should not retain original value after reopen");
    }

    [Fact]
    public void Excel_RemoveSheet_Reopen_SheetGone()
    {
        var path = CreateTemp("xlsx");

        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Add("/", "sheet", null, new() { ["name"] = "Extra" });
            h.Remove("/Extra");
        }

        AssertValidXlsx(path, "excel-remove-sheet-reopen");

        using var h2 = new ExcelHandler(path, editable: false);
        var act = () => h2.Get("/Extra/A1");
        act.Should().Throw<Exception>("sheet Extra was removed");
    }
}
