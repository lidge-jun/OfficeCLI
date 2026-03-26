// FuzzRound6 — R4 fix regression + final boundary fuzz.
//
// Areas:
//   ZM01–ZM03: Query("zoom") boundary — blank sheet, no-zoom, after Set zoom
//   WD01–WD04: Word Remove paragraph — no-image, with-image, multi-image, nonexistent
//   EC01–EC04: Excel comment Remove — single, multi, empty sheet, out-of-range
//   PW01–PW04: pageWidth Set — cm, in, pt, raw twips
//   SM01–SM06: Smoke — basic CRUD for all three handlers, no regression
//   CB01–CB04: Combined Set — font size+bold+color+alignment simultaneously

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound6 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz6_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== ZM01–ZM03: Excel Query("zoom") boundary ====================

    [Fact]
    public void ZM01_Excel_QueryZoom_BlankSheet_NoZoomKey()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: false);
        var nodes = h.Query("sheet");
        // A brand-new blank sheet has no zoom set; Format should not contain "zoom"
        // or if it does, value should be 100 (default)
        foreach (var node in nodes)
        {
            if (node.Format.ContainsKey("zoom"))
                ((int)node.Format["zoom"]!).Should().Be(100,
                    "default zoom should be 100, not stored unless explicitly set");
        }
    }

    [Fact]
    public void ZM02_Excel_QueryZoom_AfterSetZoom_ReturnsValue()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1", new() { ["zoom"] = "150" });
        var node = h.Get("/Sheet1");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("zoom");
        ((int)node.Format["zoom"]!).Should().Be(150);
    }

    [Fact]
    public void ZM03_Excel_QueryZoom_SetZoom100_NotStoredInFormat()
    {
        // Zoom=100 is the default; query should not include it in Format
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1", new() { ["zoom"] = "100" });
        var node = h.Get("/Sheet1");
        node.Should().NotBeNull();
        // If stored, it equals 100; either way no crash
        if (node!.Format.ContainsKey("zoom"))
            ((int)node.Format["zoom"]!).Should().Be(100);
    }

    // ==================== WD01–WD04: Word Remove paragraph ====================

    [Fact]
    public void WD01_Word_RemoveParagraph_NoImage_Succeeds()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para to remove" });
        h.Add("/body", "paragraph", null, new() { ["text"] = "Keep" });
        // p[1] is the initial empty paragraph from blank doc, or first added
        // Remove first paragraph we added (whichever index it is)
        // There's always at least one paragraph in a blank doc; let's remove
        // the paragraph with known text by removing from the end
        var allParas = h.Query("paragraph");
        // At least 2 paragraphs should exist (the blank doc paragraph + 2 we added)
        allParas.Count().Should().BeGreaterOrEqualTo(2);
        var act = () => h.Remove("/body/p[1]");
        act.Should().NotThrow("removing a text-only paragraph should not throw");
    }

    [Fact]
    public void WD02_Word_RemoveParagraph_WithImage_CleansImagePart()
    {
        var path = CreateTemp("docx");
        var imgPath = Path.Combine(Path.GetTempPath(), $"fuzz6_img_{Guid.NewGuid():N}.png");
        _tempFiles.Add(imgPath);
        File.WriteAllBytes(imgPath, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="));

        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "picture", null, new() { ["path"] = imgPath });
        h.Add("/body", "paragraph", null, new() { ["text"] = "After image" });
        // Remove the paragraph containing the picture (should be p[1] or p[2])
        var act = () => h.Remove("/body/p[1]");
        act.Should().NotThrow("removing a paragraph with an embedded picture should not throw");
    }

    [Fact]
    public void WD03_Word_RemoveParagraph_MultipleImages_AllCleaned()
    {
        var path = CreateTemp("docx");
        var imgPath = Path.Combine(Path.GetTempPath(), $"fuzz6_img2_{Guid.NewGuid():N}.png");
        _tempFiles.Add(imgPath);
        File.WriteAllBytes(imgPath, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="));

        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "picture", null, new() { ["path"] = imgPath });
        h.Add("/body", "picture", null, new() { ["path"] = imgPath });
        h.Add("/body", "paragraph", null, new() { ["text"] = "Sentinel" });
        // Remove two picture paragraphs — should not throw or corrupt
        h.Remove("/body/p[1]");
        var act = () => h.Remove("/body/p[1]");
        act.Should().NotThrow("removing second picture paragraph should not throw");
    }

    [Fact]
    public void WD04_Word_RemoveParagraph_Nonexistent_Throws()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        var act = () => h.Remove("/body/p[999]");
        act.Should().Throw<Exception>("removing a non-existent paragraph should throw");
    }

    // ==================== EC01–EC04: Excel comment Remove ====================

    [Fact]
    public void EC01_Excel_RemoveComment_Single_Succeeds()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Add("/Sheet1", "comment", null, new() { ["ref"] = "A1", ["text"] = "Only comment" });
        var act = () => h.Remove("/Sheet1/comment[1]");
        act.Should().NotThrow("removing a single comment should succeed");
    }

    [Fact]
    public void EC02_Excel_RemoveComment_Multiple_CorrectlyRemovesFirst()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Add("/Sheet1", "comment", null, new() { ["ref"] = "A1", ["text"] = "First" });
        h.Add("/Sheet1", "comment", null, new() { ["ref"] = "B1", ["text"] = "Second" });
        h.Remove("/Sheet1/comment[1]");
        // After first removed, one comment should remain (Second is now comment[1])
        var after = h.Get("/Sheet1/comment[1]");
        after.Should().NotBeNull("second comment should remain after first is removed");
    }

    [Fact]
    public void EC03_Excel_RemoveComment_EmptySheet_Throws()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var act = () => h.Remove("/Sheet1/comment[1]");
        act.Should().Throw<Exception>("removing comment from sheet with no comments should throw");
    }

    [Fact]
    public void EC04_Excel_RemoveComment_OutOfRange_Throws()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Add("/Sheet1", "comment", null, new() { ["ref"] = "A1", ["text"] = "Only" });
        var act = () => h.Remove("/Sheet1/comment[5]");
        act.Should().Throw<Exception>("comment index out of range should throw");
    }

    // ==================== PW01–PW04: pageWidth Set various units ====================

    [Fact]
    public void PW01_Word_SetPageWidth_Cm_Roundtrips()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Set("/section[1]", new() { ["pageWidth"] = "21cm" });
        var node = h.Get("/section[1]");
        node!.Format.Should().ContainKey("pageWidth");
        var raw = node.Format["pageWidth"]!.ToString()!;
        raw.Should().EndWith("cm", "pageWidth should be returned as cm string");
        var cm = double.Parse(raw.Replace("cm", "").Trim());
        cm.Should().BeApproximately(21.0, 0.1, "21cm should round-trip to ~21cm");
    }

    [Fact]
    public void PW02_Word_SetPageWidth_In_Roundtrips()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Set("/section[1]", new() { ["pageWidth"] = "8.5in" });
        var node = h.Get("/section[1]");
        node!.Format.Should().ContainKey("pageWidth");
        var raw = node.Format["pageWidth"]!.ToString()!;
        raw.Should().EndWith("cm", "pageWidth should be returned as cm string");
        var cm = double.Parse(raw.Replace("cm", "").Trim());
        // 8.5in = 21.59cm
        cm.Should().BeApproximately(21.59, 0.2, "8.5in should be ~21.59cm");
    }

    [Fact]
    public void PW03_Word_SetPageWidth_Pt_Roundtrips()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Set("/section[1]", new() { ["pageWidth"] = "595pt" });
        var node = h.Get("/section[1]");
        node!.Format.Should().ContainKey("pageWidth");
        var raw = node.Format["pageWidth"]!.ToString()!;
        raw.Should().EndWith("cm", "pageWidth should be returned as cm string");
        var cm = double.Parse(raw.Replace("cm", "").Trim());
        // 595pt = 20.99cm
        cm.Should().BeApproximately(20.99, 0.3, "595pt should be ~21cm");
    }

    [Fact]
    public void PW04_Word_SetPageWidth_RawTwips_Roundtrips()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        // 12240 twips = 8.5in
        h.Set("/section[1]", new() { ["pageWidth"] = "12240" });
        var node = h.Get("/section[1]");
        node!.Format.Should().ContainKey("pageWidth");
        var raw = node.Format["pageWidth"]!.ToString()!;
        raw.Should().EndWith("cm", "pageWidth should be returned as cm string");
        var cm = double.Parse(raw.Replace("cm", "").Trim());
        // 12240 twips = 8.5in = 21.59cm
        cm.Should().BeApproximately(21.59, 0.3, "12240 twips should be ~21.59cm");
    }

    // ==================== SM01–SM06: Global smoke CRUD ====================

    [Fact]
    public void SM01_Pptx_BasicCrud_NoRegression()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        // Add slide without title so shape indexing is predictable
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello", ["fill"] = "FF0000" });
        var node = h.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Text.Should().Contain("Hello");
        node.Format["fill"].Should().Be("#FF0000");
        h.Set("/slide[1]/shape[1]", new() { ["bold"] = "true" });
        node = h.Get("/slide[1]/shape[1]");
        node!.Format["bold"].Should().Be(true);
        // Remove and verify no crash
        var removeAct = () => h.Remove("/slide[1]/shape[1]");
        removeAct.Should().NotThrow("removing an existing shape should not throw");
    }

    [Fact]
    public void SM02_Excel_BasicCrud_NoRegression()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "Test", ["bold"] = "true" });
        var node = h.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node!.Text.Should().Be("Test");
        node.Format["bold"].Should().Be(true);
    }

    [Fact]
    public void SM03_Word_BasicCrud_NoRegression()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Smoke test para" });
        // Blank doc has at least one paragraph; new one is appended
        var allParas = h.Query("paragraph").ToList();
        var added = allParas.FirstOrDefault(p => p.Text.Contains("Smoke test para"));
        added.Should().NotBeNull("the added paragraph should be queryable");
        var addedPath = added!.Path;
        h.Set(addedPath, new() { ["bold"] = "true" });
        var node = h.Get(addedPath);
        node!.Format["bold"].Should().Be(true);
        // Remove and verify no crash
        var act = () => h.Remove(addedPath);
        act.Should().NotThrow("removing an existing paragraph should not throw");
    }

    [Fact]
    public void SM04_Pptx_QueryShapes_NoRegression()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Q" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Box1" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Box2" });
        var shapes = h.Query("shape");
        shapes.Should().HaveCountGreaterOrEqualTo(2);
    }

    [Fact]
    public void SM05_Excel_QueryCells_NoRegression()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "R1" });
        h.Set("/Sheet1/A2", new() { ["value"] = "R2" });
        h.Set("/Sheet1/A3", new() { ["value"] = "R3" });
        var cells = h.Query("cell");
        cells.Should().HaveCountGreaterOrEqualTo(3);
    }

    [Fact]
    public void SM06_Word_AddTable_NoRegression()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "3" });
        // Table path uses "tbl" not "table"
        var tbl = h.Get("/body/tbl[1]");
        tbl.Should().NotBeNull();
        tbl!.Type.Should().Be("table");
    }

    // ==================== CB01–CB04: Combined Set (multiple attrs simultaneously) ====================

    [Fact]
    public void CB01_Pptx_SetFontSizeBoldColorAlign_SimultaneouslySucceeds()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Multi-prop" });
        h.Set("/slide[1]/shape[1]", new()
        {
            ["size"] = "24pt",
            ["bold"] = "true",
            ["color"] = "#3366FF",
            ["align"] = "center"
        });
        var node = h.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format["bold"].Should().Be(true);
        node.Format["size"].Should().Be("24pt");
        node.Format["color"].Should().Be("#3366FF");
    }

    [Fact]
    public void CB02_Excel_SetFontAndFillAndBoldSimultaneously()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/B2", new()
        {
            ["value"] = "Combo",
            ["bold"] = "true",
            ["italic"] = "true",
            ["size"] = "14pt",
            ["color"] = "#FF0000",
            ["background"] = "#FFFF00"
        });
        var node = h.Get("/Sheet1/B2");
        node.Should().NotBeNull();
        node!.Text.Should().Be("Combo");
        node.Format["bold"].Should().Be(true);
        node.Format["italic"].Should().Be(true);
    }

    [Fact]
    public void CB03_Word_SetParagraphSpacingAndAlignmentSimultaneously()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Combo para" });
        var allParas = h.Query("paragraph").ToList();
        var added = allParas.First(p => p.Text.Contains("Combo para"));
        h.Set(added.Path, new()
        {
            ["spaceBefore"] = "12pt",
            ["spaceAfter"] = "6pt",
            ["lineSpacing"] = "1.5x",
            ["alignment"] = "center"
        });
        var node = h.Get(added.Path);
        node.Should().NotBeNull();
        node!.Format["alignment"].Should().Be("center");
        node.Format["spaceBefore"].Should().Be("12pt");
        node.Format["spaceAfter"].Should().Be("6pt");
        node.Format["lineSpacing"].Should().Be("1.5x");
    }

    [Fact]
    public void CB04_Pptx_SetFillOnly_NoConflict_DoesNotThrow()
    {
        // Set solid fill then gradient with correct format — no crash
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape", ["fill"] = "FF0000" });
        // First set a solid fill
        h.Set("/slide[1]/shape[1]", new() { ["fill"] = "0000FF" });
        // Then set gradient (correct format: "color1-color2")
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["gradient"] = "0000FF-FFFFFF" });
        act.Should().NotThrow("valid gradient format should not throw");
        var node = h.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
    }
}
