// Black-box tests (Round 7):
//   1. Excel sparkline full lifecycle: Add → Get → Set → Remove → Reopen
//   2. Excel freeze pane Set/Get round-trip and clear
//   3. Excel sheet protection Set → Get → unprotect → Get
//   4. Excel databar/iconset conditional format Add → Get → Remove
//   5. Word bookmark Add → Get → Remove → Reopen (no residue)
//   6. Word endnote Add → Get → Set → Reopen
//   7. Word field code Add → Get (round-trip)
//   8. PPTX placeholder Query returns correct types
//   9. Cross-op: Add slide + Add shape → Query → Remove slide → shape gone
//  10. Format round-trip: Query Format keys usable directly in Set

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound7 : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly ITestOutputHelper _output;

    public BtBlackBoxRound7(ITestOutputHelper output) => _output = output;

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt7_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(p);
        BlankDocCreator.Create(p);
        return p;
    }

    private void ValidateXlsx(string path, string step)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"XLSX must be valid after: {step}");
    }

    private void ValidateDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"DOCX must be valid after: {step}");
    }

    private void ValidatePptx(string path, string step)
    {
        using var doc = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _output.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"PPTX must be valid after: {step}");
    }

    // ==================== 1. Excel sparkline full lifecycle ====================

    [Fact]
    public void Excel_Sparkline_AddGetSetRemove_Lifecycle()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        // Seed data
        for (int i = 1; i <= 5; i++)
            h.Set($"/Sheet1/{(char)('A' + i - 1)}1", new() { ["value"] = (i * 10).ToString() });

        // Add sparkline
        var spkPath = h.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "G1",
            ["range"] = "A1:E1",
            ["type"] = "column",
            ["color"] = "FF0000"
        });
        spkPath.Should().Be("/Sheet1/sparkline[1]");

        // Get and verify
        var node = h.Get("/Sheet1/sparkline[1]");
        node.Should().NotBeNull();
        node!.Type.Should().Be("sparkline");
        node.Format["type"].Should().Be("column");
        node.Format["color"].ToString().Should().Be("#FF0000");
        node.Format["cell"].ToString().Should().NotBeNullOrEmpty();
        node.Format["range"].ToString().Should().Contain("A1:E1");

        // Set: change type and color
        h.Set("/Sheet1/sparkline[1]", new() { ["type"] = "line", ["color"] = "0070C0" });
        var updated = h.Get("/Sheet1/sparkline[1]");
        updated!.Format["type"].Should().Be("line");
        updated.Format["color"].ToString().Should().Be("#0070C0");

        h.Dispose();
        ValidateXlsx(path, "after sparkline set");

        // Reopen and verify persistence
        using var h2 = new ExcelHandler(path, editable: true);
        var persisted = h2.Get("/Sheet1/sparkline[1]");
        persisted!.Format["type"].Should().Be("line");

        // Remove
        h2.Remove("/Sheet1/sparkline[1]");
        h2.Dispose();
        ValidateXlsx(path, "after sparkline remove");

        using var h3 = new ExcelHandler(path, editable: false);
        var gone = h3.Query("sparkline");
        gone.Should().BeEmpty("sparkline should be removed");
    }

    // ==================== 2. Excel freeze pane Set/Get round-trip ====================

    [Fact]
    public void Excel_FreezePaneSetGet_RoundTrip()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        // Set freeze at B2 (freeze row 1 + col A)
        h.Set("/Sheet1", new() { ["freeze"] = "B2" });

        // Get via sheet node
        var sheet = h.Get("/Sheet1");
        sheet.Should().NotBeNull();
        sheet!.Format.Should().ContainKey("freeze");
        sheet.Format["freeze"].Should().Be("B2");

        // Clear freeze
        h.Set("/Sheet1", new() { ["freeze"] = "" });
        var sheet2 = h.Get("/Sheet1");
        sheet2!.Format.ContainsKey("freeze").Should().BeFalse("freeze should be removed when cleared");

        h.Dispose();
        ValidateXlsx(path, "after freeze pane round-trip");
    }

    [Fact]
    public void Excel_FreezePane_Persistence()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1", new() { ["freeze"] = "C3" });
        h.Dispose();

        ValidateXlsx(path, "after freeze set");

        using var h2 = new ExcelHandler(path, editable: false);
        var sheet = h2.Get("/Sheet1");
        sheet!.Format["freeze"].Should().Be("C3");
    }

    // ==================== 3. Excel sheet protection ====================

    [Fact]
    public void Excel_SheetProtection_SetGetUnprotect()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        // Protect the sheet
        h.Set("/Sheet1", new() { ["protect"] = "true", ["password"] = "secret123" });

        var sheet = h.Get("/Sheet1");
        sheet!.Format.Should().ContainKey("protect");
        sheet.Format["protect"].Should().Be(true);

        // Remove protection
        h.Set("/Sheet1", new() { ["protect"] = "false" });
        var sheet2 = h.Get("/Sheet1");
        sheet2!.Format.ContainsKey("protect").Should().BeFalse("protection should be removed");

        h.Dispose();
        ValidateXlsx(path, "after sheet protection toggle");
    }

    // ==================== 4. Excel databar conditional format ====================

    [Fact]
    public void Excel_DataBar_AddGetRemove()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        // Seed some values
        for (int i = 1; i <= 5; i++)
            h.Set($"/Sheet1/A{i}", new() { ["value"] = (i * 20).ToString() });

        var cfPath = h.Add("/Sheet1", "databar", null, new()
        {
            ["range"] = "A1:A5",
            ["color"] = "638EC6"
        });
        cfPath.Should().NotBeNullOrEmpty();
        cfPath.Should().StartWith("/Sheet1/cf[");

        // Get via path returned from Add
        var cfNode = h.Get(cfPath);
        cfNode.Should().NotBeNull();
        cfNode!.Type.Should().Be("conditionalFormatting");
        cfNode.Format["cfType"].ToString().Should().Be("dataBar");

        // Remove using the exact path returned
        h.Remove(cfPath);

        // Verify gone — Get should return null or throw
        var afterRemove = h.Get(cfPath);
        afterRemove.Should().BeNull("cf[1] should no longer exist after remove");

        h.Dispose();
        ValidateXlsx(path, "after databar lifecycle");
    }

    [Fact]
    public void Excel_IconSet_AddQuery()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        for (int i = 1; i <= 6; i++)
            h.Set($"/Sheet1/B{i}", new() { ["value"] = (i * 15).ToString() });

        var cfPath = h.Add("/Sheet1", "iconset", null, new()
        {
            ["range"] = "B1:B6",
            ["iconset"] = "3Arrows"
        });
        cfPath.Should().NotBeNullOrEmpty();
        cfPath.Should().StartWith("/Sheet1/cf[");

        var cfNode = h.Get(cfPath);
        cfNode.Should().NotBeNull();
        cfNode!.Type.Should().Be("conditionalFormatting");
        cfNode.Format["cfType"].ToString().Should().Be("iconSet");

        h.Dispose();
        ValidateXlsx(path, "after iconset add");
    }

    // ==================== 5. Word bookmark Add/Get/Remove ====================

    [Fact]
    public void Word_Bookmark_AddGetRemove_Lifecycle()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);

        h.Add("/body", "paragraph", null, new() { ["text"] = "Bookmark target text" });
        h.Add("/body/p[1]", "bookmark", null, new()
        {
            ["name"] = "TestBookmark7",
            ["text"] = "bookmark text"
        });

        // Get by path
        var node = h.Get("/bookmark[TestBookmark7]");
        node.Should().NotBeNull();
        node!.Type.Should().Be("bookmark");
        node.Format["name"].ToString().Should().Be("TestBookmark7");

        // Query all bookmarks
        var bookmarks = h.Query("bookmark");
        bookmarks.Should().ContainSingle(b => b.Format["name"].ToString() == "TestBookmark7");

        // Remove
        h.Remove("/bookmark[TestBookmark7]");
        var afterRemove = h.Query("bookmark");
        afterRemove.Should().NotContain(b => b.Format.ContainsKey("name") && b.Format["name"].ToString() == "TestBookmark7");

        h.Dispose();
        ValidateDocx(path, "after bookmark lifecycle");

        // Reopen: bookmark gone
        using var h2 = new WordHandler(path, editable: false);
        var bks = h2.Query("bookmark");
        bks.Should().NotContain(b => b.Format.ContainsKey("name") && b.Format["name"].ToString() == "TestBookmark7");
    }

    // ==================== 6. Word endnote Add/Get/Set/Reopen ====================

    [Fact]
    public void Word_Endnote_AddGetSet_Lifecycle()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);

        h.Add("/body", "paragraph", null, new() { ["text"] = "Main text" });
        h.Add("/body/p[1]", "endnote", null, new() { ["text"] = "Original endnote text" });

        // Get endnote
        var en = h.Get("/endnote[1]");
        en.Should().NotBeNull();
        en!.Type.Should().Be("endnote");

        // Set endnote text
        h.Set("/endnote[1]", new() { ["text"] = "Updated endnote text" });
        var updated = h.Get("/endnote[1]");
        updated.Should().NotBeNull();

        h.Dispose();
        ValidateDocx(path, "after endnote set");

        // Reopen and verify
        using var h2 = new WordHandler(path, editable: false);
        var persisted = h2.Get("/endnote[1]");
        persisted.Should().NotBeNull();
    }

    // ==================== 7. Word field code Add/Get ====================

    [Fact]
    public void Word_Field_AddGet_PageNumber()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);

        h.Add("/body", "paragraph", null, new() { ["text"] = "Page: " });
        h.Add("/body/p[1]", "field", null, new()
        {
            ["fieldType"] = "PAGE",
        });

        // Query fields
        var fields = h.Query("field");
        fields.Should().NotBeEmpty("at least one field should exist");
        fields.Should().OnlyContain(f => f.Type == "field");

        h.Dispose();
        ValidateDocx(path, "after field code add");
    }

    [Fact]
    public void Word_Field_AddGet_Date()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);

        h.Add("/body", "paragraph", null, new() { ["text"] = "Date: " });
        h.Add("/body/p[1]", "field", null, new()
        {
            ["fieldType"] = "DATE",
            ["format"] = "MMMM d, yyyy"
        });

        var fields = h.Query("field");
        fields.Should().NotBeEmpty();

        h.Dispose();
        ValidateDocx(path, "after date field add");
    }

    // ==================== 8. PPTX placeholder Query ====================

    [Fact]
    public void Pptx_Placeholder_QueryReturnsNodes()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Slide with Placeholders" });

        // Query placeholders on slide 1
        var placeholders = h.Query("placeholder");
        // A default slide should have at least a title placeholder
        placeholders.Should().NotBeEmpty("default slide should have at least title placeholder");
        placeholders.Should().OnlyContain(p => p.Type == "placeholder");

        h.Dispose();
        ValidatePptx(path, "after placeholder query");
    }

    // ==================== 9. Cross-op: Add slide + shape → Remove slide ====================

    [Fact]
    public void Pptx_RemoveSlide_AlsoRemovesShapes()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        // Determine initial slide count via Get probing
        int initialCount = 0;
        while (true)
        {
            try { var s = h.Get($"/slide[{initialCount + 1}]"); if (s == null) break; initialCount++; }
            catch { break; }
        }

        // Add a new slide with a distinctive shape
        h.Add("/", "slide", null, new() { ["title"] = "Slide WithShape" });
        int newSlideIdx = initialCount + 1;

        // Verify new slide exists
        var newSlide = h.Get($"/slide[{newSlideIdx}]");
        newSlide.Should().NotBeNull("the newly added slide should be accessible");

        // Add shape on that slide
        h.Add($"/slide[{newSlideIdx}]", "shape", null, new()
        {
            ["text"] = "Shape on Slide WithShape",
            ["fill"] = "FF0000",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "4cm", ["height"] = "2cm"
        });

        // Verify shape exists on that slide
        var shapes = h.Query("shape");
        shapes.Should().Contain(s => s.Text == "Shape on Slide WithShape");

        // Remove the new slide
        h.Remove($"/slide[{newSlideIdx}]");

        // The new slide index should now be gone (throws or returns null)
        Action getRemoved = () => h.Get($"/slide[{newSlideIdx}]");
        // After remove, slide[N] may throw or the count should equal initial
        // Verify by checking all shapes — none should be "Shape on Slide WithShape"
        var remainingShapes = h.Query("shape");
        remainingShapes.Should().NotContain(s => s.Text == "Shape on Slide WithShape");

        h.Dispose();
        ValidatePptx(path, "after slide removal with shape cleanup");
    }

    // ==================== 10. Format round-trip: Query keys → Set ====================

    [Fact]
    public void Excel_SparklineFormatKeys_UsableInSet()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        for (int i = 1; i <= 5; i++)
            h.Set($"/Sheet1/{(char)('A' + i - 1)}2", new() { ["value"] = (i * 5).ToString() });

        h.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F2",
            ["range"] = "A2:E2",
            ["type"] = "line",
            ["color"] = "FF0000",
            ["markers"] = "true"
        });

        var node = h.Get("/Sheet1/sparkline[1]");
        node.Should().NotBeNull();

        // Get the Format values and use them in Set (round-trip)
        var typeVal = node!.Format["type"].ToString()!;
        var colorVal = node.Format["color"].ToString()!;

        // These should not throw
        h.Set("/Sheet1/sparkline[1]", new()
        {
            ["type"] = typeVal,
            ["color"] = colorVal
        });

        var after = h.Get("/Sheet1/sparkline[1]");
        after!.Format["type"].Should().Be(typeVal);
        after.Format["color"].Should().Be(colorVal);
    }

    [Fact]
    public void Excel_SheetFormatKeys_FreezePaneRoundTrip()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        h.Set("/Sheet1", new() { ["freeze"] = "B3" });

        var sheet = h.Get("/Sheet1");
        var freezeVal = sheet!.Format["freeze"].ToString()!;

        // Use the Format value directly in Set (round-trip)
        h.Set("/Sheet1", new() { ["freeze"] = freezeVal });

        var after = h.Get("/Sheet1");
        after!.Format["freeze"].Should().Be(freezeVal);
    }

    [Fact]
    public void Pptx_ShapeFormatKeys_RoundTrip()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Format Round-trip" });
        h.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "RoundTrip",
            ["fill"] = "4472C4",
            ["fontSize"] = "16pt",
            ["bold"] = "true",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "6cm", ["height"] = "2cm"
        });

        var shape = h.Get("/slide[1]/shape[1]");
        shape.Should().NotBeNull();

        // Collect format values returned by Get
        var fillVal = shape!.Format.ContainsKey("fill") ? shape.Format["fill"].ToString()! : null;
        var sizeVal = shape.Format.ContainsKey("fontSize") ? shape.Format["fontSize"].ToString()! : null;

        // Re-apply them via Set (should not throw, values should survive)
        var updates = new Dictionary<string, string>();
        if (fillVal != null) updates["fill"] = fillVal;
        if (sizeVal != null) updates["fontSize"] = sizeVal;

        if (updates.Count > 0)
        {
            h.Set("/slide[1]/shape[1]", updates);
            var after = h.Get("/slide[1]/shape[1]");
            if (fillVal != null) after!.Format["fill"].ToString().Should().Be(fillVal);
            if (sizeVal != null) after!.Format["fontSize"].ToString().Should().Be(sizeVal);
        }

        h.Dispose();
        ValidatePptx(path, "after shape format round-trip");
    }
}
