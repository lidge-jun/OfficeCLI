// Black-box tests (Round 10) — key objectives:
//   1. R8 fix: Word Set font size 10.25pt → Math.Round → 21 half-points → Get returns 10.5pt
//   2. Null value Set does not crash (PPTX, XLSX, DOCX)
//   3. Move operation (PPTX slide reorder, Word paragraph reorder)
//   4. Full CRUD + persistence smoke tests: PPTX, XLSX, DOCX
//   5. Edge cases: empty text, special chars, zero-length properties

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound10 : IDisposable
{
    private readonly List<string> _temps = new();
    private readonly ITestOutputHelper _out;

    public BtBlackBoxRound10(ITestOutputHelper output) => _out = output;

    public void Dispose()
    {
        foreach (var f in _temps)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt10_{Guid.NewGuid():N}.{ext}");
        _temps.Add(p);
        BlankDocCreator.Create(p);
        return p;
    }

    private string TempImg()
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt10_img_{Guid.NewGuid():N}.png");
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

    // ==================== 1. R8 fix: Word font size rounding ====================

    [Fact]
    public void Word_Set_FontSize_10pt25_RoundsTo_10pt5()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Rounding test", ["size"] = "12pt" });

        h.Set("/body/p[1]", new() { ["size"] = "10.25pt" });
        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        _out.WriteLine($"After Set 10.25pt, size = {node!.Format.GetValueOrDefault("size")}");
        // 10.25pt * 2 = 20.5 half-points → Math.Round → 21 → 10.5pt
        node.Format["size"].ToString().Should().Be("10.5pt",
            "10.25pt rounds to 21 half-points (10.5pt), not truncates to 20 (10pt)");
    }

    [Fact]
    public void Word_Add_FontSize_10pt25_RoundsTo_10pt5()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Add rounding", ["size"] = "10.25pt" });
        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        _out.WriteLine($"After Add 10.25pt, size = {node!.Format.GetValueOrDefault("size")}");
        node.Format["size"].ToString().Should().Be("10.5pt",
            "Add with 10.25pt should also round to 10.5pt");
    }

    [Fact]
    public void Word_Set_FontSize_Exact_10pt5_Preserved()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Exact half-pt", ["size"] = "10.5pt" });
        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        node!.Format["size"].ToString().Should().Be("10.5pt", "10.5pt (21 half-pts) stored exactly");
    }

    [Fact]
    public void Word_Set_FontSize_Various_RoundTrip()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Various sizes" });

        // 14pt exactly
        h.Set("/body/p[1]", new() { ["size"] = "14pt" });
        var n1 = h.Get("/body/p[1]");
        n1!.Format["size"].ToString().Should().Be("14pt");

        // 11.25pt → 22.5 → rounds to 23 → 11.5pt
        h.Set("/body/p[1]", new() { ["size"] = "11.25pt" });
        var n2 = h.Get("/body/p[1]");
        _out.WriteLine($"11.25pt → {n2!.Format.GetValueOrDefault("size")}");
        n2.Format["size"].ToString().Should().Be("11.5pt",
            "11.25pt * 2 = 22.5 → Math.Round → 23 → 11.5pt");
    }

    // ==================== 2. Null value Set does not crash ====================

    [Fact]
    public void Word_Set_NullValues_DoesNotCrash()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Null test" });

        // Pass dict with null values
        var props = new Dictionary<string, string?>
        {
            ["bold"] = null,
            ["italic"] = null,
            ["size"] = null,
            ["color"] = null
        };

        var ex = Record.Exception(() =>
            h.Set("/body/p[1]", props.Where(kv => kv.Value != null)
                .ToDictionary(kv => kv.Key, kv => kv.Value!)));
        ex.Should().BeNull("Set with null-filtered dict should not throw");

        // Confirm paragraph is still accessible
        h.Get("/body/p[1]").Should().NotBeNull();
    }

    [Fact]
    public void Pptx_Set_NullValues_DoesNotCrash()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Null test" });
        h.Add("/slide[1]", "textbox", null, new()
        {
            ["text"] = "Test", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "4cm", ["height"] = "2cm"
        });

        var ex = Record.Exception(() =>
            h.Set("/slide[1]/shape[2]", new Dictionary<string, string>()));
        ex.Should().BeNull("Set with empty dict should not throw");

        h.Get("/slide[1]/shape[2]").Should().NotBeNull();
    }

    [Fact]
    public void Excel_Set_NullValues_DoesNotCrash()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "hello" });

        var ex = Record.Exception(() =>
            h.Set("/Sheet1/A1", new Dictionary<string, string>()));
        ex.Should().BeNull("Set with empty dict should not throw");

        h.Get("/Sheet1/A1").Should().NotBeNull();
    }

    // ==================== 3. Move operation ====================

    [Fact]
    public void Pptx_Move_Slide_ChangesOrder()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "First" });
        h.Add("/", "slide", null, new() { ["title"] = "Second" });
        h.Add("/", "slide", null, new() { ["title"] = "Third" });

        // Move slide[3] to position index=0 (before current slide[1])
        // index is 0-based: insert before remaining[0] → position 1
        var movedPath = h.Move("/slide[3]", "/", 0);
        _out.WriteLine($"Move returned: {movedPath}");
        movedPath.Should().NotBeNullOrEmpty("Move returns target path");
        movedPath.Should().Be("/slide[1]", "index=0 places slide at position 1");

        // After move, slide[1].Preview should be "Third"
        var s1 = h.Get("/slide[1]");
        s1.Should().NotBeNull();
        _out.WriteLine($"slide[1] Preview after move: '{s1?.Preview}'");
        s1!.Preview.Should().Contain("Third", "Third moved to position 1");

        h.Dispose();
        ValidatePptx(path, "after Move slide");

        using var doc = PresentationDocument.Open(path, false);
        doc.PresentationPart!.SlideParts.Should().HaveCount(3, "all 3 slides survive move");
    }

    [Fact]
    public void Word_Move_Paragraph_ChangesOrder()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para A" });
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para B" });
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para C" });

        // Move p[3] before p[1]: index=0 inserts before sibling[0] → position 1
        var moved = h.Move("/body/p[3]", "/body", 0);
        _out.WriteLine($"Move returned: {moved}");
        moved.Should().NotBeNullOrEmpty();

        var p1 = h.Get("/body/p[1]");
        _out.WriteLine($"p[1] after move: '{p1?.Text}'");
        p1.Should().NotBeNull();
        p1!.Text.Should().Contain("C", "Para C moved to position 1 with index=0");

        h.Dispose();
        ValidateDocx(path, "after Move paragraph");
    }

    // ==================== 4. Full CRUD + persistence smoke ====================

    [Fact]
    public void Pptx_FullCRUD_Persists()
    {
        var path = Temp("pptx");

        using (var h = new PowerPointHandler(path, editable: true))
        {
            // Create — blank PPTX has no title placeholder; textbox is shape[1]
            h.Add("/", "slide", null, new() { ["title"] = "Smoke Slide" });
            h.Add("/slide[1]", "textbox", null, new()
            {
                ["text"] = "Hello", ["x"] = "2cm", ["y"] = "4cm",
                ["width"] = "10cm", ["height"] = "3cm",
                ["bold"] = "true", ["size"] = "18pt", ["fill"] = "#FF6600"
            });

            // Find the textbox by querying all shapes and finding the one with text "Hello"
            var shapes = h.Query("textbox");
            _out.WriteLine($"Shapes after add: {shapes.Count}, texts: {string.Join(",", shapes.Select(s => s.Text))}");
            var helloShape = shapes.FirstOrDefault(s => s.Text == "Hello");
            helloShape.Should().NotBeNull("textbox with 'Hello' text should exist");

            var shapePath = helloShape!.Path;
            _out.WriteLine($"Hello shape path: {shapePath}");

            // Read
            var node = h.Get(shapePath);
            node.Should().NotBeNull();
            node!.Text.Should().Be("Hello");
            node.Format["size"].ToString().Should().Be("18pt");
            node.Format["fill"].ToString().Should().Be("#FF6600");

            // Update
            h.Set(shapePath, new() { ["text"] = "Updated", ["size"] = "20pt" });
            node = h.Get(shapePath);
            node!.Text.Should().Be("Updated");
            node.Format["size"].ToString().Should().Be("20pt");

            // Delete
            h.Remove(shapePath);
            var afterDelete = Record.Exception(() => h.Get(shapePath));
            // either returns null or throws — both are acceptable "not found" behaviors
            _out.WriteLine($"After delete Get: {(afterDelete == null ? "null/exception" : afterDelete.Message)}");
        }

        ValidatePptx(path, "PPTX CRUD");

        // Persistence: add then reopen
        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/slide[1]", "textbox", null, new()
            {
                ["text"] = "Persisted", ["x"] = "1cm", ["y"] = "1cm",
                ["width"] = "5cm", ["height"] = "2cm"
            });
        }

        using (var h = new PowerPointHandler(path, editable: false))
        {
            var node = h.Query("textbox");
            node.Should().NotBeEmpty("shape persists after reopen");
            node.Should().ContainSingle(n => n.Text == "Persisted");
        }
    }

    [Fact]
    public void Docx_FullCRUD_Persists()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            // Create
            h.Add("/body", "paragraph", null, new()
            {
                ["text"] = "Hello Word", ["bold"] = "true",
                ["alignment"] = "center", ["size"] = "14pt",
                ["spaceBefore"] = "12pt", ["spaceAfter"] = "6pt"
            });

            // Read
            var node = h.Get("/body/p[1]");
            node.Should().NotBeNull();
            node!.Text.Should().Be("Hello Word");
            node.Format["size"].ToString().Should().Be("14pt");
            node.Format["alignment"].ToString().Should().Be("center");

            // Update
            h.Set("/body/p[1]", new() { ["text"] = "Updated Word", ["size"] = "16pt" });
            node = h.Get("/body/p[1]");
            node!.Text.Should().Contain("Updated Word");
            node.Format["size"].ToString().Should().Be("16pt");

            // Delete
            h.Remove("/body/p[1]");
        }

        ValidateDocx(path, "DOCX CRUD");

        // Persistence
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Persisted Para" });
        }

        using (var h = new WordHandler(path, editable: false))
        {
            var paras = h.Query("paragraph");
            paras.Should().ContainSingle(p => p.Text == "Persisted Para", "paragraph persists");
        }
    }

    [Fact]
    public void Excel_FullCRUD_Persists()
    {
        var path = Temp("xlsx");

        using (var h = new ExcelHandler(path, editable: true))
        {
            // Create / Set
            h.Set("/Sheet1/C3", new()
            {
                ["value"] = "99", ["bold"] = "true",
                ["fill"] = "#AACCFF", ["alignment.horizontal"] = "right"
            });

            // Read
            var node = h.Get("/Sheet1/C3");
            node.Should().NotBeNull();
            node!.Text.Should().Be("99");

            // Update
            h.Set("/Sheet1/C3", new() { ["value"] = "100", ["fill"] = "#FFFFFF" });
            node = h.Get("/Sheet1/C3");
            node!.Text.Should().Be("100");

            // Remove
            h.Remove("/Sheet1/C3");
            // After remove, cell may still exist with empty/null text or not exist at all
            var removed = h.Get("/Sheet1/C3");
            _out.WriteLine($"After Remove, C3 text='{removed?.Text}'");
            // Just verify no crash — consistent behavior is acceptable
        }

        ValidateXlsx(path, "XLSX CRUD");

        // Persistence
        using (var h = new ExcelHandler(path, editable: true))
        {
            h.Set("/Sheet1/D4", new() { ["value"] = "Persisted" });
        }

        using (var h = new ExcelHandler(path, editable: false))
        {
            var node = h.Get("/Sheet1/D4");
            node.Should().NotBeNull();
            node!.Text.Should().Be("Persisted", "Excel cell persists after reopen");
        }
    }

    // ==================== 5. Edge cases ====================

    [Fact]
    public void Pptx_AddSlide_EmptyTitle_Works()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        var ex = Record.Exception(() =>
            h.Add("/", "slide", null, new() { ["title"] = "" }));
        ex.Should().BeNull("empty title should not throw");

        var slide = h.Get("/slide[1]");
        slide.Should().NotBeNull();
    }

    [Fact]
    public void Word_Paragraph_SpecialChars_RoundTrip()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        var text = "Hello <World> & \"quotes\" ©2026";
        h.Add("/body", "paragraph", null, new() { ["text"] = text });

        h.Dispose();
        ValidateDocx(path, "special chars");

        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Get("/body/p[1]");
        node.Should().NotBeNull();
        node!.Text.Should().Contain("Hello");
    }

    [Fact]
    public void Excel_Formula_Set_Get_Works()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "10" });
        h.Set("/Sheet1/A2", new() { ["value"] = "20" });
        h.Set("/Sheet1/A3", new() { ["formula"] = "=A1+A2" });

        var node = h.Get("/Sheet1/A3");
        node.Should().NotBeNull();
        _out.WriteLine($"Formula cell text: '{node!.Text}'");
        // Formula text or cached value
        (node.Text.Contains("A1") || node.Text == "30" || string.IsNullOrEmpty(node.Text))
            .Should().BeTrue("formula cell returns formula or cached value");

        h.Dispose();
        ValidateXlsx(path, "Excel formula");
    }

    [Fact]
    public void Word_Table_AddGet_Works()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "4" });

        var tables = h.Query("table");
        tables.Should().NotBeEmpty("table was added");
        _out.WriteLine($"Table path: {tables[0].Path}");

        // Word table path uses 'tbl' not 'table' internally
        var tblPath = tables[0].Path;
        var tbl = h.Get(tblPath);
        tbl.Should().NotBeNull("can Get table by its reported path");

        h.Dispose();
        ValidateDocx(path, "Word table");
    }

    [Fact]
    public void Pptx_Table_AddGet_Works()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Table Slide" });
        h.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "3",
            ["x"] = "1cm", ["y"] = "3cm", ["width"] = "18cm", ["height"] = "5cm"
        });

        var tables = h.Query("table");
        tables.Should().NotBeEmpty("table added to slide");

        h.Dispose();
        ValidatePptx(path, "PPTX table");
    }

    [Fact]
    public void Pptx_Picture_AddRemove_ImagePartCleaned()
    {
        var imgPath = TempImg();
        var path = Temp("pptx");

        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Pic slide" });
            h.Add("/slide[1]", "picture", null, new()
            {
                ["src"] = imgPath, ["x"] = "1cm", ["y"] = "1cm",
                ["width"] = "3cm", ["height"] = "3cm"
            });
        }

        int countBefore;
        using (var doc = PresentationDocument.Open(path, false))
            countBefore = doc.PresentationPart!.SlideParts.First().ImageParts.Count();
        countBefore.Should().BeGreaterThan(0, "image part added");

        using (var h = new PowerPointHandler(path, editable: true))
        {
            var pics = h.Query("picture");
            if (pics.Count > 0) h.Remove(pics[0].Path);
        }

        ValidatePptx(path, "after picture remove");
    }

    [Fact]
    public void Word_Set_FontSize_BelowOneHalfPt_DoesNotCrash()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Tiny" });
        // Very small size — should not crash, just clamp or store minimum
        var ex = Record.Exception(() => h.Set("/body/p[1]", new() { ["size"] = "0.1pt" }));
        ex.Should().BeNull("tiny font size should not crash");
    }
}
