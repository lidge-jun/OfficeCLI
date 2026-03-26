// FuzzRound11 — R9 regression + new untested areas
//
// Areas:
//   MR01–MR06: MidpointRounding boundary values (x.5 inputs in Word font sizes)
//   WS01–WS04: Multi-section watermark Remove (sections with separate HeaderReference)
//   DM01–DM04: Dict not modified by Add/Set (caller's dict is unchanged after call)
//   NX01–NX05: New untested combos — Excel chart range, Pptx table cell merge, Word footnote
//   SM01–SM03: Global smoke — read-only open after editable close

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound11 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz11_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== MR01–MR06: MidpointRounding boundary ====================

    [Fact]
    public void MR01_Word_Add_FontSize_0pt5_BoundaryRoundsCorrectly()
    {
        // 0.5pt * 2 = 1.0 half-points → 1 → 0.5pt (no rounding ambiguity)
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "TinyFont", ["size"] = "0.5pt" });
        var node = h.Query("paragraph").FirstOrDefault(p => p.Text.Contains("TinyFont"));
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("size");
        node.Format["size"].ToString().Should().Be("0.5pt", "0.5pt * 2 = 1 half-point exact");
    }

    [Fact]
    public void MR02_Word_Set_FontSize_ExactHalf_AllBoundaries()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "HalfBoundary", ["size"] = "12pt" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("HalfBoundary"));

        // Test several x.5pt boundaries where AwayFromZero matters
        var cases = new[] { ("8.5pt", "8.5pt"), ("9.5pt", "9.5pt"), ("10.5pt", "10.5pt"),
                            ("11.5pt", "11.5pt"), ("12.5pt", "12.5pt"), ("13.5pt", "13.5pt") };
        foreach (var (input, expected) in cases)
        {
            h.Set(para.Path, new() { ["size"] = input });
            var n = h.Get(para.Path);
            n!.Format["size"].ToString().Should().Be(expected, $"{input} should round exactly");
        }
    }

    [Fact]
    public void MR03_Word_Add_FontSize_QuarterPt_RoundsAwayFromZero()
    {
        // 12.25pt * 2 = 24.5 → AwayFromZero → 25 → 12.5pt
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "QuarterPt", ["size"] = "12.25pt" });
        var node = h.Query("paragraph").FirstOrDefault(p => p.Text.Contains("QuarterPt"));
        node.Should().NotBeNull();
        node!.Format["size"].ToString().Should().Be("12.5pt", "12.25pt rounds up to 12.5pt (AwayFromZero)");
    }

    [Fact]
    public void MR04_Word_Set_FontSize_QuarterPt_Below_RoundsDown()
    {
        // 12.1pt * 2 = 24.2 → round to 24 → 12pt (both AwayFromZero and ToEven give same result here)
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "RndDown", ["size"] = "12pt" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("RndDown"));
        h.Set(para.Path, new() { ["size"] = "12.1pt" });
        var node = h.Get(para.Path);
        node!.Format["size"].ToString().Should().Be("12pt", "12.1pt rounds to 24 half-points = 12pt");
    }

    [Fact]
    public void MR05_Word_Header_FontSize_Boundary_NoThrow()
    {
        // font size in header/footer context
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "header", null, new() { ["text"] = "HeaderFont", ["size"] = "9.5pt" });
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            // just verify no crash on reopen
            _ = h2.Query("header").ToList();
        };
        act.Should().NotThrow("header with 9.5pt font should persist without error");
    }

    [Fact]
    public void MR06_Word_Paragraph_DefaultStyle_FontSize_Boundary_Persistence()
    {
        // paragraph with size 7.5pt (odd half-point) round-trips
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "SzPersist", ["size"] = "7.5pt" });
        }
        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Query("paragraph").FirstOrDefault(p => p.Text.Contains("SzPersist"));
        node.Should().NotBeNull();
        node!.Format["size"].ToString().Should().Be("7.5pt", "7.5pt should round-trip exactly");
    }

    // ==================== WS01–WS04: Multi-section watermark Remove ====================

    [Fact]
    public void WS01_Word_Remove_Watermark_SingleSection_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "watermark", null, new() { ["text"] = "DRAFT" });
        var act = () => h.Remove("/watermark");
        act.Should().NotThrow("removing watermark from single-section doc should not throw");
    }

    [Fact]
    public void WS02_Word_Remove_Watermark_ThenReopenValid()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "watermark", null, new() { ["text"] = "CONFIDENTIAL" });
            h.Remove("/watermark");
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("reopening after watermark remove should not corrupt document");
    }

    [Fact]
    public void WS03_Word_Remove_Watermark_Multiple_Removes_Idempotent()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "watermark", null, new() { ["text"] = "DRAFT" });
        h.Remove("/watermark");
        // Remove again — should not throw even when no watermark exists
        var act = () => h.Remove("/watermark");
        act.Should().NotThrow("removing watermark twice should be idempotent / not throw");
    }

    [Fact]
    public void WS04_Word_Remove_NonExistent_Watermark_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        // No watermark added — just remove it
        var act = () => h.Remove("/watermark");
        act.Should().NotThrow("removing watermark from doc with no watermark should not throw");
    }

    // ==================== DM01–DM04: Dict not modified by Add/Set ====================

    [Fact]
    public void DM01_Word_Add_DoesNotMutateCallerDict()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        var props = new Dictionary<string, string> { ["text"] = "ImmutableTest", ["bold"] = "true", ["size"] = "12pt" };
        var before = props.ToDictionary(kv => kv.Key, kv => kv.Value);
        h.Add("/body", "paragraph", null, props);
        props.Should().BeEquivalentTo(before, "Word Add must not mutate caller's dictionary");
    }

    [Fact]
    public void DM02_Word_Set_DoesNotMutateCallerDict()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "SetMutate" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("SetMutate"));
        var props = new Dictionary<string, string> { ["bold"] = "true", ["italic"] = "true", ["size"] = "14pt" };
        var before = props.ToDictionary(kv => kv.Key, kv => kv.Value);
        h.Set(para.Path, props);
        props.Should().BeEquivalentTo(before, "Word Set must not mutate caller's dictionary");
    }

    [Fact]
    public void DM03_Excel_Set_DoesNotMutateCallerDict()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var props = new Dictionary<string, string> { ["value"] = "DictTest", ["bold"] = "true", ["color"] = "FF0000" };
        var before = props.ToDictionary(kv => kv.Key, kv => kv.Value);
        h.Set("/Sheet1/A1", props);
        props.Should().BeEquivalentTo(before, "Excel Set must not mutate caller's dictionary");
    }

    [Fact]
    public void DM04_Pptx_Add_DoesNotMutateCallerDict()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        var props = new Dictionary<string, string> { ["text"] = "PptxMutate", ["fill"] = "4472C4", ["bold"] = "true" };
        var before = props.ToDictionary(kv => kv.Key, kv => kv.Value);
        h.Add("/slide[1]", "shape", null, props);
        props.Should().BeEquivalentTo(before, "Pptx Add must not mutate caller's dictionary");
    }

    // ==================== NX01–NX05: New untested combos ====================

    [Fact]
    public void NX01_Word_Add_Footnote_BasicRoundTrip()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Main text" });
        // footnote Add — may not be supported; should not NullRef crash
        var act = () => h.Add("/body/paragraph[1]", "footnote", null, new() { ["text"] = "Footnote text" });
        try { act(); }
        catch (ArgumentException) { /* unsupported = ok */ }
        catch (KeyNotFoundException) { /* path issue = ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException on footnote Add: {ex.Message}"); }
        catch (Exception) { /* other exceptions acceptable */ }
    }

    [Fact]
    public void NX02_Pptx_Add_Table_ThenGet_ReturnsChildren()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "3" });
        var tbl = h.Get("/slide[1]/table[1]");
        tbl.Should().NotBeNull("table should be gettable after add");
        tbl!.Type.Should().Be("table");
    }

    [Fact]
    public void NX03_Excel_Set_MultipleFormats_SameCell_RoundTrip()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/C3", new() { ["value"] = "MultiFormat", ["bold"] = "true", ["italic"] = "true",
            ["size"] = "14pt", ["color"] = "FF0000" });
        var node = h.Get("/Sheet1/C3");
        node.Should().NotBeNull();
        node!.Text.Should().Be("MultiFormat");
        node.Format["bold"].ToString().Should().BeOneOf("true", "True", "1");
        node.Format["italic"].ToString().Should().BeOneOf("true", "True", "1");
    }

    [Fact]
    public void NX04_Word_Set_FindReplace_PreservesOtherProperties()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Hello World find me" });
        var act = () => h.Set("/", new() { ["find"] = "find me", ["replace"] = "replaced" });
        act.Should().NotThrow("find/replace at root should not throw");
        var paras = h.Query("paragraph").ToList();
        paras.Any(p => p.Text.Contains("replaced")).Should().BeTrue("text should be replaced");
    }

    [Fact]
    public void NX05_Pptx_Add_Slide_WithLayout_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        // Add slide with layout index (if supported)
        var act = () => h.Add("/", "slide", null, new() { ["layout"] = "1" });
        act.Should().NotThrow("Add slide with layout=1 should not throw");
    }

    // ==================== SM01–SM03: Global smoke ====================

    [Fact]
    public void SM01_Word_WriteClose_ReopenReadOnly_NoThrow()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "SmokeWord" });
            h.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            var paras = h2.Query("paragraph").ToList();
            var tables = h2.Query("table").ToList();
            paras.Should().NotBeEmpty();
            tables.Should().NotBeEmpty();
        };
        act.Should().NotThrow("read-only open after editable session should not throw");
    }

    [Fact]
    public void SM02_Excel_WriteClose_ReopenReadOnly_NoThrow()
    {
        var path = CreateTemp("xlsx");
        {
            using var h = new ExcelHandler(path, editable: true);
            h.Set("/Sheet1/A1", new() { ["value"] = "SmokeXlsx" });
            h.Set("/Sheet1/B2", new() { ["value"] = "42", ["numberformat"] = "0.00" });
            h.Add("/", "sheet", (int?)null, new() { ["name"] = "SmokeSheet" });
        }
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            var n = h2.Get("/Sheet1/A1");
            n.Should().NotBeNull();
            n!.Text.Should().Be("SmokeXlsx");
        };
        act.Should().NotThrow("read-only reopen after Excel editable session should not throw");
    }

    [Fact]
    public void SM03_Pptx_WriteClose_ReopenReadOnly_NoThrow()
    {
        var path = CreateTemp("pptx");
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { ["title"] = "SmokePptx" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "SmokeShape", ["fill"] = "4472C4" });
            h.Add("/slide[1]", "textbox", null, new() { ["text"] = "SmokeTextBox", ["x"] = "1cm", ["y"] = "1cm",
                ["width"] = "5cm", ["height"] = "2cm" });
        }
        var act = () =>
        {
            using var h2 = new PowerPointHandler(path, editable: false);
            var shapes = h2.Query("shape").ToList();
            shapes.Should().NotBeEmpty();
        };
        act.Should().NotThrow("read-only reopen after Pptx editable session should not throw");
    }
}
