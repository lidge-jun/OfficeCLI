// FuzzRound10 — R8 regression + new attack surfaces (Unicode surrogates, long keys, null paths, dual-open).
//
// Areas:
//   FS01–FS03: font size rounding boundaries (10.25pt, 10.75pt, 0.5pt) — R8 regression
//   NV01–NV03: null value Set confirmed no crash across handlers — R8 regression
//   UN01–UN04: Unicode surrogate pairs / astral-plane text in all handlers
//   LK01–LK03: extremely long property keys (1000+ chars)
//   NP01–NP03: null/empty path for Get/Set/Remove — graceful error not crash
//   DO01–DO03: two handlers open on same file simultaneously

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound10 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz10_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== FS01–FS03: Font size rounding boundaries (R8 regression) ====================

    [Fact]
    public void FS01_Word_FontSize_10_25pt_RoundTrip()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Size1025", ["size"] = "10.25pt" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("Size1025"));
        var node = h.Get(para.Path);
        node.Should().NotBeNull();
        // Should not crash; readback should be a valid pt string
        var size = node!.Format.ContainsKey("size") ? node.Format["size"]?.ToString() : null;
        size.Should().NotBeNullOrEmpty("font size 10.25pt should round-trip to a non-empty pt value");
        size.Should().EndWith("pt", "font size should be returned with pt suffix");
    }

    [Fact]
    public void FS02_Word_FontSize_10_75pt_RoundTrip()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Size1075", ["size"] = "10.75pt" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("Size1075"));
        var node = h.Get(para.Path);
        node.Should().NotBeNull();
        var size = node!.Format.ContainsKey("size") ? node.Format["size"]?.ToString() : null;
        size.Should().NotBeNullOrEmpty("font size 10.75pt should round-trip to a non-empty pt value");
        size.Should().EndWith("pt");
    }

    [Fact]
    public void FS03_Word_FontSize_0_5pt_Boundary_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        // 0.5pt is near-zero boundary; handler may clamp or store as-is, but must not crash
        var act = () => h.Add("/body", "paragraph", null, new() { ["text"] = "Size05", ["size"] = "0.5pt" });
        act.Should().NotThrow("adding paragraph with 0.5pt font size should not throw");
    }

    // ==================== NV01–NV03: null value Set — confirmed no NullReferenceException (R8 regression) ====================

    [Fact]
    public void NV01_Word_Set_NullBold_NoNullReferenceException()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "NullBold" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("NullBold"));
        try { h.Set(para.Path, new() { ["bold"] = null! }); }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException: {ex.Message}"); }
        catch (Exception) { /* ArgumentException etc. acceptable */ }
    }

    [Fact]
    public void NV02_Excel_Set_NullValue_NoNullReferenceException()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        try { h.Set("/Sheet1/A1", new() { ["value"] = null! }); }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void NV03_Pptx_Set_NullText_NoNullReferenceException()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "NullText" });
        try { h.Set("/slide[1]/shape[1]", new() { ["text"] = null! }); }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    // ==================== UN01–UN04: Unicode surrogate pairs / astral-plane text ====================

    [Fact]
    public void UN01_Word_Paragraph_SurrogatePair_Text_NoThrow()
    {
        // 𝕳𝖊𝖑𝖑𝖔 — Mathematical Fraktur letters (U+1D573 etc.) stored as surrogate pairs in UTF-16
        var astral = "𝕳𝖊𝖑𝖑𝖔 \U0001F600 \U0001F4A9";
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        var act = () => h.Add("/body", "paragraph", null, new() { ["text"] = astral });
        act.Should().NotThrow("paragraph with surrogate-pair Unicode text should not throw");
    }

    [Fact]
    public void UN02_Word_Paragraph_SurrogatePair_RoundTrip()
    {
        var astral = "𝕳𝖊𝖑𝖑𝖔";
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = astral });
        }
        using var h2 = new WordHandler(path, editable: false);
        var paras = h2.Query("paragraph").ToList();
        paras.Any(p => p.Text != null && p.Text.Contains("𝕳")).Should().BeTrue(
            "surrogate-pair text should survive write/read round trip");
    }

    [Fact]
    public void UN03_Excel_Cell_SurrogatePair_NoThrow()
    {
        var astral = "𝕳𝖊𝖑𝖑𝖔 \U0001F4CA";
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var act = () => h.Set("/Sheet1/A1", new() { ["value"] = astral });
        act.Should().NotThrow("Excel cell with surrogate-pair Unicode should not throw");
    }

    [Fact]
    public void UN04_Pptx_Shape_SurrogatePair_NoThrow()
    {
        var astral = "𝕳𝖊𝖑𝖑𝖔 \U0001F600";
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        var act = () => h.Add("/slide[1]", "shape", null, new() { ["text"] = astral });
        act.Should().NotThrow("PPTX shape with surrogate-pair Unicode should not throw");
    }

    // ==================== LK01–LK03: Extremely long property keys (1000+ chars) ====================

    [Fact]
    public void LK01_Word_Set_LongKey_NoThrow()
    {
        var longKey = new string('x', 1000);
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "LongKey" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("LongKey"));
        // Unknown key — should silently skip or throw ArgumentException, not crash
        try { h.Set(para.Path, new() { [longKey] = "value" }); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from long key: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void LK02_Excel_Set_LongKey_NoThrow()
    {
        var longKey = new string('k', 1200);
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        try { h.Set("/Sheet1/A1", new() { [longKey] = "value" }); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from long key: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void LK03_Pptx_Set_LongKey_NoThrow()
    {
        var longKey = new string('p', 999);
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "LongKeyPptx" });
        try { h.Set("/slide[1]/shape[1]", new() { [longKey] = "value" }); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from long key: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    // ==================== NP01–NP03: null/empty path for Get/Set/Remove ====================

    [Fact]
    public void NP01_Word_Get_NullPath_GracefulError()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: false);
        try { _ = h.Get(null!); }
        catch (ArgumentNullException) { /* expected */ }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from null path Get: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void NP02_Excel_Set_EmptyPath_GracefulError()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        try { h.Set("", new() { ["value"] = "test" }); }
        catch (ArgumentException) { /* expected */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from empty path Set: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void NP03_Pptx_Remove_NullPath_GracefulError()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        try { h.Remove(null!); }
        catch (ArgumentNullException) { /* expected */ }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from null path Remove: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    // ==================== DO01–DO03: Two handlers open on same file simultaneously ====================

    [Fact]
    public void DO01_Word_TwoReadOnlyHandlers_SameFile_BothWork()
    {
        var path = CreateTemp("docx");
        // Pre-populate
        {
            using var hw = new WordHandler(path, editable: true);
            hw.Add("/body", "paragraph", null, new() { ["text"] = "DualOpen" });
        }
        // Open two read-only handlers simultaneously
        using var h1 = new WordHandler(path, editable: false);
        using var h2 = new WordHandler(path, editable: false);
        var act = () =>
        {
            var p1 = h1.Query("paragraph").ToList();
            var p2 = h2.Query("paragraph").ToList();
            p1.Should().NotBeEmpty("first handler should see paragraphs");
            p2.Should().NotBeEmpty("second handler should see paragraphs");
        };
        act.Should().NotThrow("two read-only handlers on same file should both work");
    }

    [Fact]
    public void DO02_Excel_TwoReadOnlyHandlers_SameFile_BothWork()
    {
        var path = CreateTemp("xlsx");
        {
            using var hw = new ExcelHandler(path, editable: true);
            hw.Set("/Sheet1/A1", new() { ["value"] = "DualXlsx" });
        }
        using var h1 = new ExcelHandler(path, editable: false);
        using var h2 = new ExcelHandler(path, editable: false);
        var act = () =>
        {
            var n1 = h1.Get("/Sheet1/A1");
            var n2 = h2.Get("/Sheet1/A1");
            n1.Should().NotBeNull();
            n2.Should().NotBeNull();
        };
        act.Should().NotThrow("two read-only Excel handlers on same file should both work");
    }

    [Fact]
    public void DO03_Pptx_TwoReadOnlyHandlers_SameFile_BothWork()
    {
        var path = CreateTemp("pptx");
        {
            using var hw = new PowerPointHandler(path, editable: true);
            hw.Add("/", "slide", null, new() { });
            hw.Add("/slide[1]", "shape", null, new() { ["text"] = "DualPptx" });
        }
        using var h1 = new PowerPointHandler(path, editable: false);
        using var h2 = new PowerPointHandler(path, editable: false);
        var act = () =>
        {
            var n1 = h1.Get("/slide[1]/shape[1]");
            var n2 = h2.Get("/slide[1]/shape[1]");
            n1.Should().NotBeNull();
            n2.Should().NotBeNull();
        };
        act.Should().NotThrow("two read-only PPTX handlers on same file should both work");
    }
}
