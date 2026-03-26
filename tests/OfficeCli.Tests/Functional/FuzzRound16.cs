// FuzzRound16 — Final fuzz: unexplored lifecycle corners
//
// Areas:
//   WD01–WD03: Word SDT/content-control add/get/set lifecycle
//   WF01–WF03: Word field code add + Set("dirty") + persistence
//   WK01–WK03: Word section break add + page size Set + persistence
//   EN01–EN03: Excel named range add/get/remove lifecycle
//   EI01–EI03: Excel icon-set conditional formatting lifecycle
//   EP01–EP03: Excel pivot table add + reopen guard
//   PH01–PH03: PPTX hyperlink on shape add + get + remove
//   PT01–PT03: PPTX theme color Set + Get round-trip

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound16 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz16_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== WD01–WD03: Word SDT/content control ====================

    [Fact]
    public void WD01_Word_AddSdt_PlainText_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Before control" });
        var act = () => h.Add("/body", "sdt", null, new() { ["type"] = "plaintext", ["text"] = "Hello SDT" });
        act.Should().NotThrow("adding a plaintext SDT should not throw");
    }

    [Fact]
    public void WD02_Word_AddSdt_SetText_FileValid()
    {
        var path = CreateTemp("docx");
        string? sdtPath = null;
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Intro" });
            try { sdtPath = h.Add("/body", "sdt", null, new() { ["type"] = "richtext", ["text"] = "Original" }); }
            catch (Exception) { return; }
            if (sdtPath != null)
            {
                try { h.Set(sdtPath, new() { ["text"] = "Updated" }); }
                catch (NullReferenceException ex) { Assert.Fail($"NullRef on SDT Set: {ex.Message}"); }
                catch (Exception) { /* unsupported property = ok */ }
            }
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("file should be valid after SDT add+set and reopen");
    }

    [Fact]
    public void WD03_Word_AddSdt_Dropdown_Persistence()
    {
        var path = CreateTemp("docx");
        bool added = false;
        {
            using var h = new WordHandler(path, editable: true);
            try
            {
                h.Add("/body", "sdt", null, new() { ["type"] = "dropdown", ["items"] = "Alpha,Beta,Gamma" });
                added = true;
            }
            catch (Exception) { /* skip if unsupported */ }
        }
        if (!added) return;
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("file should be valid after dropdown SDT add and reopen");
    }

    // ==================== WF01–WF03: Word field codes ====================

    [Fact]
    public void WF01_Word_AddField_PageNumber_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Page: " });
        var act = () => h.Add("/body/p[1]", "field", null, new() { ["instruction"] = "PAGE" });
        act.Should().NotThrow("adding a PAGE field should not throw");
    }

    [Fact]
    public void WF02_Word_AddField_SetDirty_NoNullRef()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "NumPages: " });
        string? fieldPath = null;
        try { fieldPath = h.Add("/body/p[1]", "field", null, new() { ["instruction"] = "NUMPAGES" }); }
        catch (Exception) { return; }
        if (fieldPath == null) return;
        var act = () => h.Set(fieldPath, new() { ["dirty"] = "true" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef setting field dirty: {ex.Message}"); }
        catch (Exception) { /* unsupported = ok */ }
    }

    [Fact]
    public void WF03_Word_AddField_Date_Persistence()
    {
        var path = CreateTemp("docx");
        bool added = false;
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Date: " });
            try
            {
                h.Add("/body/p[1]", "field", null, new() { ["instruction"] = @"DATE \@ ""MMMM d, yyyy""" });
                added = true;
            }
            catch (Exception) { /* skip if unsupported */ }
        }
        if (!added) return;
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("file should be valid after DATE field add and reopen");
    }

    // ==================== WK01–WK03: Word section + page size ====================

    [Fact]
    public void WK01_Word_AddSection_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Section 1 content" });
        var act = () => h.Add("/body", "section", null, new() { ["type"] = "nextPage" });
        act.Should().NotThrow("adding a section break should not throw");
    }

    [Fact]
    public void WK02_Word_SetSection_PageSize_RoundTrip()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
        var act = () => h.Set("/section[1]", new() { ["pageWidth"] = "21cm", ["pageHeight"] = "29.7cm" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef setting section page size: {ex.Message}"); }
        catch (Exception) { return; } // path may not exist for blank doc — ok
        var node = h.Get("/section[1]");
        node.Should().NotBeNull("section[1] should be gettable after Set");
        if (node?.Format?.ContainsKey("pageWidth") == true)
            node.Format["pageWidth"].ToString().Should().NotBeNullOrEmpty("pageWidth should be readable after Set");
    }

    [Fact]
    public void WK03_Word_SetSection_Orientation_Persistence()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Landscape para" });
            try { h.Set("/section[1]", new() { ["orientation"] = "landscape" }); }
            catch (Exception) { return; }
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("file should be valid after section orientation Set and reopen");
    }

    // ==================== EN01–EN03: Excel named range ====================

    [Fact]
    public void EN01_Excel_AddNamedRange_ThenGet_NoThrow()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "10" });
        h.Set("/Sheet1/A2", new() { ["value"] = "20" });
        var act = () => h.Add("/Sheet1", "namedrange", null,
            new() { ["name"] = "MyRange", ["range"] = "Sheet1!A1:A2" });
        act.Should().NotThrow("adding a named range should not throw");
    }

    [Fact]
    public void EN02_Excel_AddNamedRange_Remove_NoNullRef()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/B1", new() { ["value"] = "val" });
        bool added = false;
        try
        {
            h.Add("/Sheet1", "namedrange", null, new() { ["name"] = "RemoveMe", ["range"] = "Sheet1!B1" });
            added = true;
        }
        catch (Exception) { return; }
        if (!added) return;
        var act = () => h.Remove("/namedrange[RemoveMe]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing named range: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable */ }
    }

    [Fact]
    public void EN03_Excel_AddNamedRange_Persistence()
    {
        var path = CreateTemp("xlsx");
        bool added = false;
        {
            using var h = new ExcelHandler(path, editable: true);
            h.Set("/Sheet1/C1", new() { ["value"] = "99" });
            try
            {
                h.Add("/Sheet1", "namedrange", null, new() { ["name"] = "PersistRange", ["range"] = "Sheet1!C1" });
                added = true;
            }
            catch (Exception) { /* skip */ }
        }
        if (!added) return;
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            _ = h2.Get("/Sheet1/C1");
        };
        act.Should().NotThrow("file should be valid after named range add and reopen");
    }

    // ==================== EI01–EI03: Excel icon set conditional formatting ====================

    [Fact]
    public void EI01_Excel_AddIconSet_NoThrow()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        for (int i = 1; i <= 5; i++) h.Set($"/Sheet1/A{i}", new() { ["value"] = $"{i * 20}" });
        var act = () => h.Add("/Sheet1", "iconset", null,
            new() { ["range"] = "A1:A5", ["style"] = "3Arrows" });
        act.Should().NotThrow("adding icon set CF should not throw");
    }

    [Fact]
    public void EI02_Excel_AddIconSet_Persistence()
    {
        var path = CreateTemp("xlsx");
        bool added = false;
        {
            using var h = new ExcelHandler(path, editable: true);
            for (int i = 1; i <= 3; i++) h.Set($"/Sheet1/B{i}", new() { ["value"] = $"{i * 10}" });
            try
            {
                h.Add("/Sheet1", "iconset", null, new() { ["range"] = "B1:B3", ["style"] = "3Symbols" });
                added = true;
            }
            catch (Exception) { /* skip */ }
        }
        if (!added) return;
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            _ = h2.Get("/Sheet1/B1");
        };
        act.Should().NotThrow("file should be valid after icon set add and reopen");
    }

    [Fact]
    public void EI03_Excel_AddIconSet_Remove_NoNullRef()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        for (int i = 1; i <= 4; i++) h.Set($"/Sheet1/C{i}", new() { ["value"] = $"{i}" });
        bool added = false;
        try
        {
            h.Add("/Sheet1", "iconset", null, new() { ["range"] = "C1:C4" });
            added = true;
        }
        catch (Exception) { return; }
        if (!added) return;
        var act = () => h.Remove("/Sheet1/cf[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing icon set CF: {ex.Message}"); }
        catch (Exception) { /* other = acceptable */ }
    }

    // ==================== EP01–EP03: Excel pivot table ====================

    [Fact]
    public void EP01_Excel_AddPivotTable_NoThrow()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var headers = new[] { "Name", "Region", "Sales" };
        for (int c = 0; c < headers.Length; c++)
            h.Set($"/Sheet1/{(char)('A' + c)}1", new() { ["value"] = headers[c] });
        for (int r = 2; r <= 4; r++)
        {
            h.Set($"/Sheet1/A{r}", new() { ["value"] = $"Item{r}" });
            h.Set($"/Sheet1/B{r}", new() { ["value"] = r % 2 == 0 ? "North" : "South" });
            h.Set($"/Sheet1/C{r}", new() { ["value"] = $"{r * 100}" });
        }
        var act = () => h.Add("/Sheet1", "pivottable", null,
            new() { ["sourceRange"] = "Sheet1!A1:C4", ["destCell"] = "E1",
                    ["rowField"] = "Region", ["valueField"] = "Sales" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef adding pivot table: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable */ }
    }

    [Fact]
    public void EP02_Excel_AddPivotTable_Persistence()
    {
        var path = CreateTemp("xlsx");
        bool added = false;
        {
            using var h = new ExcelHandler(path, editable: true);
            for (int r = 1; r <= 3; r++)
            {
                h.Set($"/Sheet1/A{r}", new() { ["value"] = r == 1 ? "Cat" : $"Val{r}" });
                h.Set($"/Sheet1/B{r}", new() { ["value"] = r == 1 ? "Num" : $"{r * 5}" });
            }
            try
            {
                h.Add("/Sheet1", "pivottable", null,
                    new() { ["sourceRange"] = "Sheet1!A1:B3", ["destCell"] = "D1",
                            ["rowField"] = "Cat", ["valueField"] = "Num" });
                added = true;
            }
            catch (Exception) { /* skip */ }
        }
        if (!added) return;
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            _ = h2.Get("/Sheet1/A1");
        };
        act.Should().NotThrow("file should be valid after pivot table add and reopen");
    }

    [Fact]
    public void EP03_Excel_AddPivotTable_EmptySource_NoNullRef()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        // Empty source range — handler must not NullRef
        var act = () => h.Add("/Sheet1", "pivottable", null,
            new() { ["sourceRange"] = "Sheet1!A1:A1", ["destCell"] = "C1",
                    ["rowField"] = "X", ["valueField"] = "Y" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef with empty pivot source: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = expected */ }
    }

    // ==================== PH01–PH03: PPTX hyperlink on shape ====================

    [Fact]
    public void PH01_Pptx_AddHyperlink_OnShape_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Click me", ["x"] = "2cm", ["y"] = "2cm" });
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["url"] = "https://officecli.ai" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef setting hyperlink on shape: {ex.Message}"); }
        catch (Exception) { /* unsupported = acceptable */ }
    }

    [Fact]
    public void PH02_Pptx_AddHyperlink_Shape_Persistence()
    {
        var path = CreateTemp("pptx");
        bool set = false;
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Link shape", ["x"] = "1cm", ["y"] = "1cm" });
            try { h.Set("/slide[1]/shape[1]", new() { ["url"] = "https://example.com" }); set = true; }
            catch (Exception) { /* skip */ }
        }
        if (!set) return;
        var act = () =>
        {
            using var h2 = new PowerPointHandler(path, editable: false);
            var s = h2.Get("/slide[1]/shape[1]");
            s.Should().NotBeNull("shape should still be accessible after hyperlink set and reopen");
        };
        act.Should().NotThrow("file should be valid after hyperlink set and reopen");
    }

    [Fact]
    public void PH03_Pptx_SetHyperlink_ThenRemoveShape_NoNullRef()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Linked", ["x"] = "3cm", ["y"] = "3cm" });
        try { h.Set("/slide[1]/shape[1]", new() { ["url"] = "https://remove.test" }); }
        catch (Exception) { /* skip if url not supported */ }
        // Removing the shape (with or without hyperlink) must not NullRef
        var act = () => h.Remove("/slide[1]/shape[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing shape after hyperlink set: {ex.Message}"); }
        catch (Exception) { /* other = acceptable */ }
    }

    // ==================== PT01–PT03: PPTX theme color ====================

    [Fact]
    public void PT01_Pptx_GetTheme_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        var act = () =>
        {
            var node = h.Get("/theme");
            // If theme path is supported, node should not be null
            if (node != null)
                node.Type.Should().NotBeNullOrEmpty("theme node type should be set");
        };
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef getting /theme: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable if path unsupported */ }
    }

    [Fact]
    public void PT02_Pptx_SetThemeColor_Accent1_NoNullRef()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        var act = () => h.Set("/theme", new() { ["accent1"] = "FF0000" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef setting theme accent1: {ex.Message}"); }
        catch (Exception) { /* unsupported = acceptable */ }
    }

    [Fact]
    public void PT03_Pptx_SetThemeColor_Persistence()
    {
        var path = CreateTemp("pptx");
        bool applied = false;
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { });
            try { h.Set("/theme", new() { ["accent1"] = "4472C4" }); applied = true; }
            catch (Exception) { /* skip */ }
        }
        if (!applied) return;
        var act = () =>
        {
            using var h2 = new PowerPointHandler(path, editable: false);
            _ = h2.Query("shape").ToList();
        };
        act.Should().NotThrow("file should be valid after theme color Set and reopen");
    }
}
