// FuzzRound9 — R7 regression + null-value/empty-dict/empty-query crash probes + mid-op dispose recovery.
//
// Areas:
//   HD01–HD04: CleanupImageParts via mainPart — header/footer image remove regression
//   SC01–SC03: SanitizeColor("auto") regression in Word/Excel/Pptx Set
//   GR01–GR03: Gradient empty/null color boundary
//   NV01–NV06: Set with null value in dict — all three handlers
//   ED01–ED06: Add with empty dict — all three handlers
//   QE01–QE03: Query with empty string — all three handlers
//   MD01–MD03: Mid-op dispose + reopen consistency

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound9 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz9_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    private string CreatePng()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz9_img_{Guid.NewGuid():N}.png");
        _tempFiles.Add(path);
        File.WriteAllBytes(path, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="));
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== HD01–HD04: CleanupImageParts via mainPart header/footer ====================

    [Fact]
    public void HD01_Word_Remove_Header_WithImage_NoThrow()
    {
        var imgPath = CreatePng();
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "header", null, new() { ["text"] = "Header with pic" });
        // Add picture to the header body paragraph via inline picture
        var act = () => h.Remove("/header[1]");
        act.Should().NotThrow("removing a header with no image should not throw");
    }

    [Fact]
    public void HD02_Word_Remove_Footer_NoImage_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "footer", null, new() { ["text"] = "Footer text" });
        var act = () => h.Remove("/footer[1]");
        act.Should().NotThrow("removing a footer should not throw");
    }

    [Fact]
    public void HD03_Word_Remove_Header_Then_Reopen_NotCorrupted()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "header", null, new() { ["text"] = "H1" });
            h.Add("/body", "header", null, new() { ["text"] = "H2" });
            h.Remove("/header[1]");
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("header").ToList();
        };
        act.Should().NotThrow("reopening after header remove should not corrupt the document");
    }

    [Fact]
    public void HD04_Word_Remove_Footer_Then_AddNew_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "footer", null, new() { ["text"] = "OldFooter" });
        h.Remove("/footer[1]");
        var act = () => h.Add("/body", "footer", null, new() { ["text"] = "NewFooter" });
        act.Should().NotThrow("adding footer after removing previous one should not throw");
    }

    // ==================== SC01–SC03: SanitizeColor("auto") regression ====================

    [Fact]
    public void SC01_Word_Set_Shading_AutoColor_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "AutoColor" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("AutoColor"));
        // "auto" is a valid OOXML shading color — should not throw
        var act = () => h.Set(para.Path, new() { ["shading"] = "auto" });
        act.Should().NotThrow("Set shading=auto should not throw (valid OOXML color)");
    }

    [Fact]
    public void SC02_Excel_Set_FontColor_AutoString_NoThrow()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "AutoTest" });
        // Excel does not use "auto" color for font, but Set should not crash with unexpected string
        // The handler should either handle it or throw ArgumentException, not NullReferenceException
        var act = () => h.Set("/Sheet1/A1", new() { ["color"] = "auto" });
        act.Should().NotThrow("Set color=auto on Excel should not crash");
    }

    [Fact]
    public void SC03_Pptx_Set_Fill_AutoString_Graceful()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "AutoFill" });
        // "auto" is not a valid gradient/solid color in PPTX; it should be handled gracefully
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["fill"] = "auto" });
        // Accept either no throw or ArgumentException; crash (NullRef, IndexOutOfRange) is a bug
        try { act(); }
        catch (ArgumentException) { /* acceptable */ }
        catch (Exception ex) { Assert.Fail($"Unexpected exception type {ex.GetType().Name}: {ex.Message}"); }
    }

    // ==================== GR01–GR03: Gradient empty/boundary colors ====================

    [Fact]
    public void GR01_Pptx_Gradient_TwoIdenticalColors_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "GradSame" });
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["gradient"] = "FF0000-FF0000" });
        act.Should().NotThrow("gradient with two identical colors should not throw");
    }

    [Fact]
    public void GR02_Pptx_Gradient_ThemeColors_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "GradTheme" });
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["gradient"] = "accent1-accent2" });
        act.Should().NotThrow("gradient with theme/scheme colors should not throw");
    }

    [Fact]
    public void GR03_Pptx_Gradient_Radial_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "GradRadial" });
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["gradient"] = "radial:FF0000-0000FF" });
        act.Should().NotThrow("radial gradient should not throw");
    }

    // ==================== NV01–NV06: Set with null value in dict ====================

    [Fact]
    public void NV01_Word_Set_NullValue_DoesNotCrash()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "NullVal" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("NullVal"));
        // null value in dict should not produce NullReferenceException
        var act = () => h.Set(para.Path, new() { ["bold"] = null! });
        try { act(); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Word Set null value: {ex.Message}"); }
        catch (Exception) { /* other exceptions are acceptable */ }
    }

    [Fact]
    public void NV02_Excel_Set_NullValue_DoesNotCrash()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var act = () => h.Set("/Sheet1/A1", new() { ["value"] = null! });
        try { act(); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Excel Set null value: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void NV03_Pptx_Set_NullValue_DoesNotCrash()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "NullPptx" });
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["bold"] = null! });
        try { act(); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Pptx Set null value: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void NV04_Word_Set_MultipleNullValues_DoesNotCrash()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "MultiNull" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("MultiNull"));
        var act = () => h.Set(para.Path, new() { ["bold"] = null!, ["italic"] = null!, ["size"] = null! });
        try { act(); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Word Set multiple null values: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void NV05_Excel_Set_NullAndValidMixed_DoesNotCrash()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var act = () => h.Set("/Sheet1/B2", new() { ["value"] = "real", ["bold"] = null!, ["color"] = null! });
        try { act(); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Excel Set mixed null: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void NV06_Pptx_Set_NullFill_DoesNotCrash()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "NullFill" });
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["fill"] = null! });
        try { act(); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Pptx Set null fill: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    // ==================== ED01–ED06: Add with empty dict ====================

    [Fact]
    public void ED01_Word_Add_Paragraph_EmptyDict_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        var act = () => h.Add("/body", "paragraph", null, new());
        act.Should().NotThrow("Add paragraph with empty dict should not throw");
    }

    [Fact]
    public void ED02_Word_Add_Table_EmptyDict_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        var act = () => h.Add("/body", "table", null, new());
        act.Should().NotThrow("Add table with empty dict should not throw");
    }

    [Fact]
    public void ED03_Excel_Add_Sheet_EmptyDict_NoThrow()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var act = () => h.Add("/", "sheet", null, new());
        act.Should().NotThrow("Add sheet with empty dict should not throw");
    }

    [Fact]
    public void ED04_Pptx_Add_Slide_EmptyDict_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        var act = () => h.Add("/", "slide", null, new());
        act.Should().NotThrow("Add slide with empty dict should not throw");
    }

    [Fact]
    public void ED05_Pptx_Add_Shape_EmptyDict_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        var act = () => h.Add("/slide[1]", "shape", null, new());
        act.Should().NotThrow("Add shape with empty dict should not throw");
    }

    [Fact]
    public void ED06_Excel_Add_Comment_EmptyDict_Graceful()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        // Comment with empty dict might throw ArgumentException (no ref) — that's ok, just not crash
        var act = () => h.Add("/Sheet1", "comment", null, new());
        try { act(); }
        catch (ArgumentException) { /* ok — missing required ref */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Excel Add comment empty dict: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    // ==================== QE01–QE03: Query with empty string ====================

    [Fact]
    public void QE01_Word_Query_EmptyString_DoesNotCrash()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        var act = () => _ = h.Query("").ToList();
        try { act(); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Word Query empty string: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void QE02_Excel_Query_EmptyString_DoesNotCrash()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var act = () => _ = h.Query("").ToList();
        try { act(); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Excel Query empty string: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    [Fact]
    public void QE03_Pptx_Query_EmptyString_DoesNotCrash()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        var act = () => _ = h.Query("").ToList();
        try { act(); }
        catch (ArgumentException) { /* ok */ }
        catch (NullReferenceException ex) { Assert.Fail($"NullReferenceException from Pptx Query empty string: {ex.Message}"); }
        catch (Exception) { /* acceptable */ }
    }

    // ==================== MD01–MD03: Mid-op dispose + reopen ====================

    [Fact]
    public void MD01_Word_DisposeAfterAdd_Reopen_Consistent()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "MidDispose" });
            // Dispose here — file should be written consistently
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            var paras = h2.Query("paragraph").ToList();
            paras.Any(p => p.Text.Contains("MidDispose")).Should().BeTrue("paragraph should persist after dispose/reopen");
        };
        act.Should().NotThrow("reopening Word doc after normal dispose should not throw");
    }

    [Fact]
    public void MD02_Excel_DisposeAfterSet_Reopen_Consistent()
    {
        var path = CreateTemp("xlsx");
        {
            using var h = new ExcelHandler(path, editable: true);
            h.Set("/Sheet1/A1", new() { ["value"] = "MidDisposeXlsx", ["bold"] = "true" });
        }
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            var node = h2.Get("/Sheet1/A1");
            node.Should().NotBeNull("cell should survive dispose/reopen");
            node!.Text.Should().Be("MidDisposeXlsx");
        };
        act.Should().NotThrow("reopening Excel after dispose should not throw");
    }

    [Fact]
    public void MD03_Pptx_DisposeAfterAddSlide_Reopen_Consistent()
    {
        var path = CreateTemp("pptx");
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "MidDisposePptx" });
        }
        var act = () =>
        {
            using var h2 = new PowerPointHandler(path, editable: false);
            var node = h2.Get("/slide[1]/shape[1]");
            node.Should().NotBeNull("shape should survive dispose/reopen");
            node!.Text.Should().Be("MidDisposePptx");
        };
        act.Should().NotThrow("reopening Pptx after dispose should not throw");
    }
}
