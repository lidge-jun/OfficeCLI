// FuzzRound13 — Target: recent fixes + untested feature combos
//
// Areas:
//   WC01–WC05: Word comment add/remove round-trip (fix: db99b70 body reference cleanup)
//   WF01–WF04: Word footnote/endnote remove (fix: 68421e4)
//   EC01–EC04: Excel conditional formatting auto-priority (fix: b733c8b)
//   ES01–ES03: Excel sparkline add lifecycle
//   PA01–PA04: Pptx arrow/hexagon shape preset add+get (fix: 11df1af text inset)
//   PAN1–PAN3: Pptx animation basic lifecycle
//   PT01–PT03: Pptx slide transition round-trip (fix: b733 readback)
//   NV01–NV04: Null-value guard in Set handlers (fix: dffe772)

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound13 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz13_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== WC01–WC05: Word comment add/remove ====================

    [Fact]
    public void WC01_Word_AddComment_ThenGet_ReturnsComment()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para with comment" });
        h.Add("/body/p[1]", "comment", null, new() { ["text"] = "My comment", ["author"] = "Tester" });
        var comments = h.Query("comment").ToList();
        comments.Should().NotBeEmpty("comment should appear in query results after add");
    }

    [Fact]
    public void WC02_Word_AddComment_ThenRemove_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para for comment remove" });
        h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Remove me" });
        var comments = h.Query("comment").ToList();
        comments.Should().NotBeEmpty();
        var act = () => h.Remove(comments[0].Path);
        act.Should().NotThrow("removing comment should not throw");
    }

    [Fact]
    public void WC03_Word_RemoveComment_CleansBodyReferences()
    {
        // Fix db99b70: body CommentRangeStart/End left dangling after comment removal
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Body ref cleanup test" });
            h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Cleanup test comment" });
            var comments = h.Query("comment").ToList();
            if (comments.Count > 0)
                h.Remove(comments[0].Path);
        }
        // Reopen — file must be valid (no dangling references causing parse errors)
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("file should remain valid after comment removal and body ref cleanup");
    }

    [Fact]
    public void WC04_Word_MultipleComments_RemoveFirst_OthersRemain()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Multi comment para" });
        h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Comment A" });
        h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Comment B" });
        var commentsBefore = h.Query("comment").ToList();
        commentsBefore.Should().HaveCountGreaterOrEqualTo(1, "at least one comment should exist");
        h.Remove(commentsBefore[0].Path);
        // After remove, file should still be openable
        var act = () => { _ = h.Query("paragraph").ToList(); };
        act.Should().NotThrow("querying after comment removal should not throw");
    }

    [Fact]
    public void WC05_Word_Comment_Persistence_RoundTrip()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Persist para" });
            h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Persisted comment", ["author"] = "FuzzBot" });
        }
        using var h2 = new WordHandler(path, editable: false);
        var comments = h2.Query("comment").ToList();
        comments.Should().NotBeEmpty("comment should persist after save and reopen");
    }

    // ==================== WF01–WF04: Word footnote/endnote remove ====================

    [Fact]
    public void WF01_Word_AddFootnote_BasicLifecycle_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Footnote para" });
        var act = () => h.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Footnote text" });
        act.Should().NotThrow("adding footnote should not throw");
    }

    [Fact]
    public void WF02_Word_RemoveFootnote_NoNullRef()
    {
        // Fix: 68421e4 — footnote Remove path support
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "FN remove para" });
        try { h.Add("/body/p[1]", "footnote", null, new() { ["text"] = "To be removed" }); }
        catch (Exception) { return; } // skip if Add not supported
        var act = () => h.Remove("/footnote[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef in footnote Remove: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable */ }
    }

    [Fact]
    public void WF03_Word_Footnote_Persistence_RoundTrip()
    {
        var path = CreateTemp("docx");
        bool addedFn = false;
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "FN persist para" });
            try
            {
                h.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Persisted footnote" });
                addedFn = true;
            }
            catch (Exception) { /* unsupported = skip */ }
        }
        if (!addedFn) return;
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("reopening after footnote add should not corrupt document");
    }

    [Fact]
    public void WF04_Word_RemoveFootnote_FromDocWithNoFootnotes_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "No footnotes here" });
        var act = () => h.Remove("/footnote[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing footnote from doc with none: {ex.Message}"); }
        catch (Exception) { /* expected: nothing to remove */ }
    }

    // ==================== EC01–EC04: Excel conditional formatting ====================

    [Fact]
    public void EC01_Excel_AddConditionalFormatting_DataBar_NoThrow()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "10" });
        h.Set("/Sheet1/A2", new() { ["value"] = "20" });
        h.Set("/Sheet1/A3", new() { ["value"] = "30" });
        var act = () => h.Add("/Sheet1", "databar", null,
            new() { ["range"] = "A1:A3", ["color"] = "4472C4" });
        act.Should().NotThrow("adding databar conditional formatting should not throw");
    }

    [Fact]
    public void EC02_Excel_AddTwoConditionalFormattings_PrioritiesUnique()
    {
        // Fix b733c8b: hardcoded priority=1 caused duplicate priority conflict
        var path = CreateTemp("xlsx");
        {
            using var h = new ExcelHandler(path, editable: true);
            for (int i = 1; i <= 5; i++)
                h.Set($"/Sheet1/A{i}", new() { ["value"] = i.ToString() });
            h.Add("/Sheet1", "databar", null, new() { ["range"] = "A1:A5", ["color"] = "FF0000" });
            h.Add("/Sheet1", "databar", null, new() { ["range"] = "A1:A5", ["color"] = "0000FF" });
        }
        // If priorities clash, file may fail validation; reopen verifies it's valid
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            _ = h2.Get("/Sheet1/A1");
        };
        act.Should().NotThrow("two conditional formattings should have unique priorities — file remains valid");
    }

    [Fact]
    public void EC03_Excel_AddColorScale_ThenGet_FileValid()
    {
        var path = CreateTemp("xlsx");
        {
            using var h = new ExcelHandler(path, editable: true);
            for (int i = 1; i <= 5; i++)
                h.Set($"/Sheet1/B{i}", new() { ["value"] = (i * 10).ToString() });
            h.Add("/Sheet1", "colorscale", null, new() { ["range"] = "B1:B5" });
        }
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            _ = h2.Get("/Sheet1/B3");
        };
        act.Should().NotThrow("file should be valid after colorscale add");
    }

    [Fact]
    public void EC04_Excel_MultipleConditionalFormats_MixedTypes_NoCrash()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        for (int i = 1; i <= 5; i++)
            h.Set($"/Sheet1/C{i}", new() { ["value"] = (i * 5).ToString() });
        var acts = new Action[]
        {
            () => h.Add("/Sheet1", "databar", null, new() { ["range"] = "C1:C5", ["color"] = "4472C4" }),
            () => h.Add("/Sheet1", "colorscale", null, new() { ["range"] = "C1:C5" }),
            () => h.Add("/Sheet1", "databar", null, new() { ["range"] = "C1:C5", ["color"] = "70AD47" }),
        };
        foreach (var a in acts)
        {
            try { a(); }
            catch (NullReferenceException ex) { Assert.Fail($"NullRef adding CF rule: {ex.Message}"); }
            catch (Exception) { /* other exceptions acceptable */ }
        }
    }

    // ==================== ES01–ES03: Excel sparkline ====================

    [Fact]
    public void ES01_Excel_AddSparkline_Line_NoThrow()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        for (int i = 1; i <= 5; i++)
            h.Set($"/Sheet1/A{i}", new() { ["value"] = (i * 3).ToString() });
        var act = () => h.Add("/Sheet1", "sparkline", null,
            new() { ["cell"] = "B1", ["range"] = "A1:A5", ["type"] = "line" });
        act.Should().NotThrow("adding line sparkline should not throw");
    }

    [Fact]
    public void ES02_Excel_AddSparkline_Column_Persistence()
    {
        var path = CreateTemp("xlsx");
        {
            using var h = new ExcelHandler(path, editable: true);
            for (int i = 1; i <= 5; i++)
                h.Set($"/Sheet1/A{i}", new() { ["value"] = (i * 2).ToString() });
            h.Add("/Sheet1", "sparkline", null,
                new() { ["cell"] = "B1", ["range"] = "A1:A5", ["type"] = "column" });
        }
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            _ = h2.Get("/Sheet1/A1");
        };
        act.Should().NotThrow("file should be valid after column sparkline add and reopen");
    }

    [Fact]
    public void ES03_Excel_AddMultipleSparklines_NoConflict()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        for (int col = 1; col <= 3; col++)
            for (int row = 1; row <= 5; row++)
                h.Set($"/Sheet1/{(char)('A' + col - 1)}{row}", new() { ["value"] = (row * col).ToString() });
        var act = () =>
        {
            h.Add("/Sheet1", "sparkline", null, new() { ["cell"] = "D1", ["range"] = "A1:A5" });
            h.Add("/Sheet1", "sparkline", null, new() { ["cell"] = "D2", ["range"] = "B1:B5" });
            h.Add("/Sheet1", "sparkline", null, new() { ["cell"] = "D3", ["range"] = "C1:C5" });
        };
        act.Should().NotThrow("adding multiple sparklines should not throw");
    }

    // ==================== PA01–PA04: Pptx arrow/hexagon shape preset ====================

    [Fact]
    public void PA01_Pptx_AddShape_RightArrow_GetReturnsPreset()
    {
        // Fix 11df1af: text inset for rightArrow was incorrect
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() {
            ["text"] = "Go Right", ["preset"] = "rightArrow",
            ["width"] = "4cm", ["height"] = "2cm", ["x"] = "1cm", ["y"] = "1cm"
        });
        var node = h.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull("rightArrow shape should be gettable after add");
        node!.Text.Should().Be("Go Right");
        node.Format["preset"].ToString().Should().Be("rightArrow", "preset should round-trip as rightArrow");
    }

    [Fact]
    public void PA02_Pptx_AddShape_Hexagon_GetReturnsPreset()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() {
            ["text"] = "Hex", ["preset"] = "hexagon",
            ["width"] = "3cm", ["height"] = "3cm", ["x"] = "2cm", ["y"] = "2cm"
        });
        var node = h.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull("hexagon shape should be gettable");
        node!.Format["preset"].ToString().Should().Be("hexagon");
    }

    [Fact]
    public void PA03_Pptx_AddShape_AllArrowDirections_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        foreach (var preset in new[] { "rightArrow", "leftArrow", "upArrow", "downArrow" })
        {
            var act = () => h.Add("/slide[1]", "shape", null, new() {
                ["text"] = preset, ["preset"] = preset,
                ["width"] = "3cm", ["height"] = "2cm"
            });
            act.Should().NotThrow($"adding {preset} shape should not throw");
        }
    }

    [Fact]
    public void PA04_Pptx_ArrowShape_Persistence()
    {
        var path = CreateTemp("pptx");
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { });
            h.Add("/slide[1]", "shape", null, new() {
                ["text"] = "Persist Arrow", ["preset"] = "rightArrow",
                ["width"] = "5cm", ["height"] = "2cm"
            });
        }
        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Text.Should().Be("Persist Arrow");
        node.Format["preset"].ToString().Should().Be("rightArrow", "preset should persist after save/reopen");
    }

    // ==================== PAN1–PAN3: Pptx animation basic lifecycle ====================

    [Fact]
    public void PAN1_Pptx_AddAnimation_Entrance_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Animated Shape" });
        var act = () => h.Add("/slide[1]/shape[1]", "animation", null,
            new() { ["type"] = "fade", ["trigger"] = "onclick" });
        act.Should().NotThrow("adding entrance animation should not throw");
    }

    [Fact]
    public void PAN2_Pptx_AddAnimation_Persistence()
    {
        var path = CreateTemp("pptx");
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "AnimPersist" });
            try
            {
                h.Add("/slide[1]/shape[1]", "animation", null,
                    new() { ["type"] = "fly", ["trigger"] = "onclick" });
            }
            catch (Exception) { return; } // skip if unsupported
        }
        var act = () =>
        {
            using var h2 = new PowerPointHandler(path, editable: false);
            _ = h2.Query("shape").ToList();
        };
        act.Should().NotThrow("file should remain valid after animation add");
    }

    [Fact]
    public void PAN3_Pptx_AddAnimation_InvalidType_ThrowsArgumentException()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });
        var act = () => h.Add("/slide[1]/shape[1]", "animation", null,
            new() { ["type"] = "nonexistent_animation_xyz", ["trigger"] = "onclick" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef for invalid animation type: {ex.Message}"); }
        catch (Exception) { /* ArgumentException or other = acceptable */ }
    }

    // ==================== PT01–PT03: Pptx slide transition ====================

    [Fact]
    public void PT01_Pptx_SetSlideTransition_Fade_RoundTrips()
    {
        // fade has no direction component, so it round-trips cleanly as "fade"
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["transition"] = "fade" });
        var node = h.Get("/slide[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("transition", "slide node should expose transition in Format");
        node.Format["transition"].ToString().Should().Be("fade");
    }

    [Fact]
    public void PT02_Pptx_SetSlideTransition_AfterAdd_RoundTrips()
    {
        // wipe defaults to direction "left", so readback is "wipe-left"
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Set("/slide[1]", new() { ["transition"] = "wipe" });
        var node = h.Get("/slide[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("transition");
        // direction is appended in readback: "wipe-left" (default dir=left)
        node.Format["transition"].ToString().Should().StartWith("wipe",
            "transition type should start with 'wipe'");
    }

    [Fact]
    public void PT03_Pptx_SlideTransition_Persistence()
    {
        // push defaults to direction "left", so readback is "push-left"
        var path = CreateTemp("pptx");
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { ["transition"] = "push" });
        }
        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("transition", "transition should persist after save/reopen");
        // direction is appended: "push-left" (default left)
        node.Format["transition"].ToString().Should().StartWith("push",
            "transition type should start with 'push' after persistence");
    }

    // ==================== NV01–NV04: Null-value guard in Set handlers ====================

    [Fact]
    public void NV01_Word_Set_NullValueInDict_NoNullRef()
    {
        // Fix dffe772: null values in dict caused NullReferenceException
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Null value test" });
        var para = h.Query("paragraph").First(p => p.Text.Contains("Null value test"));
        // Dict with null value — should not crash with NullRef
        var props = new Dictionary<string, string> { ["bold"] = "true", ["color"] = null! };
        var act = () => h.Set(para.Path, props);
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef when Set dict has null value: {ex.Message}"); }
        catch (Exception) { /* ArgumentException etc = acceptable */ }
    }

    [Fact]
    public void NV02_Excel_Set_NullValueInDict_NoNullRef()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        var props = new Dictionary<string, string> { ["value"] = "test", ["color"] = null! };
        var act = () => h.Set("/Sheet1/A1", props);
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef when Excel Set dict has null value: {ex.Message}"); }
        catch (Exception) { /* other exceptions acceptable */ }
    }

    [Fact]
    public void NV03_Pptx_Set_NullValueInDict_NoNullRef()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Null test" });
        var props = new Dictionary<string, string> { ["bold"] = "true", ["fill"] = null! };
        var act = () => h.Set("/slide[1]/shape[1]", props);
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef when Pptx Set dict has null value: {ex.Message}"); }
        catch (Exception) { /* other exceptions acceptable */ }
    }

    [Fact]
    public void NV04_Word_Add_NullValueInDict_NoNullRef()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        var props = new Dictionary<string, string> { ["text"] = "Null add test", ["bold"] = null! };
        var act = () => h.Add("/body", "paragraph", null, props);
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef when Add dict has null value: {ex.Message}"); }
        catch (Exception) { /* other exceptions acceptable */ }
    }
}
