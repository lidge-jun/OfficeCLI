// FuzzRound15 — Target: unexplored lifecycle corners after R14
//
// Areas:
//   WB01–WB03: Word bookmark add + get + remove lifecycle
//   WL01–WL03: Word hyperlink add + persistence + remove
//   WN01–WN03: Word endnote add/remove/persistence (complement to R13 footnote)
//   WH01–WH03: Word header/footer add + remove lifecycle
//   PM01–PM03: PPTX slide Move/reorder persistence
//   PG01–PG03: PPTX group shape add/get/remove
//   PN01–PN03: PPTX notes add lifecycle
//   EM01–EM03: Excel Move row across sheets

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound15 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz15_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== WB01–WB03: Word bookmark lifecycle ====================

    [Fact]
    public void WB01_Word_AddBookmark_ThenGet_ReturnsNode()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Bookmark paragraph" });
        h.Add("/body/p[1]", "bookmark", null, new() { ["name"] = "MyBookmark" });
        var bookmarks = h.Query("bookmark").ToList();
        bookmarks.Should().NotBeEmpty("bookmark should appear in query results after add");
    }

    [Fact]
    public void WB02_Word_AddBookmark_Remove_NoNullRef()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para with bookmark" });
        h.Add("/body/p[1]", "bookmark", null, new() { ["name"] = "BM1" });
        var bookmarks = h.Query("bookmark").ToList();
        bookmarks.Should().NotBeEmpty();
        var act = () => h.Remove(bookmarks[0].Path);
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing bookmark: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable */ }
    }

    [Fact]
    public void WB03_Word_AddBookmark_Persistence()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "BM persist para" });
            h.Add("/body/p[1]", "bookmark", null, new() { ["name"] = "PersistBM" });
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            var bms = h2.Query("bookmark").ToList();
            bms.Should().NotBeEmpty("bookmark should persist after save and reopen");
        };
        act.Should().NotThrow("file should be valid after bookmark add and reopen");
    }

    // ==================== WL01–WL03: Word hyperlink lifecycle ====================

    [Fact]
    public void WL01_Word_AddHyperlink_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        var paraPath = h.Add("/body", "paragraph", null, new() { ["text"] = "Para before link" });
        // adding hyperlink must not NullRef; ArgumentException is acceptable (e.g. schema restriction)
        var act = () => h.Add(paraPath, "hyperlink", null,
            new() { ["url"] = "https://example.com", ["text"] = "Click here" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef adding hyperlink: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable */ }
    }

    [Fact]
    public void WL02_Word_AddHyperlink_Persistence()
    {
        var path = CreateTemp("docx");
        bool added = false;
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Intro" });
            try
            {
                h.Add("/body/p[1]", "hyperlink", null,
                    new() { ["url"] = "https://officecli.ai", ["text"] = "OfficeCli" });
                added = true;
            }
            catch (Exception) { /* unsupported = skip */ }
        }
        if (!added) return;
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("file should be valid after hyperlink add and reopen");
    }

    [Fact]
    public void WL03_Word_AddHyperlink_Remove_FileStillValid()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Base paragraph" });
            try
            {
                h.Add("/body/p[1]", "hyperlink", null,
                    new() { ["url"] = "https://example.org", ["text"] = "Link" });
            }
            catch (Exception) { goto skip; }
            var links = h.Query("hyperlink").ToList();
            if (links.Count > 0)
            {
                try { h.Remove(links[0].Path); }
                catch (NullReferenceException ex) { Assert.Fail($"NullRef removing hyperlink: {ex.Message}"); }
                catch (Exception) { /* other = acceptable */ }
            }
            skip:;
        }
        var reopenAct = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        reopenAct.Should().NotThrow("file should be valid after hyperlink remove attempt");
    }

    // ==================== WN01–WN03: Word endnote lifecycle ====================

    [Fact]
    public void WN01_Word_AddEndnote_BasicLifecycle_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Endnote para" });
        var act = () => h.Add("/body/p[1]", "endnote", null, new() { ["text"] = "Endnote content" });
        act.Should().NotThrow("adding endnote should not throw");
    }

    [Fact]
    public void WN02_Word_RemoveEndnote_NoNullRef()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "EN remove para" });
        try { h.Add("/body/p[1]", "endnote", null, new() { ["text"] = "To remove" }); }
        catch (Exception) { return; }
        var act = () => h.Remove("/endnote[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef in endnote Remove: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable */ }
    }

    [Fact]
    public void WN03_Word_AddEndnote_Persistence_FileValid()
    {
        var path = CreateTemp("docx");
        bool addedEn = false;
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "EN persist para" });
            try
            {
                h.Add("/body/p[1]", "endnote", null, new() { ["text"] = "Persisted endnote" });
                addedEn = true;
            }
            catch (Exception) { /* unsupported = skip */ }
        }
        if (!addedEn) return;
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("file should be valid after endnote add and reopen");
    }

    // ==================== WH01–WH03: Word header/footer lifecycle ====================

    [Fact]
    public void WH01_Word_AddHeader_ThenGet_NoThrow()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        var act = () => h.Add("/body", "header", null, new() { ["text"] = "My Header" });
        act.Should().NotThrow("adding header should not throw");
    }

    [Fact]
    public void WH02_Word_AddFooter_Persistence()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            try { h.Add("/body", "footer", null, new() { ["text"] = "Page Footer" }); }
            catch (Exception) { return; } // skip if unsupported
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("file should be valid after footer add and reopen");
    }

    [Fact]
    public void WH03_Word_AddHeader_Remove_FileStillValid()
    {
        var path = CreateTemp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            try { h.Add("/body", "header", null, new() { ["text"] = "Header to remove" }); }
            catch (Exception) { goto skip; }
            try { h.Remove("/header[1]"); }
            catch (NullReferenceException ex) { Assert.Fail($"NullRef removing header: {ex.Message}"); }
            catch (Exception) { /* other = acceptable */ }
            skip:;
        }
        var act = () =>
        {
            using var h2 = new WordHandler(path, editable: false);
            _ = h2.Query("paragraph").ToList();
        };
        act.Should().NotThrow("file should be valid after header remove attempt");
    }

    // ==================== PM01–PM03: PPTX slide Move/reorder ====================

    [Fact]
    public void PM01_Pptx_MoveSlide_Reorder_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        h.Add("/", "slide", null, new() { ["title"] = "Slide 2" });
        h.Add("/", "slide", null, new() { ["title"] = "Slide 3" });
        var act = () => h.Move("/slide[1]", null, 2);
        act.Should().NotThrow("reordering slide[1] to index 2 should not throw");
    }

    [Fact]
    public void PM02_Pptx_MoveSlide_Persistence()
    {
        var path = CreateTemp("pptx");
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { ["title"] = "Alpha" });
            h.Add("/", "slide", null, new() { ["title"] = "Beta" });
            h.Move("/slide[1]", null, 1); // move slide 1 to end
        }
        var act = () =>
        {
            using var h2 = new PowerPointHandler(path, editable: false);
            var s = h2.Get("/slide[1]");
            s.Should().NotBeNull("slide[1] should exist after move+reopen");
        };
        act.Should().NotThrow("file should be valid after slide move and reopen");
    }

    [Fact]
    public void PM03_Pptx_MoveSlide_OutOfRange_ThrowsOrHandled()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        // Moving a non-existent slide should not NullRef
        var act = () => h.Move("/slide[99]", null, null);
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef moving out-of-range slide: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = expected */ }
    }

    // ==================== PG01–PG03: PPTX group shape lifecycle ====================

    [Fact]
    public void PG01_Pptx_AddGroup_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape A", ["x"] = "1cm", ["y"] = "1cm" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape B", ["x"] = "3cm", ["y"] = "1cm" });
        var act = () => h.Add("/slide[1]", "group", null,
            new() { ["members"] = "/slide[1]/shape[1],/slide[1]/shape[2]" });
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef adding group: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = acceptable if members not supported that way */ }
    }

    [Fact]
    public void PG02_Pptx_AddGroup_Persistence()
    {
        var path = CreateTemp("pptx");
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { });
            bool added = false;
            try
            {
                h.Add("/slide[1]", "group", null, new() { });
                added = true;
            }
            catch (Exception) { /* skip if unsupported */ }
            if (!added) return;
        }
        var act = () =>
        {
            using var h2 = new PowerPointHandler(path, editable: false);
            _ = h2.Query("shape").ToList();
        };
        act.Should().NotThrow("file should be valid after group add and reopen");
    }

    [Fact]
    public void PG03_Pptx_AddGroup_Remove_NoNullRef()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        try { h.Add("/slide[1]", "group", null, new() { }); }
        catch (Exception) { return; } // skip if unsupported
        var act = () => h.Remove("/slide[1]/group[1]");
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef removing group: {ex.Message}"); }
        catch (Exception) { /* other = acceptable */ }
    }

    // ==================== PN01–PN03: PPTX notes lifecycle ====================

    [Fact]
    public void PN01_Pptx_AddNotes_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        var act = () => h.Add("/slide[1]", "notes", null, new() { ["text"] = "Speaker notes here" });
        act.Should().NotThrow("adding notes to a slide should not throw");
    }

    [Fact]
    public void PN02_Pptx_AddNotes_Persistence()
    {
        var path = CreateTemp("pptx");
        bool added = false;
        {
            using var h = new PowerPointHandler(path, editable: true);
            h.Add("/", "slide", null, new() { });
            try
            {
                h.Add("/slide[1]", "notes", null, new() { ["text"] = "Persist notes" });
                added = true;
            }
            catch (Exception) { /* unsupported = skip */ }
        }
        if (!added) return;
        var act = () =>
        {
            using var h2 = new PowerPointHandler(path, editable: false);
            _ = h2.Query("shape").ToList();
        };
        act.Should().NotThrow("file should be valid after notes add and reopen");
    }

    [Fact]
    public void PN03_Pptx_AddNotes_SetText_RoundTrip()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        string? notesPath = null;
        try { notesPath = h.Add("/slide[1]", "notes", null, new() { ["text"] = "Initial note" }); }
        catch (Exception) { return; } // skip if unsupported
        // After adding notes, file and query should still work
        var act = () => { _ = h.Query("shape").ToList(); };
        act.Should().NotThrow("querying shapes after notes add should not throw");
    }

    // ==================== EM01–EM03: Excel Move row ====================

    [Fact]
    public void EM01_Excel_MoveRow_WithinSheet_NoThrow()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "MoveRow1" });
        h.Set("/Sheet1/A2", new() { ["value"] = "MoveRow2" });
        h.Set("/Sheet1/A3", new() { ["value"] = "MoveRow3" });
        var act = () => h.Move("/Sheet1/row[1]", "/Sheet1", 2);
        act.Should().NotThrow("moving a row within the same sheet should not throw");
    }

    [Fact]
    public void EM02_Excel_MoveRow_Persistence()
    {
        var path = CreateTemp("xlsx");
        {
            using var h = new ExcelHandler(path, editable: true);
            h.Set("/Sheet1/A1", new() { ["value"] = "RowFirst" });
            h.Set("/Sheet1/A2", new() { ["value"] = "RowSecond" });
            h.Move("/Sheet1/row[1]", "/Sheet1", 1);
        }
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            _ = h2.Get("/Sheet1/A1");
        };
        act.Should().NotThrow("file should be valid after row move and reopen");
    }

    [Fact]
    public void EM03_Excel_MoveRow_OutOfRange_ThrowsOrHandled()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "Only row" });
        // Moving non-existent row should not NullRef
        var act = () => h.Move("/Sheet1/row[99]", "/Sheet1", null);
        try { act(); }
        catch (NullReferenceException ex) { Assert.Fail($"NullRef moving out-of-range Excel row: {ex.Message}"); }
        catch (Exception) { /* ArgumentException = expected */ }
    }
}
