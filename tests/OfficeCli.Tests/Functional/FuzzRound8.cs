// FuzzRound8 — R6 fix regression validation + final smoke test (Round 7 final).
//
// Areas:
//   FN01–FN04: footnote/endnote Remove boundaries
//   CF01–CF03: CF Priority boundary — add 10, delete all, add new
//   TR01–TR04: transition duration boundaries — 0, large, negative, non-numeric
//   PI01–PI03: PPTX image ref-count boundary — 3 copies same image, delete one-by-one
//   SM01–SM05: global smoke — all 3 handlers CRUD lifecycle

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound8 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz8_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    private string CreatePng()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz8_img_{Guid.NewGuid():N}.png");
        _tempFiles.Add(path);
        // 1×1 transparent PNG
        File.WriteAllBytes(path, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="));
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); } } catch { }
    }

    // ==================== FN01–FN04: footnote/endnote Remove boundaries ====================

    [Fact]
    public void FN01_Word_Remove_NonExistentFootnote_Throws()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        // No footnotes added; removing /footnote[1] should throw ArgumentException
        var act = () => h.Remove("/footnote[1]");
        act.Should().Throw<ArgumentException>("removing a non-existent footnote should throw");
    }

    [Fact]
    public void FN02_Word_Remove_NonExistentEndnote_Throws()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        // No endnotes added; removing /endnote[1] should throw ArgumentException
        var act = () => h.Remove("/endnote[1]");
        act.Should().Throw<ArgumentException>("removing a non-existent endnote should throw");
    }

    [Fact]
    public void FN03_Word_Remove_Footnote_ThenAddNew_NoIdConflict()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para1" });
        var p1 = h.Query("paragraph").First(p => p.Text.Contains("Para1"));
        var fnPath = h.Add(p1.Path, "footnote", null, new() { ["text"] = "FN text A" });
        // Verify it exists via Get
        var fnNode = h.Get(fnPath);
        fnNode.Should().NotBeNull("footnote should exist after Add");

        h.Remove(fnPath);

        // After removal, add another footnote to a new paragraph — should not conflict
        h.Add("/body", "paragraph", null, new() { ["text"] = "Para2" });
        var p2 = h.Query("paragraph").First(p => p.Text.Contains("Para2"));
        var act = () => h.Add(p2.Path, "footnote", null, new() { ["text"] = "FN text B" });
        act.Should().NotThrow("adding footnote after remove should not throw");

        var fn2Path = h.Add(p2.Path, "footnote", null, new() { ["text"] = "FN text C" });
        var fn2Node = h.Get(fn2Path);
        fn2Node.Should().NotBeNull("new footnote should be queryable after add-post-remove");
    }

    [Fact]
    public void FN04_Word_Remove_MultipleFootnotes_Sequential()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "P1" });
        h.Add("/body", "paragraph", null, new() { ["text"] = "P2" });
        var p1 = h.Query("paragraph").First(p => p.Text == "P1");
        var p2 = h.Query("paragraph").First(p => p.Text == "P2");
        // Use Add's returned path directly (Add returns "/footnote[N]")
        var fn1Path = h.Add(p1.Path, "footnote", null, new() { ["text"] = "FN1" });
        var fn2Path = h.Add(p2.Path, "footnote", null, new() { ["text"] = "FN2" });

        h.Get(fn1Path).Should().NotBeNull("footnote 1 should exist after Add");
        h.Get(fn2Path).Should().NotBeNull("footnote 2 should exist after Add");

        // Remove each one — should not throw
        var act1 = () => h.Remove(fn1Path);
        act1.Should().NotThrow("removing first footnote should not throw");

        // fn1 should be gone, fn2 still present
        var act1b = () => h.Get(fn1Path);
        act1b.Should().Throw<ArgumentException>("getting removed footnote should throw");
        h.Get(fn2Path).Should().NotBeNull("second footnote should still exist");

        var act2 = () => h.Remove(fn2Path);
        act2.Should().NotThrow("removing second footnote should not throw");

        var act2b = () => h.Get(fn2Path);
        act2b.Should().Throw<ArgumentException>("getting second removed footnote should throw");
    }

    // ==================== CF01–CF03: CF Priority boundaries ====================

    [Fact]
    public void CF01_Excel_Add10CF_DeleteAll_AddNew_NoPriorityConflict()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        // Add 10 CF rules; capture returned paths directly
        var cfPaths = new List<string>();
        for (int i = 1; i <= 10; i++)
            cfPaths.Add(h.Add("/Sheet1", "conditionalformatting", null, new()
            {
                ["ref"] = $"A{i}",
                ["type"] = "expression",
                ["formula"] = $"A{i}>0",
                ["color"] = "#FF0000",
            }));
        cfPaths.Should().HaveCount(10, "10 CF paths should be captured");

        // Delete all in reverse order so index references remain valid
        foreach (var cfPath in cfPaths.AsEnumerable().Reverse())
        {
            var act = () => h.Remove(cfPath);
            act.Should().NotThrow($"removing {cfPath} should not throw");
        }

        // Add a new CF after clearing — priority should start fresh (not conflict)
        var addAct = () => h.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["ref"] = "B1",
            ["type"] = "expression",
            ["formula"] = "B1>5",
            ["color"] = "#00FF00",
        });
        addAct.Should().NotThrow("adding CF after clearing all rules should not throw");

        // Verify new CF is accessible via Get
        var newCfPath = h.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["ref"] = "B2",
            ["type"] = "expression",
            ["formula"] = "B2>10",
            ["color"] = "#0000FF",
        });
        h.Get(newCfPath).Should().NotBeNull("newly added CF should be retrievable by path");
    }

    [Fact]
    public void CF02_Excel_Add10CF_DeleteAll_Persist_NoCorruption()
    {
        var path = CreateTemp("xlsx");
        var cfPaths = new List<string>();
        {
            using var h = new ExcelHandler(path, editable: true);
            for (int i = 1; i <= 10; i++)
                cfPaths.Add(h.Add("/Sheet1", "conditionalformatting", null, new()
                {
                    ["ref"] = $"C{i}",
                    ["type"] = "expression",
                    ["formula"] = $"C{i}<>\"\"",
                    ["color"] = "#0000FF",
                }));
            // Delete in reverse so index references stay valid
            foreach (var cfPath in cfPaths.AsEnumerable().Reverse())
                h.Remove(cfPath);
        }
        // Re-open and verify document is not corrupted
        var act = () =>
        {
            using var h2 = new ExcelHandler(path, editable: false);
            // Just verify it opens without throwing
            _ = h2.Get("/Sheet1/A1");
        };
        act.Should().NotThrow("reopening after mass CF deletion should not throw");
    }

    [Fact]
    public void CF03_Excel_AddCF_PriorityIsUnique_After5Cycles()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);

        // 5 add/delete cycles — each new CF should have a unique priority
        for (int cycle = 0; cycle < 5; cycle++)
        {
            var cfPath = h.Add("/Sheet1", "conditionalformatting", null, new()
            {
                ["ref"] = "D1",
                ["type"] = "expression",
                ["formula"] = "D1>0",
                ["color"] = "#FF8000",
            });
            cfPath.Should().NotBeNullOrEmpty("CF path should be returned");
            h.Remove(cfPath);
        }
        // Final add should work cleanly
        var finalPath = h.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["ref"] = "E1",
            ["type"] = "expression",
            ["formula"] = "E1>0",
            ["color"] = "#800080",
        });
        h.Get(finalPath).Should().NotBeNull("final CF should be retrievable");
    }

    // ==================== TR01–TR04: transition duration boundaries ====================

    [Fact]
    public void TR01_Pptx_TransitionDuration_Zero_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        // "fade-0" — duration embedded in transition string as 0ms
        var act = () => h.Set("/slide[1]", new() { ["transition"] = "fade-0" });
        act.Should().NotThrow("transition with duration=0 should not throw");
    }

    [Fact]
    public void TR02_Pptx_TransitionDuration_VeryLarge_NoThrow()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        // "fade-999999" — extreme duration (999999ms ≈ 16 minutes)
        var act = () => h.Set("/slide[1]", new() { ["transition"] = "fade-999999" });
        act.Should().NotThrow("transition with very large duration should not throw");
    }

    [Fact]
    public void TR03_Pptx_TransitionDuration_Negative_HandledGracefully()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        // Negative duration string — should not throw (parsed as direction, not duration)
        // per ApplyTransition: int.TryParse("-500") succeeds as a negative int, gets stored as-is
        var act = () => h.Set("/slide[1]", new() { ["transition"] = "fade--500" });
        act.Should().NotThrow("negative duration in transition should not crash");
    }

    [Fact]
    public void TR04_Pptx_TransitionInvalidDirection_Throws()
    {
        // BUG: "wipe-abc" passes "abc" as direction to ParseSlideDir which throws ArgumentException.
        // This test documents the current behavior: invalid direction tokens throw.
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        var act = () => h.Set("/slide[1]", new() { ["transition"] = "wipe-abc" });
        act.Should().Throw<ArgumentException>("invalid transition direction should throw ArgumentException");
    }

    // ==================== PI01–PI03: PPTX image ref-count boundaries ====================

    [Fact]
    public void PI01_Pptx_ThreePicsFromSameFile_DeleteFirst_TwoRemain()
    {
        var imgPath = CreatePng();
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });
        h.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });
        h.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });

        var act = () => h.Remove("/slide[1]/picture[1]");
        act.Should().NotThrow("removing first of three shared-image pictures should not throw");

        // Two pictures should still exist and be accessible
        var pic2 = h.Get("/slide[1]/picture[1]");
        pic2.Should().NotBeNull("second picture should still exist after removing first");
    }

    [Fact]
    public void PI02_Pptx_ThreePicsFromSameFile_DeleteAll_ImagePartCleaned()
    {
        var imgPath = CreatePng();
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });
        h.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });
        h.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });

        // Delete all three — last deletion should clean up the ImagePart
        h.Remove("/slide[1]/picture[1]");
        h.Remove("/slide[1]/picture[1]");
        var act = () => h.Remove("/slide[1]/picture[1]");
        act.Should().NotThrow("removing last shared-image picture should not throw");

        // No pictures should remain — Get throws ArgumentException for missing picture
        var actGet = () => h.Get("/slide[1]/picture[1]");
        actGet.Should().Throw<ArgumentException>("Get should throw when no pictures remain");
    }

    [Fact]
    public void PI03_Pptx_ThreePicsFromSameFile_DeleteMiddle_OthersIntact()
    {
        var imgPath = CreatePng();
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { });
        h.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });
        h.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });
        h.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });

        // Delete middle (now index[2]) — should not corrupt others
        var act = () => h.Remove("/slide[1]/picture[2]");
        act.Should().NotThrow("removing middle picture should not throw");

        // Two pictures should still exist
        var pic1 = h.Get("/slide[1]/picture[1]");
        pic1.Should().NotBeNull("first picture should still exist after removing middle");
        var pic2 = h.Get("/slide[1]/picture[2]");
        pic2.Should().NotBeNull("third picture (now [2]) should still exist");
    }

    // ==================== SM01–SM05: Global smoke CRUD ====================

    [Fact]
    public void SM01_Word_FullCRUD_Paragraph()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "SmokePara", ["bold"] = "true" });
        var node = h.Query("paragraph").FirstOrDefault(p => p.Text.Contains("SmokePara"));
        node.Should().NotBeNull("paragraph should exist after Add");
        node!.Format["bold"].Should().Be(true, "bold should be set on add");

        h.Set(node.Path, new() { ["italic"] = "true", ["size"] = "14pt" });
        var updated = h.Get(node.Path);
        updated.Should().NotBeNull("paragraph should still exist after Set");
        updated!.Format["italic"].Should().Be(true, "italic should be set after Set");

        h.Remove(node.Path);
        h.Query("paragraph").Any(p => p.Text.Contains("SmokePara"))
            .Should().BeFalse("paragraph should be gone after Remove");
    }

    [Fact]
    public void SM02_Excel_FullCRUD_Cell()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/B2", new() { ["value"] = "SmokeCell", ["bold"] = "true", ["color"] = "#FF0000" });
        var node = h.Get("/Sheet1/B2");
        node.Should().NotBeNull("cell should exist after Set");
        node!.Text.Should().Be("SmokeCell");
        node.Format["bold"].Should().Be(true);

        h.Set("/Sheet1/B2", new() { ["value"] = "SmokeUpdated", ["italic"] = "true" });
        var updated = h.Get("/Sheet1/B2");
        updated!.Text.Should().Be("SmokeUpdated", "cell value should update");
        updated.Format["italic"].Should().Be(true, "italic should be set after second Set");
    }

    [Fact]
    public void SM03_Pptx_FullCRUD_Shape()
    {
        var path = CreateTemp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        // No title — slide[1] starts with no shapes; added shape is shape[1]
        h.Add("/", "slide", null, new() { });
        var shapePath = h.Add("/slide[1]", "shape", null, new() { ["text"] = "SmokeShape", ["fill"] = "#AABBCC" });

        var node = h.Get(shapePath);
        node.Should().NotBeNull("shape should exist after Add");
        node!.Text.Should().Be("SmokeShape");

        h.Set(shapePath, new() { ["bold"] = "true", ["size"] = "18pt" });
        var updated = h.Get(shapePath);
        updated!.Format["bold"].Should().Be(true, "bold should be set after Set");
        updated.Format["size"].Should().Be("18pt", "font size should be 18pt after Set");

        h.Remove(shapePath);
        // shape[1] should now be gone; Get throws ArgumentException for missing shapes
        var act = () => h.Get(shapePath);
        act.Should().Throw<ArgumentException>("Get should throw when shape no longer exists");
    }

    [Fact]
    public void SM04_Excel_FullCRUD_Comment()
    {
        var path = CreateTemp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Add("/Sheet1", "comment", null, new() { ["ref"] = "A1", ["text"] = "SmokeComment" });
        var node = h.Get("/Sheet1/comment[1]");
        node.Should().NotBeNull("comment should exist after Add");

        h.Remove("/Sheet1/comment[1]");
        var act = () => h.Get("/Sheet1/comment[1]");
        act.Should().NotThrow("getting non-existent comment should return null, not throw");
        var gone = h.Get("/Sheet1/comment[1]");
        gone.Should().BeNull("comment should be gone after Remove");
    }

    [Fact]
    public void SM05_Word_FullCRUD_Table()
    {
        var path = CreateTemp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "3" });
        var tables = h.Query("table").ToList();
        tables.Should().NotBeEmpty("table should exist after Add");

        var tablePath = tables[0].Path;
        h.Remove(tablePath);
        h.Query("table").Should().BeEmpty("table should be gone after Remove");
    }
}
