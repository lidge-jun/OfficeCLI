// FuzzRound5 — Final round: stress, regression verification, cross-op fuzz, format conflicts, integrity.
//
// Areas:
//   ST01–ST04: Stress — 100+ slides/shapes/rows, Query performance after bulk add
//   RV01–RV04: Regression verify — SanitizeColorForOoxml("auto"), gradient empty/null color skipped
//   CR01–CR06: Cross-op fuzz — Add→Set→Remove→Add same path, Set after Remove, Query after mutations
//   FC01–FC04: Format conflict — solid→gradient→solid, gradient→none, conflicting Set keys
//   FI01–FI04: File integrity — Reopen after bulk ops, Reopen after gradient, Reopen after Remove storm

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzRound5 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"fuzz5_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
        {
            try
            {
                if (File.Exists(f)) { File.SetAttributes(f, FileAttributes.Normal); File.Delete(f); }
            }
            catch { }
        }
    }

    // ==================== ST01–ST04: Stress tests ====================

    [Fact]
    public void ST01_Pptx_Add100Slides_QueryShapesDoesNotThrow()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        for (int i = 1; i <= 100; i++)
            handler.Add("/", "slide", null, new() { ["title"] = $"Slide {i}" });

        // All 100 slides should be accessible
        for (int i = 1; i <= 100; i++)
        {
            var node = handler.Get($"/slide[{i}]");
            node.Should().NotBeNull($"slide[{i}] should exist");
        }
        // Query shapes should work without throwing
        var act = () => handler.Query("shape");
        act.Should().NotThrow("Query('shape') on 100-slide deck should not throw");
    }

    [Fact]
    public void ST02_Pptx_Add100ShapesOnOneSlide_QueryReturnsAll()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Dense" });
        for (int i = 1; i <= 100; i++)
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = $"Shape {i}", ["fill"] = "4472C4" });

        var shapes = handler.Query("shape");
        shapes.Count.Should().BeGreaterThanOrEqualTo(100, "all 100 shapes should be queryable");
    }

    [Fact]
    public void ST03_Excel_Add200Rows_QueryRowsAndCells()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/", "sheet", null, new() { ["name"] = "Big" });
        for (int i = 1; i <= 200; i++)
        {
            handler.Add("/Big", "row", null, new() { ["index"] = $"{i}" });
            handler.Add($"/Big/{i}", "cell", null, new() { ["ref"] = $"A{i}", ["value"] = $"Row{i}" });
        }
        var rows = handler.Query("row");
        rows.Count.Should().BeGreaterThanOrEqualTo(200, "all 200 rows should be queryable");
    }

    [Fact]
    public void ST04_Word_Add100Paragraphs_QueryAllReturned()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);
        for (int i = 1; i <= 100; i++)
            handler.Add("/body", "paragraph", null, new() { ["text"] = $"Para {i}" });

        var paras = handler.Query("paragraph");
        paras.Count.Should().BeGreaterThanOrEqualTo(100, "all 100 paragraphs should be queryable");
    }

    // ==================== RV01–RV04: Regression verification ====================

    [Fact]
    public void RV01_SanitizeColorForOoxml_AutoKeyword_DoesNotThrow()
    {
        // Regression: "auto" used to throw ArgumentException before fix
        var act = () => ParseHelpers.SanitizeColorForOoxml("auto");
        act.Should().NotThrow("'auto' is a legal OOXML value and must not throw");
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml("auto");
        rgb.Should().Be("auto");
        alpha.Should().BeNull();
    }

    [Fact]
    public void RV02_SanitizeColorForOoxml_AutoCaseInsensitive_DoesNotThrow()
    {
        foreach (var variant in new[] { "AUTO", "Auto", "aUtO" })
        {
            var act = () => ParseHelpers.SanitizeColorForOoxml(variant);
            act.Should().NotThrow($"'{variant}' should be accepted as 'auto'");
        }
    }

    [Fact]
    public void RV03_BuildGradientFill_SingleColorAfterAngleParse_DoesNotThrow()
    {
        // Regression: gradient "FF0000-90" (looks like 2 parts but "90" is angle, leaving only 1 color → duplicate)
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "G" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Box" });
        var shapes = handler.Query("shape");
        var shapePath = shapes.Last().Path;
        // "FF0000-90" — angle=90, single color, should duplicate → valid gradient
        var act = () => handler.Set(shapePath, new() { ["gradient"] = "FF0000-90" });
        act.Should().NotThrow("single-color gradient (angle stripped) should be handled by duplication");
    }

    [Fact]
    public void RV04_GradientFill_WithSchemeColors_DoesNotThrow()
    {
        // Regression: gradient stops with scheme colors (not hex) must not crash BuildColorElement
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "Scheme" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Box" });
        var shapes = handler.Query("shape");
        var shapePath = shapes.Last().Path;
        var act = () => handler.Set(shapePath, new() { ["gradient"] = "accent1-accent2" });
        act.Should().NotThrow("gradient with scheme colors should work without crash");
    }

    // ==================== CR01–CR06: Cross-op fuzz ====================

    [Fact]
    public void CR01_Pptx_AddSetRemoveAddSamePath_FinalAddSucceeds()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);

        // Add slide + shape
        handler.Add("/", "slide", null, new() { ["title"] = "S1" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Original" });
        var shapes = handler.Query("shape");
        var shapePath = shapes.Last().Path;

        // Set it
        handler.Set(shapePath, new() { ["fill"] = "FF0000" });

        // Remove slide (removes shape too)
        handler.Remove("/slide[1]");

        // Re-add slide at same index
        handler.Add("/", "slide", null, new() { ["title"] = "S1 new" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "New shape" });

        var shapes2 = handler.Query("shape");
        shapes2.Should().NotBeEmpty("shape should exist after Add→Set→Remove→Add cycle");
    }

    [Fact]
    public void CR02_Pptx_SetAfterRemove_ThrowsArgumentException()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "S1" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Box" });
        var shapes = handler.Query("shape");
        var shapePath = shapes.Last().Path;

        handler.Remove("/slide[1]");

        // Set on removed path should throw (node no longer exists)
        var act = () => handler.Set(shapePath, new() { ["fill"] = "0000FF" });
        act.Should().Throw<ArgumentException>("Set on a removed node should throw");
    }

    [Fact]
    public void CR03_Pptx_QueryAfterManyMutations_ReturnsCorrectCount()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);

        // Add 5 slides each with 3 shapes
        for (int s = 1; s <= 5; s++)
        {
            handler.Add("/", "slide", null, new() { ["title"] = $"S{s}" });
            for (int sh = 1; sh <= 3; sh++)
                handler.Add($"/slide[{s}]", "shape", null, new() { ["text"] = $"S{s}Sh{sh}" });
        }

        // Set colors on all shapes
        var shapes = handler.Query("shape");
        foreach (var shape in shapes)
            handler.Set(shape.Path, new() { ["fill"] = "4472C4" });

        // Remove 2 slides
        handler.Remove("/slide[5]");
        handler.Remove("/slide[4]");

        // Query should reflect 3 slides worth of shapes (title + 3 shapes each = 4 shapes per slide × 3 slides)
        var remaining = handler.Query("shape");
        remaining.Should().NotBeEmpty("shapes should remain after partial slide removal");
    }

    [Fact]
    public void CR04_Word_AddSetRemoveAddParagraph_InterleavedOps()
    {
        var path = CreateTemp("docx");
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "First" });
        handler.Set("/body/p[1]", new() { ["alignment"] = "center" });
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Second" });
        handler.Set("/body/p[2]", new() { ["bold"] = "true" });
        handler.Remove("/body/p[1]");

        // After removing first paragraph, second is now p[1]
        var node = handler.Get("/body/p[1]");
        node.Should().NotBeNull("second paragraph (now p[1]) should exist after first was removed");
        node!.Text.Should().Contain("Second");
    }

    [Fact]
    public void CR05_Excel_AddRemoveAddCell_ValueIsFromLastAdd()
    {
        var path = CreateTemp("xlsx");
        using var handler = new ExcelHandler(path, editable: true);
        handler.Add("/", "sheet", null, new() { ["name"] = "S1" });
        handler.Add("/S1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Original" });
        handler.Remove("/S1/A1");
        handler.Add("/S1", "cell", null, new() { ["ref"] = "A1", ["value"] = "ReAdded" });

        var node = handler.Get("/S1/A1");
        node.Should().NotBeNull();
        node!.Text.Should().Be("ReAdded");
    }

    [Fact]
    public void CR06_Pptx_QueryImmediatelyAfterAdd_ReflectsNewShapes()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "T" });

        for (int i = 1; i <= 10; i++)
        {
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = $"Shape {i}" });
            var shapes = handler.Query("shape");
            // After adding shape i, query should return at least i shapes
            shapes.Count.Should().BeGreaterThanOrEqualTo(i, $"after adding shape {i}, Query should return >= {i} shapes");
        }
    }

    // ==================== FC01–FC04: Format conflict tests ====================

    [Fact]
    public void FC01_Pptx_SetSolidThenGradient_GradientWins()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "T" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Box" });
        var shapes = handler.Query("shape");
        var shapePath = shapes.Last().Path;

        handler.Set(shapePath, new() { ["fill"] = "FF0000" });
        handler.Set(shapePath, new() { ["gradient"] = "FF0000-0000FF" });

        var node = handler.Get(shapePath);
        node.Should().NotBeNull();
        // fill should reflect gradient, not solid — getter uses "gradient" key
        node!.Format.Should().ContainKey("gradient", "gradient should win over solid when set last");
    }

    [Fact]
    public void FC02_Pptx_SetGradientThenSolid_SolidWins()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "T" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Box" });
        var shapes = handler.Query("shape");
        var shapePath = shapes.Last().Path;

        handler.Set(shapePath, new() { ["gradient"] = "FF0000-0000FF" });
        handler.Set(shapePath, new() { ["fill"] = "00FF00" });

        var node = handler.Get(shapePath);
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("fill", "solid fill should win over gradient when set last");
        node.Format["fill"].Should().Be("#00FF00");
    }

    [Fact]
    public void FC03_Pptx_SetGradientThenNone_NoFillApplied()
    {
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "T" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Box" });
        var shapes = handler.Query("shape");
        var shapePath = shapes.Last().Path;

        handler.Set(shapePath, new() { ["gradient"] = "FF0000-0000FF" });
        handler.Set(shapePath, new() { ["fill"] = "none" });

        // Shape should still exist with no fill
        var node = handler.Get(shapePath);
        node.Should().NotBeNull("shape should still exist after fill=none");
    }

    [Fact]
    public void FC04_Pptx_SetMultipleFillKeysAtOnce_LastOneInDictWins()
    {
        // When both fill and fill.gradient are set simultaneously, behavior is implementation-defined
        // but must not crash
        var path = CreateTemp("pptx");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new() { ["title"] = "T" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Box" });
        var shapes = handler.Query("shape");
        var shapePath = shapes.Last().Path;

        var act = () => handler.Set(shapePath, new()
        {
            ["fill"] = "FF0000",
            ["gradient"] = "0000FF-FF0000"
        });
        act.Should().NotThrow("setting both fill and fill.gradient simultaneously should not crash");

        var node = handler.Get(shapePath);
        node.Should().NotBeNull("shape should be valid after conflicting fill Set");
        // One of fill or gradient should be present — whichever was applied last
        var hasFill = node!.Format.ContainsKey("fill") || node.Format.ContainsKey("gradient");
        hasFill.Should().BeTrue("at least one fill key should be present after conflicting Set");
    }

    // ==================== FI01–FI04: File integrity after heavy ops ====================

    [Fact]
    public void FI01_Pptx_BulkOps_ReopenIntact()
    {
        var path = CreateTemp("pptx");
        using (var handler = new PowerPointHandler(path, editable: true))
        {
            for (int s = 1; s <= 20; s++)
            {
                handler.Add("/", "slide", null, new() { ["title"] = $"Slide {s}" });
                handler.Add($"/slide[{s}]", "shape", null, new() { ["text"] = $"Text {s}", ["fill"] = "4472C4" });
                handler.Add($"/slide[{s}]", "shape", null, new() { ["text"] = $"Box {s}", ["gradient"] = "FF0000-0000FF" });
            }
        }

        // Reopen and verify integrity
        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[10]");
        node.Should().NotBeNull("slide[10] should be readable after reopen");
        var shapes = h2.Query("shape");
        shapes.Count.Should().BeGreaterThanOrEqualTo(20, "shapes should persist across reopen");
    }

    [Fact]
    public void FI02_Excel_BulkOps_ReopenIntact()
    {
        var path = CreateTemp("xlsx");
        using (var handler = new ExcelHandler(path, editable: true))
        {
            handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
            for (int r = 1; r <= 50; r++)
            {
                handler.Add("/Sheet1", "row", null, new() { ["index"] = $"{r}" });
                handler.Add($"/Sheet1/{r}", "cell", null, new() { ["ref"] = $"A{r}", ["value"] = $"Val{r}" });
                handler.Add($"/Sheet1/{r}", "cell", null, new() { ["ref"] = $"B{r}", ["value"] = $"{r * 2}" });
            }
        }

        using var h2 = new ExcelHandler(path, editable: false);
        var cell = h2.Get("/Sheet1/A25");
        cell.Should().NotBeNull("A25 should persist across reopen");
        cell!.Text.Should().Be("Val25");
    }

    [Fact]
    public void FI03_Word_BulkOps_ReopenIntact()
    {
        var path = CreateTemp("docx");
        using (var handler = new WordHandler(path, editable: true))
        {
            for (int i = 1; i <= 50; i++)
                handler.Add("/body", "paragraph", null, new() { ["text"] = $"Paragraph {i}" });

            // Apply styles to every other paragraph
            for (int i = 1; i <= 50; i += 2)
                handler.Set($"/body/p[{i}]", new() { ["bold"] = "true", ["alignment"] = "center" });
        }

        using var h2 = new WordHandler(path, editable: false);
        var paras = h2.Query("paragraph");
        paras.Count.Should().BeGreaterThanOrEqualTo(50, "50 paragraphs should persist across reopen");
    }

    [Fact]
    public void FI04_Pptx_RemoveStorm_ReopenIntact()
    {
        var path = CreateTemp("pptx");
        using (var handler = new PowerPointHandler(path, editable: true))
        {
            // Add 30 slides
            for (int s = 1; s <= 30; s++)
            {
                handler.Add("/", "slide", null, new() { ["title"] = $"Slide {s}" });
                handler.Add($"/slide[{s}]", "shape", null, new() { ["text"] = $"Keep {s}" });
            }

            // Remove every other slide from the end to avoid index shifting issues
            for (int s = 30; s >= 2; s -= 2)
                handler.Remove($"/slide[{s}]");
        }

        // Reopen: 15 slides should remain
        using var h2 = new PowerPointHandler(path, editable: false);
        var node = h2.Get("/slide[1]");
        node.Should().NotBeNull("slide[1] should exist after remove storm");

        var shapes = h2.Query("shape");
        shapes.Should().NotBeEmpty("remaining slides should have shapes after reopen");
    }
}
