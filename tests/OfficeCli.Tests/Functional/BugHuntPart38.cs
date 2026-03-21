// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt round 38: Bugs found via BT regression testing.
/// Focus areas:
///   - BT-5: PPTX Query shape[text=X] never matches (text field not in Format dict),
///           and ~= operator not parsed (returns all shapes)
///   - BT-NEW-1: XLSX Get /Sheet1 does not include chart children (only rows)
///   - BT-NEW-2: XLSX Add cell without ref always writes to A1 (default "A1"),
///               so multiple cells to same row all overwrite each other
/// </summary>
public class BugHuntPart38 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // =====================================================================
    // BT-5 Group: PPTX Query shape[text=...] / [text~=...] broken
    //
    // Root cause in MatchesGenericAttributes (PowerPointHandler.Selector.cs):
    //   Only checks node.Format dictionary, never node.Text.
    //   DocumentNode.Text is NOT stored in Format, so [text=X] always fails.
    //
    // Root cause for ~= operator in ParseShapeSelector:
    //   Regex `(\\?!?=)` matches `=` and `!=` but NOT `~=`.
    //   `~=` is never captured as an attribute filter.
    //   So shape[text~=NOTHING] is parsed with no attribute filter at all
    //   and returns ALL shapes instead of zero.
    // =====================================================================

    // Bug3800: shape[text=Hello World] returns 0 results — should return 1
    // MatchesGenericAttributes only checks Format dict; node.Text is separate
    [Fact]
    public void Bug3800_Pptx_Query_TextEquals_Filter_Returns_Zero_Instead_Of_One()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello World" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Other Shape" });

        // Should return exactly 1 shape with text matching "Hello World"
        var results = handler.Query("shape[text=Hello World]");
        results.Should().HaveCount(1,
            "Query shape[text=Hello World] should match the shape whose Text is 'Hello World'");
        results[0].Text.Should().Be("Hello World");
    }

    // Bug3801: shape[text~=NOTHING] returns all shapes — should return 0
    // ~= is not a valid operator but should be treated as "contains" or reject;
    // instead ParseShapeSelector fails to parse it and applies NO filter,
    // so all shapes pass through.
    [Fact]
    public void Bug3801_Pptx_Query_TildeEquals_Operator_Returns_All_Instead_Of_Zero()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Alpha" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Beta" });

        // ~= is not supported; should return 0 (no match) or throw, not return all shapes
        var results = handler.Query("shape[text~=NOTHING]");
        results.Should().HaveCount(0,
            "shape[text~=NOTHING] should match nothing; ~= operator is unsupported or contains-match that finds no result");
    }

    // Bug3802: shape[text!=Hello World] should return 1 (the other shape)
    // But since "text" is not in Format, negate with missing key passes vacuously,
    // so BOTH shapes return (negate=true, key not found => not rejected => passes).
    // Result: [text!=Hello World] returns ALL 2 shapes instead of 1.
    [Fact]
    public void Bug3802_Pptx_Query_TextNotEquals_Returns_All_Instead_Of_One()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello World" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Other Shape" });

        // [text!=Hello World] should return 1 shape (the one whose text != "Hello World")
        var results = handler.Query("shape[text!=Hello World]");
        results.Should().HaveCount(1,
            "shape[text!=Hello World] should match only the shape whose Text is NOT 'Hello World'");
        results[0].Text.Should().Be("Other Shape");
    }

    // Bug3803 (PASS baseline): shape[fill=FF0000] correctly filters by Format["fill"]
    // This verifies that Format-based attribute filtering works as expected.
    [Fact]
    public void Bug3803_Pptx_Query_FillEquals_Filter_Works_Correctly_Baseline()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Red", ["fill"] = "FF0000" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Blue", ["fill"] = "0000FF" });

        // fill is in Format dict, so this should work correctly
        var results = handler.Query("shape[fill=FF0000]");
        results.Should().HaveCount(1, "shape[fill=FF0000] should match exactly the red-filled shape");
        results[0].Format["fill"].ToString().Should().Be("#FF0000");
    }

    // Bug3804: shape[text=Hello] should exact-match only that shape, not "Hello World"
    // Currently returns 0 (text field missing from Format), but after fix it should be 1.
    [Fact]
    public void Bug3804_Pptx_Query_TextEquals_Exact_Match_Not_Contains()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello World" });

        // shape[text=Hello] should return exactly 1 shape (exact match)
        var results = handler.Query("shape[text=Hello]");
        results.Should().HaveCount(1,
            "shape[text=Hello] should exactly match the shape with text 'Hello', not 'Hello World'");
        results[0].Text.Should().Be("Hello");
    }

    // =====================================================================
    // BT-NEW-1 Group: XLSX Get /Sheet1 doesn't include chart children
    //
    // Root cause in GetSheetChildNodes (ExcelHandler.Helpers.cs line 132):
    //   Only iterates sheetData.Elements<Row>().
    //   DrawingsPart with ChartParts is never examined.
    //   So Get /Sheet1 with depth>0 returns only row nodes, no chart nodes.
    //   But Query("chart") correctly reads DrawingsPart and returns charts.
    // =====================================================================

    // Bug3805: Get /Sheet1 children doesn't include chart after adding one
    [Fact]
    public void Bug3805_Xlsx_Get_Sheet_Children_Missing_Chart()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:100,200,300"
        });

        var sheetNode = handler.Get("/Sheet1", depth: 1);
        sheetNode.Type.Should().Be("sheet");

        // Chart should appear as a child of the sheet node
        var chartChildren = sheetNode.Children.Where(c => c.Type == "chart").ToList();
        chartChildren.Should().HaveCountGreaterThan(0,
            "Get /Sheet1 should include chart children in the sheet node, but currently only rows are returned");
    }

    // Bug3806: Query("chart") returns chart but Get /Sheet1 doesn't — inconsistency
    [Fact]
    public void Bug3806_Xlsx_Get_Sheet_And_Query_Chart_Are_Inconsistent()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:100,200,300"
        });

        // Query finds the chart
        var queryResults = handler.Query("chart");
        queryResults.Should().HaveCount(1, "Query('chart') should return the added chart");

        // Get /Sheet1 should ALSO find the chart in children — but currently does not
        var sheetNode = handler.Get("/Sheet1", depth: 1);
        var chartInGet = sheetNode.Children.Any(c => c.Type == "chart");
        chartInGet.Should().BeTrue(
            "Get /Sheet1 children should match what Query('chart') finds; they are inconsistent");
    }

    // Bug3807 (PASS baseline): Direct get /Sheet1/chart[1] works after add
    [Fact]
    public void Bug3807_Xlsx_Get_Direct_ChartPath_Works_After_Add_Baseline()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:100,200,300"
        });

        // Direct path access works
        var chartNode = handler.Get("/Sheet1/chart[1]");
        chartNode.Type.Should().Be("chart", "Direct Get /Sheet1/chart[1] should work correctly");
    }

    // =====================================================================
    // BT-NEW-2 Group: XLSX Add cell without ref always writes to A1
    //
    // Root cause in ExcelHandler.Add case "cell" (ExcelHandler.Add.cs line 81):
    //   var cellRef = properties.GetValueOrDefault("ref", "A1");
    //   Default is hardcoded "A1". When adding multiple cells to the same row
    //   without specifying ref, all go to A1, each overwriting the previous.
    //   The expected behavior is auto-increment: first cell → A1, second → B1, etc.
    // =====================================================================

    // Bug3808: Adding 3 cells to row[1] without ref — all land in A1, last value wins
    [Fact]
    public void Bug3808_Xlsx_Add_Multiple_Cells_Without_Ref_All_Write_To_A1()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "A" });
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "B" });
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "C" });

        // Expected: A1=A, B1=B, C1=C (auto-placed in next available column)
        // Actual: A1=C (all overwrite A1)
        var rowNode = handler.Get("/Sheet1/row[1]", depth: 1);
        rowNode.Children.Should().HaveCount(3,
            "Adding 3 cells to row[1] without explicit ref should create 3 distinct cells (A1, B1, C1), not overwrite A1 three times");
    }

    // Bug3809: After adding 3 cells without ref, A1 has value of LAST cell added
    [Fact]
    public void Bug3809_Xlsx_Add_Cells_Without_Ref_Last_Value_Overwrites_A1()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "First" });
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "Second" });
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "Third" });

        // A1 should be "First" if auto-placement starts at A1 and increments
        // But currently A1 = "Third" (last one wins because all default to A1)
        var a1 = handler.Get("/Sheet1/A1");
        a1.Text.Should().Be("First",
            "A1 should contain the first cell added; subsequent adds without ref should go to B1, C1");
    }

    // Bug3810 (PASS baseline): Adding cells with explicit ref creates distinct cells correctly
    [Fact]
    public void Bug3810_Xlsx_Add_Cells_With_Explicit_Ref_Works_Correctly_Baseline()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "A", ["ref"] = "A1" });
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "B", ["ref"] = "B1" });
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "C", ["ref"] = "C1" });

        var a1 = handler.Get("/Sheet1/A1");
        var b1 = handler.Get("/Sheet1/B1");
        var c1 = handler.Get("/Sheet1/C1");

        a1.Text.Should().Be("A", "A1 should be 'A' when explicitly specified");
        b1.Text.Should().Be("B", "B1 should be 'B' when explicitly specified");
        c1.Text.Should().Be("C", "C1 should be 'C' when explicitly specified");
    }

    // Bug3811: Two sequential cell adds to same row without ref should return different paths
    [Fact]
    public void Bug3811_Xlsx_Add_Cell_Sequential_Returns_Different_Paths()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        var path1 = handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "X" });
        var path2 = handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "Y" });

        // The returned paths should be different cells
        path1.Should().NotBe(path2,
            "Two sequential cell adds to the same row without ref should return different paths (different columns)");
    }
}
