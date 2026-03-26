// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt tests targeting potential chart issues:
/// - Multi-chart index stability after reopen
/// - Pie chart with axis properties (silently ignored)
/// - Edge cases in chart operations
/// </summary>
public class ChartBugHuntTests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _docxPath;
    private ExcelHandler _excel;
    private WordHandler _word;

    public ChartBugHuntTests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_docxPath);
        _excel = new ExcelHandler(_xlsxPath, editable: true);
        _word = new WordHandler(_docxPath, editable: true);
    }

    public void Dispose()
    {
        _excel.Dispose();
        _word.Dispose();
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    private void ReopenExcel() { _excel.Dispose(); _excel = new ExcelHandler(_xlsxPath, editable: true); }
    private void ReopenWord() { _word.Dispose(); _word = new WordHandler(_docxPath, editable: true); }

    // ==================== BUG 1: Multi-chart index stability (Word) ====================
    // After adding 2+ charts and reopening, does /chart[1] still point to the first chart?

    [Fact]
    public void Word_MultiChart_IndexStable_AfterReopen()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "First Chart", ["data"] = "S1:1,2,3"
        });
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Second Chart", ["data"] = "S1:4,5,6"
        });
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "pie", ["title"] = "Third Chart", ["data"] = "S1:7,8,9", ["categories"] = "A,B,C"
        });

        // Verify before reopen
        ((string)_word.Get("/chart[1]").Format["title"]).Should().Be("First Chart");
        ((string)_word.Get("/chart[2]").Format["title"]).Should().Be("Second Chart");
        ((string)_word.Get("/chart[3]").Format["title"]).Should().Be("Third Chart");

        // Reopen and verify index stability
        ReopenWord();
        ((string)_word.Get("/chart[1]").Format["title"]).Should().Be("First Chart",
            "chart[1] should still be the first chart after reopen");
        ((string)_word.Get("/chart[2]").Format["title"]).Should().Be("Second Chart",
            "chart[2] should still be the second chart after reopen");
        ((string)_word.Get("/chart[3]").Format["title"]).Should().Be("Third Chart",
            "chart[3] should still be the third chart after reopen");
    }

    // ==================== BUG 2: Multi-chart index stability (Excel) ====================

    [Fact]
    public void Excel_MultiChart_IndexStable_AfterReopen()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "First Chart", ["data"] = "S1:1,2,3"
        });
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Second Chart", ["data"] = "S1:4,5,6"
        });
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "Third Chart", ["data"] = "S1:7,8,9"
        });

        // Verify before reopen
        ((string)_excel.Get("/Sheet1/chart[1]").Format["title"]).Should().Be("First Chart");
        ((string)_excel.Get("/Sheet1/chart[2]").Format["title"]).Should().Be("Second Chart");
        ((string)_excel.Get("/Sheet1/chart[3]").Format["title"]).Should().Be("Third Chart");

        // Reopen and verify
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]").Format["title"]).Should().Be("First Chart",
            "chart[1] should still be the first chart after reopen");
        ((string)_excel.Get("/Sheet1/chart[2]").Format["title"]).Should().Be("Second Chart",
            "chart[2] should still be the second chart after reopen");
        ((string)_excel.Get("/Sheet1/chart[3]").Format["title"]).Should().Be("Third Chart",
            "chart[3] should still be the third chart after reopen");
    }

    // ==================== BUG 3: Set on specific chart in multi-chart doc ====================
    // Set on chart[2] should not affect chart[1] or chart[3]

    [Fact]
    public void Word_MultiChart_SetOnSecond_DoesNotAffectOthers()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Chart A", ["data"] = "S1:1,2"
        });
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Chart B", ["data"] = "S1:3,4"
        });
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "Chart C", ["data"] = "S1:5,6"
        });

        // Set only chart[2]
        _word.Set("/chart[2]", new() { ["title"] = "Modified B", ["legend"] = "top" });

        // Verify chart[1] and chart[3] are untouched
        ((string)_word.Get("/chart[1]").Format["title"]).Should().Be("Chart A");
        ((string)_word.Get("/chart[2]").Format["title"]).Should().Be("Modified B");
        ((string)_word.Get("/chart[3]").Format["title"]).Should().Be("Chart C");

        // Reopen and verify
        ReopenWord();
        ((string)_word.Get("/chart[1]").Format["title"]).Should().Be("Chart A");
        ((string)_word.Get("/chart[2]").Format["title"]).Should().Be("Modified B");
        ((string)_word.Get("/chart[3]").Format["title"]).Should().Be("Chart C");
    }

    [Fact]
    public void Excel_MultiChart_SetOnSecond_DoesNotAffectOthers()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Chart A", ["data"] = "S1:1,2"
        });
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "Chart B", ["data"] = "S1:3,4"
        });

        _excel.Set("/Sheet1/chart[2]", new() { ["title"] = "Modified B" });

        ((string)_excel.Get("/Sheet1/chart[1]").Format["title"]).Should().Be("Chart A");
        ((string)_excel.Get("/Sheet1/chart[2]").Format["title"]).Should().Be("Modified B");

        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]").Format["title"]).Should().Be("Chart A");
        ((string)_excel.Get("/Sheet1/chart[2]").Format["title"]).Should().Be("Modified B");
    }

    // ==================== BUG 4: Pie chart with axis properties ====================
    // Pie charts have no axes. Axis properties should be silently ignored, not crash.

    [Fact]
    public void Word_PieChart_AxisProperties_DoNotCrash()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "pie", ["title"] = "Pie", ["data"] = "S1:10,20,30", ["categories"] = "A,B,C"
        });

        // These should not throw — pie charts have no axes
        var unsupported = _word.Set("/chart[1]", new()
        {
            ["axistitle"] = "Should Ignore",
            ["cattitle"] = "Should Ignore",
            ["axismin"] = "0",
            ["axismax"] = "100"
        });

        // Verify chart is still valid
        var node = _word.Get("/chart[1]");
        node.Type.Should().Be("chart");
        ((string)node.Format["chartType"]).Should().Be("pie");
        ((string)node.Format["title"]).Should().Be("Pie");
        // Axis properties should NOT be present (no axes on pie)
        node.Format.Should().NotContainKey("axisTitle");
    }

    [Fact]
    public void Excel_PieChart_DataLabels_WorkCorrectly()
    {
        // Pie charts should support dataLabels even though they lack axes
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "pie", ["title"] = "Pie DL",
            ["data"] = "S1:10,20,30", ["categories"] = "A,B,C",
            ["datalabels"] = "value,percent"
        });

        var node = _excel.Get("/Sheet1/chart[1]");
        ((string)node.Format["dataLabels"]).Should().Contain("value");
        ((string)node.Format["dataLabels"]).Should().Contain("percent");

        ReopenExcel();
        node = _excel.Get("/Sheet1/chart[1]");
        ((string)node.Format["dataLabels"]).Should().Contain("value");
    }

    // ==================== BUG 5: Doughnut chart with axis properties ====================

    [Fact]
    public void Excel_DoughnutChart_AxisProperties_DoNotCrash()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "doughnut", ["title"] = "Donut",
            ["data"] = "S1:10,20,30", ["categories"] = "A,B,C",
            ["axistitle"] = "Should Ignore", ["axismin"] = "0"
        });

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Type.Should().Be("chart");
        ((string)node.Format["chartType"]).Should().Be("doughnut");
        // No axes → no axisTitle
        node.Format.Should().NotContainKey("axisTitle");
    }

    // ==================== BUG 6: Legend=none at Add time ====================

    [Fact]
    public void Word_AddChart_LegendNone_NoLegendInGet()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "No Legend",
            ["data"] = "S1:1,2;S2:3,4", ["legend"] = "none"
        });

        var node = _word.Get("/chart[1]");
        node.Format.Should().NotContainKey("legend",
            "legend=none at Add time should result in no legend element");
    }

    [Fact]
    public void Excel_AddChart_LegendFalse_NoLegendInGet()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "No Legend",
            ["data"] = "S1:1,2;S2:3,4", ["legend"] = "false"
        });

        var node = _excel.Get("/Sheet1/chart[1]");
        node.Format.Should().NotContainKey("legend",
            "legend=false at Add time should result in no legend element");
    }

    // ==================== BUG 7: Title removal via Set ====================

    [Fact]
    public void Word_SetTitle_None_RemovesTitle()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Has Title", ["data"] = "S1:1,2"
        });

        ((string)_word.Get("/chart[1]").Format["title"]).Should().Be("Has Title");

        _word.Set("/chart[1]", new() { ["title"] = "none" });

        _word.Get("/chart[1]").Format.Should().NotContainKey("title",
            "title=none should remove the title");

        ReopenWord();
        _word.Get("/chart[1]").Format.Should().NotContainKey("title");
    }

    // ==================== BUG 8: Legend removal via Set ====================

    [Fact]
    public void Excel_SetLegend_None_RemovesLegend()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T",
            ["data"] = "S1:1,2;S2:3,4", ["legend"] = "top"
        });

        ((string)_excel.Get("/Sheet1/chart[1]").Format["legend"]).Should().Be("top");

        _excel.Set("/Sheet1/chart[1]", new() { ["legend"] = "none" });

        _excel.Get("/Sheet1/chart[1]").Format.Should().NotContainKey("legend",
            "legend=none should remove the legend");

        ReopenExcel();
        _excel.Get("/Sheet1/chart[1]").Format.Should().NotContainKey("legend");
    }

    // ==================== BUG 9: Empty data edge case ====================

    [Fact]
    public void Word_AddChart_NoData_Throws()
    {
        var act = () => _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "No Data"
        });

        act.Should().Throw<ArgumentException>().WithMessage("*requires data*");
    }

    // ==================== BUG 10: Scatter chart with axis properties ====================

    [Fact]
    public void Excel_ScatterChart_AxisProperties_Work()
    {
        // Scatter charts have TWO ValueAxes (no CategoryAxis)
        // axisTitle should work (targets first ValueAxis)
        // cattitle should NOT work (no CategoryAxis)
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "scatter", ["title"] = "Scatter",
            ["data"] = "S1:1,2,3", ["categories"] = "10,20,30",
            ["axistitle"] = "Y Values"
        });

        var node = _excel.Get("/Sheet1/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("scatter");
        ((string)node.Format["axisTitle"]).Should().Be("Y Values");
    }

    // ==================== BUG 11: Query returns correct chart count ====================

    [Fact]
    public void Word_Query_Chart_ReturnsCorrectCount_AfterReopen()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "A", ["data"] = "S1:1,2"
        });
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "B", ["data"] = "S1:3,4"
        });

        _word.Query("chart").Should().HaveCount(2);

        ReopenWord();
        _word.Query("chart").Should().HaveCount(2,
            "Query should find same number of charts after reopen");
    }
}
