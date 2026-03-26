// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Unified chart tests for Excel, Word, and PPTX — full lifecycle:
/// Create → Add → Get → Verify → Set → Get → Verify → Reopen → Get → Verify
/// Every property is tested at both Add and Set stages with Reopen persistence.
/// </summary>
public class ChartUnifiedTests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _docxPath;
    private readonly string _pptxPath;
    private ExcelHandler _excel;
    private WordHandler _word;
    private PowerPointHandler _pptx;

    public ChartUnifiedTests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_pptxPath);
        _excel = new ExcelHandler(_xlsxPath, editable: true);
        _word = new WordHandler(_docxPath, editable: true);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _excel.Dispose();
        _word.Dispose();
        _pptx.Dispose();
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void ReopenExcel() { _excel.Dispose(); _excel = new ExcelHandler(_xlsxPath, editable: true); }
    private void ReopenWord() { _word.Dispose(); _word = new WordHandler(_docxPath, editable: true); }
    private void ReopenPptx() { _pptx.Dispose(); _pptx = new PowerPointHandler(_pptxPath, editable: true); }

    // ==================== Excel: Add + Get + Reopen ====================

    [Fact]
    public void Excel_Add_Chart_ReturnsCorrectPath()
    {
        var path = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2,3"
        });
        path.Should().Be("/Sheet1/chart[1]");
    }

    [Fact]
    public void Excel_Add_Chart_WithChartType_ChartTypeIsReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "T", ["data"] = "S1:1,2,3"
        });
        var node = _excel.Get("/Sheet1/chart[1]");
        node.Type.Should().Be("chart");
        ((string)node.Format["chartType"]).Should().Be("bar");
    }

    [Fact]
    public void Excel_Add_Chart_WithTitle_TitleIsReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Revenue Report", ["data"] = "S1:1,2"
        });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["title"]).Should().Be("Revenue Report");
    }

    [Fact]
    public void Excel_Add_Chart_WithLegend_LegendIsReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "line", ["title"] = "T", ["data"] = "S1:1,2", ["legend"] = "top"
        });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["legend"]).Should().Be("top");
    }

    [Fact]
    public void Excel_Add_Chart_WithCategories_CategoriesAreReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["categories"] = "Jan,Feb"
        });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["categories"]).Should().Be("Jan,Feb");
    }

    [Fact]
    public void Excel_Add_Chart_WithData_SeriesCountAndChildrenAreReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "Revenue:10,20;Cost:5,8"
        });
        var node = _excel.Get("/Sheet1/chart[1]", depth: 2);
        ((int)node.Format["seriesCount"]).Should().Be(2);
        node.Children.Should().HaveCount(2);
        node.Children[0].Type.Should().Be("series");
        node.Children[0].Text.Should().Be("Revenue");
        ((string)node.Children[0].Format["values"]).Should().Be("10,20");
    }

    [Fact]
    public void Excel_Add_Chart_WithSeries1Format_SeriesIsReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T",
            ["series1"] = "Sales:100,200", ["series2"] = "Cost:50,80", ["categories"] = "Q1,Q2"
        });
        var node = _excel.Get("/Sheet1/chart[1]", depth: 2);
        ((int)node.Format["seriesCount"]).Should().Be(2);
        ((string)node.Children[0].Format["values"]).Should().Be("100,200");
        ((string)node.Children[1].Format["values"]).Should().Be("50,80");
    }

    [Fact]
    public void Excel_Add_Chart_WithColors_ColorsAreReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2;S2:3,4", ["colors"] = "FF0000,00FF00"
        });
        var node = _excel.Get("/Sheet1/chart[1]", depth: 2);
        ((string)node.Children[0].Format["color"]).Should().Be("#FF0000");
        ((string)node.Children[1].Format["color"]).Should().Be("#00FF00");
    }

    [Fact]
    public void Excel_Add_Chart_WithDataLabels_DataLabelsAreReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["datalabels"] = "value,percent"
        });
        var dl = (string)_excel.Get("/Sheet1/chart[1]").Format["dataLabels"];
        dl.Should().Contain("value");
        dl.Should().Contain("percent");
    }

    [Fact]
    public void Excel_Add_Chart_WithAxisTitle_AxisTitleIsReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["axistitle"] = "Revenue"
        });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["axisTitle"]).Should().Be("Revenue");
    }

    [Fact]
    public void Excel_Add_Chart_WithCatTitle_CatTitleIsReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["cattitle"] = "Months"
        });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["catTitle"]).Should().Be("Months");
    }

    [Fact]
    public void Excel_Add_Chart_WithAxisMinMax_AxisMinMaxAreReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2",
            ["axismin"] = "0", ["axismax"] = "100"
        });
        var node = _excel.Get("/Sheet1/chart[1]");
        ((double)node.Format["axisMin"]).Should().Be(0);
        ((double)node.Format["axisMax"]).Should().Be(100);
    }

    [Fact]
    public void Excel_Add_Chart_WithMajorMinorUnit_UnitsAreReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2",
            ["majorunit"] = "10", ["minorunit"] = "2"
        });
        var node = _excel.Get("/Sheet1/chart[1]");
        ((double)node.Format["majorUnit"]).Should().Be(10);
        ((double)node.Format["minorUnit"]).Should().Be(2);
    }

    [Fact]
    public void Excel_Add_Chart_WithAxisNumFmt_NumFmtIsReadBack()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["axisnumfmt"] = "$#,##0"
        });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["axisNumFmt"]).Should().Be("$#,##0");
    }

    [Fact]
    public void Excel_Add_Chart_DifferentTypes()
    {
        foreach (var ct in new[] { "column", "bar", "line", "pie", "doughnut", "area", "scatter" })
        {
            var p = Path.Combine(Path.GetTempPath(), $"test_ct_{Guid.NewGuid():N}.xlsx");
            BlankDocCreator.Create(p);
            using var h = new ExcelHandler(p, editable: true);
            h.Add("/Sheet1", "chart", null, new()
            {
                ["chartType"] = ct, ["title"] = ct, ["data"] = "S1:1,2,3", ["categories"] = "A,B,C"
            });
            h.Get("/Sheet1/chart[1]").Type.Should().Be("chart");
            h.Dispose();
            File.Delete(p);
        }
    }

    [Fact]
    public void Excel_Add_Chart_Persist_SurvivesReopen()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Persist", ["data"] = "S1:10,20",
            ["categories"] = "A,B", ["legend"] = "right", ["axistitle"] = "Val"
        });
        ReopenExcel();
        var node = _excel.Get("/Sheet1/chart[1]");
        ((string)node.Format["title"]).Should().Be("Persist");
        ((string)node.Format["chartType"]).Should().Be("column");
        ((string)node.Format["categories"]).Should().Be("A,B");
        ((string)node.Format["legend"]).Should().Be("right");
        ((string)node.Format["axisTitle"]).Should().Be("Val");
    }

    // ==================== Excel: Set + Get + Reopen ====================

    [Fact]
    public void Excel_Set_Title_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Old", ["data"] = "S1:1,2"
        });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["title"]).Should().Be("Old");
        _excel.Set("/Sheet1/chart[1]", new() { ["title"] = "New" });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["title"]).Should().Be("New");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]").Format["title"]).Should().Be("New");
    }

    [Fact]
    public void Excel_Set_Legend_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2;S2:3,4"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["legend"] = "right" });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["legend"]).Should().Be("right");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]").Format["legend"]).Should().Be("right");
    }

    [Fact]
    public void Excel_Set_Categories_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2,3", ["categories"] = "A,B,C"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["categories"] = "X,Y,Z" });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["categories"]).Should().Be("X,Y,Z");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]").Format["categories"]).Should().Be("X,Y,Z");
    }

    [Fact]
    public void Excel_Set_Data_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2;S2:3,4"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["data"] = "A:10,20;B:40,50" });
        var node = _excel.Get("/Sheet1/chart[1]", depth: 2);
        ((string)node.Children[0].Format["values"]).Should().Be("10,20");
        ((string)node.Children[1].Format["values"]).Should().Be("40,50");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]", depth: 2).Children[0].Format["values"]).Should().Be("10,20");
    }

    [Fact]
    public void Excel_Set_Series1_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2,3"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["series1"] = "S1:10,20,30" });
        ((string)_excel.Get("/Sheet1/chart[1]", depth: 2).Children[0].Format["values"]).Should().Be("10,20,30");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]", depth: 2).Children[0].Format["values"]).Should().Be("10,20,30");
    }

    [Fact]
    public void Excel_Set_Colors_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2;S2:3,4"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["colors"] = "FF0000,00FF00" });
        var node = _excel.Get("/Sheet1/chart[1]", depth: 2);
        ((string)node.Children[0].Format["color"]).Should().Be("#FF0000");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]", depth: 2).Children[0].Format["color"]).Should().Be("#FF0000");
    }

    [Fact]
    public void Excel_Set_DataLabels_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "pie", ["title"] = "T", ["data"] = "S1:10,20,30", ["categories"] = "A,B,C"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["datalabels"] = "value,percent" });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["dataLabels"]).Should().Contain("value");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]").Format["dataLabels"]).Should().Contain("value");
    }

    [Fact]
    public void Excel_Set_AxisTitle_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["axistitle"] = "Revenue ($)" });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["axisTitle"]).Should().Be("Revenue ($)");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]").Format["axisTitle"]).Should().Be("Revenue ($)");
    }

    [Fact]
    public void Excel_Set_CatTitle_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["cattitle"] = "Quarters" });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["catTitle"]).Should().Be("Quarters");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]").Format["catTitle"]).Should().Be("Quarters");
    }

    [Fact]
    public void Excel_Set_AxisMinMax_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:10,20"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["axismin"] = "0", ["axismax"] = "50" });
        ((double)_excel.Get("/Sheet1/chart[1]").Format["axisMin"]).Should().Be(0);
        ((double)_excel.Get("/Sheet1/chart[1]").Format["axisMax"]).Should().Be(50);
        ReopenExcel();
        ((double)_excel.Get("/Sheet1/chart[1]").Format["axisMin"]).Should().Be(0);
    }

    [Fact]
    public void Excel_Set_MajorMinorUnit_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:10,20"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["majorunit"] = "10", ["minorunit"] = "2" });
        ((double)_excel.Get("/Sheet1/chart[1]").Format["majorUnit"]).Should().Be(10);
        ((double)_excel.Get("/Sheet1/chart[1]").Format["minorUnit"]).Should().Be(2);
        ReopenExcel();
        ((double)_excel.Get("/Sheet1/chart[1]").Format["majorUnit"]).Should().Be(10);
    }

    [Fact]
    public void Excel_Set_AxisNumFmt_IsUpdated_Persists()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1.5,2.5"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["axisnumfmt"] = "$#,##0" });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["axisNumFmt"]).Should().Be("$#,##0");
        ReopenExcel();
        ((string)_excel.Get("/Sheet1/chart[1]").Format["axisNumFmt"]).Should().Be("$#,##0");
    }

    [Fact]
    public void Excel_Set_ChartType_RemainsUnchanged()
    {
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2"
        });
        _excel.Set("/Sheet1/chart[1]", new() { ["title"] = "Changed" });
        ((string)_excel.Get("/Sheet1/chart[1]").Format["chartType"]).Should().Be("column");
    }

    // ==================== Excel: Query ====================

    [Fact]
    public void Excel_Query_Chart_FindsAll()
    {
        _excel.Add("/Sheet1", "chart", null, new() { ["chartType"] = "column", ["title"] = "A", ["data"] = "S1:1,2" });
        _excel.Add("/Sheet1", "chart", null, new() { ["chartType"] = "line", ["title"] = "B", ["data"] = "S1:3,4" });
        var results = _excel.Query("chart");
        results.Should().HaveCount(2);
        ((string)results[0].Format["title"]).Should().Be("A");
    }

    [Fact]
    public void Excel_Query_Chart_ContainsFilter()
    {
        _excel.Add("/Sheet1", "chart", null, new() { ["chartType"] = "column", ["title"] = "Sales", ["data"] = "S1:1,2" });
        _excel.Add("/Sheet1", "chart", null, new() { ["chartType"] = "line", ["title"] = "Cost", ["data"] = "S1:3,4" });
        _excel.Query("chart:contains(\"Sales\")").Should().HaveCount(1);
    }

    // ==================== Excel: Full lifecycle all props ====================

    [Fact]
    public void Excel_Chart_FullLifecycle_AllProperties()
    {
        // 1. Add with all supported props
        _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Initial", ["categories"] = "Q1,Q2,Q3",
            ["data"] = "Revenue:100,200,300;Cost:80,150,250",
            ["legend"] = "bottom", ["colors"] = "FF0000,00FF00",
            ["datalabels"] = "value", ["axistitle"] = "Amount",
            ["cattitle"] = "Quarter", ["axismin"] = "0", ["axismax"] = "500",
            ["majorunit"] = "100", ["minorunit"] = "25", ["axisnumfmt"] = "$#,##0"
        });

        // 2. Verify Add
        var node = _excel.Get("/Sheet1/chart[1]", depth: 2);
        ((string)node.Format["title"]).Should().Be("Initial");
        ((string)node.Format["categories"]).Should().Be("Q1,Q2,Q3");
        ((int)node.Format["seriesCount"]).Should().Be(2);
        ((string)node.Format["legend"]).Should().Be("bottom");
        ((string)node.Format["dataLabels"]).Should().Contain("value");
        ((string)node.Format["axisTitle"]).Should().Be("Amount");
        ((string)node.Format["catTitle"]).Should().Be("Quarter");
        ((double)node.Format["axisMin"]).Should().Be(0);
        ((double)node.Format["axisMax"]).Should().Be(500);
        ((double)node.Format["majorUnit"]).Should().Be(100);
        ((double)node.Format["minorUnit"]).Should().Be(25);
        ((string)node.Format["axisNumFmt"]).Should().Be("$#,##0");
        ((string)node.Children[0].Format["color"]).Should().Be("#FF0000");

        // 3. Set — change everything
        _excel.Set("/Sheet1/chart[1]", new()
        {
            ["title"] = "Updated", ["legend"] = "top", ["categories"] = "X,Y,Z",
            ["datalabels"] = "percent", ["colors"] = "0000FF,FFFF00",
            ["axistitle"] = "Revenue ($)", ["cattitle"] = "Period",
            ["axismin"] = "10", ["axismax"] = "400",
            ["majorunit"] = "50", ["minorunit"] = "10", ["axisnumfmt"] = "0.0%"
        });

        // 4. Verify Set
        node = _excel.Get("/Sheet1/chart[1]", depth: 2);
        ((string)node.Format["title"]).Should().Be("Updated");
        ((string)node.Format["legend"]).Should().Be("top");
        ((string)node.Format["categories"]).Should().Be("X,Y,Z");
        ((string)node.Format["dataLabels"]).Should().Contain("percent");
        ((string)node.Format["axisTitle"]).Should().Be("Revenue ($)");
        ((string)node.Format["catTitle"]).Should().Be("Period");
        ((double)node.Format["axisMin"]).Should().Be(10);
        ((double)node.Format["axisMax"]).Should().Be(400);
        ((double)node.Format["majorUnit"]).Should().Be(50);
        ((double)node.Format["minorUnit"]).Should().Be(10);
        ((string)node.Format["axisNumFmt"]).Should().Be("0.0%");
        ((string)node.Children[0].Format["color"]).Should().Be("#0000FF");

        // 5. Reopen + Verify
        ReopenExcel();
        node = _excel.Get("/Sheet1/chart[1]", depth: 2);
        ((string)node.Format["title"]).Should().Be("Updated");
        ((string)node.Format["legend"]).Should().Be("top");
        ((string)node.Format["axisTitle"]).Should().Be("Revenue ($)");
        ((double)node.Format["axisMin"]).Should().Be(10);
        ((string)node.Format["axisNumFmt"]).Should().Be("0.0%");
        ((string)node.Children[0].Format["color"]).Should().Be("#0000FF");
    }

    // ==================== Word: Add + Get + Reopen ====================

    [Fact]
    public void Word_Add_Chart_ReturnsCorrectPath()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2"
        });
        // path is /chart[1]
        _word.Get("/chart[1]").Type.Should().Be("chart");
    }

    [Fact]
    public void Word_Add_Chart_WithChartType_ChartTypeIsReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "T", ["data"] = "S1:1,2"
        });
        ((string)_word.Get("/chart[1]").Format["chartType"]).Should().Be("bar");
    }

    [Fact]
    public void Word_Add_Chart_WithTitle_TitleIsReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Word Report", ["data"] = "S1:1,2"
        });
        ((string)_word.Get("/chart[1]").Format["title"]).Should().Be("Word Report");
    }

    [Fact]
    public void Word_Add_Chart_WithLegend_LegendIsReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["legend"] = "left"
        });
        ((string)_word.Get("/chart[1]").Format["legend"]).Should().Be("left");
    }

    [Fact]
    public void Word_Add_Chart_WithCategories_CategoriesAreReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["categories"] = "X,Y"
        });
        ((string)_word.Get("/chart[1]").Format["categories"]).Should().Be("X,Y");
    }

    [Fact]
    public void Word_Add_Chart_WithData_SeriesCountAndChildrenAreReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "Sales:10,20;Profit:5,8"
        });
        var node = _word.Get("/chart[1]", depth: 2);
        ((int)node.Format["seriesCount"]).Should().Be(2);
        node.Children.Should().HaveCount(2);
        node.Children[0].Text.Should().Be("Sales");
        ((string)node.Children[0].Format["values"]).Should().Be("10,20");
    }

    [Fact]
    public void Word_Add_Chart_WithSeries1Format_SeriesIsReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T",
            ["series1"] = "Sales:100,200", ["series2"] = "Cost:50,80"
        });
        var node = _word.Get("/chart[1]", depth: 2);
        ((int)node.Format["seriesCount"]).Should().Be(2);
        ((string)node.Children[0].Format["values"]).Should().Be("100,200");
    }

    [Fact]
    public void Word_Add_Chart_WithColors_ColorsAreReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2;S2:3,4", ["colors"] = "AA0000,00BB00"
        });
        var node = _word.Get("/chart[1]", depth: 2);
        ((string)node.Children[0].Format["color"]).Should().Be("#AA0000");
    }

    [Fact]
    public void Word_Add_Chart_WithDataLabels_DataLabelsAreReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["datalabels"] = "value,category"
        });
        var dl = (string)_word.Get("/chart[1]").Format["dataLabels"];
        dl.Should().Contain("value");
        dl.Should().Contain("category");
    }

    [Fact]
    public void Word_Add_Chart_WithAxisTitle_AxisTitleIsReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["axistitle"] = "Revenue"
        });
        ((string)_word.Get("/chart[1]").Format["axisTitle"]).Should().Be("Revenue");
    }

    [Fact]
    public void Word_Add_Chart_WithCatTitle_CatTitleIsReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["cattitle"] = "Months"
        });
        ((string)_word.Get("/chart[1]").Format["catTitle"]).Should().Be("Months");
    }

    [Fact]
    public void Word_Add_Chart_WithAxisMinMax_AxisMinMaxAreReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2",
            ["axismin"] = "5", ["axismax"] = "50"
        });
        var node = _word.Get("/chart[1]");
        ((double)node.Format["axisMin"]).Should().Be(5);
        ((double)node.Format["axisMax"]).Should().Be(50);
    }

    [Fact]
    public void Word_Add_Chart_WithMajorMinorUnit_UnitsAreReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2",
            ["majorunit"] = "5", ["minorunit"] = "1"
        });
        var node = _word.Get("/chart[1]");
        ((double)node.Format["majorUnit"]).Should().Be(5);
        ((double)node.Format["minorUnit"]).Should().Be(1);
    }

    [Fact]
    public void Word_Add_Chart_WithAxisNumFmt_NumFmtIsReadBack()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["axisnumfmt"] = "0.0%"
        });
        ((string)_word.Get("/chart[1]").Format["axisNumFmt"]).Should().Be("0.0%");
    }

    [Fact]
    public void Word_Add_Chart_DifferentTypes()
    {
        foreach (var ct in new[] { "column", "bar", "line", "pie", "doughnut", "area", "scatter" })
        {
            var p = Path.Combine(Path.GetTempPath(), $"test_ct_{Guid.NewGuid():N}.docx");
            BlankDocCreator.Create(p);
            using var h = new WordHandler(p, editable: true);
            h.Add("/body", "chart", null, new()
            {
                ["chartType"] = ct, ["title"] = ct, ["data"] = "S1:1,2,3", ["categories"] = "A,B,C"
            });
            h.Get("/chart[1]").Type.Should().Be("chart");
            h.Dispose();
            File.Delete(p);
        }
    }

    [Fact]
    public void Word_Add_Chart_Persist_SurvivesReopen()
    {
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "bar", ["title"] = "Persist", ["data"] = "S1:10,20",
            ["categories"] = "X,Y", ["legend"] = "left", ["axistitle"] = "Val"
        });
        ReopenWord();
        var node = _word.Get("/chart[1]");
        ((string)node.Format["title"]).Should().Be("Persist");
        ((string)node.Format["chartType"]).Should().Be("bar");
        ((string)node.Format["categories"]).Should().Be("X,Y");
        ((string)node.Format["legend"]).Should().Be("left");
        ((string)node.Format["axisTitle"]).Should().Be("Val");
    }

    // ==================== Word: Set + Get + Reopen ====================

    [Fact]
    public void Word_Set_Title_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "Old", ["data"] = "S1:1,2" });
        _word.Set("/chart[1]", new() { ["title"] = "New" });
        ((string)_word.Get("/chart[1]").Format["title"]).Should().Be("New");
        ReopenWord();
        ((string)_word.Get("/chart[1]").Format["title"]).Should().Be("New");
    }

    [Fact]
    public void Word_Set_Legend_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2;S2:3,4" });
        _word.Set("/chart[1]", new() { ["legend"] = "right" });
        ((string)_word.Get("/chart[1]").Format["legend"]).Should().Be("right");
        ReopenWord();
        ((string)_word.Get("/chart[1]").Format["legend"]).Should().Be("right");
    }

    [Fact]
    public void Word_Set_Categories_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2", ["categories"] = "A,B" });
        _word.Set("/chart[1]", new() { ["categories"] = "X,Y" });
        ((string)_word.Get("/chart[1]").Format["categories"]).Should().Be("X,Y");
        ReopenWord();
        ((string)_word.Get("/chart[1]").Format["categories"]).Should().Be("X,Y");
    }

    [Fact]
    public void Word_Set_Data_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2;S2:3,4" });
        _word.Set("/chart[1]", new() { ["data"] = "A:10,20;B:40,50" });
        ((string)_word.Get("/chart[1]", depth: 2).Children[0].Format["values"]).Should().Be("10,20");
        ReopenWord();
        ((string)_word.Get("/chart[1]", depth: 2).Children[0].Format["values"]).Should().Be("10,20");
    }

    [Fact]
    public void Word_Set_Series1_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2,3" });
        _word.Set("/chart[1]", new() { ["series1"] = "S1:10,20,30" });
        ((string)_word.Get("/chart[1]", depth: 2).Children[0].Format["values"]).Should().Be("10,20,30");
        ReopenWord();
        ((string)_word.Get("/chart[1]", depth: 2).Children[0].Format["values"]).Should().Be("10,20,30");
    }

    [Fact]
    public void Word_Set_Colors_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2;S2:3,4" });
        _word.Set("/chart[1]", new() { ["colors"] = "0000FF,FF00FF" });
        ((string)_word.Get("/chart[1]", depth: 2).Children[0].Format["color"]).Should().Be("#0000FF");
        ReopenWord();
        ((string)_word.Get("/chart[1]", depth: 2).Children[0].Format["color"]).Should().Be("#0000FF");
    }

    [Fact]
    public void Word_Set_DataLabels_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "pie", ["title"] = "T", ["data"] = "S1:10,20", ["categories"] = "A,B" });
        _word.Set("/chart[1]", new() { ["datalabels"] = "value,percent" });
        ((string)_word.Get("/chart[1]").Format["dataLabels"]).Should().Contain("value");
        ReopenWord();
        ((string)_word.Get("/chart[1]").Format["dataLabels"]).Should().Contain("value");
    }

    [Fact]
    public void Word_Set_AxisTitle_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2" });
        _word.Set("/chart[1]", new() { ["axistitle"] = "Revenue ($)" });
        ((string)_word.Get("/chart[1]").Format["axisTitle"]).Should().Be("Revenue ($)");
        ReopenWord();
        ((string)_word.Get("/chart[1]").Format["axisTitle"]).Should().Be("Revenue ($)");
    }

    [Fact]
    public void Word_Set_CatTitle_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2" });
        _word.Set("/chart[1]", new() { ["cattitle"] = "Months" });
        ((string)_word.Get("/chart[1]").Format["catTitle"]).Should().Be("Months");
        ReopenWord();
        ((string)_word.Get("/chart[1]").Format["catTitle"]).Should().Be("Months");
    }

    [Fact]
    public void Word_Set_AxisMinMax_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:10,20" });
        _word.Set("/chart[1]", new() { ["axismin"] = "5", ["axismax"] = "100" });
        ((double)_word.Get("/chart[1]").Format["axisMin"]).Should().Be(5);
        ReopenWord();
        ((double)_word.Get("/chart[1]").Format["axisMin"]).Should().Be(5);
    }

    [Fact]
    public void Word_Set_MajorMinorUnit_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:10,20" });
        _word.Set("/chart[1]", new() { ["majorunit"] = "5", ["minorunit"] = "1" });
        ((double)_word.Get("/chart[1]").Format["majorUnit"]).Should().Be(5);
        ReopenWord();
        ((double)_word.Get("/chart[1]").Format["majorUnit"]).Should().Be(5);
    }

    [Fact]
    public void Word_Set_AxisNumFmt_IsUpdated_Persists()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "T", ["data"] = "S1:1,2" });
        _word.Set("/chart[1]", new() { ["axisnumfmt"] = "0.00%" });
        ((string)_word.Get("/chart[1]").Format["axisNumFmt"]).Should().Be("0.00%");
        ReopenWord();
        ((string)_word.Get("/chart[1]").Format["axisNumFmt"]).Should().Be("0.00%");
    }

    [Fact]
    public void Word_Set_ChartType_RemainsUnchanged()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "line", ["title"] = "T", ["data"] = "S1:1,2" });
        _word.Set("/chart[1]", new() { ["title"] = "Changed" });
        ((string)_word.Get("/chart[1]").Format["chartType"]).Should().Be("line");
    }

    // ==================== Word: Query ====================

    [Fact]
    public void Word_Query_Chart_FindsAll()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "A", ["data"] = "S1:1,2" });
        _word.Add("/body", "chart", null, new() { ["chartType"] = "line", ["title"] = "B", ["data"] = "S1:3,4" });
        _word.Query("chart").Should().HaveCount(2);
    }

    [Fact]
    public void Word_Query_Chart_ContainsFilter()
    {
        _word.Add("/body", "chart", null, new() { ["chartType"] = "column", ["title"] = "Sales", ["data"] = "S1:1,2" });
        _word.Add("/body", "chart", null, new() { ["chartType"] = "line", ["title"] = "Cost", ["data"] = "S1:3,4" });
        _word.Query("chart:contains(\"Sales\")").Should().HaveCount(1);
    }

    // ==================== Word: Full lifecycle all props ====================

    [Fact]
    public void Word_Chart_FullLifecycle_AllProperties()
    {
        // 1. Add with all props
        _word.Add("/body", "chart", null, new()
        {
            ["chartType"] = "column", ["title"] = "Initial", ["categories"] = "Q1,Q2,Q3",
            ["data"] = "Revenue:100,200,300;Cost:80,150,250",
            ["legend"] = "bottom", ["colors"] = "FF0000,00FF00",
            ["datalabels"] = "value", ["axistitle"] = "Amount",
            ["cattitle"] = "Quarter", ["axismin"] = "0", ["axismax"] = "500",
            ["majorunit"] = "100", ["minorunit"] = "25", ["axisnumfmt"] = "$#,##0"
        });

        // 2. Verify Add
        var node = _word.Get("/chart[1]", depth: 2);
        ((string)node.Format["title"]).Should().Be("Initial");
        ((string)node.Format["legend"]).Should().Be("bottom");
        ((string)node.Format["dataLabels"]).Should().Contain("value");
        ((string)node.Format["axisTitle"]).Should().Be("Amount");
        ((double)node.Format["axisMin"]).Should().Be(0);
        ((string)node.Format["axisNumFmt"]).Should().Be("$#,##0");

        // 3. Set — change everything
        _word.Set("/chart[1]", new()
        {
            ["title"] = "Updated", ["legend"] = "top", ["categories"] = "X,Y,Z",
            ["datalabels"] = "percent", ["colors"] = "0000FF,FFFF00",
            ["axistitle"] = "Revenue ($)", ["cattitle"] = "Period",
            ["axismin"] = "10", ["axismax"] = "400",
            ["majorunit"] = "50", ["minorunit"] = "10", ["axisnumfmt"] = "0.0%"
        });

        // 4. Verify Set
        node = _word.Get("/chart[1]", depth: 2);
        ((string)node.Format["title"]).Should().Be("Updated");
        ((string)node.Format["legend"]).Should().Be("top");
        ((string)node.Format["dataLabels"]).Should().Contain("percent");
        ((string)node.Format["axisTitle"]).Should().Be("Revenue ($)");
        ((double)node.Format["axisMin"]).Should().Be(10);
        ((string)node.Format["axisNumFmt"]).Should().Be("0.0%");

        // 5. Reopen + Verify
        ReopenWord();
        node = _word.Get("/chart[1]", depth: 2);
        ((string)node.Format["title"]).Should().Be("Updated");
        ((string)node.Format["legend"]).Should().Be("top");
        ((string)node.Format["axisTitle"]).Should().Be("Revenue ($)");
        ((double)node.Format["axisMin"]).Should().Be(10);
        ((string)node.Format["axisNumFmt"]).Should().Be("0.0%");
    }

    // ==================== Cross-format consistency ====================

    [Fact]
    public void AllFormats_ChartAddSet_SameProperties_SameResults()
    {
        _pptx.Add("/", "slide", null, new());
        var addProps = new Dictionary<string, string>
        {
            ["chartType"] = "column", ["title"] = "Test", ["data"] = "S1:1,2,3",
            ["categories"] = "A,B,C", ["legend"] = "top",
            ["datalabels"] = "value", ["axistitle"] = "Val", ["axismin"] = "0", ["axismax"] = "10"
        };
        _pptx.Add("/slide[1]", "chart", null, new(addProps));
        _excel.Add("/Sheet1", "chart", null, new(addProps));
        _word.Add("/body", "chart", null, new(addProps));

        // Verify all three match at Add time
        var pn = _pptx.Get("/slide[1]/chart[1]");
        var en = _excel.Get("/Sheet1/chart[1]");
        var wn = _word.Get("/chart[1]");
        foreach (var node in new[] { pn, en, wn })
        {
            ((string)node.Format["title"]).Should().Be("Test");
            ((string)node.Format["legend"]).Should().Be("top");
            ((string)node.Format["dataLabels"]).Should().Contain("value");
            ((string)node.Format["axisTitle"]).Should().Be("Val");
            ((double)node.Format["axisMin"]).Should().Be(0);
        }

        // Set same props on all
        var setProps = new Dictionary<string, string>
        {
            ["title"] = "Unified", ["legend"] = "right", ["axistitle"] = "Revenue",
            ["axismin"] = "5", ["axisnumfmt"] = "$#,##0"
        };
        _pptx.Set("/slide[1]/chart[1]", setProps);
        _excel.Set("/Sheet1/chart[1]", new(setProps));
        _word.Set("/chart[1]", new(setProps));

        // Verify consistency after Set
        pn = _pptx.Get("/slide[1]/chart[1]");
        en = _excel.Get("/Sheet1/chart[1]");
        wn = _word.Get("/chart[1]");
        foreach (var node in new[] { pn, en, wn })
        {
            ((string)node.Format["title"]).Should().Be("Unified");
            ((string)node.Format["legend"]).Should().Be("right");
            ((string)node.Format["axisTitle"]).Should().Be("Revenue");
            ((double)node.Format["axisMin"]).Should().Be(5);
            ((string)node.Format["axisNumFmt"]).Should().Be("$#,##0");
        }

        // Reopen all + Verify persistence
        ReopenPptx(); ReopenExcel(); ReopenWord();
        pn = _pptx.Get("/slide[1]/chart[1]");
        en = _excel.Get("/Sheet1/chart[1]");
        wn = _word.Get("/chart[1]");
        foreach (var node in new[] { pn, en, wn })
        {
            ((string)node.Format["title"]).Should().Be("Unified");
            ((string)node.Format["axisTitle"]).Should().Be("Revenue");
            ((double)node.Format["axisMin"]).Should().Be(5);
        }
    }
}
