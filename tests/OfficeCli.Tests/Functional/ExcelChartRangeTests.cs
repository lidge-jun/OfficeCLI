// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for chart series cell references in Excel.
/// Dotted syntax: series1.name=Sales, series1.values=Sheet1!B2:B13, series1.categories=Sheet1!A2:A13
/// </summary>
public class ExcelChartRangeTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelChartRangeTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Cell Reference Series ====================

    [Fact]
    public void Add_ChartWithValuesRef_StoresNumberReference()
    {
        // 1. Add some data to cells
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Q1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Q2" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "100" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "200" });

        // 2. Add chart with cell references
        var path = _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Sales by Quarter",
            ["series1.name"] = "Revenue",
            ["series1.values"] = "Sheet1!B1:B2"
        });
        path.Should().Be("/Sheet1/chart[1]");

        // 3. Get chart and verify
        var chart = _handler.Get("/Sheet1/chart[1]", depth: 2);
        chart.Type.Should().Be("chart");
        ((string)chart.Format["title"]).Should().Be("Sales by Quarter");
        chart.Children.Should().HaveCount(1);
        var series = chart.Children[0];
        series.Format["name"].Should().Be("Revenue");
        ((string)series.Format["valuesRef"]).Should().Be("Sheet1!$B$1:$B$2");
    }

    [Fact]
    public void Add_ChartWithCategoriesRef_StoresStringReference()
    {
        // 1. Add data
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Jan" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Feb" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "10" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "20" });

        // 2. Add chart with both values and categories references
        _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Monthly Sales",
            ["series1.name"] = "Sales",
            ["series1.values"] = "Sheet1!B1:B2",
            ["series1.categories"] = "Sheet1!A1:A2"
        });

        // 3. Get chart
        var chart = _handler.Get("/Sheet1/chart[1]", depth: 2);
        chart.Children.Should().HaveCount(1);
        var series = chart.Children[0];
        ((string)series.Format["valuesRef"]).Should().Be("Sheet1!$B$1:$B$2");
        ((string)series.Format["categoriesRef"]).Should().Be("Sheet1!$A$1:$A$2");
    }

    [Fact]
    public void Add_ChartWithTopLevelCategoriesRef_AppliedToAllSeries()
    {
        // Data
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Q1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Q2" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "10" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "20" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "C1", ["value"] = "30" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "C2", ["value"] = "40" });

        // Add chart with top-level categories reference and two series
        _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "Sheet1!A1:A2",
            ["series1.name"] = "Revenue",
            ["series1.values"] = "Sheet1!B1:B2",
            ["series2.name"] = "Cost",
            ["series2.values"] = "Sheet1!C1:C2"
        });

        var chart = _handler.Get("/Sheet1/chart[1]", depth: 2);
        chart.Children.Should().HaveCount(2);

        // Both series should have categoriesRef from top-level
        var s1 = chart.Children[0];
        ((string)s1.Format["valuesRef"]).Should().Be("Sheet1!$B$1:$B$2");
        ((string)s1.Format["categoriesRef"]).Should().Be("Sheet1!$A$1:$A$2");

        var s2 = chart.Children[1];
        ((string)s2.Format["valuesRef"]).Should().Be("Sheet1!$C$1:$C$2");
        ((string)s2.Format["categoriesRef"]).Should().Be("Sheet1!$A$1:$A$2");

        // Chart-level categoriesRef should also be set
        ((string)chart.Format["categoriesRef"]).Should().Be("Sheet1!$A$1:$A$2");
    }

    [Fact]
    public void Add_ChartWithRef_Persistence()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Jan" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "100" });

        _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Persisted Chart",
            ["series1.name"] = "Sales",
            ["series1.values"] = "Sheet1!B1:B1",
            ["series1.categories"] = "Sheet1!A1:A1"
        });

        // Reopen and verify references survive
        Reopen();
        var chart = _handler.Get("/Sheet1/chart[1]", depth: 2);
        ((string)chart.Format["title"]).Should().Be("Persisted Chart");
        chart.Children.Should().HaveCount(1);
        var series = chart.Children[0];
        ((string)series.Format["valuesRef"]).Should().Be("Sheet1!$B$1:$B$1");
        ((string)series.Format["categoriesRef"]).Should().Be("Sheet1!$A$1:$A$1");
    }

    [Fact]
    public void Add_ChartWithRef_AlreadyAbsolute_PreservesDollarSigns()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "50" });

        _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["series1.name"] = "Test",
            ["series1.values"] = "Sheet1!$B$1:$B$1"
        });

        var chart = _handler.Get("/Sheet1/chart[1]", depth: 2);
        ((string)chart.Children[0].Format["valuesRef"]).Should().Be("Sheet1!$B$1:$B$1");
    }

    [Fact]
    public void Add_ChartWithRef_NoSheetPrefix_UsesRawReference()
    {
        // Range without sheet prefix — should still be treated as reference
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "50" });

        _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["series1.name"] = "Test",
            ["series1.values"] = "B1:B2"
        });

        var chart = _handler.Get("/Sheet1/chart[1]", depth: 2);
        // Should be normalized to absolute references
        ((string)chart.Children[0].Format["valuesRef"]).Should().Be("$B$1:$B$2");
    }

    // ==================== Legacy Format Still Works ====================

    [Fact]
    public void Add_LegacySeriesFormat_StillWorks()
    {
        _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["series1"] = "Sales:10,20,30"
        });

        var chart = _handler.Get("/Sheet1/chart[1]", depth: 2);
        chart.Children.Should().HaveCount(1);
        var series = chart.Children[0];
        ((string)series.Format["name"]).Should().Be("Sales");
        ((string)series.Format["values"]).Should().Be("10,20,30");
        series.Format.Should().NotContainKey("valuesRef");
    }

    [Fact]
    public void Add_DottedSyntaxWithLiteralValues_StillWorks()
    {
        _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["series1.name"] = "Revenue",
            ["series1.values"] = "10,20,30"
        });

        var chart = _handler.Get("/Sheet1/chart[1]", depth: 2);
        chart.Children.Should().HaveCount(1);
        var series = chart.Children[0];
        ((string)series.Format["name"]).Should().Be("Revenue");
        ((string)series.Format["values"]).Should().Be("10,20,30");
        series.Format.Should().NotContainKey("valuesRef");
    }

    // ==================== Line Chart with References ====================

    [Fact]
    public void Add_LineChartWithRefs_StoresReferences()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Jan" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Feb" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "10" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "20" });

        _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Line Ref Chart",
            ["series1.name"] = "Trend",
            ["series1.values"] = "Sheet1!B1:B2",
            ["series1.categories"] = "Sheet1!A1:A2"
        });

        var chart = _handler.Get("/Sheet1/chart[1]", depth: 2);
        ((string)chart.Format["chartType"]).Should().Be("line");
        var series = chart.Children[0];
        ((string)series.Format["valuesRef"]).Should().Be("Sheet1!$B$1:$B$2");
        ((string)series.Format["categoriesRef"]).Should().Be("Sheet1!$A$1:$A$2");
    }

    // ==================== IsRangeReference Helper Tests ====================

    [Theory]
    [InlineData("Sheet1!B2:B13", true)]
    [InlineData("Sheet1!$B$2:$B$13", true)]
    [InlineData("B2:B13", true)]
    [InlineData("$A$1:$Z$100", true)]
    [InlineData("AA1:ZZ999", true)]
    [InlineData("10,20,30", false)]
    [InlineData("Sales", false)]
    [InlineData("", false)]
    public void IsRangeReference_DetectsCorrectly(string value, bool expected)
    {
        OfficeCli.Core.ChartHelper.IsRangeReference(value).Should().Be(expected);
    }

    [Theory]
    [InlineData("Sheet1!B2:B13", "Sheet1!$B$2:$B$13")]
    [InlineData("Sheet1!$B$2:$B$13", "Sheet1!$B$2:$B$13")]
    [InlineData("B2:B13", "$B$2:$B$13")]
    public void NormalizeRangeReference_AddsAbsoluteMarkers(string input, string expected)
    {
        OfficeCli.Core.ChartHelper.NormalizeRangeReference(input).Should().Be(expected);
    }
}
