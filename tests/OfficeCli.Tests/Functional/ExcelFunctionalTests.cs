// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for XLSX: each test creates a blank file, adds elements,
/// queries them, and modifies them — exercising the full Create→Add→Get→Set lifecycle.
/// </summary>
public class ExcelFunctionalTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelFunctionalTests()
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

    // Reopen the file to verify persistence
    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Sheet lifecycle ====================

    [Fact]
    public void BlankFile_HasSheet1()
    {
        var node = _handler.Get("/");
        node.Children.Should().Contain(c => c.Path == "/Sheet1");
    }

    [Fact]
    public void AddSheet_ReturnsPath()
    {
        var path = _handler.Add("/", "sheet", null,
            new Dictionary<string, string> { ["name"] = "Sales" });
        path.Should().Be("/Sales");
    }

    [Fact]
    public void AddSheet_Get_ReturnsSheetType()
    {
        _handler.Add("/", "sheet", null, new Dictionary<string, string> { ["name"] = "Report" });
        var node = _handler.Get("/Report");
        node.Type.Should().Be("sheet");
    }

    [Fact]
    public void AddSheet_Multiple_AllVisible()
    {
        _handler.Add("/", "sheet", null, new Dictionary<string, string> { ["name"] = "Alpha" });
        _handler.Add("/", "sheet", null, new Dictionary<string, string> { ["name"] = "Beta" });

        var root = _handler.Get("/");
        var sheetPaths = root.Children.Select(c => c.Path).ToList();
        sheetPaths.Should().Contain("/Alpha");
        sheetPaths.Should().Contain("/Beta");
    }

    // ==================== Cell lifecycle ====================

    [Fact]
    public void AddCell_NumberValue_TextIsReadBack()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "42" });

        var node = _handler.Get("/Sheet1/A1");
        node.Type.Should().Be("cell");
        node.Text.Should().Be("42");
    }

    [Fact]
    public void AddCell_StringValue_TextIsReadBack()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "B2", ["value"] = "Hello", ["type"] = "string" });

        var node = _handler.Get("/Sheet1/B2");
        node.Text.Should().Be("Hello");
    }

    [Fact]
    public void AddCell_Formula_FormulaIsReadBack()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "C1", ["formula"] = "A1+B1" });

        var node = _handler.Get("/Sheet1/C1");
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].Should().Be("A1+B1");
    }

    [Fact]
    public void AddCell_MultipleCells_AllReadBack()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "10" });
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "B1", ["value"] = "20" });
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "C1", ["value"] = "30" });

        _handler.Get("/Sheet1/A1").Text.Should().Be("10");
        _handler.Get("/Sheet1/B1").Text.Should().Be("20");
        _handler.Get("/Sheet1/C1").Text.Should().Be("30");
    }

    // ==================== Set: modify cell properties ====================

    [Fact]
    public void SetCell_Value_ValueIsUpdated()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "old" });

        _handler.Set("/Sheet1/A1", new Dictionary<string, string> { ["value"] = "new" });

        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("new");
    }

    [Fact]
    public void SetCell_Bold_DoesNotThrow()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "100" });

        var act = () => _handler.Set("/Sheet1/A1",
            new Dictionary<string, string> { ["font.bold"] = "true" });
        act.Should().NotThrow();

        // Cell value should still be intact
        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("100");
    }

    [Fact]
    public void SetCell_Fill_DoesNotThrow()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "styled" });

        var act = () => _handler.Set("/Sheet1/A1",
            new Dictionary<string, string> { ["fill"] = "4472C4" });
        act.Should().NotThrow();
    }

    [Fact]
    public void SetCell_NumFmt_DoesNotThrow()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "0.5" });

        var act = () => _handler.Set("/Sheet1/A1",
            new Dictionary<string, string> { ["numFmt"] = "0.00%" });
        act.Should().NotThrow();
    }

    [Fact]
    public void SetCell_Alignment_DoesNotThrow()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "text" });

        var act = () => _handler.Set("/Sheet1/A1",
            new Dictionary<string, string> { ["alignment.horizontal"] = "center" });
        act.Should().NotThrow();
    }

    // ==================== Query ====================

    [Fact]
    public void GetSheet_WithCells_ChildrenContainCells()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "1" });
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "B1", ["value"] = "2" });

        var sheet = _handler.Get("/Sheet1", depth: 2);
        var allCells = sheet.Children.SelectMany(row => row.Children).ToList();
        allCells.Should().HaveCountGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void GetRange_ReturnsAllCellsInRange()
    {
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "1" });
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "B1", ["value"] = "2" });
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "C1", ["value"] = "3" });

        var range = _handler.Get("/Sheet1/A1:C1");
        range.Children.Should().HaveCount(3);
    }

    [Fact]
    public void GetRange_ChildrenHaveCorrectValues()
    {
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "10" });
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string> { ["ref"] = "B1", ["value"] = "20" });

        var range = _handler.Get("/Sheet1/A1:B1");
        var values = range.Children.Select(c => c.Text).ToList();
        values.Should().Contain("10");
        values.Should().Contain("20");
    }

    // ==================== Persistence ====================

    [Fact]
    public void AddCell_Persist_SurvivesReopenFile()
    {
        _handler.Add("/Sheet1", "cell", null,
            new Dictionary<string, string> { ["ref"] = "A1", ["value"] = "persistent" });

        Reopen();
        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("persistent");
    }

    [Fact]
    public void AddSheet_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "sheet", null, new Dictionary<string, string> { ["name"] = "Saved" });

        Reopen();
        var root = _handler.Get("/");
        root.Children.Should().Contain(c => c.Path == "/Saved");
    }

    // ==================== Row lifecycle ====================

    [Fact]
    public void AddRow_RowIsQueryable()
    {
        _handler.Add("/Sheet1", "row", null,
            new Dictionary<string, string> { ["cols"] = "3" });

        var sheet = _handler.Get("/Sheet1", depth: 1);
        sheet.Children.Should().HaveCountGreaterThanOrEqualTo(1);
        sheet.Children.Should().Contain(c => c.Type == "row");
    }

    // ==================== XLSX Hyperlinks ====================

    [Fact]
    public void CellLink_Lifecycle()
    {
        // 1. Set cell value + link
        _handler.Set("/Sheet1/A1", new Dictionary<string, string>
        {
            ["value"] = "Visit us",
            ["link"] = "https://first.com"
        });

        // 2. Get + Verify
        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("Visit us");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://first.com");

        // 3. Set updated link + Verify
        _handler.Set("/Sheet1/A1", new Dictionary<string, string> { ["link"] = "https://updated.com" });
        node = _handler.Get("/Sheet1/A1");
        ((string)node.Format["link"]).Should().StartWith("https://updated.com");

        // 4. Remove link + Verify
        _handler.Set("/Sheet1/A1", new Dictionary<string, string> { ["link"] = "none" });
        node = _handler.Get("/Sheet1/A1");
        node.Format.Should().NotContainKey("link");
    }

    [Fact]
    public void CellLink_Persist_SurvivesReopenFile()
    {
        _handler.Set("/Sheet1/B1", new Dictionary<string, string>
        {
            ["value"] = "Link cell",
            ["link"] = "https://original.com"
        });
        _handler.Set("/Sheet1/B1", new Dictionary<string, string> { ["link"] = "https://persist.com" });

        var handler2 = Reopen();
        var node = handler2.Get("/Sheet1/B1");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://persist.com");
    }

    // ==================== Border Lifecycle ====================

    [Fact]
    public void Border_FullLifecycle()
    {
        // 1. Add cell
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Bordered" });

        // 2. Set border
        _handler.Set("/Sheet1/A1", new() { ["border.all"] = "thin", ["border.color"] = "000000" });

        // 3. Get + Verify borders readable
        var node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("Bordered");
        node.Format.Should().ContainKey("border.left");
        ((string)node.Format["border.left"]).Should().Be("thin");

        // 4. Set (modify border)
        _handler.Set("/Sheet1/A1", new() { ["border.bottom"] = "thick" });

        // 5. Get + Verify modification
        node = _handler.Get("/Sheet1/A1");
        ((string)node.Format["border.bottom"]).Should().Be("thick");

        // 6. Persistence
        Reopen();
        node = _handler.Get("/Sheet1/A1");
        node.Text.Should().Be("Bordered");
        node.Format.Should().ContainKey("border.left");
    }

    // ==================== Merge Cells Lifecycle ====================

    [Fact]
    public void MergeCells_FullLifecycle()
    {
        // 1. Add cells
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Merged" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "" });

        // 2. Merge
        _handler.Set("/Sheet1/A1:D1", new() { ["merge"] = "true" });

        // 3. Get + Verify merge info
        var cell = _handler.Get("/Sheet1/A1");
        cell.Format.Should().ContainKey("merge");
        ((string)cell.Format["merge"]).Should().Be("A1:D1");

        // 4. Persistence
        Reopen();
        cell = _handler.Get("/Sheet1/A1");
        cell.Format.Should().ContainKey("merge");

        // 5. Unmerge
        _handler.Set("/Sheet1/A1:D1", new() { ["merge"] = "false" });
        cell = _handler.Get("/Sheet1/A1");
        cell.Format.Should().NotContainKey("merge");
    }

    // ==================== Column Width Lifecycle ====================

    [Fact]
    public void ColumnWidth_FullLifecycle()
    {
        // 1. Add data
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Wide column" });

        // 2. Set column width
        _handler.Set("/Sheet1/col[A]", new() { ["width"] = "25" });

        // 3. Get + Verify
        var col = _handler.Get("/Sheet1/col[A]");
        col.Type.Should().Be("column");
        ((double)col.Format["width"]).Should().Be(25);

        // 4. Set (modify)
        _handler.Set("/Sheet1/col[A]", new() { ["width"] = "30" });

        // 5. Get + Verify
        col = _handler.Get("/Sheet1/col[A]");
        ((double)col.Format["width"]).Should().Be(30);

        // 6. Persistence
        Reopen();
        col = _handler.Get("/Sheet1/col[A]");
        ((double)col.Format["width"]).Should().Be(30);
    }

    // ==================== Row Height Lifecycle ====================

    [Fact]
    public void RowHeight_FullLifecycle()
    {
        // 1. Add data
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Tall row" });

        // 2. Set row height
        _handler.Set("/Sheet1/row[1]", new() { ["height"] = "30" });

        // 3. Get + Verify
        var row = _handler.Get("/Sheet1/row[1]");
        row.Type.Should().Be("row");
        ((double)row.Format["height"]).Should().Be(30);

        // 4. Set (modify)
        _handler.Set("/Sheet1/row[1]", new() { ["height"] = "40" });

        // 5. Get + Verify
        row = _handler.Get("/Sheet1/row[1]");
        ((double)row.Format["height"]).Should().Be(40);

        // 6. Persistence
        Reopen();
        row = _handler.Get("/Sheet1/row[1]");
        ((double)row.Format["height"]).Should().Be(40);
    }

    // ==================== Freeze Panes Lifecycle ====================

    [Fact]
    public void FreezePanes_FullLifecycle()
    {
        // 1. Add data
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Header" });

        // 2. Set freeze (freeze row 1)
        _handler.Set("/Sheet1", new() { ["freeze"] = "A2" });

        // 3. Get + Verify
        var sheet = _handler.Get("/Sheet1");
        sheet.Format.Should().ContainKey("freeze");
        ((string)sheet.Format["freeze"]).Should().Be("A2");

        // 4. Set (modify freeze)
        _handler.Set("/Sheet1", new() { ["freeze"] = "B3" });
        sheet = _handler.Get("/Sheet1");
        ((string)sheet.Format["freeze"]).Should().Be("B3");

        // 5. Persistence
        Reopen();
        sheet = _handler.Get("/Sheet1");
        ((string)sheet.Format["freeze"]).Should().Be("B3");
    }

    // ==================== AutoFilter Lifecycle ====================

    [Fact]
    public void AutoFilter_FullLifecycle()
    {
        // 1. Add data
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Name" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "Score" });

        // 2. Add autofilter
        _handler.Add("/Sheet1", "autofilter", null, new() { ["range"] = "A1:B10" });

        // 3. Get + Verify
        var sheet = _handler.Get("/Sheet1");
        sheet.Format.Should().ContainKey("autoFilter");
        ((string)sheet.Format["autoFilter"]).Should().Be("A1:B10");

        // 4. Set (modify range)
        _handler.Set("/Sheet1/autofilter", new() { ["range"] = "A1:B20" });
        sheet = _handler.Get("/Sheet1");
        ((string)sheet.Format["autoFilter"]).Should().Be("A1:B20");

        // 5. Persistence
        Reopen();
        sheet = _handler.Get("/Sheet1");
        ((string)sheet.Format["autoFilter"]).Should().Be("A1:B20");
    }

    // ==================== ColorScale Lifecycle ====================

    [Fact]
    public void ColorScale_FullLifecycle()
    {
        // 1. Add data
        for (int i = 1; i <= 5; i++)
            _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = $"A{i}", ["value"] = $"{i * 20}" });

        // 2. Add colorscale
        var path = _handler.Add("/Sheet1", "colorscale", null, new()
        {
            ["sqref"] = "A1:A5", ["mincolor"] = "F8696B", ["maxcolor"] = "63BE7B"
        });
        path.Should().Be("/Sheet1/cf[1]");

        // 3. Get + Verify
        var cf = _handler.Get("/Sheet1/cf[1]");
        cf.Type.Should().Be("conditionalFormatting");
        ((string)cf.Format["cfType"]).Should().Be("colorScale");
        ((string)cf.Format["sqref"]).Should().Be("A1:A5");
        ((string)cf.Format["mincolor"]).Should().Contain("#F8696B");
        ((string)cf.Format["maxcolor"]).Should().Contain("#63BE7B");

        // 4. Set (modify colors)
        _handler.Set("/Sheet1/cf[1]", new() { ["mincolor"] = "0000FF", ["maxcolor"] = "FF0000" });

        // 5. Get + Verify
        cf = _handler.Get("/Sheet1/cf[1]");
        ((string)cf.Format["mincolor"]).Should().Contain("#0000FF");
        ((string)cf.Format["maxcolor"]).Should().Contain("#FF0000");

        // 6. Persistence
        Reopen();
        cf = _handler.Get("/Sheet1/cf[1]");
        ((string)cf.Format["cfType"]).Should().Be("colorScale");
        ((string)cf.Format["mincolor"]).Should().Contain("#0000FF");
    }

    // ==================== IconSet Lifecycle ====================

    [Fact]
    public void IconSet_FullLifecycle()
    {
        // 1. Add data
        for (int i = 1; i <= 5; i++)
            _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = $"A{i}", ["value"] = $"{i * 20}" });

        // 2. Add iconset
        var path = _handler.Add("/Sheet1", "iconset", null, new()
        {
            ["sqref"] = "A1:A5", ["iconset"] = "3Arrows"
        });
        path.Should().Be("/Sheet1/cf[1]");

        // 3. Get + Verify
        var cf = _handler.Get("/Sheet1/cf[1]");
        cf.Type.Should().Be("conditionalFormatting");
        ((string)cf.Format["cfType"]).Should().Be("iconSet");
        ((string)cf.Format["iconset"]).Should().Be("3Arrows");

        // 4. Set (modify)
        _handler.Set("/Sheet1/cf[1]", new() { ["iconset"] = "3TrafficLights1", ["reverse"] = "true" });

        // 5. Get + Verify
        cf = _handler.Get("/Sheet1/cf[1]");
        ((string)cf.Format["iconset"]).Should().Be("3TrafficLights1");
        ((bool)cf.Format["reverse"]).Should().BeTrue();

        // 6. Persistence
        Reopen();
        cf = _handler.Get("/Sheet1/cf[1]");
        ((string)cf.Format["cfType"]).Should().Be("iconSet");
        ((string)cf.Format["iconset"]).Should().Be("3TrafficLights1");
    }

    // ==================== Formula CF Lifecycle ====================

    [Fact]
    public void FormulaCF_FullLifecycle()
    {
        // 1. Add data
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "200" });

        // 2. Add formula CF
        var path = _handler.Add("/Sheet1", "formulacf", null, new()
        {
            ["sqref"] = "A1:A10", ["formula"] = "$A1>100", ["fill"] = "FF0000"
        });
        path.Should().Be("/Sheet1/cf[1]");

        // 3. Get + Verify
        var cf = _handler.Get("/Sheet1/cf[1]");
        cf.Type.Should().Be("conditionalFormatting");
        ((string)cf.Format["cfType"]).Should().Be("formula");
        ((string)cf.Format["formula"]).Should().Be("$A1>100");
        ((string)cf.Format["sqref"]).Should().Be("A1:A10");

        // 4. Set (modify range)
        _handler.Set("/Sheet1/cf[1]", new() { ["sqref"] = "A1:A20" });

        // 5. Get + Verify
        cf = _handler.Get("/Sheet1/cf[1]");
        ((string)cf.Format["sqref"]).Should().Be("A1:A20");

        // 6. Persistence
        Reopen();
        cf = _handler.Get("/Sheet1/cf[1]");
        ((string)cf.Format["cfType"]).Should().Be("formula");
        ((string)cf.Format["sqref"]).Should().Be("A1:A20");
    }

    // ==================== Chart Lifecycle ====================

    [Fact]
    public void Chart_FullLifecycle()
    {
        // 1. Add data
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Q1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Q2" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "100" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "200" });

        // 2. Add chart
        var path = _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Sales",
            ["categories"] = "Q1,Q2",
            ["data"] = "Revenue:100,200"
        });
        path.Should().Be("/Sheet1/chart[1]");

        // 3. Get + Verify
        var chart = _handler.Get("/Sheet1/chart[1]");
        chart.Type.Should().Be("chart");
        ((string)chart.Format["title"]).Should().Be("Sales");
        ((string)chart.Format["chartType"]).Should().Be("column");

        // 4. Set (modify title)
        _handler.Set("/Sheet1/chart[1]", new() { ["title"] = "Updated Sales" });

        // 5. Get + Verify
        chart = _handler.Get("/Sheet1/chart[1]");
        ((string)chart.Format["title"]).Should().Be("Updated Sales");

        // 6. Persistence
        Reopen();
        chart = _handler.Get("/Sheet1/chart[1]");
        chart.Type.Should().Be("chart");
        ((string)chart.Format["title"]).Should().Be("Updated Sales");
    }

    // ==================== Chart Enhanced Properties ====================

    [Fact]
    public void Chart_TitleFont_Lifecycle()
    {
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:1,2,3",
            ["title"] = "My Chart"
        });

        _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["title.font"] = "Impact",
            ["title.size"] = "28",
            ["title.color"] = "FF0000",
            ["title.bold"] = "true"
        });

        // Title text should still be readable
        var chart = _handler.Get("/Sheet1/chart[1]");
        ((string)chart.Format["title"]).Should().Be("My Chart");
    }

    [Fact]
    public void Chart_SeriesShadowOutline()
    {
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["categories"] = "A,B",
            ["series1"] = "S1:10,20",
            ["colors"] = "FF0000"
        });

        _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["series.shadow"] = "000000-8-135-4-50",
            ["series.outline"] = "FFFFFF-0.5"
        });

        // Should not throw, chart still readable
        var chart = _handler.Get("/Sheet1/chart[1]");
        chart.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_GradientFill_And_View3d()
    {
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["categories"] = "Q1,Q2",
            ["series1"] = "Rev:100,200"
        });

        _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["chartfill"] = "0D1117-161B22:270",
            ["plotfill"] = "161B22",
            ["gradient"] = "FF0000-0000FF:90",
            ["gap"] = "80",
            ["axisfont"] = "9:8B949E",
            ["legendfont"] = "9:CCCCCC"
        });

        var chart = _handler.Get("/Sheet1/chart[1]");
        chart.Type.Should().Be("chart");
    }

    [Fact]
    public void Chart_Column3d()
    {
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column3d",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:1,2,3",
            ["title"] = "3D Chart"
        });

        _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["view3d"] = "15,20,30"
        });

        var chart = _handler.Get("/Sheet1/chart[1]");
        chart.Type.Should().Be("chart");
        ((string)chart.Format["title"]).Should().Be("3D Chart");
    }

    // ==================== Excel Shape Lifecycle ====================

    [Fact]
    public void Shape_Add_Get_Lifecycle()
    {
        // 1. Add shape with text, font, color
        _handler.Add("/Sheet1", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Hello",
            ["font"] = "Arial",
            ["size"] = "24",
            ["bold"] = "true",
            ["color"] = "FF0000",
            ["fill"] = "0000FF",
            ["align"] = "center",
            ["x"] = "1", ["y"] = "2", ["width"] = "6", ["height"] = "3"
        });

        // 2. Get + Verify
        var node = _handler.Get("/Sheet1/shape[1]");
        node.Type.Should().Be("shape");
        node.Text.Should().Be("Hello");
        node.Format["font"].ToString()!.Should().Be("Arial");
        node.Format["size"].ToString()!.Should().Be("24pt");
        node.Format["bold"].Should().Be(true);
        node.Format["color"].ToString()!.Should().Be("#FF0000");
        node.Format["fill"].ToString()!.Should().Be("#0000FF");
        node.Format["x"].ToString()!.Should().Be("1");
        node.Format["width"].ToString()!.Should().Be("6");
    }

    [Fact]
    public void Shape_Set_Font_Color_Size()
    {
        _handler.Add("/Sheet1", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Test", ["font"] = "Arial", ["size"] = "20", ["color"] = "000000"
        });

        // Set new font properties
        _handler.Set("/Sheet1/shape[1]", new Dictionary<string, string>
        {
            ["font"] = "Georgia", ["size"] = "36", ["color"] = "FF4500", ["bold"] = "true"
        });

        var node = _handler.Get("/Sheet1/shape[1]");
        node.Format["font"].ToString()!.Should().Be("Georgia");
        node.Format["size"].ToString()!.Should().Be("36pt");
        node.Format["color"].ToString()!.Should().Be("#FF4500");
        node.Format["bold"].Should().Be(true);
    }

    [Fact]
    public void Shape_Set_Text()
    {
        _handler.Add("/Sheet1", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Original", ["font"] = "Arial"
        });

        _handler.Set("/Sheet1/shape[1]", new Dictionary<string, string> { ["text"] = "Updated" });

        var node = _handler.Get("/Sheet1/shape[1]");
        node.Text.Should().Be("Updated");
        // Font should be preserved after text change
        node.Format["font"].ToString()!.Should().Be("Arial");
    }

    [Fact]
    public void Shape_Effects_Shadow_Glow()
    {
        _handler.Add("/Sheet1", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Effects", ["fill"] = "none",
            ["shadow"] = "FF0000-8-135-4-60",
            ["glow"] = "FF6600-12-80"
        });

        var node = _handler.Get("/Sheet1/shape[1]");
        node.Format.Should().ContainKey("shadow");
        node.Format.Should().ContainKey("glow");
        node.Format["shadow"].ToString()!.Should().StartWith("#FF0000");
        node.Format["glow"].ToString()!.Should().StartWith("#FF6600");

        // Remove effects
        _handler.Set("/Sheet1/shape[1]", new Dictionary<string, string>
        {
            ["shadow"] = "none", ["glow"] = "none"
        });
        node = _handler.Get("/Sheet1/shape[1]");
        node.Format.Should().NotContainKey("shadow");
        node.Format.Should().NotContainKey("glow");
    }

    [Fact]
    public void Shape_Persist_AfterReopen()
    {
        _handler.Add("/Sheet1", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Persist", ["font"] = "Impact", ["size"] = "48",
            ["bold"] = "true", ["color"] = "FFD700", ["fill"] = "1A1A2E"
        });

        Reopen();

        var node = _handler.Get("/Sheet1/shape[1]");
        node.Type.Should().Be("shape");
        node.Text.Should().Be("Persist");
        node.Format["font"].ToString()!.Should().Be("Impact");
        node.Format["size"].ToString()!.Should().Be("48pt");
        node.Format["bold"].Should().Be(true);
        node.Format["color"].ToString()!.Should().Be("#FFD700");
        node.Format["fill"].ToString()!.Should().Be("#1A1A2E");
    }
}
