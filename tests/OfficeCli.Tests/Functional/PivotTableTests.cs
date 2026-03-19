// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Pivot table tests: Create → Get → Verify → Set → Get → Verify → Reopen → Verify
/// </summary>
public class PivotTableTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public PivotTableTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
        PopulateSourceData();
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen() { _handler.Dispose(); _handler = new ExcelHandler(_path, editable: true); }

    /// <summary>
    /// Create sample data: Region | Product | Sales | Quantity
    /// </summary>
    private void PopulateSourceData()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "Region", ["c2"] = "Product", ["c3"] = "Sales", ["c4"] = "Quantity" });
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "East", ["c2"] = "Widget", ["c3"] = "100", ["c4"] = "10" });
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "West", ["c2"] = "Widget", ["c3"] = "200", ["c4"] = "20" });
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "East", ["c2"] = "Gadget", ["c3"] = "150", ["c4"] = "15" });
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "West", ["c2"] = "Gadget", ["c3"] = "250", ["c4"] = "25" });
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "East", ["c2"] = "Widget", ["c3"] = "120", ["c4"] = "12" });

        // Set numeric types for Sales and Quantity
        for (int r = 2; r <= 6; r++)
        {
            _handler.Set($"/Sheet1/C{r}", new() { ["type"] = "number" });
            _handler.Set($"/Sheet1/D{r}", new() { ["type"] = "number" });
        }
    }

    // ==================== Add ====================

    [Fact]
    public void Add_PivotTable_Basic()
    {
        var ptPath = _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["values"] = "Sales:sum"
        });

        ptPath.Should().Be("/Sheet1/pivottable[1]");
    }

    [Fact]
    public void Add_PivotTable_WithPosition()
    {
        var ptPath = _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "Sheet1!A1:D6",
            ["position"] = "F1",
            ["rows"] = "Region",
            ["values"] = "Sales:sum"
        });

        ptPath.Should().Be("/Sheet1/pivottable[1]");
    }

    [Fact]
    public void Add_PivotTable_RowsAndCols()
    {
        var ptPath = _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["position"] = "F1",
            ["rows"] = "Region",
            ["cols"] = "Product",
            ["values"] = "Sales:sum"
        });

        ptPath.Should().Be("/Sheet1/pivottable[1]");
    }

    [Fact]
    public void Add_PivotTable_MultipleValues()
    {
        var ptPath = _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["values"] = "Sales:sum,Quantity:count"
        });

        ptPath.Should().Be("/Sheet1/pivottable[1]");
    }

    [Fact]
    public void Add_PivotTable_WithFilter()
    {
        var ptPath = _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["filters"] = "Product",
            ["values"] = "Sales:sum"
        });

        ptPath.Should().Be("/Sheet1/pivottable[1]");
    }

    [Fact]
    public void Add_PivotTable_CustomName()
    {
        var ptPath = _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["values"] = "Sales:sum",
            ["name"] = "SalesReport"
        });

        ptPath.Should().Be("/Sheet1/pivottable[1]");
    }

    // ==================== Get ====================

    [Fact]
    public void Get_PivotTable_ReturnsProperties()
    {
        _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["position"] = "F1",
            ["rows"] = "Region",
            ["values"] = "Sales:sum",
            ["name"] = "SalesReport"
        });

        var node = _handler.Get("/Sheet1/pivottable[1]");
        node.Type.Should().Be("pivottable");
        node.Format["name"].Should().Be("SalesReport");
        node.Format.Should().ContainKey("fieldCount");
        ((int)node.Format["fieldCount"]).Should().Be(4);
    }

    [Fact]
    public void Get_PivotTable_RowFields()
    {
        _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["cols"] = "Product",
            ["values"] = "Sales:sum"
        });

        var node = _handler.Get("/Sheet1/pivottable[1]");
        node.Format.Should().ContainKey("fieldCount");
        ((int)node.Format["fieldCount"]).Should().Be(4);
    }

    // ==================== Set ====================

    [Fact]
    public void Set_PivotTable_Name()
    {
        _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["values"] = "Sales:sum"
        });

        _handler.Set("/Sheet1/pivottable[1]", new() { ["name"] = "UpdatedName" });

        var node = _handler.Get("/Sheet1/pivottable[1]");
        node.Format["name"].Should().Be("UpdatedName");
    }

    [Fact]
    public void Set_PivotTable_Style()
    {
        _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["values"] = "Sales:sum"
        });

        _handler.Set("/Sheet1/pivottable[1]", new() { ["style"] = "PivotStyleMedium9" });

        var node = _handler.Get("/Sheet1/pivottable[1]");
        node.Format["style"].Should().Be("PivotStyleMedium9");
    }

    // ==================== Persistence ====================

    [Fact]
    public void PivotTable_Persists_AfterReopen()
    {
        _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["position"] = "F1",
            ["rows"] = "Region",
            ["values"] = "Sales:sum",
            ["name"] = "SalesReport"
        });

        Reopen();

        var node = _handler.Get("/Sheet1/pivottable[1]");
        node.Type.Should().Be("pivottable");
        node.Format["name"].Should().Be("SalesReport");
    }

    [Fact]
    public void PivotTable_Set_Persists_AfterReopen()
    {
        _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["values"] = "Sales:sum"
        });

        _handler.Set("/Sheet1/pivottable[1]", new() { ["name"] = "Renamed", ["style"] = "PivotStyleDark1" });

        Reopen();

        var node = _handler.Get("/Sheet1/pivottable[1]");
        node.Format["name"].Should().Be("Renamed");
        node.Format["style"].Should().Be("PivotStyleDark1");
    }

    // ==================== Query ====================

    [Fact]
    public void Query_PivotTable_FindsAll()
    {
        _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6", ["rows"] = "Region", ["values"] = "Sales:sum", ["name"] = "PT1"
        });
        _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6", ["rows"] = "Product", ["values"] = "Quantity:count", ["name"] = "PT2"
        });

        var results = _handler.Query("pivottable");
        results.Should().HaveCount(2);
        results[0].Format["name"].Should().Be("PT1");
        results[1].Format["name"].Should().Be("PT2");
    }

    // ==================== Aggregation Functions ====================

    [Fact]
    public void Add_PivotTable_AverageFunctiong()
    {
        var ptPath = _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["values"] = "Sales:average"
        });

        var node = _handler.Get(ptPath);
        node.Format.Should().ContainKey("fieldCount");
    }

    [Fact]
    public void Add_PivotTable_MaxFunction()
    {
        var ptPath = _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "Region",
            ["values"] = "Sales:max"
        });

        var node = _handler.Get(ptPath);
        node.Format.Should().ContainKey("fieldCount");
    }

    // ==================== FieldIndex Reference ====================

    [Fact]
    public void Add_PivotTable_ByColumnIndex()
    {
        var ptPath = _handler.Add("/Sheet1", "pivottable", null, new()
        {
            ["source"] = "A1:D6",
            ["rows"] = "0",     // Column A (Region)
            ["values"] = "2:sum" // Column C (Sales)
        });

        var node = _handler.Get(ptPath);
        node.Format.Should().ContainKey("fieldCount");
    }
}
