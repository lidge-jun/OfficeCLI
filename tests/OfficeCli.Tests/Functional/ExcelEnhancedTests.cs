// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for Excel enhanced features:
/// #1 rotation/indent, #2 tabColor, #3 zoom, #4 row/col grouping,
/// #5 bubble/radar/stock charts, #6 picture rotation/shadow/glow
/// </summary>
public class ExcelEnhancedTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelEnhancedTests()
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

    private void Reopen() { _handler.Dispose(); _handler = new ExcelHandler(_path, editable: true); }

    // ==================== #1 Text Rotation / Indent ====================

    [Fact]
    public void Set_TextRotation_90Degrees()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "Header" });
        _handler.Set("/Sheet1/A1", new() { ["rotation"] = "90" });

        Reopen();
        // Verify no crash on reopen — rotation is a style property
        var node = _handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Set_Indent_Level2()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "Indented" });
        _handler.Set("/Sheet1/A1", new() { ["indent"] = "2" });

        Reopen();
        var node = _handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Set_ShrinkToFit()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "Long text here" });
        _handler.Set("/Sheet1/A1", new() { ["shrinktofit"] = "true" });

        Reopen();
        var node = _handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
    }

    [Fact]
    public void Set_RotationAndIndent_Combined()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "Test" });
        _handler.Set("/Sheet1/A1", new() { ["rotation"] = "45", ["indent"] = "1", ["halign"] = "left" });

        Reopen();
        var node = _handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
    }

    // ==================== #2 Sheet Tab Color ====================

    [Fact]
    public void Set_TabColor()
    {
        _handler.Set("/Sheet1", new() { ["tabColor"] = "FF0000" });

        var node = _handler.Get("/Sheet1");
        node.Format["tabColor"].Should().Be("#FF0000");

        Reopen();
        var node2 = _handler.Get("/Sheet1");
        node2.Format["tabColor"].Should().Be("#FF0000");
    }

    [Fact]
    public void Set_TabColor_WithHash()
    {
        _handler.Set("/Sheet1", new() { ["tabColor"] = "#4472C4" });

        var node = _handler.Get("/Sheet1");
        node.Format["tabColor"].Should().Be("#4472C4");
    }

    [Fact]
    public void Set_TabColor_None_Removes()
    {
        _handler.Set("/Sheet1", new() { ["tabColor"] = "FF0000" });
        _handler.Set("/Sheet1", new() { ["tabColor"] = "none" });

        var node = _handler.Get("/Sheet1");
        node.Format.Should().NotContainKey("tabColor");
    }

    // ==================== #3 Zoom Level ====================

    [Fact]
    public void Set_Zoom_150()
    {
        _handler.Set("/Sheet1", new() { ["zoom"] = "150" });

        var node = _handler.Get("/Sheet1");
        node.Format["zoom"].Should().Be(150u);

        Reopen();
        var node2 = _handler.Get("/Sheet1");
        node2.Format["zoom"].Should().Be(150u);
    }

    [Fact]
    public void Set_Zoom_75()
    {
        _handler.Set("/Sheet1", new() { ["zoom"] = "75" });

        var node = _handler.Get("/Sheet1");
        node.Format["zoom"].Should().Be(75u);
    }

    // ==================== #4 Row/Column Grouping ====================

    [Fact]
    public void Set_RowOutlineLevel()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "Data" });
        _handler.Set("/Sheet1/row[1]", new() { ["outline"] = "1" });

        var node = _handler.Get("/Sheet1/row[1]");
        node.Format["outlineLevel"].Should().Be((byte)1);

        Reopen();
        var node2 = _handler.Get("/Sheet1/row[1]");
        node2.Format["outlineLevel"].Should().Be((byte)1);
    }

    [Fact]
    public void Set_RowOutlineLevel_Collapsed()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "Data" });
        _handler.Set("/Sheet1/row[1]", new() { ["outline"] = "2", ["collapsed"] = "true" });

        var node = _handler.Get("/Sheet1/row[1]");
        node.Format["outlineLevel"].Should().Be((byte)2);
        node.Format["collapsed"].Should().Be(true);
    }

    [Fact]
    public void Set_ColumnOutlineLevel()
    {
        _handler.Set("/Sheet1/col[B]", new() { ["outline"] = "1" });

        var node = _handler.Get("/Sheet1/col[B]");
        node.Format["outlineLevel"].Should().Be((byte)1);

        Reopen();
        var node2 = _handler.Get("/Sheet1/col[B]");
        node2.Format["outlineLevel"].Should().Be((byte)1);
    }

    [Fact]
    public void Set_MultipleRows_Grouping()
    {
        for (int i = 1; i <= 5; i++)
            _handler.Add("/Sheet1", "row", null, new() { ["c1"] = $"Row {i}" });

        // Group rows 2-4 at level 1
        _handler.Set("/Sheet1/row[2]", new() { ["outline"] = "1" });
        _handler.Set("/Sheet1/row[3]", new() { ["outline"] = "1" });
        _handler.Set("/Sheet1/row[4]", new() { ["outline"] = "1" });

        var node2 = _handler.Get("/Sheet1/row[2]");
        var node3 = _handler.Get("/Sheet1/row[3]");
        var node4 = _handler.Get("/Sheet1/row[4]");
        node2.Format["outlineLevel"].Should().Be((byte)1);
        node3.Format["outlineLevel"].Should().Be((byte)1);
        node4.Format["outlineLevel"].Should().Be((byte)1);
    }

    // ==================== #5 New Chart Types ====================

    [Fact]
    public void Add_BubbleChart()
    {
        var chartPath = _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "bubble",
            ["title"] = "Bubble Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "1,2,3"
        });

        var node = _handler.Get(chartPath, depth: 0);
        node.Format["chartType"].Should().Be("bubble");
        node.Format["title"].Should().Be("Bubble Test");

        Reopen();
        var node2 = _handler.Get(chartPath, depth: 0);
        node2.Format["chartType"].Should().Be("bubble");
    }

    [Fact]
    public void Add_RadarChart()
    {
        var chartPath = _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "radar",
            ["title"] = "Radar Test",
            ["data"] = "S1:5,4,3,2,5;S2:3,5,2,4,3",
            ["categories"] = "Speed,Power,Range,Defense,Magic"
        });

        var node = _handler.Get(chartPath, depth: 0);
        node.Format["chartType"].Should().Be("radar");
        node.Format["title"].Should().Be("Radar Test");
        node.Format["seriesCount"].Should().Be(2);

        Reopen();
        var node2 = _handler.Get(chartPath, depth: 1);
        node2.Format["chartType"].Should().Be("radar");
        node2.Children.Should().HaveCount(2);
    }

    [Fact]
    public void Add_RadarChart_Filled()
    {
        var chartPath = _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "radar",
            ["radarStyle"] = "filled",
            ["title"] = "Filled Radar",
            ["data"] = "S1:5,4,3,2,5",
            ["categories"] = "A,B,C,D,E"
        });

        var node = _handler.Get(chartPath, depth: 0);
        node.Format["chartType"].Should().Be("radar");
    }

    [Fact]
    public void Add_StockChart()
    {
        var chartPath = _handler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "stock",
            ["title"] = "Stock Price",
            ["data"] = "High:150,155,148;Low:140,142,135;Close:145,150,140",
            ["categories"] = "Mon,Tue,Wed"
        });

        var node = _handler.Get(chartPath, depth: 0);
        node.Format["chartType"].Should().Be("stock");
        node.Format["seriesCount"].Should().Be(3);

        Reopen();
        var node2 = _handler.Get(chartPath, depth: 0);
        node2.Format["chartType"].Should().Be("stock");
    }

    // ==================== #6 Picture Rotation/Effects ====================

    [Fact]
    public void Set_PictureRotation()
    {
        // Create a small test image
        var imgPath = CreateTestImage();
        try
        {
            _handler.Add("/Sheet1", "picture", null, new()
            {
                ["path"] = imgPath, ["x"] = "0", ["y"] = "0", ["width"] = "3", ["height"] = "3"
            });

            _handler.Set("/Sheet1/picture[1]", new() { ["rotation"] = "45" });

            Reopen();
            // Verify no corruption
            var node = _handler.Get("/Sheet1");
            node.Should().NotBeNull();
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }

    [Fact]
    public void Set_PictureShadow()
    {
        var imgPath = CreateTestImage();
        try
        {
            _handler.Add("/Sheet1", "picture", null, new()
            {
                ["path"] = imgPath, ["x"] = "0", ["y"] = "0", ["width"] = "3", ["height"] = "3"
            });

            _handler.Set("/Sheet1/picture[1]", new() { ["shadow"] = "000000:4:3:45" });

            Reopen();
            var node = _handler.Get("/Sheet1");
            node.Should().NotBeNull();
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }

    [Fact]
    public void Set_PictureGlow()
    {
        var imgPath = CreateTestImage();
        try
        {
            _handler.Add("/Sheet1", "picture", null, new()
            {
                ["path"] = imgPath, ["x"] = "0", ["y"] = "0", ["width"] = "3", ["height"] = "3"
            });

            _handler.Set("/Sheet1/picture[1]", new() { ["glow"] = "4472C4:8" });

            Reopen();
            var node = _handler.Get("/Sheet1");
            node.Should().NotBeNull();
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }

    [Fact]
    public void Set_PictureRotation_Shadow_Glow_Combined()
    {
        var imgPath = CreateTestImage();
        try
        {
            _handler.Add("/Sheet1", "picture", null, new()
            {
                ["path"] = imgPath, ["x"] = "0", ["y"] = "0", ["width"] = "3", ["height"] = "3"
            });

            _handler.Set("/Sheet1/picture[1]", new()
            {
                ["rotation"] = "30",
                ["shadow"] = "333333:6:4:60",
                ["glow"] = "FF6600:5"
            });

            Reopen();
            var node = _handler.Get("/Sheet1");
            node.Should().NotBeNull();
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }

    [Fact]
    public void Set_PictureShadow_None_Removes()
    {
        var imgPath = CreateTestImage();
        try
        {
            _handler.Add("/Sheet1", "picture", null, new()
            {
                ["path"] = imgPath, ["x"] = "0", ["y"] = "0", ["width"] = "3", ["height"] = "3"
            });

            _handler.Set("/Sheet1/picture[1]", new() { ["shadow"] = "000000:4:3:45" });
            _handler.Set("/Sheet1/picture[1]", new() { ["shadow"] = "none" });

            Reopen();
            var node = _handler.Get("/Sheet1");
            node.Should().NotBeNull();
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }

    // ==================== Helper ====================

    private static string CreateTestImage()
    {
        // Create a minimal valid PNG (1x1 pixel, red)
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        byte[] png = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, // 8-bit RGB
            0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
            0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, // compressed
            0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33, // data
            0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND chunk
            0xAE, 0x42, 0x60, 0x82
        ];
        File.WriteAllBytes(path, png);
        return path;
    }
}
