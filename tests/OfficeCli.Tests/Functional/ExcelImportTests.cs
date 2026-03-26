// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for CSV/TSV import into Excel sheets.
/// </summary>
public class ExcelImportTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelImportTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen() { _handler.Dispose(); _handler = new ExcelHandler(_path, editable: true); }

    // ==================== CSV Parsing ====================

    [Fact]
    public void ParseCsv_SimpleFields()
    {
        var rows = ExcelHandler.ParseCsv("a,b,c\n1,2,3", ',');
        rows.Should().HaveCount(2);
        rows[0].Should().Equal("a", "b", "c");
        rows[1].Should().Equal("1", "2", "3");
    }

    [Fact]
    public void ParseCsv_QuotedFields_WithComma()
    {
        var rows = ExcelHandler.ParseCsv("\"hello, world\",b\n1,2", ',');
        rows.Should().HaveCount(2);
        rows[0][0].Should().Be("hello, world");
        rows[0][1].Should().Be("b");
    }

    [Fact]
    public void ParseCsv_EscapedQuotes()
    {
        var rows = ExcelHandler.ParseCsv("\"he said \"\"hi\"\"\",b", ',');
        rows.Should().HaveCount(1);
        rows[0][0].Should().Be("he said \"hi\"");
    }

    [Fact]
    public void ParseCsv_NewlineInsideQuotes()
    {
        var rows = ExcelHandler.ParseCsv("\"line1\nline2\",b\nc,d", ',');
        rows.Should().HaveCount(2);
        rows[0][0].Should().Be("line1\nline2");
        rows[0][1].Should().Be("b");
        rows[1].Should().Equal("c", "d");
    }

    [Fact]
    public void ParseCsv_TabDelimiter()
    {
        var rows = ExcelHandler.ParseCsv("a\tb\tc\n1\t2\t3", '\t');
        rows.Should().HaveCount(2);
        rows[0].Should().Equal("a", "b", "c");
    }

    [Fact]
    public void ParseCsv_BomStripped()
    {
        var rows = ExcelHandler.ParseCsv("\uFEFFa,b\n1,2", ',');
        rows[0][0].Should().Be("a");
    }

    [Fact]
    public void ParseCsv_CrLfLineEndings()
    {
        var rows = ExcelHandler.ParseCsv("a,b\r\n1,2\r\n", ',');
        rows.Should().HaveCount(2);
        rows[0].Should().Equal("a", "b");
        rows[1].Should().Equal("1", "2");
    }

    [Fact]
    public void ParseCsv_EmptyContent()
    {
        var rows = ExcelHandler.ParseCsv("", ',');
        rows.Should().BeEmpty();
    }

    // ==================== Import: Basic ====================

    [Fact]
    public void Import_BasicCsv_CellsPopulated()
    {
        var csv = "Name,Age,City\nAlice,30,NYC\nBob,25,LA";
        _handler.Import("/Sheet1", csv, ',', false, "A1");

        var a1 = _handler.Get("/Sheet1/A1", 0);
        a1.Text.Should().Be("Name");

        var b2 = _handler.Get("/Sheet1/B2", 0);
        b2.Text.Should().Be("30");

        var c3 = _handler.Get("/Sheet1/C3", 0);
        c3.Text.Should().Be("LA");
    }

    [Fact]
    public void Import_StartCellOffset_CellsAtCorrectPosition()
    {
        var csv = "x,y\n1,2";
        _handler.Import("/Sheet1", csv, ',', false, "C5");

        var c5 = _handler.Get("/Sheet1/C5", 0);
        c5.Text.Should().Be("x");

        var d6 = _handler.Get("/Sheet1/D6", 0);
        d6.Text.Should().Be("2");
    }

    // ==================== Import: Type Detection ====================

    [Fact]
    public void Import_NumberDetection()
    {
        var csv = "42\n3.14\n-100";
        _handler.Import("/Sheet1", csv, ',', false, "A1");

        // Numbers should be stored without DataType (numeric default)
        var a1 = _handler.Get("/Sheet1/A1", 0);
        a1.Text.Should().Be("42");

        var a2 = _handler.Get("/Sheet1/A2", 0);
        a2.Text.Should().Be("3.14");

        var a3 = _handler.Get("/Sheet1/A3", 0);
        a3.Text.Should().Be("-100");
    }

    [Fact]
    public void Import_BooleanDetection()
    {
        var csv = "TRUE\nFALSE\ntrue";
        _handler.Import("/Sheet1", csv, ',', false, "A1");

        var a1 = _handler.Get("/Sheet1/A1", 0);
        // Boolean cells display as TRUE/FALSE
        a1.Text.Should().BeOneOf("1", "TRUE");
    }

    [Fact]
    public void Import_FormulaDetection()
    {
        var csv = "10\n20\n=A1+A2";
        _handler.Import("/Sheet1", csv, ',', false, "A1");

        var a3 = _handler.Get("/Sheet1/A3", 0);
        a3.Format.Should().ContainKey("formula");
        ((string)a3.Format["formula"]).Should().Be("A1+A2");
    }

    [Fact]
    public void Import_DateDetection()
    {
        var csv = "2024-01-15";
        _handler.Import("/Sheet1", csv, ',', false, "A1");

        // Date stored as OLE Automation number
        var a1 = _handler.Get("/Sheet1/A1", 0);
        double.TryParse(a1.Text, out var oaDate).Should().BeTrue();
        var dt = DateTime.FromOADate(oaDate);
        dt.Year.Should().Be(2024);
        dt.Month.Should().Be(1);
        dt.Day.Should().Be(15);
    }

    [Fact]
    public void Import_StringFallback()
    {
        var csv = "hello world";
        _handler.Import("/Sheet1", csv, ',', false, "A1");

        var a1 = _handler.Get("/Sheet1/A1", 0);
        a1.Text.Should().Be("hello world");
    }

    // ==================== Import: --header ====================

    [Fact]
    public void Import_WithHeader_AutoFilterSet()
    {
        var csv = "Name,Age\nAlice,30\nBob,25";
        _handler.Import("/Sheet1", csv, ',', true, "A1");

        var sheet = _handler.Get("/Sheet1", 0);
        sheet.Format.Should().ContainKey("autoFilter");
        ((string)sheet.Format["autoFilter"]).Should().Be("A1:B3");
    }

    [Fact]
    public void Import_WithHeader_FreezePaneSet()
    {
        var csv = "Name,Age\nAlice,30\nBob,25";
        _handler.Import("/Sheet1", csv, ',', true, "A1");

        var sheet = _handler.Get("/Sheet1", 0);
        sheet.Format.Should().ContainKey("freeze");
        // Freeze below row 1: pane at A2
        ((string)sheet.Format["freeze"]).Should().Be("A2");
    }

    [Fact]
    public void Import_WithHeader_OffsetStartCell()
    {
        var csv = "Col1,Col2\nv1,v2";
        _handler.Import("/Sheet1", csv, ',', true, "B3");

        var sheet = _handler.Get("/Sheet1", 0);
        sheet.Format.Should().ContainKey("autoFilter");
        ((string)sheet.Format["autoFilter"]).Should().Be("B3:C4");
    }

    // ==================== Import: Persistence ====================

    [Fact]
    public void Import_Persistence_SurvivesReopen()
    {
        var csv = "Name,Value\nTest,42";
        _handler.Import("/Sheet1", csv, ',', true, "A1");
        Reopen();

        var a1 = _handler.Get("/Sheet1/A1", 0);
        a1.Text.Should().Be("Name");

        var b2 = _handler.Get("/Sheet1/B2", 0);
        b2.Text.Should().Be("42");

        var sheet = _handler.Get("/Sheet1", 0);
        sheet.Format.Should().ContainKey("autoFilter");
        sheet.Format.Should().ContainKey("freeze");
    }

    // ==================== Import: TSV ====================

    [Fact]
    public void Import_Tsv_TabDelimiter()
    {
        var tsv = "Name\tAge\nAlice\t30";
        _handler.Import("/Sheet1", tsv, '\t', false, "A1");

        var a1 = _handler.Get("/Sheet1/A1", 0);
        a1.Text.Should().Be("Name");

        var b1 = _handler.Get("/Sheet1/B1", 0);
        b1.Text.Should().Be("Age");
    }

    // ==================== Import: Edge Cases ====================

    [Fact]
    public void Import_EmptyContent_ReturnsNoData()
    {
        var result = _handler.Import("/Sheet1", "", ',', false, "A1");
        result.Should().Be("No data to import");
    }

    [Fact]
    public void Import_SingleCell()
    {
        _handler.Import("/Sheet1", "hello", ',', false, "A1");

        var a1 = _handler.Get("/Sheet1/A1", 0);
        a1.Text.Should().Be("hello");
    }
}
