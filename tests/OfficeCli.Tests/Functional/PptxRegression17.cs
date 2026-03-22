// Bug hunt Part 17 — deeper edge cases, persistence bugs, cross-handler issues.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class PptxRegression17 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public PptxRegression17()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt17_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt17_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt17_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        using (var pptx = new PowerPointHandler(_pptxPath, editable: true))
            pptx.Add("/", "slide", null, new());
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _excelHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private WordHandler ReopenWord()
    {
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
        return _wordHandler;
    }

    private ExcelHandler ReopenExcel()
    {
        _excelHandler.Dispose();
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        return _excelHandler;
    }


    // ==================== BUG #1: PPTX shape bold value false in Get is missing ====================
    // NodeBuilder.cs:268 only reports bold=true when Bold.Value is true.
    // If bold is explicitly set to false (to override inherited bold), Get doesn't report it.
    // After Set bold=false, Get should show bold=false, but it shows nothing.
    [Fact]
    public void Pptx_Shape_Bold_False_ShouldBeInFormat()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["bold"] = "true"
        });

        // Verify bold is true
        var shape1 = pptx.Get("/slide[1]/shape[1]");
        shape1.Format.Should().ContainKey("bold");

        // Now set bold to false
        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["bold"] = "false"
        });

        var shape2 = pptx.Get("/slide[1]/shape[1]");
        // BUG: After explicitly setting bold=false, Get doesn't report bold at all.
        // The format should include bold=false to confirm the override was applied.
        shape2.Format.Should().ContainKey("bold",
            "explicitly setting bold=false should be reported in Format, not omitted");
    }


    // ==================== BUG #3: Excel sheet Add returns wrong path when sheets exist ====================
    // When adding a second sheet, the returned path should use the new sheet's name.
    [Fact]
    public void Excel_Add_Sheet_ReturnedPath_ShouldUseNewName()
    {
        var sheetPath = _excelHandler.Add("/", "sheet", null, new()
        {
            ["name"] = "MySheet"
        });

        sheetPath.Should().Contain("MySheet",
            "returned path should include the new sheet's name");

        var sheet = _excelHandler.Get(sheetPath);
        sheet.Should().NotBeNull();
        sheet.Type.Should().Be("sheet");
    }


    // ==================== BUG #4: Word table cell width not reported by Get ====================
    // After creating a table with column widths, Get on cell should show width.
    [Fact]
    public void Word_TableCell_Get_ShouldIncludeWidth()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "2",
            ["colwidths"] = "3000,5000"
        });

        var cell1 = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell1.Should().NotBeNull();

        // BUG: The cell width was set during creation but Get doesn't report it
        cell1.Format.Should().ContainKey("width",
            "table cell Get should expose width when column widths are specified");
    }


    // ==================== BUG #5: PPTX shape rotation not in Get Format ====================
    // After setting rotation on a shape, Get should report it.
    [Fact]
    public void Pptx_Shape_Get_ShouldIncludeRotation()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Rotated"
        });

        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["rotation"] = "45"
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        shape.Should().NotBeNull();

        shape.Format.Should().ContainKey("rotation",
            "shape Get should expose rotation property when it's set");
    }


    // ==================== BUG #8: PPTX table cell vertical alignment not in Get ====================
    // After setting valign on a table cell, Get should report it.
    [Fact]
    public void Pptx_TableCell_Get_ShouldIncludeVerticalAlignment()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Centered",
            ["valign"] = "middle"
        });

        var cell = pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Should().NotBeNull();

        cell.Format.Should().ContainKey("valign",
            "table cell Get should expose vertical alignment when it's set");
    }


    // ==================== BUG #9: Word Add run returns path but Set on run adds extra RunProperties ====================
    // When you Set properties on an existing run, EnsureRunProperties should check for existing
    // rather than always creating new. Duplicate RunProperties in a single Run is invalid.
    [Fact]
    public void Word_Run_MultipleSetCalls_ShouldNotDuplicate()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Hello"
        });

        // Multiple Set calls on the same run
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["bold"] = "true" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["italic"] = "true" });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["color"] = "FF0000" });

        var run = _wordHandler.Get("/body/p[1]/r[1]");
        run.Should().NotBeNull();
        run.Format.Should().ContainKey("bold");
        run.Format.Should().ContainKey("italic");
        run.Format.Should().ContainKey("color");

        // Validate the document to check for duplicate elements
        var errors = _wordHandler.Validate();
        var runErrors = errors.Where(e =>
            e.Description.Contains("RunProperties", StringComparison.OrdinalIgnoreCase)).ToList();
        runErrors.Should().BeEmpty(
            "multiple Set calls on the same run should not create duplicate RunProperties");
    }


}
