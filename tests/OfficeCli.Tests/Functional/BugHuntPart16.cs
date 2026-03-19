// Bug hunt Part 16 — more edge cases: persistence, property round-trips, path issues.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart16 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public BugHuntPart16()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt16_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt16_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt16_{Guid.NewGuid():N}.pptx");
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


    // ==================== BUG #1: Word Set table cell alignment only affects first paragraph ====================
    // WordHandler.Set.cs:881-895 sets alignment only on the first paragraph in the cell.
    // If a cell has multiple paragraphs, only the first is aligned.
    [Fact]
    public void Word_TableCell_SetAlignment_ShouldAffectAllParagraphs()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        // Add text to cell
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "First line"
        });

        // Add a second paragraph to the cell
        _wordHandler.Add("/body/tbl[1]/tr[1]/tc[1]", "paragraph", null, new()
        {
            ["text"] = "Second line"
        });

        // Set center alignment on the cell
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["alignment"] = "center"
        });

        // Check both paragraphs
        var cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]", depth: 2);
        cell.Children.Count.Should().BeGreaterThanOrEqualTo(2);

        // BUG: Only the first paragraph gets center alignment
        // The second paragraph retains its default (left) alignment
        foreach (var para in cell.Children)
        {
            if (para.Type == "paragraph")
            {
                para.Format.Should().ContainKey("alignment",
                    "all paragraphs in a cell should get the alignment, not just the first");
            }
        }
    }


    // ==================== BUG #2: Excel cell value "false" gets converted to "0" ====================
    // Same as boolean "true" → "1" bug. "false" → boolean type → stored as "0".
    [Fact]
    public void Excel_SetCellValue_False_RoundTrip()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "false"
        });

        var cell = _excelHandler.Get("/Sheet1/A1");
        cell.Should().NotBeNull();

        // BUG: "false" is auto-detected as boolean and stored as "0"
        cell.Text.Should().Be("false",
            "setting value 'false' should preserve the original string, not convert to '0'");
    }


    // ==================== BUG #4: Word Add section break returns inconsistent path ====================
    // Section break returns "/section[N]" but the section count may not match
    // because it counts sections in paragraph properties, not body-level SectionProperties.
    [Fact]
    public void Word_Add_SectionBreak_ReturnedPath_ShouldBeUsable()
    {
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Page 1" });

        var secPath = _wordHandler.Add("/", "section", null, new()
        {
            ["type"] = "nextpage"
        });

        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Page 2" });

        // The returned path should work with Get
        var act = () => _wordHandler.Get(secPath);
        act.Should().NotThrow("the section path returned by Add should be usable with Get");

        var section = _wordHandler.Get(secPath);
        section.Type.Should().Be("section");
    }


    // ==================== BUG #5: PPTX shape strikethrough naming inconsistency ====================
    // Set uses "strike" as the property key, but Get returns it as "strikethrough".
    // This means code that does: Set(strike=true) → Get → check Format["strike"] will fail.
    // The property name should be consistent between Set and Get.
    [Fact]
    public void Pptx_Shape_Strikethrough_NamingConsistency()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Struck"
        });

        // Set uses "strike" as the key
        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["strike"] = "true"
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        shape.Should().NotBeNull();

        // BUG: Get returns "strikethrough" key, not "strike"
        // The Set key and Get key should match for consistent round-trips
        shape.Format.Should().ContainKey("strike",
            "Get should use the same property name as Set ('strike'), not 'strikethrough'");

        // Also, the value should be a simple boolean, not the raw OOXML value "sngStrike"
        if (shape.Format.ContainsKey("strikethrough"))
        {
            var val = shape.Format["strikethrough"]?.ToString();
            val.Should().Be("true",
                "strikethrough value should be normalized to 'true'/'false', not raw OOXML 'sngStrike'");
        }
    }


    // ==================== BUG #6: Word footer Get doesn't include type (default/even/first) ====================
    // When multiple footers exist with different types, Get should show which type each is.
    [Fact]
    public void Word_Footer_Get_ShouldIncludeType()
    {
        _wordHandler.Add("/", "footer", null, new()
        {
            ["text"] = "Default Footer",
            ["type"] = "default"
        });

        var footer = _wordHandler.Get("/footer[1]");
        footer.Should().NotBeNull();
        footer.Text.Should().Contain("Default Footer");

        // The footer should report its type
        footer.Format.Should().ContainKey("type",
            "footer Get should include the type (default/even/first) in Format");
    }


    // ==================== BUG #8: Word table Set width with percent ====================
    // WordHandler.Add.cs:428 uses int.Parse(tv.TrimEnd('%')) * 50
    // This converts percentage to OOXML pct50 unit. But if someone passes "100%",
    // it stores 5000. The Get handler should report this back consistently.
    [Fact]
    public void Word_Table_Width_Percent_ShouldRoundTrip()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "2",
            ["width"] = "100%"
        });

        var table = _wordHandler.Get("/body/tbl[1]");
        table.Should().NotBeNull();

        // The table width should be reported as a percentage
        table.Format.Should().ContainKey("width",
            "table Get should include width property for verification");

        if (table.Format.ContainsKey("width"))
        {
            var width = table.Format["width"]?.ToString();
            // Should be reported back in a user-friendly format like "100%"
            width.Should().Contain("%",
                "table width set as percentage should be reported back as percentage");
        }
    }


    // ==================== BUG #9: Excel named range Add then Get by name ====================
    // After Add("/", "namedrange", ...) which returns /namedrange[1],
    // Get("/namedrange[MyRange]") should also work (by name lookup).
    [Fact]
    public void Excel_NamedRange_GetByName_ShouldWork()
    {
        _excelHandler.Add("/", "namedrange", null, new()
        {
            ["name"] = "TestRange",
            ["ref"] = "Sheet1!$A$1:$B$10"
        });

        // Get by index should work
        var byIndex = _excelHandler.Get("/namedrange[1]");
        byIndex.Should().NotBeNull();
        byIndex.Format.Should().ContainKey("name");

        // Get by name should also work
        var byName = _excelHandler.Get("/namedrange[TestRange]");
        byName.Should().NotBeNull();
        byName.Format["name"]?.ToString().Should().Be("TestRange");
    }


    // ==================== BUG #10: PPTX table Get doesn't include tableStyleId readback ====================
    // When creating a table with a style, the style ID should be readable via Get.
    [Fact]
    public void Pptx_Table_Set_Style_ThenGet_ShouldShowStyle()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        // Set a table style
        pptx.Set("/slide[1]/table[1]", new()
        {
            ["style"] = "medium1"
        });

        var table = pptx.Get("/slide[1]/table[1]");
        table.Should().NotBeNull();
        table.Format.Should().ContainKey("tableStyleId",
            "table Get should include tableStyleId after Set style");
    }
}
