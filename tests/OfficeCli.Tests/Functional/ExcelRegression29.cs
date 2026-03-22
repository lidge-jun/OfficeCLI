// Bug hunt Part 29 — Excel handler confirmed bugs and edge cases:
// validation type raw enum, font.underline "single" not recognized,
// wraptext not in IsStyleKey, errorTitle case-sensitive, formula roundtrip,
// conditional formatting, merge cells, comments, named ranges, query shorthand.

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class ExcelRegression29 : IDisposable
{
    private readonly string _xlsxPath;
    private ExcelHandler _excelHandler;

    public ExcelRegression29()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt29_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_xlsxPath);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _excelHandler.Dispose();
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
    }

    private ExcelHandler Reopen()
    {
        _excelHandler.Dispose();
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        return _excelHandler;
    }

    // =================================================================
    // CONFIRMED BUG: Validation type returns raw enum "DataValidationValues { }"
    // instead of InnerText like "list", "whole", "decimal".
    // Fixed: changed .Value.ToString() to .InnerText in DataValidationToNode.
    // =================================================================

    [Fact]
    public void Bug_Excel_Validation_Type_Returns_FriendlyName()
    {
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A1:A10",
            ["type"] = "list",
            ["formula1"] = "X,Y,Z"
        });

        var node = _excelHandler.Get("/Sheet1/validation[1]");
        var typeVal = node.Format["type"]?.ToString();
        typeVal.Should().NotContain("DataValidationValues",
            "validation type should return friendly name, not raw enum");
        typeVal.Should().Be("list");
    }

    [Fact]
    public void Bug_Excel_Validation_Whole_Type_And_Operator()
    {
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "C1:C10",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "100"
        });

        var node = _excelHandler.Get("/Sheet1/validation[1]");
        node.Format["type"]?.ToString().Should().Be("whole");
        node.Format["operator"]?.ToString().Should().Be("between");
        node.Format["formula1"]?.ToString().Should().Be("1");
        node.Format["formula2"]?.ToString().Should().Be("100");
    }

    // =================================================================
    // CONFIRMED BUG: font.underline="single" not recognized by
    // GetOrCreateFont. IsTruthy("single") returns false.
    // Fixed: added explicit "single" check in underline parsing.
    // =================================================================

    [Fact]
    public void Bug_Excel_FontUnderline_Single_Set_Roundtrip()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Underlined",
            ["font.underline"] = "single"
        });

        var node = _excelHandler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("font.underline",
            "font.underline='single' should be recognized and readable");
        node.Format["font.underline"]?.ToString().Should().Be("single");
    }

    [Fact]
    public void Bug_Excel_FontUnderline_Double_Set_Roundtrip()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Double underlined",
            ["font.underline"] = "double"
        });

        var node = _excelHandler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("font.underline");
        node.Format["font.underline"]?.ToString().Should().Be("double");
    }

    // =================================================================
    // CONFIRMED BUG: "wraptext" not in IsStyleKey — falls to unsupported.
    // Fixed: added "wraptext" to IsStyleKey and mapped it in ApplyStyle.
    // =================================================================

    [Fact]
    public void Bug_Excel_WrapText_Set_Roundtrip()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Long text",
            ["wraptext"] = "true"
        });

        var node = _excelHandler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("alignment.wrapText",
            "wraptext shorthand should apply and be readable as alignment.wrapText");
    }

    // =================================================================
    // CONFIRMED BUG: Validation errorTitle case-sensitive — "errortitle"
    // not matching "errorTitle" in properties.TryGetValue.
    // Fixed: use case-insensitive dictionary for validation properties.
    // =================================================================

    [Fact]
    public void Bug_Excel_Validation_ErrorTitle_CaseInsensitive()
    {
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "D1:D10",
            ["type"] = "list",
            ["formula1"] = "A,B,C",
            ["showerror"] = "true",
            ["errortitle"] = "Invalid",
            ["error"] = "Please select from the list"
        });

        var node = _excelHandler.Get("/Sheet1/validation[1]");
        node.Format.Should().ContainKey("errorTitle",
            "errorTitle should be readable even when set as 'errortitle'");
        node.Format["errorTitle"]?.ToString().Should().Be("Invalid");
        node.Format["error"]?.ToString().Should().Be("Please select from the list");
    }

    [Fact]
    public void Bug_Excel_Validation_Prompt_CaseInsensitive()
    {
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "E1:E10",
            ["type"] = "list",
            ["formula1"] = "A,B,C",
            ["showinput"] = "true",
            ["prompttitle"] = "Choose",
            ["prompt"] = "Select a value"
        });

        var node = _excelHandler.Get("/Sheet1/validation[1]");
        node.Format.Should().ContainKey("promptTitle");
        node.Format["promptTitle"]?.ToString().Should().Be("Choose");
    }

    // =================================================================
    // Excel formula Set/Clear roundtrip
    // =================================================================

    [Fact]
    public void Bug_Excel_Cell_Formula_Roundtrip()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _excelHandler.Set("/Sheet1/A2", new() { ["value"] = "20" });
        _excelHandler.Set("/Sheet1/A3", new() { ["formula"] = "SUM(A1:A2)" });

        var node = _excelHandler.Get("/Sheet1/A3");
        node.Format["formula"]?.ToString().Should().Be("SUM(A1:A2)");
    }

    [Fact]
    public void Bug_Excel_Cell_Formula_Persists_After_Reopen()
    {
        _excelHandler.Set("/Sheet1/B1", new() { ["formula"] = "TODAY()" });
        Reopen();

        _excelHandler.Get("/Sheet1/B1").Format["formula"]?.ToString().Should().Be("TODAY()");
    }

    [Fact]
    public void Bug_Excel_Cell_Value_Clears_Formula()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["formula"] = "1+1" });
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "42" });

        var node = _excelHandler.Get("/Sheet1/A1");
        node.Text.Should().Be("42");
        node.Format.ContainsKey("formula").Should().BeFalse(
            "setting value should clear the formula");
    }

    [Fact]
    public void Bug_Excel_Cell_Clear_Removes_Everything()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Important", ["font.bold"] = "true"
        });
        _excelHandler.Set("/Sheet1/A1", new() { ["clear"] = "true" });

        var node = _excelHandler.Get("/Sheet1/A1");
        (string.IsNullOrEmpty(node.Text) || node.Text == "(empty)").Should().BeTrue();
        node.Format.ContainsKey("font.bold").Should().BeFalse();
    }

    // =================================================================
    // Excel merge/unmerge roundtrip
    // =================================================================

    [Fact]
    public void Bug_Excel_MergeCell_Roundtrip()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Merged" });
        _excelHandler.Set("/Sheet1/A1:C1", new() { ["merge"] = "true" });

        var node = _excelHandler.Get("/Sheet1/A1");
        node.Format["merge"]?.ToString().Should().Be("A1:C1");
    }

    [Fact]
    public void Bug_Excel_Unmerge_Roundtrip()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "M" });
        _excelHandler.Set("/Sheet1/A1:C1", new() { ["merge"] = "true" });
        _excelHandler.Set("/Sheet1/A1:C1", new() { ["merge"] = "false" });

        _excelHandler.Get("/Sheet1/A1").Format.ContainsKey("merge").Should().BeFalse();
    }

    // =================================================================
    // Excel conditional formatting roundtrip
    // =================================================================

    [Fact]
    public void Bug_Excel_CF_DataBar_Roundtrip()
    {
        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["sqref"] = "A1:A10", ["type"] = "databar", ["color"] = "FF6600"
        });

        var node = _excelHandler.Get("/Sheet1/cf[1]");
        node.Format["cfType"]?.ToString().Should().Be("dataBar");
    }

    [Fact]
    public void Bug_Excel_CF_ColorScale_Roundtrip()
    {
        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["sqref"] = "B1:B10", ["type"] = "colorscale",
            ["mincolor"] = "FF0000", ["maxcolor"] = "00FF00"
        });

        var node = _excelHandler.Get("/Sheet1/cf[1]");
        node.Format["cfType"]?.ToString().Should().Be("colorScale");
        node.Format.Should().ContainKey("mincolor");
        node.Format.Should().ContainKey("maxcolor");
    }

    [Fact]
    public void Bug_Excel_CF_IconSet_Roundtrip()
    {
        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["sqref"] = "C1:C10", ["type"] = "iconset", ["iconset"] = "3Arrows"
        });

        var node = _excelHandler.Get("/Sheet1/cf[1]");
        node.Format["cfType"]?.ToString().Should().Be("iconSet");
        node.Format.Should().ContainKey("iconset");
    }

    [Fact]
    public void Bug_Excel_CF_Set_Sqref_Roundtrip()
    {
        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["sqref"] = "A1:A10", ["type"] = "databar", ["color"] = "339966"
        });
        _excelHandler.Set("/Sheet1/cf[1]", new() { ["sqref"] = "A1:A20" });

        _excelHandler.Get("/Sheet1/cf[1]").Format["sqref"]?.ToString().Should().Be("A1:A20");
    }

    // =================================================================
    // Excel comment roundtrip
    // =================================================================

    [Fact]
    public void Bug_Excel_Comment_Add_And_Set_Roundtrip()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Data" });
        _excelHandler.Add("/Sheet1", "comment", null, new()
        {
            ["ref"] = "A1", ["text"] = "Original", ["author"] = "Tester"
        });

        var node = _excelHandler.Get("/Sheet1/comment[1]");
        node.Format["ref"]?.ToString().Should().Be("A1");
        node.Format["author"]?.ToString().Should().Be("Tester");

        _excelHandler.Set("/Sheet1/comment[1]", new() { ["text"] = "Updated" });
        _excelHandler.Get("/Sheet1/comment[1]").Text.Should().Contain("Updated");

        _excelHandler.Set("/Sheet1/comment[1]", new() { ["author"] = "Author2" });
        _excelHandler.Get("/Sheet1/comment[1]").Format["author"]?.ToString().Should().Be("Author2");
    }

    // =================================================================
    // Excel named range roundtrip
    // =================================================================

    [Fact]
    public void Bug_Excel_NamedRange_Add_Roundtrip()
    {
        _excelHandler.Add("/", "namedrange", null, new()
        {
            ["name"] = "TestRange", ["ref"] = "Sheet1!$A$1:$B$10"
        });

        var node = _excelHandler.Get("/namedrange[1]");
        node.Format["name"]?.ToString().Should().Be("TestRange");
        node.Format["ref"]?.ToString().Should().Contain("Sheet1");
    }

    [Fact]
    public void Bug_Excel_NamedRange_Scope_Roundtrip()
    {
        _excelHandler.Add("/", "namedrange", null, new()
        {
            ["name"] = "ScopedRange",
            ["ref"] = "Sheet1!$A$1:$C$5",
            ["scope"] = "Sheet1"
        });

        var node = _excelHandler.Get("/namedrange[1]");
        node.Format["scope"]?.ToString().Should().Be("Sheet1");
    }

    // =================================================================
    // Excel table roundtrip
    // =================================================================

    [Fact]
    public void Bug_Excel_Table_Add_With_Style()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Name" });
        _excelHandler.Set("/Sheet1/B1", new() { ["value"] = "Value" });
        _excelHandler.Set("/Sheet1/A2", new() { ["value"] = "Alpha" });
        _excelHandler.Set("/Sheet1/B2", new() { ["value"] = "100" });

        _excelHandler.Add("/Sheet1", "table", null, new()
        {
            ["ref"] = "A1:B2", ["name"] = "MyTable",
            ["displayname"] = "MyTable", ["style"] = "TableStyleMedium2"
        });

        var node = _excelHandler.Get("/Sheet1/table[1]");
        node.Format["name"]?.ToString().Should().Be("MyTable");
        node.Format["style"]?.ToString().Should().Be("TableStyleMedium2");
    }

    // =================================================================
    // Excel freeze/autofilter roundtrip
    // =================================================================

    [Fact]
    public void Bug_Excel_Freeze_Set_And_Remove()
    {
        _excelHandler.Set("/Sheet1", new() { ["freeze"] = "B2" });
        _excelHandler.Get("/Sheet1").Format["freeze"]?.ToString().Should().Be("B2");

        _excelHandler.Set("/Sheet1", new() { ["freeze"] = "none" });
        _excelHandler.Get("/Sheet1").Format.ContainsKey("freeze").Should().BeFalse();
    }

    [Fact]
    public void Bug_Excel_AutoFilter_Set()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Header" });
        _excelHandler.Set("/Sheet1", new() { ["autofilter"] = "A1:C1" });

        _excelHandler.Get("/Sheet1").Format["autoFilter"]?.ToString().Should().Be("A1:C1");
    }

    // =================================================================
    // Excel hyperlink roundtrip
    // =================================================================

    [Fact]
    public void Bug_Excel_Hyperlink_Set_And_Remove()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Click", ["link"] = "https://example.com"
        });
        _excelHandler.Get("/Sheet1/A1").Format.Should().ContainKey("link");

        _excelHandler.Set("/Sheet1/A1", new() { ["link"] = "none" });
        _excelHandler.Get("/Sheet1/A1").Format.ContainsKey("link").Should().BeFalse();
    }

    // =================================================================
    // Excel column/row properties
    // =================================================================

    [Fact]
    public void Bug_Excel_Column_Width_And_Hidden()
    {
        _excelHandler.Set("/Sheet1/col[A]", new() { ["width"] = "20" });
        _excelHandler.Set("/Sheet1/col[B]", new() { ["hidden"] = "true" });

        double.Parse(_excelHandler.Get("/Sheet1/col[A]").Format["width"]!.ToString()!).Should().Be(20.0);
        _excelHandler.Get("/Sheet1/col[B]").Format.Should().ContainKey("hidden");
    }

    [Fact]
    public void Bug_Excel_Row_Height_And_Hidden()
    {
        _excelHandler.Set("/Sheet1/row[1]", new() { ["height"] = "30" });
        double.Parse(_excelHandler.Get("/Sheet1/row[1]").Format["height"]!.ToString()!).Should().Be(30.0);

        _excelHandler.Set("/Sheet1/A2", new() { ["value"] = "data" });
        _excelHandler.Set("/Sheet1/row[2]", new() { ["hidden"] = "true" });
        _excelHandler.Get("/Sheet1/row[2]").Format.Should().ContainKey("hidden");
    }

    // =================================================================
    // Excel cell border and number format
    // =================================================================

    [Fact]
    public void Bug_Excel_Cell_Border_Set()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Bordered",
            ["border.bottom"] = "thin",
            ["border.bottom.color"] = "FF0000"
        });

        var node = _excelHandler.Get("/Sheet1/A1");
        node.Format["border.bottom"]?.ToString().Should().Be("thin");
    }

    [Fact]
    public void Bug_Excel_Cell_NumberFormat_Roundtrip()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "0.123456", ["numberformat"] = "0.00"
        });

        _excelHandler.Get("/Sheet1/A1").Format["numberformat"]?.ToString().Should().Be("0.00");
    }

    // =================================================================
    // Excel cell font properties — consolidated
    // =================================================================

    [Fact]
    public void Bug_Excel_Cell_FontProperties_Set()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Styled",
            ["font.bold"] = "true",
            ["font.italic"] = "true",
            ["font.name"] = "Arial",
            ["font.size"] = "20",
            ["font.strike"] = "true"
        });

        var node = _excelHandler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("font.bold");
        node.Format.Should().ContainKey("font.italic");
        node.Format["font.name"]?.ToString().Should().Be("Arial");
        node.Format["font.size"]?.ToString().Should().Be("20pt");
        node.Format.Should().ContainKey("font.strike");
    }

    // =================================================================
    // Excel multiple sheets
    // =================================================================

    [Fact]
    public void Bug_Excel_AddSheet_Navigate()
    {
        _excelHandler.Add("/", "sheet", null, new() { ["name"] = "Data" });
        _excelHandler.Set("/Data/A1", new() { ["value"] = "Hello" });

        _excelHandler.Get("/Data/A1").Text.Should().Be("Hello");
    }

    // =================================================================
    // Excel sheet Remove
    // =================================================================

    [Fact]
    public void Bug_Excel_Sheet_Remove()
    {
        _excelHandler.Add("/", "sheet", null, new() { ["name"] = "ToDelete" });
        _excelHandler.Set("/ToDelete/A1", new() { ["value"] = "bye" });

        _excelHandler.Remove("/ToDelete");
        ((Action)(() => _excelHandler.Get("/ToDelete")))
            .Should().Throw<ArgumentException>();
    }

    // =================================================================
    // Excel sheet-level merge
    // =================================================================

    [Fact]
    public void Bug_Excel_SheetLevel_Merge()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Merged" });
        _excelHandler.Set("/Sheet1", new() { ["merge"] = "A1:D1" });

        _excelHandler.Get("/Sheet1/A1").Format.Should().ContainKey("merge");
    }

    // =================================================================
    // Excel validation Set properties
    // =================================================================

    [Fact]
    public void Bug_Excel_Validation_Set_Prompt()
    {
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "F1:F10", ["type"] = "list", ["formula1"] = "A,B,C"
        });
        _excelHandler.Set("/Sheet1/validation[1]", new()
        {
            ["showinput"] = "true",
            ["prompttitle"] = "Choose",
            ["prompt"] = "Select a value"
        });

        var node = _excelHandler.Get("/Sheet1/validation[1]");
        node.Format["promptTitle"]?.ToString().Should().Be("Choose");
        node.Format["prompt"]?.ToString().Should().Be("Select a value");
    }

    // =================================================================
    // Excel validation list formula — no double quoting
    // =================================================================

    [Fact]
    public void Bug_Excel_Validation_List_Formula_NoDoubleQuote()
    {
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "G1:G10", ["type"] = "list",
            ["formula1"] = "Red,Green,Blue"
        });

        _excelHandler.Get("/Sheet1/validation[1]").Format["formula1"]?.ToString()
            .Should().Be("Red,Green,Blue");
    }

    [Fact]
    public void Bug_Excel_Validation_AlreadyQuoted_Formula()
    {
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "H1:H10", ["type"] = "list",
            ["formula1"] = "\"Yes,No,Maybe\""
        });

        _excelHandler.Get("/Sheet1/validation[1]").Format["formula1"]?.ToString()
            .Should().Be("Yes,No,Maybe");
    }

    // =================================================================
    // Excel multiple cells batch formatting
    // =================================================================

    [Fact]
    public void Bug_Excel_Multiple_Cells_Formatting()
    {
        for (int i = 1; i <= 5; i++)
            _excelHandler.Set($"/Sheet1/A{i}", new()
            {
                ["value"] = $"Row {i}", ["font.bold"] = "true"
            });

        for (int i = 1; i <= 5; i++)
        {
            var node = _excelHandler.Get($"/Sheet1/A{i}");
            node.Text.Should().Be($"Row {i}");
            node.Format.Should().ContainKey("font.bold");
        }
    }

    // =================================================================
    // CONFIRMED BUG: Excel Query shorthand — "cell:text" syntax
    // Fixed: added shorthand regex in ParseCellSelector.
    // =================================================================

    [Fact]
    public void Bug_Excel_Query_Cell_ByText_Shorthand()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Apple" });
        _excelHandler.Set("/Sheet1/A2", new() { ["value"] = "Banana" });
        _excelHandler.Set("/Sheet1/A3", new() { ["value"] = "Cherry" });

        var results = _excelHandler.Query("cell:Banana");
        results.Should().NotBeEmpty("cell:text shorthand should find matching cells");
        results[0].Text.Should().Be("Banana");
    }
}
