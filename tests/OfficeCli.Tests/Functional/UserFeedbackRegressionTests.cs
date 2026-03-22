// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Regression tests from user feedback during deep-review testing.
/// Each test documents a real bug reported by users.
/// </summary>
public class UserFeedbackRegressionTests : IDisposable
{
    private readonly string _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
    private readonly string _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
    private readonly string _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");

    // ==================== tblLook schema validation ====================
    // Bug: add table --prop style=TableGrid generated w:tblLook with Office 2010+
    // boolean attributes (firstRow, lastRow, etc.) that caused 6 schema validation errors.

    [Fact]
    public void DocxAddTable_WithStyle_ValidatesCleanly()
    {
        BlankDocCreator.Create(_docxPath);
        using var handler = new WordHandler(_docxPath, editable: true);

        handler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2", ["style"] = "TableGrid" });

        var errors = handler.Validate();
        errors.Should().BeEmpty("table with style=TableGrid should not produce tblLook validation errors");
    }

    // ==================== Transition OpenXmlUnknownElement ====================
    // Bug: add slide with transition=fade inserted transition as OpenXmlUnknownElement
    // (workaround for old SDK bug). This caused validate to report
    // "invalid type 'OpenXmlUnknownElement'" on the transition element.

    [Fact]
    public void PptxAddSlide_WithTransition_ValidatesCleanly()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);

        handler.Add("/", "slide", null, new() { ["transition"] = "fade" });

        var errors = handler.Validate();
        errors.Where(e => e.Description.Contains("transition", StringComparison.OrdinalIgnoreCase))
            .Should().BeEmpty("transition=fade should use typed SDK element, not OpenXmlUnknownElement");
    }

    [Fact]
    public void PptxSetTransition_OnExistingSlide_ValidatesCleanly()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "dissolve" });

        var errors = handler.Validate();
        errors.Where(e => e.Description.Contains("transition", StringComparison.OrdinalIgnoreCase))
            .Should().BeEmpty();
    }

    [Theory]
    [InlineData("fade")]
    [InlineData("wipe-left")]
    [InlineData("push-right")]
    [InlineData("circle")]
    [InlineData("zoom-out")]
    public void PptxAddSlide_VariousTransitions_ValidateCleanly(string transition)
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);

        handler.Add("/", "slide", null, new() { ["transition"] = transition });

        var errors = handler.Validate();
        errors.Where(e => e.Description.Contains("transition", StringComparison.OrdinalIgnoreCase))
            .Should().BeEmpty($"transition={transition} should validate cleanly");
    }

    // ==================== Malformed selector detection ====================
    // Bug: query with paragraph[style===Heading1] silently returned 0 matches
    // instead of reporting a syntax error. Agent couldn't distinguish
    // "wrong syntax" from "no matching elements".

    [Fact]
    public void Parse_TripleEquals_ThrowsCliException()
    {
        var act = () => AttributeFilter.Parse("paragraph[style===Heading1]");
        act.Should().Throw<CliException>()
            .Which.Code.Should().Be("invalid_selector");
    }

    [Fact]
    public void Parse_UnclosedBracket_ThrowsCliException()
    {
        var act = () => AttributeFilter.Parse("paragraph[style");
        act.Should().Throw<CliException>()
            .Which.Code.Should().Be("invalid_selector");
    }

    [Fact]
    public void Parse_EmptyBrackets_ThrowsCliException()
    {
        var act = () => AttributeFilter.Parse("paragraph[]");
        act.Should().Throw<CliException>()
            .Which.Code.Should().Be("invalid_selector");
    }

    [Fact]
    public void Parse_ValidSelector_DoesNotThrow()
    {
        var conditions = AttributeFilter.Parse("paragraph[style=Heading 1]");
        conditions.Should().HaveCount(1);
        conditions[0].Key.Should().Be("style");
        conditions[0].Op.Should().Be(AttributeFilter.FilterOp.Equal);
        conditions[0].Value.Should().Be("Heading 1");
    }

    // ==================== [attr] exists filter ====================
    // Bug: cell[formula] returned all cells instead of only those with formulas.
    // The [formula] without an operator was silently ignored.

    [Fact]
    public void Parse_HasAttribute_ReturnsExistsCondition()
    {
        var conditions = AttributeFilter.Parse("cell[formula]");
        conditions.Should().HaveCount(1);
        conditions[0].Key.Should().Be("formula");
        conditions[0].Op.Should().Be(AttributeFilter.FilterOp.Exists);
    }

    [Fact]
    public void ExistsFilter_MatchesOnlyNodesWithAttribute()
    {
        var nodes = new List<DocumentNode>
        {
            new() { Path = "/A1", Type = "cell", Format = new() { ["formula"] = "SUM(B1:B5)" } },
            new() { Path = "/A2", Type = "cell", Format = new() { ["value"] = "hello" } },
            new() { Path = "/A3", Type = "cell", Format = new() { ["formula"] = "A1+A2" } },
        };

        var conditions = new List<AttributeFilter.Condition>
        {
            new("formula", AttributeFilter.FilterOp.Exists, "")
        };

        var result = AttributeFilter.Apply(nodes, conditions);
        result.Should().HaveCount(2);
        result.Select(n => n.Path).Should().BeEquivalentTo("/A1", "/A3");
    }

    [Fact]
    public void ExistsFilter_DoesNotMatchEmptyValue()
    {
        var nodes = new List<DocumentNode>
        {
            new() { Path = "/A1", Type = "cell", Format = new() { ["formula"] = "" } },
            new() { Path = "/A2", Type = "cell", Format = new() { ["formula"] = "SUM(B1:B5)" } },
        };

        var conditions = new List<AttributeFilter.Condition>
        {
            new("formula", AttributeFilter.FilterOp.Exists, "")
        };

        var result = AttributeFilter.Apply(nodes, conditions);
        result.Should().HaveCount(1);
        result[0].Path.Should().Be("/A2");
    }

    [Fact]
    public void ExcelQuery_CellWithFormula_OnlyReturnsFormulaCells()
    {
        BlankDocCreator.Create(_xlsxPath);
        using var handler = new ExcelHandler(_xlsxPath, editable: true);

        handler.Set("/Sheet1/A1", new() { ["value"] = "Name" });
        handler.Set("/Sheet1/B1", new() { ["value"] = "100" });
        handler.Set("/Sheet1/C1", new() { ["formula"] = "B1*2" });

        var allCells = handler.Query("cell");
        allCells.Should().HaveCount(3);

        // Use exists filter
        var conditions = AttributeFilter.Parse("cell[formula]");
        var filtered = AttributeFilter.Apply(allCells, conditions);
        filtered.Should().HaveCount(1);
        filtered[0].Path.Should().Contain("C1");
    }

    // ==================== Unsupported prop returns success:false ====================
    // Bug: set with unsupported property returned success:true with "Updated" message.
    // The unsupported list contained help text like "key (valid props: ...)"
    // so Contains(key) never matched, making all props appear "applied".

    [Fact]
    public void DocxSet_AllUnsupported_ReturnsUnsupportedInList()
    {
        BlankDocCreator.Create(_docxPath);
        using var handler = new WordHandler(_docxPath, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        var unsupported = handler.Set("/body/p[1]", new() { ["noSuchProp"] = "1" });
        unsupported.Should().NotBeEmpty("unknown property should be reported as unsupported");
        // The first unsupported entry contains help text — verify the key is at the start
        unsupported[0].Should().StartWith("noSuchProp");
    }

    [Fact]
    public void DocxSet_MixedProps_OnlyValidPropsApplied()
    {
        BlankDocCreator.Create(_docxPath);
        using var handler = new WordHandler(_docxPath, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        var unsupported = handler.Set("/body/p[1]", new() { ["bold"] = "true", ["noSuchProp"] = "1" });
        unsupported.Should().NotBeEmpty();

        // bold should have been applied
        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("bold");
    }

    // ==================== Structured JSON view output ====================
    // Bug: view stats/outline/text --json wrapped human-readable text in a Content
    // string field instead of returning structured JSON fields.

    [Fact]
    public void PptxViewAsStatsJson_ReturnsStructuredFields()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Test" });

        var json = handler.ViewAsStatsJson();
        var obj = json.AsObject();
        obj.Should().ContainKey("slides");
        obj.Should().ContainKey("totalShapes");
        obj.Should().ContainKey("textBoxes");
        obj.Should().ContainKey("pictures");
        obj["slides"]!.GetValue<int>().Should().Be(1);
    }

    [Fact]
    public void PptxViewAsOutlineJson_ReturnsSlideArray()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Intro" });
        handler.Add("/", "slide", null, new() { ["title"] = "Details" });

        var json = handler.ViewAsOutlineJson();
        var obj = json.AsObject();
        obj.Should().ContainKey("slides");
        obj.Should().ContainKey("totalSlides");
        obj["totalSlides"]!.GetValue<int>().Should().Be(2);

        var slides = obj["slides"]!.AsArray();
        slides.Should().HaveCount(2);
        slides[0]!.AsObject().Should().ContainKey("index");
        slides[0]!.AsObject().Should().ContainKey("title");
    }

    [Fact]
    public void DocxViewAsStatsJson_ReturnsStructuredFields()
    {
        BlankDocCreator.Create(_docxPath);
        using var handler = new WordHandler(_docxPath, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello World" });

        var json = handler.ViewAsStatsJson();
        var obj = json.AsObject();
        obj.Should().ContainKey("paragraphs");
        obj.Should().ContainKey("totalCharacters");
        obj.Should().ContainKey("styleDistribution");
    }

    [Fact]
    public void ExcelViewAsOutlineJson_ReturnsSheetsArray()
    {
        BlankDocCreator.Create(_xlsxPath);
        using var handler = new ExcelHandler(_xlsxPath, editable: true);

        var json = handler.ViewAsOutlineJson();
        var obj = json.AsObject();
        obj.Should().ContainKey("fileName");
        obj.Should().ContainKey("sheets");

        var sheets = obj["sheets"]!.AsArray();
        sheets.Should().HaveCountGreaterOrEqualTo(1);
        sheets[0]!.AsObject().Should().ContainKey("name");
        sheets[0]!.AsObject().Should().ContainKey("rows");
        sheets[0]!.AsObject().Should().ContainKey("cols");
    }

    [Fact]
    public void ExcelViewAsTextJson_ReturnsCellObjects()
    {
        BlankDocCreator.Create(_xlsxPath);
        using var handler = new ExcelHandler(_xlsxPath, editable: true);

        handler.Set("/Sheet1/A1", new() { ["value"] = "Name" });
        handler.Set("/Sheet1/B1", new() { ["value"] = "Score" });

        var json = handler.ViewAsTextJson();
        var obj = json.AsObject();
        obj.Should().ContainKey("sheets");

        var sheets = obj["sheets"]!.AsArray();
        var rows = sheets[0]!.AsObject()["rows"]!.AsArray();
        rows.Should().HaveCount(1);
        rows[0]!.AsObject()["cells"]!.AsObject().Should().ContainKey("A1");
    }

    public void Dispose()
    {
        try { File.Delete(_docxPath); } catch { }
        try { File.Delete(_pptxPath); } catch { }
        try { File.Delete(_xlsxPath); } catch { }
    }
}
