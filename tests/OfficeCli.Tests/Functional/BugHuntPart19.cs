// Bug hunt Part 19 — more property readback gaps, Set/Get inconsistencies, edge cases.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart19 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public BugHuntPart19()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt19_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt19_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt19_{Guid.NewGuid():N}.pptx");
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


    // ==================== BUG #1: Word run Get doesn't include link (hyperlink wrapper) ====================
    // After wrapping a run in a hyperlink via Set, Get on the run should show the link.
    [Fact]
    public void Word_Run_Get_ShouldIncludeLink()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Click here"
        });

        _wordHandler.Set("/body/p[1]/r[1]", new()
        {
            ["link"] = "https://example.com"
        });

        var run = _wordHandler.Get("/body/p[1]/r[1]");
        run.Should().NotBeNull();

        // BUG: After setting a link, the run is wrapped in a Hyperlink element,
        // but Get on the run doesn't report the link URL
        run.Format.Should().ContainKey("link",
            "run Get should include link URL when the run is wrapped in a hyperlink");
    }


    // ==================== BUG #2: PPTX shape italic=false not reported ====================
    // Same as bold=false bug: italic only reports when true, not when explicitly false
    [Fact]
    public void Pptx_Shape_Italic_False_ShouldBeInFormat()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["italic"] = "true"
        });

        var shape1 = pptx.Get("/slide[1]/shape[1]");
        shape1.Format.Should().ContainKey("italic");

        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["italic"] = "false"
        });

        var shape2 = pptx.Get("/slide[1]/shape[1]");
        // BUG: italic=false is not reported, making it impossible to verify the override
        shape2.Format.Should().ContainKey("italic",
            "explicitly setting italic=false should be reported in Format");
    }


    // ==================== BUG #4: Word comment Get doesn't include author ====================
    [Fact]
    public void Word_Comment_Get_ShouldIncludeAuthor()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Commented text"
        });

        _wordHandler.Add("/body/p[1]", "comment", null, new()
        {
            ["text"] = "This is a comment",
            ["author"] = "TestAuthor"
        });

        // Query for comments
        var comments = _wordHandler.Query("comment");
        comments.Should().NotBeEmpty();

        var comment = comments[0];
        comment.Format.Should().ContainKey("author",
            "comment Get/Query should include the author name");
    }


    // ==================== BUG #5: PPTX shape text with literal \n should create multi-line ====================
    // The Add handler should interpret \\n as literal newline for multi-paragraph text.
    [Fact]
    public void Pptx_AddShape_TextWithNewlines()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line1\\nLine2"
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        shape.Should().NotBeNull();

        // The text should contain both lines
        shape.Text.Should().Contain("Line1");
        shape.Text.Should().Contain("Line2");
    }


    // ==================== BUG #6: Word table row height not in Get ====================
    [Fact]
    public void Word_TableRow_Get_ShouldIncludeHeight()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        // Add a row with specific height
        _wordHandler.Add("/body/tbl[1]", "row", null, new()
        {
            ["height"] = "720"  // 720 twips = 0.5 inch
        });

        var table = _wordHandler.Get("/body/tbl[1]", depth: 2);
        table.Children.Count.Should().BeGreaterThanOrEqualTo(2);

        // The second row (added) should report its height
        var row2 = table.Children[1];
        row2.Format.Should().ContainKey("height",
            "table row Get should expose height when it's been set");
    }


    // ==================== BUG #8: PPTX slide Set transition then Get should read it back ====================
    [Fact]
    public void Pptx_Slide_Get_ShouldIncludeTransition()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Set("/slide[1]", new()
        {
            ["transition"] = "fade"
        });

        // Read back in the same session (persistence across reopen is a separate SDK limitation)
        var slide = pptx.Get("/slide[1]");
        slide.Should().NotBeNull();

        slide.Format.Should().ContainKey("transition",
            "slide Get should include transition type after Set in the same session");
        slide.Format["transition"]?.ToString().Should().Be("fade");
    }


    // ==================== BUG #10: Excel cell alignment not in Get ====================
    [Fact]
    public void Excel_Cell_Alignment_RoundTrip()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Centered",
            ["halign"] = "center"
        });

        var cell = _excelHandler.Get("/Sheet1/A1");
        cell.Should().NotBeNull();

        cell.Format.Should().ContainKey("halign",
            "cell Get should include horizontal alignment when it's been set");
    }


    // ==================== BUG #11: Word table Get missing style property ====================
    [Fact]
    public void Word_Table_Get_ShouldIncludeStyle()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        _wordHandler.Set("/body/tbl[1]", new()
        {
            ["style"] = "TableGrid"
        });

        var table = _wordHandler.Get("/body/tbl[1]");
        table.Should().NotBeNull();

        table.Format.Should().ContainKey("style",
            "table Get should include style when one is set");
    }
}
