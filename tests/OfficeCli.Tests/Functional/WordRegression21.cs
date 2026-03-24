// Bug hunt Part 21 — more Set/Get asymmetries, edge cases in all three handlers.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class WordRegression21 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public WordRegression21()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt21_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt21_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt21_{Guid.NewGuid():N}.pptx");
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


    // ==================== BUG #1: Word run Get doesn't include shading ====================
    [Fact]
    public void Word_Run_Get_ShouldIncludeShading()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Highlighted",
            ["shading"] = "FFFF00"
        });

        var run = _wordHandler.Get("/body/p[1]/r[1]");
        run.Should().NotBeNull();

        // Shading was set on the paragraph during Add which applies to the run
        // But does the run have shading in Format?
        run.Format.Should().ContainKey("shading",
            "run Get should include shading/background color when it's set");
    }


    // ==================== BUG #2: PPTX slide background not persisted correctly ====================
    [Fact]
    public void Pptx_Slide_Background_Persistence()
    {
        using (var pptx = new PowerPointHandler(_pptxPath, editable: true))
        {
            pptx.Set("/slide[1]", new()
            {
                ["background"] = "003366"
            });
        }

        // Reopen
        using var pptx2 = new PowerPointHandler(_pptxPath, editable: true);
        var slide = pptx2.Get("/slide[1]");

        slide.Format.Should().ContainKey("background",
            "slide background color should persist after file close and reopen");

        slide.Format["background"]?.ToString().Should().Contain("003366",
            "background color value should survive reopen");
    }


    // ==================== BUG #3: Excel cell font.bold readback value ====================
    // Font properties use bare true/false in Get but Format stores as boxed bool.
    [Fact]
    public void Excel_Cell_FontBold_Get_ValueType()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Bold",
            ["font.bold"] = "true"
        });

        var cell = _excelHandler.Get("/Sheet1/A1");
        cell.Format.Should().ContainKey("font.bold");

        // The value should be a boolean true, not a string "true"
        var boldVal = cell.Format["font.bold"];
        boldVal.Should().NotBeNull();
        // Verify it can be used in boolean comparisons
        (boldVal is true || boldVal?.ToString() == "True" || boldVal?.ToString() == "true")
            .Should().BeTrue("font.bold should be truthy");
    }


    // ==================== BUG #4: Word paragraph leftindent not in Get ====================
    [Fact]
    public void Word_Paragraph_Get_ShouldIncludeLeftIndent()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Indented",
            ["leftindent"] = "720"
        });

        var para = _wordHandler.Get("/body/p[1]");
        para.Format.Should().ContainKey("leftIndent",
            "paragraph Get should expose leftIndent when it's set");
    }


    // ==================== BUG #5: PPTX shape Get missing margin/padding info ====================
    [Fact]
    public void Pptx_Shape_Get_ShouldIncludeMargin()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Padded",
            ["margin"] = "0.5cm"
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        shape.Format.Should().ContainKey("margin",
            "shape Get should include margin/padding when it's been set");
    }


    // ==================== BUG #7: Word run highlight readback value format ====================
    // Set uses color name like "yellow", but Get reports it as the enum InnerText
    [Fact]
    public void Word_Run_Highlight_ValueFormat()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Marked"
        });

        _wordHandler.Set("/body/p[1]/r[1]", new()
        {
            ["highlight"] = "yellow"
        });

        var run = _wordHandler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("highlight");

        var hl = run.Format["highlight"]?.ToString();
        hl.Should().Be("yellow",
            "highlight value should be 'yellow' (lowercase), matching the Set input format");
    }


    // ==================== BUG #8: PPTX shape shadow not in Get Format ====================
    [Fact]
    public void Pptx_Shape_Get_ShouldIncludeShadow()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shadow"
        });

        pptx.Set("/slide[1]/shape[1]", new()
        {
            ["shadow"] = "000000"
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        shape.Format.Should().ContainKey("shadow",
            "shape Get should include shadow when it's been set");
    }


    // ==================== BUG #9: Word section margins not in Get ====================
    [Fact]
    public void Word_Section_Get_ShouldIncludeMargins()
    {
        _wordHandler.Set("/section[1]", new()
        {
            ["margintop"] = "1440",
            ["marginbottom"] = "1440"
        });

        var section = _wordHandler.Get("/section[1]");
        section.Format.Should().ContainKey("margintop",
            "section Get should include margins when they've been set");
    }


}
