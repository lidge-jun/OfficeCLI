// Bug hunt Part 28 — Word handler & PPTX handler confirmed bugs and edge cases:
// dstrike missing from paragraph Add, alignment "both" not accepted, font size
// 10.5pt edge case, table cell formatting, run property roundtrips, paragraph
// properties, PPTX lineSpacing Add, text handling, element operations.

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart28 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private PowerPointHandler _pptxHandler;

    public BugHuntPart28()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt28_{Guid.NewGuid():N}.docx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt28_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_pptxPath);
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _pptxHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private WordHandler ReopenWord()
    {
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
        return _wordHandler;
    }

    private PowerPointHandler ReopenPptx()
    {
        _pptxHandler.Dispose();
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
        return _pptxHandler;
    }

    // =================================================================
    // CONFIRMED BUG: Word paragraph Add — dstrike missing from inline
    // run properties. Only available via "run" Add, not "paragraph" Add.
    // Fixed: added dstrike to paragraph Add run property block.
    // =================================================================

    [Fact]
    public void Bug_Word_Paragraph_Add_DStrike_Roundtrip()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Double strike",
            ["dstrike"] = "true"
        });

        var node = _wordHandler.Get("/body/p[1]", depth: 2);
        node.Children[0].Format.Should().ContainKey("dstrike",
            "dstrike via paragraph Add should be readable in Get");
    }

    // =================================================================
    // CONFIRMED BUG: Word Add/Set alignment="both" not accepted.
    // "both" is the OOXML name for justify; should be a synonym.
    // Fixed: added "both" to all alignment switch expressions.
    // =================================================================

    [Fact]
    public void Bug_Word_Paragraph_Alignment_Both_Via_Add()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Justified",
            ["alignment"] = "both"
        });

        var node = _wordHandler.Get("/body/p[1]");
        node.Format.Should().ContainKey("alignment");
        var align = node.Format["alignment"]?.ToString();
        (align == "justify" || align == "both").Should().BeTrue(
            "alignment 'both' should be accepted and stored (normalize to 'justify' or keep 'both')");
    }

    [Fact]
    public void Bug_Word_Paragraph_Alignment_Both_Via_Set()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]", new() { ["alignment"] = "both" });

        var node = _wordHandler.Get("/body/p[1]");
        var align = node.Format["alignment"]?.ToString();
        (align == "justify" || align == "both").Should().BeTrue(
            "Set alignment=both should produce 'justify' or 'both' in Get");
    }

    // =================================================================
    // EDGE CASE: Word font size 10.5pt — common in Chinese documents.
    // 21 half-points must round-trip through int→double division.
    // =================================================================

    [Fact]
    public void Bug_Word_Run_FontSize_HalfPoint_Roundtrip()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "五号字体",
            ["size"] = "10.5"
        });

        var node = _wordHandler.Get("/body/p[1]", depth: 2);
        node.Children[0].Format["size"]?.ToString().Should().Be("10.5pt",
            "10.5pt (五号) should round-trip correctly");
    }

    // =================================================================
    // EDGE CASE: Word table cell Set text preserves bold formatting
    // =================================================================

    [Fact]
    public void Bug_Word_TableCell_SetText_PreservesFormatting()
    {
        _wordHandler.Add("/body", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Initial", ["bold"] = "true"
        });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Updated"
        });

        var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]", depth: 3);
        node.Text.Should().Be("Updated");
        node.Children[0].Children[0].Format.Should().ContainKey("bold",
            "bold should survive text replacement");
    }

    // =================================================================
    // EDGE CASE: Word run bold=false / italic=false should REMOVE
    // =================================================================

    [Fact]
    public void Bug_Word_Run_Bold_False_Removes()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Bold", ["bold"] = "true"
        });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["bold"] = "false" });

        var node = _wordHandler.Get("/body/p[1]", depth: 2);
        node.Children[0].Format.ContainsKey("bold").Should().BeFalse(
            "bold=false should remove bold, not keep it");
    }

    [Fact]
    public void Bug_Word_Run_Italic_False_Removes()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Italic", ["italic"] = "true"
        });
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["italic"] = "false" });

        var node = _wordHandler.Get("/body/p[1]", depth: 2);
        node.Children[0].Format.ContainsKey("italic").Should().BeFalse(
            "italic=false should remove italic");
    }

    // =================================================================
    // Word run properties — consolidated test for all rare properties
    // via "run" Add (dstrike tested separately via paragraph Add above)
    // =================================================================

    [Fact]
    public void Bug_Word_Run_Add_AllRareProperties()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "base" });

        var props = new (string key, string formatKey)[]
        {
            ("vanish", "vanish"), ("outline", "outline"), ("shadow", "shadow"),
            ("emboss", "emboss"), ("imprint", "imprint"), ("noproof", "noproof"),
            ("rtl", "rtl"), ("caps", "caps"), ("smallcaps", "smallcaps")
        };

        foreach (var (key, formatKey) in props)
        {
            _wordHandler.Add("/body/p[1]", "run", null, new()
            {
                ["text"] = key, [key] = "true"
            });
        }

        var node = _wordHandler.Get("/body/p[1]", depth: 2);
        // Skip first run ("base"), check each subsequent run
        for (int i = 0; i < props.Length; i++)
        {
            var run = node.Children[i + 1];
            run.Format.Should().ContainKey(props[i].formatKey,
                $"run property '{props[i].key}' should be readable in Get");
        }
    }

    [Fact]
    public void Bug_Word_Run_Superscript_Subscript_Add()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "H" });
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "2", ["subscript"] = "true"
        });
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "3", ["superscript"] = "true"
        });

        var node = _wordHandler.Get("/body/p[1]", depth: 2);
        node.Children[1].Format.Should().ContainKey("subscript");
        node.Children[2].Format.Should().ContainKey("superscript");
    }

    // =================================================================
    // Word table properties — padding, width, borders, colWidths
    // =================================================================

    [Fact]
    public void Bug_Word_Table_Padding_Persists()
    {
        _wordHandler.Add("/body", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["padding"] = "200"
        });
        ReopenWord();

        var node = _wordHandler.Get("/body/tbl[1]", depth: 0);
        (node.Format.ContainsKey("padding.left") || node.Format.ContainsKey("padding.right"))
            .Should().BeTrue("table padding should persist after reopen");
    }

    [Fact]
    public void Bug_Word_Table_Width_And_ColWidths()
    {
        _wordHandler.Add("/body", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "3", ["width"] = "100%"
        });

        var node = _wordHandler.Get("/body/tbl[1]", depth: 0);
        node.Format.Should().ContainKey("colWidths");
        node.Format["colWidths"]?.ToString()!.Split(',').Length.Should().Be(3,
            "colWidths should have 3 values for 3-column table");
    }

    [Fact]
    public void Bug_Word_Table_Border_Add()
    {
        _wordHandler.Add("/body", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["border"] = "single;4;000000"
        });

        var node = _wordHandler.Get("/body/tbl[1]", depth: 0);
        (node.Format.ContainsKey("border.top") || node.Format.ContainsKey("border.bottom"))
            .Should().BeTrue("table borders from Add should be readable");
    }

    // =================================================================
    // Word table cell properties
    // =================================================================

    [Fact]
    public void Bug_Word_TableCell_Properties_Roundtrip()
    {
        _wordHandler.Add("/body", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "4"
        });

        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["shd"] = "FF9900", ["valign"] = "center",
            ["gridspan"] = "2", ["nowrap"] = "true",
            ["border.bottom"] = "single;8;FF0000"
        });

        var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Format["shd"]?.ToString().Should().Be("#FF9900");
        node.Format["valign"]?.ToString().Should().Be("center");
        int.Parse(node.Format["gridspan"]!.ToString()!).Should().Be(2);
        node.Format.Should().ContainKey("nowrap");
        node.Format["border.bottom"]?.ToString().Should().Contain("single");
    }

    [Fact]
    public void Bug_Word_TableCell_TextDirection_Set()
    {
        _wordHandler.Add("/body", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["textDirection"] = "btLr"
        });

        var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("textDirection");
    }

    // =================================================================
    // Word table row height and header
    // =================================================================

    [Fact]
    public void Bug_Word_TableRow_Height_And_Header()
    {
        _wordHandler.Add("/body", "table", null, new()
        {
            ["rows"] = "3", ["cols"] = "2"
        });
        _wordHandler.Set("/body/tbl[1]/tr[1]", new()
        {
            ["height"] = "500", ["header"] = "true"
        });

        var node = _wordHandler.Get("/body/tbl[1]", depth: 1);
        var row = node.Children[0];
        row.Format.Should().ContainKey("height");
        row.Format.Should().ContainKey("header");
    }

    // =================================================================
    // Word paragraph properties — consolidated
    // =================================================================

    [Fact]
    public void Bug_Word_Paragraph_AllProperties_Add()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Full featured paragraph",
            ["alignment"] = "center",
            ["linespacing"] = "360",
            ["spacebefore"] = "240",
            ["spaceafter"] = "120",
            ["leftindent"] = "720",
            ["rightindent"] = "360",
            ["hangingindent"] = "480",
            ["keepnext"] = "true",
            ["keeplines"] = "true",
            ["pagebreakbefore"] = "true",
            ["widowcontrol"] = "true",
            ["highlight"] = "yellow",
            ["color"] = "#FF0000"
        });

        var node = _wordHandler.Get("/body/p[1]");
        node.Format["alignment"]?.ToString().Should().Be("center");
        node.Format["lineSpacing"]?.ToString().Should().Be("1.5x");
        node.Format.Should().ContainKey("spaceBefore");
        node.Format.Should().ContainKey("spaceAfter");
        node.Format["leftindent"]?.ToString().Should().Be("720");
        node.Format["rightindent"]?.ToString().Should().Be("360");
        node.Format["hangingindent"]?.ToString().Should().Be("480");
        node.Format.Should().ContainKey("keepnext");
        node.Format.Should().ContainKey("keeplines");
        node.Format.Should().ContainKey("pagebreakbefore");
        node.Format.Should().ContainKey("widowcontrol");

        var run = _wordHandler.Get("/body/p[1]", depth: 2).Children[0];
        run.Format["highlight"]?.ToString().Should().Be("yellow");
        run.Format["color"]?.ToString().Should().Be("#FF0000");
    }

    [Fact]
    public void Bug_Word_Paragraph_Style_Set()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Heading", ["style"] = "Heading1"
        });
        var node = _wordHandler.Get("/body/p[1]");
        node.Format["style"]?.ToString().Should().Be("Heading1");
    }

    [Fact]
    public void Bug_Word_Paragraph_ListStyle_Add()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Bullet", ["liststyle"] = "bullet"
        });
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Ordered", ["liststyle"] = "ordered"
        });

        _wordHandler.Get("/body/p[1]").Format["listStyle"]?.ToString().Should().Be("bullet");
        _wordHandler.Get("/body/p[2]").Format["listStyle"]?.ToString().Should().Be("ordered");
    }

    [Fact]
    public void Bug_Word_Paragraph_Shd_Set()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "shaded" });
        _wordHandler.Set("/body/p[1]", new() { ["shd"] = "CCFFCC" });

        var node = _wordHandler.Get("/body/p[1]");
        (node.Format.ContainsKey("shd") || node.Format.ContainsKey("shading")).Should().BeTrue(
            "paragraph shading should be readable via 'shd' or 'shading' key");
    }

    // =================================================================
    // Word element operations — Remove path shift, Move
    // =================================================================

    [Fact]
    public void Bug_Word_Paragraph_Remove_PathShift()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "First" });
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Second" });
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Third" });

        _wordHandler.Remove("/body/p[2]");
        _wordHandler.Get("/body/p[2]").Text.Should().Be("Third",
            "after removing p[2], Third should shift to p[2]");
    }

    [Fact]
    public void Bug_Word_TableRow_Remove_PathShift()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Row1" });
        _wordHandler.Set("/body/tbl[1]/tr[2]/tc[1]", new() { ["text"] = "Row2" });
        _wordHandler.Set("/body/tbl[1]/tr[3]/tc[1]", new() { ["text"] = "Row3" });

        _wordHandler.Remove("/body/tbl[1]/tr[2]");
        _wordHandler.Get("/body/tbl[1]/tr[2]/tc[1]").Text.Should().Be("Row3");
    }

    [Fact]
    public void Bug_Word_Paragraph_Move()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "First" });
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Second" });
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Third" });

        _wordHandler.Move("/body/p[3]", "/body", 0);
        _wordHandler.Get("/body/p[1]").Text.Should().Be("Third",
            "after move to position 0 (0-based), Third should be at p[1]");
    }

    // =================================================================
    // CONFIRMED BUG: PPTX shape Add — lineSpacing not supported.
    // Fixed: added lineSpacing/spaceBefore/spaceAfter to Add.
    // =================================================================

    [Fact]
    public void Bug_Pptx_Shape_Add_LineSpacing()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Double spaced",
            ["lineSpacing"] = "2.0"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineSpacing",
            "lineSpacing from Add should be readable in Get");
        node.Format["lineSpacing"]?.ToString().Should().Be("2x");
    }

    // =================================================================
    // PPTX shape text with newlines and multiline replace
    // =================================================================

    [Fact]
    public void Bug_Pptx_Shape_Text_Newline_Creates_Paragraphs()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line1\\nLine2\\nLine3"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Text.Should().Contain("Line3");
        node.ChildCount.Should().Be(3, "3 lines should create 3 paragraphs");
    }

    [Fact]
    public void Bug_Pptx_Shape_SetText_Multiline_To_Single()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line1\\nLine2\\nLine3"
        });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["text"] = "Single" });

        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Single");
        node.ChildCount.Should().Be(1, "single line should produce 1 paragraph");
    }

    // =================================================================
    // PPTX shape Add with all properties at once
    // =================================================================

    [Fact]
    public void Bug_Pptx_Shape_Add_AllProperties()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Full featured", ["fill"] = "FFCC00",
            ["font"] = "Arial", ["size"] = "18",
            ["bold"] = "true", ["italic"] = "true",
            ["color"] = "000080", ["align"] = "right",
            ["valign"] = "bottom", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "8cm", ["height"] = "4cm"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Full featured");
        node.Format["fill"]?.ToString().Should().Be("#FFCC00");
        node.Format["font"]?.ToString().Should().Be("Arial");
        node.Format["bold"].Should().NotBeNull();
        node.Format.Should().ContainKey("align");
        node.Format.Should().ContainKey("valign");
    }

    // =================================================================
    // PPTX shape text persistence after reopen
    // =================================================================

    [Fact]
    public void Bug_Pptx_Shape_Text_Persists_After_Reopen()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Persistent", ["bold"] = "true", ["color"] = "FF0000"
        });
        ReopenPptx();

        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Persistent");
        node.Format.Should().ContainKey("bold");
        node.Format.Should().ContainKey("color");
    }

    // =================================================================
    // PPTX element operations — Remove, Swap, CopyFrom
    // =================================================================

    [Fact]
    public void Bug_Pptx_Shape_Remove_PathShift()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "C" });

        _pptxHandler.Remove("/slide[1]/shape[2]");
        _pptxHandler.Get("/slide[1]/shape[2]").Text.Should().Be("C");
    }

    [Fact]
    public void Bug_Pptx_Slide_Swap()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Slide A" });
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[2]", "shape", null, new() { ["text"] = "Slide B" });

        _pptxHandler.Swap("/slide[1]", "/slide[2]");
        _pptxHandler.Get("/slide[1]/shape[1]").Text.Should().Be("Slide B");
        _pptxHandler.Get("/slide[2]/shape[1]").Text.Should().Be("Slide A");
    }

    [Fact]
    public void Bug_Pptx_Shape_CopyFrom()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original", ["fill"] = "FF0000"
        });

        _pptxHandler.CopyFrom("/slide[1]/shape[1]", "/slide[1]", null);
        var copy = _pptxHandler.Get("/slide[1]/shape[2]");
        copy.Text.Should().Be("Original");
        copy.Format.Should().ContainKey("fill");
    }

    [Fact]
    public void Bug_Pptx_Slide_Remove_PathShift()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "S1" });
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[2]", "shape", null, new() { ["text"] = "S2" });
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[3]", "shape", null, new() { ["text"] = "S3" });

        _pptxHandler.Remove("/slide[2]");
        _pptxHandler.Get("/slide[2]/shape[1]").Text.Should().Be("S3");
    }

    // =================================================================
    // PPTX Query shorthand — "shape:text" syntax
    // =================================================================

    [Fact]
    public void Bug_Pptx_Query_Shape_ByText_Shorthand()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Find me" });
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Not this" });

        var results = _pptxHandler.Query("shape:Find me");
        results.Should().NotBeEmpty("shape:text shorthand should find matching shapes");
    }
}
