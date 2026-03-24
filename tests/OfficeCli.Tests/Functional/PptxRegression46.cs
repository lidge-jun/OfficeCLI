// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class PptxRegression46 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}{ext}");
        _tempFiles.Add(p);
        return p;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // =====================================================================
    // Bug4600: PPTX table cell text roundtrip — set text on cell,
    // verify it reads back correctly
    // =====================================================================
    [Fact]
    public void Bug4600_Pptx_Table_Cell_Text_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Hello Cell"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = node.Children[0].Children[0]; // tr[1]/tc[1]
        cellNode.Text.Should().Be("Hello Cell",
            because: "table cell text should roundtrip after Set");
    }

    // =====================================================================
    // Bug4601: PPTX table cell bold roundtrip
    // =====================================================================
    [Fact]
    public void Bug4601_Pptx_Table_Cell_Bold_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold Cell", ["bold"] = "true"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = node.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("bold",
            because: "bold should be readable on table cell after Set");
    }

    // =====================================================================
    // Bug4602: PPTX table cell fill roundtrip
    // =====================================================================
    [Fact]
    public void Bug4602_Pptx_Table_Cell_Fill_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["fill"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = node.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("fill",
            because: "fill should be readable on table cell after Set");
        cellNode.Format["fill"].ToString().Should().Be("#FF0000",
            because: "fill=FF0000 should roundtrip");
    }

    // =====================================================================
    // Bug4603: PPTX table cell valign roundtrip
    // =====================================================================
    [Fact]
    public void Bug4603_Pptx_Table_Cell_VAlign_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["valign"] = "center"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = node.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("valign",
            because: "valign should be readable on table cell after Set");
        // Note: Set uses "center" but NodeBuilder maps to "middle" (line 185)
        // BUG if inconsistent naming between Set and Get
        var val = cellNode.Format["valign"].ToString();
        val.Should().BeOneOf(new[] { "center", "middle" },
            because: "valign should roundtrip with consistent naming");
    }

    // =====================================================================
    // Bug4604: PPTX connector lineDash Set then Get roundtrip
    // Connector Set handler now supports lineDash (lines 941-957).
    // Verify it roundtrips correctly.
    // =====================================================================
    [Fact]
    public void Bug4604_Pptx_Connector_LineDash_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["line"] = "0000FF"
        });

        handler.Set("/slide[1]/connector[1]", new()
        {
            ["lineDash"] = "dash"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Format.Should().ContainKey("lineDash",
            because: "lineDash should be readable after Set");
        node.Format["lineDash"].ToString().Should().Be("dash",
            because: "lineDash=dash should roundtrip correctly");
    }

    // =====================================================================
    // Bug4605: PPTX connector rotation Set then Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4605_Pptx_Connector_Rotation_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["line"] = "FF0000"
        });

        handler.Set("/slide[1]/connector[1]", new()
        {
            ["rotation"] = "45"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Format.Should().ContainKey("rotation",
            because: "rotation should be readable after Set");
        node.Format["rotation"].ToString().Should().Be("45",
            because: "rotation=45 should roundtrip correctly");
    }

    // =====================================================================
    // Bug4606: PPTX connector lineOpacity requires existing lineColor
    // If no lineColor is set, lineOpacity silently does nothing.
    // =====================================================================
    [Fact]
    public void Bug4606_Pptx_Connector_LineOpacity_Without_LineColor()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new());

        // Set lineOpacity without lineColor — should it work?
        handler.Set("/slide[1]/connector[1]", new()
        {
            ["lineOpacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        // BUG: lineOpacity silently ignored without existing SolidFill
        node.Format.Should().ContainKey("lineOpacity",
            because: "lineOpacity should work even without explicit lineColor");
    }

    // =====================================================================
    // Bug4607: Word paragraph Add with superscript and subscript
    // Both can be specified — subscript should win (last one)
    // =====================================================================
    [Fact]
    public void Bug4607_Word_Add_Superscript_And_Subscript()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Both superscript and subscript specified — which wins?
        handler.Add("/body", "p", null, new()
        {
            ["text"] = "H2O",
            ["superscript"] = "true",
            ["subscript"] = "true"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        // Both set — subscript comes after superscript in Add handler (line 187)
        // so subscript should overwrite superscript
        node.Format.Should().ContainKey("subscript",
            because: "subscript should overwrite superscript when both specified");
        node.Format.Should().NotContainKey("superscript",
            because: "superscript should be overwritten by subscript");
    }

    // =====================================================================
    // Bug4608: PPTX shape Add with "lineColor" key (not "line") for line
    // Add handler checks "line" OR "linecolor" OR "lineColor" (line 392)
    // =====================================================================
    [Fact]
    public void Bug4608_Pptx_Shape_Add_LineColor_Key()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Outlined",
            ["lineColor"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("line",
            because: "lineColor key should set the line property");
        node.Format["line"].ToString().Should().Be("#FF0000",
            because: "lineColor=FF0000 should be readable as line=FF0000");
    }

    // =====================================================================
    // Bug4609: PPTX shape Set "line" key — should work as alias for "lineColor"
    // Set handler at SetRunOrShapeProperties checks "linecolor" (line 320+)
    // but what about "line" key?
    // =====================================================================
    [Fact]
    public void Bug4609_Pptx_Shape_Set_Line_Key()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "NoLine"
        });

        // Set "line" to add a border color
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["line"] = "0000FF"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("line",
            because: "Set with 'line' key should add a border color");
        node.Format["line"].ToString().Should().Be("#0000FF",
            because: "line=0000FF should roundtrip");
    }

    // =====================================================================
    // Bug4610: PPTX slide background Set and Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4610_Pptx_Slide_Background_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new()
        {
            ["background"] = "FFFF00"
        });

        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background",
            because: "background should be readable after Set");
        node.Format["background"].ToString().Should().Be("#FFFF00",
            because: "background=FFFF00 should roundtrip");
    }

    // =====================================================================
    // Bug4611: PPTX slide notes text roundtrip
    // =====================================================================
    [Fact]
    public void Bug4611_Pptx_Slide_Notes_Text_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]/notes", new()
        {
            ["text"] = "Speaker notes here"
        });

        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("notes",
            because: "notes should be readable after Set via /slide[1]/notes path");
    }

    // =====================================================================
    // Bug4612: Word paragraph keepNext roundtrip
    // =====================================================================
    [Fact]
    public void Bug4612_Word_Paragraph_KeepNext_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "Keep", ["keepnext"] = "true"
        });

        var node = handler.Get("/body/p[1]");
        // Does NodeBuilder report keepNext?
        node.Format.Should().ContainKey("keepNext",
            because: "keepNext should be readable after Add");
    }

    // =====================================================================
    // Bug4613: Word paragraph pageBreakBefore roundtrip
    // =====================================================================
    [Fact]
    public void Bug4613_Word_Paragraph_PageBreakBefore_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "Break", ["pagebreakbefore"] = "true"
        });

        var node = handler.Get("/body/p[1]");
        // Does NodeBuilder report pageBreakBefore?
        node.Format.Should().ContainKey("pagebreakbefore",
            because: "pageBreakBefore should be readable after Add");
    }

    // =====================================================================
    // Bug4614: Word paragraph firstLineIndent roundtrip
    // =====================================================================
    [Fact]
    public void Bug4614_Word_Paragraph_FirstLineIndent_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "Indented", ["firstlineindent"] = "720"
        });

        var node = handler.Get("/body/p[1]");
        // Does NodeBuilder report firstLineIndent?
        node.Format.Should().ContainKey("firstLineIndent",
            because: "firstLineIndent should be readable after Add");
    }

    // =====================================================================
    // Bug4615: PPTX shape margin Set then Get — verify "margin" key exists
    // Set uses "margin" key in SetRunOrShapeProperties
    // =====================================================================
    [Fact]
    public void Bug4615_Pptx_Shape_Margin_Set_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Margins"
        });

        // Set all margins to 0
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["margin"] = "0"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("margin",
            because: "margin should be readable after Set");
        node.Format["margin"].ToString().Should().Be("0cm",
            because: "margin=0 should be '0cm' after FormatEmu");
    }

    // =====================================================================
    // Bug4616: PPTX table cell font Set then Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4616_Pptx_Table_Cell_Font_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Fonted", ["font"] = "Arial"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = node.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("font",
            because: "font should be readable on table cell after Set");
        cellNode.Format["font"].ToString().Should().Be("Arial",
            because: "font=Arial should roundtrip");
    }

    // =====================================================================
    // Bug4617: PPTX table cell color Set then Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4617_Pptx_Table_Cell_Color_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Red", ["color"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = node.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("color",
            because: "color should be readable on table cell after Set");
        cellNode.Format["color"].ToString().Should().Be("#FF0000",
            because: "color=FF0000 should roundtrip");
    }

    // =====================================================================
    // Bug4618: Word document properties roundtrip
    // Set title/author/subject then verify readable
    // =====================================================================
    [Fact]
    public void Bug4618_Word_Document_Properties_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Set("/", new()
        {
            ["title"] = "My Document",
            ["author"] = "Test Author"
        });

        var node = handler.Get("/");
        node.Format.Should().ContainKey("title",
            because: "title should be readable after Set");
        node.Format["title"].ToString().Should().Be("My Document",
            because: "title should roundtrip correctly");
    }

    // =====================================================================
    // Bug4619: Excel named range ref roundtrip
    // =====================================================================
    [Fact]
    public void Bug4619_Excel_Named_Range_Ref_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "100" });
        handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "Revenue",
            ["ref"] = "Sheet1!$A$1"
        });

        handler.Set("/namedrange[Revenue]", new() { ["ref"] = "Sheet1!$A$1:$A$10" });

        var node = handler.Get("/namedrange[Revenue]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("ref",
            because: "named range ref should be readable");
        node.Format["ref"].ToString().Should().Contain("$A$1:$A$10",
            because: "ref should be updated to the new range");
    }

    // =====================================================================
    // Bug4620: PPTX shape depth (3D extrusion) roundtrip
    // =====================================================================
    [Fact]
    public void Bug4620_Pptx_Shape_Depth_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D",
            ["fill"] = "0000FF",
            ["depth"] = "10"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("depth",
            because: "depth should be readable after Add");
    }

    // =====================================================================
    // Bug4621: PPTX shape bevel roundtrip
    // =====================================================================
    [Fact]
    public void Bug4621_Pptx_Shape_Bevel_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Beveled",
            ["fill"] = "00FF00",
            ["bevel"] = "circle"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("bevel",
            because: "bevel should be readable after Add");
    }

    // =====================================================================
    // Bug4622: PPTX shape material roundtrip
    // =====================================================================
    [Fact]
    public void Bug4622_Pptx_Shape_Material_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Matte",
            ["fill"] = "FF0000",
            ["material"] = "matte"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("material",
            because: "material should be readable after Add");
        node.Format["material"].ToString().Should().Be("matte",
            because: "material=matte should roundtrip");
    }

    // =====================================================================
    // Bug4623: Excel cell clear should reset everything including style
    // =====================================================================
    [Fact]
    public void Bug4623_Excel_Cell_Clear_Resets_Style()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Styled",
            ["font.bold"] = "true", ["fill"] = "FFFF00"
        });

        var n1 = handler.Get("/Sheet1/A1");
        n1.Format.Should().ContainKey("font.bold");

        handler.Set("/Sheet1/A1", new() { ["clear"] = "true" });

        var n2 = handler.Get("/Sheet1/A1");
        n2.Text.Should().BeNullOrEmpty();
        n2.Format.Should().NotContainKey("font.bold",
            because: "clear should reset all styling");
    }

    // =====================================================================
    // Bug4624: Word run highlight roundtrip
    // =====================================================================
    [Fact]
    public void Bug4624_Word_Run_Highlight_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "Highlighted", ["highlight"] = "yellow"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("highlight",
            because: "highlight should be readable on run after Add");
        node.Format["highlight"].ToString().Should().Be("yellow",
            because: "highlight=yellow should roundtrip");
    }

    // =====================================================================
    // Bug4625: Word run caps roundtrip
    // =====================================================================
    [Fact]
    public void Bug4625_Word_Run_Caps_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "caps text", ["caps"] = "true"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("caps",
            because: "caps should be readable after Add");
    }

    // =====================================================================
    // Bug4626: Excel cell formula roundtrip — verify formula is preserved
    // after setting value then formula
    // =====================================================================
    [Fact]
    public void Bug4626_Excel_Cell_Formula_Replaces_Value()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "100"
        });
        handler.Set("/Sheet1/A1", new() { ["formula"] = "SUM(B1:B10)" });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].ToString().Should().Be("SUM(B1:B10)");
        // After setting formula, the CellValue is cleared but Text shows formula
        // This is by design — GetCellDisplayValue returns formula text when no value
        node.Text.Should().NotBe("100",
            because: "old value '100' should be replaced by formula");
    }

    // =====================================================================
    // Bug4627: PPTX shape text color via Add and Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4627_Pptx_Shape_Text_Color_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Red text",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("color",
            because: "text color should be readable after Add");
        node.Format["color"].ToString().Should().Be("#FF0000",
            because: "color=FF0000 should roundtrip");
    }

    // =====================================================================
    // Bug4628: PPTX shape link roundtrip
    // =====================================================================
    [Fact]
    public void Bug4628_Pptx_Shape_Link_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Click me",
            ["link"] = "https://example.com"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("link",
            because: "link should be readable after Add");
        // URL may have trailing slash added by Uri normalization
        node.Format["link"].ToString().Should().StartWith("https://example.com",
            because: "link URL should roundtrip");
    }

    // =====================================================================
    // Bug4629: Word run double-strike roundtrip
    // =====================================================================
    [Fact]
    public void Bug4629_Word_Run_DoubleStrike_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "dstrike", ["dstrike"] = "true"
        });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("dstrike",
            because: "double-strike should be readable after Add");
    }
}
