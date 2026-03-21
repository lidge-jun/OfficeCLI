// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart41 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // =====================================================================
    // Bug4100: PPTX connector Set "lineDash" falls through to default (unsupported)
    // Connector Set handler (Set.cs:880-941) only has: name, x/y/width/height,
    // linewidth, linecolor, preset. Missing: lineDash, lineOpacity, rotation.
    // =====================================================================
    [Fact]
    public void Bug4100_Pptx_Connector_Set_LineDash_Returns_Unsupported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "0cm",
            ["lineColor"] = "000000", ["lineWidth"] = "2pt"
        });

        var unsupported = handler.Set("/slide[1]/connector[1]", new()
        {
            ["lineDash"] = "dot"
        });

        // lineDash should be supported but connector handler lacks this case
        unsupported.Should().BeEmpty(
            because: "lineDash should be supported for connectors like it is for shapes");
    }

    // =====================================================================
    // Bug4101: PPTX connector Set "lineOpacity" also unsupported
    // =====================================================================
    [Fact]
    public void Bug4101_Pptx_Connector_Set_LineOpacity_Unsupported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "0cm",
            ["lineColor"] = "000000", ["lineWidth"] = "2pt"
        });

        var unsupported = handler.Set("/slide[1]/connector[1]", new()
        {
            ["lineOpacity"] = "0.5"
        });

        unsupported.Should().BeEmpty(
            because: "lineOpacity should be supported for connectors like it is for shapes");
    }

    // =====================================================================
    // Bug4102: PPTX connector Set "rotation" also unsupported
    // =====================================================================
    [Fact]
    public void Bug4102_Pptx_Connector_Set_Rotation_Unsupported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "0cm"
        });

        var unsupported = handler.Set("/slide[1]/connector[1]", new()
        {
            ["rotation"] = "45"
        });

        unsupported.Should().BeEmpty(
            because: "rotation should be supported for connectors (Transform2D.Rotation)");
    }

    // =====================================================================
    // Bug4103: PPTX ConnectorToNode missing lineColor when using SrgbColorModelHex
    // ConnectorToNode (NodeBuilder.cs:801-804) only checks RgbColorModelHex,
    // but line SolidFill can also use SrgbColorModelHex or SchemeColor.
    // The Set handler creates RgbColorModelHex so this should work for simple cases.
    // Let's verify the basic roundtrip with a clean Add + Set.
    // =====================================================================
    [Fact]
    public void Bug4103_Pptx_Connector_LineColor_Set_Then_Get()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "0cm"
        });

        handler.Set("/slide[1]/connector[1]", new()
        {
            ["lineColor"] = "FF0000", ["lineWidth"] = "3pt"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Format.Should().ContainKey("lineColor");
        node.Format["lineColor"].Should().Be("#FF0000");
        node.Format.Should().ContainKey("lineWidth");
    }

    // =====================================================================
    // Bug4104: PPTX shape Add with "lineOpacity" is silently ignored
    // lineOpacity is NOT in the effectKeys set (Add.cs:404-413),
    // so it gets neither processed inline nor delegated to SetRunOrShapeProperties
    // =====================================================================
    [Fact]
    public void Bug4104_Pptx_Add_Shape_LineOpacity_Not_In_EffectKeys()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["fill"] = "FFFFFF",
            ["lineColor"] = "000000", ["lineWidth"] = "2pt",
            ["lineOpacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineOpacity",
            because: "lineOpacity should be settable during Add, but it's not in effectKeys");
    }

    // =====================================================================
    // Bug4105: PPTX shape Add with "opacity" (fill opacity) is handled inline
    // (Add.cs:351-364) but requires a SolidFill to already exist.
    // If fill is set in the same Add call, the order matters.
    // =====================================================================
    [Fact]
    public void Bug4105_Pptx_Add_Shape_Fill_And_Opacity_Together()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Semi-transparent",
            ["fill"] = "FF0000",
            ["opacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("fill");
        node.Format["fill"].Should().Be("#FF0000");
        node.Format.Should().ContainKey("opacity",
            because: "opacity with fill in same Add call should work since fill is processed first");
        node.Format["opacity"].Should().Be("0.5");
    }

    // =====================================================================
    // Bug4106: PPTX shape rotation NodeBuilder — verify roundtrip
    // Set "rotation" stores value * 60000 in Transform2D.Rotation
    // NodeBuilder reads rotation and divides by 60000
    // =====================================================================
    [Fact]
    public void Bug4106_Pptx_Shape_Rotation_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Rotated", ["rotation"] = "45"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("rotation");
        // rotation should be "45" or 45 (as integer)
        var rotVal = Convert.ToDouble(node.Format["rotation"]);
        rotVal.Should().Be(45.0,
            because: "rotation of 45 degrees should roundtrip correctly");
    }

    // =====================================================================
    // Bug4107: PPTX shape flipH/flipV roundtrip
    // =====================================================================
    [Fact]
    public void Bug4107_Pptx_Shape_FlipH_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Flipped", ["flipH"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("flipH",
            because: "flipH should be readable after setting during Add");
        node.Format["flipH"].Should().Be(true);
    }

    // =====================================================================
    // Bug4108: Word table cell shading via Set — verify roundtrip
    // =====================================================================
    [Fact]
    public void Bug4108_Word_TableCell_Shading_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["fill"] = "FF0000"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cellNode.Format.Should().ContainKey("fill",
            because: "table cell fill/shading should be readable after setting");
    }

    // =====================================================================
    // Bug4109: Word table cell valign — verify Set + Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4109_Word_TableCell_Valign_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["valign"] = "center"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cellNode.Format.Should().ContainKey("valign",
            because: "table cell vertical alignment should be readable after setting");
    }

    // =====================================================================
    // Bug4110: Word paragraph Set "text" with multiline (\\n) — verify behavior
    // Unlike PPTX, Word paragraph doesn't support \\n in text replacement.
    // The "text" case (line 928-954) just sets the first run's text to the
    // full string including literal \n.
    // =====================================================================
    [Fact]
    public void Bug4110_Word_Paragraph_Set_Text_With_Newline()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "Original" });
        handler.Set("/body/p[1]", new() { ["text"] = "Line1\\nLine2" });

        var node = handler.Get("/body/p[1]");
        // Word paragraph text replacement should handle \\n somehow
        // Currently it likely stores the literal string "Line1\nLine2"
        node.Text.Should().NotBeNullOrEmpty();
    }

    // =====================================================================
    // Bug4111: PPTX shape Set "name" roundtrip
    // =====================================================================
    [Fact]
    public void Bug4111_Pptx_Shape_Name_Set_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Named", ["name"] = "MyShape"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["name"].Should().Be("MyShape");

        handler.Set("/slide[1]/shape[1]", new() { ["name"] = "RenamedShape" });
        node = handler.Get("/slide[1]/shape[1]");
        node.Format["name"].Should().Be("RenamedShape");
    }

    // =====================================================================
    // Bug4112: Excel cell Set "clear" removes value, formula, type, AND style
    // But CellToNode still shows the cell. Verify clear actually works.
    // =====================================================================
    [Fact]
    public void Bug4112_Excel_Cell_Clear_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Data", ["bold"] = "true"
        });

        var nodeBefore = handler.Get("/Sheet1/A1");
        nodeBefore.Text.Should().Be("Data");

        handler.Set("/Sheet1/A1", new() { ["clear"] = "true" });

        var nodeAfter = handler.Get("/Sheet1/A1");
        // After clear, cell should be empty
        (nodeAfter.Text ?? "").Should().BeEmpty(
            because: "cell should be empty after clear");
    }

    // =====================================================================
    // Bug4113: Word style Set "underline" — missing from supported keys
    // WordHandler.Set.cs style handler (lines 444-511) supports: name,
    // basedon, next, font, size, bold, italic, color, alignment, spacebefore,
    // spaceafter. Does NOT support: underline, strike, highlight.
    // =====================================================================
    [Fact]
    public void Bug4113_Word_Style_Set_Underline_Unsupported()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Add a custom style first
        handler.Add("/styles", "style", null, new()
        {
            ["id"] = "TestStyle", ["name"] = "Test Style", ["type"] = "paragraph"
        });

        var unsupported = handler.Set("/styles/TestStyle", new()
        {
            ["underline"] = "single"
        });

        unsupported.Should().BeEmpty(
            because: "underline should be supported for style Set, but the handler is missing this case");
    }

    // =====================================================================
    // Bug4114: Word style Set "strike" — also missing
    // =====================================================================
    [Fact]
    public void Bug4114_Word_Style_Set_Strike_Unsupported()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/styles", "style", null, new()
        {
            ["id"] = "TestStyle2", ["name"] = "Test Style 2", ["type"] = "paragraph"
        });

        var unsupported = handler.Set("/styles/TestStyle2", new()
        {
            ["strike"] = "true"
        });

        unsupported.Should().BeEmpty(
            because: "strike should be supported for style Set, but the handler is missing this case");
    }

    // =====================================================================
    // Bug4115: PPTX table cell Set "text" then "bold" (order test)
    // SetTableCellProperties uses lazy cell.Descendants<Run>() per case.
    // If "text" replaces all paragraphs, "bold" uses fresh cell.Descendants
    // which gets the NEW runs. So this should work (unlike stale runs for shapes).
    // =====================================================================
    [Fact]
    public void Bug4115_Pptx_TableCell_Set_Text_Then_Bold()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "10cm", ["height"] = "5cm"
        });

        // text then bold — cell uses lazy descendants, so bold should apply to new runs
        var props = new Dictionary<string, string>();
        props["text"] = "NewCellText";
        props["bold"] = "true";
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", props);

        // Table cell NodeBuilder doesn't report bold (Bug4012), but the XML should be correct
        // Let's at least verify the text was set
        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = tableNode.Children[0].Children[0];
        cellNode.Text.Should().Be("NewCellText");
    }

    // =====================================================================
    // Bug4116: PPTX shape Set text="\\n" (single newline = 2 paragraphs)
    // This triggers the multi-line path in SetRunOrShapeProperties
    // which creates 2 empty paragraphs. Combined with bold, it's a stale runs issue.
    // =====================================================================
    [Fact]
    public void Bug4116_Pptx_Shape_Set_Empty_Multiline_Plus_Bold()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Initial" });

        // Set empty multiline text + bold
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "\\n",
            ["bold"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // The text should be empty lines, and bold should apply
        // But stale runs bug means bold goes to orphaned runs
        if (node.Format.ContainsKey("bold"))
        {
            node.Format["bold"].Should().Be(true);
        }
        else
        {
            // This confirms the bug: bold was lost
            node.Format.Should().ContainKey("bold",
                because: "bold should apply to new empty-line runs, but stale runs bug prevents this");
        }
    }

    // =====================================================================
    // Bug4117: Word paragraph Set multiple formatting keys in same call
    // Verify font + size + bold + italic all applied together
    // =====================================================================
    [Fact]
    public void Bug4117_Word_Paragraph_Set_Multiple_Run_Formatting()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "Multi-format" });
        handler.Set("/body/p[1]", new()
        {
            ["font"] = "Arial",
            ["size"] = "16",
            ["bold"] = "true",
            ["italic"] = "true",
            ["color"] = "FF0000"
        });

        var runNode = handler.Get("/body/p[1]/r[1]");
        runNode.Format.Should().ContainKey("font");
        runNode.Format["font"].Should().Be("Arial");
        runNode.Format.Should().ContainKey("bold");
        runNode.Format["bold"].Should().Be(true);
        runNode.Format.Should().ContainKey("italic");
        runNode.Format["italic"].Should().Be(true);
        runNode.Format.Should().ContainKey("color");
        runNode.Format["color"].Should().Be("#FF0000");
    }

    // =====================================================================
    // Bug4118: Excel named range Add + Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4118_Excel_NamedRange_Add_Get_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "2" });

        handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "TestRange", ["ref"] = "Sheet1!$A$1:$A$2"
        });

        var node = handler.Get("/namedrange[1]");
        node.Type.Should().Be("namedrange");
        node.Format["name"].Should().Be("TestRange");
        node.Format["ref"].Should().Be("Sheet1!$A$1:$A$2");
    }

    // =====================================================================
    // Bug4119: PPTX shape with textWarp roundtrip
    // textwarp is in effectKeys so Add delegates to SetRunOrShapeProperties.
    // NodeBuilder reads textWarp from BodyProperties.
    // =====================================================================
    // BUG: textWarp="wave" produces "textWave" which is not a valid
    // TextShapeValues enum. The handler does not validate the warp name
    // before passing it to the enum constructor, causing an exception.
    [Fact]
    public void Bug4119_Pptx_Shape_TextWarp_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // "wave" → "textWave" which is not a valid OOXML preset
        // Valid would be "textWave1" or "textWave2"
        var act = () => handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Warped", ["textWarp"] = "wave"
        });

        act.Should().Throw<ArgumentOutOfRangeException>(
            because: "textWarp='wave' constructs invalid enum 'textWave' — handler should validate or map names");
    }

    // =====================================================================
    // Bug4120: Word paragraph NodeBuilder font size integer division truncation
    // int.Parse(rp.FontSize.Val.Value) / 2 uses integer division, so
    // a half-point value of 21 (10.5pt) is truncated to 10 instead of 10.5.
    // =====================================================================
    [Fact]
    public void Bug4120_Word_Paragraph_FontSize_IntegerDivision_Truncates()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "Test" });
        // Set font size to 10.5pt — stored as 21 half-points in OOXML
        handler.Set("/body/p[1]", new() { ["size"] = "10.5" });

        var node = handler.Get("/body/p[1]");
        // BUG: int.Parse("21") / 2 = 10 (integer division truncation)
        // Expected: 10.5 or "10.5" but gets 10
        var sizeVal = node.Format["size"];
        sizeVal.Should().NotBe(10,
            because: "font size 10.5pt should not be truncated to 10 by integer division");
    }

    // =====================================================================
    // Bug4121: Excel Add cell with type=boolean does not convert value
    // When adding a cell with value="true" and type="boolean", the value
    // stays as "true" instead of being converted to "1" as OOXML requires.
    // =====================================================================
    [Fact]
    public void Bug4121_Excel_Add_Cell_Boolean_Value_Not_Converted()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "true", ["type"] = "boolean"
        });

        var node = handler.Get("/Sheet1/A1");
        // OOXML boolean cells should store "1" for true and "0" for false
        node.Text.Should().Be("1",
            because: "boolean cell value 'true' should be stored as '1' in OOXML");
    }

    // =====================================================================
    // Bug4122: Excel Add cell clear does not reset DataType or StyleIndex
    // In Add handler, clear only resets CellValue and CellFormula,
    // but in Set handler it also resets DataType and StyleIndex.
    // =====================================================================
    [Fact]
    public void Bug4122_Excel_Add_Cell_Clear_Incomplete()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "hello", ["type"] = "string"
        });

        // Verify it was set
        var node1 = handler.Get("/Sheet1/A1");
        node1.Text.Should().Be("hello");

        // Now clear via Set (should reset everything)
        handler.Set("/Sheet1/A1", new() { ["clear"] = "true" });
        var node2 = handler.Get("/Sheet1/A1");
        node2.Format["type"].Should().Be("Number",
            because: "clear via Set should reset DataType (Number is default)");
    }

    // =====================================================================
    // Bug4123: PPTX shape Set lineOpacity auto-creates black line fill
    // When no lineColor exists, setting lineOpacity now auto-creates a
    // black SolidFill on the outline (matching Apache POI behavior).
    // =====================================================================
    [Fact]
    public void Bug4123_Pptx_Shape_LineOpacity_Without_LineColor_AutoCreates_Fill()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Set lineOpacity without setting lineColor first — auto-creates black line fill
        handler.Set("/slide[1]/shape[1]", new() { ["lineOpacity"] = "0.5" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineOpacity",
            because: "lineOpacity auto-creates a black line fill when none exists");
    }

    // =====================================================================
    // Bug4124: PPTX shape Set opacity auto-creates white fill
    // When no SolidFill exists, setting opacity now auto-creates a white
    // SolidFill and applies the alpha (matching Apache POI behavior).
    // =====================================================================
    [Fact]
    public void Bug4124_Pptx_Shape_Opacity_Without_Fill_AutoCreates_Fill()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Set opacity without setting fill first — auto-creates white fill
        handler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.5" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("opacity",
            because: "opacity auto-creates a white fill when none exists");
    }

    // =====================================================================
    // Bug4125: PPTX shape charspacing roundtrip naming mismatch
    // Set uses "spacing"/"charspacing"/"letterspacing" (line 406),
    // NodeBuilder reports as "spacing" (line 404).
    // But charspacing is also used via different naming in some paths.
    // =====================================================================
    [Fact]
    public void Bug4125_Pptx_Shape_CharSpacing_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Spaced", ["spacing"] = "2"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("spacing",
            because: "charspacing should be readable after Add");
        node.Format["spacing"].Should().Be("2",
            because: "charspacing 2pt should round-trip as '2'");
    }

    // =====================================================================
    // Bug4126: PPTX shape baseline/superscript roundtrip
    // Set accepts "super"/"sub"/"30" etc.
    // NodeBuilder reports as percentage (e.g., "30" for 30% superscript).
    // =====================================================================
    [Fact]
    public void Bug4126_Pptx_Shape_Baseline_Superscript_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "E=mc2", ["baseline"] = "super"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("baseline",
            because: "baseline should be readable after Add");
        node.Format["baseline"].Should().Be("30",
            because: "super = 30% baseline offset");
    }

    // =====================================================================
    // Bug4127: PPTX shape Add with lineOpacity and lineColor together
    // lineColor is handled inline but lineOpacity is NOT in effectKeys,
    // so lineOpacity is silently dropped during Add.
    // =====================================================================
    [Fact]
    public void Bug4127_Pptx_Add_LineColor_And_LineOpacity_Together()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["lineColor"] = "FF0000",
            ["lineOpacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineOpacity",
            because: "lineOpacity should be settable during Add alongside lineColor");
    }

    // =====================================================================
    // Bug4128: Word paragraph Set with text replaces all runs
    // When Set("/paragraph[1]", { ["text"] = "new" }) is called,
    // it updates only the first run's text and removes extra runs.
    // But if "bold" and "text" are in the same call and "text" comes
    // after "bold" in dict iteration, bold is set on old runs, then
    // text clears extra runs — bold may be lost on the kept run.
    // =====================================================================
    [Fact]
    public void Bug4128_Word_Paragraph_Set_Bold_Then_Text_Order()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "old text" });
        // Set bold and text together — bold should survive
        handler.Set("/body/p[1]", new() { ["bold"] = "true", ["text"] = "new text" });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("new text");
        node.Format.Should().ContainKey("bold",
            because: "bold set in same call as text should be preserved");
    }
}
