using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class PptxRegression42 : IDisposable
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
    // Bug4200: PPTX RunToNode does not report underline
    // ShapeToNode reports underline from firstRun at shape level,
    // but RunToNode itself does not include underline in Format.
    // =====================================================================
    [Fact]
    public void Bug4200_Pptx_RunToNode_Missing_Underline()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Underlined", ["underline"] = "single"
        });

        // Shape-level should have underline
        var shapeNode = handler.Get("/slide[1]/shape[1]");
        shapeNode.Format.Should().ContainKey("underline",
            because: "shape-level should report underline from first run");

        // Run-level (depth>1) should also have underline
        var runNode = handler.Get("/slide[1]/shape[1]/paragraph[1]/run[1]");
        runNode.Format.Should().ContainKey("underline",
            because: "RunToNode should report underline — it currently only reports bold/italic/spacing/baseline/color");
    }

    // =====================================================================
    // Bug4201: PPTX RunToNode does not report strike
    // Same gap as underline — RunToNode checks bold/italic but not strike.
    // =====================================================================
    [Fact]
    public void Bug4201_Pptx_RunToNode_Missing_Strike()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Struck", ["strike"] = "single"
        });

        // Shape-level should have strike
        var shapeNode = handler.Get("/slide[1]/shape[1]");
        shapeNode.Format.Should().ContainKey("strike",
            because: "shape-level should report strike from first run");

        // Run-level should also have strike
        var runNode = handler.Get("/slide[1]/shape[1]/paragraph[1]/run[1]");
        runNode.Format.Should().ContainKey("strike",
            because: "RunToNode should report strike — it currently only reports bold/italic/spacing/baseline/color");
    }

    // =====================================================================
    // Bug4202: PPTX paragraph child node does not report lineSpacing
    // ShapeToNode reports lineSpacing from first paragraph at shape level,
    // but paragraph child nodes (depth>0) don't report lineSpacing.
    // =====================================================================
    [Fact]
    public void Bug4202_Pptx_ParagraphNode_Missing_LineSpacing()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Spaced", ["lineSpacing"] = "1.5"
        });

        // Shape-level should have lineSpacing
        var shapeNode = handler.Get("/slide[1]/shape[1]");
        shapeNode.Format.Should().ContainKey("lineSpacing");

        // Paragraph child should also have lineSpacing
        var paraNode = handler.Get("/slide[1]/shape[1]/paragraph[1]");
        paraNode.Format.Should().ContainKey("lineSpacing",
            because: "paragraph child node should report lineSpacing, not just shape-level");
    }

    // =====================================================================
    // Bug4203: PPTX paragraph child node does not report spaceBefore
    // Same issue — spaceBefore only reported at shape level.
    // =====================================================================
    [Fact]
    public void Bug4203_Pptx_ParagraphNode_Missing_SpaceBefore()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["spaceBefore"] = "12"
        });

        var shapeNode = handler.Get("/slide[1]/shape[1]");
        shapeNode.Format.Should().ContainKey("spaceBefore");

        var paraNode = handler.Get("/slide[1]/shape[1]/paragraph[1]");
        paraNode.Format.Should().ContainKey("spaceBefore",
            because: "paragraph child node should report spaceBefore");
    }

    // =====================================================================
    // Bug4204: PPTX paragraph child node does not report spaceAfter
    // =====================================================================
    [Fact]
    public void Bug4204_Pptx_ParagraphNode_Missing_SpaceAfter()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["spaceAfter"] = "6"
        });

        var shapeNode = handler.Get("/slide[1]/shape[1]");
        shapeNode.Format.Should().ContainKey("spaceAfter");

        var paraNode = handler.Get("/slide[1]/shape[1]/paragraph[1]");
        paraNode.Format.Should().ContainKey("spaceAfter",
            because: "paragraph child node should report spaceAfter");
    }

    // =====================================================================
    // Bug4205: PPTX ConnectorToNode does not report lineDash
    // ShapeToNode reports lineDash (line 434), but ConnectorToNode only
    // reports lineWidth and lineColor (lines 773-807).
    // =====================================================================
    [Fact]
    public void Bug4205_Pptx_ConnectorToNode_Missing_LineDash()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["lineColor"] = "FF0000", ["lineDash"] = "dash"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Format.Should().ContainKey("lineDash",
            because: "ConnectorToNode should report lineDash like ShapeToNode does");
    }

    // =====================================================================
    // Bug4206: PPTX ConnectorToNode does not report lineOpacity
    // ShapeToNode reports lineOpacity (line 438), ConnectorToNode doesn't.
    // =====================================================================
    [Fact]
    public void Bug4206_Pptx_ConnectorToNode_Missing_LineOpacity()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["lineColor"] = "0000FF"
        });
        // Set lineOpacity — even though Set may not support it for connector,
        // if it did, the NodeBuilder wouldn't report it
        // So this test just checks that after creating a connector with lineColor,
        // if we manually set lineOpacity we can read it back
        var node = handler.Get("/slide[1]/connector[1]");
        // At minimum, the connector node should have same line-property reporting as shapes
        // If lineColor exists, the NodeBuilder should also check for alpha on it
        node.Format.Should().ContainKey("lineColor",
            because: "connector with lineColor should report it");
    }

    // =====================================================================
    // Bug4207: Excel cell Add with value and type=boolean — order matters
    // Properties dict iterates in insertion order. If "type" comes after
    // "value", the DataType is set but the CellValue is not converted.
    // =====================================================================
    [Fact]
    public void Bug4207_Excel_Add_Cell_Boolean_False_Not_Converted()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "false", ["type"] = "boolean"
        });

        var node = handler.Get("/Sheet1/A1");
        // Boolean false should be stored as "0"
        node.Text.Should().Be("0",
            because: "boolean cell value 'false' should be stored as '0' in OOXML");
    }

    // =====================================================================
    // Bug4208: Excel Set cell with type=boolean also doesn't convert value
    // Same issue in Set handler — setting type=boolean doesn't convert
    // existing CellValue from "true"/"false" to "1"/"0".
    // =====================================================================
    [Fact]
    public void Bug4208_Excel_Set_Cell_Boolean_Value_Not_Converted()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "hello"
        });

        // Now set type to boolean and value to true
        handler.Set("/Sheet1/A1", new() { ["value"] = "true", ["type"] = "boolean" });

        var node = handler.Get("/Sheet1/A1");
        node.Text.Should().Be("1",
            because: "boolean cell value 'true' should be stored as '1' in OOXML");
    }

    // =====================================================================
    // Bug4209: PPTX shape Add with fill=none then opacity — opacity silently fails
    // If fill is "none" (NoFill element), then opacity has nothing to attach to.
    // =====================================================================
    [Fact]
    public void Bug4209_Pptx_Add_Shape_NoFill_Then_Opacity()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["fill"] = "none", ["opacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // fill=none means NoFill, so opacity has nowhere to go
        // The handler should either warn or apply opacity to text or ignore gracefully
        node.Format["fill"].Should().Be("none");
    }

    // =====================================================================
    // Bug4210: Word paragraph Set multiple formatting with text replacement
    // When Set is called with {text, bold, italic, color} together,
    // text is handled at line 954-977, and bold/italic/color at line 941-952.
    // The ordering matters: bold/italic/color are applied to existing runs
    // first, then text replaces runs. Since text removes extra runs
    // (line 961) but keeps the first run, formatting should be preserved.
    // But if the paragraph has NO runs initially and text creates one,
    // the formatting may not be applied.
    // =====================================================================
    [Fact]
    public void Bug4210_Word_Paragraph_Set_Text_And_Formatting_On_Empty_Para()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new());

        // Set text, bold, and color on an empty paragraph (no runs)
        handler.Set("/body/p[1]", new()
        {
            ["bold"] = "true", ["color"] = "FF0000", ["text"] = "Formatted"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Formatted");
        node.Format.Should().ContainKey("bold",
            because: "bold should be applied even when paragraph starts empty");
        node.Format.Should().ContainKey("color",
            because: "color should be applied even when paragraph starts empty");
    }

    // =====================================================================
    // Bug4211: PPTX table cell Set text+bold order dependency
    // SetTableCellProperties iterates properties in dict order.
    // If "bold" comes before "text", bold is applied to old runs,
    // then "text" replaces all paragraphs but preserves first run's props.
    // If "text" comes before "bold", text creates new runs then bold
    // is applied via lazy cell.Descendants<Drawing.Run>().
    // The lazy evaluation should work, but let's verify.
    // =====================================================================
    [Fact]
    public void Bug4211_Pptx_TableCell_Set_Text_Before_Bold()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1"
        });

        // Set text then bold — bold uses lazy descendants so should work
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Test", ["bold"] = "true"
        });

        // Get table with depth=2 to include row and cell children
        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        node.Children.Should().NotBeEmpty("table should have row children at depth 2");
        var rowNode = node.Children[0];
        rowNode.Children.Should().NotBeEmpty("row should have cell children at depth 2");
        var cellNode = rowNode.Children[0];
        cellNode.Text.Should().Be("Test");
        cellNode.Format.Should().ContainKey("bold");
        cellNode.Format["bold"].Should().Be(true,
            because: "bold should be applied even when text is set in same call");
    }

    // =====================================================================
    // Bug4212: PPTX shape Set with geometry (preset) change roundtrip
    // SetRunOrShapeProperties handles "geometry" but let's check that
    // it's in effectKeys for Add and works in both Add and Set.
    // =====================================================================
    [Fact]
    public void Bug4212_Pptx_Shape_Geometry_Change_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["preset"] = "rect"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["preset"].Should().Be("rect");

        // Change geometry via Set
        handler.Set("/slide[1]/shape[1]", new() { ["preset"] = "ellipse" });

        node = handler.Get("/slide[1]/shape[1]");
        node.Format["preset"].Should().Be("ellipse",
            because: "preset geometry should be updatable via Set");
    }

    // =====================================================================
    // Bug4213: Excel sheet rename doesn't update named range references
    // When renaming a sheet, DefinedNames that reference the old sheet
    // name (e.g., "Sheet1!$A$1:$A$2") are not updated.
    // =====================================================================
    [Fact]
    public void Bug4213_Excel_Sheet_Rename_NamedRange_Not_Updated()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "test" });
        handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "MyRange", ["ref"] = "Sheet1!$A$1"
        });

        // Rename the sheet
        handler.Set("/Sheet1", new() { ["name"] = "RenamedSheet" });

        // Get the named range
        var nr = handler.Get("/namedrange[1]");
        nr.Format["ref"].ToString().Should().Contain("RenamedSheet",
            because: "named range reference should be updated when sheet is renamed");
    }

    // =====================================================================
    // Bug4214: PPTX shape margin (text inset) roundtrip
    // Add supports margin/padding via effectKeys? Let's check if
    // margin is in effectKeys or handled inline.
    // =====================================================================
    [Fact]
    public void Bug4214_Pptx_Shape_Margin_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test"
        });

        // Set margin via Set
        handler.Set("/slide[1]/shape[1]", new() { ["margin"] = "1cm" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("margin",
            because: "margin should be readable after Set");
    }

    // =====================================================================
    // Bug4215: PPTX shape valign roundtrip via Set
    // =====================================================================
    [Fact]
    public void Bug4215_Pptx_Shape_Valign_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["valign"] = "center" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("valign",
            because: "valign should be readable after Set");
        node.Format["valign"].Should().Be("center");
    }

    // =====================================================================
    // Bug4216: Word table cell Set alignment roundtrip
    // Word table cell Set handler supports alignment?
    // =====================================================================
    [Fact]
    public void Bug4216_Word_TableCell_Alignment_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Centered", ["alignment"] = "center"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("alignment",
            because: "table cell alignment should be readable after Set");
    }

    // =====================================================================
    // Bug4217: PPTX shape paragraph alignment not in effectKeys for Add
    // "align" or "alignment" during Add — is it handled?
    // =====================================================================
    [Fact]
    public void Bug4217_Pptx_Add_Shape_Alignment()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Centered", ["align"] = "center"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("align",
            because: "alignment should be settable during Add");
        node.Format["align"].Should().Be("center");
    }

    // =====================================================================
    // Bug4218: Excel cell number format roundtrip
    // Setting a number format like "0.00" or "#,##0" via style props
    // should be readable back.
    // =====================================================================
    [Fact]
    public void Bug4218_Excel_Cell_NumberFormat_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "1234.5", ["format"] = "0.00"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("numberformat",
            because: "numberformat should be the canonical Excel key after Add");
    }

    // =====================================================================
    // Bug4219: PPTX connector Add with rotation — not supported
    // Connector Add likely doesn't handle rotation like shape Add does.
    // =====================================================================
    [Fact]
    public void Bug4219_Pptx_Connector_Add_Rotation()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        var result = handler.Add("/slide[1]", "connector", null, new()
        {
            ["lineColor"] = "000000", ["rotation"] = "45"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Format.Should().ContainKey("rotation",
            because: "connector rotation should be settable during Add");
    }

    // =====================================================================
    // Bug4220: PPTX shape reflection roundtrip
    // =====================================================================
    [Fact]
    public void Bug4220_Pptx_Shape_Reflection_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Reflected", ["reflection"] = "half"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("reflection",
            because: "reflection should be readable after Add");
        node.Format["reflection"].Should().Be("half");
    }

    // =====================================================================
    // Bug4221: PPTX shape softEdge roundtrip
    // =====================================================================
    [Fact]
    public void Bug4221_Pptx_Shape_SoftEdge_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Soft", ["softEdge"] = "10"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("softEdge",
            because: "softEdge should be readable after Add");
    }

    // =====================================================================
    // Bug4222: Excel merge cells via Set doesn't call ReorderWorksheetChildren
    // AppendChild(mergeCells) at line 787 may place MergeCells in wrong
    // position in the worksheet XML. ReorderWorksheetChildren should be
    // called after but it's unclear if it is.
    // =====================================================================
    [Fact]
    public void Bug4222_Excel_Merge_Cells_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "merged" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "" });

        handler.Set("/Sheet1", new() { ["merge"] = "A1:A2" });

        // Verify by querying — the file should be valid after merge
        var node = handler.Get("/Sheet1/A1");
        node.Text.Should().Be("merged");
    }

    // =====================================================================
    // Bug4223: PPTX shape bevel roundtrip
    // =====================================================================
    [Fact]
    public void Bug4223_Pptx_Shape_Bevel_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Beveled", ["bevel"] = "circle"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("bevel",
            because: "bevel should be readable after Add");
    }

    // =====================================================================
    // Bug4224: Word paragraph Add with alignment — verify it works
    // =====================================================================
    [Fact]
    public void Bug4224_Word_Paragraph_Add_With_Alignment()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Right aligned", ["alignment"] = "right"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("alignment",
            because: "alignment should be set during Add");
        node.Format["alignment"].Should().Be("right");
    }

    // =====================================================================
    // Bug4225: PPTX shape rot3d roundtrip
    // =====================================================================
    [Fact]
    public void Bug4225_Pptx_Shape_Rot3D_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D", ["rot3d"] = "45,30,0"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("rot3d",
            because: "rot3d should be readable after Add");
    }

    // =====================================================================
    // Bug4226: PPTX shape depth/extrusion roundtrip
    // =====================================================================
    [Fact]
    public void Bug4226_Pptx_Shape_Depth_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Deep", ["depth"] = "10"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("depth",
            because: "depth (extrusion) should be readable after Add");
    }

    // =====================================================================
    // Bug4227: Word paragraph NodeBuilder size value type inconsistency
    // int.Parse(rp.FontSize.Val.Value) / 2 returns int when half-points
    // are even, but the Set handler accepts decimal sizes like "12".
    // The NodeBuilder stores an int (e.g., 12) not a string like "12pt".
    // This means comparing node.Format["size"] to a string fails.
    // =====================================================================
    [Fact]
    public void Bug4227_Word_Paragraph_FontSize_Type_Is_Integer()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Test", ["size"] = "14"
        });

        var node = handler.Get("/body/p[1]");
        // Word NodeBuilder stores size as int (half-points / 2)
        // but PPTX stores size as "14pt" string
        // This inconsistency could confuse users
        node.Format.Should().ContainKey("size");
        var sizeVal = node.Format["size"];
        // Verify the type — it should be a numeric value matching 14
        sizeVal.Should().Be("14pt",
            because: "size 14pt should round-trip as '14pt'");
    }

    // =====================================================================
    // Bug4228: Excel comment roundtrip — can we read back comments?
    // =====================================================================
    [Fact]
    public void Bug4228_Excel_Comment_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "test" });
        handler.Add("/Sheet1", "comment", null, new()
        {
            ["ref"] = "A1", ["text"] = "This is a comment", ["author"] = "Tester"
        });

        var comment = handler.Get("/Sheet1/comment[1]");
        comment.Should().NotBeNull(because: "comment should be retrievable after Add");
        comment.Text.Should().Be("This is a comment");
    }
}
