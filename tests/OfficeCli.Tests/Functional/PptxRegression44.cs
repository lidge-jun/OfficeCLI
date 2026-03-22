using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class PptxRegression44 : IDisposable
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
    // Bug4400: PPTX connector lineWidth format inconsistency
    // ShapeToNode reports lineWidth as FormatEmu (e.g., "0.07cm"),
    // but ConnectorToNode reports lineWidth as "Xpt" format (e.g., "1pt").
    // This inconsistency means the same property has different formats
    // depending on whether it's on a shape or connector.
    // =====================================================================
    [Fact]
    public void Bug4400_Pptx_Connector_LineWidth_Format_Different_From_Shape()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["lineColor"] = "000000", ["lineWidth"] = "2pt"
        });
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["line"] = "000000", ["linewidth"] = "2pt"
        });

        var shapeNode = handler.Get("/slide[1]/shape[1]");
        var cxnNode = handler.Get("/slide[1]/connector[1]");

        // Both should have the same lineWidth property value format
        shapeNode.Format.Should().ContainKey("lineWidth");
        cxnNode.Format.Should().ContainKey("lineWidth");
        var shapeLineWidth = shapeNode.Format["lineWidth"].ToString();
        var cxnLineWidth = cxnNode.Format["lineWidth"].ToString();
        // Both now use FormatEmu consistently
        shapeLineWidth.Should().Be(cxnLineWidth,
            because: "lineWidth format should be consistent between shapes and connectors");
    }

    // =====================================================================
    // Bug4401: PPTX connector Add doesn't support "lineColor" key (only "line")
    // Verified in code: Add handler checks "line" but not "lineColor".
    // This is the same as Bug4300 but tests directly with the "line" key
    // to show it works, proving the key inconsistency.
    // =====================================================================
    [Fact]
    public void Bug4401_Pptx_Connector_Add_Line_Key_Works()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["line"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Format.Should().ContainKey("lineColor");
        node.Format["lineColor"].Should().Be("#FF0000",
            because: "'line' key during connector Add should set line color");
    }

    // =====================================================================
    // Bug4402: PPTX connector lineDash Set via Set handler
    // Connector Set handler doesn't support lineDash (Bug4100 confirmed).
    // Even after creating with dash, Set can't change it.
    // =====================================================================
    [Fact]
    public void Bug4402_Pptx_Connector_Set_LineDash_Via_Set()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["line"] = "000000"
        });

        var unsupported = handler.Set("/slide[1]/connector[1]", new()
        {
            ["lineDash"] = "dash"
        });

        unsupported.Should().BeEmpty(
            because: "lineDash should be supported for connectors via Set");
    }

    // =====================================================================
    // Bug4403: PPTX shape Set "lineColor" vs "line" key
    // ShapeToNode uses "line" key for line color (line 430), but
    // SetRunOrShapeProperties uses "linecolor" (line 315).
    // So the key returned by Get ("line") can't be fed back to Set directly.
    // =====================================================================
    [Fact]
    public void Bug4403_Pptx_Shape_Line_Key_Naming_Inconsistency()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["lineColor"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // NodeBuilder uses "line" as the key name
        node.Format.Should().ContainKey("line",
            because: "ShapeToNode reports line color as 'line' key");
        // But Set uses "linecolor" — verify both work
        handler.Set("/slide[1]/shape[1]", new() { ["line"] = "0000FF" });
        var node2 = handler.Get("/slide[1]/shape[1]");
        // The "line" key probably doesn't match any Set case
        // and falls through to GenericXmlQuery
        node2.Format["line"].Should().Be("#0000FF",
            because: "setting 'line' should update the line color");
    }

    // =====================================================================
    // Bug4404: PPTX shape Set "line" key not handled as alias for lineColor
    // SetRunOrShapeProperties has "linecolor" (line 315) but not "line".
    // Setting {"line": "FF0000"} goes to default case and is unsupported.
    // =====================================================================
    [Fact]
    public void Bug4404_Pptx_Shape_Set_Line_Key_Unsupported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["lineColor"] = "FF0000"
        });

        var unsupported = handler.Set("/slide[1]/shape[1]", new() { ["line"] = "0000FF" });
        unsupported.Should().BeEmpty(
            because: "'line' should be accepted as an alias for 'lineColor' in Set");
    }

    // =====================================================================
    // Bug4405: PPTX shape lineDash naming inconsistency between Set and Get
    // Set maps user-friendly names (e.g., "longdash") to OOXML enums.
    // Get/NodeBuilder returns a mix: ShapeToNode returns OOXML InnerText
    // lowercase (e.g., "lgdash"), while ConnectorToNode now maps to
    // user-friendly names. Inconsistency.
    // =====================================================================
    [Fact]
    public void Bug4405_Pptx_Shape_LineDash_Naming_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["lineDash"] = "longdash"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineDash");
        // Set sends "longdash" which maps to LargeDash OOXML enum
        // NodeBuilder reads InnerText "lgDash" and lowercases to "lgdash"
        // The value should round-trip as the same user-friendly name
        node.Format["lineDash"].Should().Be("longdash",
            because: "lineDash should round-trip with user-friendly names, not OOXML InnerText");
    }

    // =====================================================================
    // Bug4406: PPTX shape Add with "align" uses effectKeys
    // But the effectKeys set doesn't include "align" — let me verify
    // how it actually works. (Bug4308 passed, so it works somehow)
    // =====================================================================
    [Fact]
    public void Bug4406_Pptx_Shape_Add_Align_Center_Roundtrip()
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
        node.Format.Should().ContainKey("align");
        node.Format["align"].Should().Be("center");
    }

    // =====================================================================
    // Bug4407: Excel cell Set with font.bold — verify style property
    // =====================================================================
    [Fact]
    public void Bug4407_Excel_Cell_Set_FontBold_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Bold text", ["font.bold"] = "true"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("font.bold",
            because: "font.bold should be readable after Add");
        node.Format["font.bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug4408: Excel cell Set with font.color — verify style property
    // =====================================================================
    [Fact]
    public void Bug4408_Excel_Cell_Set_FontColor_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Red text", ["font.color"] = "FF0000"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("font.color",
            because: "font.color should be readable after Add");
        node.Format["font.color"].Should().Be("#FF0000");
    }

    // =====================================================================
    // Bug4409: Excel cell Set with fill color — verify style property
    // =====================================================================
    [Fact]
    public void Bug4409_Excel_Cell_Set_Fill_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Filled", ["fill"] = "FFFF00"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("fill",
            because: "fill should be readable after Add");
    }

    // =====================================================================
    // Bug4410: PPTX slide background solid color roundtrip
    // =====================================================================
    [Fact]
    public void Bug4410_Pptx_Slide_Background_SolidColor_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["background"] = "FF0000" });

        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background",
            because: "slide background should be readable after Set");
    }

    // =====================================================================
    // Bug4411: PPTX notes roundtrip
    // =====================================================================
    [Fact]
    public void Bug4411_Pptx_Notes_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["notes"] = "Speaker notes here" });

        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("notes",
            because: "notes should be readable after Set");
        node.Format["notes"].Should().Be("Speaker notes here");
    }

    // =====================================================================
    // Bug4412: Word Add paragraph with multiple formatting props
    // =====================================================================
    [Fact]
    public void Bug4412_Word_Add_Paragraph_Bold_Italic_Color()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Formatted",
            ["bold"] = "true",
            ["italic"] = "true",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Formatted");
        node.Format.Should().ContainKey("bold");
        node.Format.Should().ContainKey("italic");
        node.Format.Should().ContainKey("color");
    }

    // =====================================================================
    // Bug4413: Word paragraph Set text preserves formatting
    // When changing text on a paragraph that has bold/italic,
    // the formatting should be preserved on the remaining run.
    // =====================================================================
    [Fact]
    public void Bug4413_Word_Paragraph_Set_Text_Preserves_Bold()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Original", ["bold"] = "true"
        });

        handler.Set("/body/p[1]", new() { ["text"] = "Changed" });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Changed");
        node.Format.Should().ContainKey("bold",
            because: "bold should be preserved when only text is changed");
    }

    // =====================================================================
    // Bug4414: PPTX shape Set text preserves bold (stale runs fix)
    // After the stale runs fix (line 65-67), setting text should preserve
    // formatting because the first run's RunProperties are cloned.
    // =====================================================================
    [Fact]
    public void Bug4414_Pptx_Shape_Set_Text_Preserves_Bold()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original", ["bold"] = "true"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["text"] = "Changed" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Changed");
        node.Format.Should().ContainKey("bold",
            because: "bold should be preserved when only text is changed via Set");
    }

    // =====================================================================
    // Bug4415: PPTX shape noFill line roundtrip
    // Setting line="none" should create a NoFill on the outline.
    // =====================================================================
    [Fact]
    public void Bug4415_Pptx_Shape_Line_None_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "No border", ["lineColor"] = "none"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("line");
        node.Format["line"].Should().Be("none",
            because: "line='none' should round-trip as 'none'");
    }

    // =====================================================================
    // Bug4416: Excel cell border roundtrip
    // =====================================================================
    [Fact]
    public void Bug4416_Excel_Cell_Border_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Bordered",
            ["border.all"] = "thin"
        });

        var node = handler.Get("/Sheet1/A1");
        // border.all sets all 4 sides, readback should have border.left etc.
        node.Format.Should().ContainKey("border.left",
            because: "border should be readable after Add with border.all");
    }

    // =====================================================================
    // Bug4417: PPTX presentation slideSize roundtrip
    // =====================================================================
    [Fact]
    public void Bug4417_Pptx_Presentation_SlideSize_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Set("/", new() { ["slideSize"] = "16:9" });

        var node = handler.Get("/");
        node.Format.Should().ContainKey("slideWidth",
            because: "slide width should be readable after setting slideSize");
    }

    // =====================================================================
    // Bug4418: Word section pageWidth roundtrip
    // =====================================================================
    [Fact]
    public void Bug4418_Word_Section_PageWidth_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Set("/section[1]", new() { ["pageWidth"] = "15840" });

        var node = handler.Get("/section[1]");
        // BUG: Set uses "pageWidth" (camelCase) but Get returns "pagewidth" (lowercase)
        // This is a casing inconsistency in the section properties
        node.Format.Should().ContainKey("pageWidth",
            because: "section Set accepts 'pageWidth' so Get should return the same key casing");
    }

    // =====================================================================
    // Bug4419: PPTX shape Add with image fill — not tested
    // "imagefill" is in SetRunOrShapeProperties but requires a file path.
    // Skip actual image test but verify the error handling.
    // =====================================================================
    [Fact]
    public void Bug4419_Pptx_Shape_ImageFill_Without_File()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Setting imageFill to a non-existent file should throw
        var act = () => handler.Set("/slide[1]/shape[1]", new()
        {
            ["imageFill"] = "/nonexistent/image.png"
        });

        act.Should().Throw<Exception>(
            because: "imageFill with non-existent file should throw an error");
    }

    // =====================================================================
    // Bug4420: Excel sheet name with special characters
    // =====================================================================
    [Fact]
    public void Bug4420_Excel_Sheet_Name_Special_Characters()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet 1" });
        handler.Add("/Sheet 1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "test"
        });

        var node = handler.Get("/Sheet 1/A1");
        node.Should().NotBeNull();
        node.Text.Should().Be("test",
            because: "sheet names with spaces should work");
    }

    // =====================================================================
    // Bug4421: PPTX shape Remove roundtrip
    // After removing a shape, the shape count should decrease.
    // =====================================================================
    [Fact]
    public void Bug4421_Pptx_Shape_Remove()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Keep" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Remove" });

        // Verify 2 shapes exist
        var shapes = handler.Query("shape");
        shapes.Should().HaveCount(2);

        handler.Remove("/slide[1]/shape[2]");

        shapes = handler.Query("shape");
        shapes.Should().HaveCount(1);
        shapes[0].Text.Should().Be("Keep");
    }

    // =====================================================================
    // Bug4422: Word paragraph "size" format is double (12.0) but
    // run "size" format is string ("12pt"). Inconsistency in type.
    // Paragraph NodeBuilder line 300: node.Format["size"] = int.Parse(...) / 2.0
    // Run NodeBuilder line 328: node.Format["size"] = GetRunFontSize(run) which returns "12pt"
    // =====================================================================
    [Fact]
    public void Bug4422_Word_Paragraph_Size_Type_Differs_From_Run_Size_Type()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new() { ["text"] = "Hello", ["size"] = "14" });

        // Get paragraph-level size
        var paraNode = handler.Get("/body/p[1]");
        paraNode.Format.Should().ContainKey("size");
        var paraSize = paraNode.Format["size"];

        // Get run-level size
        var runNode = handler.Get("/body/p[1]/r[1]");
        runNode.Format.Should().ContainKey("size");
        var runSize = runNode.Format["size"];

        // BUG: paragraph stores double (14.0) but run stores string ("14pt")
        // They should be the same type for consistency
        paraSize.GetType().Should().Be(runSize.GetType(),
            because: "paragraph-level 'size' and run-level 'size' should be the same type");
    }

    // =====================================================================
    // Bug4423: Excel Set type=boolean doesn't convert "yes"/"no" to "1"/"0"
    // Line 667-672: only handles "true" and "false" but not "yes"/"no"
    // But line 636-639 (Set value with existing boolean type) handles "yes"/"no"
    // =====================================================================
    [Fact]
    public void Bug4423_Excel_Set_Type_Boolean_Does_Not_Convert_Yes()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "yes" });
        handler.Set("/Sheet1/A1", new() { ["type"] = "boolean" });

        var node = handler.Get("/Sheet1/A1");
        // When type is set to boolean, "yes" should be converted to "1"
        node.Text.Should().Be("1",
            because: "setting type=boolean should convert 'yes' to '1' like Set value does");
    }

    // =====================================================================
    // Bug4424: Word section Set uses "pagewidth" (lowercase) key
    // but the format dictionary also stores "pagewidth" (lowercase).
    // This is fine internally, but the page size is stored as uint,
    // not string. If user sets pageWidth="15840", the value stored
    // in format is 15840u (uint), not the string "15840".
    // So round-tripping through Set then Get loses the string type.
    // =====================================================================
    [Fact]
    public void Bug4424_Word_Section_PageSize_Type_Is_Uint_Not_String()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Set("/section[1]", new() { ["pageWidth"] = "12240" });

        var node = handler.Get("/section[1]");
        node.Format.Should().ContainKey("pageWidth");
        // The value is stored as uint, so comparing to string "12240" would fail
        var pwValue = node.Format["pageWidth"];
        pwValue.ToString().Should().Be("12240",
            because: "pageWidth should round-trip correctly");
        // BUG: The value is uint (12240u), not string ("12240").
        // This means node.Format["pageWidth"].Should().Be("12240") would fail
        // because uint 12240 != string "12240".
        pwValue.Should().BeOfType<string>(
            because: "format values should be strings for consistency with other properties");
    }

    // =====================================================================
    // Bug4425: Word paragraph Set "text" removes all runs except first.
    // This means formatting from second+ runs is lost entirely.
    // If user has "Hello" (bold) + " World" (italic) and sets text="NewText",
    // the italic run is simply deleted.
    // =====================================================================
    [Fact]
    public void Bug4425_Word_Paragraph_SetText_Preserves_Only_FirstRun_Format()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Create a paragraph with text and formatting
        handler.Add("/body", "p", null, new() { ["text"] = "Hello", ["bold"] = "true" });

        var node1 = handler.Get("/body/p[1]");
        node1.Format.Should().ContainKey("bold");

        // Now set text again — should preserve bold
        handler.Set("/body/p[1]", new() { ["text"] = "NewText" });

        var node2 = handler.Get("/body/p[1]");
        node2.Text.Should().Be("NewText");
        // BUG: If bold was on the run, setting text should preserve the
        // first run's formatting (which it does — text goes to first run).
        // But paragraph mark properties may not be preserved.
        // Actually this should pass — let's check bold is retained.
        node2.Format.Should().ContainKey("bold",
            because: "setting text on paragraph should preserve existing formatting");
    }

    // =====================================================================
    // Bug4426: PPTX slide notes Set then Get roundtrip
    // Check if notes can be added and retrieved
    // =====================================================================
    [Fact]
    public void Bug4426_Pptx_Slide_Notes_Add_And_Get()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Set("/slide[1]", new() { ["notes"] = "My speaker notes" });

        var node = handler.Get("/slide[1]");
        node.Format.Should().ContainKey("notes",
            because: "notes should be readable after being set");
        node.Format["notes"].ToString().Should().Contain("My speaker notes");
    }

    // =====================================================================
    // Bug4427: PPTX presentation slideSize roundtrip via Set then Get
    // =====================================================================
    [Fact]
    public void Bug4427_Pptx_Presentation_SlideSize_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Set("/", new() { ["slideSize"] = "4:3" });

        var node = handler.Get("/");
        node.Format.Should().ContainKey("slideWidth",
            because: "presentation should report slide dimensions");
        // 4:3 = 9144000 EMU = 25.4cm width. Verify the value changed from default widescreen.
        // Default widescreen 16:9 = 12192000 EMU = 33.87cm width. After 4:3 it should be 25.4cm.
        node.Format["slideWidth"].ToString().Should().Be("25.4cm",
            because: "4:3 aspect ratio should have 25.4cm width");
    }

    // =====================================================================
    // Bug4428: Excel sheet rename doesn't update formula references
    // If a formula references "Sheet1!A1" and the sheet is renamed to
    // "Data", the formula still says "Sheet1!A1" — stale reference.
    // =====================================================================
    [Fact]
    public void Bug4428_Excel_Sheet_Rename_Does_Not_Update_Formula_References()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "100" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["formula"] = "Sheet1!A1*2" });

        // Rename the sheet
        handler.Set("/Sheet1", new() { ["name"] = "Data" });

        // Check if formula was updated
        var node = handler.Get("/Data/B1");
        // BUG: formula still references "Sheet1!A1" after rename
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].ToString().Should().Contain("Data",
            because: "formula should be updated when sheet is renamed");
    }

    // =====================================================================
    // Bug4429: Word Run Set "superscript" then "subscript" — second
    // should overwrite first, not leave both.
    // =====================================================================
    [Fact]
    public void Bug4429_Word_Run_Superscript_Then_Subscript()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new() { ["text"] = "H2O" });
        handler.Set("/body/p[1]/r[1]", new() { ["superscript"] = "true" });

        var node1 = handler.Get("/body/p[1]/r[1]");
        node1.Format.Should().ContainKey("superscript");

        handler.Set("/body/p[1]/r[1]", new() { ["subscript"] = "true" });

        var node2 = handler.Get("/body/p[1]/r[1]");
        node2.Format.Should().ContainKey("subscript");
        // When setting subscript, superscript should be removed
        node2.Format.Should().NotContainKey("superscript",
            because: "setting subscript should clear superscript");
    }

    // =====================================================================
    // Bug4430: PPTX shape with "geometry" property Add + Get roundtrip
    // Verify geometry is properly stored and readable
    // =====================================================================
    [Fact]
    public void Bug4430_Pptx_Shape_Geometry_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Diamond",
            ["geometry"] = "diamond"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("geometry",
            because: "geometry should be readable after Add");
        node.Format["geometry"].ToString().Should().Be("diamond",
            because: "geometry should round-trip correctly");
    }

    // =====================================================================
    // Bug4431: Word run shading via Set then Get roundtrip
    // Set "shd" = "FF0000" on a run, verify it's readable as "shading"
    // =====================================================================
    [Fact]
    public void Bug4431_Word_Run_Shading_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new() { ["text"] = "Highlighted" });
        handler.Set("/body/p[1]/r[1]", new() { ["shd"] = "FF0000" });

        var node = handler.Get("/body/p[1]/r[1]");
        node.Format.Should().ContainKey("shading",
            because: "run shading should be readable after Set");
        node.Format["shading"].ToString().Should().Be("#FF0000",
            because: "shading color should round-trip correctly");
    }

    // =====================================================================
    // Bug4432: PPTX shape rotation Add roundtrip
    // Verify rotation set during Add is readable
    // =====================================================================
    [Fact]
    public void Bug4432_Pptx_Shape_Rotation_Add_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Rotated",
            ["rotation"] = "45"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("rotation",
            because: "rotation should be readable after Add");
        // rotation in OOXML is stored as degrees * 60000
        // The NodeBuilder should return something representing 45 degrees
        var rotVal = node.Format["rotation"].ToString();
        rotVal.Should().Be("45",
            because: "rotation should round-trip as the original degree value");
    }

    // =====================================================================
    // Bug4433: Excel validation Add then Get roundtrip
    // Add a list validation, verify it's readable
    // =====================================================================
    [Fact]
    public void Bug4433_Excel_Validation_Add_Get_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A1:A10",
            ["type"] = "list",
            ["formula1"] = "Yes,No,Maybe"
        });

        var node = handler.Get("/Sheet1/validation[1]");
        node.Should().NotBeNull("validation should be retrievable after Add");
        node.Format.Should().ContainKey("type",
            because: "validation type should be readable");
        node.Format["type"].ToString().Should().Be("list",
            because: "validation type should round-trip correctly");
    }
}
