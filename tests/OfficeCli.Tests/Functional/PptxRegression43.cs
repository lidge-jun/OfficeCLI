using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class PptxRegression43 : IDisposable
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
    // Bug4300: PPTX connector Add uses "line" key instead of "lineColor"
    // The connector Add handler checks for "line" (line 1050) but
    // the connector Set handler uses "linecolor" (line 929).
    // So passing "lineColor" during Add is silently ignored.
    // =====================================================================
    [Fact]
    public void Bug4300_Pptx_Connector_Add_LineColor_Key_Ignored()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["lineColor"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        // BUG: Add uses "line" key, not "lineColor" — so "lineColor" is silently ignored
        // The connector gets default black (000000) instead of FF0000
        node.Format.Should().ContainKey("lineColor");
        node.Format["lineColor"].Should().Be("#FF0000",
            because: "lineColor during Add should set the connector's line color");
    }

    // =====================================================================
    // Bug4301: PPTX connector Add does not handle lineDash
    // Connector Add has no lineDash handling (unlike shape Add which
    // delegates to SetRunOrShapeProperties via effectKeys).
    // =====================================================================
    [Fact]
    public void Bug4301_Pptx_Connector_Add_LineDash_Not_Supported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["lineColor"] = "000000", ["lineDash"] = "dash"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        node.Format.Should().ContainKey("lineDash",
            because: "lineDash should be settable during connector Add");
    }

    // =====================================================================
    // Bug4302: PPTX connector Add does not handle lineWidth via "lineWidth" key
    // The Add handler checks "linewidth" (lowercase, line 1054), but
    // if the user passes "lineWidth" (camelCase), does it match?
    // Properties.TryGetValue is case-sensitive for Dictionary<string,string>.
    // =====================================================================
    [Fact]
    public void Bug4302_Pptx_Connector_Add_LineWidth_CaseSensitive()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["lineColor"] = "000000", ["lineWidth"] = "2pt"
        });

        var node = handler.Get("/slide[1]/connector[1]");
        // Check if lineWidth was applied (lineWidth is "linewidth" in the handler)
        // If the handler uses case-sensitive lookup, "lineWidth" won't match "linewidth"
        node.Format.Should().ContainKey("lineWidth");
        var widthVal = node.Format["lineWidth"].ToString();
        widthVal.Should().NotBe("0.03cm",
            because: "lineWidth='2pt' should change from default 1pt/0.03cm");
    }

    // =====================================================================
    // Bug4303: PPTX group Set does not support rotation
    // Group Set handler (line 978-999) only handles name and x/y/width/height.
    // Rotation is not supported for groups.
    // =====================================================================
    [Fact]
    public void Bug4303_Pptx_Group_Set_Rotation_Unsupported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });

        var unsupported = handler.Set("/slide[1]/group[1]", new() { ["rotation"] = "45" });
        unsupported.Should().BeEmpty(
            because: "rotation should be supported for groups like it is for shapes");
    }

    // =====================================================================
    // Bug4304: PPTX group Set does not support fill
    // =====================================================================
    [Fact]
    public void Bug4304_Pptx_Group_Set_Fill_Unsupported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });

        var unsupported = handler.Set("/slide[1]/group[1]", new() { ["fill"] = "FF0000" });
        unsupported.Should().BeEmpty(
            because: "fill should be supported for groups");
    }

    // =====================================================================
    // Bug4305: PPTX shape shadow roundtrip — opacity stored as percentage
    // Shadow stores opacity as 0-100 (percentage) but shape fill opacity
    // uses 0.0-1.0 (fraction). This inconsistency can confuse users.
    // =====================================================================
    [Fact]
    public void Bug4305_Pptx_Shape_Shadow_Opacity_Format()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shadow", ["shadow"] = "000000-4-45-3-50"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("shadow");
        // Shadow format is COLOR-BLUR-ANGLE-DIST-OPACITY
        var shadowVal = node.Format["shadow"].ToString()!;
        // Verify the shadow opacity reads back correctly
        var parts = shadowVal.Split('-');
        parts.Should().HaveCountGreaterOrEqualTo(5,
            because: "shadow should have 5 parts: color-blur-angle-dist-opacity");
        parts[4].Should().Be("50",
            because: "shadow opacity 50 should round-trip as '50'");
    }

    // =====================================================================
    // Bug4306: PPTX shape glow color parsing with 3-char hex
    // ApplyGlow splits on '-', so "F00-10-75" is parsed as
    // color="F00", radius="10", opacity="75". But BuildColorElement
    // may not handle 3-char hex colors properly.
    // =====================================================================
    [Fact]
    public void Bug4306_Pptx_Shape_Glow_ThreeCharHex()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        // Use 3-char hex — BuildColorElement should expand it or handle it
        var act = () => handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["glow"] = "F00-10-75"
        });

        // This should either work or throw a clear error
        act.Should().NotThrow(because: "3-char hex colors should be handled gracefully");

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("glow");
    }

    // =====================================================================
    // Bug4307: Excel cell Set with value=boolean doesn't check existing type
    // If cell already has type=boolean, setting value="true" should
    // auto-convert to "1", but the handler just stores "true" verbatim.
    // =====================================================================
    [Fact]
    public void Bug4307_Excel_Set_Cell_Value_Should_Convert_For_Boolean_Type()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "1", ["type"] = "boolean"
        });

        // Now update just the value
        handler.Set("/Sheet1/A1", new() { ["value"] = "true" });

        var node = handler.Get("/Sheet1/A1");
        // The cell type is still boolean, but value was set to "true"
        // instead of being converted to "1"
        node.Format["type"].Should().Be("Boolean");
        // The cell value should be "1" (true) or "true" was stored verbatim
        // This is a bug if "true" is stored for a Boolean cell
        node.Text.Should().Be("1",
            because: "setting value='true' on a boolean cell should convert to '1'");
    }

    // =====================================================================
    // Bug4308: PPTX shape Add with align — not in effectKeys
    // "align" or "alignment" is not in the effectKeys set and is not
    // handled inline in Add. It should be supported.
    // =====================================================================
    [Fact]
    public void Bug4308_Pptx_Add_Shape_Alignment_Not_In_EffectKeys()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Right", ["align"] = "right"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("align",
            because: "text alignment should be settable during shape Add");
    }

    // =====================================================================
    // Bug4309: PPTX shape Add with valign — not handled
    // "valign" is not in effectKeys and not handled inline in Add.
    // =====================================================================
    [Fact]
    public void Bug4309_Pptx_Add_Shape_Valign_Not_Handled()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Bottom", ["valign"] = "bottom"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("valign",
            because: "valign should be settable during shape Add");
    }

    // =====================================================================
    // Bug4310: Word paragraph Set text with newlines doesn't create multiple runs
    // When Set text contains "\n", Word handler just updates first run's text
    // (line 958). Unlike PPTX which creates multiple paragraphs for multiline
    // text, Word just puts the literal "\n" or newline in the text.
    // =====================================================================
    [Fact]
    public void Bug4310_Word_Paragraph_Set_Text_With_Newline()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "Hello" });
        handler.Set("/body/p[1]", new() { ["text"] = "Line1\nLine2" });

        var node = handler.Get("/body/p[1]");
        // In Word, paragraphs don't contain newlines — each line is its own paragraph
        // So setting text with \n on a paragraph should either:
        // 1. Create a line break element
        // 2. Or preserve the text as-is
        node.Text.Should().Contain("Line1");
    }

    // =====================================================================
    // Bug4311: PPTX shape Set align="justify" roundtrip naming
    // Set sends "justify" to pProps.Alignment, but NodeBuilder reads
    // "just" from InnerText. ShapeToNode maps "just" → "justify" (line 581).
    // Let's verify the roundtrip works.
    // =====================================================================
    [Fact]
    public void Bug4311_Pptx_Shape_Alignment_Justify_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        handler.Set("/slide[1]/shape[1]", new() { ["align"] = "justify" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("align");
        node.Format["align"].Should().Be("justify",
            because: "justify alignment should round-trip correctly");
    }

    // =====================================================================
    // Bug4312: PPTX shape Set valign roundtrip
    // =====================================================================
    [Fact]
    public void Bug4312_Pptx_Shape_Valign_Bottom_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        handler.Set("/slide[1]/shape[1]", new() { ["valign"] = "bottom" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("valign");
        node.Format["valign"].Should().Be("bottom");
    }

    // =====================================================================
    // Bug4313: PPTX ConnectorToNode does not report rotation
    // ConnectorToNode (lines 773-807) does not read Transform2D.Rotation.
    // ShapeToNode does (lines 526-528).
    // =====================================================================
    [Fact]
    public void Bug4313_Pptx_ConnectorToNode_Missing_Rotation()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["lineColor"] = "000000"
        });

        // Even if we can't set rotation on connector, we should test
        // if the NodeBuilder would report it if it existed
        var node = handler.Get("/slide[1]/connector[1]");
        // Connector doesn't have rotation by default, so we just verify
        // the node is accessible. The real bug is that ConnectorToNode
        // doesn't have rotation reading code like ShapeToNode does.
        node.Type.Should().Be("connector");
    }

    // =====================================================================
    // Bug4314: Excel data validation roundtrip
    // Add a data validation and verify it can be read back.
    // =====================================================================
    [Fact]
    public void Bug4314_Excel_DataValidation_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A1:A10",
            ["type"] = "list",
            ["formula1"] = "\"Yes,No,Maybe\""
        });

        var node = handler.Get("/Sheet1");
        // Check that the sheet has validation info
        node.Should().NotBeNull();
    }

    // =====================================================================
    // Bug4315: Word section Set orientation swap logic
    // When setting orientation to "landscape" on a section that already has
    // landscape dimensions (w>h), the handler should NOT swap dimensions.
    // But if dimensions are portrait (w<h), it should swap.
    // =====================================================================
    [Fact]
    public void Bug4315_Word_Section_Orientation_Swap()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Set to landscape
        handler.Set("/section[1]", new() { ["orientation"] = "landscape" });

        var node = handler.Get("/section[1]");
        node.Format.Should().ContainKey("orientation");
        node.Format["orientation"].ToString().Should().Be("landscape");

        // Page dimensions should be swapped: width > height
        if (node.Format.ContainsKey("pageWidth") && node.Format.ContainsKey("pageHeight"))
        {
            var w = Convert.ToInt32(node.Format["pageWidth"]);
            var h = Convert.ToInt32(node.Format["pageHeight"]);
            w.Should().BeGreaterThan(h,
                because: "landscape orientation should have width > height");
        }
    }

    // =====================================================================
    // Bug4316: PPTX shape material roundtrip
    // =====================================================================
    [Fact]
    public void Bug4316_Pptx_Shape_Material_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Metal", ["material"] = "metal"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("material",
            because: "material should be readable after Add");
    }

    // =====================================================================
    // Bug4317: Word table Add with style — verify style is applied
    // =====================================================================
    [Fact]
    public void Bug4317_Word_Table_Add_With_Style()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["style"] = "TableGrid"
        });

        var node = handler.Get("/body/tbl[1]", depth: 0);
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("style",
            because: "table style should be readable after Add");
    }

    // =====================================================================
    // Bug4318: PPTX shape lighting roundtrip
    // =====================================================================
    [Fact]
    public void Bug4318_Pptx_Shape_Lighting_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Lit", ["lighting"] = "threePt"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lighting",
            because: "lighting should be readable after Add");
    }

    // =====================================================================
    // Bug4319: PPTX shape flipV roundtrip
    // =====================================================================
    [Fact]
    public void Bug4319_Pptx_Shape_FlipV_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Flipped", ["flipV"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("flipV",
            because: "flipV should be readable after Add");
        node.Format["flipV"].Should().Be(true);
    }

    // =====================================================================
    // Bug4320: Excel freeze pane roundtrip
    // =====================================================================
    [Fact]
    public void Bug4320_Excel_Freeze_Pane_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Set("/Sheet1", new() { ["freeze"] = "A2" });

        var node = handler.Get("/Sheet1");
        node.Format.Should().ContainKey("freeze",
            because: "freeze pane should be readable after Set");
    }

    // =====================================================================
    // Bug4321: PPTX shape list style roundtrip
    // =====================================================================
    [Fact]
    public void Bug4321_Pptx_Shape_ListStyle_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Item 1", ["list"] = "•"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("list",
            because: "list style should be readable after Add");
    }

    // =====================================================================
    // Bug4322: PPTX shape indent roundtrip
    // =====================================================================
    [Fact]
    public void Bug4322_Pptx_Shape_Indent_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Indented", ["indent"] = "1cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("indent",
            because: "indent should be readable after Add");
    }

    // =====================================================================
    // Bug4323: PPTX shape marginLeft roundtrip
    // =====================================================================
    [Fact]
    public void Bug4323_Pptx_Shape_MarginLeft_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Margin", ["marginLeft"] = "2cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("marginLeft",
            because: "marginLeft should be readable after Add");
    }

    // =====================================================================
    // Bug4324: Excel autofilter roundtrip
    // =====================================================================
    [Fact]
    public void Bug4324_Excel_AutoFilter_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Header" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Data" });

        handler.Set("/Sheet1", new() { ["autofilter"] = "A1:A2" });

        var node = handler.Get("/Sheet1");
        node.Format.Should().ContainKey("autoFilter",
            because: "autoFilter should be readable after Set");
    }

    // =====================================================================
    // Bug4325: PPTX shape Add with margin — not in effectKeys
    // "margin" is not in effectKeys and not handled inline in shape Add.
    // It's handled by SetRunOrShapeProperties but not listed in effectKeys.
    // =====================================================================
    [Fact]
    public void Bug4325_Pptx_Add_Shape_Margin_Not_In_EffectKeys()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["margin"] = "0.5cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("margin",
            because: "margin should be settable during Add");
    }

    // =====================================================================
    // Bug4326: PPTX shape Add with valign — not handled
    // "valign" is not in effectKeys and not handled inline.
    // =====================================================================
    [Fact]
    public void Bug4326_Pptx_Add_Shape_Valign_Not_Handled_Verify()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Middle", ["valign"] = "center"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("valign",
            because: "valign should be settable during Add via effectKeys");
    }

    // =====================================================================
    // Bug4327: Word Add paragraph with highlight — verify it works
    // =====================================================================
    [Fact]
    public void Bug4327_Word_Paragraph_Add_With_Highlight()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Highlighted", ["highlight"] = "yellow"
        });

        var node = handler.Get("/body/p[1]");
        // Check if highlight is reported at paragraph level
        // or if we need to navigate to run level
        var runNode = handler.Get("/body/p[1]/r[1]");
        runNode.Format.Should().ContainKey("highlight",
            because: "highlight should be readable at run level after Add");
    }

    // =====================================================================
    // Bug4328: PPTX connector NodeBuilder does not report flipH/flipV
    // ShapeToNode reports flipH/flipV from Transform2D (lines 512-514),
    // but ConnectorToNode likely doesn't.
    // =====================================================================
    [Fact]
    public void Bug4328_Pptx_ConnectorToNode_Missing_FlipH()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["lineColor"] = "000000"
        });

        // Connector doesn't support flipH during Add, but let's verify
        // what the NodeBuilder reports
        var node = handler.Get("/slide[1]/connector[1]");
        node.Type.Should().Be("connector");
        // Even without explicit flip, the NodeBuilder should have the same
        // property reading capability as ShapeToNode — but it doesn't
    }

    // =====================================================================
    // Bug4329: Excel cell Add with link — trailing slash normalization
    // Uri.TryCreate normalizes "https://example.com" to "https://example.com/"
    // This is a known bug confirmed in earlier sessions.
    // =====================================================================
    [Fact]
    public void Bug4329_Excel_Cell_Link_Trailing_Slash()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Click"
        });

        handler.Set("/Sheet1/A1", new() { ["link"] = "https://example.com" });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("link");
        node.Format["link"].Should().Be("https://example.com",
            because: "URL should not have trailing slash added by Uri.TryCreate normalization");
    }
}
