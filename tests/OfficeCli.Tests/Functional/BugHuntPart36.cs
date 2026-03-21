// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart36 : IDisposable
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
    // Bug3600: PPTX table cell Get — font/size/bold/italic not reported
    // The NodeBuilder for table cells reads text, fill, borders, but does
    // NOT read font, size, bold, italic from the cell runs.
    // =====================================================================
    [Fact]
    public void Bug3600_Pptx_TableCell_Font_Not_Reported_In_Get()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["font"] = "Calibri"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = node.Children[0].Children[0];
        cell.Format.Should().ContainKey("font",
            "Table cell node should report font property from runs");
    }

    // =====================================================================
    // Bug3601: PPTX table cell Get — bold not reported
    // =====================================================================
    [Fact]
    public void Bug3601_Pptx_TableCell_Bold_Not_Reported_In_Get()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold cell",
            ["bold"] = "true"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = node.Children[0].Children[0];
        cell.Format.Should().ContainKey("bold",
            "Table cell node should report bold property");
    }

    // =====================================================================
    // Bug3602: PPTX table cell Get — size not reported
    // =====================================================================
    [Fact]
    public void Bug3602_Pptx_TableCell_Size_Not_Reported_In_Get()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["size"] = "24"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = node.Children[0].Children[0];
        cell.Format.Should().ContainKey("size",
            "Table cell node should report font size property");
    }

    // =====================================================================
    // Bug3603: PPTX table cell Get — color not reported
    // =====================================================================
    [Fact]
    public void Bug3603_Pptx_TableCell_Color_Not_Reported_In_Get()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Red text",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = node.Children[0].Children[0];
        cell.Format.Should().ContainKey("color",
            "Table cell node should report text color");
    }

    // =====================================================================
    // Bug3604: Word table cell Set text+bold together — deferred text loses bold
    // In WordHandler.Set for TableCell, "text" is deferred (line 970-972).
    // Bold (line 982-1018) is applied to existing runs first.
    // Then deferred text replaces all paragraphs/runs, creating new ones
    // that don't have the bold formatting.
    // =====================================================================
    [Fact]
    public void Bug3604_Word_TableCell_Set_TextAndBold_DeferredText_LosesBold()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1"
        });

        // Set text and bold together in one call
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold text",
            ["bold"] = "true"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Bold text");
        node.Format.Should().ContainKey("bold",
            "Text set via deferred text should inherit the bold formatting applied in the same Set call");
        node.Format["bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug3605: Word table cell Set text+font together — deferred text loses font
    // Same deferred text issue as Bug3604 but for font property.
    // =====================================================================
    [Fact]
    public void Bug3605_Word_TableCell_Set_TextAndFont_DeferredText_LosesFont()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Styled text",
            ["font"] = "Courier New"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Styled text");
        node.Format.Should().ContainKey("font",
            "Text set via deferred text should inherit the font from the same Set call");
        node.Format["font"].ToString().Should().Contain("Courier New");
    }

    // =====================================================================
    // Bug3606: Word table cell Set text+color — deferred text loses color
    // =====================================================================
    [Fact]
    public void Bug3606_Word_TableCell_Set_TextAndColor_DeferredText_LosesColor()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "table", null, new()
        {
            ["rows"] = "1", ["cols"] = "1"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Red text",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Red text");
        node.Format.Should().ContainKey("color",
            "Text set via deferred text should inherit the color from the same Set call");
        node.Format["color"].ToString()!.Should().Contain("#FF0000");
    }

    // =====================================================================
    // Bug3607: PPTX Add shape with autofit property — processed twice
    // The Add handler processes autofit both inline (lines 298-310) and
    // again via SetRunOrShapeProperties delegation (effectKeys, line 409).
    // This may cause duplicate autofit elements in the body properties.
    // =====================================================================
    [Fact]
    public void Bug3607_Pptx_Add_Shape_Autofit_Processed_Twice()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Autofit test",
            ["autofit"] = "true"
        });

        // Get the shape and verify only one autofit element exists
        var node = handler.Get("/slide[1]/shape[1]");
        // If autofit is processed twice, there could be duplicate NormalAutoFit elements
        // in BodyProperties. We can't directly inspect XML via Get, but we can verify
        // the shape is valid and doesn't crash on reopening.
        node.Should().NotBeNull();
        handler.Dispose();

        // Reopen to verify file integrity
        using var handler2 = new PowerPointHandler(path, editable: false);
        var node2 = handler2.Get("/slide[1]/shape[1]");
        node2.Should().NotBeNull();
    }

    // =====================================================================
    // Bug3608: PPTX shape Set text with multiple lines + opacity
    // Opacity requires a SolidFill on the shape. Setting multiline text
    // replaces paragraphs but doesn't affect shape fill. However, if there's
    // no fill, opacity is silently ignored — not reported as unsupported.
    // =====================================================================
    [Fact]
    public void Bug3608_Pptx_Set_MultilineText_And_Opacity_WithFill()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial",
            ["fill"] = "0000FF"
        });

        // Set multiline text AND opacity together
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line 1\\nLine 2",
            ["opacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line 1");
        node.Format.Should().ContainKey("opacity",
            "Opacity should be applied to shape fill");
    }

    // =====================================================================
    // Bug3609: PPTX shape Set text multiline + fill together
    // When both text (multiline) and fill are set together, fill is processed
    // after text. The stale runs issue only applies to run-level properties.
    // Fill should work correctly since it modifies ShapeProperties, not runs.
    // But verify the shape remains valid after both operations.
    // =====================================================================
    [Fact]
    public void Bug3609_Pptx_Set_MultilineText_And_Fill_Together()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Initial" });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line A\\nLine B",
            ["fill"] = "00FF00"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line A");
        node.Format.Should().ContainKey("fill");
        node.Format["fill"].ToString().Should().Be("#00FF00");
    }

    // =====================================================================
    // Bug3610: PPTX lineDash — NodeBuilder outputs "lineDash" but some
    // values are stored with OOXML internal names. The Set handler uses
    // "longdash" but the NodeBuilder reads InnerText which could be "lgDash".
    // Round-trip test: set lineDash=longdash, read back, verify readable.
    // =====================================================================
    [Fact]
    public void Bug3610_Pptx_LineDash_LongDash_Roundtrip_KeyName()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Dash test",
            ["line"] = "000000"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["linedash"] = "longdash"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // NodeBuilder stores as node.Format["lineDash"] but inner text is OOXML enum name
        // The value should be readable as "lgDash" or "longdash"
        node.Format.Should().ContainKey("lineDash",
            "lineDash key should be present after setting");
    }

    // =====================================================================
    // Bug3611: PPTX shape Set name property
    // Verify that setting the "name" property on a shape updates the shape name
    // and it's readable back via Get.
    // =====================================================================
    [Fact]
    public void Bug3611_Pptx_Set_Shape_Name_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Named" });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["name"] = "MyCustomShape"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["name"].ToString().Should().Be("MyCustomShape");
    }

    // =====================================================================
    // Bug3612: Word Set footnote with unsupported property — not reported
    // The footnote Set handler only handles "text". Any other property
    // should be added to the unsupported list, but the code doesn't have
    // a default case for non-text properties.
    // =====================================================================
    [Fact]
    public void Bug3612_Word_Set_Footnote_Unsupported_Property_Not_Reported()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Footnoted text" });
        handler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Note" });

        var unsupported = handler.Set("/footnote[1]", new()
        {
            ["bold"] = "true"
        });

        unsupported.Should().Contain("bold",
            "Footnote Set should report unsupported properties like 'bold'");
    }

    // =====================================================================
    // Bug3613: Word Set endnote with unsupported property — not reported
    // Same issue as Bug3612 but for endnotes.
    // =====================================================================
    [Fact]
    public void Bug3613_Word_Set_Endnote_Unsupported_Property_Not_Reported()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Endnoted text" });
        handler.Add("/body/p[1]", "endnote", null, new() { ["text"] = "End note" });

        var unsupported = handler.Set("/endnote[1]", new()
        {
            ["font"] = "Arial"
        });

        unsupported.Should().Contain("font",
            "Endnote Set should report unsupported properties like 'font'");
    }

    // =====================================================================
    // Bug3614: PPTX shape multiline text + align together
    // Align modifies paragraph properties. When text replaces all paragraphs,
    // alignment was applied to the OLD paragraphs. The new paragraphs
    // created by multiline text won't have the alignment.
    // =====================================================================
    [Fact]
    public void Bug3614_Pptx_Set_MultilineText_And_Align_Together()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Init" });

        // Set multiline text AND alignment together in one call
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line 1\\nLine 2",
            ["align"] = "center"
        });

        // "align" iterates shape.TextBody.Elements<Paragraph>() which was just replaced
        // by "text". Since "text" case replaces paragraphs, and "align" also uses
        // the TextBody (fresh access), alignment SHOULD be applied. But dict iteration
        // order means "text" comes first (alphabetically), then "align" runs on
        // the NEW paragraphs. Wait... dict ordering depends on insertion order:
        // { ["text"] = ..., ["align"] = ... } — "text" first, then "align".
        // "text" creates new paragraphs. "align" iterates shape.TextBody paragraphs
        // which are now the new ones. So this should work.
        // But the question is whether the align code uses the shape's TextBody
        // or the stale `runs` list. Let's check: "align" case (line 244-253)
        // uses shape.TextBody.Elements<Paragraph>() — fresh access, not stale.
        // So this should work correctly. Let's verify.

        // We can test a case where the order matters differently
        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line 1");
    }

    // =====================================================================
    // Bug3615: PPTX shape multiline text + lineSpacing — uses stale paragraphs?
    // lineSpacing (line 451-463) iterates shape.TextBody.Elements<Paragraph>().
    // If "text" is processed first (alphabetically "l" > "t"? No, "l" < "t").
    // So "linespacing" comes before "text" in iteration order.
    // This means lineSpacing is applied to OLD paragraphs, then "text"
    // replaces them all! The new paragraphs won't have lineSpacing.
    // =====================================================================
    [Fact]
    public void Bug3615_Pptx_Set_MultilineText_And_LineSpacing_OrderBug()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Init" });

        // Dict iteration: "linespacing" (l) before "text" (t)
        // So lineSpacing is applied to old paragraphs, then text replaces them
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["linespacing"] = "1.5",
            ["text"] = "Line 1\\nLine 2"
        });

        // The new paragraphs created by "text" should NOT have lineSpacing
        // because lineSpacing was applied to the old (now removed) paragraphs.
        // But paraProps from first paragraph are cloned to new ones by "text" case.
        // Let's check: "text" case (line 46-47) clones paraProps from first paragraph.
        // So if lineSpacing was set on the first paragraph BEFORE text replaces it,
        // the cloned paraProps should carry lineSpacing forward.
        // Actually, dict iteration order in .NET is insertion order, not alphabetical.
        // So { ["linespacing"] = "1.5", ["text"] = ... } processes linespacing first.
        // lineSpacing sets on old paragraphs, text clones old first para props to new.
        // This should work. Let's verify anyway.
        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line 1");
    }

    // =====================================================================
    // Bug3616: PPTX Set shape rotation then Get — verify roundtrip
    // =====================================================================
    [Fact]
    public void Bug3616_Pptx_Set_Shape_Rotation_Roundtrip()
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
        // Rotation should be reported in the Format
        // But NodeBuilder may not read rotation from Transform2D
        // Let's check if it's reported
        node.Format.Should().ContainKey("rotation",
            "Shape rotation should be reported in Format");
    }

    // =====================================================================
    // Bug3617: PPTX Set connector linewidth then Get — verify roundtrip
    // =====================================================================
    [Fact]
    public void Bug3617_Pptx_Set_Connector_LineWidth_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["linecolor"] = "FF0000",
            ["linewidth"] = "2pt"
        });

        var connectors = handler.Query("connector");
        connectors.Should().NotBeEmpty();
        var conn = connectors[0];
        conn.Format.Should().ContainKey("lineWidth",
            "Connector should report lineWidth");
    }

    // =====================================================================
    // Bug3618: PPTX shape Set valign then Get — verify roundtrip
    // The NodeBuilder ShapeToNode doesn't seem to read valign (anchor).
    // =====================================================================
    [Fact]
    public void Bug3618_Pptx_Shape_Valign_Not_Reported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Centered",
            ["valign"] = "center"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("valign",
            "Shape vertical alignment should be reported in Format");
    }

    // =====================================================================
    // Bug3619: PPTX shape Set margin then Get — verify roundtrip
    // The NodeBuilder ShapeToNode doesn't seem to read text margins/insets.
    // =====================================================================
    [Fact]
    public void Bug3619_Pptx_Shape_Margin_Not_Reported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Margins",
            ["margin"] = "1cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("margin",
            "Shape text margin/inset should be reported in Format");
    }

    // =====================================================================
    // Bug3620: PPTX shape Set list style then Get — roundtrip
    // The NodeBuilder reads "list" from CharacterBullet/AutoNumberedBullet.
    // Verify bullet list style roundtrip.
    // =====================================================================
    [Fact]
    public void Bug3620_Pptx_Shape_List_Bullet_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Item 1",
            ["list"] = "bullet"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("list",
            "Shape with bullet list should report 'list' in Format");
    }

    // =====================================================================
    // Bug3621: PPTX shape Add with geometry="ellipse" — processed twice
    // The Add handler has inline preset handling (line 334-337) AND
    // effectKeys includes "geometry" (line 407). When "geometry" is provided
    // alongside "preset", it gets processed twice.
    // =====================================================================
    [Fact]
    public void Bug3621_Pptx_Add_Shape_Geometry_Processed_Twice()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // "geometry" is in effectKeys, so it will be processed by
        // SetRunOrShapeProperties after inline processing already set preset.
        // However, "geometry" sets CustomGeometry in SetRunOrShapeProperties,
        // while inline code sets PresetGeometry. They're different code paths.
        // But passing "preset" + "geometry" may conflict.
        // Let's just test "geometry" alone via Add to check for double processing.
        var result = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Custom shape",
            ["geometry"] = "M 0,0 L 100,0 L 100,100 L 0,100 Z"
        });

        var node = handler.Get(result);
        // The shape should have custom geometry, not preset
        // But inline code always sets PresetGeometry (line 334-337) using
        // properties.GetValueOrDefault("preset", "rect"), then effectKeys
        // delegation processes "geometry" which removes PresetGeometry and
        // sets CustomGeometry. This should result in a custom geometry shape.
        node.Should().NotBeNull();
    }

    // =====================================================================
    // Bug3622: Word Set paragraph text replaces runs — verifying font survives
    // When Set("/body/p[1]", {"text": "new", "font": "Arial"}),
    // the paragraph switch handles "font" via default case (applies to runs),
    // then "text" replaces runs. The font setting is lost.
    // =====================================================================
    [Fact]
    public void Bug3622_Word_Set_Paragraph_Text_And_Font_Together()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Initial" });

        // Dict order: "font" (f) inserted before "text" (t)
        // font is applied to the existing run, then "text" replaces runs.
        // "text" (line 929-952) either updates first run text or creates new run.
        // It updates firstText, then removes extra runs (i=1..N).
        // Since there's only one run, it just sets the text. Font survives.
        handler.Set("/body/p[1]", new()
        {
            ["font"] = "Courier New",
            ["text"] = "Updated"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Updated");
        node.Format.Should().ContainKey("font");
        node.Format["font"].ToString().Should().Contain("Courier New");
    }

    // =====================================================================
    // Bug3623: Word watermark color — SanitizeHex strips # but VML needs it
    // The watermark Set handler calls SanitizeHex(value) which strips the
    // '#' prefix. But VML fillcolor attribute expects "#RRGGBB" format.
    // So setting color="FF0000" results in fillcolor="FF0000" instead of
    // fillcolor="#FF0000".
    // =====================================================================
    [Fact]
    public void Bug3623_Word_Watermark_Color_Missing_Hash_Prefix()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Add a watermark first
        handler.Add("/body", "watermark", null, new()
        {
            ["text"] = "DRAFT"
        });

        // Set the color
        handler.Set("/watermark", new()
        {
            ["color"] = "FF0000"
        });

        // Get the watermark and check the color
        var node = handler.Get("/watermark");
        // The color should be readable. If SanitizeHex strips # and VML
        // expects it, the color may not render correctly.
        // But from the Get perspective, we just check it was stored.
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("color",
            "Watermark should report its color");
        // VML fillcolor should have # prefix
        var colorVal = node.Format["color"].ToString()!;
        colorVal.Should().StartWith("#",
            "VML fillcolor requires # prefix, but SanitizeHex strips it");
    }

    // =====================================================================
    // Bug3624: PPTX Set shape alignment — verify Get reports it
    // The NodeBuilder doesn't read paragraph alignment.
    // =====================================================================
    [Fact]
    public void Bug3624_Pptx_Shape_Alignment_Not_Reported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Centered",
            ["align"] = "center"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("align",
            "Shape text alignment should be reported in Format");
    }

    // =====================================================================
    // Bug3625: PPTX shape lineSpacing Set then Get — verify reported
    // NodeBuilder doesn't read lineSpacing from paragraph properties.
    // =====================================================================
    [Fact]
    public void Bug3625_Pptx_Shape_LineSpacing_Not_Reported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Spaced",
            ["linespacing"] = "2.0"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineSpacing",
            "Shape lineSpacing should be reported in Format");
    }

    // =====================================================================
    // Bug3626: Word section Set orientation without swapping width/height
    // When changing orientation from portrait to landscape, the page size
    // width and height should be swapped. But the code only sets the Orient
    // attribute without performing the swap (line 376-378).
    // =====================================================================
    [Fact]
    public void Bug3626_Word_Section_Orientation_NoSwap()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Get initial dimensions (portrait: width < height)
        var secBefore = handler.Get("/section[1]");
        var widthBefore = secBefore.Format.ContainsKey("pageWidth")
            ? Convert.ToInt32(secBefore.Format["pageWidth"]) : 12240;
        var heightBefore = secBefore.Format.ContainsKey("pageHeight")
            ? Convert.ToInt32(secBefore.Format["pageHeight"]) : 15840;

        // Set to landscape
        handler.Set("/section[1]", new()
        {
            ["orientation"] = "landscape"
        });

        var secAfter = handler.Get("/section[1]");
        var widthAfter = secAfter.Format.ContainsKey("pageWidth")
            ? Convert.ToInt32(secAfter.Format["pageWidth"]) : widthBefore;
        var heightAfter = secAfter.Format.ContainsKey("pageHeight")
            ? Convert.ToInt32(secAfter.Format["pageHeight"]) : heightBefore;

        // For landscape, width should be > height
        widthAfter.Should().BeGreaterThan(heightAfter,
            "Landscape orientation should swap width and height so width > height. " +
            "The code only sets Orient attribute without swapping dimensions.");
    }

    // =====================================================================
    // Bug3627: PPTX shape multiline text + font + size + bold + color
    // The ultimate stale runs test: all run-level properties set together
    // with multiline text replacement.
    // =====================================================================
    [Fact]
    public void Bug3627_Pptx_Set_MultilineText_AllRunProperties_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Init" });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line 1\\nLine 2\\nLine 3",
            ["font"] = "Impact",
            ["size"] = "36",
            ["bold"] = "true",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line 1");
        node.Format.Should().ContainKey("font",
            "Font should be applied to new runs after multiline text replacement");
        node.Format["font"].ToString().Should().Be("Impact");
        node.Format.Should().ContainKey("size",
            "Size should be applied to new runs");
        node.Format.Should().ContainKey("bold",
            "Bold should be applied to new runs");
        node.Format.Should().ContainKey("color",
            "Color should be applied to new runs");
    }
}
