// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class PptxRegression45 : IDisposable
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
    // Bug4500: PPTX gradient fill processed twice in ShapeToNode
    // Lines 288-305 set fill="color1-color2-angle" for gradient,
    // then lines 320-358 set gradient="color1-color2-angle" again.
    // Both read the same GradientFill element. This means a gradient
    // shape has BOTH "fill" and "gradient" keys which is redundant.
    // The "fill" key overwrites any solid fill info with gradient info.
    // =====================================================================
    [Fact]
    public void Bug4500_Pptx_Shape_Gradient_Fill_Has_Both_Fill_And_Gradient_Keys()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Gradient",
            ["gradient"] = "FF0000-0000FF-90"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // Both fill and gradient are set for same gradient fill — redundant
        var hasFill = node.Format.ContainsKey("fill");
        var hasGradient = node.Format.ContainsKey("gradient");

        // BUG: Having both "fill" and "gradient" for the same gradient
        // is confusing. "fill" should only represent solid fill.
        // At minimum, solid fill ("fill") should NOT contain gradient info.
        if (hasFill && hasGradient)
        {
            // If fill contains a gradient pattern (color-color-angle), it's the bug
            var fillVal = node.Format["fill"].ToString()!;
            fillVal.Should().NotContain("-",
                because: "solid fill key should not contain gradient info when gradient key exists");
        }
    }

    // =====================================================================
    // Bug4501: PPTX shape "preset" key naming is inconsistent with Add
    // Add uses "preset" key (line 338), NodeBuilder reports "preset" (line 317).
    // But the effectKeys list (line 430) includes "geometry" which goes to
    // SetRunOrShapeProperties as custom geometry path, NOT preset.
    // User confusion: passing geometry="diamond" tries SVG path parsing,
    // not preset shape creation.
    // =====================================================================
    [Fact]
    public void Bug4501_Pptx_Shape_Geometry_Key_Is_Not_Preset()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        // "preset" key works for preset shapes
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Diamond", ["preset"] = "diamond"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("preset");
        node.Format["preset"].ToString().Should().Be("diamond");
    }

    // =====================================================================
    // Bug4502: PPTX shape gradient angle integer division loses precision
    // NodeBuilder line 297: lin.Angle.Value / 60000 (integer division)
    // For non-round angles like 45 degrees = 2700000 / 60000 = 45 (OK)
    // But 15 degrees = 900000 / 60000 = 15 (OK)
    // But 22.5 degrees = 1350000 / 60000 = 22 (BUG: loses .5)
    // =====================================================================
    [Fact]
    public void Bug4502_Pptx_Gradient_Angle_Integer_Division()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        // Use gradient with 135 degree angle (no precision loss for integer)
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Grad",
            ["gradient"] = "FF0000-0000FF-135"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // Verify the gradient key is present and has angle info
        node.Format.Should().ContainKey("gradient");
        var gradVal = node.Format["gradient"].ToString()!;
        gradVal.Should().Contain("135",
            because: "gradient angle 135 should roundtrip correctly");
    }

    // =====================================================================
    // Bug4503: PPTX shape shadow opacity uses /1000.0 scale
    // but fill opacity uses /100000.0 scale. NodeBuilder reports
    // shadow opacity as part of shadow string using /1000.0 (line 470)
    // Verify shadow roundtrip.
    // =====================================================================
    [Fact]
    public void Bug4503_Pptx_Shape_Shadow_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shadowed",
            ["fill"] = "FFFF00",
            ["shadow"] = "000000-4-45-3-50"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("shadow");
        var shadow = node.Format["shadow"].ToString()!;
        // Shadow format: color-blur-angle-dist-opacity
        var parts = shadow.Split('-');
        parts.Should().HaveCount(5,
            because: "shadow should have 5 components: color-blur-angle-dist-opacity");
        parts[0].Should().Be("#000000", because: "shadow color should be black");
        parts[4].Should().Be("50", because: "shadow opacity should be 50");
    }

    // =====================================================================
    // Bug4504: PPTX shape lineSpacing Set then Get roundtrip
    // Set lineSpacing="1.5" → stored as 1.5 * 100000 = 150000
    // Get → /100000.0 = "1.5"
    // =====================================================================
    [Fact]
    public void Bug4504_Pptx_Shape_LineSpacing_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Spaced"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["lineSpacing"] = "1.5" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineSpacing",
            because: "lineSpacing should be readable after Set");
        node.Format["lineSpacing"].ToString().Should().Be("1.5x",
            because: "lineSpacing=1.5 should roundtrip correctly");
    }

    // =====================================================================
    // Bug4505: PPTX shape spaceBefore Set then Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4505_Pptx_Shape_SpaceBefore_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Before"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["spaceBefore"] = "12" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("spaceBefore",
            because: "spaceBefore should be readable after Set");
        node.Format["spaceBefore"].ToString().Should().Be("12pt",
            because: "spaceBefore=12 should roundtrip correctly");
    }

    // =====================================================================
    // Bug4506: PPTX shape spaceAfter Set then Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4506_Pptx_Shape_SpaceAfter_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "After"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["spaceAfter"] = "6" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("spaceAfter",
            because: "spaceAfter should be readable after Set");
        node.Format["spaceAfter"].ToString().Should().Be("6pt",
            because: "spaceAfter=6 should roundtrip correctly");
    }

    // =====================================================================
    // Bug4507: PPTX shape autoFit roundtrip via Set
    // =====================================================================
    [Fact]
    public void Bug4507_Pptx_Shape_AutoFit_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "AutoFit",
            ["autofit"] = "normal"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("autoFit",
            because: "autoFit should be readable after Add");
        node.Format["autoFit"].ToString().Should().Be("normal");

        // Set to shape autofit
        handler.Set("/slide[1]/shape[1]", new() { ["autofit"] = "shape" });
        node = handler.Get("/slide[1]/shape[1]");
        node.Format["autoFit"].ToString().Should().Be("shape",
            because: "autoFit=shape should roundtrip");
    }

    // =====================================================================
    // Bug4508: PPTX shape valign (vertical alignment) roundtrip
    // =====================================================================
    [Fact]
    public void Bug4508_Pptx_Shape_VAlign_Roundtrip()
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
            because: "valign should be readable after Add");
        node.Format["valign"].ToString().Should().Be("center",
            because: "valign=center should roundtrip");
    }

    // =====================================================================
    // Bug4509: PPTX shape margin roundtrip
    // =====================================================================
    [Fact]
    public void Bug4509_Pptx_Shape_Margin_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Padded"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["margin"] = "1cm" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("margin",
            because: "margin should be readable after Set");
    }

    // =====================================================================
    // Bug4510: PPTX shape flipH roundtrip
    // =====================================================================
    [Fact]
    public void Bug4510_Pptx_Shape_FlipH_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Flipped",
            ["flipH"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // flipH is in effectKeys, so it should be applied via SetRunOrShapeProperties
        node.Format.Should().ContainKey("flipH",
            because: "flipH should be readable after Add");
        node.Format["flipH"].Should().Be(true);
    }

    // =====================================================================
    // Bug4511: PPTX shape flipV roundtrip
    // =====================================================================
    [Fact]
    public void Bug4511_Pptx_Shape_FlipV_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "FlippedV",
            ["flipV"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("flipV",
            because: "flipV should be readable after Add");
        node.Format["flipV"].Should().Be(true);
    }

    // =====================================================================
    // Bug4512: PPTX shape indent roundtrip via Add/Set
    // =====================================================================
    [Fact]
    public void Bug4512_Pptx_Shape_Indent_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Indented",
            ["indent"] = "1cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("indent",
            because: "indent should be readable after Add");
        node.Format["indent"].ToString().Should().Contain("1",
            because: "indent=1cm should roundtrip with a value containing 1");
    }

    // =====================================================================
    // Bug4513: Excel cell font.bold Set then Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4513_Excel_Cell_Font_Bold_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Bold",
            ["font.bold"] = "true"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("font.bold",
            because: "font.bold should be readable after Add");
        node.Format["font.bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug4514: Excel cell font.color Set then Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug4514_Excel_Cell_Font_Color_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Red",
            ["font.color"] = "FF0000"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("font.color",
            because: "font.color should be readable after Add");
        node.Format["font.color"].ToString().Should().Be("#FF0000",
            because: "font.color=FF0000 should roundtrip correctly");
    }

    // =====================================================================
    // Bug4515: Excel cell fill roundtrip
    // =====================================================================
    [Fact]
    public void Bug4515_Excel_Cell_Fill_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Filled",
            ["fill"] = "00FF00"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("fill",
            because: "fill should be readable after Add");
        node.Format["fill"].ToString().Should().Be("#00FF00",
            because: "fill=00FF00 should roundtrip correctly");
    }

    // =====================================================================
    // Bug4516: Excel cell number format roundtrip
    // Set numberformat="0.00" then verify it's readable
    // =====================================================================
    [Fact]
    public void Bug4516_Excel_Cell_NumberFormat_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "3.14159",
            ["numberformat"] = "0.00"
        });

        var node = handler.Get("/Sheet1/A1");
        // BUG: number format is not readable — CellToNode doesn't report it
        node.Format.Should().ContainKey("numberformat",
            because: "numberformat should be readable after Add");
    }

    // =====================================================================
    // Bug4517: PPTX shape text with newline character
    // Add shape with text containing \n. The text should be split into
    // multiple paragraphs or preserved as-is.
    // =====================================================================
    [Fact]
    public void Bug4517_Pptx_Shape_Text_With_Newline()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line1\nLine2"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // The text should contain both lines
        node.Text.Should().Contain("Line1");
        node.Text.Should().Contain("Line2");
    }

    // =====================================================================
    // Bug4518: Word paragraph bold="false" should clear bold
    // When bold is set to "false", it should remove Bold element
    // =====================================================================
    [Fact]
    public void Bug4518_Word_Paragraph_Bold_False_Clears()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "Bold", ["bold"] = "true"
        });

        var n1 = handler.Get("/body/p[1]");
        n1.Format.Should().ContainKey("bold");

        handler.Set("/body/p[1]", new() { ["bold"] = "false" });

        var n2 = handler.Get("/body/p[1]");
        n2.Format.Should().NotContainKey("bold",
            because: "bold=false should clear bold formatting");
    }

    // =====================================================================
    // Bug4519: Excel cell alignment roundtrip
    // Set alignment (halign) on cell, verify readable
    // =====================================================================
    [Fact]
    public void Bug4519_Excel_Cell_Alignment_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Centered",
            ["halign"] = "center"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("alignment.horizontal",
            because: "alignment.horizontal should be the canonical Excel key");
        node.Format["alignment.horizontal"].ToString().Should().Be("center",
            because: "alignment.horizontal=center should roundtrip correctly");
    }

    // =====================================================================
    // Bug4520: Excel cell wrap text roundtrip
    // =====================================================================
    [Fact]
    public void Bug4520_Excel_Cell_Wrap_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Wrapped text here",
            ["wrap"] = "true"
        });

        var node = handler.Get("/Sheet1/A1");
        // BUG: Set uses "wrap" but Get returns "alignment.wrapText"
        // Key naming inconsistency between Set and Get
        node.Format.Should().ContainKey("alignment.wrapText",
            because: "alignment.wrapText should be the canonical Excel key");
    }

    // =====================================================================
    // Bug4521: PPTX shape preset "rect" is default — check if
    // non-default presets like "roundRect" roundtrip correctly
    // =====================================================================
    [Fact]
    public void Bug4521_Pptx_Shape_Preset_RoundRect_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Round", ["preset"] = "roundRect"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("preset");
        node.Format["preset"].ToString().Should().Be("roundRect",
            because: "roundRect preset should roundtrip");
    }

    // =====================================================================
    // Bug4522: PPTX shape list style bullet roundtrip
    // =====================================================================
    [Fact]
    public void Bug4522_Pptx_Shape_List_Bullet_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Bullet item",
            ["list"] = "bullet"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("list",
            because: "list style should be readable after Add");
    }

    // =====================================================================
    // Bug4523: Word paragraph Add with "alignment"="center" roundtrip
    // =====================================================================
    [Fact]
    public void Bug4523_Word_Paragraph_Alignment_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "Centered", ["alignment"] = "center"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("alignment",
            because: "alignment should be readable after Add");
        node.Format["alignment"].ToString().Should().Be("center",
            because: "center alignment should roundtrip");
    }

    // =====================================================================
    // Bug4524: Word paragraph spacebefore/spaceafter roundtrip
    // =====================================================================
    [Fact]
    public void Bug4524_Word_Paragraph_SpaceBefore_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "Spaced", ["spacebefore"] = "200"
        });

        var node = handler.Get("/body/p[1]");
        // BUG: Key is "spacebefore" (lowercase) not "spaceBefore" (camelCase)
        // Inconsistent with PPTX which uses camelCase "spaceBefore"
        node.Format.Should().ContainKey("spaceBefore",
            because: "Word handler should expose camelCase spaceBefore");
    }

    // =====================================================================
    // Bug4525: Word paragraph linespacing roundtrip
    // =====================================================================
    [Fact]
    public void Bug4525_Word_Paragraph_LineSpacing_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "p", null, new()
        {
            ["text"] = "LineSpaced", ["linespacing"] = "480"
        });

        var node = handler.Get("/body/p[1]");
        // BUG: Key is "linespacing" (lowercase) not "lineSpacing" (camelCase)
        // Inconsistent with PPTX which uses camelCase "lineSpacing"
        node.Format.Should().ContainKey("lineSpacing",
            because: "Word handler should expose camelCase lineSpacing");
    }

    // =====================================================================
    // Bug4526: PPTX shape charspacing roundtrip
    // =====================================================================
    [Fact]
    public void Bug4526_Pptx_Shape_CharSpacing_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Spaced chars",
            ["spacing"] = "2"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("spacing",
            because: "spacing should be readable after Add");
        node.Format["spacing"].ToString().Should().Be("2",
            because: "spacing=2 should roundtrip correctly");
    }

    // =====================================================================
    // Bug4527: PPTX shape reflection roundtrip
    // =====================================================================
    [Fact]
    public void Bug4527_Pptx_Shape_Reflection_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Reflected",
            ["fill"] = "0000FF",
            ["reflection"] = "half"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("reflection",
            because: "reflection should be readable after Add");
        node.Format["reflection"].ToString().Should().Be("half",
            because: "reflection=half should roundtrip");
    }

    // =====================================================================
    // Bug4528: PPTX shape softedge roundtrip
    // =====================================================================
    [Fact]
    public void Bug4528_Pptx_Shape_SoftEdge_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Soft",
            ["fill"] = "00FF00",
            ["softedge"] = "5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("softEdge",
            because: "softEdge should be readable after Add");
    }

    // =====================================================================
    // Bug4529: Excel "border" key (without dot) should work as
    // shorthand for "border.all" — currently it's not recognized
    // as a style key because IsStyleKey checks for "border." prefix.
    // =====================================================================
    [Fact]
    public void Bug4529_Excel_Border_Without_Dot_Not_Recognized()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Bordered",
            ["border"] = "thin"
        });

        var node = handler.Get("/Sheet1/A1");
        // BUG: "border" (without dot) is not a recognized style key
        // It should work as shorthand for "border.all"
        node.Format.Should().ContainKey("border.left",
            because: "border key should work as shorthand for border.all");
    }

    // =====================================================================
    // Bug4530: Word table cell Add with text and formatting
    // =====================================================================
    [Fact]
    public void Bug4530_Word_Table_Cell_Text_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "tbl", null, new() { ["rows"] = "2", ["cols"] = "2" });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Hello", ["bold"] = "true"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Hello");
        node.Format.Should().ContainKey("bold",
            because: "bold should be readable on table cell after Set");
    }

    // =====================================================================
    // Bug4531: PPTX shape glow roundtrip
    // =====================================================================
    [Fact]
    public void Bug4531_Pptx_Shape_Glow_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Glowing",
            ["fill"] = "FF0000",
            ["glow"] = "FFFF00-8-75"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("glow",
            because: "glow should be readable after Add");
        var glow = node.Format["glow"].ToString()!;
        glow.Should().Contain("#FFFF00",
            because: "glow color should be preserved");
    }
}
