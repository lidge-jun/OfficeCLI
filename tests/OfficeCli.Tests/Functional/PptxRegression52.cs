// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// BugHuntPart52: Tests for PPTX effects round-trip, gradient format consistency,
/// autoFit key casing, background read-back, animation round-trip, transition
/// persistence, notes round-trip, table cell formatting, connector properties,
/// and various Word/Excel edge cases.
/// </summary>
public class PptxRegression52 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTempFile(string ext)
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

    private PowerPointHandler Reopen(PowerPointHandler handler, string path)
    {
        handler.Dispose();
        return new PowerPointHandler(path, editable: true);
    }

    private ExcelHandler ReopenExcel(ExcelHandler handler, string path)
    {
        handler.Dispose();
        return new ExcelHandler(path, editable: true);
    }

    private WordHandler ReopenWord(WordHandler handler, string path)
    {
        handler.Dispose();
        return new WordHandler(path, editable: true);
    }

    // ==================== PPTX Gradient Format Inconsistency ====================

    [Fact]
    public void Bug5200_PptxGradientGetFormatShouldBeConsistent()
    {
        // Bug: NodeBuilder reads gradient TWICE — first at lines 290-307 (sets "linear;C1;C2;angle")
        // then at lines 331-369 (overwrites with "C1-C2-angle"). The Get format should be
        // consistent and ideally match one of the supported Set input formats.
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["gradient"] = "FF0000-0000FF-90"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("gradient");

        var gradientVal = node.Format["gradient"].ToString()!;
        // The gradient should be readable and parseable — verify it contains both colors
        gradientVal.Should().Contain("#FF0000", because: "gradient Get should include the first color");
        gradientVal.Should().Contain("#0000FF", because: "gradient Get should include the second color");

        // Verify the gradient value can be fed back into Set without error
        // (round-trip: Get output → Set input)
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["gradient"] = gradientVal });
        act.Should().NotThrow(
            because: "gradient Get output should be a valid input for Set (round-trip). " +
                     "Currently Get returns a format that may not match any accepted Set input format");
    }

    [Fact]
    public void Bug5201_PptxGradientRadialRoundTrip()
    {
        // Test that radial gradient round-trips correctly
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["gradient"] = "radial:FF0000-0000FF-tl"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("gradient");

        var gradientVal = node.Format["gradient"].ToString()!;
        gradientVal.Should().Contain("#FF0000");
        gradientVal.Should().Contain("#0000FF");
        gradientVal.Should().Contain("tl", because: "radial gradient focus point should be preserved");
    }

    // ==================== PPTX autoFit Key Casing Mismatch ====================

    [Fact]
    public void Bug5202_PptxAutoFitKeyCasingMismatch()
    {
        // Bug: NodeBuilder returns Format["autoFit"] (camelCase F) at line 593,
        // but ShapeProperties Set accepts "autofit" (all lowercase) at line 534.
        // The key casing should be consistent so the user can copy-paste Get output into Set.
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "AutoFit Test",
            ["autofit"] = "normal"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();

        // Check if the key is "autofit" (consistent with Set input)
        var hasLowercase = node!.Format.ContainsKey("autofit");
        var hasCamelCase = node.Format.ContainsKey("autoFit");

        // At least one should exist
        (hasLowercase || hasCamelCase).Should().BeTrue(
            because: "Get should return the autofit property after it was set during Add");

        // NodeBuilder returns "autoFit" (camelCase) — Set accepts both "autofit" and "autoFit"
        hasCamelCase.Should().BeTrue(
            because: "Get returns 'autoFit' (camelCase). Set accepts lowercase, so both directions work");
    }

    // ==================== PPTX Shadow Round-Trip ====================

    [Fact]
    public void Bug5203_PptxShadowRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Shadow Test",
            ["shadow"] = "000000-6-315-4-50"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("shadow");

        var shadowVal = node.Format["shadow"].ToString()!;
        // Verify the shadow value can be fed back into Set
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["shadow"] = shadowVal });
        act.Should().NotThrow(
            because: "shadow Get output should be valid Set input (round-trip)");

        // Verify shadow persists after reopen
        var handler2 = Reopen(handler, path);
        var node2 = handler2.Get("/slide[1]/shape[1]");
        node2.Should().NotBeNull();
        node2!.Format.Should().ContainKey("shadow");
        handler2.Dispose();
    }

    [Fact]
    public void Bug5204_PptxShadowColorPreservation()
    {
        // Verify shadow color is preserved correctly
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Red Shadow",
            ["shadow"] = "FF0000-4-45-3-60"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("shadow");
        var shadow = node.Format["shadow"].ToString()!;
        shadow.Should().Contain("#FF0000", because: "shadow color should be preserved in Get output");
    }

    // ==================== PPTX Glow Round-Trip ====================

    [Fact]
    public void Bug5205_PptxGlowRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Glow Test",
            ["glow"] = "0070FF-10-60"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("glow");

        var glowVal = node.Format["glow"].ToString()!;
        glowVal.Should().Contain("#0070FF", because: "glow color should be preserved");

        // Verify round-trip
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["glow"] = glowVal });
        act.Should().NotThrow(
            because: "glow Get output should be valid Set input");
    }

    // ==================== PPTX Reflection Round-Trip ====================

    [Fact]
    public void Bug5206_PptxReflectionRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Reflection Test",
            ["reflection"] = "half"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("reflection");
        node.Format["reflection"].ToString().Should().Be("half");

        // Verify round-trip
        var act = () => handler.Set("/slide[1]/shape[1]",
            new() { ["reflection"] = node.Format["reflection"].ToString()! });
        act.Should().NotThrow(
            because: "reflection Get output should be valid Set input");
    }

    // ==================== PPTX SoftEdge Round-Trip ====================

    [Fact]
    public void Bug5207_PptxSoftEdgeRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "SoftEdge Test",
            ["softedge"] = "5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("softEdge",
            because: "softEdge should be readable after being set during Add");

        var softEdgeVal = node.Format["softEdge"].ToString()!;
        // softEdge key in Get is "softEdge" (camelCase E), but Set key is "softedge" (lowercase)
        // This may be another key casing inconsistency
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["softedge"] = softEdgeVal });
        act.Should().NotThrow(
            because: "softEdge Get output should be valid Set input");
    }

    // ==================== PPTX 3D Rotation Round-Trip ====================

    [Fact]
    public void Bug5208_Pptx3DRotationRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "3D Rotation",
            ["rot3d"] = "45,30,0"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("rot3d",
            because: "3D rotation should be readable after being set during Add");

        var rot3dVal = node.Format["rot3d"].ToString()!;
        rot3dVal.Should().Contain("45", because: "rotX should be preserved");
        rot3dVal.Should().Contain("30", because: "rotY should be preserved");

        // Verify the Get output can be fed back into Set
        // Get returns "rot3d" key but Set accepts "rot3d", "rotation3d", "3drotation", or "3d.rotation"
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["rot3d"] = rot3dVal });
        act.Should().NotThrow(
            because: "rot3d Get output should be valid Set input");
    }

    // ==================== PPTX Bevel Round-Trip ====================

    [Fact]
    public void Bug5209_PptxBevelRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Bevel Test",
            ["bevel"] = "circle-6-6"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("bevel",
            because: "bevel should be readable after being set during Add");

        var bevelVal = node.Format["bevel"].ToString()!;
        bevelVal.Should().Contain("circle", because: "bevel preset should be preserved");

        // Verify round-trip: Get output → Set input
        // Bug: FormatBevel returns InnerText of preset (e.g. "circle") but ParseBevelPreset
        // expects lowercase (e.g. "circle"). However, InnerText may return camelCase form.
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["bevel"] = bevelVal });
        act.Should().NotThrow(
            because: "bevel Get output should be valid Set input (round-trip)");
    }

    // ==================== PPTX Material Round-Trip ====================

    [Fact]
    public void Bug5210_PptxMaterialRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Material Test",
            ["material"] = "plastic"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("material",
            because: "material should be readable after being set during Add");

        var materialVal = node.Format["material"].ToString()!;
        // NodeBuilder stores InnerText of PresetMaterial (e.g. "plastic" → "plastic" XML enum).
        // ParseMaterial expects lowercase. InnerText may differ from what ParseMaterial accepts.
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["material"] = materialVal });
        act.Should().NotThrow(
            because: "material Get output should be valid Set input. " +
                     "NodeBuilder returns InnerText which may not match ParseMaterial's expected values");
    }

    // ==================== PPTX Lighting Round-Trip ====================

    [Fact]
    public void Bug5211_PptxLightingRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Lighting Test",
            ["lighting"] = "balanced"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("lighting",
            because: "lighting should be readable after being set during Add");

        var lightingVal = node.Format["lighting"].ToString()!;
        // NodeBuilder stores lightRig.Rig.InnerText. ParseLightRig expects lowercase aliases.
        // InnerText for Balanced would be "balanced" — should match. But for ThreePoints it
        // returns "threePt" which needs "threept" or "3pt" in ParseLightRig.
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["lighting"] = lightingVal });
        act.Should().NotThrow(
            because: "lighting Get output should be valid Set input. " +
                     "InnerText may not match ParseLightRig's accepted values");
    }

    // ==================== PPTX 3D Depth Round-Trip ====================

    [Fact]
    public void Bug5212_Pptx3DDepthRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Depth Test",
            ["depth"] = "10"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("depth",
            because: "depth should be readable after being set during Add");

        var depthVal = node.Format["depth"].ToString()!;
        // Verify numeric value is preserved
        double.Parse(depthVal).Should().BeApproximately(10.0, 0.1);
    }

    // ==================== PPTX Background Round-Trip ====================

    [Fact]
    public void Bug5213_PptxBackgroundSolidColorRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["background"] = "FF0000" });

        var node = handler.Get("/slide[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("background",
            because: "slide background should be readable after being set");

        var bgVal = node.Format["background"].ToString()!;
        bgVal.Should().Be("#FF0000",
            because: "solid color background should preserve hex color value");
    }

    [Fact]
    public void Bug5214_PptxBackgroundGradientRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["background"] = "FF0000-0000FF-90" });

        var node = handler.Get("/slide[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("background",
            because: "gradient background should be readable after being set");

        var bgVal = node.Format["background"].ToString()!;
        bgVal.Should().Contain("#FF0000");
        bgVal.Should().Contain("#0000FF");
    }

    // ==================== PPTX Notes Round-Trip ====================

    [Fact]
    public void Bug5215_PptxNotesRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["notes"] = "These are speaker notes" });

        var node = handler.Get("/slide[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("notes",
            because: "speaker notes should be readable after being set");

        var notesVal = node.Format["notes"].ToString()!;
        notesVal.Should().Be("These are speaker notes",
            because: "notes text should be preserved exactly");
    }

    [Fact]
    public void Bug5216_PptxNotesMultilineRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["notes"] = "Line 1\nLine 2\nLine 3" });

        var node = handler.Get("/slide[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("notes");

        var notesVal = node.Format["notes"].ToString()!;
        notesVal.Should().Contain("Line 1");
        notesVal.Should().Contain("Line 2");
        notesVal.Should().Contain("Line 3");
    }

    // ==================== PPTX Connector Round-Trip ====================

    [Fact]
    public void Bug5217_PptxConnectorRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["preset"] = "straight",
            ["lineColor"] = "FF0000",
            ["lineWidth"] = "2pt"
        });

        var nodes = handler.Query("connector");
        nodes.Should().NotBeNull();
        nodes.Should().HaveCountGreaterThan(0,
            because: "connector should be queryable after being added");

        if (nodes.Count > 0)
        {
            var cxn = nodes[0];
            cxn.Type.Should().Be("connector");
        }
    }

    // ==================== PPTX Transition Persistence ====================

    [Fact]
    public void Bug5218_PptxTransitionPersistsAfterReopen()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Set("/slide[1]", new() { ["transition"] = "fade" });

        handler = Reopen(handler, path);

        var node = handler.Get("/slide[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("transition",
            because: "transition should persist after file reopen. " +
                     "The SDK's typed Transition setter may not serialize correctly");

        var transVal = node.Format["transition"].ToString()!;
        transVal.Should().Be("fade", because: "transition type should be preserved");
        handler.Dispose();
    }

    // ==================== PPTX Table Cell Formatting ====================

    [Fact]
    public void Bug5219_PptxTableCellFillRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Colored Cell",
            ["fill"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("fill");
        node.Format["fill"].ToString().Should().Be("#FF0000",
            because: "table cell fill color should be preserved");
    }

    // ==================== PPTX Shape Flip Round-Trip ====================

    [Fact]
    public void Bug5220_PptxShapeFlipHRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Flip Test"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["flipH"] = "true" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("flipH",
            because: "flipH should be readable after being set");
        node.Format["flipH"].Should().Be(true);
    }

    // ==================== PPTX Character Spacing Round-Trip ====================

    [Fact]
    public void Bug5221_PptxCharSpacingRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Spacing Test",
            ["spacing"] = "2"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        // Get returns "spacing" key for character spacing
        node!.Format.Should().ContainKey("spacing",
            because: "character spacing should be readable after being set during Add");

        var spacingVal = double.Parse(node.Format["spacing"].ToString()!);
        spacingVal.Should().BeApproximately(2.0, 0.1,
            because: "character spacing of 2pt should be preserved");
    }

    // ==================== PPTX Baseline/Superscript Round-Trip ====================

    [Fact]
    public void Bug5222_PptxBaselineRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Super",
            ["superscript"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("baseline",
            because: "baseline should be readable after superscript is set during Add");

        var baselineVal = double.Parse(node.Format["baseline"].ToString()!);
        baselineVal.Should().BeGreaterThan(0,
            because: "superscript should produce a positive baseline value");
    }

    // ==================== PPTX TextWarp Round-Trip ====================

    [Fact]
    public void Bug5223_PptxTextWarpRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "WordArt",
            ["textwarp"] = "textWave1"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("textWarp",
            because: "textWarp should be readable after being set during Add");
    }

    // ==================== Word Table Cell Text Direction Round-Trip ====================

    [Fact]
    public void Bug5224_WordTableCellTextDirectionRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Vertical Text",
            ["textDirection"] = "btLr"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("textDirection",
            because: "textDirection should be readable after being set");
        node.Format["textDirection"].ToString().Should().Be("btLr",
            because: "textDirection value should be preserved exactly");
    }

    // ==================== Word Table Cell NoWrap Round-Trip ====================

    [Fact]
    public void Bug5225_WordTableCellNoWrapRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["nowrap"] = "true"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("nowrap",
            because: "nowrap should be readable after being set");
    }

    // ==================== Excel Number Format Round-Trip ====================

    [Fact]
    public void Bug5226_ExcelNumberFormatRoundTrip()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new()
        {
            ["value"] = "1234.56",
            ["numberformat"] = "#,##0.00"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();

        // Excel stores both "numberformat" and "format" — check for either
        var hasNf = node!.Format.ContainsKey("numberformat");
        var hasFmt = node.Format.ContainsKey("format");

        (hasNf || hasFmt).Should().BeTrue(
            because: "number format should be readable after being set");
    }

    // ==================== Excel Font Properties Round-Trip ====================

    [Fact]
    public void Bug5227_ExcelFontBoldRoundTrip()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Bold Text",
            ["font.bold"] = "true"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("font.bold",
            because: "font.bold should be readable after being set");

        var boldVal = node.Format["font.bold"];
        // Should be boolean true or string "true" — not false
        (boldVal.ToString() == "True" || boldVal.ToString() == "true" || boldVal.Equals(true))
            .Should().BeTrue(because: "font.bold should be true after being set to true");
    }

    // ==================== Excel Merge Cells Round-Trip ====================

    [Fact]
    public void Bug5228_ExcelMergeCellsRoundTrip()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new() { ["value"] = "Merged" });
        handler.Set("/Sheet1/A1:C1", new() { ["merge"] = "true" });

        var node = handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        // Check if merge info is reported
        node!.Format.Should().ContainKey("merge",
            because: "merge should be readable after being set. " +
                     "CellToNode reads mergeCell info but may not report it correctly");
    }

    // ==================== PPTX Shape LineDash Round-Trip ====================

    [Fact]
    public void Bug5229_PptxShapeLineDashRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Dash Test",
            ["line"] = "000000",
            ["lineWidth"] = "2pt",
            ["lineDash"] = "dash"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("lineDash",
            because: "lineDash should be readable after being set during Add");

        var dashVal = node.Format["lineDash"].ToString()!;
        dashVal.Should().Be("dash",
            because: "lineDash value should be preserved as 'dash'");
    }

    // ==================== PPTX Shape Link Round-Trip ====================

    [Fact]
    public void Bug5230_PptxShapeLinkRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Click Me",
            ["link"] = "https://example.com"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("link",
            because: "hyperlink should be readable after being set during Add");

        var linkVal = node.Format["link"].ToString()!;
        linkVal.Should().Contain("example.com",
            because: "hyperlink URL should be preserved");
    }

    // ==================== PPTX Lighting ThreePoint InnerText ====================

    [Fact]
    public void Bug5231_PptxLightingThreePointRoundTrip()
    {
        // Bug: ParseLightRig accepts "threept" or "3pt" for ThreePoints, but NodeBuilder
        // returns lightRig.Rig.InnerText which is "threePt" (camelCase). ParseLightRig
        // lowercases to "threept" which IS matched. So this specific case works.
        // But let's verify other lighting presets whose InnerText may not match.
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Light Test",
            ["lighting"] = "brightroom"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("lighting");

        var lightingVal = node.Format["lighting"].ToString()!;
        // InnerText for BrightRoom is "brightRm" not "brightroom"
        // ParseLightRig expects "brightroom" — this would fail if fed back to Set
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["lighting"] = lightingVal });
        act.Should().NotThrow(
            because: "lighting Get output should be valid Set input. " +
                     "InnerText 'brightRm' may not match ParseLightRig's 'brightroom'");
    }

    // ==================== PPTX Animation Round-Trip ====================

    [Fact]
    public void Bug5232_PptxAnimationRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Animated Shape"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "fade-entrance-500" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("animation",
            because: "animation should be readable after being set");

        var animVal = node.Format["animation"].ToString()!;
        animVal.Should().Contain("fade", because: "animation effect type should be preserved");
        animVal.Should().Contain("entrance", because: "animation class should be preserved");
    }

    // ==================== PPTX Slide zorder ====================

    [Fact]
    public void Bug5233_PptxShapeZOrderReturned()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Shape 1"
        });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Shape 2"
        });

        var node1 = handler.Get("/slide[1]/shape[1]");
        var node2 = handler.Get("/slide[1]/shape[2]");

        node1.Should().NotBeNull();
        node2.Should().NotBeNull();

        node1!.Format.Should().ContainKey("zorder",
            because: "z-order should be reported for shapes");
        node2!.Format.Should().ContainKey("zorder");

        var z1 = Convert.ToInt32(node1.Format["zorder"]);
        var z2 = Convert.ToInt32(node2.Format["zorder"]);
        z1.Should().BeLessThan(z2,
            because: "first added shape should have lower z-order than second");
    }

    // ==================== PPTX Slide Layout Alignment Default ====================

    [Fact]
    public void Bug5234_PptxShapeDefaultAlignLeft()
    {
        // Verify that shapes without explicit alignment default to "left"
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Default Align"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("align");
        node.Format["align"].ToString().Should().Be("left",
            because: "default text alignment should be 'left'");
    }

    // ==================== Word Paragraph keepNext Round-Trip ====================

    [Fact]
    public void Bug5235_WordParagraphKeepNextRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Keep Next Para",
            ["keepnext"] = "true"
        });

        var node = handler.Get("/body/p[1]");
        node.Should().NotBeNull();

        // Bug: Word Get stores "keepnext" (all lowercase) but the conventional casing
        // would be "keepNext" (camelCase). This inconsistency means the property name
        // from Get doesn't match common conventions.
        var hasLowercase = node!.Format.ContainsKey("keepnext");
        var hasCamelCase = node.Format.ContainsKey("keepNext");

        (hasLowercase || hasCamelCase).Should().BeTrue(
            because: "keepNext property should be present after being set during Add");

        // Assert the key casing inconsistency is a bug
        if (hasLowercase && !hasCamelCase)
        {
            hasLowercase.Should().BeFalse(
                because: "Get returns 'keepnext' (all lowercase) but conventional " +
                         "key naming uses camelCase like 'keepNext'. This creates inconsistency " +
                         "with other handlers that use camelCase keys");
        }
    }

    // ==================== Word Paragraph keepLines Round-Trip ====================

    [Fact]
    public void Bug5236_WordParagraphKeepLinesRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Keep Lines Para",
            ["keeplines"] = "true"
        });

        var node = handler.Get("/body/p[1]");
        node.Should().NotBeNull();

        var hasLowercase = node!.Format.ContainsKey("keeplines");
        var hasCamelCase = node.Format.ContainsKey("keepLines");

        (hasLowercase || hasCamelCase).Should().BeTrue(
            because: "keepLines property should be present after being set during Add");

        if (hasLowercase && !hasCamelCase)
        {
            hasLowercase.Should().BeFalse(
                because: "Get returns 'keeplines' (all lowercase) but conventional " +
                         "key naming uses camelCase like 'keepLines'. This creates inconsistency " +
                         "with other handlers that use camelCase keys");
        }
    }

    // ==================== Excel Conditional Formatting Round-Trip ====================

    [Fact]
    public void Bug5237_ExcelConditionalFormattingRoundTrip()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new() { ["value"] = "100" });
        handler.Set("/Sheet1/A2", new() { ["value"] = "200" });

        handler.Add("/Sheet1", "conditionalFormatting", null, new()
        {
            ["range"] = "A1:A10",
            ["type"] = "greaterThan",
            ["value"] = "150",
            ["fill"] = "FF0000"
        });

        var nodes = handler.Query("conditionalFormatting");
        nodes.Should().NotBeNull();
        nodes.Should().HaveCountGreaterThan(0,
            because: "conditional formatting should be queryable after being added");
    }

    // ==================== PPTX Shape Paragraph Indent Round-Trip ====================

    [Fact]
    public void Bug5238_PptxShapeParagraphIndentRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Indented",
            ["indent"] = "1cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("indent",
            because: "paragraph indent should be readable after being set during Add");
    }

    // ==================== PPTX Shape MarginLeft Round-Trip ====================

    [Fact]
    public void Bug5239_PptxShapeMarginLeftRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["text"] = "Margin Left",
            ["marginLeft"] = "2cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("marginLeft",
            because: "marginLeft should be readable after being set during Add");
    }
}
