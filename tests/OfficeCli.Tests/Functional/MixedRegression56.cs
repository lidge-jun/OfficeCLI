// Bug hunt Part 56 — Schema order and transition bugs:
// 1. chartFill spPr must come AFTER chart element in chartSpace (was InsertBefore, now InsertAfter)
// 2. effectLst children must follow schema order: glow before outerShdw (was AppendChild, now InsertEffectInOrder)
// 3. p14 modern transitions (flip/prism/doors/vortex) must use mc:AlternateContent wrapper
// 4. FlipTransition/SwitchTransition must have dir attribute (missing dir crashes PowerPoint)

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class MixedRegression56 : IDisposable
{
    private readonly string _pptxPath;
    private PowerPointHandler _pptxHandler;

    public MixedRegression56()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt56_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_pptxPath);
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _pptxHandler.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void Reopen()
    {
        _pptxHandler.Dispose();
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    private DocumentFormat.OpenXml.Packaging.PresentationDocument GetDoc() =>
        (DocumentFormat.OpenXml.Packaging.PresentationDocument)_pptxHandler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_pptxHandler)!;

    // ── Bug 1: chartFill spPr must come AFTER chart in chartSpace ────

    [Fact]
    public void Pptx_ChartFill_SpPr_ComesAfterChart()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:10,20,30",
            ["chartFill"] = "1A2744"
        });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var chartPart = slidePart.ChartParts.First();
        var chartSpace = chartPart.ChartSpace;

        var children = chartSpace.ChildElements.Select(c => c.LocalName).ToList();
        var chartIdx = children.IndexOf("chart");
        var spPrIdx = children.IndexOf("spPr");

        chartIdx.Should().BeGreaterThanOrEqualTo(0, "chart element should exist");
        spPrIdx.Should().BeGreaterThanOrEqualTo(0, "spPr element should exist");
        spPrIdx.Should().BeGreaterThan(chartIdx, "spPr must come AFTER chart in chartSpace schema order");
    }

    [Fact]
    public void Pptx_ChartFill_Set_SpPr_StillAfterChart()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:10,20,30"
        });

        // Set chartFill after creation
        _pptxHandler.Set("/slide[1]/chart[1]", new() { ["chartFill"] = "0D1B2A" });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var chartPart = slidePart.ChartParts.First();
        var children = chartPart.ChartSpace.ChildElements.Select(c => c.LocalName).ToList();

        var chartIdx = children.IndexOf("chart");
        var spPrIdx = children.IndexOf("spPr");
        spPrIdx.Should().BeGreaterThan(chartIdx, "spPr must come AFTER chart even when Set after creation");
    }

    [Fact]
    public void Pptx_ChartFill_Persist_SpPrOrderSurvivesReopen()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:10,20,30",
            ["chartFill"] = "1A2744"
        });

        Reopen();

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var chartPart = slidePart.ChartParts.First();
        var children = chartPart.ChartSpace.ChildElements.Select(c => c.LocalName).ToList();

        var chartIdx = children.IndexOf("chart");
        var spPrIdx = children.IndexOf("spPr");
        spPrIdx.Should().BeGreaterThan(chartIdx, "spPr order must survive reopen");
    }

    // ── Bug 2: effectLst children must follow schema order ───────────

    [Fact]
    public void Set_Shadow_ThenGlow_GlowBeforeShadowInEffectLst()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "test",
            ["shadow"] = "000000-6-315-4-30"
        });

        // Add glow AFTER shadow already exists
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["glow"] = "6C63FF-12-25" });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var shape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().First();
        var effectLst = shape.Descendants<DocumentFormat.OpenXml.Drawing.EffectList>().First();
        var children = effectLst.ChildElements.Select(c => c.LocalName).ToList();

        var glowIdx = children.IndexOf("glow");
        var outerShdwIdx = children.IndexOf("outerShdw");

        glowIdx.Should().BeGreaterThanOrEqualTo(0, "glow should exist");
        outerShdwIdx.Should().BeGreaterThanOrEqualTo(0, "outerShdw should exist");
        glowIdx.Should().BeLessThan(outerShdwIdx,
            "glow must come BEFORE outerShdw in effectLst schema order");
    }

    [Fact]
    public void Add_Shape_WithShadowAndGlow_GlowBeforeShadow()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "test",
            ["shadow"] = "000000-6-315-4-30",
            ["glow"] = "FF0000-10-50"
        });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var shape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().First();
        var effectLst = shape.Descendants<DocumentFormat.OpenXml.Drawing.EffectList>().First();
        var children = effectLst.ChildElements.Select(c => c.LocalName).ToList();

        var glowIdx = children.IndexOf("glow");
        var outerShdwIdx = children.IndexOf("outerShdw");
        glowIdx.Should().BeLessThan(outerShdwIdx,
            "glow must come BEFORE outerShdw even when both set in Add");
    }

    [Fact]
    public void Set_Shadow_ThenGlow_ThenReflection_SchemaOrder()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });

        // Add effects in reverse schema order
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["reflection"] = "half" });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["shadow"] = "000000-6-315-4-30" });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["glow"] = "FF0000-8" });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var shape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().First();
        var effectLst = shape.Descendants<DocumentFormat.OpenXml.Drawing.EffectList>().First();
        var children = effectLst.ChildElements.Select(c => c.LocalName).ToList();

        var glowIdx = children.IndexOf("glow");
        var outerShdwIdx = children.IndexOf("outerShdw");
        var reflectionIdx = children.IndexOf("reflection");

        glowIdx.Should().BeLessThan(outerShdwIdx, "glow < outerShdw");
        outerShdwIdx.Should().BeLessThan(reflectionIdx, "outerShdw < reflection");
    }

    [Fact]
    public void Set_AllEffects_WithSoftEdge_FullSchemaOrder()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });

        // Add in scrambled order
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["softEdge"] = "5" });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["shadow"] = "000000-4" });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["reflection"] = "tight" });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["glow"] = "00FF00-10" });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var shape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().First();
        var effectLst = shape.Descendants<DocumentFormat.OpenXml.Drawing.EffectList>().First();
        var children = effectLst.ChildElements.Select(c => c.LocalName).ToList();

        var glowIdx = children.IndexOf("glow");
        var outerShdwIdx = children.IndexOf("outerShdw");
        var reflectionIdx = children.IndexOf("reflection");
        var softEdgeIdx = children.IndexOf("softEdge");

        glowIdx.Should().BeLessThan(outerShdwIdx, "glow < outerShdw");
        outerShdwIdx.Should().BeLessThan(reflectionIdx, "outerShdw < reflection");
        reflectionIdx.Should().BeLessThan(softEdgeIdx, "reflection < softEdge");
    }

    [Fact]
    public void EffectLstOrder_GlowShadow_Persist_SurvivesReopen()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "test",
            ["shadow"] = "000000-6-315-4-30"
        });
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["glow"] = "6C63FF-12-25" });

        Reopen();

        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Format["glow"].Should().NotBeNull();
        node.Format["shadow"].Should().NotBeNull();

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("effectLst order should produce no validation errors after reopen");
    }

    // ── Bug 3: p14 transitions need mc:AlternateContent wrapper ──────

    [Theory]
    [InlineData("flip")]
    [InlineData("prism")]
    [InlineData("doors")]
    [InlineData("vortex")]
    [InlineData("honeycomb")]
    [InlineData("flash")]
    public void P14Transition_HasMcAlternateContent(string transType)
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/", "slide", null, new() { ["transition"] = transType });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.Skip(1).First();
        var slide = slidePart.Slide;
        var xml = slide.OuterXml;

        // Should have AlternateContent wrapper
        xml.Should().Contain("AlternateContent", $"{transType} transition should be wrapped in mc:AlternateContent");
        xml.Should().Contain("mc:Choice", $"{transType} transition should have mc:Choice");
        xml.Should().Contain("mc:Fallback", $"{transType} transition should have mc:Fallback for graceful degradation");
        xml.Should().Contain($"p14:{transType}", $"mc:Choice should contain p14:{transType}");

        // Should NOT have a standalone p:transition outside AlternateContent
        var typedTransitions = slide.Elements<DocumentFormat.OpenXml.Presentation.Transition>().Count();
        typedTransitions.Should().Be(0,
            $"{transType} should be inside AlternateContent, not a standalone p:transition");
    }

    [Fact]
    public void P14Transition_Fallback_ContainsFade()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/", "slide", null, new() { ["transition"] = "vortex" });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.Skip(1).First();
        var xml = slidePart.Slide.OuterXml;

        // Fallback should contain a standard fade transition
        var fallbackMatch = System.Text.RegularExpressions.Regex.Match(
            xml, @"<mc:Fallback>(.*?)</mc:Fallback>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        fallbackMatch.Success.Should().BeTrue();
        fallbackMatch.Groups[1].Value.Should().Contain("fade",
            "mc:Fallback should contain p:fade for graceful degradation in older PowerPoint");
    }

    [Fact]
    public void P14Transition_ReadBack_ReturnsCorrectType()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/", "slide", null, new() { ["transition"] = "flip" });

        var node = _pptxHandler.Get("/slide[2]");
        node.Format["transition"].Should().Be("flip");
    }

    [Fact]
    public void P14Transition_Persist_SurvivesReopen()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/", "slide", null, new() { ["transition"] = "prism" });

        Reopen();

        var node = _pptxHandler.Get("/slide[2]");
        node.Format["transition"].Should().Be("prism");
    }

    // ── Bug 4: FlipTransition must have dir attribute ────────────────

    [Fact]
    public void FlipTransition_HasDirAttribute()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/", "slide", null, new() { ["transition"] = "flip" });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.Skip(1).First();
        var xml = slidePart.Slide.OuterXml;

        // flip element should have dir attribute
        var flipMatch = System.Text.RegularExpressions.Regex.Match(xml, @"<p14:flip([^/]*)/>");
        flipMatch.Success.Should().BeTrue("flip element should exist");
        flipMatch.Groups[1].Value.Should().Contain("dir=",
            "FlipTransition must have dir attribute — missing dir crashes PowerPoint");
    }

    [Fact]
    public void SwitchTransition_HasDirAttribute()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/", "slide", null, new() { ["transition"] = "switch" });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.Skip(1).First();
        var xml = slidePart.Slide.OuterXml;

        var switchMatch = System.Text.RegularExpressions.Regex.Match(xml, @"<p14:switch([^/]*)/>");
        switchMatch.Success.Should().BeTrue("switch element should exist");
        switchMatch.Groups[1].Value.Should().Contain("dir=",
            "SwitchTransition must have dir attribute — same issue as flip");
    }

    [Theory]
    [InlineData("flip-left", "l")]
    [InlineData("flip-right", "r")]
    [InlineData("switch-left", "l")]
    [InlineData("switch-right", "r")]
    public void FlipSwitch_DirValue_MatchesInput(string transition, string expectedDir)
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/", "slide", null, new() { ["transition"] = transition });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.Skip(1).First();
        var xml = slidePart.Slide.OuterXml;

        var typeName = transition.Split('-')[0];
        var match = System.Text.RegularExpressions.Regex.Match(xml, $@"<p14:{typeName}[^/]*dir=""(\w+)""");
        match.Success.Should().BeTrue($"{typeName} should have dir attribute");
        match.Groups[1].Value.Should().Be(expectedDir);
    }

    [Fact]
    public void FlipTransition_DefaultDir_IsLeft()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/", "slide", null, new() { ["transition"] = "flip" });

        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.Skip(1).First();
        var xml = slidePart.Slide.OuterXml;

        xml.Should().Contain(@"dir=""l""",
            "flip without explicit direction should default to left");
    }
}
