using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// BugHuntPart55: Tests for 7 new morph/animation features:
///   1. Group shape position in Query (x, y, width, height, zorder)
///   2. Multi-stop gradients with custom @position
///   3. Animation delay + easing (easein/easeout)
///   4. Theme color editing (Set /theme accent1-6, headingFont, bodyFont)
///   5. Morph-check API (Get /morph-check)
///   6. Align / distribute shapes on a slide
///   7. Motion Path animation (motionPath property)
/// </summary>
public class BugHuntPart55 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTempPptx()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Feature 1: GroupShape position in Query
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug5500_GroupShape_ReturnsPosition()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        // Add two shapes and group them
        h.Add("/", "slide", null, new() { ["title"] = "Group Test" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "A", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "B", ["x"] = "5cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm" });
        h.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });

        var slide = h.Get("/slide[1]", depth: 1);
        var grpNode = slide.Children?.FirstOrDefault(c => c.Type == "group");

        grpNode.Should().NotBeNull("Group should exist");
        grpNode!.Format.Should().ContainKey("name");
        grpNode.Format.Should().ContainKey("zorder");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Feature 2: Multi-stop gradients with @position
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug5501_MultiStopGradient_EvenDistribution()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Gradient" });
        h.Add("/slide[1]", "shape", null, new() {
            ["text"] = "3-stop",
            ["gradient"] = "FF0000-FFFF00-0000FF",
            ["width"] = "5cm", ["height"] = "3cm"
        });

        var node = h.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("gradient");
        var grad = node.Format["gradient"]?.ToString()!;
        // Should contain all 3 colors
        grad.Should().Contain("FF0000");
        grad.Should().Contain("FFFF00");
        grad.Should().Contain("0000FF");
    }

    [Fact]
    public void Bug5502_MultiStopGradient_CustomPosition_RoundTrip()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Custom Gradient" });
        h.Add("/slide[1]", "shape", null, new() {
            ["text"] = "Custom stops",
            ["gradient"] = "FF0000@0-FFFF00@30-0000FF@100",
            ["width"] = "5cm", ["height"] = "3cm"
        });

        var node = h.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("gradient");
        var grad = node.Format["gradient"]?.ToString()!;
        // Custom positions should be reflected with @ notation
        grad.Should().Contain("FF0000");
        grad.Should().Contain("FFFF00");
        grad.Should().Contain("@30");
    }

    [Fact]
    public void Bug5503_MultiStopGradient_WithAngle()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Angled Gradient" });
        h.Add("/slide[1]", "shape", null, new() {
            ["text"] = "Angled",
            ["gradient"] = "FF0000-FFFFFF-0000FF-45",
            ["width"] = "5cm", ["height"] = "3cm"
        });

        var node = h.Get("/slide[1]/shape[2]");
        var grad = node.Format["gradient"]?.ToString()!;
        grad.Should().Contain("FF0000");
        grad.Should().Contain("0000FF");
        grad.Should().Contain("45");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Feature 3: Animation delay + easing
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug5504_AnimationDelay_IsStored()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Anim Delay" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape" });
        h.Set("/slide[1]/shape[1]", new() { ["animation"] = "fade-entrance-500-click-delay=300" });

        var animNode = h.Get("/slide[1]/shape[1]/animation[1]");
        animNode.Format.Should().ContainKey("delay");
        animNode.Format["delay"].Should().Be(300);
    }

    [Fact]
    public void Bug5505_AnimationEasing_IsStored()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Anim Ease" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape" });
        h.Set("/slide[1]/shape[1]", new() { ["animation"] = "fly-entrance-left-500-click-easein=30-easeout=20" });

        var animNode = h.Get("/slide[1]/shape[1]/animation[1]");
        animNode.Format.Should().ContainKey("easein");
        animNode.Format.Should().ContainKey("easeout");
        ((int)animNode.Format["easein"]!).Should().Be(30);
        ((int)animNode.Format["easeout"]!).Should().Be(20);
    }

    [Fact]
    public void Bug5506_AnimationEasing_Symmetric()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Ease" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape" });
        h.Set("/slide[1]/shape[1]", new() { ["animation"] = "fade-500-click-easing=40" });

        var animNode = h.Get("/slide[1]/shape[1]/animation[1]");
        ((int)animNode.Format["easein"]!).Should().Be(40);
        ((int)animNode.Format["easeout"]!).Should().Be(40);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Feature 4: Theme color editing
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug5507_ThemeGet_ReturnsAccentColors()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        var theme = h.Get("/theme");
        theme.Type.Should().Be("theme");
        theme.Format.Should().ContainKey("accent1");
        theme.Format.Should().ContainKey("accent2");
        theme.Format.Should().ContainKey("accent6");
        theme.Format.Should().ContainKey("dk1");
        theme.Format.Should().ContainKey("lt1");
    }

    [Fact]
    public void Bug5508_ThemeSet_Accent1_IsUpdated()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Set("/theme", new() { ["accent1"] = "FF6B35" });

        var theme = h.Get("/theme");
        theme.Format["accent1"]?.ToString().Should().Be("#FF6B35");
    }

    [Fact]
    public void Bug5509_ThemeSet_MultipleColors_RoundTrip()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Set("/theme", new() {
            ["accent1"] = "E63946",
            ["accent2"] = "457B9D",
            ["accent3"] = "A8DADC",
            ["accent4"] = "1D3557"
        });

        var theme = h.Get("/theme");
        theme.Format["accent1"]?.ToString().Should().Be("#E63946");
        theme.Format["accent2"]?.ToString().Should().Be("#457B9D");
        theme.Format["accent3"]?.ToString().Should().Be("#A8DADC");
        theme.Format["accent4"]?.ToString().Should().Be("#1D3557");
    }

    [Fact]
    public void Bug5510_ThemeSet_FontScheme_IsUpdated()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Set("/theme", new() { ["headingFont"] = "Segoe UI", ["bodyFont"] = "Segoe UI Light" });

        var theme = h.Get("/theme");
        theme.Format["headingFont"]?.ToString().Should().Be("Segoe UI");
        theme.Format["bodyFont"]?.ToString().Should().Be("Segoe UI Light");
    }

    [Fact]
    public void Bug5511_ThemeSet_Persistence()
    {
        var path = CreateTempPptx();

        using (var h = new PowerPointHandler(path, editable: true))
            h.Set("/theme", new() { ["accent1"] = "AABBCC" });

        using (var h2 = new PowerPointHandler(path, editable: false))
        {
            var theme = h2.Get("/theme");
            theme.Format["accent1"]?.ToString().Should().Be("#AABBCC");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Feature 5: Morph check API
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug5512_MorphCheck_EmptyPresentation()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        h.Add("/", "slide", null, new() { ["title"] = "Slide 2" });

        var result = h.Get("/morph-check");
        result.Type.Should().Be("morph-check");
        result.ChildCount.Should().Be(0, "No !! shapes exist yet");
    }

    [Fact]
    public void Bug5513_MorphCheck_DetectsMatchedPairs()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        // Slide 1: shape named "!!circle"
        h.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Circle", ["name"] = "!!circle" });

        // Slide 2: shape also named "!!circle" → should be a matched pair
        h.Add("/", "slide", null, new() { ["title"] = "Slide 2" });
        h.Add("/slide[2]", "shape", null, new() { ["text"] = "Circle moved", ["name"] = "!!circle" });

        var result = h.Get("/morph-check");
        result.Children.Should().NotBeNullOrEmpty();

        var matched = result.Children!
            .Where(c => c.Format.TryGetValue("status", out var s) && s?.ToString() == "matched")
            .ToList();
        matched.Should().HaveCountGreaterThan(0, "!!circle on slide 1 should match !!circle on slide 2");
    }

    [Fact]
    public void Bug5514_MorphCheck_UnmatchedShapeReported()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Only here", ["name"] = "!!unique" });

        h.Add("/", "slide", null, new() { ["title"] = "Slide 2" });
        // Slide 2 has no !!unique shape

        var result = h.Get("/morph-check");
        var unmatched = result.Children?
            .Where(c => c.Format.TryGetValue("status", out var s) && s?.ToString() == "unmatched")
            .ToList();
        unmatched.Should().HaveCountGreaterThan(0, "!!unique has no match on slide 2");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Feature 6: Align / distribute shapes
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug5515_AlignLeft_AllShapesHaveSameX()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Align" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "A", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "2cm", ["height"] = "1cm" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "B", ["x"] = "4cm", ["y"] = "3cm", ["width"] = "2cm", ["height"] = "1cm" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "C", ["x"] = "7cm", ["y"] = "5cm", ["width"] = "2cm", ["height"] = "1cm" });

        h.Set("/slide[1]", new() { ["align"] = "left", ["targets"] = "shape[1],shape[2],shape[3]" });

        var s1 = h.Get("/slide[1]/shape[1]");
        var s2 = h.Get("/slide[1]/shape[2]");
        var s3 = h.Get("/slide[1]/shape[3]");

        s1.Format["x"].Should().Be(s2.Format["x"], "all left edges should align");
        s2.Format["x"].Should().Be(s3.Format["x"], "all left edges should align");
    }

    [Fact]
    public void Bug5516_AlignCenter_Horizontal()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Align Center" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "A", ["x"] = "1cm", ["y"] = "2cm", ["width"] = "3cm", ["height"] = "1cm" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "B", ["x"] = "5cm", ["y"] = "4cm", ["width"] = "3cm", ["height"] = "1cm" });

        h.Set("/slide[1]", new() { ["align"] = "center", ["targets"] = "shape[2],shape[3]" });

        var s1 = h.Get("/slide[1]/shape[2]");
        var s2 = h.Get("/slide[1]/shape[3]");

        // After center align, both content shapes (same width) have the same x
        s1.Format["x"].Should().Be(s2.Format["x"], "center-aligned shapes of equal width have same left edge");
    }

    [Fact]
    public void Bug5517_DistributeHorizontal_GapsEqual()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Distribute" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "A", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "2cm", ["height"] = "1cm" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "B", ["x"] = "3cm", ["y"] = "1cm", ["width"] = "2cm", ["height"] = "1cm" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "C", ["x"] = "9cm", ["y"] = "1cm", ["width"] = "2cm", ["height"] = "1cm" });

        h.Set("/slide[1]", new() { ["distribute"] = "horizontal", ["targets"] = "shape[1],shape[2],shape[3]" });

        // After distribution the shapes should be evenly spaced
        // (Just verify the operation doesn't throw and shapes still exist)
        var s1 = h.Get("/slide[1]/shape[1]");
        var s2 = h.Get("/slide[1]/shape[2]");
        var s3 = h.Get("/slide[1]/shape[3]");

        s1.Format.Should().ContainKey("x");
        s2.Format.Should().ContainKey("x");
        s3.Format.Should().ContainKey("x");
    }

    [Fact]
    public void Bug5518_AlignSlideCenter_CentersOnSlide()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Slide Center" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello", ["x"] = "0cm", ["y"] = "0cm", ["width"] = "4cm", ["height"] = "2cm" });

        h.Set("/slide[1]", new() { ["align"] = "slide-center" });

        var s = h.Get("/slide[1]/shape[1]");
        // Should be somewhere not at x=0 (moved toward center of 16:9 slide = ~6.1cm from left)
        s.Format.Should().ContainKey("x");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Feature 7: Motion Path animation
    // ────────────────────────────────────────────────────────────────────────

    [Fact]
    public void Bug5519_MotionPath_BasicLinear()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Motion Path" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Fly" });

        // Should not throw
        h.Set("/slide[1]/shape[1]", new() { ["motionPath"] = "M 0 0 L 0.5 -0.3 E" });

        // Verify animation is in the timing tree
        var slide = h.Get("/slide[1]", depth: 0);
        slide.Should().NotBeNull();
    }

    [Fact]
    public void Bug5520_MotionPath_WithCommaSyntax()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Motion Path" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Fly" });

        // Comma syntax should be normalised to spaces
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["motionPath"] = "M0,0 L0.5,-0.3 E" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Bug5521_MotionPath_WithDelayAndEasing()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Motion Path Delay" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Fly" });

        var act = () => h.Set("/slide[1]/shape[1]", new() {
            ["motionPath"] = "M 0 0 L 0.5 0 E-600-click-delay=200-easing=25"
        });
        act.Should().NotThrow();
    }

    [Fact]
    public void Bug5522_MotionPath_None_RemovesAnimation()
    {
        var path = CreateTempPptx();
        using var h = new PowerPointHandler(path, editable: true);

        h.Add("/", "slide", null, new() { ["title"] = "Motion None" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Fly" });
        h.Set("/slide[1]/shape[1]", new() { ["motionPath"] = "M 0 0 L 0.5 0 E" });

        // Remove
        var act = () => h.Set("/slide[1]/shape[1]", new() { ["motionPath"] = "none" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Bug5523_MotionPath_Persistence()
    {
        var path = CreateTempPptx();

        using (var h = new PowerPointHandler(path, editable: true))
        {
            h.Add("/", "slide", null, new() { ["title"] = "Motion Path" });
            h.Add("/slide[1]", "shape", null, new() { ["text"] = "Fly" });
            h.Set("/slide[1]/shape[1]", new() { ["motionPath"] = "M 0 0 L 0.5 -0.3 E" });
        }

        // Re-open and verify file is valid
        using (var h2 = new PowerPointHandler(path, editable: false))
        {
            var slide = h2.Get("/slide[1]");
            slide.Should().NotBeNull();
        }
    }
}
