// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// BugHuntPart50: Tests for PPTX opacity (alpha) value handling.
/// Verifies that opacity values in 0-100 percentage form and 0.0-1.0 decimal form
/// both produce valid OOXML alpha values (0-100000 range).
/// </summary>
public class MixedRegression50 : IDisposable
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

    [Fact]
    public void Add_OpacityAsPercentage_ShouldProduceValidAlpha()
    {
        // Bug: opacity='30' (meaning 30%) was computed as 30*100000=3000000,
        // exceeding OOXML alpha max of 100000, corrupting the file.
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["fill"] = "CCCCCC",
            ["opacity"] = "30"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        // opacity='30' should be normalized to 0.3 (30%)
        node!.Format.Should().ContainKey("opacity");
        var opacityVal = double.Parse(node.Format["opacity"].ToString()!);
        opacityVal.Should().BeInRange(0.0, 1.0, "opacity must be in valid 0.0-1.0 range after normalization");
        opacityVal.Should().BeApproximately(0.3, 0.01);
    }

    [Fact]
    public void Add_OpacityAsDecimal_ShouldProduceValidAlpha()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["fill"] = "FF0000",
            ["opacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("opacity");
        var opacityVal = double.Parse(node.Format["opacity"].ToString()!);
        opacityVal.Should().BeApproximately(0.5, 0.01);
    }

    [Fact]
    public void Set_OpacityAsPercentage_ShouldProduceValidAlpha()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["fill"] = "0000FF"
        });

        // Set opacity using percentage form
        handler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "50" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("opacity");
        var opacityVal = double.Parse(node.Format["opacity"].ToString()!);
        opacityVal.Should().BeInRange(0.0, 1.0, "opacity must be in valid 0.0-1.0 range after normalization");
        opacityVal.Should().BeApproximately(0.5, 0.01);
    }

    [Fact]
    public void Set_OpacityAsDecimal_ShouldProduceValidAlpha()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["fill"] = "0000FF"
        });

        handler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.75" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("opacity");
        var opacityVal = double.Parse(node.Format["opacity"].ToString()!);
        opacityVal.Should().BeApproximately(0.75, 0.01);
    }

    [Fact]
    public void Add_Opacity100Percent_ShouldClampToOne()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["fill"] = "AABBCC",
            ["opacity"] = "100"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("opacity");
        var opacityVal = double.Parse(node.Format["opacity"].ToString()!);
        opacityVal.Should().BeApproximately(1.0, 0.01);
    }

    [Fact]
    public void Opacity_PersistsAfterReopen()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["preset"] = "rect",
            ["fill"] = "CCCCCC",
            ["opacity"] = "30"
        });

        handler = Reopen(handler, path);

        var node = handler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("opacity");
        var opacityVal = double.Parse(node.Format["opacity"].ToString()!);
        opacityVal.Should().BeApproximately(0.3, 0.01);
        handler.Dispose();
    }
}
