using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class WatermarkTests : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public WatermarkTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"wm_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Add_TextWatermark_DefaultProperties()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello World" });

        var result = _handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });

        result.Should().Be("/watermark");

        // Verify via Get
        var node = _handler.Get("/watermark");
        node.Type.Should().Be("watermark");
        node.Format["text"].Should().Be("DRAFT");
    }

    [Fact]
    public void Add_TextWatermark_CustomProperties()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });

        _handler.Add("/", "watermark", null, new()
        {
            ["text"] = "CONFIDENTIAL",
            ["color"] = "FF0000",
            ["font"] = "Arial",
            ["opacity"] = ".3",
            ["rotation"] = "315"
        });

        var node = _handler.Get("/watermark");
        node.Format["text"].Should().Be("CONFIDENTIAL");
        node.Format["color"].ToString().Should().Contain("FF0000".ToLowerInvariant());
    }

    [Fact]
    public void Add_TextWatermark_PersistsAfterReopen()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
        _handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });

        Reopen();

        var node = _handler.Get("/watermark");
        node.Format["text"].Should().Be("DRAFT");
    }

    [Fact]
    public void Set_WatermarkText_UpdatesExisting()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
        _handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });

        _handler.Set("/watermark", new() { ["text"] = "FINAL" });

        var node = _handler.Get("/watermark");
        node.Format["text"].Should().Be("FINAL");
    }

    [Fact]
    public void Set_WatermarkColor_UpdatesExisting()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
        _handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT", ["color"] = "silver" });

        _handler.Set("/watermark", new() { ["color"] = "ff0000" });

        var node = _handler.Get("/watermark");
        node.Format["color"].ToString().Should().Contain("ff0000");
    }

    [Fact]
    public void Remove_Watermark_RemovesFromDocument()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
        _handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });

        // Verify it exists
        _handler.Get("/watermark").Format.Should().ContainKey("text");

        // Remove
        _handler.Remove("/watermark");

        // Verify it's gone
        var node = _handler.Get("/watermark");
        node.Text.Should().Be("(no watermark)");
    }

    [Fact]
    public void Add_Watermark_ReplacesExisting()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
        _handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });
        _handler.Add("/", "watermark", null, new() { ["text"] = "FINAL" });

        var node = _handler.Get("/watermark");
        node.Format["text"].Should().Be("FINAL");
    }

    [Fact]
    public void Add_Watermark_CreatesThreeHeaders()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });
        _handler.Add("/", "watermark", null, new() { ["text"] = "DRAFT" });

        Reopen();

        // Should have default, first, even headers
        var root = _handler.Get("/");
        var headerCount = root.Children.Count(c => c.Type == "header");
        headerCount.Should().BeGreaterOrEqualTo(3,
            "watermark should create default, first, and even page headers");
    }
}
