using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class ParagraphBorderTest : IDisposable
{
    private readonly string _docxPath;
    private WordHandler _wordHandler;

    public ParagraphBorderTest()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"pbdr_test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_docxPath);
        _wordHandler = new WordHandler(_docxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    private WordHandler Reopen()
    {
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
        return _wordHandler;
    }

    [Fact]
    public void Add_ParagraphWithBottomBorder_GetReturnsBorder()
    {
        // Add a paragraph with a bottom border (simulating header underline)
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Header with underline",
            ["pbdr.bottom"] = "single;12;000000;1"
        });

        // Get and verify
        var node = _wordHandler.Get("/body/p[1]", depth: 0);
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("pbdr.bottom");
        var border = node.Format["pbdr.bottom"].ToString()!;
        border.Should().Contain("single");
        border.Should().Contain("12");
    }

    [Fact]
    public void Set_ParagraphBottomBorder_IsUpdated()
    {
        // Add a plain paragraph
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Test paragraph"
        });

        // Set bottom border
        _wordHandler.Set("/body/p[1]", new() { ["pbdr.bottom"] = "single;4;FF0000" });

        // Get and verify
        var node = _wordHandler.Get("/body/p[1]", depth: 0);
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("pbdr.bottom");
        var border = node.Format["pbdr.bottom"].ToString()!;
        border.Should().Contain("single");
        border.Should().Contain("FF0000");

        // Update to double border
        _wordHandler.Set("/body/p[1]", new() { ["pbdr.bottom"] = "double;8;0000FF" });
        node = _wordHandler.Get("/body/p[1]", depth: 0);
        border = node!.Format["pbdr.bottom"].ToString()!;
        border.Should().Contain("double");
        border.Should().Contain("0000FF");
    }

    [Fact]
    public void Set_ParagraphAllBorders_AllPresent()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Bordered paragraph"
        });

        // Set all borders at once
        _wordHandler.Set("/body/p[1]", new() { ["pbdr.all"] = "single;4;333333" });

        var node = _wordHandler.Get("/body/p[1]", depth: 0);
        node!.Format.Should().ContainKey("pbdr.top");
        node.Format.Should().ContainKey("pbdr.bottom");
        node.Format.Should().ContainKey("pbdr.left");
        node.Format.Should().ContainKey("pbdr.right");
        node.Format.Should().ContainKey("pbdr.between");
    }

    [Fact]
    public void Add_ParagraphBorder_PersistsAfterReopen()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Persistent border",
            ["pbdr.bottom"] = "thick;12;000000;1"
        });

        // Reopen and verify persistence
        var handler = Reopen();
        var node = handler.Get("/body/p[1]", depth: 0);
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("pbdr.bottom");
        var border = node.Format["pbdr.bottom"].ToString()!;
        border.Should().Contain("thick");
    }
}
