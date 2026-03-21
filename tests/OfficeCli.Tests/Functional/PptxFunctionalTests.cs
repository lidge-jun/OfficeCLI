// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for PPTX: each test creates a blank file, adds elements,
/// queries them, and modifies them — exercising the full Create→Add→Get→Set lifecycle.
/// </summary>
public class PptxFunctionalTests : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxFunctionalTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    // Reopen the file to verify persistence
    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Slide lifecycle ====================

    [Fact]
    public void AddSlide_ReturnsPath_Slide1()
    {
        var path = _handler.Add("/", "slide", null, new Dictionary<string, string>());
        path.Should().Be("/slide[1]");
    }

    [Fact]
    public void AddSlide_Get_ReturnsSlideType()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var node = _handler.Get("/slide[1]");
        node.Type.Should().Be("slide");
    }

    [Fact]
    public void AddSlide_Multiple_PathIncrements()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var path3 = _handler.Add("/", "slide", null, new Dictionary<string, string>());
        path3.Should().Be("/slide[3]");
    }

    [Fact]
    public void AddSlide_WithTitle_TitleVisibleInText()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Hello World" });
        var node = _handler.Get("/slide[1]", depth: 2);
        var allText = node.Children.SelectMany(c => c.Children).Select(c => c.Text).Concat(
                      node.Children.Select(c => c.Text))
                      .Where(t => t != null).ToList();
        allText.Any(t => t!.Contains("Hello World")).Should().BeTrue();
    }

    // ==================== Shape lifecycle ====================

    [Fact]
    public void AddShape_ReturnsPath_Shape1()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var shapePath = _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Test" });
        shapePath.Should().Be("/slide[1]/shape[1]");
    }

    [Fact]
    public void AddShape_WithText_TextIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Hello Shape" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Hello Shape");
    }

    [Fact]
    public void AddShape_WithFill_FillIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Filled", ["fill"] = "FF0000" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("fill");
        node.Format["fill"].Should().Be("#FF0000");
    }

    [Fact]
    public void AddShape_WithPosition_PositionIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Positioned",
            ["x"] = "2cm",
            ["y"] = "3cm",
            ["width"] = "5cm",
            ["height"] = "2cm"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["x"].Should().Be("2cm");
        node.Format["y"].Should().Be("3cm");
        node.Format["width"].Should().Be("5cm");
        node.Format["height"].Should().Be("2cm");
    }

    [Fact]
    public void AddShape_WithName_NameIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Named", ["name"] = "MyBox" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["name"].Should().Be("MyBox");
    }

    // ==================== Set: modify shape properties ====================

    [Fact]
    public void SetShape_Bold_BoldIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Normal" });

        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["bold"] = "true" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true);
    }

    [Fact]
    public void SetShape_Italic_ItalicIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Normal" });

        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["italic"] = "true" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("italic");
        node.Format["italic"].Should().Be(true);
    }

    [Fact]
    public void SetShape_Fill_FillIsUpdated()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "A", ["fill"] = "0000FF" });

        // Change fill from blue to red
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["fill"] = "FF0000" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["fill"].Should().Be("#FF0000");
    }

    [Fact]
    public void SetShape_FontSize_SizeIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Big Text" });

        // size property accepts a raw point number (stored as pt*100 internally)
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["size"] = "24" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("size");
        node.Format["size"].Should().Be("24pt");
    }

    [Fact]
    public void SetShape_Position_PositionIsUpdated()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "A" });

        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string>
        {
            ["x"] = "4cm",
            ["y"] = "5cm"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["x"].Should().Be("4cm");
        node.Format["y"].Should().Be("5cm");
    }

    // ==================== Query ====================

    [Fact]
    public void QueryShapes_AfterAddTwo_ReturnsBoth()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "A" });
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "B" });

        var nodes = _handler.Query("shape");
        nodes.Should().HaveCountGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void GetRoot_AfterAddThreeSlides_HasThreeChildren()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>());

        var root = _handler.Get("/");
        root.Children.Should().HaveCount(3);
        root.Children.Should().AllSatisfy(c => c.Type.Should().Be("slide"));
    }

    // ==================== Table lifecycle ====================

    [Fact]
    public void AddTable_ReturnsTablePath()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var path = _handler.Add("/slide[1]", "table", null,
            new Dictionary<string, string> { ["rows"] = "2", ["cols"] = "3" });
        path.Should().Be("/slide[1]/table[1]");
    }

    [Fact]
    public void AddTable_Get_HasCorrectDimensions()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "table", null,
            new Dictionary<string, string> { ["rows"] = "3", ["cols"] = "4" });

        var node = _handler.Get("/slide[1]/table[1]");
        node.Type.Should().Be("table");
        node.Format["rows"].Should().Be(3);
        node.Format["cols"].Should().Be(4);
    }

    [Fact]
    public void SetTableCell_TextIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "table", null,
            new Dictionary<string, string> { ["rows"] = "2", ["cols"] = "2" });

        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]",
            new Dictionary<string, string> { ["text"] = "Cell A1" });

        var table = _handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = table.Children
            .FirstOrDefault(r => r.Type == "tr")
            ?.Children.FirstOrDefault(c => c.Type == "tc");
        cell.Should().NotBeNull();
        cell!.Text.Should().Be("Cell A1");
    }

    // ==================== Table Row Add Lifecycle ====================

    [Fact]
    public void AddRow_FullLifecycle()
    {
        // 1. Create slide + table
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });

        // 2. Add row with cell text
        var path = _handler.Add("/slide[1]/table[1]", "row", null, new() { ["c1"] = "Hello", ["c2"] = "World" });
        path.Should().Be("/slide[1]/table[1]/tr[2]");

        // 3. Get + Verify
        var table = _handler.Get("/slide[1]/table[1]", depth: 2);
        var row2 = table.Children.Where(c => c.Type == "tr").ElementAt(1);
        row2.Children[0].Text.Should().Be("Hello");
        row2.Children[1].Text.Should().Be("World");

        // 4. Set (modify cell text)
        _handler.Set("/slide[1]/table[1]/tr[2]/tc[1]", new() { ["text"] = "Modified" });

        // 5. Get + Verify again
        table = _handler.Get("/slide[1]/table[1]", depth: 2);
        row2 = table.Children.Where(c => c.Type == "tr").ElementAt(1);
        row2.Children[0].Text.Should().Be("Modified");
        row2.Children[1].Text.Should().Be("World");

        // 6. Persistence: Reopen + Verify
        Reopen();
        table = _handler.Get("/slide[1]/table[1]", depth: 2);
        row2 = table.Children.Where(c => c.Type == "tr").ElementAt(1);
        row2.Children[0].Text.Should().Be("Modified");
        row2.Children[1].Text.Should().Be("World");
    }

    [Fact]
    public void AddRow_AtIndex_FullLifecycle()
    {
        // 1. Create slide + table with 2 rows
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "1" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "First" });
        _handler.Set("/slide[1]/table[1]/tr[2]/tc[1]", new() { ["text"] = "Last" });

        // 2. Add row at index 1
        var path = _handler.Add("/slide[1]/table[1]", "row", 1, new() { ["c1"] = "Middle" });
        path.Should().Be("/slide[1]/table[1]/tr[2]");

        // 3. Get + Verify order
        var table = _handler.Get("/slide[1]/table[1]", depth: 2);
        var rows = table.Children.Where(c => c.Type == "tr").ToList();
        rows[0].Children[0].Text.Should().Be("First");
        rows[1].Children[0].Text.Should().Be("Middle");
        rows[2].Children[0].Text.Should().Be("Last");

        // 4. Set (modify inserted row)
        _handler.Set("/slide[1]/table[1]/tr[2]/tc[1]", new() { ["text"] = "Center" });

        // 5. Get + Verify
        table = _handler.Get("/slide[1]/table[1]", depth: 2);
        rows = table.Children.Where(c => c.Type == "tr").ToList();
        rows[1].Children[0].Text.Should().Be("Center");

        // 6. Persistence
        Reopen();
        table = _handler.Get("/slide[1]/table[1]", depth: 2);
        rows = table.Children.Where(c => c.Type == "tr").ToList();
        rows[1].Children[0].Text.Should().Be("Center");
    }

    // ==================== Table Cell Add Lifecycle ====================

    [Fact]
    public void AddCell_FullLifecycle()
    {
        // 1. Create slide + table
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "1" });

        // 2. Add cell
        var path = _handler.Add("/slide[1]/table[1]/tr[1]", "cell", null, new() { ["text"] = "NewCell" });
        path.Should().Be("/slide[1]/table[1]/tr[1]/tc[2]");

        // 3. Get + Verify
        var table = _handler.Get("/slide[1]/table[1]", depth: 2);
        var row = table.Children.First(c => c.Type == "tr");
        row.Children.Should().HaveCount(2);
        row.Children[1].Text.Should().Be("NewCell");

        // 4. Set (modify)
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[2]", new() { ["text"] = "Updated" });

        // 5. Get + Verify
        table = _handler.Get("/slide[1]/table[1]", depth: 2);
        row = table.Children.First(c => c.Type == "tr");
        row.Children[1].Text.Should().Be("Updated");

        // 6. Persistence
        Reopen();
        table = _handler.Get("/slide[1]/table[1]", depth: 2);
        row = table.Children.First(c => c.Type == "tr");
        row.Children[1].Text.Should().Be("Updated");
    }

    [Fact]
    public void AddCell_AtIndex_FullLifecycle()
    {
        // 1. Create slide + table with 2 cells
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "A" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[2]", new() { ["text"] = "C" });

        // 2. Add cell at index 1
        var path = _handler.Add("/slide[1]/table[1]/tr[1]", "cell", 1, new() { ["text"] = "B" });
        path.Should().Be("/slide[1]/table[1]/tr[1]/tc[2]");

        // 3. Get + Verify order
        var table = _handler.Get("/slide[1]/table[1]", depth: 2);
        var row = table.Children.First(c => c.Type == "tr");
        row.Children[0].Text.Should().Be("A");
        row.Children[1].Text.Should().Be("B");
        row.Children[2].Text.Should().Be("C");

        // 4. Set (modify inserted cell)
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[2]", new() { ["text"] = "Beta" });

        // 5. Get + Verify
        table = _handler.Get("/slide[1]/table[1]", depth: 2);
        row = table.Children.First(c => c.Type == "tr");
        row.Children[1].Text.Should().Be("Beta");

        // 6. Persistence
        Reopen();
        table = _handler.Get("/slide[1]/table[1]", depth: 2);
        row = table.Children.First(c => c.Type == "tr");
        row.Children[0].Text.Should().Be("A");
        row.Children[1].Text.Should().Be("Beta");
        row.Children[2].Text.Should().Be("C");
    }

    // ==================== Slide background ====================

    [Fact]
    public void AddSlide_WithBackground_BackgroundIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("#FF0000");
    }

    [Fact]
    public void AddSlide_WithGradientBackground_BackgroundIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000-0000FF" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        var bg = node.Format["background"]?.ToString();
        bg.Should().NotBeNull();
        bg!.Should().Contain("#FF0000");
        bg.Should().Contain("#0000FF");
    }

    [Fact]
    public void SetSlideBackground_SolidColor_IsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "FF0000" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("#FF0000");
    }

    [Fact]
    public void SetSlideBackground_UpdateColor_NewColorIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000" });

        _handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "0000FF" });

        var node = _handler.Get("/slide[1]");
        node.Format["background"].Should().Be("#0000FF");
    }

    [Fact]
    public void SetSlideBackground_None_RemovesBackground()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000" });
        _handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "none" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().NotContainKey("background");
    }

    [Fact]
    public void AddSlide_WithBackground_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "00FF00" });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("#00FF00");
    }

    // ==================== Persistence ====================

    [Fact]
    public void AddShape_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Persistent" });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Persistent");
    }

    // ==================== Speaker Notes ====================

    [Fact]
    public void Notes_Lifecycle()
    {
        // 1. Add slide + notes
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var path = _handler.Add("/slide[1]", "notes", null, new Dictionary<string, string> { ["text"] = "Original note" });
        path.Should().Be("/slide[1]/notes");

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/notes");
        node.Type.Should().Be("notes");
        node.Text.Should().Be("Original note");

        // 3. Query + Verify
        var results = _handler.Query("notes");
        results.Should().Contain(n => n.Type == "notes" && n.Text == "Original note");

        // 4. Set + Verify
        _handler.Set("/slide[1]/notes", new Dictionary<string, string> { ["text"] = "Updated note" });
        node = _handler.Get("/slide[1]/notes");
        node.Text.Should().Be("Updated note");

        // 5. Query reflects update
        results = _handler.Query("notes");
        results.Should().Contain(n => n.Text == "Updated note");
    }

    [Fact]
    public void Notes_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "notes", null, new Dictionary<string, string> { ["text"] = "Persist me" });
        _handler.Set("/slide[1]/notes", new Dictionary<string, string> { ["text"] = "Persisted note" });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]/notes");
        node.Text.Should().Be("Persisted note");
    }

    // ==================== PPTX Hyperlinks ====================

    [Fact]
    public void ShapeLink_Lifecycle()
    {
        // 1. Add slide + shape with link
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Click me",
            ["link"] = "https://first.com"
        });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://first.com");

        // 3. Set new link + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["link"] = "https://updated.com" });
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["link"]).Should().StartWith("https://updated.com");

        // 4. Remove link + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["link"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("link");
    }

    [Fact]
    public void ShapeLink_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Persistent link",
            ["link"] = "https://persist.com"
        });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://persist.com");
    }

    // ==================== PPTX lineDash ====================

    [Fact]
    public void ShapeLineDash_Lifecycle()
    {
        // 1. Add shape with lineDash
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "dashed border",
            ["line"] = "FF0000",
            ["lineDash"] = "dash"
        });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineDash");
        node.Format["lineDash"].Should().Be("dash");

        // 3. Set new dash style + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["lineDash"] = "dot" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["lineDash"].Should().Be("dot");

        // 4. Set solid (remove dash) + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["lineDash"] = "solid" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["lineDash"].Should().Be("solid");
    }

    [Fact]
    public void ShapeLineDash_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "persist dash",
            ["line"] = "0000FF",
            ["lineDash"] = "dashdot"
        });

        Reopen();
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineDash");
        node.Format["lineDash"].Should().Be("dashdot");
    }

    // ==================== PPTX Effects (shadow / glow / reflection) ====================

    [Fact]
    public void ShapeShadow_Lifecycle()
    {
        // 1. Add shape with shadow
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "shadowed",
            ["shadow"] = "000000"
        });

        // 2. Get + Verify — shadow now includes full params: color-blur-angle-dist-opacity
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("shadow");
        node.Format["shadow"]!.ToString()!.Should().StartWith("#000000");

        // 3. Set new shadow color + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["shadow"] = "404040" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["shadow"]!.ToString()!.Should().StartWith("#404040");

        // 4. Remove shadow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["shadow"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("shadow");
    }

    [Fact]
    public void ShapeGlow_Lifecycle()
    {
        // 1. Add shape with glow
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "glowing",
            ["glow"] = "0070FF"
        });

        // 2. Get + Verify — glow now includes full params: color-radius-opacity
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("glow");
        node.Format["glow"]!.ToString()!.Should().StartWith("#0070FF");

        // 3. Set new glow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["glow"] = "FF0000-10" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["glow"]!.ToString()!.Should().StartWith("#FF0000");

        // 4. Remove glow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["glow"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("glow");
    }

    [Fact]
    public void TextShadow_NoFillShape_AppliedToRun()
    {
        // When fill=none, shadow should be applied to text runs (a:rPr/a:effectLst)
        // instead of shape properties, so it renders visually on text.
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Shadow Text",
            ["fill"] = "none",
            ["shadow"] = "333333"
        });

        // Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("shadow");
        node.Format["shadow"]!.ToString()!.Should().StartWith("#333333");

        // Set new shadow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["shadow"] = "FF0000-6-90-4-60" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["shadow"]!.ToString()!.Should().StartWith("#FF0000");
        node.Format["shadow"]!.ToString()!.Should().Contain("6");

        // Remove shadow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["shadow"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("shadow");
    }

    [Fact]
    public void TextGlow_NoFillShape_AppliedToRun()
    {
        // When fill=none, glow should be applied to text runs
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Glow Text",
            ["fill"] = "none",
            ["glow"] = "E94560-8-75"
        });

        // Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("glow");
        node.Format["glow"]!.ToString()!.Should().StartWith("#E94560");

        // Set new glow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["glow"] = "0000FF-12-90" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["glow"]!.ToString()!.Should().StartWith("#0000FF");

        // Remove glow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["glow"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("glow");
    }

    [Fact]
    public void TextShadow_NoFillShape_PersistsAfterReopen()
    {
        // Verify text-level shadow survives save/reopen
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Persistent Shadow",
            ["fill"] = "none",
            ["shadow"] = "222222-5-30-3-50"
        });

        Reopen();

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("shadow");
        node.Format["shadow"]!.ToString()!.Should().StartWith("#222222");
    }

    [Fact]
    public void TextGlow_NoFillShape_PersistsAfterReopen()
    {
        // Verify text-level glow survives save/reopen
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Persistent Glow",
            ["fill"] = "none",
            ["glow"] = "FF6600-10-80"
        });

        Reopen();

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("glow");
        node.Format["glow"]!.ToString()!.Should().StartWith("#FF6600");
    }

    [Fact]
    public void TextFill_NoFillShape_GradientRendered()
    {
        // textFill gradient on fill=none shape must insert gradFill before latin/ea in rPr
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Gradient",
            ["fill"] = "none",
            ["font"] = "Arial",
            ["textFill"] = "FF0000-0000FF-90"
        });

        // Get + Verify textFill is readable
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("textFill");
        node.Format["textFill"]!.ToString()!.Should().Contain("#FF0000");

        // Set new textFill + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["textFill"] = "00FF00-FFFF00-180" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["textFill"]!.ToString()!.Should().Contain("#00FF00");

        // Persistence
        Reopen();
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("textFill");
        node.Format["textFill"]!.ToString()!.Should().Contain("#00FF00");
    }

    /// <summary>
    /// Helper: get the first Drawing.RunProperties from slide 1 shape 1 via reflection.
    /// </summary>
    private DocumentFormat.OpenXml.Drawing.RunProperties GetFirstRunProperties()
    {
        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var shape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().First();
        return shape.Descendants<DocumentFormat.OpenXml.Drawing.RunProperties>().First();
    }

    /// <summary>
    /// Helper: get child element index in parent, or -1 if not found.
    /// </summary>
    private static int ChildIndex<T>(DocumentFormat.OpenXml.OpenXmlElement parent) where T : DocumentFormat.OpenXml.OpenXmlElement
    {
        int i = 0;
        foreach (var child in parent.ChildElements)
        {
            if (child is T) return i;
            i++;
        }
        return -1;
    }

    [Fact]
    public void RunProperties_SolidFillBeforeLatinFont()
    {
        // CT_TextCharacterProperties schema: solidFill must come before latin/ea
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Test",
            ["color"] = "FF0000",
            ["font"] = "Arial"
        });

        var rPr = GetFirstRunProperties();
        var fillIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.SolidFill>(rPr);
        var latinIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.LatinFont>(rPr);
        var eaIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.EastAsianFont>(rPr);

        fillIdx.Should().BeGreaterOrEqualTo(0, "solidFill should exist");
        latinIdx.Should().BeGreaterOrEqualTo(0, "latin font should exist");
        fillIdx.Should().BeLessThan(latinIdx, "solidFill must come before latin in a:rPr schema order");
        fillIdx.Should().BeLessThan(eaIdx, "solidFill must come before ea in a:rPr schema order");
    }

    [Fact]
    public void RunProperties_GradFillBeforeLatinFont()
    {
        // CT_TextCharacterProperties schema: gradFill must come before latin/ea
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Test",
            ["font"] = "Arial",
            ["fill"] = "none",
            ["textFill"] = "FF0000-0000FF-90"
        });

        var rPr = GetFirstRunProperties();
        var fillIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.GradientFill>(rPr);
        var latinIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.LatinFont>(rPr);

        fillIdx.Should().BeGreaterOrEqualTo(0, "gradFill should exist");
        latinIdx.Should().BeGreaterOrEqualTo(0, "latin font should exist");
        fillIdx.Should().BeLessThan(latinIdx, "gradFill must come before latin in a:rPr schema order");
    }

    [Fact]
    public void RunProperties_EffectLstBeforeLatinFont()
    {
        // CT_TextCharacterProperties schema: effectLst must come before latin/ea
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Test",
            ["font"] = "Arial",
            ["fill"] = "none",
            ["shadow"] = "000000"
        });

        var rPr = GetFirstRunProperties();
        var effectIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.EffectList>(rPr);
        var latinIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.LatinFont>(rPr);

        effectIdx.Should().BeGreaterOrEqualTo(0, "effectLst should exist");
        latinIdx.Should().BeGreaterOrEqualTo(0, "latin font should exist");
        effectIdx.Should().BeLessThan(latinIdx, "effectLst must come before latin in a:rPr schema order");
    }

    [Fact]
    public void RunProperties_SchemaOrder_FillBeforeEffectBeforeFont()
    {
        // Full ordering: solidFill < effectLst < latin/ea
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Test",
            ["color"] = "FFFFFF",
            ["font"] = "Arial",
            ["fill"] = "none",
            ["shadow"] = "000000",
            ["glow"] = "FF0000-8-75"
        });

        var rPr = GetFirstRunProperties();
        var fillIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.SolidFill>(rPr);
        var effectIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.EffectList>(rPr);
        var latinIdx = ChildIndex<DocumentFormat.OpenXml.Drawing.LatinFont>(rPr);

        fillIdx.Should().BeGreaterOrEqualTo(0, "solidFill should exist");
        effectIdx.Should().BeGreaterOrEqualTo(0, "effectLst should exist");
        latinIdx.Should().BeGreaterOrEqualTo(0, "latin should exist");

        fillIdx.Should().BeLessThan(effectIdx, "solidFill must come before effectLst");
        effectIdx.Should().BeLessThan(latinIdx, "effectLst must come before latin");
    }

    [Fact]
    public void TextReflection_NoFillShape_AppliedToRun()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Reflected",
            ["fill"] = "none",
            ["reflection"] = "half"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("reflection");
        node.Format["reflection"]!.ToString()!.Should().Be("half");

        // Set + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["reflection"] = "full" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["reflection"]!.ToString()!.Should().Be("full");

        // Remove + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["reflection"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("reflection");
    }

    [Fact]
    public void TextSoftEdge_NoFillShape_AppliedToRun()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Soft",
            ["fill"] = "none",
            ["softEdge"] = "8"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("softEdge");
        node.Format["softEdge"]!.ToString()!.Should().Be("8");

        // Set + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["softEdge"] = "12" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["softEdge"]!.ToString()!.Should().Be("12");

        // Remove + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["softEdge"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("softEdge");
    }

    [Fact]
    public void TextReflection_NoFillShape_PersistsAfterReopen()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Reflected",
            ["fill"] = "none",
            ["reflection"] = "tight"
        });

        Reopen();

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("reflection");
        node.Format["reflection"]!.ToString()!.Should().Be("tight");
    }

    [Fact]
    public void TextSoftEdge_NoFillShape_PersistsAfterReopen()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Soft",
            ["fill"] = "none",
            ["softEdge"] = "5"
        });

        Reopen();

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("softEdge");
        node.Format["softEdge"]!.ToString()!.Should().Be("5");
    }

    [Fact]
    public void ShapeReflection_Lifecycle()
    {
        // 1. Add shape with reflection
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "reflected",
            ["reflection"] = "half"
        });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("reflection");
        // reflection now returns type: "tight", "half", or "full"
        node.Format["reflection"]!.ToString()!.Should().Be("half");

        // 3. Remove reflection + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["reflection"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("reflection");

        // 4. Re-add via Set + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["reflection"] = "full" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("reflection");
    }

    // ==================== PPTX Animation ====================

    [Fact]
    public void ShapeAnimation_Lifecycle()
    {
        // 1. Add shape with animation
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "animated",
            ["animation"] = "fade"
        });

        // 2. Get shape — shape itself is accessible
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("animated");

        // 3. Set different animation + Verify shape still accessible
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["animation"] = "fly" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("animated");

        // 4. Remove animation + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["animation"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("animated");
    }

    [Fact]
    public void ShapeAnimation_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "persist anim",
            ["animation"] = "fade-entrance-500"
        });

        Reopen();
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("persist anim");
    }

    // ==================== Animation Effect Types ====================

    [Fact]
    public void Animation_Fly_UsesPropertyAnim()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "fly test",
            ["animation"] = "fly-entrance-500"
        });

        // Verify shape is accessible and animation readback works
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("fly test");
        node.Format.Should().ContainKey("animation");
        node.Format["animation"]!.ToString()!.Should().Contain("fly");

        // Verify via raw XML: should have p:anim (not p:animEffect filter="fly")
        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var xml = slidePart.Slide.OuterXml;
        xml.Should().Contain("ppt_"); // p:anim uses ppt_x or ppt_y
        xml.Should().NotContain("filter=\"fly\""); // should NOT use invalid filter
    }

    [Fact]
    public void Animation_Zoom_UsesAnimScale()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "zoom test",
            ["animation"] = "zoom-entrance-600"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("animation");
        node.Format["animation"]!.ToString()!.Should().Contain("zoom");

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var xml = slidePart.Slide.OuterXml;
        xml.Should().Contain("animScale"); // p:animScale for zoom
        xml.Should().NotContain("filter=\"zoom\"");
    }

    [Fact]
    public void Animation_Swivel_UsesAnimRot()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "swivel test",
            ["animation"] = "swivel-entrance-700"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("animation");
        node.Format["animation"]!.ToString()!.Should().Contain("swivel");

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var xml = slidePart.Slide.OuterXml;
        xml.Should().Contain("animRot"); // p:animRot for swivel
        xml.Should().NotContain("filter=\"swivel\"");
    }

    // ==================== Morph + AdvanceTime ====================

    [Fact]
    public void Morph_WithAdvanceTime_NoSchemaDuplicate()
    {
        // Morph transition + advanceTime should produce a single transition inside mc:AlternateContent
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>
        {
            ["transition"] = "morph",
            ["advanceTime"] = "2000",
            ["advanceClick"] = "false"
        });

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.Skip(1).First();
        var slide = slidePart.Slide;

        // Should have AlternateContent with morph
        slide.OuterXml.Should().Contain("morph");

        // Should NOT have a standalone p:transition outside AlternateContent
        var typedTransitions = slide.Elements<DocumentFormat.OpenXml.Presentation.Transition>().Count();
        typedTransitions.Should().Be(0, "advanceTime should be merged into morph AlternateContent, not a separate p:transition");

        // advTm should be inside the AlternateContent transitions
        var acXml = slide.ChildElements.First(c => c.LocalName == "AlternateContent").OuterXml;
        acXml.Should().Contain("advTm=\"2000\"");
        acXml.Should().Contain("advClick=\"0\"");
    }

    [Fact]
    public void Morph_SetAdvanceTime_AfterCreation()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["transition"] = "morph" });

        // Set advanceTime after slide creation
        _handler.Set("/slide[2]", new Dictionary<string, string>
        {
            ["advanceTime"] = "1500",
            ["advanceClick"] = "false"
        });

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.Skip(1).First();
        var slide = slidePart.Slide;

        var typedTransitions = slide.Elements<DocumentFormat.OpenXml.Presentation.Transition>().Count();
        typedTransitions.Should().Be(0, "should not have duplicate typed p:transition");

        var acXml = slide.ChildElements.First(c => c.LocalName == "AlternateContent").OuterXml;
        acXml.Should().Contain("advTm=\"1500\"");
    }

    // ==================== Radial Gradient ====================

    [Fact]
    public void RadialGradient_Shape_Lifecycle()
    {
        // 1. Add slide + shape
        _handler.Add("/", "slide", null, new() { ["title"] = "Gradient Test" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Radial" });

        // 2. Set radial gradient (blue→purple, top-right focus)
        _handler.Set("/slide[1]/shape[1]", new() { ["gradient"] = "radial:1E90FF-4B0082-tr" });

        // 3. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("gradient");
        var grad = (string)node.Format["gradient"];
        grad.Should().StartWith("radial:");
        grad.Should().Contain("#1E90FF");
        grad.Should().Contain("#4B0082");
        grad.Should().EndWith("tr");

        // 4. Change to center focus
        _handler.Set("/slide[1]/shape[1]", new() { ["gradient"] = "radial:FF0000-FFFF00-center" });
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["gradient"]).Should().Contain("center");

        // 5. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["gradient"]).Should().StartWith("radial:");
    }

    [Fact]
    public void RadialGradient_Background_Lifecycle()
    {
        // 1. Add slide
        _handler.Add("/", "slide", null, new() { ["title"] = "BG Gradient" });

        // 2. Set radial gradient as background
        _handler.Set("/slide[1]", new() { ["background"] = "radial:4B0082-1E90FF-bl" });

        // 3. Get + Verify
        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        var bg = (string)node.Format["background"];
        bg.Should().StartWith("radial:");
        bg.Should().Contain("bl");

        // 4. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]");
        ((string)node.Format["background"]).Should().StartWith("radial:");
    }

    // ==================== Line Opacity ====================

    [Fact]
    public void LineOpacity_Lifecycle()
    {
        // 1. Add slide + shape with line
        _handler.Add("/", "slide", null, new() { ["title"] = "Line Test" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Bordered" });

        // 2. Set line color + opacity
        _handler.Set("/slide[1]/shape[1]", new()
        {
            ["line"] = "FFFFFF",
            ["linewidth"] = "2pt",
            ["lineopacity"] = "0.5"
        });

        // 3. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("line");
        ((string)node.Format["line"]).Should().Be("#FFFFFF");
        node.Format.Should().ContainKey("lineOpacity");
        ((string)node.Format["lineOpacity"]).Should().Be("0.5");

        // 4. Change opacity
        _handler.Set("/slide[1]/shape[1]", new() { ["lineopacity"] = "0.3" });
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["lineOpacity"]).Should().Be("0.3");

        // 5. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["lineOpacity"]).Should().Be("0.3");
    }

    [Fact]
    public void LineWidth_BareNumber_TreatedAsPoints()
    {
        // Bug: lineWidth=3 was treated as 3 EMU (≈0), should be 3pt (38100 EMU)
        // Apache POI's setLineWidth() accepts points for bare numbers
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test", ["line"] = "FF0000", ["lineWidth"] = "3"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        // 3pt = 38100 EMU = 0.11cm (approx)
        var lw = (string)node.Format["lineWidth"];
        // Should NOT be "0cm" (which would mean 3 EMU)
        lw.Should().NotBe("0cm", "bare number lineWidth should be treated as points, not EMU");

        // Verify 3pt with explicit suffix gives same result
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test2", ["line"] = "FF0000", ["lineWidth"] = "3pt"
        });
        var node2 = _handler.Get("/slide[1]/shape[2]");
        ((string)node2.Format["lineWidth"]).Should().Be(lw,
            "lineWidth=3 and lineWidth=3pt should produce identical results");
    }

    [Fact]
    public void CustomGeometry_CoordinatesScaledTo100000()
    {
        // Bug: path w/h used raw user coordinates (e.g. w=100), which is too small for PowerPoint
        // to render. OOXML standard uses 0-100000 coordinate space for custom geometry.
        // Fix: multiply user coordinates by 1000 internally (0-100 → 0-100000).
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["fill"] = "FF0000",
            ["geometry"] = "M 50,0 L 100,100 L 0,100 Z" // triangle in 0-100 space
        });

        // Read raw XML to check path dimensions
        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var shape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().First();
        var custGeom = shape.ShapeProperties!.GetFirstChild<DocumentFormat.OpenXml.Drawing.CustomGeometry>()!;
        var path = custGeom.PathList!.GetFirstChild<DocumentFormat.OpenXml.Drawing.Path>()!;

        // Path w/h should be scaled to OOXML standard range (×1000)
        path.Width!.Value.Should().Be(100000L,
            "path width should be scaled from 100 to 100000 for OOXML compatibility");
        path.Height!.Value.Should().Be(100000L,
            "path height should be scaled from 100 to 100000 for OOXML compatibility");

        // Point coordinates should also be scaled
        var moveTo = path.GetFirstChild<DocumentFormat.OpenXml.Drawing.MoveTo>()!;
        var pt = moveTo.GetFirstChild<DocumentFormat.OpenXml.Drawing.Point>()!;
        pt.X!.Value.Should().Be("50000", "x=50 should be scaled to 50000");
        pt.Y!.Value.Should().Be("0", "y=0 should remain 0");
    }

    [Fact]
    public void CustomGeometry_InsertedAfterXfrm_BeforeFill()
    {
        // Bug: custGeom was appended after solidFill/ln, PowerPoint ignores it.
        // OOXML requires: xfrm → custGeom → fill → ln
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["fill"] = "FF0000",
            ["line"] = "0000FF",
            ["geometry"] = "M 0,0 L 100,0 L 100,100 L 0,100 Z"
        });

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var shape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().First();
        var spPr = shape.ShapeProperties!;

        // Verify element order: xfrm must come before custGeom, custGeom before fill
        var children = spPr.ChildElements.ToList();
        var xfrmIdx = children.FindIndex(c => c is DocumentFormat.OpenXml.Drawing.Transform2D);
        var custGeomIdx = children.FindIndex(c => c is DocumentFormat.OpenXml.Drawing.CustomGeometry);
        var fillIdx = children.FindIndex(c => c is DocumentFormat.OpenXml.Drawing.SolidFill);
        var lnIdx = children.FindIndex(c => c is DocumentFormat.OpenXml.Drawing.Outline);

        custGeomIdx.Should().BeGreaterThan(xfrmIdx,
            "custGeom must come after xfrm in OOXML element order");
        custGeomIdx.Should().BeLessThan(fillIdx,
            "custGeom must come before solidFill in OOXML element order");
        custGeomIdx.Should().BeLessThan(lnIdx,
            "custGeom must come before ln in OOXML element order");
    }

    // ==================== Shape Image Fill ====================

    [Fact]
    public void ShapeImageFill_Lifecycle()
    {
        // 1. Create a tiny test image
        var imgPath = Path.Combine(Path.GetTempPath(), $"test_img_{Guid.NewGuid():N}.png");
        try
        {
            // Write a minimal 1x1 PNG
            var pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A, // PNG signature
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,  // IHDR chunk
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,0xDE,
                0x00,0x00,0x00,0x0C,0x49,0x44,0x41,0x54,  // IDAT chunk
                0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,0x00,0x00,0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,
                0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82 // IEND chunk
            };
            File.WriteAllBytes(imgPath, pngBytes);

            // 2. Add slide + shape
            _handler.Add("/", "slide", null, new() { ["title"] = "Image Fill" });
            _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Filled" });

            // 3. Set image fill
            _handler.Set("/slide[1]/shape[1]", new() { ["image"] = imgPath });

            // 4. Get + Verify
            var node = _handler.Get("/slide[1]/shape[1]");
            node.Format.Should().ContainKey("image");
            ((string)node.Format["image"]).Should().Be("true");

            // 5. Persist + Verify
            Reopen();
            node = _handler.Get("/slide[1]/shape[1]");
            node.Format.Should().ContainKey("image");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // ==================== Slide Layout ====================

    [Fact]
    public void SlideLayout_Default_HasLayoutInfo()
    {
        // 1. Add slide without specifying layout
        _handler.Add("/", "slide", null, new() { ["title"] = "Default Layout" });

        // 2. Get + Verify layout info is returned
        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("layout");
        ((string)node.Format["layout"]).Should().NotBeNullOrEmpty();

        // 3. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("layout");
    }

    [Fact]
    public void SlideLayout_ByIndex_SelectsDifferentLayout()
    {
        // 1. Add two slides with different layout indices
        _handler.Add("/", "slide", null, new() { ["title"] = "Layout 1", ["layout"] = "1" });
        _handler.Add("/", "slide", null, new() { ["title"] = "Layout 2", ["layout"] = "2" });

        // 2. Get + Verify they have different layouts
        var node1 = _handler.Get("/slide[1]");
        var node2 = _handler.Get("/slide[2]");
        node1.Format.Should().ContainKey("layout");
        node2.Format.Should().ContainKey("layout");

        // Layout names should be different (assuming blank doc has >1 layout)
        var layout1 = (string)node1.Format["layout"];
        var layout2 = (string)node2.Format["layout"];
        layout1.Should().NotBeNullOrEmpty();
        layout2.Should().NotBeNullOrEmpty();
        layout1.Should().NotBe(layout2, "different layout indices should yield different layouts");

        // 3. Persist + Verify
        Reopen();
        node1 = _handler.Get("/slide[1]");
        node2 = _handler.Get("/slide[2]");
        ((string)node1.Format["layout"]).Should().Be(layout1);
        ((string)node2.Format["layout"]).Should().Be(layout2);
    }

    [Fact]
    public void SlideLayout_ByType_Blank()
    {
        // 1. Add slide with blank layout type
        _handler.Add("/", "slide", null, new() { ["layout"] = "blank" });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("layoutType");
        ((string)node.Format["layoutType"]).Should().Be("blank");

        // 3. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]");
        ((string)node.Format["layoutType"]).Should().Be("blank");
    }

    [Fact]
    public void SlideLayout_ByName_MatchesLayoutName()
    {
        // 1. Get the name of layout #1 to use as a name lookup
        _handler.Add("/", "slide", null, new() { ["layout"] = "1" });
        var node = _handler.Get("/slide[1]");
        var layoutName = (string)node.Format["layout"];

        // 2. Add another slide using that layout name
        _handler.Add("/", "slide", null, new() { ["title"] = "By Name", ["layout"] = layoutName });
        var node2 = _handler.Get("/slide[2]");
        ((string)node2.Format["layout"]).Should().Be(layoutName);
    }

    [Fact]
    public void SlideLayout_RootGet_ShowsLayoutPerSlide()
    {
        // 1. Add slides with different layouts
        _handler.Add("/", "slide", null, new() { ["title"] = "Slide A", ["layout"] = "1" });
        _handler.Add("/", "slide", null, new() { ["title"] = "Slide B", ["layout"] = "blank" });

        // 2. Get root
        var root = _handler.Get("/");
        root.Children.Should().HaveCount(2);
        root.Children[0].Format.Should().ContainKey("layout");
        root.Children[1].Format.Should().ContainKey("layout");
    }

    // ==================== Charts ====================

    [Fact]
    public void Chart_Column_Lifecycle()
    {
        // 1. Add slide + column chart
        _handler.Add("/", "slide", null, new() { ["title"] = "Chart Test" });
        var path = _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["title"] = "Q1 Sales",
            ["categories"] = "Jan,Feb,Mar",
            ["series1"] = "Revenue:100,200,300",
            ["series2"] = "Cost:80,150,250"
        });
        path.Should().Be("/slide[1]/chart[1]");

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
        node.Format.Should().ContainKey("chartType");
        ((string)node.Format["chartType"]).Should().Be("column");
        ((string)node.Format["title"]).Should().Be("Q1 Sales");
        ((int)node.Format["seriesCount"]).Should().Be(2);
        ((string)node.Format["categories"]).Should().Be("Jan,Feb,Mar");

        // 3. Set — change title
        _handler.Set("/slide[1]/chart[1]", new() { ["title"] = "Updated Sales" });

        // 4. Get + Verify title changed
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["title"]).Should().Be("Updated Sales");

        // 5. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("column");
        ((string)node.Format["title"]).Should().Be("Updated Sales");
        ((int)node.Format["seriesCount"]).Should().Be(2);
    }

    [Fact]
    public void Chart_Bar_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Bar Chart" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "bar",
            ["title"] = "Horizontal",
            ["categories"] = "A,B,C",
            ["series1"] = "Data:10,20,30"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("bar");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("bar");
    }

    [Fact]
    public void Chart_Line_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Line Chart" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "line",
            ["title"] = "Trend",
            ["categories"] = "Q1,Q2,Q3,Q4",
            ["data"] = "Sales:10,25,30,45;Profit:5,12,18,30"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("line");
        ((int)node.Format["seriesCount"]).Should().Be(2);

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("line");
    }

    [Fact]
    public void Chart_Pie_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Pie Chart" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "pie",
            ["title"] = "Market Share",
            ["categories"] = "Apple,Google,Microsoft",
            ["series1"] = "Share:40,30,30"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("pie");
        ((string)node.Format["categories"]).Should().Be("Apple,Google,Microsoft");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("pie");
    }

    [Fact]
    public void Chart_SeriesData_ReadbackAtDepth()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Data Test" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["categories"] = "A,B",
            ["series1"] = "S1:10,20",
            ["series2"] = "S2:30,40"
        });

        // depth=0: just seriesCount
        var node0 = _handler.Get("/slide[1]/chart[1]", 0);
        ((int)node0.Format["seriesCount"]).Should().Be(2);
        node0.Children.Should().BeEmpty();

        // depth=1: series children with values
        var node1 = _handler.Get("/slide[1]/chart[1]", 1);
        node1.Children.Should().HaveCount(2);
        node1.Children[0].Type.Should().Be("series");
        node1.Children[0].Text.Should().Be("S1");
        ((string)node1.Children[0].Format["values"]).Should().Be("10,20");
        node1.Children[1].Text.Should().Be("S2");
        ((string)node1.Children[1].Format["values"]).Should().Be("30,40");
    }

    [Fact]
    public void Chart_Query_FindsCharts()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["title"] = "Revenue",
            ["series1"] = "Data:1,2,3"
        });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Not a chart" });

        var charts = _handler.Query("chart");
        charts.Should().HaveCount(1);
        charts[0].Type.Should().Be("chart");
        ((string)charts[0].Format["title"]).Should().Be("Revenue");
    }

    [Fact]
    public void Chart_SetLegend_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Legend Test" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["series1"] = "A:1,2",
            ["legend"] = "top"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("legend");
        ((string)node.Format["legend"]).Should().Be("t");

        // Change legend to none
        _handler.Set("/slide[1]/chart[1]", new() { ["legend"] = "none" });
        node = _handler.Get("/slide[1]/chart[1]");
        node.Format.Should().NotContainKey("legend");

        // Set legend back
        _handler.Set("/slide[1]/chart[1]", new() { ["legend"] = "right" });
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["legend"]).Should().Be("r");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["legend"]).Should().Be("r");
    }

    [Fact]
    public void Chart_Doughnut_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Donut" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "doughnut",
            ["title"] = "Budget",
            ["categories"] = "Rent,Food,Transport",
            ["series1"] = "Spending:1200,800,400"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("doughnut");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("doughnut");
    }

    [Fact]
    public void Chart_SlideChildNodes_IncludesChart()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Mixed" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "line",
            ["series1"] = "Data:1,2,3"
        });

        var slide = _handler.Get("/slide[1]");
        slide.Children.Should().Contain(c => c.Type == "chart");
        slide.Children.Should().Contain(c => c.Type == "textbox" || c.Type == "title");
    }

    // ==================== Theme Colors ====================

    [Fact]
    public void ThemeColor_Fill_Lifecycle()
    {
        // 1. Add slide + shape with theme color fill
        _handler.Add("/", "slide", null, new() { ["title"] = "Theme Test" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Accent", ["fill"] = "accent1" });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["fill"]).Should().Be("accent1");

        // 3. Set to different theme color
        _handler.Set("/slide[1]/shape[2]", new() { ["fill"] = "accent3" });
        node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["fill"]).Should().Be("accent3");

        // 4. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["fill"]).Should().Be("accent3");
    }

    [Fact]
    public void ThemeColor_TextAndLine_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Theme" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Styled" });

        // Set theme colors for text and line
        _handler.Set("/slide[1]/shape[2]", new()
        {
            ["color"] = "dk1",
            ["line"] = "accent2"
        });

        var node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["color"]).Should().Be("dk1");
        ((string)node.Format["line"]).Should().Be("accent2");

        Reopen();
        node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["color"]).Should().Be("dk1");
    }

    // ==================== Connectors ====================

    [Fact]
    public void Connector_Lifecycle()
    {
        // 1. Add slide + connector
        _handler.Add("/", "slide", null, new() { ["title"] = "Flow" });
        var path = _handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "2cm",
            ["y"] = "5cm",
            ["width"] = "6cm",
            ["height"] = "0cm",
            ["line"] = "000000",
            ["linewidth"] = "2pt"
        });
        path.Should().StartWith("/slide[1]/connector[");

        // 2. Persist + Verify (connector should survive reopen)
        Reopen();
        // Connectors are GraphicFrame-less, they're direct children
        var slide = _handler.Get("/slide[1]");
        slide.Should().NotBeNull();
    }

    [Fact]
    public void Connector_Elbow_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Elbow" });
        _handler.Add("/slide[1]", "connector", null, new()
        {
            ["preset"] = "elbow",
            ["x"] = "1cm",
            ["y"] = "2cm",
            ["width"] = "5cm",
            ["height"] = "3cm",
            ["line"] = "accent1"
        });

        Reopen();
        var slide = _handler.Get("/slide[1]");
        slide.Should().NotBeNull();
    }

    // ==================== Group Shapes ====================

    [Fact]
    public void Group_Lifecycle()
    {
        // 1. Add slide + 3 shapes
        _handler.Add("/", "slide", null, new() { ["title"] = "Group Test" });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "B", ["x"] = "5cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });

        // Shapes 2 and 3 (title is shape 1)
        var slide = _handler.Get("/slide[1]");
        var shapesBeforeGroup = slide.Children.Count(c => c.Type == "textbox" || c.Type == "title");
        shapesBeforeGroup.Should().BeGreaterThanOrEqualTo(3);

        // 2. Group shapes 2 and 3
        var groupPath = _handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "2,3" });
        groupPath.Should().StartWith("/slide[1]/group[");

        // 3. Persist + Verify
        Reopen();
        var root = _handler.Get("/slide[1]");
        root.Should().NotBeNull();
    }

    // ==================== Chart Data Modification ====================

    [Fact]
    public void Chart_SetData_Lifecycle()
    {
        // 1. Add chart
        _handler.Add("/", "slide", null, new() { ["title"] = "Data Mod" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["categories"] = "A,B",
            ["series1"] = "S1:10,20",
            ["series2"] = "S2:30,40"
        });

        // 2. Verify original data
        var node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["values"]).Should().Be("10,20");

        // 3. Set — modify series1 data
        _handler.Set("/slide[1]/chart[1]", new() { ["series1"] = "S1:100,200" });

        // 4. Get + Verify
        node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["values"]).Should().Be("100,200");
        node.Children[0].Text.Should().Be("S1");

        // 5. Set — modify all data at once via data property
        _handler.Set("/slide[1]/chart[1]", new() { ["data"] = "X:1,2;Y:3,4" });
        node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["values"]).Should().Be("1,2");
        ((string)node.Children[1].Format["values"]).Should().Be("3,4");

        // 6. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["values"]).Should().Be("1,2");
    }

    [Fact]
    public void Chart_SetCategories_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Cat Mod" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "D:1,2,3"
        });

        // Change categories
        _handler.Set("/slide[1]/chart[1]", new() { ["categories"] = "X,Y,Z" });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["categories"]).Should().Be("X,Y,Z");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["categories"]).Should().Be("X,Y,Z");
    }

    // ==================== Slide Size ====================

    [Fact]
    public void SlideSize_GetDefault()
    {
        // Blank doc should have default slide size
        var root = _handler.Get("/");
        root.Format.Should().ContainKey("slideWidth");
        root.Format.Should().ContainKey("slideHeight");
    }

    [Fact]
    public void SlideSize_SetPreset_Lifecycle()
    {
        // 1. Set to 4:3
        _handler.Set("/", new() { ["slidesize"] = "4:3" });

        // 2. Get + Verify
        var root = _handler.Get("/");
        ((string)root.Format["slideSize"]).Should().Be("screen4x3");

        // 3. Set to 16:9
        _handler.Set("/", new() { ["slidesize"] = "16:9" });
        root = _handler.Get("/");
        ((string)root.Format["slideSize"]).Should().Be("screen16x9");

        // 4. Persist + Verify
        Reopen();
        root = _handler.Get("/");
        ((string)root.Format["slideSize"]).Should().Be("screen16x9");
    }

    [Fact]
    public void SlideSize_SetCustom_Lifecycle()
    {
        // Set custom dimensions
        _handler.Set("/", new() { ["slidewidth"] = "30cm", ["slideheight"] = "20cm" });

        var root = _handler.Get("/");
        ((string)root.Format["slideWidth"]).Should().Contain("cm");
        ((string)root.Format["slideSize"]).Should().Be("custom");

        Reopen();
        root = _handler.Get("/");
        ((string)root.Format["slideSize"]).Should().Be("custom");
    }

    // ==================== Chart Formatting ====================

    [Fact]
    public void Chart_SeriesColors_Lifecycle()
    {
        // 1. Add chart with custom colors
        _handler.Add("/", "slide", null, new() { ["title"] = "Colors" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["categories"] = "A,B",
            ["series1"] = "S1:10,20",
            ["series2"] = "S2:30,40",
            ["colors"] = "FF0000,00FF00"
        });

        // 2. Get + Verify series colors at depth 1
        var node = _handler.Get("/slide[1]/chart[1]", 1);
        node.Children.Should().HaveCount(2);
        ((string)node.Children[0].Format["color"]).Should().Be("#FF0000");
        ((string)node.Children[1].Format["color"]).Should().Be("#00FF00");

        // 3. Set — change colors
        _handler.Set("/slide[1]/chart[1]", new() { ["colors"] = "0000FF,FFFF00" });
        node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["color"]).Should().Be("#0000FF");

        // 4. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["color"]).Should().Be("#0000FF");
    }

    [Fact]
    public void Chart_DataLabels_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Labels" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["series1"] = "S1:10,20,30"
        });

        // 1. Set data labels
        _handler.Set("/slide[1]/chart[1]", new() { ["datalabels"] = "value" });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("dataLabels");
        ((string)node.Format["dataLabels"]).Should().Contain("value");

        // 3. Set to none
        _handler.Set("/slide[1]/chart[1]", new() { ["datalabels"] = "none" });
        node = _handler.Get("/slide[1]/chart[1]");
        node.Format.Should().NotContainKey("dataLabels");

        // 4. Set multiple
        _handler.Set("/slide[1]/chart[1]", new() { ["datalabels"] = "value,percent" });
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["dataLabels"]).Should().Contain("value");
        ((string)node.Format["dataLabels"]).Should().Contain("percent");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["dataLabels"]).Should().Contain("value");
    }

    [Fact]
    public void Chart_AxisTitle_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Axis" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["series1"] = "Revenue:100,200"
        });

        // Set axis titles
        _handler.Set("/slide[1]/chart[1]", new() { ["axistitle"] = "Amount ($)" });

        // Persist + Verify (axis title is in raw XML, verify no crash)
        Reopen();
        var node = _handler.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("chartType");
    }

    // ==================== Combo Chart ====================

    [Fact]
    public void Chart_Combo_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Combo" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "combo",
            ["categories"] = "Q1,Q2,Q3,Q4",
            ["series1"] = "Revenue:100,200,300,400",
            ["series2"] = "Trend:150,200,250,350",
            ["combosplit"] = "1"
        });

        // Get + Verify combo type detected
        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("combo");
        ((int)node.Format["seriesCount"]).Should().Be(2);

        // Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("combo");
    }

    // ==================== Table Style ====================

    [Fact]
    public void TableStyle_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Table Style" });
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // Set table style
        _handler.Set("/slide[1]/table[1]", new() { ["tablestyle"] = "medium2" });

        // Persist + Verify
        Reopen();
        var node = _handler.Get("/slide[1]/table[1]");
        node.Type.Should().Be("table");
    }

    // ==================== Picture Cropping ====================

    [Fact]
    public void PictureCrop_Lifecycle()
    {
        // Create test image
        var imgPath = Path.Combine(Path.GetTempPath(), $"test_crop_{Guid.NewGuid():N}.png");
        try
        {
            var pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,0xDE,
                0x00,0x00,0x00,0x0C,0x49,0x44,0x41,0x54,
                0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,0x00,0x00,0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,
                0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
            };
            File.WriteAllBytes(imgPath, pngBytes);

            _handler.Add("/", "slide", null, new() { ["title"] = "Crop Test" });
            _handler.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });

            // Set crop (10% from each side)
            _handler.Set("/slide[1]/picture[1]", new() { ["crop"] = "10,10,10,10" });

            // Get + Verify
            var node = _handler.Get("/slide[1]/picture[1]");
            node.Format.Should().ContainKey("crop");
            ((string)node.Format["crop"]).Should().Be("10,10,10,10");

            // Set individual crop
            _handler.Set("/slide[1]/picture[1]", new() { ["cropleft"] = "20" });
            node = _handler.Get("/slide[1]/picture[1]");
            ((string)node.Format["crop"]).Should().StartWith("20,");

            // Persist + Verify
            Reopen();
            node = _handler.Get("/slide[1]/picture[1]");
            node.Format.Should().ContainKey("crop");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // ==================== Ungroup ====================

    [Fact]
    public void Ungroup_Lifecycle()
    {
        // 1. Add slide + shapes, group them
        _handler.Add("/", "slide", null, new() { ["title"] = "Ungroup" });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "B", ["x"] = "5cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        _handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "2,3" });

        // 2. Remove (ungroup) the group
        _handler.Remove("/slide[1]/group[1]");

        // 3. Shapes should be back as individual shapes
        Reopen();
        var slide = _handler.Get("/slide[1]");
        // Title + 2 ungrouped shapes should exist
        slide.Children.Count(c => c.Type == "textbox" || c.Type == "title").Should().BeGreaterThanOrEqualTo(3);
    }

    // ==================== WordArt ====================

    [Fact]
    public void WordArt_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "WordArt" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Wavy Text" });

        // 1. Set text warp
        _handler.Set("/slide[1]/shape[2]", new() { ["textwarp"] = "textWave1" });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("textWarp");
        ((string)node.Format["textWarp"]).Should().Be("textWave1");

        // 3. Remove warp
        _handler.Set("/slide[1]/shape[2]", new() { ["textwarp"] = "none" });
        node = _handler.Get("/slide[1]/shape[2]");
        node.Format.Should().NotContainKey("textWarp");

        // 4. Set again + Persist
        _handler.Set("/slide[1]/shape[2]", new() { ["textwarp"] = "textChevron" });
        Reopen();
        node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["textWarp"]).Should().Be("textChevron");
    }

    // ==================== Picture Replace ====================

    [Fact]
    public void PictureReplace_Lifecycle()
    {
        var img1 = Path.Combine(Path.GetTempPath(), $"test_r1_{Guid.NewGuid():N}.png");
        var img2 = Path.Combine(Path.GetTempPath(), $"test_r2_{Guid.NewGuid():N}.png");
        try
        {
            var png = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,0xDE,
                0x00,0x00,0x00,0x0C,0x49,0x44,0x41,0x54,
                0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,0x00,0x00,0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,
                0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
            };
            File.WriteAllBytes(img1, png);
            File.WriteAllBytes(img2, png);

            _handler.Add("/", "slide", null, new() { ["title"] = "Replace" });
            _handler.Add("/slide[1]", "picture", null, new() { ["path"] = img1 });

            // Replace image
            _handler.Set("/slide[1]/picture[1]", new() { ["path"] = img2 });

            // Should not throw
            Reopen();
            var node = _handler.Get("/slide[1]/picture[1]");
            node.Type.Should().Be("picture");
        }
        finally
        {
            if (File.Exists(img1)) File.Delete(img1);
            if (File.Exists(img2)) File.Delete(img2);
        }
    }

    // ==================== Chart Axis Formatting ====================

    [Fact]
    public void Chart_AxisFormatting_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Axis" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["series1"] = "Data:10,50,100"
        });

        // Set axis min/max/unit
        _handler.Set("/slide[1]/chart[1]", new()
        {
            ["axismin"] = "0",
            ["axismax"] = "150",
            ["majorunit"] = "25",
            ["axisnumfmt"] = "0.0"
        });

        // Persist + Verify no crash
        Reopen();
        var node = _handler.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("chartType");
    }

    // ==================== Master Editing ====================

    [Fact]
    public void MasterEdit_LayoutName()
    {
        // Set layout name
        _handler.Set("/slideLayout[1]", new() { ["name"] = "My Custom Blank" });

        // After reopen, the layout name should persist
        Reopen();
        _handler.Add("/", "slide", null, new() { ["layout"] = "My Custom Blank" });
        var node = _handler.Get("/slide[1]");
        ((string)node.Format["layout"]).Should().Be("My Custom Blank");
    }

    // ==================== Video/Audio ====================

    [Fact]
    public void Media_Video_Lifecycle()
    {
        var videoPath = Path.Combine(Path.GetTempPath(), $"test_vid_{Guid.NewGuid():N}.mp4");
        try
        {
            File.WriteAllBytes(videoPath, new byte[] { 0x00, 0x00, 0x00, 0x20 });

            // 1. Add video
            _handler.Add("/", "slide", null, new() { ["title"] = "Video" });
            _handler.Add("/slide[1]", "video", null, new()
            {
                ["path"] = videoPath,
                ["width"] = "10cm",
                ["height"] = "6cm",
                ["volume"] = "60",
                ["autoplay"] = "true"
            });

            // 2. Get — should show as "video" type
            var slide = _handler.Get("/slide[1]");
            slide.Children.Should().Contain(c => c.Type == "video");
            var videoNode = slide.Children.First(c => c.Type == "video");
            videoNode.Format.Should().ContainKey("volume");
            ((int)videoNode.Format["volume"]).Should().Be(60);
            videoNode.Format.Should().ContainKey("autoplay");

            // 3. Set — change volume
            _handler.Set("/slide[1]/video[1]", new() { ["volume"] = "40" });
            slide = _handler.Get("/slide[1]");
            videoNode = slide.Children.First(c => c.Type == "video");
            ((int)videoNode.Format["volume"]).Should().Be(40);

            // 4. Query — find videos
            var videos = _handler.Query("video");
            videos.Should().HaveCount(1);
            videos[0].Type.Should().Be("video");

            // 5. Persist + Verify
            Reopen();
            slide = _handler.Get("/slide[1]");
            slide.Children.Should().Contain(c => c.Type == "video");

            // 6. Remove video
            _handler.Remove("/slide[1]/video[1]");
            slide = _handler.Get("/slide[1]");
            slide.Children.Should().NotContain(c => c.Type == "video");
        }
        finally
        {
            if (File.Exists(videoPath)) File.Delete(videoPath);
        }
    }

    // ==================== Hyperlink Remove ====================

    [Fact]
    public void ShapeHyperlink_Remove()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Link" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Click me" });

        // Add link
        _handler.Set("/slide[1]/shape[2]", new() { ["link"] = "https://example.com" });
        var node = _handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("link");

        // Remove link
        _handler.Set("/slide[1]/shape[2]", new() { ["link"] = "none" });
        node = _handler.Get("/slide[1]/shape[2]");
        node.Format.Should().NotContainKey("link");
    }

    // ==================== Remove shape with animation cleanup ====================

    [Fact]
    public void RemoveShape_WithAnimation_AnimationIsCleanedUp()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Animated", ["fill"] = "FF0000" });
        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "fade-entrance-500" });

        // Verify animation exists
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("animation");

        // Remove the shape
        _handler.Remove("/slide[1]/shape[1]");

        // File should be valid after reopen (no orphaned animation references)
        Reopen();
        var slide = _handler.Get("/slide[1]");
        slide.Children.Should().BeEmpty("shape was removed, no shapes should remain");
    }

    [Fact]
    public void RemoveShape_WithAnimation_OtherShapeAnimationSurvives()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Keep", ["fill"] = "00FF00" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Remove", ["fill"] = "FF0000" });
        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "fade-entrance-500" });
        _handler.Set("/slide[1]/shape[2]", new() { ["animation"] = "fly-entrance-600" });

        // Remove shape[2]
        _handler.Remove("/slide[1]/shape[2]");

        // shape[1] should still have its animation
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Keep");
        node.Format.Should().ContainKey("animation");
        ((string)node.Format["animation"]).Should().Contain("fade");
    }

    [Fact]
    public void RemoveShape_WithAnimation_Persist_FileRemainsValid()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Stay" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Go" });
        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "zoom-entrance-400" });
        _handler.Set("/slide[1]/shape[2]", new() { ["animation"] = "fade-entrance-300" });

        _handler.Remove("/slide[1]/shape[2]");

        // Reopen — this will fail if animation XML references a deleted shape
        Reopen();

        var shapes = _handler.Get("/slide[1]").Children;
        shapes.Should().HaveCount(1);
        shapes[0].Text.Should().Be("Stay");
        shapes[0].Format.Should().ContainKey("animation");
    }

    [Fact]
    public void RemoveShape_NoAnimation_DoesNotCorrupt()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Plain" });

        _handler.Remove("/slide[1]/shape[1]");

        Reopen();
        var slide = _handler.Get("/slide[1]");
        slide.Children.Should().BeEmpty();
    }

    // ==================== Animation trigger structure tests ====================

    [Fact]
    public void Animation_WithTrigger_NestedInsidePreviousAnimationPar()
    {
        // Setup: 3 shapes with click, after, with triggers
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A", ["animation"] = "fade-entrance-500"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "B", ["animation"] = "fade-entrance-500-after"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "C", ["animation"] = "fade-entrance-500-with"
        });

        // Access the raw slide XML to inspect timing structure
        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var timing = slidePart.Slide.GetFirstChild<DocumentFormat.OpenXml.Presentation.Timing>()!;

        // Find mainSeq
        var mainSeqCTn = timing.Descendants<DocumentFormat.OpenXml.Presentation.CommonTimeNode>()
            .First(c => c.NodeType?.Value == DocumentFormat.OpenXml.Presentation.TimeNodeValues.MainSequence);

        // mainSeq should have exactly 2 direct child pars:
        //   par[0]: click group (shape A) — delay="indefinite"
        //   par[1]: after group (shape B + shape C nested) — delay="0"
        // Shape C (with) should NOT be a 3rd separate par sibling
        var topPars = mainSeqCTn.ChildTimeNodeList!
            .Elements<DocumentFormat.OpenXml.Presentation.ParallelTimeNode>().ToList();
        topPars.Should().HaveCount(2,
            "withEffect animation should be nested inside the previous par, not a separate sibling");

        // Verify par[0] is click (delay=indefinite)
        var par0Delay = topPars[0].CommonTimeNode!.StartConditionList!
            .Elements<DocumentFormat.OpenXml.Presentation.Condition>().First().Delay!.Value;
        par0Delay.Should().Be("indefinite", "first animation group should be click-triggered");

        // Verify par[1] is after (delay=0) and contains BOTH shape B and shape C animations
        var par1Delay = topPars[1].CommonTimeNode!.StartConditionList!
            .Elements<DocumentFormat.OpenXml.Presentation.Condition>().First().Delay!.Value;
        par1Delay.Should().Be("0", "second animation group should be after-triggered");

        // par[1] should contain 2 mid-pars (one for B=afterEffect, one for C=withEffect)
        var par1Children = topPars[1].CommonTimeNode!.ChildTimeNodeList!
            .Elements<DocumentFormat.OpenXml.Presentation.ParallelTimeNode>().ToList();
        par1Children.Should().HaveCount(2,
            "after group should contain both the afterEffect and the withEffect animations");
    }

    [Fact]
    public void Animation_WithTrigger_NestedInsidePreviousAnimationPar_Persist()
    {
        // Same as above but verify after reopen
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A", ["animation"] = "fade-entrance-500"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "B", ["animation"] = "fade-entrance-500-after"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "C", ["animation"] = "fade-entrance-500-with"
        });

        Reopen();

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var timing = slidePart.Slide.GetFirstChild<DocumentFormat.OpenXml.Presentation.Timing>()!;

        var mainSeqCTn = timing.Descendants<DocumentFormat.OpenXml.Presentation.CommonTimeNode>()
            .First(c => c.NodeType?.Value == DocumentFormat.OpenXml.Presentation.TimeNodeValues.MainSequence);

        var topPars = mainSeqCTn.ChildTimeNodeList!
            .Elements<DocumentFormat.OpenXml.Presentation.ParallelTimeNode>().ToList();
        topPars.Should().HaveCount(2,
            "withEffect should remain nested after persist, not become a separate sibling");
    }

    [Fact]
    public void Animation_AfterTrigger_IsAutoPlayNotClick()
    {
        // Verify "after" trigger creates delay="0" (auto-play), not delay="indefinite" (click)
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Click me", ["animation"] = "fade-entrance-500"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Auto after", ["animation"] = "fade-entrance-500-after"
        });

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var timing = slidePart.Slide.GetFirstChild<DocumentFormat.OpenXml.Presentation.Timing>()!;

        var mainSeqCTn = timing.Descendants<DocumentFormat.OpenXml.Presentation.CommonTimeNode>()
            .First(c => c.NodeType?.Value == DocumentFormat.OpenXml.Presentation.TimeNodeValues.MainSequence);

        var topPars = mainSeqCTn.ChildTimeNodeList!
            .Elements<DocumentFormat.OpenXml.Presentation.ParallelTimeNode>().ToList();
        topPars.Should().HaveCount(2);

        // par[0] = click (indefinite), par[1] = after (0)
        topPars[0].CommonTimeNode!.StartConditionList!
            .Elements<DocumentFormat.OpenXml.Presentation.Condition>().First()
            .Delay!.Value.Should().Be("indefinite");
        topPars[1].CommonTimeNode!.StartConditionList!
            .Elements<DocumentFormat.OpenXml.Presentation.Condition>().First()
            .Delay!.Value.Should().Be("0");

        // Verify nodeType values
        var effectNodes = timing.Descendants<DocumentFormat.OpenXml.Presentation.CommonTimeNode>()
            .Where(c => c.PresetId != null).ToList();
        effectNodes.Should().HaveCount(2);
        effectNodes[0].NodeType!.Value.Should().Be(DocumentFormat.OpenXml.Presentation.TimeNodeValues.ClickEffect);
        effectNodes[1].NodeType!.Value.Should().Be(DocumentFormat.OpenXml.Presentation.TimeNodeValues.AfterEffect);
    }

    [Fact]
    public void Animation_ClickThenWith_OnlyOneClickGroup()
    {
        // Repro: cool-morph slide 3 — TitleText=click, SubText=with
        // Bug: with animation became a separate p:par sibling, consuming an extra click
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Make it", ["animation"] = "zoom-entrance-600"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Subtitle", ["animation"] = "fade-entrance-500-with"
        });

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var timing = slidePart.Slide.GetFirstChild<DocumentFormat.OpenXml.Presentation.Timing>()!;

        var mainSeqCTn = timing.Descendants<DocumentFormat.OpenXml.Presentation.CommonTimeNode>()
            .First(c => c.NodeType?.Value == DocumentFormat.OpenXml.Presentation.TimeNodeValues.MainSequence);

        // Should be exactly 1 top-level par (one click group),
        // not 2 (which would mean the with-animation consumes an extra click)
        var topPars = mainSeqCTn.ChildTimeNodeList!
            .Elements<DocumentFormat.OpenXml.Presentation.ParallelTimeNode>().ToList();
        topPars.Should().HaveCount(1,
            "with-animation should be nested inside the click group, not a separate click step");

        // The single click group should contain 2 mid-pars (zoom + fade)
        var midPars = topPars[0].CommonTimeNode!.ChildTimeNodeList!
            .Elements<DocumentFormat.OpenXml.Presentation.ParallelTimeNode>().ToList();
        midPars.Should().HaveCount(2,
            "click group should contain both the click animation and the with animation");
    }

    [Fact]
    public void Animation_ClickThenWith_OnlyOneClickGroup_Persist()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Make it", ["animation"] = "zoom-entrance-600"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Subtitle", ["animation"] = "fade-entrance-500-with"
        });

        Reopen();

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var timing = slidePart.Slide.GetFirstChild<DocumentFormat.OpenXml.Presentation.Timing>()!;

        var mainSeqCTn = timing.Descendants<DocumentFormat.OpenXml.Presentation.CommonTimeNode>()
            .First(c => c.NodeType?.Value == DocumentFormat.OpenXml.Presentation.TimeNodeValues.MainSequence);

        var topPars = mainSeqCTn.ChildTimeNodeList!
            .Elements<DocumentFormat.OpenXml.Presentation.ParallelTimeNode>().ToList();
        topPars.Should().HaveCount(1,
            "with-animation should remain nested after persist, not become a separate click step");
    }

    [Fact]
    public void MotionPath_WithTrigger_NestedInsidePreviousAnimationPar()
    {
        // motionPath with "with" trigger has the same bug as shape animation
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A", ["animation"] = "fade-entrance-500"
        });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });

        // Set motionPath with "with" trigger on shape B
        _handler.Set("/slide[1]/shape[2]", new()
        {
            ["motionPath"] = "M 0.0 0.0 L 0.5 0.5 E-500-with"
        });

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var timing = slidePart.Slide.GetFirstChild<DocumentFormat.OpenXml.Presentation.Timing>()!;

        var mainSeqCTn = timing.Descendants<DocumentFormat.OpenXml.Presentation.CommonTimeNode>()
            .First(c => c.NodeType?.Value == DocumentFormat.OpenXml.Presentation.TimeNodeValues.MainSequence);

        // Should be 1 par (click group containing both the fade and the motionPath),
        // not 2 separate pars
        var topPars = mainSeqCTn.ChildTimeNodeList!
            .Elements<DocumentFormat.OpenXml.Presentation.ParallelTimeNode>().ToList();
        topPars.Should().HaveCount(1,
            "motionPath with 'with' trigger should be nested inside the click group, not a separate sibling");
    }

    [Fact]
    public void MotionPath_WithTrigger_NestedInsidePreviousAnimationPar_Persist()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A", ["animation"] = "fade-entrance-500"
        });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        _handler.Set("/slide[1]/shape[2]", new()
        {
            ["motionPath"] = "M 0.0 0.0 L 0.5 0.5 E-500-with"
        });

        Reopen();

        var doc = _handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler) as DocumentFormat.OpenXml.Packaging.PresentationDocument;
        var slidePart = doc!.PresentationPart!.SlideParts.First();
        var timing = slidePart.Slide.GetFirstChild<DocumentFormat.OpenXml.Presentation.Timing>()!;

        var mainSeqCTn = timing.Descendants<DocumentFormat.OpenXml.Presentation.CommonTimeNode>()
            .First(c => c.NodeType?.Value == DocumentFormat.OpenXml.Presentation.TimeNodeValues.MainSequence);

        var topPars = mainSeqCTn.ChildTimeNodeList!
            .Elements<DocumentFormat.OpenXml.Presentation.ParallelTimeNode>().ToList();
        topPars.Should().HaveCount(1,
            "motionPath with 'with' trigger should remain nested after persist");
    }

    // ==================== Slide Zoom ====================

    [Fact]
    public void Add_Zoom_ReturnsCorrectPath()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        var path = _handler.Add("/slide[1]", "zoom", null, new() { ["target"] = "2" });
        path.Should().Be("/slide[1]/zoom[1]");
    }

    [Fact]
    public void Add_Zoom_GetReturnsZoomType()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new() { ["target"] = "2" });

        var node = _handler.Get("/slide[1]/zoom[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("zoom");
        node.Format.Should().ContainKey("target");
        node.Format["target"].Should().Be(2);
    }

    [Fact]
    public void Add_Zoom_WithPosition_PositionIsReadBack()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new()
        {
            ["target"] = "2", ["x"] = "3cm", ["y"] = "5cm",
            ["width"] = "9cm", ["height"] = "5cm"
        });

        var node = _handler.Get("/slide[1]/zoom[1]");
        node.Format["x"].Should().Be("3cm");
        node.Format["y"].Should().Be("5cm");
        node.Format["width"].Should().Be("9cm");
        node.Format["height"].Should().Be("5cm");
    }

    [Fact]
    public void Add_Zoom_WithReturnToParent_PropertyIsReadBack()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new()
        {
            ["target"] = "2", ["returntoparent"] = "true"
        });

        var node = _handler.Get("/slide[1]/zoom[1]");
        node.Format["returnToParent"].Should().Be("1");
    }

    [Fact]
    public void Add_Zoom_WithTransitionDur_PropertyIsReadBack()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new()
        {
            ["target"] = "2", ["transitiondur"] = "2000"
        });

        var node = _handler.Get("/slide[1]/zoom[1]");
        node.Format["transitionDur"].Should().Be("2000");
    }

    [Fact]
    public void Add_Zoom_MultipleTargets_QueryReturnsAll()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new() { ["target"] = "2" });
        _handler.Add("/slide[1]", "zoom", null, new() { ["target"] = "3" });

        var results = _handler.Query("zoom");
        results.Should().HaveCount(2);
        results[0].Format["target"].Should().Be(2);
        results[1].Format["target"].Should().Be(3);
    }

    [Fact]
    public void Set_Zoom_ReturnToParentIsUpdated()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new() { ["target"] = "2" });

        var before = _handler.Get("/slide[1]/zoom[1]");
        before.Format["returnToParent"].Should().Be("0");

        _handler.Set("/slide[1]/zoom[1]", new() { ["returnToParent"] = "true" });

        var after = _handler.Get("/slide[1]/zoom[1]");
        after.Format["returnToParent"].Should().Be("1");
    }

    [Fact]
    public void Set_Zoom_TransitionDurIsUpdated()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new() { ["target"] = "2" });

        _handler.Set("/slide[1]/zoom[1]", new() { ["transitionDur"] = "500" });

        var node = _handler.Get("/slide[1]/zoom[1]");
        node.Format["transitionDur"].Should().Be("500");
    }

    [Fact]
    public void Set_Zoom_PositionIsUpdated()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new()
        {
            ["target"] = "2", ["x"] = "1cm", ["y"] = "1cm"
        });

        _handler.Set("/slide[1]/zoom[1]", new() { ["x"] = "5cm", ["y"] = "8cm" });

        var node = _handler.Get("/slide[1]/zoom[1]");
        node.Format["x"].Should().Be("5cm");
        node.Format["y"].Should().Be("8cm");
    }

    [Fact]
    public void Remove_Zoom_ElementIsRemoved()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new() { ["target"] = "2" });
        _handler.Add("/slide[1]", "zoom", null, new() { ["target"] = "3" });

        _handler.Query("zoom").Should().HaveCount(2);

        _handler.Remove("/slide[1]/zoom[1]");

        var remaining = _handler.Query("zoom");
        remaining.Should().HaveCount(1);
        remaining[0].Format["target"].Should().Be(3);
    }

    [Fact]
    public void Zoom_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "zoom", null, new()
        {
            ["target"] = "2", ["x"] = "3cm", ["y"] = "5cm",
            ["width"] = "9cm", ["height"] = "5cm",
            ["returntoparent"] = "true", ["transitiondur"] = "1500"
        });

        Reopen();

        var node = _handler.Get("/slide[1]/zoom[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("zoom");
        node.Format["target"].Should().Be(2);
        node.Format["x"].Should().Be("3cm");
        node.Format["y"].Should().Be("5cm");
        node.Format["width"].Should().Be("9cm");
        node.Format["height"].Should().Be("5cm");
        node.Format["returnToParent"].Should().Be("1");
        node.Format["transitionDur"].Should().Be("1500");
    }
}
