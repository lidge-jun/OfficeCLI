// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt round 37: New white-box bugs found via deep code review.
/// Focus areas:
///   - PPTX table cell Set: font/color lost when text replaces empty cell runs
///   - Word watermark color: VML needs # prefix but SanitizeHex strips it
///   - Word section orientation: only sets Orient attribute, doesn't swap W/H
///   - PPTX stale-runs: multiline text + run-level props on shape
///   - PPTX table cell Get: font/bold/size/color not reported
///   - Word footnote/endnote Set: non-text props silently ignored (not in unsupported)
/// </summary>
public class PptxRegression37 : IDisposable
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
    // Bug3700: PPTX table cell — font set on empty cell, then text replaces
    // In SetTableCellProperties, "font" case uses cell.Descendants<Run>().
    // On a fresh table cell with no text, there are no runs.
    // The "font" case iterates over empty runs list (no-op).
    // Then "text" case creates new runs WITHOUT font (only Language=en-US).
    // Result: font is lost even though user set it in same call.
    // =====================================================================
    [Fact]
    public void Bug3700_Pptx_TableCell_Font_Lost_On_EmptyCell_With_Text()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        // On a fresh empty table cell, setting font + text together:
        // font is applied to existing (empty) run list - no-op
        // then text creates a new run without font
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["font"] = "Impact",
            ["text"] = "Hello"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = tableNode.Children[0].Children[0];
        cell.Text.Should().Be("Hello");
        cell.Format.Should().ContainKey("font",
            "Font should be applied to the run created by 'text' when set together on empty cell");
        cell.Format["font"].ToString().Should().Be("Impact");
    }

    // =====================================================================
    // Bug3701: PPTX table cell — color set on empty cell, then text replaces
    // Same as Bug3700 but for color property.
    // =====================================================================
    [Fact]
    public void Bug3701_Pptx_TableCell_Color_Lost_On_EmptyCell_With_Text()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["color"] = "FF0000",
            ["text"] = "Red text"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = tableNode.Children[0].Children[0];
        cell.Text.Should().Be("Red text");
        cell.Format.Should().ContainKey("color",
            "Color should appear on the run created by 'text' for empty cell");
        cell.Format["color"].ToString().Should().Be("#FF0000");
    }

    // =====================================================================
    // Bug3702: PPTX table cell — bold set on empty cell, then text replaces
    // Same family: bold is applied to empty runs list (no-op), then text
    // creates a new run that doesn't inherit the bold setting.
    // =====================================================================
    [Fact]
    public void Bug3702_Pptx_TableCell_Bold_Lost_On_EmptyCell_With_Text()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["bold"] = "true",
            ["text"] = "Bold text"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = tableNode.Children[0].Children[0];
        cell.Text.Should().Be("Bold text");
        cell.Format.Should().ContainKey("bold",
            "Bold should be preserved on new run created by 'text' for empty cell");
        cell.Format["bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug3703: Word watermark Set color — VML fillcolor needs '#' prefix
    // WordHandler.Set for /watermark does:
    //   var clr = SanitizeHex(value);  // strips '#' -> "FF0000"
    //   xml = Regex.Replace(xml, @"fillcolor=""[^""]*""", $@"fillcolor=""{clr}""");
    // But VML spec requires fillcolor="#FF0000" with the '#' prefix.
    // So the watermark color is stored as fillcolor="FF0000" which is invalid VML.
    // =====================================================================
    [Fact]
    public void Bug3703_Word_Watermark_Color_VML_Missing_Hash()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "watermark", null, new()
        {
            ["text"] = "DRAFT"
        });

        handler.Set("/watermark", new()
        {
            ["color"] = "FF0000"
        });

        var node = handler.Get("/watermark");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("color");
        var colorVal = node.Format["color"].ToString()!;
        colorVal.Should().StartWith("#",
            "VML fillcolor attribute requires '#' prefix (e.g., '#FF0000'). " +
            "SanitizeHex strips '#' but VML needs it back.");
    }

    // =====================================================================
    // Bug3704: Word section Set orientation — only sets Orient, doesn't swap W/H
    // The code at WordHandler.Set:
    //   ps.Orient = value == "landscape" ? Landscape : Portrait;
    // Only sets the Orient attribute without swapping W/H values.
    // For landscape, width should be greater than height.
    // =====================================================================
    [Fact]
    public void Bug3704_Word_Section_Orientation_Swap_WH_Not_Done()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Set("/section[1]", new()
        {
            ["orientation"] = "landscape"
        });

        var secAfter = handler.Get("/section[1]");

        if (secAfter.Format.ContainsKey("pageWidth") && secAfter.Format.ContainsKey("pageHeight"))
        {
            var w = Convert.ToInt32(secAfter.Format["pageWidth"]);
            var h = Convert.ToInt32(secAfter.Format["pageHeight"]);
            w.Should().BeGreaterThan(h,
                "Landscape orientation requires width > height, " +
                "but the code only sets the Orient attribute without swapping W and H values.");
        }
        else
        {
            // At minimum, verify the orientation attribute is set
            secAfter.Format.Should().ContainKey("orientation",
                "Section should expose orientation in Get after Set");
        }
    }

    // =====================================================================
    // Bug3705: PPTX shape Set multiline text + underline (stale runs)
    // SetRunOrShapeProperties captures allRuns BEFORE "text" replaces paragraphs.
    // "underline" case iterates stale runs (already removed from DOM).
    // =====================================================================
    [Fact]
    public void Bug3705_Pptx_StaleRuns_MultilineText_Underline()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["underline"] = "single"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("underline",
            "Underline should be applied to new runs after multiline text replacement (stale runs bug)");
        node.Format["underline"].ToString().Should().Be("single");
    }

    // =====================================================================
    // Bug3706: PPTX shape Set multiline text + bold (stale runs)
    // =====================================================================
    [Fact]
    public void Bug3706_Pptx_StaleRuns_MultilineText_Bold()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["bold"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("bold",
            "Bold should be applied to new runs after multiline text replacement (stale runs bug)");
        node.Format["bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug3707: PPTX shape Set multiline text + color (stale runs)
    // =====================================================================
    [Fact]
    public void Bug3707_Pptx_StaleRuns_MultilineText_Color()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("color",
            "Color should be applied to new runs after multiline text replacement (stale runs bug)");
        node.Format["color"].ToString().Should().Be("#FF0000");
    }

    // =====================================================================
    // Bug3708: PPTX shape Set multiline text + font (stale runs)
    // =====================================================================
    [Fact]
    public void Bug3708_Pptx_StaleRuns_MultilineText_Font()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["font"] = "Wingdings"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("font",
            "Font should be applied to new runs after multiline text replacement (stale runs bug)");
        node.Format["font"].ToString().Should().Be("Wingdings");
    }

    // =====================================================================
    // Bug3709: PPTX shape Set multiline text + size (stale runs)
    // =====================================================================
    [Fact]
    public void Bug3709_Pptx_StaleRuns_MultilineText_Size()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["size"] = "36"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("size",
            "Size should be applied to new runs after multiline text replacement (stale runs bug)");
        node.Format["size"].ToString().Should().Contain("36");
    }

    // =====================================================================
    // Bug3710: PPTX table cell Get — bold not reported by TableToNode
    // In PowerPointHandler.NodeBuilder.cs TableToNode, the cell node is built
    // with text, fill, and borders, but run-level formatting (bold, font, color,
    // size) is NOT read from cell runs and NOT added to cellNode.Format.
    // =====================================================================
    [Fact]
    public void Bug3710_Pptx_TableCell_Get_Does_Not_Report_Bold()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        // Two separate calls to avoid stale-runs issue
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold cell"
        });
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["bold"] = "true"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = tableNode.Children[0].Children[0];
        cell.Text.Should().Be("Bold cell");
        cell.Format.Should().ContainKey("bold",
            "TableToNode should read bold from cell runs, " +
            "but it currently doesn't report any run-level formatting");
    }

    // =====================================================================
    // Bug3711: PPTX table cell Get — font not reported by TableToNode
    // =====================================================================
    [Fact]
    public void Bug3711_Pptx_TableCell_Get_Does_Not_Report_Font()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Styled"
        });
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["font"] = "Calibri"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = tableNode.Children[0].Children[0];
        cell.Format.Should().ContainKey("font",
            "TableToNode should read font from cell runs, but currently doesn't");
    }

    // =====================================================================
    // Bug3712: PPTX table cell Get — color not reported by TableToNode
    // =====================================================================
    [Fact]
    public void Bug3712_Pptx_TableCell_Get_Does_Not_Report_Color()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Colored"
        });
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["color"] = "FF0000"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = tableNode.Children[0].Children[0];
        cell.Format.Should().ContainKey("color",
            "TableToNode should read text color from cell runs");
        cell.Format["color"].ToString().Should().Be("#FF0000");
    }

    // =====================================================================
    // Bug3713: PPTX table cell Get — size not reported by TableToNode
    // =====================================================================
    [Fact]
    public void Bug3713_Pptx_TableCell_Get_Does_Not_Report_Size()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Big text"
        });
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["size"] = "24"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = tableNode.Children[0].Children[0];
        cell.Format.Should().ContainKey("size",
            "TableToNode should read font size from cell runs");
        cell.Format["size"].ToString().Should().Contain("24");
    }

    // =====================================================================
    // Bug3714: Word Set paragraph bold + text (single run) — bold preserved
    // With ONE existing run: "bold" applies to that run, "text" updates
    // its Text element (doesn't remove RunProperties). Should PASS.
    // Verifies the non-bug case to ensure correctness.
    // =====================================================================
    [Fact]
    public void Bug3714_Word_Set_Paragraph_BoldAndText_SingleRun_Works()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Initial" });

        handler.Set("/body/p[1]", new()
        {
            ["bold"] = "true",
            ["text"] = "Updated"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Updated");

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("bold",
            "Bold should survive when text and bold are set together on single-run paragraph");
    }

    // =====================================================================
    // Bug3715: Word Set footnote — non-text properties silently ignored
    // The footnote Set handler only processes "text" key. All other keys
    // (bold, color, font, etc.) are NOT added to the `unsupported` list.
    // They are silently ignored — a data-loss bug.
    // Code path: WordHandler.Set lines 265-287 — no default/else branch.
    // =====================================================================
    [Fact]
    public void Bug3715_Word_Set_Footnote_NonText_Props_Silently_Ignored()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Footnoted text" });
        handler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Footnote content" });

        var unsupported = handler.Set("/footnote[1]", new()
        {
            ["bold"] = "true",
            ["color"] = "FF0000"
        });

        unsupported.Should().Contain("bold",
            "Non-text properties in footnote Set should be reported as unsupported, " +
            "currently they are silently ignored");
        unsupported.Should().Contain("color",
            "Non-text properties in footnote Set should be reported as unsupported");
    }

    // =====================================================================
    // Bug3716: Word Set endnote — non-text properties silently ignored
    // Same silent-ignore bug as Bug3715 but for endnotes.
    // Code path: WordHandler.Set lines 308-328 — no default/else branch.
    // =====================================================================
    [Fact]
    public void Bug3716_Word_Set_Endnote_NonText_Props_Silently_Ignored()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Endnoted text" });
        handler.Add("/body/p[1]", "endnote", null, new() { ["text"] = "Endnote content" });

        var unsupported = handler.Set("/endnote[1]", new()
        {
            ["font"] = "Arial",
            ["italic"] = "true"
        });

        unsupported.Should().Contain("font",
            "Non-text properties in endnote Set should be reported as unsupported");
        unsupported.Should().Contain("italic",
            "Non-text properties in endnote Set should be reported as unsupported");
    }

    // =====================================================================
    // Bug3717: PPTX Add shape with "geometry" — inline preset applied first
    // then effectKeys re-processes "geometry" which removes PresetGeometry
    // and adds CustomGeometry. End result should be custom geometry only.
    // =====================================================================
    [Fact]
    public void Bug3717_Pptx_Add_Shape_CustomGeometry_Override_Preset()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        var result = handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Custom",
            ["geometry"] = "M 0,0 L 100,0 L 50,100 Z"
        });

        var node = handler.Get(result);
        node.Should().NotBeNull();
        node.Text.Should().Be("Custom");
        // After inline preset="rect" + geometry override:
        // final shape should have custom geometry (no "preset" key)
        node.Format.Should().NotContainKey("preset",
            "Custom geometry should override the inline-applied rect preset");
    }

    // =====================================================================
    // Bug3718: Excel Set cell link — file remains valid/readable after set
    // =====================================================================
    [Fact]
    public void Bug3718_Excel_Set_Cell_Link_File_Readable_After_Set()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Click here",
            ["link"] = "https://example.com"
        });

        handler.Dispose();
        using var handler2 = new ExcelHandler(path, editable: false);
        var node = handler2.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node.Text.Should().Be("Click here");
    }

    // =====================================================================
    // Bug3719: Excel Set cell link — link readable via Get after setting
    // =====================================================================
    [Fact]
    public void Bug3719_Excel_Set_Cell_Link_Readable_Via_Get()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Go to example",
            ["link"] = "https://example.com"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Text.Should().Be("Go to example");
        node.Format.Should().ContainKey("link",
            "Hyperlink should be readable after being set on a cell");
        node.Format["link"].ToString().Should().Contain("example.com");
    }

    // =====================================================================
    // Bug3720: PPTX shape Set single-line text — preserves RunProperties
    // Single-line + single-run path: runs[0].Text = new Drawing.Text(...)
    // This preserves RunProperties on runs[0]. Verify bold/font survive.
    // =====================================================================
    [Fact]
    public void Bug3720_Pptx_Shape_Set_SingleLine_Text_Preserves_RunProps()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original",
            ["bold"] = "true",
            ["font"] = "Arial"
        });

        // Set single-line text only — should preserve RunProperties
        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Updated"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Updated");
        node.Format.Should().ContainKey("bold",
            "Single-line text Set should preserve existing RunProperties (bold)");
        node.Format.Should().ContainKey("font");
        node.Format["font"].ToString().Should().Be("Arial");
    }

    // =====================================================================
    // Bug3721: PPTX shape Set multiline text + spacing (stale runs)
    // =====================================================================
    [Fact]
    public void Bug3721_Pptx_StaleRuns_MultilineText_Spacing()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["spacing"] = "2"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("spacing",
            "Character spacing should be applied to new runs after multiline text replacement (stale runs bug)");
    }

    // =====================================================================
    // Bug3722: PPTX shape Set multiline text + strike (stale runs)
    // =====================================================================
    [Fact]
    public void Bug3722_Pptx_StaleRuns_MultilineText_Strike()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["strike"] = "single"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("strike",
            "Strike should be applied to new runs after multiline text replacement (stale runs bug)");
    }

    // =====================================================================
    // Bug3723: Word table cell Set text+bold (deferred) on EMPTY cell
    // When text is DEFERRED in Word TableCell Set on fresh empty cell:
    //   1. "bold" case: no existing runs -> stores bold in ParagraphMarkRunProperties
    //   2. deferred "text": clones PMR props into new run's RunProperties
    // This SHOULD work via PMR cloning. Verify it does.
    // =====================================================================
    [Fact]
    public void Bug3723_Word_TableCell_DeferredText_Bold_Via_PMR()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["bold"] = "true",
            ["text"] = "Bold content"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Bold content");
        node.Format.Should().ContainKey("bold",
            "Bold set on empty cell with deferred text should be preserved via PMR cloning");
        node.Format["bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug3724: Word table cell Set text+color (deferred) on EMPTY cell
    // =====================================================================
    [Fact]
    public void Bug3724_Word_TableCell_DeferredText_Color_Via_PMR()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["color"] = "FF0000",
            ["text"] = "Red content"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Red content");
        node.Format.Should().ContainKey("color",
            "Color set on empty cell with deferred text should be preserved via PMR cloning");
        node.Format["color"].ToString().Should().Be("#FF0000");
    }

    // =====================================================================
    // Bug3725: Word table cell Set text+font (deferred) on EMPTY cell
    // =====================================================================
    [Fact]
    public void Bug3725_Word_TableCell_DeferredText_Font_Via_PMR()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["font"] = "Courier New",
            ["text"] = "Monospace content"
        });

        var node = handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Monospace content");
        node.Format.Should().ContainKey("font",
            "Font set on empty cell with deferred text should be preserved via PMR cloning");
        node.Format["font"].ToString().Should().Contain("Courier");
    }
}
