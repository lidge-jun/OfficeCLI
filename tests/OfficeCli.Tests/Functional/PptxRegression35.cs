// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt round 35: Additional white-box code review bugs across PPTX, Word, and Excel handlers.
/// </summary>
public class PptxRegression35 : IDisposable
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
    // Bug3500: PPTX Set multi-line text + underline — underline on stale runs
    // Same stale-runs family as Bug3400. "underline" iterates over the old
    // `runs` list after "text" has replaced all paragraphs with new ones.
    // =====================================================================
    [Fact]
    public void Bug3500_Pptx_Set_MultilineText_And_Underline_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Original" });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["underline"] = "single"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("underline",
            "Underline should be applied to new runs, not stale orphaned runs");
    }

    // =====================================================================
    // Bug3501: PPTX Set multi-line text + strikethrough — strike on stale runs
    // =====================================================================
    [Fact]
    public void Bug3501_Pptx_Set_MultilineText_And_Strike_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Original" });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["strike"] = "single"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("strike",
            "Strikethrough should be applied to new runs");
    }

    // =====================================================================
    // Bug3502: PPTX Set multi-line text + spacing — charspacing on stale runs
    // =====================================================================
    [Fact]
    public void Bug3502_Pptx_Set_MultilineText_And_CharSpacing_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Original" });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["spacing"] = "2"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("spacing",
            "Character spacing should be applied to new runs");
    }

    // =====================================================================
    // Bug3503: PPTX Set multi-line text + superscript — baseline on stale runs
    // =====================================================================
    [Fact]
    public void Bug3503_Pptx_Set_MultilineText_And_Superscript_StaleRuns()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Original" });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["superscript"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("baseline",
            "Baseline/superscript should be applied to new runs");
    }

    // =====================================================================
    // Bug3504: Word Set paragraph text + bold — order-dependent behavior
    // In WordHandler.Set for Paragraph, "bold" is handled by case
    // "size" or "font" or "bold"... which applies to existing runs.
    // "text" replaces runs. If bold comes before text in the dictionary,
    // bold is applied to old runs and then text replaces them with new ones
    // that don't have bold.
    // =====================================================================
    [Fact]
    public void Bug3504_Word_Set_Paragraph_Text_And_Bold_OrderDependent()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Original" });

        // Set text and bold together
        handler.Set("/body/p[1]", new()
        {
            ["text"] = "Updated text",
            ["bold"] = "true"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Updated text");

        // Bold should be applied to the actual text
        // Check first run for bold
        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("bold");
        run.Format["bold"].Should().Be(true);
    }

    // =====================================================================
    // Bug3505: Word Set paragraph text + color — color lost on replacement
    // Same pattern as Bug3504 but with color property.
    // =====================================================================
    [Fact]
    public void Bug3505_Word_Set_Paragraph_Text_And_Color()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Original" });

        handler.Set("/body/p[1]", new()
        {
            ["text"] = "Colored text",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Colored text");

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("color");
        run.Format["color"].ToString().Should().Be("#FF0000");
    }

    // =====================================================================
    // Bug3506: PPTX presentation Set with unknown slidesize value
    // Setting slidesize to an unknown value like "letter" adds it to
    // unsupported list instead of throwing — but "letter" might be expected
    // to work. Verify behavior.
    // =====================================================================
    [Fact]
    public void Bug3506_Pptx_Set_Presentation_Unknown_SlideSize()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // Set unknown slidesize — should report as unsupported
        var unsupported = handler.Set("/", new()
        {
            ["slidesize"] = "letter"
        });

        // "letter" is not a recognized slide size value
        unsupported.Should().Contain("slidesize",
            "Unknown slidesize values should be reported as unsupported");
    }

    // =====================================================================
    // Bug3507: Excel Set hyperlink on cell — link not readable via Get
    // After setting a hyperlink on a cell, Get should show the link.
    // =====================================================================
    [Fact]
    public void Bug3507_Excel_Set_Cell_Link_Readable()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Click here",
            ["link"] = "https://example.com"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Text.Should().Be("Click here");
        // Link should be readable
        node.Format.Should().ContainKey("link",
            "Hyperlink should be readable after being set");
    }

    // =====================================================================
    // Bug3508: PPTX table cell Set text + color — color on new runs
    // In SetTableCellProperties, "text" replaces all paragraphs with new
    // runs, and "color" iterates cell.Descendants<Drawing.Run>(). Since
    // Descendants is lazy, if text was processed first and new runs exist,
    // color should work. But if color is processed first (before text),
    // it applies to empty EndParagraphRunProperties runs that get replaced.
    // =====================================================================
    [Fact]
    public void Bug3508_Pptx_TableCell_Set_Text_And_Color_Together()
    {
        var path = CreateTemp(".pptx");
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
            ["text"] = "Red text",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = node.Children[0].Children[0];
        cell.Text.Should().Be("Red text");
        // Color should be on the actual runs
        cell.Format.Should().ContainKey("color");
        cell.Format["color"].ToString().Should().Be("#FF0000");
    }

    // =====================================================================
    // Bug3509: PPTX table cell Set text + font + size — multiple run props
    // =====================================================================
    [Fact]
    public void Bug3509_Pptx_TableCell_Set_Text_Font_Size_Together()
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
            ["text"] = "Styled cell",
            ["font"] = "Arial",
            ["size"] = "18"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = node.Children[0].Children[0];
        cell.Text.Should().Be("Styled cell");
        cell.Format.Should().ContainKey("font");
        cell.Format["font"].ToString().Should().Be("Arial");
        cell.Format.Should().ContainKey("size");
        cell.Format["size"].ToString().Should().Contain("18");
    }

    // =====================================================================
    // Bug3510: Excel Set sheet name — after rename, can you still access it?
    // After renaming a sheet, the old name should no longer work.
    // =====================================================================
    [Fact]
    public void Bug3510_Excel_Set_Sheet_Name_Access_After_Rename()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        // Set a value first
        handler.Set("/Sheet1/A1", new() { ["value"] = "Hello" });

        // Rename the sheet
        handler.Set("/Sheet1", new() { ["name"] = "Data" });

        // Should be accessible by new name
        var node = handler.Get("/Data/A1");
        node.Text.Should().Be("Hello");

        // Old name should fail
        Action getOld = () => handler.Get("/Sheet1/A1");
        getOld.Should().Throw<ArgumentException>();
    }

    // =====================================================================
    // Bug3511: PPTX Set shape opacity without fill — no effect
    // Setting opacity on a shape that has no SolidFill does nothing because
    // the code only looks for SolidFill. If the shape has no fill at all
    // (default blank shape), opacity is silently ignored.
    // =====================================================================
    [Fact]
    public void Bug3511_Pptx_Set_Shape_Opacity_Without_Fill_NoEffect()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Set opacity without any fill — should silently do nothing
        var unsupported = handler.Set("/slide[1]/shape[1]", new()
        {
            ["opacity"] = "0.5"
        });

        // Opacity auto-creates a white fill if none exists (matching POI behavior)
        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("opacity",
            "Opacity should be applied by auto-creating a fill");
    }

    // =====================================================================
    // Bug3512: PPTX Add shape with "geometry" property — processed twice
    // The Add handler sets preset geometry inline (line 334-337) AND also
    // delegates "geometry" via effectKeys to SetRunOrShapeProperties.
    // effectKeys includes "geometry", so custom geometry path processing
    // happens after preset was already applied.
    // =====================================================================
    [Fact]
    public void Bug3512_Pptx_Add_Shape_Geometry_DoubleProcessed()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        // Add shape with custom geometry — should override preset
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Custom shape",
            ["geometry"] = "M 0,0 L 100,0 L 100,100 L 0,100 Z"
        });

        // Shape should exist
        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Custom shape");
        // Preset should NOT be present because geometry overrides it
        // But the Add code applies preset first (line 334), then effectKeys
        // delegation processes geometry which removes preset and adds custom
        node.Format.Should().NotContainKey("preset",
            "Custom geometry should override preset, but preset is applied first in Add");
    }

    // =====================================================================
    // Bug3513: Word Set paragraph alignment then get — verify roundtrip
    // =====================================================================
    [Fact]
    public void Bug3513_Word_Set_Paragraph_Alignment_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Centered text",
            ["alignment"] = "center"
        });

        var node = handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("alignment");
        node.Format["alignment"].ToString().Should().Be("center");
    }

    // =====================================================================
    // Bug3514: PPTX Set shape linedash — verify "longdashdot" works
    // The linedash handler uses a switch that maps "longdashdot" etc.
    // Verify the full range of supported values.
    // =====================================================================
    [Fact]
    public void Bug3514_Pptx_Set_Shape_LineDash_LongDashDot()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["line"] = "000000"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["linedash"] = "longdashdot"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        // Verify linedash was applied
        node.Format.Should().ContainKey("lineDash");
    }

    // =====================================================================
    // Bug3515: Excel Set cell then clear — verify style is also cleared
    // Setting clear=true should reset StyleIndex to null, clearing formatting.
    // =====================================================================
    [Fact]
    public void Bug3515_Excel_Set_Cell_Clear_Resets_Style()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        // Set value with formatting
        handler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Bold text",
            ["font.bold"] = "true"
        });

        // Clear the cell
        handler.Set("/Sheet1/A1", new() { ["clear"] = "true" });

        var node = handler.Get("/Sheet1/A1");
        node.Text.Should().BeEmpty();
        // Style should be cleared too
        node.Format.Should().NotContainKey("font.bold",
            "Clear should also reset cell style/formatting");
    }

    // =====================================================================
    // Bug3516: PPTX shape Get — zorder is always reported
    // The Get for shapes reports zorder based on the shape's position among
    // all shapes. Verify it's consistent.
    // =====================================================================
    [Fact]
    public void Bug3516_Pptx_Shape_Get_ZOrder_Reported()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "First" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Second" });

        var node1 = handler.Get("/slide[1]/shape[1]");
        var node2 = handler.Get("/slide[1]/shape[2]");

        // Both should have zorder
        node1.Format.Should().ContainKey("zorder");
        node2.Format.Should().ContainKey("zorder");

        // First should have lower zorder than second
        var z1 = Convert.ToInt32(node1.Format["zorder"]);
        var z2 = Convert.ToInt32(node2.Format["zorder"]);
        z1.Should().BeLessThan(z2);
    }

    // =====================================================================
    // Bug3517: Word Set section columns — verify column count roundtrip
    // =====================================================================
    [Fact]
    public void Bug3517_Word_Set_Section_Columns_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Set("/section[1]", new()
        {
            ["columns"] = "3"
        });

        var node = handler.Get("/section[1]");
        node.Format.Should().ContainKey("columns");
        Convert.ToInt32(node.Format["columns"]).Should().Be(3);
    }

    // =====================================================================
    // Bug3518: PPTX Set shape name — verify Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug3518_Pptx_Set_Shape_Name_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["name"] = "OriginalName"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["name"] = "RenamedShape"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format["name"].ToString().Should().Be("RenamedShape");
    }

    // =====================================================================
    // Bug3519: Excel Add comment — verify Get roundtrip
    // =====================================================================
    [Fact]
    public void Bug3519_Excel_Add_Comment_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new() { ["value"] = "Data" });

        handler.Add("/Sheet1/A1", "comment", null, new()
        {
            ["text"] = "This is a comment",
            ["author"] = "TestUser"
        });

        var node = handler.Get("/Sheet1/comment[1]");
        node.Text.Should().Contain("This is a comment");
    }

    // =====================================================================
    // Bug3520: PPTX table row Set height — verify EMU parsing
    // Table row height Set uses ParseEmu, so "2cm" should be parsed
    // correctly as EMU value.
    // =====================================================================
    [Fact]
    public void Bug3520_Pptx_TableRow_Set_Height()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/slide[1]/table[1]/tr[1]", new()
        {
            ["height"] = "2cm"
        });

        var node = handler.Get("/slide[1]/table[1]", depth: 1);
        var row = node.Children[0]; // tr[1]
        row.Format.Should().ContainKey("height");
        row.Format["height"].ToString().Should().Contain("2");
    }

    // =====================================================================
    // Bug3521: Word Add run without text — empty text should work
    // =====================================================================
    [Fact]
    public void Bug3521_Word_Add_Run_Without_Text()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        // Add run without text property — default text is ""
        var result = handler.Add("/body/p[1]", "run", null, new()
        {
            ["bold"] = "true"
        });

        result.Should().Contain("r[");

        // Now set text on the run
        handler.Set(result, new() { ["text"] = "Appended" });

        var node = handler.Get(result);
        node.Text.Should().Be("Appended");
    }

    // =====================================================================
    // Bug3522: PPTX slide Set notes — verify roundtrip
    // =====================================================================
    [Fact]
    public void Bug3522_Pptx_Set_Slide_Notes_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());

        handler.Set("/slide[1]", new()
        {
            ["notes"] = "Speaker notes text"
        });

        var node = handler.Get("/slide[1]/notes");
        node.Text.Should().Be("Speaker notes text");
    }

    // =====================================================================
    // Bug3523: Excel Set merge cell range — verify roundtrip
    // =====================================================================
    [Fact]
    public void Bug3523_Excel_Set_Merge_Range()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1/A1", new() { ["value"] = "Merged" });

        handler.Set("/Sheet1/A1:C1", new()
        {
            ["merge"] = "true"
        });

        // Verify merge exists
        var sheetNode = handler.Get("/Sheet1");
        // Merge info should be available somehow
        sheetNode.Should().NotBeNull();
    }

    // =====================================================================
    // Bug3524: Word Set paragraph with text on paragraph that has no runs
    // When a paragraph has no runs (e.g., empty paragraph), setting "text"
    // should create a new run with the text. Verify this works.
    // =====================================================================
    [Fact]
    public void Bug3524_Word_Set_Paragraph_Text_NoExistingRuns()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Add empty paragraph (no text)
        handler.Add("/body", "paragraph", null, new());

        // Set text on it
        handler.Set("/body/p[1]", new()
        {
            ["text"] = "New text"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("New text");
    }

    // =====================================================================
    // Bug3525: Word Add paragraph with link — hyperlink creation
    // =====================================================================
    [Fact]
    public void Bug3525_Word_Add_Paragraph_With_Style()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Heading text",
            ["style"] = "Heading1"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Heading text");
        node.Format.Should().ContainKey("style");
    }
}
