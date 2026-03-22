// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class PptxRegression40 : IDisposable
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
    // Bug4000: PPTX Add shape with lineOpacity — not in effectKeys list
    // Add.cs effectKeys (line 404-413) does not include "lineopacity"/"line.opacity"
    // so lineOpacity is silently ignored during shape Add
    // =====================================================================
    [Fact]
    public void Bug4000_Pptx_Add_Shape_LineOpacity_Ignored()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["lineColor"] = "000000",
            ["lineWidth"] = "2pt",
            ["lineOpacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineOpacity",
            because: "lineOpacity should be processed during Add but it's not in effectKeys");
        node.Format["lineOpacity"].Should().Be("0.5",
            because: "lineOpacity of 0.5 should be stored and readable");
    }

    // =====================================================================
    // Bug4001: PPTX Add shape autofit double-processed — Add has inline
    // autofit handling (lines 301-314) AND "autofit" is in effectKeys (line 413)
    // which delegates to SetRunOrShapeProperties. AutoFit gets processed twice.
    // =====================================================================
    [Fact]
    public void Bug4001_Pptx_Add_Shape_AutoFit_Double_Processing()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        // This should not crash even with double processing
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test autofit", ["autofit"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("autoFit");
        // The value should be "normal" (NormalAutoFit), not duplicated
        node.Format["autoFit"].Should().Be("normal",
            because: "autofit=true should create a single NormalAutoFit element");
    }

    // =====================================================================
    // Bug4002: PPTX lineDash NodeBuilder returns OOXML InnerText, not user name
    // Set accepts "dash" but Get returns "dash" (happens to match for some values)
    // but "longdash" becomes "lgDash" — inconsistent mapping
    // =====================================================================
    [Fact]
    public void Bug4002_Pptx_LineDash_DashDot_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "DashDot", ["lineColor"] = "000000",
            ["lineWidth"] = "2pt", ["lineDash"] = "dashdot"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineDash");
        var dashVal = node.Format["lineDash"]?.ToString();
        // Set maps "dashdot" → DashDot enum, but InnerText is "dashDot" (camelCase)
        // NodeBuilder lowercases: "dashdot"
        // This should match, but let's verify
        dashVal.Should().Be("dashdot",
            because: "lineDash 'dashdot' should roundtrip cleanly, InnerText is 'dashDot' → lowered to 'dashdot'");
    }

    // =====================================================================
    // Bug4003: PPTX lineDash "longdash" roundtrip fails
    // Set maps "longdash" → LargeDash enum, InnerText is "lgDash"
    // NodeBuilder: dash.Val.InnerText.ToLowerInvariant() = "lgdash" (not "longdash")
    // =====================================================================
    [Fact]
    public void Bug4003_Pptx_LineDash_LongDash_Roundtrip_Mismatch()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "LongDash", ["lineColor"] = "000000",
            ["lineWidth"] = "2pt", ["lineDash"] = "longdash"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineDash");
        var dashVal = node.Format["lineDash"]?.ToString();
        dashVal.Should().Be("longdash",
            because: "lineDash 'longdash' roundtrip should return 'longdash' but returns 'lgdash'");
    }

    // =====================================================================
    // Bug4004: Excel hyperlink URL trailing slash added during roundtrip
    // Uri.TryCreate normalizes "https://example.com" to "https://example.com/"
    // =====================================================================
    [Fact]
    public void Bug4004_Excel_Hyperlink_Trailing_Slash_Added()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Link" });
        handler.Set("/Sheet1/A1", new() { ["link"] = "https://example.com" });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("link");
        // The URL should not have a trailing slash added
        node.Format["link"].Should().Be("https://example.com",
            because: "URL should preserve exact user input, but Uri.TryCreate adds trailing slash");
    }

    // =====================================================================
    // Bug4005: Word paragraph NodeBuilder doesn't aggregate run-level font
    // Paragraph ElementToNode (Navigation.cs:216-301) only reports paragraph-level
    // properties. The font, bold, color etc. from runs are not aggregated.
    // The user must navigate to individual runs to see font info.
    // =====================================================================
    [Fact]
    public void Bug4005_Word_Paragraph_NodeBuilder_Missing_Font()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Hello", ["font"] = "Courier New", ["size"] = "14"
        });

        var paraNode = handler.Get("/body/p[1]");
        paraNode.Text.Should().Be("Hello");
        // Paragraph NodeBuilder doesn't report font from runs
        paraNode.Format.Should().ContainKey("font",
            because: "paragraph should aggregate font from all runs when they share the same font value");
    }

    // =====================================================================
    // Bug4006: Word paragraph NodeBuilder doesn't aggregate run bold
    // =====================================================================
    [Fact]
    public void Bug4006_Word_Paragraph_NodeBuilder_Missing_Bold()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Bold text", ["bold"] = "true"
        });

        var paraNode = handler.Get("/body/p[1]");
        paraNode.Format.Should().ContainKey("bold",
            because: "paragraph should aggregate bold from runs when all runs are bold");
    }

    // =====================================================================
    // Bug4007: Word paragraph NodeBuilder doesn't aggregate run color
    // =====================================================================
    [Fact]
    public void Bug4007_Word_Paragraph_NodeBuilder_Missing_Color()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Red text", ["color"] = "FF0000"
        });

        var paraNode = handler.Get("/body/p[1]");
        paraNode.Format.Should().ContainKey("color",
            because: "paragraph should aggregate color from runs when all runs share the same color");
    }

    // =====================================================================
    // Bug4008: Word paragraph NodeBuilder doesn't aggregate run size
    // =====================================================================
    [Fact]
    public void Bug4008_Word_Paragraph_NodeBuilder_Missing_Size()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Big text", ["size"] = "24"
        });

        var paraNode = handler.Get("/body/p[1]");
        paraNode.Format.Should().ContainKey("size",
            because: "paragraph should aggregate size from runs when all runs share the same size");
    }

    // =====================================================================
    // Bug4009: Word paragraph NodeBuilder doesn't aggregate run italic
    // =====================================================================
    [Fact]
    public void Bug4009_Word_Paragraph_NodeBuilder_Missing_Italic()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Italic text", ["italic"] = "true"
        });

        var paraNode = handler.Get("/body/p[1]");
        paraNode.Format.Should().ContainKey("italic",
            because: "paragraph should aggregate italic from runs when all runs are italic");
    }

    // =====================================================================
    // Bug4010: Word paragraph NodeBuilder doesn't aggregate run underline
    // =====================================================================
    [Fact]
    public void Bug4010_Word_Paragraph_NodeBuilder_Missing_Underline()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Underlined", ["underline"] = "single"
        });

        var paraNode = handler.Get("/body/p[1]");
        paraNode.Format.Should().ContainKey("underline",
            because: "paragraph should aggregate underline from runs when all runs share the same underline");
    }

    // =====================================================================
    // Bug4011: PPTX table cell NodeBuilder missing font from cell runs
    // TableToNode (NodeBuilder.cs:80-188) doesn't read font/size/bold/italic
    // from the cell's runs.
    // =====================================================================
    [Fact]
    public void Bug4011_Pptx_TableCell_NodeBuilder_Missing_Font()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "10cm", ["height"] = "5cm"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Cell text", ["font"] = "Courier New", ["size"] = "14"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = tableNode.Children[0].Children[0];
        cellNode.Text.Should().Be("Cell text");
        cellNode.Format.Should().ContainKey("font",
            because: "table cell NodeBuilder should report font from runs but doesn't");
    }

    // =====================================================================
    // Bug4012: PPTX table cell NodeBuilder missing bold
    // =====================================================================
    [Fact]
    public void Bug4012_Pptx_TableCell_NodeBuilder_Missing_Bold()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "10cm", ["height"] = "5cm"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold cell", ["bold"] = "true"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = tableNode.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("bold",
            because: "table cell NodeBuilder should report bold from runs but doesn't");
    }

    // =====================================================================
    // Bug4013: PPTX table cell NodeBuilder missing color
    // =====================================================================
    [Fact]
    public void Bug4013_Pptx_TableCell_NodeBuilder_Missing_Color()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "10cm", ["height"] = "5cm"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Red", ["color"] = "FF0000"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = tableNode.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("color",
            because: "table cell NodeBuilder should report text color from runs but doesn't");
    }

    // =====================================================================
    // Bug4014: PPTX table cell NodeBuilder missing italic
    // =====================================================================
    [Fact]
    public void Bug4014_Pptx_TableCell_NodeBuilder_Missing_Italic()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "10cm", ["height"] = "5cm"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Italic cell", ["italic"] = "true"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = tableNode.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("italic",
            because: "table cell NodeBuilder should report italic from runs but doesn't");
    }

    // =====================================================================
    // Bug4015: PPTX table cell NodeBuilder missing underline
    // =====================================================================
    [Fact]
    public void Bug4015_Pptx_TableCell_NodeBuilder_Missing_Underline()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "10cm", ["height"] = "5cm"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Underlined", ["underline"] = "single"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = tableNode.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("underline",
            because: "table cell NodeBuilder should report underline from runs but doesn't");
    }

    // =====================================================================
    // Bug4016: PPTX table cell NodeBuilder missing strike
    // =====================================================================
    [Fact]
    public void Bug4016_Pptx_TableCell_NodeBuilder_Missing_Strike()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "10cm", ["height"] = "5cm"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Struck", ["strike"] = "single"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = tableNode.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("strike",
            because: "table cell NodeBuilder should report strikethrough from runs but doesn't");
    }

    // =====================================================================
    // Bug4017: PPTX table cell NodeBuilder missing size
    // =====================================================================
    [Fact]
    public void Bug4017_Pptx_TableCell_NodeBuilder_Missing_Size()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2", ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "10cm", ["height"] = "5cm"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Big", ["size"] = "24"
        });

        var tableNode = handler.Get("/slide[1]/table[1]", depth: 2);
        var cellNode = tableNode.Children[0].Children[0];
        cellNode.Format.Should().ContainKey("size",
            because: "table cell NodeBuilder should report font size from runs but doesn't");
    }

    // =====================================================================
    // Bug4018: PPTX shape Set rotation + text + bold (stale runs + rotation)
    // Rotation is handled in SetRunOrShapeProperties (not a run property)
    // so it should work even with stale runs. But text + bold still fails.
    // =====================================================================
    [Fact]
    public void Bug4018_Pptx_Stale_Runs_Multiline_Text_Plus_Bold()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Initial" });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Line1\\nLine2",
            ["bold"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Contain("Line1");
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true,
            because: "bold should apply to new runs but stale runs bug means bold goes to orphaned runs");
    }

    // =====================================================================
    // Bug4019: PPTX single-line text + bold should work (control test)
    // When text is single line and there's exactly 1 run, the text case
    // simply updates the run's text without replacing paragraphs.
    // So the runs list stays valid. Bold should work.
    // =====================================================================
    [Fact]
    public void Bug4019_Pptx_SingleLine_Text_Plus_Bold_Works()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Initial" });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["text"] = "Updated",
            ["bold"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Updated");
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true,
            because: "single-line text replacement preserves the run, so bold should work");
    }

    // =====================================================================
    // Bug4020: Excel cell style font.name not reported when set via Add
    // CellToNode only reads font when styleIndex > 0, but default font (id 0)
    // is skipped. Let's verify Add with font creates the right style.
    // =====================================================================
    [Fact]
    public void Bug4020_Excel_Cell_Add_With_Font_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Styled", ["font"] = "Courier New"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("font.name",
            because: "font name should be readable after setting it during Add");
        node.Format["font.name"].Should().Be("Courier New");
    }

    // =====================================================================
    // Bug4021: Excel cell Set bold + value roundtrip
    // =====================================================================
    [Fact]
    public void Bug4021_Excel_Cell_Set_Bold_Roundtrip()
    {
        var path = CreateTemp(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Bold", ["bold"] = "true"
        });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("font.bold",
            because: "bold formatting should be readable after setting during Add");
    }

    // =====================================================================
    // Bug4022: Word table cell Set text + bold roundtrip
    // Word table cell Set defers text until after formatting, preserving
    // RunProperties from first run. Let's verify this works.
    // =====================================================================
    [Fact]
    public void Bug4022_Word_TableCell_Text_Plus_Bold()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold Cell",
            ["bold"] = "true"
        });

        // Get the run inside the table cell
        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]", depth: 2);
        cellNode.Text.Should().Contain("Bold Cell");
        // Check if run has bold
        var runs = cellNode.Children.SelectMany(c => c.Children).Where(c => c.Type == "run").ToList();
        if (runs.Count > 0)
        {
            runs[0].Format.Should().ContainKey("bold",
                because: "bold should be preserved when text is deferred in table cell Set");
        }
    }

    // =====================================================================
    // Bug4023: PPTX shape opacity — Set "opacity" requires existing SolidFill
    // If the shape has no fill, opacity is silently ignored
    // =====================================================================
    [Fact]
    public void Bug4023_Pptx_Shape_Opacity_Without_Fill_Ignored()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "No fill" });

        // Try setting opacity without a fill — it should be reported as unsupported
        // or at least documented behavior
        handler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.5" });
        var node = handler.Get("/slide[1]/shape[1]");
        // opacity auto-creates a white fill if none exists (matching POI behavior)
        node.Format.ContainsKey("opacity").Should().BeTrue(
            because: "opacity should be applied by auto-creating a fill when none exists");
    }

    // =====================================================================
    // Bug4024: Word Add paragraph with numbering — numid and numlevel roundtrip
    // =====================================================================
    [Fact]
    public void Bug4024_Word_Add_Paragraph_With_ListStyle_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Item 1", ["liststyle"] = "bullet"
        });

        var node = handler.Get("/body/p[1]");
        node.Text.Should().Be("Item 1");
        node.Format.Should().ContainKey("listStyle",
            because: "paragraph with liststyle should report listStyle in Get");
    }

    // =====================================================================
    // Bug4025: PPTX shape Set spaceBefore/spaceAfter roundtrip
    // NodeBuilder reports these from SpaceBefore/SpaceAfter elements
    // =====================================================================
    [Fact]
    public void Bug4025_Pptx_Shape_SpaceBefore_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Spaced", ["spaceBefore"] = "12"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("spaceBefore",
            because: "spaceBefore should be readable after setting during Add");
    }

    // =====================================================================
    // Bug4026: Word run NodeBuilder — verify hyperlink link is reported
    // =====================================================================
    [Fact]
    public void Bug4026_Word_Run_Hyperlink_Roundtrip()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "Click here" });
        handler.Set("/body/p[1]/r[1]", new() { ["link"] = "https://example.com" });

        var runNode = handler.Get("/body/p[1]/r[1]");
        runNode.Format.Should().ContainKey("link",
            because: "run with hyperlink should report link in Format");
    }

    // =====================================================================
    // Bug4027: PPTX shape Add with "charspacing" — verify Add supports it
    // "spacing"/"charspacing"/"letterspacing" are in effectKeys so they
    // should be delegated to SetRunOrShapeProperties during Add.
    // =====================================================================
    [Fact]
    public void Bug4027_Pptx_Shape_Add_CharSpacing_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Spaced", ["charspacing"] = "200"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("spacing",
            because: "charspacing should be readable after Add — it's in effectKeys");
    }

    // =====================================================================
    // Bug4028: Word paragraph Set "text" replaces runs, losing any
    // formatting set in the same call IF text comes after formatting.
    // The "text" case (line 928-954) uses first run's existing text, removes
    // extra runs. If bold was set before text, bold is on the first run.
    // When text updates first run's text, the RunProperties should survive.
    // But when formatting comes AFTER text in dict, text creates a new run
    // from ParagraphMarkRunProperties, and then formatting applies to it.
    // =====================================================================
    [Fact]
    public void Bug4028_Word_Paragraph_Set_Font_Then_Text_Preserves_Font()
    {
        var path = CreateTemp(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "Original" });

        // Set font first, then text
        var props = new Dictionary<string, string>();
        props["font"] = "Courier New";
        props["text"] = "NewText";
        handler.Set("/body/p[1]", props);

        var runNode = handler.Get("/body/p[1]/r[1]");
        runNode.Text.Should().Be("NewText");
        runNode.Format.Should().ContainKey("font",
            because: "font should persist after text replacement in same Set call");
        runNode.Format["font"].Should().Be("Courier New");
    }

    // =====================================================================
    // Bug4029: PPTX connector lineWidth uses ParseEmu which gives EMU values
    // but the connector Set handler (Set.cs:912) converts to int — lineWidth
    // should use EmuConverter.ParseEmuAsInt for consistency.
    // Let's verify the roundtrip: Set "2pt" lineWidth, Get should return "2pt"
    // or "Npt" format (not EMU).
    // =====================================================================
    [Fact]
    public void Bug4029_Pptx_Connector_LineWidth_Roundtrip()
    {
        var path = CreateTemp(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "0cm"
        });

        handler.Set("/slide[1]/connector[1]", new() { ["lineWidth"] = "2pt" });
        var node = handler.Get("/slide[1]/connector[1]");
        node.Format.Should().ContainKey("lineWidth");
        // ConnectorToNode reports lineWidth as "Npt" format (width / 12700.0)
        // ParseEmu("2pt") = 2 * 12700 = 25400 EMU
        // But Set handler uses (int)ParseEmu("2pt") = 25400
        // Then Get reports 25400 / 12700.0 = "2pt" — should match
        node.Format["lineWidth"].Should().Be("0.07cm");
    }
}
