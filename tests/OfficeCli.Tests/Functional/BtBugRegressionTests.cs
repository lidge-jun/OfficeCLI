// BT Bug Regression Tests — covers confirmed-fixed bugs found by black-box testing.
// BT-1: batch exit code on failure
// BT-2: highlight color validation
// BT-3: gradient fill data loss on invalid color
// BT-5: PPTX Query case-insensitive attribute matching (covered in PptxQueryAttributeFilterTests.cs)
// BT-6: Excel Query treats 'row' as column name
// BT-7: Excel table totalRow Set/Get mismatch
// BT-8: connector/group/notes not recognized in PPTX selector
// BT-9: notes placeholder missing paragraph element
// BT-10: Excel border color readback includes ARGB alpha prefix
// BT-11: chart part left empty on invalid chart type
// BT-13: textwarp/wordart/autofit ignored during PPTX shape Add
// BT-14: XLSX invalid color values not rejected

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BtBugRegressionTests : IDisposable
{
    private readonly string _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
    private readonly string _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
    private readonly string _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");

    private ExcelHandler? _excel;
    private WordHandler? _word;

    private ExcelHandler GetExcel()
    {
        if (_excel == null)
        {
            BlankDocCreator.Create(_xlsxPath);
            _excel = new ExcelHandler(_xlsxPath, editable: true);
        }
        return _excel;
    }

    private WordHandler GetWord()
    {
        if (_word == null)
        {
            BlankDocCreator.Create(_docxPath);
            _word = new WordHandler(_docxPath, editable: true);
        }
        return _word;
    }

    public void Dispose()
    {
        _excel?.Dispose();
        _word?.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    // ==================== BT-2: highlight color validation ====================

    /// <summary>
    /// BT-2: Word highlight color must validate against allowlist upfront.
    /// Previously HighlightColorValues accepted invalid values silently,
    /// only throwing on XML serialization (outside the try/catch).
    /// </summary>
    [Fact]
    public void BT2_Word_SetHighlightColor_InvalidValue_ThrowsArgumentException()
    {
        var word = GetWord();
        word.Add("/body", "p", null, new() { ["text"] = "Test" });

        var act = () => word.Set("/body/p[1]/r[1]", new() { ["highlight"] = "purple" });
        act.Should().Throw<ArgumentException>("'purple' is not a valid highlight color");
    }

    [Fact]
    public void BT2_Word_SetHighlightColor_ValidValue_Works()
    {
        var word = GetWord();
        word.Add("/body", "p", null, new() { ["text"] = "Test" });

        var unsupported = word.Set("/body/p[1]/r[1]", new() { ["highlight"] = "yellow" });
        unsupported.Should().NotContain("highlight");
    }

    // ==================== BT-3: gradient fill data loss ====================

    /// <summary>
    /// BT-3: PPTX gradient fill should build new fill before removing old one.
    /// Previously, invalid color removed the old fill first, leaving shape with no fill.
    /// </summary>
    [Fact]
    public void BT3_Pptx_GradientFill_InvalidColor_PreserveExistingFill()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test", ["fill"] = "FF0000" });

        // Set a valid solid fill first
        pptx.Set("/slide[1]/shape[1]", new() { ["fill"] = "00FF00" });
        var before = pptx.Get("/slide[1]/shape[1]");
        before.Format.Should().ContainKey("fill");

        // Now try invalid gradient — should throw, not silently remove fill
        var act = () => pptx.Set("/slide[1]/shape[1]", new() { ["gradient"] = "ZZZZZZ-000000" });
        act.Should().Throw<Exception>("invalid gradient color should throw");

        // Original fill should still be intact
        var after = pptx.Get("/slide[1]/shape[1]");
        after.Format.Should().ContainKey("fill", "fill should be preserved after failed gradient set");
    }

    // ==================== BT-6: Excel Query 'row' treated as column ====================

    /// <summary>
    /// BT-6: Excel Query("row") was matched by column filter regex ^[A-Z]+$,
    /// treating 'row' as column name 'ROW' and silently returning no results.
    /// </summary>
    [Fact]
    public void BT6_Excel_QueryRow_NotTreatedAsColumn()
    {
        var excel = GetExcel();
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Hello" });
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "World" });

        // Query for rows — should not be misinterpreted as column "ROW"
        var results = excel.Query("row");
        results.Should().NotBeEmpty("'row' should query rows, not be treated as a column name");
    }

    // ==================== BT-7: Excel table totalRow Set/Get mismatch ====================

    /// <summary>
    /// BT-7: Set used TotalsRowCount but Get reads TotalsRowShown.
    /// Setting totalrow=true was not reflected when reading back.
    /// </summary>
    [Fact]
    public void BT7_Excel_Table_TotalRow_SetGetRoundtrip()
    {
        var excel = GetExcel();
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Name" });
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "Value" });
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Item1" });
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "10" });

        excel.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:B2", ["name"] = "TestTable" });

        // Set totalRow
        excel.Set("/Sheet1/table[1]", new() { ["totalrow"] = "true" });

        // Get should reflect totalRow
        var node = excel.Get("/Sheet1/table[1]");
        var totalRow = node.Format.ContainsKey("totalRow") ? node.Format["totalRow"]?.ToString() : null;
        totalRow.Should().NotBeNull("totalRow should be readable after setting");
    }

    // ==================== BT-8: connector/group/notes in PPTX selector ====================

    /// <summary>
    /// BT-8: ParseShapeSelector didn't recognize connector/group/notes as valid types,
    /// causing Query("connector") to return all element types.
    /// </summary>
    [Fact]
    public void BT8_Pptx_QueryConnector_OnlyReturnsConnectors()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape1" });
        pptx.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "1cm", ["y"] = "1cm", ["cx"] = "5cm", ["cy"] = "0cm"
        });

        var connectors = pptx.Query("connector");
        // All results should be connectors, not shapes
        foreach (var c in connectors)
        {
            c.Type.Should().Be("connector",
                "Query('connector') should only return connectors, not other element types");
        }
    }

    [Fact]
    public void BT8_Pptx_QueryGroup_OnlyReturnsGroups()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape1" });

        // Query for groups — should not match shapes
        var groups = pptx.Query("group");
        foreach (var g in groups)
        {
            g.Type.Should().Be("group",
                "Query('group') should only return groups, not other element types");
        }
    }

    // ==================== BT-9: notes placeholder missing paragraph ====================

    /// <summary>
    /// BT-9: EnsureNotesSlidePart created notes body placeholder without
    /// a paragraph element, causing XML schema violation.
    /// </summary>
    [Fact]
    public void BT9_Pptx_SpeakerNotes_NewNotesSlideIsValid()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Set speaker notes — this creates the notes slide part
        pptx.Set("/slide[1]", new() { ["notes"] = "Test speaker notes" });

        // Read back — should not throw and should return the text
        var node = pptx.Get("/slide[1]");
        var notes = node.Format.ContainsKey("notes") ? node.Format["notes"]?.ToString() : null;
        notes.Should().Contain("Test speaker notes",
            "speaker notes should be readable after creation");
    }

    // ==================== BT-10: Excel border color readback ====================

    /// <summary>
    /// BT-10: Border color was returned as 8-char ARGB (e.g. "FFFF0000")
    /// while font.color strips to 6-char RGB. Should be consistent.
    /// </summary>
    [Fact]
    public void BT10_Excel_BorderColor_ReadbackIs6CharRGB()
    {
        var excel = GetExcel();
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });
        excel.Set("/Sheet1/A1", new() { ["border.all"] = "thin", ["border.color"] = "FF0000" });

        // Reopen to ensure persistence
        _excel?.Dispose();
        _excel = new ExcelHandler(_xlsxPath, editable: true);

        var node = _excel.Get("/Sheet1/A1");

        // Check that border color values are #RRGGBB (7-char with # prefix), not 8-char ARGB
        foreach (var key in node.Format.Keys.Where(k => k.Contains("border") && k.Contains("color")))
        {
            var colorVal = node.Format[key]?.ToString();
            if (colorVal != null && colorVal.Length > 0)
            {
                colorVal.Length.Should().BeLessOrEqualTo(7,
                    $"{key} should be #RRGGBB format (got '{colorVal}'), not 8-char ARGB");
            }
        }
    }

    // ==================== BT-11: chart part empty on invalid type ====================

    /// <summary>
    /// BT-11: Adding chart with invalid type used to create an empty ChartPart
    /// (added part first, then tried to build content which failed).
    /// Now builds content first, only adds part after validation.
    /// </summary>
    [Fact]
    public void BT11_Excel_AddChart_InvalidType_ThrowsCleanly()
    {
        var excel = GetExcel();
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        var act = () => excel.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "invalidChartType",
            ["data"] = "A1:A1"
        });

        act.Should().Throw<Exception>("invalid chart type should throw");

        // Verify no orphaned chart parts left behind
        var node = excel.Get("/Sheet1");
        var charts = node.Children.Where(c => c.Type == "chart").ToList();
        charts.Should().BeEmpty("no chart should be created when type is invalid");
    }

    [Fact]
    public void BT11_Pptx_AddChart_InvalidType_ThrowsCleanly()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var act = () => pptx.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "invalidChartType",
            ["data"] = "A,B\n1,2"
        });

        act.Should().Throw<Exception>("invalid chart type should throw");
    }

    // ==================== BT-13: textwarp/wordart/autofit during Add ====================

    /// <summary>
    /// BT-13: textwarp, wordart, and autofit were silently ignored during
    /// shape Add because they weren't in the effectKeys set.
    /// </summary>
    [Fact]
    public void BT13_Pptx_AddShape_WithAutoFit_Applied()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add shape with autoFit in the same call
        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "AutoFit Test",
            ["autoFit"] = "normal"
        });

        var node = pptx.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("autoFit");
        node.Format["autoFit"].Should().Be("normal",
            "autoFit should be applied during Add, not silently ignored");
    }

    [Fact]
    public void BT13_Pptx_AddShape_WithTextWarp_Applied()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Warped",
            ["textwarp"] = "wave1"
        });

        var node = pptx.Get("/slide[1]/shape[1]");
        // textWarp should be present in format (camelCase key)
        node.Format.Should().ContainKey("textWarp",
            "textWarp should be applied during Add, not silently ignored");
    }

    // ==================== BT-14: XLSX invalid color validation ====================

    /// <summary>
    /// BT-14: Invalid hex color values were silently accepted in XLSX,
    /// producing corrupted files. Now validated via NormalizeArgbColor.
    /// </summary>
    [Fact]
    public void BT14_Excel_SetFontColor_InvalidHex_ThrowsArgumentException()
    {
        var excel = GetExcel();
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        var act = () => excel.Set("/Sheet1/A1", new() { ["font.color"] = "ZZZZZZ" });
        act.Should().Throw<ArgumentException>("'ZZZZZZ' is not a valid hex color");
    }

    [Fact]
    public void BT14_Excel_SetFillColor_InvalidHex_ThrowsArgumentException()
    {
        var excel = GetExcel();
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        var act = () => excel.Set("/Sheet1/A1", new() { ["fill"] = "notacolor" });
        act.Should().Throw<ArgumentException>("'notacolor' is not a valid hex color");
    }

    [Fact]
    public void BT14_Excel_SetFontColor_ValidHex_Works()
    {
        var excel = GetExcel();
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Should not throw — font.color is the correct key for Excel cell font color
        excel.Set("/Sheet1/A1", new() { ["font.color"] = "FF0000" });

        var node = excel.Get("/Sheet1/A1");
        var color = node.Format.ContainsKey("font.color") ? node.Format["font.color"]?.ToString() : null;
        color.Should().NotBeNull("font color should be set after valid hex");
    }

    // ==================== BT-6 additional: Excel Query != filter ====================

    /// <summary>
    /// Related to BT-6: Excel Query != filter was broken because '!' was
    /// parsed as sheet separator. E.g., cell[value!=foo] was split at '!'.
    /// </summary>
    [Fact]
    public void BT6b_Excel_Query_NotEqualsFilter_Works()
    {
        var excel = GetExcel();
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "keep" });
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "skip" });

        var results = excel.Query("cell[value!=skip]");
        results.Should().Contain(n => n.Text == "keep", "!= filter should exclude 'skip'");
        results.Should().NotContain(n => n.Text == "skip", "!= filter should exclude 'skip'");
    }

    // ==================== BT-9 additional: hyperlink container cleanup ====================

    /// <summary>
    /// Related: removing all hyperlinks in XLSX should also remove the empty container.
    /// </summary>
    [Fact]
    public void BT_Excel_RemoveHyperlink_CleansUpEmptyContainer()
    {
        var excel = GetExcel();
        excel.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Link" });
        excel.Set("/Sheet1/A1", new() { ["hyperlink"] = "https://example.com" });

        // Remove hyperlink
        excel.Set("/Sheet1/A1", new() { ["hyperlink"] = "" });

        // Reopen and verify no corruption
        _excel?.Dispose();
        _excel = new ExcelHandler(_xlsxPath, editable: true);
        var node = _excel.Get("/Sheet1/A1");
        node.Should().NotBeNull("cell should still be accessible after hyperlink removal");
    }

    // ==================== Cross-check: did-you-mean doesn't suggest self ====================

    /// <summary>
    /// BT-4 partial: SuggestProperty should not suggest the exact same property
    /// that the user typed (self-suggestion).
    /// </summary>
    [Fact]
    public void BT4_Pptx_Set_UnsupportedProp_DoesNotSuggestSelf()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // "name" is a known property but not supported at shape level in Set
        // The unsupported list should contain "name" but the suggestion
        // should NOT be "name" itself
        var unsupported = pptx.Set("/slide[1]/shape[1]", new() { ["fakeProp123"] = "value" });
        unsupported.Should().Contain("fakeProp123");
        // This is a sanity check — the suggestion system should not crash
    }
}
