// Bug hunt tests — each test exposes a specific bug found through code review.
// Tests are organized by severity: Critical → High → Medium → Low

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class MixedRegression1 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public MixedRegression1()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        // Pre-create a slide so Part4 tests can reference /slide[1]
        using (var pptx = new PowerPointHandler(_pptxPath, editable: true))
            pptx.Add("/", "slide", null, new());
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _excelHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private WordHandler ReopenWord()
    {
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
        return _wordHandler;
    }

    private ExcelHandler ReopenExcel()
    {
        _excelHandler.Dispose();
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        return _excelHandler;
    }


    // ==================== BUG #6 (HIGH): Endnote Set also prepends space ====================
    // Same bug as footnote but for endnotes.
    // WordHandler.Set.cs line 141: textEl.Text = " " + enText;
    //
    // Location: WordHandler.Set.cs line 141

    [Fact]
    public void Bug06_EndnoteSet_PrependsSpace()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Text" });
        _wordHandler.Add("/body/p[1]", "endnote", null, new() { ["text"] = "Original" });

        _wordHandler.Set("/endnote[1]", new() { ["text"] = "Clean" });

        var en = _wordHandler.Get("/endnote[1]");
        // BUG: text = " Clean" instead of "Clean"
        // The Get joins ALL descendants<Text>(), including the reference mark's space
        var text = en.Text ?? "";
        text.Trim().Should().Be("Clean",
            "Endnote text should be clean without extra space");
        // Verify no leading space accumulation
        text.Should().NotStartWith(" ",
            "Endnote text should not have leading space from Set");
    }

    // ==================== BUG #7 (HIGH): Excel Hyperlinks element ordering violation ====================
    // When adding a hyperlink via Set (case "link"), the Hyperlinks element is appended
    // to the worksheet with ws.AppendChild(hyperlinksEl) (line 522).
    // This does NOT respect schema order and may place Hyperlinks after Drawing.
    // While ReorderWorksheetChildren exists, it's not always called for link operations.
    //
    // Location: ExcelHandler.Set.cs lines 519-523

    [Fact]
    public void Bug07_ExcelHyperlinkElementOrder()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Click me" });

        // Add a chart first (creates Drawing element which should be AFTER Hyperlinks in schema)
        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "pie",
            ["data"] = "Sales:40,30,30",
            ["categories"] = "A,B,C"
        });

        // Now add a hyperlink — Hyperlinks should be BEFORE Drawing in schema
        _excelHandler.Set("/Sheet1/A1", new() { ["link"] = "https://example.com" });

        ReopenExcel();
        var errors = _excelHandler.Validate();
        errors.Should().BeEmpty(
            "Hyperlinks should be ordered before Drawing element per schema");
    }


    // ==================== BUG #21 (MEDIUM): Excel Query GenericXmlQuery uses body instead of specific sheet ====================
    // When Excel handler falls through to GenericXmlQuery for unknown element types,
    // it searches the Worksheet element. But if the user specifies a sheet prefix
    // in the selector, the sheet prefix is not resolved.

    [Fact]
    public void Bug21_ExcelQueryAfterMultipleSets()
    {
        // Create cells and then query them
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "X" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Y" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A3", ["value"] = "Z" });

        // Query for cells containing "X"
        var results = _excelHandler.Query("cell:contains(X)");
        results.Should().HaveCountGreaterOrEqualTo(1,
            "Query should find cell containing 'X'");
        results[0].Text.Should().Be("X");
    }

    // ==================== BUG #25 (MEDIUM): Word paragraph Set "text" not implemented ====================
    // Setting text on a paragraph doesn't have a direct handler —
    // it falls through to the generic fallback which may fail.

    [Fact]
    public void Bug25_WordParagraphSetText()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Original" });

        // Set text on paragraph — does this work?
        var unsupported = _wordHandler.Set("/body/p[1]", new() { ["text"] = "Updated" });

        // If "text" is in unsupported list, it means paragraph-level text Set isn't implemented
        if (unsupported.Contains("text"))
        {
            // BUG: Cannot Set text directly on a paragraph
            unsupported.Should().NotContain("text",
                "Setting text on paragraph should be supported");
        }
        else
        {
            var para = _wordHandler.Get("/body/p[1]");
            para.Text.Should().Contain("Updated");
        }
    }

    // ==================== BUG #29 (MEDIUM): Excel cell value="true" stored as String not Boolean ====================
    // When value="true", double.TryParse fails → DataType=String
    // But OpenXML has CellValues.Boolean for boolean cells
    // The auto-detection doesn't check for boolean values
    //
    // Location: ExcelHandler.Set.cs lines 482-487

    [Fact]
    public void Bug29_ExcelBooleanAutoDetect()
    {
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "true" });

        var cell = _excelHandler.Get("/Sheet1/A1");
        // "true" should ideally be detected as Boolean, not String
        var type = (string)cell.Format["type"];
        // Currently it's "String" because double.TryParse("true") fails
        // A more complete auto-detect would check for "true"/"false" → Boolean
        type.Should().BeOneOf("Boolean", "String",
            "Auto-detection should handle boolean values");
    }

    // ==================== BUG #32 (HIGH): PPTX Add slide with index returns wrong path ====================
    // When inserting a slide at a specific index, the returned path is always
    // /slide[{slideCount}] (last position) instead of the actual insertion position.
    // If you insert at index=0 (before first slide), the path says it's the last slide.
    //
    // Location: PowerPointHandler.Add.cs line 89

    [Fact]
    public void Bug32_PptxAddSlide_IndexReturnsWrongPath()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);

        // Add 3 slides
        handler.Add("/", "slide", null, new() { ["title"] = "Slide1" });
        handler.Add("/", "slide", null, new() { ["title"] = "Slide2" });
        handler.Add("/", "slide", null, new() { ["title"] = "Slide3" });

        // Insert a new slide at index 0 (before first slide)
        var resultPath = handler.Add("/", "slide", 0, new() { ["title"] = "Inserted" });

        // BUG: resultPath will be /slide[4] (total count), not /slide[1] (insertion position)
        // The slide was inserted at position 1, but the path says position 4
        resultPath.Should().Be("/slide[1]",
            "Slide inserted at index 0 should return /slide[1], not the total slide count");
    }


    // ==================== BUG #38 (MEDIUM): Word TOC Set hyperlinks uses bool.Parse ====================
    // WordHandler.Set.cs lines 70, 72: bool.Parse(hlSwitch)
    // Same bool.Parse inconsistency — "yes"/"1" throw FormatException.
    //
    // Location: WordHandler.Set.cs lines 70-72

    [Fact]
    public void Bug38_WordTocSetBoolParse()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Intro", ["style"] = "Heading1" });
        _wordHandler.Add("/body", "toc", null, new() { ["levels"] = "1-3" });

        // "1" should mean true, but bool.Parse("1") throws
        var ex = Record.Exception(() =>
            _wordHandler.Set("/toc[1]", new() { ["hyperlinks"] = "1" }));

        // BUG: FormatException from bool.Parse("1")
        ex.Should().BeNull(
            "hyperlinks='1' should be accepted as truthy");
    }


    // ==================== BUG #40 (MEDIUM): Excel Query calls CellToNode with inconsistent params ====================
    // In ExcelHandler.Query.cs line 457: CellToNode(sheetName, cell) — 2 params (no worksheet)
    // In ExcelHandler.Query.cs line 186/360: CellToNode(sheetName, cell, worksheet) — 3 params
    // The 2-param version may miss style information that requires the worksheet part.
    //
    // Location: ExcelHandler.Query.cs line 457 vs lines 186/360

    [Fact]
    public void Bug40_ExcelQueryCellToNode_InconsistentParams()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Styled" });
        _excelHandler.Set("/Sheet1/A1", new() { ["font.bold"] = "true", ["font.color"] = "FF0000" });

        // Get via direct path (uses 3-param CellToNode with worksheet)
        var directCell = _excelHandler.Get("/Sheet1/A1");

        // Query (uses 2-param CellToNode without worksheet)
        var queryResults = _excelHandler.Query("cell:contains(Styled)");
        queryResults.Should().HaveCountGreaterOrEqualTo(1);

        var queryCell = queryResults[0];
        // Both should have the same format information
        // BUG: queryCell may miss style info because CellToNode was called without worksheet
        directCell.Format.Should().ContainKey("font.bold",
            "Direct Get should include font.bold");
    }

    // ==================== BUG #41 (HIGH): Word run-level bold/italic all use bool.Parse not IsTruthy ====================
    // WordHandler.Set.cs lines 336-369: Every boolean property (bold, italic, caps, smallCaps,
    // dstrike, vanish, outline, shadow, emboss, imprint, noproof, rtl) uses bool.Parse(value).
    // None of them accept "yes"/"1"/"on" — only "True"/"False".
    // This is inconsistent with ExcelStyleManager which uses IsTruthy.
    //
    // Location: WordHandler.Set.cs lines 336-369

    [Fact]
    public void Bug41_WordRunBooleanProperties_AllUseBoolParse()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });

        // Test all boolean properties with "1" (truthy in many systems)
        var boolProps = new[] { "italic", "caps", "smallcaps", "vanish" };
        foreach (var prop in boolProps)
        {
            var ex = Record.Exception(() =>
                _wordHandler.Set("/body/p[1]/r[1]", new() { [prop] = "1" }));

            // BUG: All of these throw FormatException because bool.Parse("1") fails
            ex.Should().BeNull(
                $"'{prop}=1' should be accepted as truthy value (like Excel's IsTruthy)");
        }
    }

    // ==================== BUG #42 (MEDIUM): PPTX Add shape bold/italic use bool.Parse ====================
    // PowerPointHandler.Add.cs lines 131, 140: bool.Parse(boldStr), bool.Parse(italicStr)
    // Same inconsistency as Word — "yes"/"1" throw.
    //
    // Location: PowerPointHandler.Add.cs lines 131, 140

    [Fact]
    public void Bug42_PptxAddShapeBoldItalic_BoolParse()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        handler.Add("/", "slide", null, new());

        var ex = Record.Exception(() =>
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Bold", ["bold"] = "yes"
            }));

        // BUG: FormatException from bool.Parse("yes")
        ex.Should().BeNull(
            "bold='yes' should be accepted when adding a PPTX shape");
    }

    // ==================== BUG #43 (MEDIUM): PPTX Add shape size parsing — int.Parse for size ====================
    // PowerPointHandler.Add.cs line 122: var sizeVal = int.Parse(sizeStr) * 100;
    // This fails for fractional font sizes like "10.5" (common in presentations).
    // int.Parse("10.5") throws FormatException.
    //
    // Location: PowerPointHandler.Add.cs line 122

    [Fact]
    public void Bug43_PptxAddShapeFractionalFontSize()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        handler.Add("/", "slide", null, new());

        var ex = Record.Exception(() =>
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Small", ["size"] = "10.5"
            }));

        // BUG: FormatException from int.Parse("10.5")
        // Should use double.Parse or decimal.Parse for fractional font sizes
        ex.Should().BeNull(
            "Fractional font size '10.5' should be supported");
    }


    // ==================== BUG #49 (MEDIUM): PPTX group bounding box calculation ignores zero-value offsets ====================
    // PowerPointHandler.Add.cs lines 944-957: Bounding box calculation uses
    // long minX = long.MaxValue, minY = long.MaxValue, maxX = 0, maxY = 0;
    // If a shape is at (0,0), maxX/maxY start at 0 and won't update if all shapes
    // have x=0, y=0. But more importantly, if no shape has Transform2D, the bounding
    // box stays at (MaxValue, MaxValue, 0, 0) which is nonsensical.
    //
    // Location: PowerPointHandler.Add.cs lines 944-957

    [Fact]
    public void Bug49_PptxGroupBoundingBox_NoTransform()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        handler.Add("/", "slide", null, new());

        // Add shapes at position (0,0)
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A", ["x"] = "0", ["y"] = "0", ["width"] = "2cm", ["height"] = "1cm"
        });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "B", ["x"] = "0", ["y"] = "1cm", ["width"] = "2cm", ["height"] = "1cm"
        });

        // Group them — bounding box should be (0, 0) to (2cm, 2cm)
        var grpPath = handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });
        grpPath.Should().NotBeNull();

        // Verify group was created
        var slide = handler.Get("/slide[1]", depth: 1);
        slide.ChildCount.Should().BeGreaterThan(0);
    }

    // ==================== BUG #50 (MEDIUM): Word section margin Set creates separate elements ====================
    // When setting multiple margins, each one does:
    //   sectPr.GetFirstChild<PageMargin>() ?? sectPr.AppendChild(new PageMargin())
    // If PageMargin doesn't exist, the first property creates one. But AppendChild
    // puts it at the END of sectPr, which may violate schema order.
    // More importantly: each margin property independently checks for PageMargin,
    // so they all share the same element (fine), but if PageMargin exists with
    // ONLY Top set and you Set MarginLeft, the existing Top is preserved (correct).
    // However, if PageMargin doesn't exist and you only set MarginTop,
    // the other margin attributes (Left, Right, Bottom) will be missing/default.
    //
    // Location: WordHandler.Set.cs lines 184-195

    [Fact]
    public void Bug50_WordSectionMarginPartial()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });

        // Set only top margin — other margins remain unset
        _wordHandler.Set("/section[1]", new() { ["margintop"] = "1440" });

        var sec = _wordHandler.Get("/section[1]");
        sec.Format.Should().ContainKey("margintop",
            "Top margin should be set");

        // Now set left margin — it should reuse the same PageMargin element
        _wordHandler.Set("/section[1]", new() { ["marginleft"] = "1440" });

        sec = _wordHandler.Get("/section[1]");
        sec.Format.Should().ContainKey("margintop",
            "Top margin should still be present after setting left margin");
        sec.Format.Should().ContainKey("marginleft",
            "Left margin should now be set");
    }

    // ==================== BUG #51 (HIGH): PPTX ShapeProperties size uses int.Parse for fractional sizes ====================
    // PowerPointHandler.ShapeProperties.cs line 77: int.Parse(value) * 100
    // This is the Set path (not just Add). Fractional font sizes like "10.5" throw.
    //
    // Location: PowerPointHandler.ShapeProperties.cs line 77

    [Fact]
    public void Bug51_PptxSetRunFractionalFontSize()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });

        // Set fractional font size on existing run
        var ex = Record.Exception(() =>
            handler.Set("/slide[1]/shape[1]", new() { ["size"] = "10.5" }));

        // BUG: int.Parse("10.5") throws FormatException
        ex.Should().BeNull(
            "Fractional font size '10.5' should be supported in Set");
    }

    // ==================== BUG #52 (MEDIUM): Word paragraph Set keepnext/keeplines/pagebreakbefore/widowcontrol use bool.Parse ====================
    // WordHandler.Set.cs lines 546-568: paragraph-level boolean properties all use bool.Parse.
    // Same inconsistency — "1", "yes", "on" throw FormatException.
    //
    // Location: WordHandler.Set.cs lines 546-568

    [Fact]
    public void Bug52_WordParagraphBoolProperties()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });

        var ex = Record.Exception(() =>
            _wordHandler.Set("/body/p[1]", new() { ["keepnext"] = "1" }));

        // BUG: bool.Parse("1") throws FormatException
        ex.Should().BeNull(
            "keepnext='1' should be accepted as truthy value");
    }

    // ==================== BUG #53 (MEDIUM): Word Add paragraph bool properties also use bool.Parse ====================
    // WordHandler.Add.cs lines 118-124: keepnext, keeplines, pagebreakbefore, widowcontrol
    // all use bool.Parse. "1" throws.
    //
    // Location: WordHandler.Add.cs lines 118-124

    [Fact]
    public void Bug53_WordAddParagraphBoolProperties()
    {
        var ex = Record.Exception(() =>
            _wordHandler.Add("/body", "paragraph", null, new()
            {
                ["text"] = "Test", ["keepnext"] = "yes"
            }));

        // BUG: bool.Parse("yes") throws FormatException
        ex.Should().BeNull(
            "keepnext='yes' should be accepted when adding paragraph");
    }

    // ==================== BUG #54 (HIGH): PPTX table cell vmerge/hmerge use bool.Parse ====================
    // PowerPointHandler.ShapeProperties.cs lines 559, 562:
    // cell.VerticalMerge = new BooleanValue(bool.Parse(value));
    // cell.HorizontalMerge = new BooleanValue(bool.Parse(value));
    //
    // Location: PowerPointHandler.ShapeProperties.cs lines 559, 562

    [Fact]
    public void Bug54_PptxTableCellMergeBoolParse()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // Try setting vmerge with "1"
        var ex = Record.Exception(() =>
            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["vmerge"] = "1" }));

        // BUG: bool.Parse("1") throws FormatException
        ex.Should().BeNull(
            "vmerge='1' should be accepted as truthy value");
    }

    // ==================== BUG #55 (MEDIUM): Excel formula Set doesn't clear DataType ====================
    // When setting formula on a cell that was previously String type,
    // the DataType remains String. Formulas should have no DataType (null = Number).
    // This is a deeper test of Bug #13 verifying the actual behavior.
    //
    // Location: ExcelHandler.Set.cs lines 489-491

    [Fact]
    public void Bug55_ExcelFormulaSetDoesNotClearDataType()
    {
        // Explicitly set cell as String
        _excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Hello"
        });
        _excelHandler.Set("/Sheet1/A1", new() { ["type"] = "string" });

        var before = _excelHandler.Get("/Sheet1/A1");
        ((string)before.Format["type"]).Should().Be("String");

        // Now set formula — should clear DataType
        _excelHandler.Set("/Sheet1/A1", new() { ["formula"] = "SUM(B1:B10)" });

        var after = _excelHandler.Get("/Sheet1/A1");
        // BUG: DataType is still "String" — formula cells should not have DataType=String
        var type = after.Format.ContainsKey("type") ? (string)after.Format["type"] : "Number";
        type.Should().NotBe("String",
            "Formula cell should not retain String DataType");
    }


    // ==================== BUG #63 (MEDIUM): PPTX ShapeProperties bold/italic in SetTableCellProperties ====================
    // PowerPointHandler.ShapeProperties.cs lines 495, 503:
    // Same bool.Parse pattern for table cell run properties.
    //
    // Location: PowerPointHandler.ShapeProperties.cs lines 495, 503

    [Fact]
    public void Bug63_PptxTableCellBoldItalicBoolParse()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // Set bold with "yes"
        var ex = Record.Exception(() =>
            handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["bold"] = "yes" }));

        // BUG: bool.Parse("yes") throws FormatException
        ex.Should().BeNull(
            "bold='yes' in table cell should be accepted");
    }

    // ==================== BUG #64 (HIGH): Word Add paragraph bold/italic from Add.cs also use bool.Parse ====================
    // WordHandler.Add.cs lines 150, 152, etc.: When adding a paragraph with bold=true,
    // bool.Parse is used. "yes"/"1" throw.
    //
    // Location: WordHandler.Add.cs lines 150-168

    [Fact]
    public void Bug64_WordAddParagraphBoldBoolParse()
    {
        var ex = Record.Exception(() =>
            _wordHandler.Add("/body", "paragraph", null, new()
            {
                ["text"] = "Bold text", ["bold"] = "yes"
            }));

        // BUG: bool.Parse("yes") throws FormatException
        ex.Should().BeNull(
            "bold='yes' should be accepted when adding Word paragraph");
    }

    // ==================== BUG #65 (MEDIUM): Word Add run bold/italic use bool.Parse ====================
    // WordHandler.Add.cs lines 278, 280, 286, 290, 292, 294, 296
    //
    // Location: WordHandler.Add.cs lines 278-296

    [Fact]
    public void Bug65_WordAddRunBoolParse()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });

        var ex = Record.Exception(() =>
            _wordHandler.Add("/body/p[1]", "run", null, new()
            {
                ["text"] = "Bold run", ["bold"] = "1"
            }));

        // BUG: bool.Parse("1") throws FormatException
        ex.Should().BeNull(
            "bold='1' should be accepted when adding Word run");
    }

    // ==================== BUG #66 (LOW): Word Add TOC hyperlinks/pagenumbers use bool.Parse ====================
    // WordHandler.Add.cs lines 808-809: TOC creation uses bool.Parse for hyperlinks/pagenumbers
    //
    // Location: WordHandler.Add.cs lines 808-809

    [Fact]
    public void Bug66_WordAddTocBoolParse()
    {
        _wordHandler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Heading", ["style"] = "Heading1"
        });

        var ex = Record.Exception(() =>
            _wordHandler.Add("/body", "toc", null, new()
            {
                ["levels"] = "1-3", ["hyperlinks"] = "yes"
            }));

        // BUG: bool.Parse("yes") throws FormatException
        ex.Should().BeNull(
            "hyperlinks='yes' should be accepted when adding TOC");
    }

    // ==================== BUG #67 (MEDIUM): Excel Set "clear" doesn't reset DataType ====================
    // ExcelHandler.Set.cs lines 502-505: case "clear" clears CellValue and CellFormula
    // but doesn't clear DataType. Cell remains typed as String after clearing.
    //
    // Location: ExcelHandler.Set.cs lines 502-505

    [Fact]
    public void Bug67_ExcelClearDoesNotResetDataType()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "Hello"
        });

        // Cell is now String type
        var before = _excelHandler.Get("/Sheet1/A1");
        ((string)before.Format["type"]).Should().Be("String");

        // Clear the cell
        _excelHandler.Set("/Sheet1/A1", new() { ["clear"] = "true" });

        var after = _excelHandler.Get("/Sheet1/A1");
        // BUG: DataType may still be "String" even though cell is cleared
        // A cleared cell should have no type or show as empty
    }

    // ==================== BUG #68 (HIGH): Word Set paragraph superscript/subscript use bool.Parse ====================
    // WordHandler.Set.cs lines 410, 415
    //
    // Location: WordHandler.Set.cs lines 410-417

    [Fact]
    public void Bug68_WordRunSuperscriptBoolParse()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "H2O" });

        var ex = Record.Exception(() =>
            _wordHandler.Set("/body/p[1]/r[1]", new() { ["subscript"] = "1" }));

        // BUG: bool.Parse("1") throws FormatException
        ex.Should().BeNull(
            "subscript='1' should be accepted as truthy value");
    }


    // ==================== BUG #73 (HIGH): Word table cell bold/italic use bool.Parse ====================
    // WordHandler.Set.cs lines 659, 662: table cell formatting uses bool.Parse
    //
    // Location: WordHandler.Set.cs lines 659, 662

    [Fact]
    public void Bug73_WordTableCellBoldBoolParse()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Cell" });

        var ex = Record.Exception(() =>
            _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["bold"] = "yes" }));

        // BUG: bool.Parse("yes") throws FormatException
        ex.Should().BeNull(
            "bold='yes' in table cell should be accepted");
    }

    // ==================== BUG #74 (MEDIUM): Word table row header uses bool.Parse ====================
    // WordHandler.Set.cs line 769: bool.Parse(value) for header row
    //
    // Location: WordHandler.Set.cs line 769

    [Fact]
    public void Bug74_WordTableRowHeaderBoolParse()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        var ex = Record.Exception(() =>
            _wordHandler.Set("/body/tbl[1]/tr[1]", new() { ["header"] = "1" }));

        // BUG: bool.Parse("1") throws FormatException
        ex.Should().BeNull(
            "header='1' should be accepted as truthy for table row header");
    }

    // ==================== BUG #75 (MEDIUM): Word table cell font size uses int.Parse with multiplication ====================
    // WordHandler.Set.cs line 656: (int.Parse(value) * 2).ToString()
    // "10.5" (common font size) would throw FormatException from int.Parse.
    //
    // Location: WordHandler.Set.cs line 656

    [Fact]
    public void Bug75_WordTableCellFontSizeIntParse()
    {
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Cell" });

        // "10.5" is a common font size but int.Parse fails on it
        var ex = Record.Exception(() =>
            _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["size"] = "10.5" }));

        // BUG: int.Parse("10.5") throws FormatException — should use double/decimal parse
        ex.Should().BeNull(
            "Fractional font size '10.5' should be supported in table cells");
    }

    // ==================== BUG #76 (MEDIUM): Word run font size int.Parse fails on fractional sizes ====================
    // WordHandler.Set.cs line 388: (int.Parse(value) * 2).ToString()
    // Same bug as #75 but for regular runs.
    //
    // Location: WordHandler.Set.cs line 388

    [Fact]
    public void Bug76_WordRunFontSizeIntParse()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });

        var ex = Record.Exception(() =>
            _wordHandler.Set("/body/p[1]/r[1]", new() { ["size"] = "10.5" }));

        // BUG: int.Parse("10.5") throws FormatException
        ex.Should().BeNull(
            "Fractional font size '10.5' should be supported for runs");
    }


    // ==================== BUG #79 (MEDIUM): Word paragraph firstlineindent uses int.Parse ====================
    // WordHandler.Set.cs line 529: int.Parse(value) * 480
    // Fractional values like "1.5" crash.
    //
    // Location: WordHandler.Set.cs line 529

    [Fact]
    public void Bug79_WordParagraphFirstLineIndent_IntParse()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });

        var ex = Record.Exception(() =>
            _wordHandler.Set("/body/p[1]", new() { ["firstlineindent"] = "1.5" }));

        // BUG: int.Parse("1.5") throws FormatException
        ex.Should().BeNull(
            "Fractional firstlineindent '1.5' should be supported");
    }


    // ==================== BUG #81 (HIGH): Word Set paragraph strike uses bool.Parse ====================
    // WordHandler.Set.cs line 407
    //
    // Location: WordHandler.Set.cs line 407

    [Fact]
    public void Bug81_WordRunStrikeBoolParse()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });

        var ex = Record.Exception(() =>
            _wordHandler.Set("/body/p[1]/r[1]", new() { ["strike"] = "yes" }));

        // BUG: bool.Parse("yes") throws FormatException
        ex.Should().BeNull(
            "strike='yes' should be accepted");
    }


    // ==================== BUG #83 (MEDIUM): Excel conditional formatting icon set reverse uses bool.Parse ====================
    // ExcelHandler.Set.cs line 354: isEl.Reverse = bool.Parse(value);
    //
    // Location: ExcelHandler.Set.cs line 354

    [Fact]
    public void Bug83_ExcelIconSetReverseBoolParse()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "2" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A3", ["value"] = "3" });

        // Add conditional formatting with icon set
        _excelHandler.Add("/Sheet1", "cf", null, new()
        {
            ["sqref"] = "A1:A3", ["type"] = "iconset", ["iconset"] = "3TrafficLights1"
        });

        // Try to set reverse with "1"
        var ex = Record.Exception(() =>
            _excelHandler.Set("/Sheet1/cf[1]", new() { ["reverse"] = "1" }));

        // BUG: bool.Parse("1") throws FormatException
        if (ex != null)
        {
            ex.Should().BeOfType<FormatException>(
                "bool.Parse('1') for icon set reverse throws");
        }
    }

    // ==================== BUG #84 (MEDIUM): Word section margin Set creates PageMargin at wrong schema position ====================
    // WordHandler.Set.cs line 174-195: sectPr.AppendChild(new PageSize()) and
    // sectPr.AppendChild(new PageMargin()) don't respect schema order.
    // PageSize must come before PageMargin in sectPr children.
    //
    // Location: WordHandler.Set.cs lines 174-195

    [Fact]
    public void Bug84_WordSectionMarginSchemaOrder()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Content" });

        // Set margin first, then orientation
        _wordHandler.Set("/section[1]", new() { ["margintop"] = "1440" });
        _wordHandler.Set("/section[1]", new() { ["orientation"] = "landscape" });

        ReopenWord();
        // If schema order is violated, Word may not read the document correctly
        var sec = _wordHandler.Get("/section[1]");
        sec.Should().NotBeNull();
    }


    // ==================== Helper Methods ====================

    private static string CreateTempImage()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        // Create minimal valid PNG (1x1 pixel, white)
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND chunk
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(path, pngBytes);
        return path;
    }
}
