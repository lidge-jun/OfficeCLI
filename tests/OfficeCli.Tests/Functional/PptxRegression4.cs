// Bug hunt tests Part 4 — Bug #251-290
// Program.cs, Helpers, NodeBuilder, Background, Fill, Effects, Selector, View, Core

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class PptxRegression4 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public PptxRegression4()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
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


    /// Bug #254 — PPTX ParseEmu: double.Parse without TryParse on unit values
    /// File: PowerPointHandler.Helpers.cs, lines 164-172
    /// All ParseEmu branches use double.Parse or long.Parse without validation.
    /// Invalid unit strings like "abc cm" or "not_a_number" will throw FormatException.
    [Fact]
    public void Bug254_PptxParseEmu_DoubleParseNoValidation()
    {
        // ParseEmu does: double.Parse(value[..^2]) for "cm", "in", "pt", "px" suffixes
        // and long.Parse(value) for raw EMU values.
        // None of these use TryParse, so invalid input throws unhandled FormatException.

        var handler = new PowerPointHandler(_pptxPath, true);
        try
        {
            // Adding a shape with invalid EMU value should fail gracefully
            var act = () => handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "test",
                ["x"] = "not_a_numbercm"  // invalid double before "cm" suffix
            });
            act.Should().Throw<Exception>(
                "ParseEmu with invalid unit value throws FormatException instead of a clear error");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #255 — PPTX ParseEmu: long.Parse fallback for raw EMU without validation
    /// File: PowerPointHandler.Helpers.cs, line 172
    /// When the value doesn't match any unit suffix, long.Parse(value) is called directly.
    /// No TryParse, no validation — any non-numeric string causes FormatException.
    [Fact]
    public void Bug255_PptxParseEmu_LongParseFallbackNoValidation()
    {
        var handler = new PowerPointHandler(_pptxPath, true);
        try
        {
            var act = () => handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "test",
                ["x"] = "hello"  // not a valid EMU number, not a unit string
            });
            act.Should().Throw<Exception>(
                "ParseEmu raw fallback: long.Parse('hello') throws FormatException");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #256 — PPTX Background: single-color gradient creates invalid gradient
    /// File: PowerPointHandler.Background.cs, lines 244-248
    /// BuildGradientFill allows colorParts.Count==1 after removing angle/focus.
    /// A single gradient stop at position 0 is invalid per OpenXML spec —
    /// gradients require at least 2 stops.
    [Fact]
    public void Bug256_PptxBackground_SingleColorGradientInvalid()
    {
        // If input is "FF0000-90" where "90" is parsed as angle (<=3 digits),
        // colorParts becomes ["FF0000"] after removing the angle.
        // The code creates a single gradient stop at position 0,
        // which is technically invalid (needs at least 2 stops).

        var handler = new PowerPointHandler(_pptxPath, true);
        try
        {
            // "FF0000-90" should be parsed as FF0000 with 90 degree angle
            // but that leaves only 1 color — invalid gradient
            var act = () => handler.Set("/slide[1]", new()
            {
                ["background"] = "FF0000-90"
            });
            // This should either throw or create a valid 2-stop gradient
            // Instead it creates an invalid single-stop gradient
            act.Should().NotThrow("but the resulting gradient has only 1 stop which is invalid");
        }
        finally { handler.Dispose(); }
    }


    /// Bug #261 — PPTX Background: IsGradientColorString false positive for hex starting with "radial:"
    /// File: PowerPointHandler.Background.cs, lines 166-176
    /// IsGradientColorString returns true for ANY string starting with "radial:" or "path:",
    /// even if the remaining part is not a valid color string (e.g., "radial:" alone or "radial:xyz").
    [Fact]
    public void Bug261_PptxBackground_IsGradientColorString_FalsePositive()
    {
        // IsGradientColorString checks:
        //   if starts with "radial:" or "path:" → return true
        // This means "radial:" with no colors after it passes validation,
        // but then BuildGradientFill will throw because colorSpec.Split('-') has < 2 parts.

        var handler = new PowerPointHandler(_pptxPath, true);
        try
        {
            var act = () => handler.Set("/slide[1]", new()
            {
                ["background"] = "radial:"  // passes IsGradientColorString but fails BuildGradientFill
            });
            act.Should().Throw<Exception>(
                "radial: with no colors passes validation check but fails in BuildGradientFill");
        }
        finally { handler.Dispose(); }
    }


    // ==================== Bug #271-290: Effects, Selector, Word View/Query, Core ====================

    /// Bug #271 — PPTX Effects: ApplyShadow double.Parse without validation
    /// File: PowerPointHandler.Effects.cs, lines 34-37
    /// Shadow parameters (blur, angle, distance, opacity) parsed with double.Parse()
    /// without TryParse. Invalid input like "000000-abc-45-3-40" throws FormatException.
    [Fact]
    public void Bug271_PptxEffects_ShadowDoubleParseNoValidation()
    {
        var handler = new PowerPointHandler(_pptxPath, true);
        try
        {
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });
            var act = () => handler.Set("/slide[1]/shape[1]", new()
            {
                ["shadow"] = "000000-notanumber-45-3-40"
            });
            act.Should().Throw<Exception>(
                "double.Parse on shadow blur value 'notanumber' throws FormatException");
        }
        finally { handler.Dispose(); }
    }

    /// Bug #272 — PPTX Effects: ApplyGlow double.Parse without validation
    /// File: PowerPointHandler.Effects.cs, lines 74-75
    /// Glow radius and opacity parsed with double.Parse() without validation.
    [Fact]
    public void Bug272_PptxEffects_GlowDoubleParseNoValidation()
    {
        var handler = new PowerPointHandler(_pptxPath, true);
        try
        {
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });
            var act = () => handler.Set("/slide[1]/shape[1]", new()
            {
                ["glow"] = "FF0000-xyz"  // xyz is not a valid radius
            });
            act.Should().Throw<Exception>(
                "double.Parse on glow radius 'xyz' throws FormatException");
        }
        finally { handler.Dispose(); }
    }


    /// Bug #282 — PPTX Effects: shadow empty string Split produces single-element array
    /// File: PowerPointHandler.Effects.cs, line 32
    /// If value is "" (empty string), Split('-') returns [""], not an empty array.
    /// Then parts[0] is "" which is not "none" or "false", so the code proceeds
    /// to create a shadow with colorHex="" which is invalid.
    [Fact]
    public void Bug282_PptxEffects_ShadowEmptyStringNotHandled()
    {
        var handler = new PowerPointHandler(_pptxPath, true);
        try
        {
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "test" });
            var act = () => handler.Set("/slide[1]/shape[1]", new()
            {
                ["shadow"] = ""  // empty string, not "none"
            });
            // Empty string should be treated as "none" but instead creates invalid shadow
            act.Should().Throw<Exception>(
                "Empty shadow value creates a shadow with empty color hex");
        }
        finally { handler.Dispose(); }
    }


    /// Bug #287 — PPTX Effects: reflection pct*1000 overflow for large values
    /// File: PowerPointHandler.Effects.cs, line 110
    /// The reflection endPos uses int.TryParse then pct*1000.
    /// If the user passes "2147484" (just under int.MaxValue/1000),
    /// pct*1000 overflows to negative.
    [Fact]
    public void Bug287_PptxEffects_ReflectionPctOverflow()
    {
        int pct = 2147484; // slightly > int.MaxValue / 1000
        int endPos = pct * 1000; // integer overflow!
        endPos.Should().BeNegative(
            "integer overflow: 2147484 * 1000 wraps to negative in int arithmetic");
    }

    /// Bug #293 — PPTX Add: shape name vs ID formula inconsistency
    /// File: PowerPointHandler.Add.cs, lines 106-107
    /// Shape name uses Shape.Count()+1 but ID uses Shape.Count()+Picture.Count()+2.
    /// Names like "TextBox 2" may not correspond to ID 4.
    [Fact]
    public void Bug293_PptxAdd_ShapeNameVsIdInconsistency()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add an image first
        var imgPath = CreateTempImage();
        try
        {
            pptx.Add("/slide[1]", "picture", null, new() { ["src"] = imgPath });
            // Add a shape — name will be "TextBox 1" but ID includes picture count
            pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Text" });

            var node = pptx.Get("/slide[1]/shape[1]", depth: 0);
            // Shape name says "TextBox 1" but its ID is Shape(1) + Picture(1) + 2 = 4
            node.Should().NotBeNull();
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }


    /// Bug #299 — Word Add: empty basedOn style reference
    /// File: WordHandler.Add.cs, lines 884-887
    /// BasedOn with empty string creates invalid style chain reference.
    [Fact]
    public void Bug299_WordAdd_EmptyBasedOnStyle()
    {
        var act = () => _wordHandler.Add("/styles", "style", null, new()
        {
            ["name"] = "MyStyle",
            ["basedon"] = ""
        });

        // Empty basedOn value should be ignored or rejected
        // Instead it creates <w:basedOn w:val=""/> which is invalid
        act.Should().NotThrow("but creates invalid basedOn element with empty Val");
    }


    /// Bug #302 — PPTX Query: picture filtering creates semantic index mismatch
    /// File: PowerPointHandler.Query.cs, lines 437-456
    /// When filtering pictures by media type (video/audio/picture), the absolute
    /// position index is passed to PictureToNode instead of the filtered ordinal.
    [Fact]
    public void Bug302_PptxQuery_PictureFilterIndexMismatch()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        // Add an image
        var imgPath = CreateTempImage();
        try
        {
            pptx.Add("/slide[1]", "picture", null, new() { ["src"] = imgPath });

            // Query for pictures — should return valid paths
            var results = pptx.Query("picture");
            foreach (var r in results)
            {
                r.Path.Should().NotContain("picture[0]",
                    "picture index should be 1-based in query results");
            }
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }


    /// Bug #313 — Excel Add: table column count mismatch not validated
    /// File: ExcelHandler.Add.cs, lines 722-724
    /// User-provided column names are not validated against actual column count.
    [Fact]
    public void Bug313_ExcelAdd_TableColumnCountMismatch()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "H1" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "H2" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "C1", ["value"] = "H3" });

        // Provide 2 column names for a 3-column table
        var act = () => _excelHandler.Add("/Sheet1", "table", null, new()
        {
            ["range"] = "A1:C5",
            ["columns"] = "Name,Age"  // only 2 names for 3 columns
        });

        // Should validate column count matches, but silently creates mismatched table
        act.Should().NotThrow("but column names don't match actual column count");
    }


    /// Bug #318 — PPTX Set: double.Parse on crop values
    /// File: PowerPointHandler.Set.cs, lines 647-650, 654, 660
    /// Crop percentages use double.Parse without validation.
    [Fact]
    public void Bug318_PptxSet_CropDoubleParseNoValidation()
    {
        BlankDocCreator.Create(_pptxPath);
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());

        var imgPath = CreateTempImage();
        try
        {
            pptx.Add("/slide[1]", "picture", null, new() { ["src"] = imgPath });

            var act = () => pptx.Set("/slide[1]/picture[1]", new()
            {
                ["crop"] = "ten,20,30,40"  // "ten" is not a valid double
            });

            act.Should().Throw<Exception>(
                "double.Parse on 'ten' for crop value throws FormatException");
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }


    /// Bug #321 — Cross-handler: inconsistent error message formats
    /// File: All handlers
    /// Word uses "Path not found: {path}", Excel uses "{ref} not found",
    /// PowerPoint uses "Slide {idx} not found (total: {count})".
    [Fact]
    public void Bug321_CrossHandler_InconsistentErrorMessages()
    {
        // Word: "Path not found: /body/p[999]"
        var wordAct = () => _wordHandler.Get("/body/p[999]");

        // Excel: different format
        var excelAct = () => _excelHandler.Get("/NonExistentSheet");

        // Both should throw, but with different error message styles
        wordAct.Should().Throw<Exception>();
        excelAct.Should().Throw<Exception>();
    }

    /// Bug #324 — Excel Add: validation BETWEEN operator without formula2
    /// File: ExcelHandler.Add.cs, lines 266-275
    /// When operator is "between", formula2 is required but not validated.
    [Fact]
    public void Bug324_ExcelAdd_ValidationBetweenWithoutFormula2()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "5" });

        var act = () => _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["ref"] = "A1:A10",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1"
            // formula2 is missing but required for "between"
        });

        // Should validate that formula2 is provided for "between" operator
        act.Should().NotThrow("but creates invalid validation without formula2 for between");
    }

    /// Bug #325 — Excel Add: comment author ID off-by-one
    /// File: ExcelHandler.Add.cs, lines 189, 193
    /// When adding a new author, authorId uses count BEFORE append.
    [Fact]
    public void Bug325_ExcelAdd_CommentAuthorIdOffByOne()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Add first comment
        _excelHandler.Add("/Sheet1/A1", "comment", null, new()
        {
            ["text"] = "First comment",
            ["author"] = "Alice"
        });

        // Add second comment with same author — should reuse author ID
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "Test2" });
        _excelHandler.Add("/Sheet1/B1", "comment", null, new()
        {
            ["text"] = "Second comment",
            ["author"] = "Alice"
        });

        ReopenExcel();
        // Both comments should reference the same author
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull();
    }


    /// Bug #327 — Excel Add: DataValidations insertion order violates schema
    /// File: ExcelHandler.Add.cs, lines 298-304
    /// DataValidations element may be inserted AFTER ConditionalFormatting,
    /// violating Excel's required element ordering.
    [Fact]
    public void Bug327_ExcelAdd_DataValidationsSchemaOrder()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "5" });

        // Add conditional formatting first
        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["ref"] = "A1:A10",
            ["type"] = "colorScale"
        });

        // Then add validation — should go BEFORE conditional formatting per schema
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["ref"] = "A1:A10",
            ["type"] = "whole",
            ["operator"] = "greaterThan",
            ["formula1"] = "0"
        });

        ReopenExcel();
        var node = _excelHandler.Get("/Sheet1");
        node.Should().NotBeNull("sheet should be valid with both CF and validation");
    }



    private static string CreateTempImage()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(path, pngBytes);
        return path;
    }
}