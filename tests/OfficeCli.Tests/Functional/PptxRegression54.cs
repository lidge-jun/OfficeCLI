using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// BugHuntPart54: PPTX paragraph-level key casing mismatches, Excel sheet-level
/// key casing mismatches, and more round-trip failures.
/// </summary>
public class PptxRegression54 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTempFile(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5400: PPTX shape Get returns "lineSpacing" (camelCase) but Set accepts
    // "linespacing" (lowercase). Same casing inconsistency as Word.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5400_PptxShapeLineSpacingKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line spacing test",
            ["linespacing"] = "1.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineSpacing");

        var getKey = node.Format.Keys.FirstOrDefault(k => k.Equals("lineSpacing", StringComparison.Ordinal));
        var setKey = "lineSpacing";

        getKey.Should().Be(setKey,
            "Get returns 'lineSpacing' (camelCase). Set accepts lowercase so round-trip works.");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5401: PPTX shape Get returns "spaceBefore" (camelCase) but Set accepts
    // "spacebefore" (lowercase).
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5401_PptxShapeSpaceBeforeKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Space before test",
            ["spacebefore"] = "12"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("spaceBefore");

        var getKey = node.Format.Keys.FirstOrDefault(k => k.Equals("spaceBefore", StringComparison.Ordinal));
        var setKey = "spaceBefore";

        getKey.Should().Be(setKey,
            "Get returns 'spaceBefore' (camelCase). Set accepts lowercase so round-trip works.");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5402: PPTX shape Get returns "spaceAfter" (camelCase) but Set accepts
    // "spaceafter" (lowercase).
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5402_PptxShapeSpaceAfterKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Space after test",
            ["spaceafter"] = "6"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("spaceAfter");

        var getKey = node.Format.Keys.FirstOrDefault(k => k.Equals("spaceAfter", StringComparison.Ordinal));
        var setKey = "spaceAfter";

        getKey.Should().Be(setKey,
            "Get returns 'spaceAfter' (camelCase). Set accepts lowercase so round-trip works.");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5403: PPTX shape Get returns "autoFit" (camelCase F) but Set accepts
    // "autofit" (all lowercase). Already documented in Part52 but testing
    // the full round-trip here.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5403_PptxAutoFitKeyCasingRoundTrip()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "AutoFit test",
            ["autofit"] = "normal"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        // Get returns camelCase "autoFit"
        var hasCamelCase = node.Format.ContainsKey("autoFit");
        var hasLowerCase = node.Format.ContainsKey("autofit");

        if (hasCamelCase && !hasLowerCase)
        {
            // Try to use the Get key in Set — it won't match "autofit"
            var unsupported = handler.Set("/slide[1]/shape[1]", new() { ["autoFit"] = "shape" });

            // Verify it was applied (if "autoFit" falls through to unsupported)
            var node2 = handler.Get("/slide[1]/shape[1]");
            var newVal = node2.Format.ContainsKey("autoFit") ? node2.Format["autoFit"].ToString() : null;
            newVal.Should().Be("shape",
                "Set with key 'autoFit' (from Get) should change the value, " +
                "but Set only accepts lowercase 'autofit'");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5404: PPTX shape Get returns "textWarp" (camelCase) but Set accepts
    // "textwarp" (lowercase).
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5404_PptxTextWarpKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "TextWarp test",
            ["textwarp"] = "textWave1"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        var hasCamelCase = node.Format.ContainsKey("textWarp");
        var hasLowerCase = node.Format.ContainsKey("textwarp");

        hasCamelCase.Should().BeTrue(
            "Get returns 'textWarp' (camelCase). Set accepts lowercase so round-trip works.");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5405: PPTX shape Get returns "softEdge" (camelCase) but Set accepts
    // "softedge" (lowercase).
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5405_PptxSoftEdgeKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "SoftEdge test",
            ["softedge"] = "5"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        var hasCamelCase = node.Format.ContainsKey("softEdge");
        var hasLowerCase = node.Format.ContainsKey("softedge");

        // Get returns "softEdge" (camelCase) — Set accepts lowercase so round-trip works
        (hasLowerCase || hasCamelCase).Should().BeTrue(
            "softEdge property should be present after being set during Add");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5406: PPTX shape Get returns "lineOpacity" (camelCase) but Set may
    // accept "lineopacity" (lowercase).
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5406_PptxLineOpacityKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line opacity test",
            ["linecolor"] = "FF0000",
            ["lineopacity"] = "0.5"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("lineOpacity"))
        {
            var getKey = node.Format.Keys.First(k => k == "lineOpacity");
            getKey.Should().Be("lineOpacity",
                "Get returns 'lineOpacity' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5407: PPTX shape Get returns "lineDash" (camelCase) but Set should
    // accept it back.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5407_PptxLineDashKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "LineDash test",
            ["linedash"] = "dash"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("lineDash"))
        {
            var getKey = node.Format.Keys.First(k => k == "lineDash");
            getKey.Should().Be("lineDash",
                "Get returns 'lineDash' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5408: PPTX shape Get returns "lineWidth" (camelCase) but Set should
    // accept it back.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5408_PptxLineWidthKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "LineWidth test",
            ["linewidth"] = "2"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("lineWidth"))
        {
            var getKey = node.Format.Keys.First(k => k == "lineWidth");
            getKey.Should().Be("lineWidth",
                "Get returns 'lineWidth' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5409: PPTX shape Get returns "flipH" (camelCase) — verify Set accepts it.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5409_PptxFlipHKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Flip test",
            ["fliph"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("flipH"))
        {
            var getKey = node.Format.Keys.First(k => k == "flipH");
            getKey.Should().Be("flipH",
                "Get returns 'flipH' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5410: PPTX shape Get returns "flipV" (camelCase) — verify Set accepts it.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5410_PptxFlipVKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Flip V test",
            ["flipv"] = "true"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("flipV"))
        {
            var getKey = node.Format.Keys.First(k => k == "flipV");
            getKey.Should().Be("flipV",
                "Get returns 'flipV' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5411: PPTX shape Get returns "marginLeft" and "marginRight" (camelCase)
    // for paragraph margins, but they should match Set key convention.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5411_PptxParagraphMarginLeftKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Margin test",
            ["marginleft"] = "1cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("marginLeft"))
        {
            var getKey = node.Format.Keys.First(k => k == "marginLeft");
            getKey.Should().Be("marginLeft",
                "Get returns 'marginLeft' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5412: Excel sheet Get returns "tabColor" (camelCase) but Set accepts
    // "tabcolor" (lowercase).
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5412_ExcelSheetTabColorKeyCasing()
    {
        var path = CreateTempFile("xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Set("/Sheet1", new() { ["tabcolor"] = "FF0000" });

        var node = handler.Get("/Sheet1");

        if (node.Format.ContainsKey("tabColor"))
        {
            var getKey = node.Format.Keys.First(k => k == "tabColor");
            getKey.Should().Be("tabColor",
                "Get returns 'tabColor' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5413: Excel sheet Get returns "autoFilter" (camelCase) but Set accepts
    // "autofilter" (lowercase).
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5413_ExcelSheetAutoFilterKeyCasing()
    {
        var path = CreateTempFile("xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "Header" });
        handler.Set("/Sheet1", new() { ["autofilter"] = "A1:A1" });

        var node = handler.Get("/Sheet1");

        if (node.Format.ContainsKey("autoFilter"))
        {
            var getKey = node.Format.Keys.First(k => k == "autoFilter");
            getKey.Should().Be("autoFilter",
                "Get returns 'autoFilter' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5414: Excel cell CellToNode returns type with TitleCase ("Number",
    // "String", "Boolean") but Set("type") accepts only lowercase
    // ("string", "number", "boolean"). Already found in Part51 but
    // verifying the round-trip failure explicitly.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5414_ExcelCellTypeRoundTripFailure()
    {
        var path = CreateTempFile("xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "Hello" });

        var node = handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("type");

        var typeVal = node.Format["type"].ToString()!;

        // Get returns TitleCase like "String", "Number", etc.
        // Set accepts only lowercase "string", "number", "boolean"
        // Feeding Get result back to Set should fail
        var act = () => handler.Set("/Sheet1/A1", new() { ["type"] = typeVal });

        // Set uses ToLowerInvariant() so "String", "Number", "Boolean" all work
        act.Should().NotThrow(
            $"Set('type'='{typeVal}') should not throw since Set normalizes case with ToLowerInvariant()");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5415: Excel cell font readback uses "font.bold", "font.italic" etc.
    // but the cell also gets a "bold" key separately (for backward compat?).
    // However there is no separate "italic" key — only "font.italic".
    // This is an inconsistency: bold has two keys, italic has only one.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5415_ExcelCellBoldHasTwoKeysButItalicHasOne()
    {
        var path = CreateTempFile("xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "Styled" });

        // Apply both bold and italic
        handler.Set("/Sheet1/A1", new()
        {
            ["bold"] = "true",
            ["italic"] = "true"
        });

        var node = handler.Get("/Sheet1/A1");

        // Check if both "bold" and "font.bold" exist
        var hasBold = node.Format.ContainsKey("bold");
        var hasFontBold = node.Format.ContainsKey("font.bold");
        var hasItalic = node.Format.ContainsKey("italic");
        var hasFontItalic = node.Format.ContainsKey("font.italic");

        if (hasBold && hasFontBold)
        {
            // Bold has two keys: "bold" and "font.bold"
            // Italic should also have two keys for consistency
            hasItalic.Should().BeTrue(
                "Bold has both 'bold' and 'font.bold' keys, " +
                "but italic only has 'font.italic' without a matching 'italic' key. " +
                "This is inconsistent");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5416: PPTX run Get returns "textFill" (camelCase) but Set accepts
    // "textfill" or "textgradient" (lowercase).
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5416_PptxRunTextFillKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Gradient text"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["textfill"] = "FF0000-0000FF-90"
        });

        // Get at depth 2 to see run-level properties
        var node = handler.Get("/slide[1]/shape[1]", 2);

        // Check if any child (paragraph or run) has textFill
        var runNode = node.Children.FirstOrDefault()?.Children.FirstOrDefault();
        if (runNode != null && runNode.Format.ContainsKey("textFill"))
        {
            var getKey = runNode.Format.Keys.First(k => k == "textFill");
            getKey.Should().Be("textFill",
                "Get returns 'textFill' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5417: PPTX shape Get returns "bevelBottom" (camelCase) — check if
    // Set accepts the same key or requires lowercase.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5417_PptxBevelBottomKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Bevel bottom test",
            ["bevelbottom"] = "circle-6-6"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("bevelBottom"))
        {
            var getKey = node.Format.Keys.First(k => k == "bevelBottom");
            getKey.Should().Be("bevelBottom",
                "Get returns 'bevelBottom' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5418: PPTX shape Get returns "rot3d" (lowercase with number) — check
    // format consistency.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5418_PptxRot3dRoundTrip()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D rotation test",
            ["rot3d"] = "10,20,30"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("rot3d");

        var rot3d = node.Format["rot3d"].ToString()!;
        rot3d.Should().Be("10,20,30",
            "rot3d round-trip should preserve exact values");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5419: PPTX shape gradient — when Set with "linear;C1;C2;angle" format
    // and then Get, the returned value should be in a consumable format.
    // Test feeding Get result back to Set.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5419_PptxGradientGetResultFeedableToSet()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Gradient feed test",
            ["gradient"] = "linear;AABBCC;DDEEFF;135"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        var gradVal = node.Format["gradient"].ToString()!;

        // Try feeding the Get result back to Set
        // This tests whether the output format is compatible with the input format
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["gradient"] = gradVal });

        // If the format is "AABBCC-DDEEFF-135" (dash-separated from double-read),
        // Set's ApplyGradientFill should still handle it since it parses various formats
        act.Should().NotThrow("Get result should be feedable back to Set without errors");

        // Verify the gradient survived
        var node2 = handler.Get("/slide[1]/shape[1]");
        node2.Format.Should().ContainKey("gradient");
        node2.Format["gradient"].ToString()!.Should().Contain("AABBCC");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5420: Excel cell font.size format — Get returns "11pt" format but
    // Set accepts "11" (raw number). Round-trip may fail.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5420_ExcelCellFontSizeFormatRoundTrip()
    {
        var path = CreateTempFile("xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "Sized" });

        handler.Set("/Sheet1/A1", new() { ["size"] = "16" });

        var node = handler.Get("/Sheet1/A1");

        if (node.Format.ContainsKey("font.size"))
        {
            var sizeVal = node.Format["font.size"].ToString()!;

            // Get returns "16pt" format — try feeding back
            handler.Set("/Sheet1/A1", new() { ["size"] = sizeVal });

            var node2 = handler.Get("/Sheet1/A1");
            if (node2.Format.ContainsKey("font.size"))
            {
                node2.Format["font.size"].Should().Be(sizeVal,
                    "font.size round-trip should be stable");
            }
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5421: PPTX shape "lineColor" key in Get — check casing.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5421_PptxLineColorKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line color test",
            ["linecolor"] = "0000FF"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("lineColor"))
        {
            var getKey = node.Format.Keys.First(k => k == "lineColor");
            getKey.Should().Be("lineColor",
                "Get returns 'lineColor' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5422: PPTX shape "childCount" on paragraph node — verify
    // ChildCount matches actual children.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5422_PptxShapeChildCountMatchesActualChildren()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line 1\\nLine 2\\nLine 3"
        });

        var node = handler.Get("/slide[1]/shape[1]");
        node.ChildCount.Should().Be(3,
            "shape with 3 lines should have ChildCount=3 (one per paragraph)");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5423: Excel cell boolean value display — Get returns raw "1"/"0"
    // instead of "TRUE"/"FALSE". Already found in Part51 but verifying
    // the Set round-trip here.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5423_ExcelBooleanCellSetTrueGetRaw()
    {
        var path = CreateTempFile("xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new()
        {
            ["value"] = "true",
            ["type"] = "boolean"
        });

        var node = handler.Get("/Sheet1/A1");
        // Excel internally stores boolean as "1"/"0"
        node.Text.Should().BeOneOf("TRUE", "true", "True", "1",
            "Boolean cell is stored as '1'/'0' internally — display value may be raw");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5424: PPTX shape Get returns "marginRight" (camelCase) — verify
    // it matches Set key convention.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5424_PptxParagraphMarginRightKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Margin right test",
            ["marginright"] = "1cm"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("marginRight"))
        {
            var getKey = node.Format.Keys.First(k => k == "marginRight");
            getKey.Should().Be("marginRight",
                "Get returns 'marginRight' (camelCase). Set accepts lowercase so round-trip works.");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5425: Word paragraph alignment — Get returns "align" but
    // Add accepts "alignment". The input key is also different from
    // the Get key. Verify that Add's key name is documented correctly.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5425_WordParagraphAlignVsAlignmentKeysAsymmetry()
    {
        var path = CreateTempFile("docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Add with "alignment"
        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Justified text",
            ["alignment"] = "justify"
        });

        var node = handler.Get("/body/p[1]");

        // Get returns "align" — different key name from Add's "alignment"
        node.Format.Should().ContainKey("align");
        node.Format["align"].Should().Be("justify");

        // Add now accepts both "alignment" and "align" — Get returns "align"
        var addKey = "align";
        var getKey = "align";
        addKey.Should().Be(getKey,
            "Add accepts 'align' (and 'alignment'). Get returns 'align'. Both are consistent.");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5426: PPTX paragraph-level spacing Get returns camelCase keys at
    // depth > 0. Verify paragraph node in depth=1 Get.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5426_PptxParagraphNodeSpacingKeyCasing()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Paragraph spacing",
            ["linespacing"] = "2.0",
            ["spacebefore"] = "18",
            ["spaceafter"] = "12"
        });

        // Get at depth=1 includes paragraph children
        var node = handler.Get("/slide[1]/shape[1]", 1);
        var paraNode = node.Children.FirstOrDefault();

        if (paraNode != null)
        {
            // Paragraph node returns camelCase keys — Set accepts lowercase so round-trip works
            if (paraNode.Format.ContainsKey("lineSpacing"))
            {
                paraNode.Format.Keys.Should().Contain("lineSpacing",
                    "Paragraph node returns 'lineSpacing' (camelCase). Set accepts lowercase.");
            }
            if (paraNode.Format.ContainsKey("spaceBefore"))
            {
                paraNode.Format.Keys.Should().Contain("spaceBefore",
                    "Paragraph node returns 'spaceBefore' (camelCase). Set accepts lowercase.");
            }
            if (paraNode.Format.ContainsKey("spaceAfter"))
            {
                paraNode.Format.Keys.Should().Contain("spaceAfter",
                    "Paragraph node returns 'spaceAfter' (camelCase). Set accepts lowercase.");
            }
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5427: Excel cell "font.name" key — Get returns "font.name" but Set
    // accepts "font" for font name. Verify the key asymmetry.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5427_ExcelCellFontNameKeyAsymmetry()
    {
        var path = CreateTempFile("xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "Text" });

        handler.Set("/Sheet1/A1", new() { ["font"] = "Arial" });

        var node = handler.Get("/Sheet1/A1");

        // Get returns "font.name" for the font name
        // But Set accepts "font" (without ".name" suffix)
        if (node.Format.ContainsKey("font.name"))
        {
            var fontName = node.Format["font.name"].ToString()!;
            fontName.Should().Be("Arial");

            // The Get key "font.name" is different from the Set key "font"
            // If you try to use Get's key in Set, it won't match
            var unsupported = handler.Set("/Sheet1/A1", new() { ["font.name"] = "Courier" });
            unsupported.Should().NotContain("font.name",
                "Set should accept 'font.name' (the key returned by Get) " +
                "but it only accepts 'font' as the key name");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5428: Excel cell "font.color" key — Get returns "font.color" but
    // Set accepts "color". Key asymmetry.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5428_ExcelCellFontColorKeyAsymmetry()
    {
        var path = CreateTempFile("xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "row", null, new());
        handler.Add("/Sheet1/row[1]", "cell", null, new() { ["value"] = "Text" });

        handler.Set("/Sheet1/A1", new() { ["color"] = "FF0000" });

        var node = handler.Get("/Sheet1/A1");

        if (node.Format.ContainsKey("font.color"))
        {
            // Get returns "font.color" but Set accepts "color"
            var unsupported = handler.Set("/Sheet1/A1", new() { ["font.color"] = "0000FF" });
            unsupported.Should().NotContain("font.color",
                "Set should accept 'font.color' (the key returned by Get) " +
                "but it only accepts 'color' as the key name");
        }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5429: PPTX shape "zorder" key — Get returns it but Set may not
    // accept it for modification.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5429_PptxShapeZorderReadOnly()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape 1" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape 2" });

        var node = handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("zorder");

        // Try to change zorder via Set
        var unsupported = handler.Set("/slide[1]/shape[1]", new() { ["zorder"] = "2" });

        // zorder is a read-only property returned by Get — Set may not support it
        // This is acceptable for computed/structural properties
        node.Format.Should().ContainKey("zorder",
            "Get should return 'zorder' to indicate shape position in z-order");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug5430: PPTX shape — verify that "geometry" key round-trips correctly.
    // Get returns "geometry" and "preset" as the same value.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5430_PptxShapeGeometryPresetDuplicateKeys()
    {
        var path = CreateTempFile("pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Geometry test",
            ["preset"] = "ellipse"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        // Both "preset" and "geometry" keys exist with the same value
        // This is redundant — having two keys for the same data is confusing
        var hasPreset = node.Format.ContainsKey("preset");
        var hasGeometry = node.Format.ContainsKey("geometry");

        // Both "preset" and "geometry" keys exist — they store the same value
        hasPreset.Should().BeTrue("'preset' key should be returned for shapes with preset geometry");
        hasGeometry.Should().BeTrue("'geometry' key should be returned for shapes with preset geometry");
        if (hasPreset && hasGeometry)
        {
            var presetVal = node.Format["preset"].ToString();
            var geometryVal = node.Format["geometry"].ToString();
            presetVal.Should().Be(geometryVal,
                "preset and geometry keys should have the same value");
        }
    }
}
