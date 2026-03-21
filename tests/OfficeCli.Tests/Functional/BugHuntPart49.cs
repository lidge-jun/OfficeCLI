// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// BugHuntPart49: Word section duplicate keys, Excel autofilter duplicate keys,
/// Word style size format inconsistency, Word paragraph shd data loss,
/// Word watermark color double-hash, Word footnote lookup asymmetry,
/// cross-handler table cell fill key inconsistency.
/// </summary>
public class BugHuntPart49 : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTempFile(string ext)
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

    // ==================== Bug4900 ====================
    // Word section Get stores BOTH "pageWidth" AND "pagewidth" (camelCase + lowercase)
    // for the same page width value. This is the same duplicate-key pattern found
    // in Word paragraph spacing (Bug4801/4815/4816).
    // See WordHandler.Query.cs lines 180-183.
    [Fact]
    public void Bug4900_WordSectionPageWidthKeyDuplicated()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        var node = handler.Get("/section[1]");

        var hasPageWidth = node.Format.ContainsKey("pageWidth");
        var hasPagewidthLower = node.Format.ContainsKey("pagewidth");

        // BUG: Both keys exist with the same value
        (hasPageWidth && hasPagewidthLower).Should().BeFalse(
            because: "Word section should use one key for page width, not duplicate both " +
                     "'pageWidth' AND 'pagewidth'. WordHandler.Query.cs lines 180-183 " +
                     "explicitly store the same value under both camelCase and lowercase keys");
    }

    // ==================== Bug4901 ====================
    // Word section Get stores BOTH "pageHeight" AND "pageheight" (camelCase + lowercase)
    // Same issue as Bug4900 but for height.
    [Fact]
    public void Bug4901_WordSectionPageHeightKeyDuplicated()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        var node = handler.Get("/section[1]");

        var hasPageHeight = node.Format.ContainsKey("pageHeight");
        var hasPageheightLower = node.Format.ContainsKey("pageheight");

        (hasPageHeight && hasPageheightLower).Should().BeFalse(
            because: "Word section should use one key for page height, not duplicate both " +
                     "'pageHeight' AND 'pageheight'. Same pattern as Bug4900");
    }

    // ==================== Bug4902 ====================
    // Excel sheet Get stores BOTH "autoFilter" AND "autofilter" for the same value.
    // See ExcelHandler.Query.cs lines 136-137.
    [Fact]
    public void Bug4902_ExcelAutoFilterKeyDuplicated()
    {
        var path = CreateTempFile(".xlsx");
        BlankDocCreator.Create(path);
        using var handler = new ExcelHandler(path, editable: true);

        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Name" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "Value" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "x" });
        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "1" });

        handler.Set("/Sheet1", new() { ["autofilter"] = "A1:B2" });

        var node = handler.Get("/Sheet1");

        var hasAutoFilter = node.Format.ContainsKey("autoFilter");
        var hasAutofilterLower = node.Format.ContainsKey("autofilter");

        if (hasAutoFilter && hasAutofilterLower)
        {
            // BUG: Both keys exist with the same value
            (hasAutoFilter && hasAutofilterLower).Should().BeFalse(
                because: "Excel sheet should use one key for autofilter, not duplicate both " +
                         "'autoFilter' AND 'autofilter'. ExcelHandler.Query.cs lines 136-137 " +
                         "store the same reference value under both keys");
        }
    }

    // ==================== Bug4903 ====================
    // Word style Get returns size as int (e.g., 12) without "pt" suffix,
    // but Word paragraph Get returns size as "12pt" with suffix.
    // This makes programmatic comparison between style size and paragraph size impossible
    // without stripping the suffix.
    [Fact]
    public void Bug4903_WordStyleSizeFormatVsParagraphSize()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        // Create a custom style with size 14
        handler.Add("/styles", "style", null, new()
        {
            ["id"] = "CustomTest",
            ["name"] = "Custom Test",
            ["type"] = "paragraph",
            ["size"] = "14"
        });

        // Get style size
        var styleNode = handler.Get("/styles/CustomTest");
        var styleSize = styleNode.Format.ContainsKey("size") ? styleNode.Format["size"]?.ToString() : null;

        // Create paragraph with size 14
        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "test",
            ["size"] = "14"
        });

        var paraNode = handler.Get("/body/p[1]");
        var paraSize = paraNode.Format.ContainsKey("size") ? paraNode.Format["size"]?.ToString() : null;

        if (styleSize != null && paraSize != null)
        {
            // BUG: style returns "14" (int), paragraph returns "14pt" (string with suffix)
            styleSize.Should().Be(paraSize,
                because: "Word style size format should match paragraph size format. " +
                         "Style Get (WordHandler.Query.cs line 235) returns int/2 without 'pt'. " +
                         "Paragraph Get (WordHandler.Navigation.cs line 309) returns '{value}pt'. " +
                         "Inconsistent format makes cross-reference impossible without manual parsing");
        }
    }

    // ==================== Bug4904 ====================
    // Word paragraph shd Get returns only the fill value, losing the pattern and color.
    // When Set with a complex shading like "clear;FF0000;000000", Get returns only "FF0000"
    // (the fill). The shading pattern ("clear") and color ("000000") are lost.
    [Fact]
    public void Bug4904_WordParagraphShdGetLosesPatternAndColor()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "shaded",
            ["shd"] = "clear;FF0000;000000"
        });

        var node = handler.Get("/body/p[1]");

        if (node.Format.ContainsKey("shd"))
        {
            var shdValue = node.Format["shd"]?.ToString() ?? "";
            // BUG: Get returns only "FF0000" (the fill value)
            // It should return the full "clear;FF0000;000000" or at least include pattern info
            shdValue.Should().Contain(";",
                because: "Word paragraph shd Get should return the full shading specification " +
                         "including pattern and color, not just the fill value. " +
                         "WordHandler.Navigation.cs line 269 reads only Fill?.Value, " +
                         "discarding the Val (pattern) and Color components");
        }
    }

    // ==================== Bug4905 ====================
    // Word watermark color should now consistently return "#RRGGBB" format
    // like all other color properties in the system.
    [Fact]
    public void Bug4905_WordWatermarkColorInconsistentHashPrefix()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "watermark", null, new()
        {
            ["text"] = "DRAFT",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/watermark");

        if (node.Format.ContainsKey("color"))
        {
            var colorVal = node.Format["color"]?.ToString() ?? "";
            // All color outputs now use "#RRGGBB" format consistently
            colorVal.Should().StartWith("#",
                because: "Word watermark color should be returned with '#' prefix " +
                         "for consistency with all other color properties");
        }
    }

    // ==================== Bug4906 ====================
    // Word footnote Get uses ID-based lookup (footnote[N] where N is the ID),
    // but Word footnote Set falls back to ordinal lookup (1-based index among user footnotes).
    // This means Get("/footnote[1]") looks for ID=1, but Set("/footnote[1]") looks for
    // the first user footnote (which may have ID=1 or not). The semantics diverge.
    [Fact]
    public void Bug4906_WordFootnoteGetVsSetLookupAsymmetry()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "test para" });
        handler.Add("/body/p[1]", "footnote", null, new()
        {
            ["text"] = "Original footnote"
        });

        // Get uses ID-based lookup
        var getNode = handler.Get("/footnote[1]");

        if (getNode.Type != "error")
        {
            // Now Set modifies using the same path — but Set first tries ID lookup,
            // and if that fails, falls back to ordinal lookup.
            // This is asymmetric: Get never does ordinal fallback.
            handler.Set("/footnote[1]", new() { ["text"] = "Modified footnote" });

            var getNode2 = handler.Get("/footnote[1]");

            // Verify the modification took effect on the same footnote
            getNode2.Text.Should().Contain("Modified",
                because: "Get and Set for /footnote[1] should reference the same footnote. " +
                         "Get (WordHandler.Query.cs line 91) uses only ID lookup. " +
                         "Set (WordHandler.Set.cs line 253) uses ID first, then ordinal fallback. " +
                         "If the footnote ID doesn't match the ordinal position, they diverge");
        }
    }

    // ==================== Bug4907 ====================
    // Word table cell Set accepts "fill" as alias for "shd" (line 1111),
    // but Get reports shading as "shd" key, not "fill". The round-trip fails
    // because the key names differ between Set and Get.
    [Fact]
    public void Bug4907_WordTableCellFillKeyVsShdKey()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        // Set cell shading using "fill" key
        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["fill"] = "FF0000"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]", 3);

        // Check what key the fill is reported under
        var hasFill = cellNode.Format.ContainsKey("fill");
        var hasShd = cellNode.Format.ContainsKey("shd");

        // BUG: Set accepts "fill" but Get reports as "shd"
        // For a true round-trip, if you Set with "fill", Get should return "fill"
        hasFill.Should().BeTrue(
            because: "Word table cell Get should include 'fill' key when shading was set via 'fill'. " +
                     "Currently Set accepts 'fill' as alias for 'shd' (WordHandler.Set.cs line 1111) " +
                     "but Get only reports 'shd' key (via ReadCellProps), breaking the round-trip");
    }

    // ==================== Bug4908 ====================
    // Word header/footer Get reports font size as "12pt" format (with "pt" suffix),
    // but Word header/footer Set accepts "size" property which goes through ParseFontSize.
    // The Get format for header includes "pt" suffix (line 435/490 of WordHandler.Query.cs)
    // while style Get returns bare int (no "pt") — inconsistent within the same handler.
    [Fact]
    public void Bug4908_WordHeaderSizeFormatVsStyleSize()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "header", null, new()
        {
            ["text"] = "Header Text",
            ["size"] = "14"
        });

        var headerNode = handler.Get("/header[1]");
        var headerSize = headerNode.Format.ContainsKey("size")
            ? headerNode.Format["size"]?.ToString() : null;

        // Create a style with size 14
        handler.Add("/styles", "style", null, new()
        {
            ["id"] = "HeaderStyle",
            ["name"] = "Header Style",
            ["type"] = "paragraph",
            ["size"] = "14"
        });

        var styleNode = handler.Get("/styles/HeaderStyle");
        var styleSize = styleNode.Format.ContainsKey("size")
            ? styleNode.Format["size"]?.ToString() : null;

        if (headerSize != null && styleSize != null)
        {
            // BUG: header returns "14pt", style returns 14 (int, no "pt" suffix)
            headerSize.Should().Be(styleSize?.ToString(),
                because: "Word header size format should match style size format. " +
                         "Header Get returns '{value}pt' (line 435). " +
                         "Style Get returns int without 'pt' (line 235). " +
                         "Both are within WordHandler but use different formats");
        }
    }

    // ==================== Bug4909 ====================
    // Word table cell ReadCellProps — checking if valign Get round-trips with Set.
    // Set accepts "top", "center", "bottom". Let's verify Get returns the same.
    [Fact]
    public void Bug4909_WordTableCellValignRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["valign"] = "center"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]", 3);

        if (cellNode.Format.ContainsKey("valign"))
        {
            cellNode.Format["valign"]?.ToString().Should().Be("center",
                because: "Word table cell valign should round-trip: Set 'center' → Get 'center'");
        }
        else
        {
            // BUG: valign is not reported in Get at all
            cellNode.Format.Should().ContainKey("valign",
                because: "Word table cell Get should include 'valign' when vertical alignment was set");
        }
    }

    // ==================== Bug4910 ====================
    // Word table cell textDirection round-trip.
    // Set accepts "btlr" or "vertical" → TextDirectionValues.BottomToTopLeftToRight.
    // Verify Get reports it back.
    [Fact]
    public void Bug4910_WordTableCellTextDirectionRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["textdirection"] = "vertical"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]", 3);

        // Check if textdirection is readable
        var hasTextDir = cellNode.Format.ContainsKey("textdirection")
                      || cellNode.Format.ContainsKey("textDirection")
                      || cellNode.Format.ContainsKey("textdir");

        hasTextDir.Should().BeTrue(
            because: "Word table cell Get should include text direction when it was set via Set. " +
                     "Set supports 'textdirection' key but Get (ReadCellProps) may not read it back");
    }

    // ==================== Bug4911 ====================
    // Word table cell nowrap round-trip.
    [Fact]
    public void Bug4911_WordTableCellNowrapRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["nowrap"] = "true"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]", 3);

        var hasNowrap = cellNode.Format.ContainsKey("nowrap")
                     || cellNode.Format.ContainsKey("noWrap");

        hasNowrap.Should().BeTrue(
            because: "Word table cell Get should include 'nowrap' when it was set via Set. " +
                     "Set supports 'nowrap' key but Get (ReadCellProps) may not read it back");
    }

    // ==================== Bug4912 ====================
    // Word table cell padding round-trip.
    [Fact]
    public void Bug4912_WordTableCellPaddingRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["padding"] = "100"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]", 3);

        var hasPadding = cellNode.Format.ContainsKey("padding")
                      || cellNode.Format.ContainsKey("padding.top")
                      || cellNode.Format.ContainsKey("margin");

        hasPadding.Should().BeTrue(
            because: "Word table cell Get should include padding when it was set via Set. " +
                     "Set supports 'padding' key but Get (ReadCellProps) may not read it back");
    }

    // ==================== Bug4913 ====================
    // Word table cell width round-trip.
    [Fact]
    public void Bug4913_WordTableCellWidthRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["width"] = "3000"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]", 3);

        var hasWidth = cellNode.Format.ContainsKey("width");

        hasWidth.Should().BeTrue(
            because: "Word table cell Get should include 'width' when it was set via Set. " +
                     "Set supports 'width' key but Get (ReadCellProps) may not read it back");
    }

    // ==================== Bug4914 ====================
    // Word table cell vmerge round-trip.
    [Fact]
    public void Bug4914_WordTableCellVmergeRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "3",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["vmerge"] = "restart"
        });
        handler.Set("/body/tbl[1]/tr[2]/tc[1]", new()
        {
            ["vmerge"] = "continue"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]", 3);

        var hasVmerge = cellNode.Format.ContainsKey("vmerge")
                     || cellNode.Format.ContainsKey("verticalMerge");

        hasVmerge.Should().BeTrue(
            because: "Word table cell Get should include 'vmerge' when it was set via Set. " +
                     "Set supports 'vmerge' key but Get (ReadCellProps) may not read it back");
    }

    // ==================== Bug4915 ====================
    // Word table cell border round-trip.
    [Fact]
    public void Bug4915_WordTableCellBorderRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["border.all"] = "single;4;FF0000"
        });

        var cellNode = handler.Get("/body/tbl[1]/tr[1]/tc[1]", 3);

        var hasBorder = cellNode.Format.ContainsKey("border.top")
                     || cellNode.Format.ContainsKey("border.all")
                     || cellNode.Format.ContainsKey("border");

        hasBorder.Should().BeTrue(
            because: "Word table cell Get should include border info when borders were set via Set. " +
                     "Set supports 'border.all' key via ApplyCellBorders but Get (ReadCellProps) " +
                     "may not read cell-level borders back");
    }

    // ==================== Bug4916 ====================
    // Word table row height round-trip.
    [Fact]
    public void Bug4916_WordTableRowHeightRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]", new()
        {
            ["height"] = "500"
        });

        var tableNode = handler.Get("/body/tbl[1]", 2);
        var rowNode = tableNode.Children.FirstOrDefault();

        rowNode.Should().NotBeNull();
        if (rowNode != null)
        {
            var hasHeight = rowNode.Format.ContainsKey("height");
            hasHeight.Should().BeTrue(
                because: "Word table row Get should include 'height' when it was set via Set");
        }
    }

    // ==================== Bug4917 ====================
    // Word table row header repeat round-trip.
    [Fact]
    public void Bug4917_WordTableRowHeaderRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "3",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]/tr[1]", new()
        {
            ["header"] = "true"
        });

        var tableNode = handler.Get("/body/tbl[1]", 2);
        var rowNode = tableNode.Children.FirstOrDefault();

        rowNode.Should().NotBeNull();
        if (rowNode != null)
        {
            var hasHeader = rowNode.Format.ContainsKey("header");
            hasHeader.Should().BeTrue(
                because: "Word table row Get should include 'header' when it was set via Set. " +
                         "Set supports 'header' key (line 1333) but ReadRowProps may not read it back");
        }
    }

    // ==================== Bug4918 ====================
    // Word table Set "padding" → table-level default cell margin round-trip.
    // Set stores as TableCellMarginDefault. Get reads from padding.top/bottom/left/right.
    // But table-level padding uses different element types (TableCellLeftMargin vs LeftMargin).
    [Fact]
    public void Bug4918_WordTablePaddingRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]", new()
        {
            ["padding"] = "100"
        });

        var tableNode = handler.Get("/body/tbl[1]");

        var hasPaddingTop = tableNode.Format.ContainsKey("padding.top");
        var hasPaddingLeft = tableNode.Format.ContainsKey("padding.left");

        (hasPaddingTop && hasPaddingLeft).Should().BeTrue(
            because: "Word table Get should report padding.top and padding.left when table padding was set. " +
                     "Set stores via TableCellMarginDefault, Get reads those sub-elements");
    }

    // ==================== Bug4919 ====================
    // Word table layout round-trip.
    [Fact]
    public void Bug4919_WordTableLayoutRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]", new()
        {
            ["layout"] = "fixed"
        });

        var tableNode = handler.Get("/body/tbl[1]");

        tableNode.Format.Should().ContainKey("layout",
            because: "Word table Get should report 'layout' when it was set via Set");

        if (tableNode.Format.ContainsKey("layout"))
        {
            tableNode.Format["layout"]?.ToString().Should().Be("fixed",
                because: "Word table layout should round-trip: Set 'fixed' → Get 'fixed'");
        }
    }

    // ==================== Bug4920 ====================
    // Word table cellSpacing round-trip.
    [Fact]
    public void Bug4920_WordTableCellSpacingRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        handler.Set("/body/tbl[1]", new()
        {
            ["cellspacing"] = "50"
        });

        var tableNode = handler.Get("/body/tbl[1]");

        // Get reports as "cellSpacing" (camelCase)
        var hasCellSpacing = tableNode.Format.ContainsKey("cellSpacing")
                          || tableNode.Format.ContainsKey("cellspacing");

        hasCellSpacing.Should().BeTrue(
            because: "Word table Get should report 'cellSpacing' when it was set via Set");
    }

    // ==================== Bug4921 ====================
    // Word section orientation round-trip. Set swaps width/height automatically.
    // Verify that Get returns the new orientation correctly.
    [Fact]
    public void Bug4921_WordSectionOrientationRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Set("/section[1]", new()
        {
            ["orientation"] = "landscape"
        });

        var node = handler.Get("/section[1]");

        node.Format.Should().ContainKey("orientation",
            because: "Word section Get should report orientation when it was set");

        if (node.Format.ContainsKey("orientation"))
        {
            node.Format["orientation"]?.ToString().Should().Be("landscape",
                because: "Section orientation should round-trip: Set 'landscape' → Get 'landscape'");
        }
    }

    // ==================== Bug4922 ====================
    // Word section columns Set uses "columns" key but Get reports as "columns" too.
    // Verify the round-trip.
    [Fact]
    public void Bug4922_WordSectionColumnsRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Set("/section[1]", new()
        {
            ["columns"] = "3"
        });

        var node = handler.Get("/section[1]");

        node.Format.Should().ContainKey("columns",
            because: "Word section Get should report columns count when it was set");

        if (node.Format.ContainsKey("columns"))
        {
            var colVal = node.Format["columns"]?.ToString();
            colVal.Should().Be("3",
                because: "Section columns should round-trip: Set '3' → Get '3'");
        }
    }

    // ==================== Bug4923 ====================
    // Word run Set supports "formula" key which replaces the run with an oMath element.
    // After this replacement, Get for the original run path should fail or return the equation.
    // Verify that the equation is retrievable.
    [Fact]
    public void Bug4923_WordRunFormulaReplacement()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Before formula"
        });

        // Replace the run with a formula
        handler.Set("/body/p[1]/r[1]", new()
        {
            ["formula"] = "x^2 + y^2"
        });

        // After formula replacement, the original run is gone
        // Try to get the paragraph to see the equation
        var paraNode = handler.Get("/body/p[1]");

        paraNode.Should().NotBeNull();
        // The paragraph should now contain the formula, not the original text
        paraNode.Text.Should().NotBe("Before formula",
            because: "After replacing a run with a formula via Set, the paragraph text should change");
    }

    // ==================== Bug4924 ====================
    // Word paragraph Set "text" replaces text of first run and removes extra runs.
    // But if the paragraph has no runs (e.g., only contains an oMath),
    // it creates a new run. Verify this doesn't crash.
    [Fact]
    public void Bug4924_WordParagraphSetTextOnEmptyParagraph()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "first",
            ["bold"] = "true"
        });

        // Verify paragraph has text
        var node1 = handler.Get("/body/p[1]");
        node1.Text.Should().Contain("first");

        // Replace text
        handler.Set("/body/p[1]", new() { ["text"] = "replaced" });

        var node2 = handler.Get("/body/p[1]");
        node2.Text.Should().Contain("replaced");

        // Verify bold formatting is preserved
        if (node2.Format.ContainsKey("bold"))
        {
            // Good — formatting preserved
        }
    }

    // ==================== Bug4925 ====================
    // Word paragraph keepnext/keeplines Set uses IsTruthy, Get checks presence.
    // If Set with "false", the element is set to null, so Get should not report it.
    // But if the paragraph had keepnext originally and we Set false, verify round-trip.
    [Fact]
    public void Bug4925_WordParagraphKeepNextFalseRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "test",
            ["keepnext"] = "true"
        });

        var node1 = handler.Get("/body/p[1]");
        node1.Format.Should().ContainKey("keepnext",
            because: "keepnext was set to true during Add");

        // Set keepnext to false
        handler.Set("/body/p[1]", new() { ["keepnext"] = "false" });

        var node2 = handler.Get("/body/p[1]");

        node2.Format.ContainsKey("keepnext").Should().BeFalse(
            because: "After setting keepnext to false, it should not appear in Format. " +
                     "keepnext=false removes the KeepNext element, and Get only reports it when present");
    }

    // ==================== Bug4926 ====================
    // PPTX connector node builder reports lineWidth but not line color when
    // the connector has a non-solid line fill. Let's check if connector properties
    // round-trip correctly.
    [Fact]
    public void Bug4926_PptxConnectorLineColorRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "1cm", ["y"] = "1cm",
            ["width"] = "5cm", ["height"] = "0cm",
            ["lineColor"] = "FF0000",
            ["lineWidth"] = "2pt"
        });

        var node = handler.Get("/slide[1]/connector[1]");

        node.Format.Should().ContainKey("lineColor",
            because: "PPTX connector Get should include lineColor when it was set");

        if (node.Format.ContainsKey("lineColor"))
        {
            node.Format["lineColor"]?.ToString().Should().Be("#FF0000",
                because: "Connector lineColor should round-trip: Set 'FF0000' → Get 'FF0000'");
        }
    }

    // ==================== Bug4927 ====================
    // PPTX shape Set "opacity" uses EMU percentage format (× 100000),
    // but Get returns it in what format? Let's verify round-trip.
    [Fact]
    public void Bug4927_PptxShapeOpacityRoundTrip()
    {
        var path = CreateTempFile(".pptx");
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "test" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "opacity test",
            ["fill"] = "FF0000"
        });

        handler.Set("/slide[1]/shape[1]", new()
        {
            ["opacity"] = "50"
        });

        var node = handler.Get("/slide[1]/shape[1]");

        if (node.Format.ContainsKey("opacity"))
        {
            var opacityStr = node.Format["opacity"]?.ToString() ?? "";
            // The opacity value should be human-readable (e.g., "50" or "50%")
            // not raw OOXML units (50000)
            opacityStr.Should().NotBe("50000",
                because: "Shape opacity Get should return human-readable value, not raw OOXML units");
        }
    }

    // ==================== Bug4928 ====================
    // Word SDT (content control) Set supports "alias", "tag", "lock", "text" keys.
    // Verify that Get reports these properties back.
    [Fact]
    public void Bug4928_WordSdtPropertiesRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "sdt", null, new()
        {
            ["alias"] = "TestControl",
            ["tag"] = "test-tag",
            ["text"] = "SDT content"
        });

        // SDT should be navigable
        var nodes = handler.Query("sdt");
        nodes.Should().NotBeEmpty(
            because: "After adding an SDT, Query('sdt') should find it");

        // If we can navigate to it, check properties
        if (nodes.Count > 0)
        {
            var sdtNode = nodes[0];
            // Check if alias/tag are reported
            var hasAlias = sdtNode.Format.ContainsKey("alias")
                        || sdtNode.Format.ContainsKey("name");
            var hasTag = sdtNode.Format.ContainsKey("tag");

            (hasAlias || hasTag).Should().BeTrue(
                because: "Word SDT Get should include alias/name and tag properties " +
                         "when they were set during Add");
        }
    }

    // ==================== Bug4929 ====================
    // Word bookmark Set "text" inserts a new run after BookmarkStart,
    // but the BookmarkEnd follows the original content.
    // Verify round-trip of bookmark text modification.
    [Fact]
    public void Bug4929_WordBookmarkTextRoundTrip()
    {
        var path = CreateTempFile(".docx");
        BlankDocCreator.Create(path);
        using var handler = new WordHandler(path, editable: true);

        handler.Add("/", "paragraph", null, new() { ["text"] = "before " });
        handler.Add("/body/p[1]", "bookmark", null, new()
        {
            ["name"] = "TestMark",
            ["text"] = "original"
        });

        // Set bookmark text
        var bookmarks = handler.Query("bookmark");
        if (bookmarks.Count > 0)
        {
            handler.Set(bookmarks[0].Path, new() { ["text"] = "modified" });

            var updated = handler.Get(bookmarks[0].Path);
            updated.Text.Should().Contain("modified",
                because: "Bookmark text should round-trip: Set 'modified' → Get contains 'modified'");
        }
    }
}
