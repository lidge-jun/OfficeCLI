// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Failing tests for DOCX bugs found in rounds 1-10.
/// Each test documents a specific bug and should fail against the unfixed code.
/// </summary>
public class DocxRound1Tests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private (string path, WordHandler handler) CreateDoc()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return (path, new WordHandler(path, editable: true));
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 1: listStyle=bullet ignored on Add — numPr not set on paragraph
    //
    // When Add("/body", "paragraph", ...) is called with listStyle=bullet,
    // the paragraph's NumberingProperties should be set with a valid numId.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug1_ListStyleBullet_NumPrIsSetOnAdd()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Bullet item",
            ["listStyle"] = "bullet"
        });

        var node = h.Get("/body/p[1]");
        // numid should be present — if listStyle is applied, numPr is set
        node.Format.Should().ContainKey("numid",
            "listStyle=bullet should set w:numPr with a numId on the paragraph");
        var numId = node.Format["numid"];
        numId.Should().NotBeNull();
        Convert.ToInt32(numId).Should().BeGreaterThan(0,
            "numId must reference a valid numbering definition (> 0)");

        // The listStyle key itself should be reported back
        node.Format.Should().ContainKey("listStyle",
            "Get should report listStyle=bullet based on the numFmt");
        node.Format["listStyle"].ToString().Should().Be("bullet");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 2: listStyle=numbered ignored on Add — numPr not set on paragraph
    //
    // Same as Bug 1 but for ordered/numbered lists.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug2_ListStyleNumbered_NumPrIsSetOnAdd()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Numbered item",
            ["listStyle"] = "numbered"
        });

        var node = h.Get("/body/p[1]");
        node.Format.Should().ContainKey("numid",
            "listStyle=numbered should set w:numPr with a numId on the paragraph");
        var numId = node.Format["numid"];
        Convert.ToInt32(numId).Should().BeGreaterThan(0,
            "numId must reference a valid numbering definition (> 0)");

        node.Format.Should().ContainKey("listStyle",
            "Get should report listStyle for numbered paragraphs");
        node.Format["listStyle"].ToString().Should().Be("ordered",
            "a decimal-format list should be reported as 'ordered'");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 3: Indent property key casing is inconsistent between Set (input) and Get (output).
    //
    // Set/Add accept lowercase keys: "firstlineindent", "leftindent", "hangingindent".
    // Get returns camelCase keys: "firstLineIndent", "leftIndent", "hangingIndent".
    //
    // This inconsistency means callers cannot use the same key name for Set and Get.
    // The canonical Set key "firstlineindent" differs from the Get key "firstLineIndent".
    //
    // The fix should make Set also accept the camelCase canonical key, or make Get
    // return the lowercase key to match what Set accepts. Either way, they must agree.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug3_FirstLineIndent_SetAcceptsCanonicalCamelCaseKey()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Indented paragraph"
        });

        // Use the camelCase key that Get returns — Set should also accept this
        // BUG: Set only handles "firstlineindent" (lowercase), not "firstLineIndent" (camelCase).
        // After fixing, both key forms should work in Set.
        h.Set("/body/p[1]", new Dictionary<string, string>
        {
            ["firstLineIndent"] = "720"
        });

        var node = h.Get("/body/p[1]");
        node.Format.Should().ContainKey("firstLineIndent",
            "Get should return 'firstLineIndent' after Set with camelCase 'firstLineIndent'");
        node.Format["firstLineIndent"].ToString().Should().Be("720",
            "first line indent value should round-trip when using camelCase key in Set");
    }

    [Fact]
    public void Bug3_LeftIndent_SetAcceptsCanonicalCamelCaseKey()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Left-indented paragraph"
        });

        // BUG: Set only handles "leftindent" (lowercase), not "leftIndent" (camelCase).
        h.Set("/body/p[1]", new Dictionary<string, string>
        {
            ["leftIndent"] = "720"
        });

        var node = h.Get("/body/p[1]");
        node.Format.Should().ContainKey("leftIndent",
            "Get should return 'leftIndent' after Set with camelCase 'leftIndent'");
        node.Format["leftIndent"].ToString().Should().Be("720",
            "left indent value should round-trip when using camelCase key in Set");
    }

    [Fact]
    public void Bug3_HangingIndent_SetAcceptsCanonicalCamelCaseKey()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Hanging-indented paragraph"
        });

        // BUG: Set only handles "hangingindent" (lowercase), not "hangingIndent" (camelCase).
        h.Set("/body/p[1]", new Dictionary<string, string>
        {
            ["hangingIndent"] = "360"
        });

        var node = h.Get("/body/p[1]");
        node.Format.Should().ContainKey("hangingIndent",
            "Get should return 'hangingIndent' after Set with camelCase 'hangingIndent'");
        node.Format["hangingIndent"].ToString().Should().Be("360",
            "hanging indent value should round-trip when using camelCase key in Set");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 4: vertAlign=superscript ignored on Add paragraph run
    //
    // When a paragraph is added with vertAlign=superscript, the run inside
    // should have w:vertAlign w:val="superscript". The key name used in the
    // API is "vertAlign" but the handler only checks for "superscript".
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug4_VertAlignSuperscript_IsAppliedOnAdd()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Super^text",
            ["vertAlign"] = "superscript"
        });

        // The run inside the paragraph should have superscript set
        var runNode = h.Get("/body/p[1]/r[1]");
        runNode.Format.Should().ContainKey("superscript",
            "vertAlign=superscript on Add should result in w:vertAlign=superscript on the run");
        runNode.Format["superscript"].Should().Be(true);
    }

    [Fact]
    public void Bug4_VertAlignSuperscript_ViaAddRunIsApplied()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Add a bare paragraph first
        h.Add("/body", "paragraph", null, new Dictionary<string, string>());
        // Then add a run with vertAlign=superscript
        h.Add("/body/p[1]", "run", null, new Dictionary<string, string>
        {
            ["text"] = "sup",
            ["vertAlign"] = "superscript"
        });

        var runNode = h.Get("/body/p[1]/r[1]");
        runNode.Format.Should().ContainKey("superscript",
            "vertAlign=superscript on Add run should set w:vertAlign=superscript");
        runNode.Format["superscript"].Should().Be(true);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 5: smallCaps=true (camelCase key) ignored on Add paragraph
    //
    // The Add handler checks for "smallcaps" (all lowercase). If the caller
    // passes "smallCaps" (camelCase), it is silently ignored because
    // Dictionary.TryGetValue is case-sensitive.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug5_SmallCaps_CamelCaseKey_IsAppliedOnAdd()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Pass "smallCaps" (camelCase) — matching the canonical CLAUDE.md convention
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Small caps text",
            ["smallCaps"] = "true"
        });

        var runNode = h.Get("/body/p[1]/r[1]");
        runNode.Format.Should().ContainKey("smallcaps",
            "smallCaps=true on Add should apply w:smallCaps to the run");
        runNode.Format["smallcaps"].Should().Be(true);
    }

    [Fact]
    public void Bug5_SmallCaps_LowercaseKey_IsAppliedOnAdd()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Pass lowercase "smallcaps" — the handler's expected key
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Small caps text",
            ["smallcaps"] = "true"
        });

        var runNode = h.Get("/body/p[1]/r[1]");
        runNode.Format.Should().ContainKey("smallcaps",
            "smallcaps=true on Add should apply w:smallCaps to the run");
        runNode.Format["smallcaps"].Should().Be(true);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 6: Duplicate alignment keys returned by Get on paragraph
    //
    // Navigation.cs sets both Format["alignment"] and Format["align"] for
    // the same value. The canonical key per CLAUDE.md is "alignment" only.
    // The duplicate "align" key must NOT be present.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug6_ParagraphAlignment_OnlyCanonicalKeyReturned()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Centered paragraph",
            ["alignment"] = "center"
        });

        var node = h.Get("/body/p[1]");
        node.Format.Should().ContainKey("alignment",
            "canonical key 'alignment' should be present");
        node.Format["alignment"].ToString().Should().Be("center");

        // BUG: currently both "alignment" and "align" are written
        node.Format.Should().NotContainKey("align",
            "duplicate alias 'align' must not be present — only canonical 'alignment' should be returned");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 7: Duplicate shd/fill keys returned by Get on table cell
    //
    // ReadCellProps sets both Format["shd"] and Format["fill"] for the
    // same color value. Per CLAUDE.md, only one canonical key should exist.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug7_CellFill_OnlyOneCanonicalKeyReturned()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Create a table with one cell
        h.Add("/body", "table", null, new Dictionary<string, string>
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });

        // Set the cell fill color
        h.Set("/body/tbl[1]/tr[1]/tc[1]", new Dictionary<string, string>
        {
            ["shd"] = "FF0000"
        });

        var cellNode = h.Get("/body/tbl[1]/tr[1]/tc[1]");

        // The cell should have the fill color set
        var hasShd = cellNode.Format.ContainsKey("shd");
        var hasFill = cellNode.Format.ContainsKey("fill");
        (hasShd || hasFill).Should().BeTrue("at least one fill key should be present after Set");

        // BUG: currently both "shd" and "fill" are written with the same value
        // Only one canonical key should be present.
        (hasShd && hasFill).Should().BeFalse(
            "duplicate fill keys 'shd' and 'fill' must not both be present; only one canonical key should be returned");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 8: Query("paragraph[bold=true]") returns 0 results
    //
    // MatchesParagraphAttrs does not handle "bold" as a special case.
    // It falls through to GenericXmlQuery.GetAttributeValue which looks for
    // XML attributes on the Paragraph element — but bold is a run-level
    // property in the RunProperties child. The query returns nothing.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug8_QueryParagraphBoldTrue_ReturnsMatches()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        // Add a bold paragraph
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Bold paragraph text",
            ["bold"] = "true"
        });

        // Add a non-bold paragraph
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Normal paragraph text"
        });

        var results = h.Query("paragraph[bold=true]");
        results.Should().NotBeEmpty(
            "Query('paragraph[bold=true]') should return the bold paragraph, " +
            "but currently returns 0 results because MatchesParagraphAttrs " +
            "does not check first-run bold formatting");
        results.Should().HaveCount(1,
            "only the bold paragraph should match, not the normal one");
        results[0].Text.Should().Be("Bold paragraph text");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 8b: Query("run[bold=true]") works but paragraph-level query does not
    //
    // Contrast: querying runs correctly finds bold runs, but querying paragraphs
    // by bold attribute is broken. This confirms the bug is in MatchesParagraphAttrs.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug8b_QueryRunBoldTrue_ReturnsMatchesAsBaseline()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Bold text run",
            ["bold"] = "true"
        });

        // Run-level query should work (this is the baseline)
        var runResults = h.Query("run[bold=true]");
        runResults.Should().NotBeEmpty(
            "Query('run[bold=true]') should return the bold run — " +
            "MatchesRunSelector handles bold correctly");

        // Paragraph-level query should ALSO return a result
        var paraResults = h.Query("paragraph[bold=true]");
        paraResults.Should().NotBeEmpty(
            "Query('paragraph[bold=true]') must return the paragraph whose " +
            "first run is bold — currently broken because MatchesParagraphAttrs " +
            "does not check run-level bold");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Bug 4c: vertAlign=superscript via Set on run is applied correctly
    //
    // Verify that Set with vertAlign=superscript applies the property.
    // This tests the Set path for comparison with the Add path.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug4c_VertAlignSuperscript_ViaSetOnRun_IsApplied()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "to be superscript"
        });

        // Set via vertAlign key — this is the user-facing API name
        h.Set("/body/p[1]/r[1]", new Dictionary<string, string>
        {
            ["vertAlign"] = "superscript"
        });

        var runNode = h.Get("/body/p[1]/r[1]");
        runNode.Format.Should().ContainKey("superscript",
            "Set with vertAlign=superscript should apply w:vertAlign=superscript to the run");
        runNode.Format["superscript"].Should().Be(true);
    }

    // ────────────────────────────────────────────────────────────────────────
    // Persistence verification: list style survives reopen
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void Bug1_ListStyleBullet_PersistsAfterReopen()
    {
        var (path, handler) = CreateDoc();

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Persistent bullet",
            ["listStyle"] = "bullet"
        });
        handler.Dispose();

        using var h2 = new WordHandler(path, editable: false);
        var node = h2.Get("/body/p[1]");
        node.Format.Should().ContainKey("numid",
            "bullet list paragraph should persist numPr after file reopen");
        Convert.ToInt32(node.Format["numid"]).Should().BeGreaterThan(0);
    }
}
