// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests that prove bugs found during Agent A Round 2 deep testing.
/// These tests are expected to FAIL until the bugs are fixed.
///
/// BUG-2: Word run path inconsistency — parent node lists comment-reference runs in children,
///        but direct path navigation filters them out, causing index mismatch.
/// BUG-4: Adding footer after header causes OOXML schema validation error due to
///        footerReference being prepended before headerReference in sectPr.
/// BUG-5: AddRun does not accept "strikethrough" property alias (only "strike"),
///        but Set accepts both — inconsistent API.
/// BUG-6: Set pageWidth/pageHeight uses SafeParseUint which rejects unit-qualified
///        values like "21cm" — should support cm/in/pt units.
/// </summary>
public class AgentFeedbackBugTests_Round2 : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public AgentFeedbackBugTests_Round2()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    // ==================== BUG-2: Run path inconsistency with comments ====================

    /// <summary>
    /// BUG-2: When a paragraph contains a comment, the comment creates a run with
    /// CommentReference. The parent paragraph's children list (via GetAllRuns using
    /// Descendants&lt;Run&gt;) includes this comment-reference run in its count.
    /// But NavigateToElement filters out runs with CommentReference when resolving
    /// "r[N]" paths. This means the last run index reported by the parent is
    /// unreachable via direct path navigation.
    ///
    /// Steps: Create paragraph with text, add a comment to it (which inserts a
    /// CommentReference run), then verify all child run paths are navigable.
    /// </summary>
    [Fact]
    public void Bug2_RunPathsShouldBeConsistentAfterAddingComment()
    {
        // Create a paragraph with some text runs
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello World" });
        _handler.Add("/body/p[1]", "run", null, new() { ["text"] = " More text" });

        // Add a comment to the paragraph — this inserts a CommentReference run
        _handler.Add("/body/p[1]", "comment", null, new() { ["text"] = "Test comment", ["author"] = "Tester" });

        // Get the paragraph node to see how many children it reports
        var paraNode = _handler.Get("/body/p[1]");
        paraNode.Should().NotBeNull();

        // The paragraph should have children (runs) — type is "run" in ElementToNode
        var runChildren = paraNode.Children?.Where(c => c.Type == "run").ToList();
        runChildren.Should().NotBeNull();
        runChildren!.Count.Should().BeGreaterThan(0, "paragraph should have run children");

        // Every run path reported by the parent should be navigable
        foreach (var child in runChildren)
        {
            var act = () => _handler.Get(child.Path);
            act.Should().NotThrow(
                $"run at path '{child.Path}' is listed as a child but should be navigable");
        }
    }

    // ==================== BUG-4: Footer reference schema order ====================

    /// <summary>
    /// BUG-4: When adding a footer after a header has already been added, the
    /// FooterReference is prepended to sectPr, placing it BEFORE the HeaderReference.
    /// The OOXML schema (CT_SectPr) requires headerReference elements to come before
    /// footerReference elements. Using PrependChild for both means the last-added
    /// element ends up first, violating the required element ordering.
    ///
    /// While the Open XML SDK validator may not always catch this, the XML structure
    /// is non-conformant and some consumers (e.g. strict-mode validators, LibreOffice)
    /// may reject it.
    ///
    /// Steps: Add a header, then add a footer, then verify sectPr child ordering.
    /// </summary>
    [Fact]
    public void Bug4_AddFooterAfterHeader_SectPrChildrenShouldBeInSchemaOrder()
    {
        // Add a header first
        _handler.Add("/", "header", null, new() { ["text"] = "My Header" });

        // Add a footer — this will prepend footerReference before headerReference
        _handler.Add("/", "footer", null, new() { ["text"] = "My Footer" });

        // Save and reopen to ensure XML is persisted, then close handler
        // so we can open with raw SDK
        _handler.Dispose();

        // Access the sectPr to check element ordering
        using var doc = WordprocessingDocument.Open(_path, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var sectPr = body.GetFirstChild<SectionProperties>();
        sectPr.Should().NotBeNull("sectPr should exist after adding header and footer");

        // Collect all headerReference and footerReference children in order
        var refs = sectPr!.ChildElements
            .Where(e => e is HeaderReference || e is FooterReference)
            .ToList();

        // Find the position of the last headerReference and the first footerReference
        int lastHeaderPos = -1;
        int firstFooterPos = int.MaxValue;
        for (int i = 0; i < refs.Count; i++)
        {
            if (refs[i] is HeaderReference)
                lastHeaderPos = i;
            if (refs[i] is FooterReference && i < firstFooterPos)
                firstFooterPos = i;
        }

        // Per OOXML spec (CT_SectPr), all headerReference elements must come
        // before all footerReference elements
        lastHeaderPos.Should().BeLessThan(firstFooterPos,
            "all headerReference elements must precede footerReference elements in sectPr " +
            "per OOXML schema order, but PrependChild placed footerReference before headerReference");

        doc.Dispose();
        // Reopen handler so Dispose() won't fail
        _handler = new WordHandler(_path, editable: true);
    }

    // ==================== BUG-5: strikethrough not accepted in Add ====================

    /// <summary>
    /// BUG-5: AddRun only checks for the property key "strike", not "strikethrough".
    /// But Set accepts both "strike" and "strikethrough". The Get output uses "strike"
    /// as the canonical key. Users who pass "strikethrough=true" during Add will find
    /// the property silently ignored.
    ///
    /// Steps: Add a run with strikethrough=true, then verify the strike format is set.
    /// </summary>
    [Fact]
    public void Bug5_AddRunWithStrikethrough_ShouldApplyStrike()
    {
        // Add a paragraph first
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Base text" });

        // Add a run with "strikethrough" property (not "strike")
        _handler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "Struck text",
            ["strikethrough"] = "true"
        });

        // Get the run and check that strike is applied
        var runNode = _handler.Get("/body/p[1]/r[2]");
        runNode.Should().NotBeNull();
        runNode.Format.Should().ContainKey("strike",
            "AddRun should accept 'strikethrough' as an alias for 'strike', " +
            "consistent with the Set handler which accepts both");
    }

    /// <summary>
    /// BUG-5 variant: Same issue exists in AddParagraph when creating text with properties.
    /// </summary>
    [Fact]
    public void Bug5_AddParagraphWithStrikethrough_ShouldApplyStrike()
    {
        // Add a paragraph with "strikethrough" property
        _handler.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Struck paragraph text",
            ["strikethrough"] = "true"
        });

        // Get the run inside the paragraph
        var runNode = _handler.Get("/body/p[1]/r[1]");
        runNode.Should().NotBeNull();
        runNode.Format.Should().ContainKey("strike",
            "AddParagraph should accept 'strikethrough' as an alias for 'strike', " +
            "consistent with the Set handler which accepts both");
    }

    // ==================== BUG-6: pageWidth rejects unit-qualified values ====================

    /// <summary>
    /// BUG-6: Set pageWidth uses SafeParseUint which only accepts raw integers.
    /// Values like "21cm" or "8.5in" are rejected with "Expected a non-negative integer".
    /// According to CLAUDE.md, EMU-based properties should support cm/in/pt unit input
    /// via ParseEmu, similar to how margins and other dimension properties work.
    ///
    /// Steps: Set pageWidth to "21cm" and verify it doesn't throw.
    /// </summary>
    [Fact]
    public void Bug6_SetPageWidth_ShouldAcceptCmUnit()
    {
        // Setting pageWidth with cm unit should work, not throw
        var act = () => _handler.Set("/", new() { ["pageWidth"] = "21cm" });
        act.Should().NotThrow(
            "pageWidth should accept unit-qualified values like '21cm', " +
            "but SafeParseUint rejects non-integer input");
    }

    /// <summary>
    /// BUG-6 variant: Same issue for pageHeight.
    /// </summary>
    [Fact]
    public void Bug6_SetPageHeight_ShouldAcceptInchUnit()
    {
        var act = () => _handler.Set("/", new() { ["pageHeight"] = "11in" });
        act.Should().NotThrow(
            "pageHeight should accept unit-qualified values like '11in', " +
            "but SafeParseUint rejects non-integer input");
    }

    /// <summary>
    /// BUG-6 variant: pageWidth with pt unit should also work.
    /// </summary>
    [Fact]
    public void Bug6_SetPageWidth_ShouldAcceptPtUnit()
    {
        var act = () => _handler.Set("/", new() { ["pageWidth"] = "595pt" });
        act.Should().NotThrow(
            "pageWidth should accept unit-qualified values like '595pt', " +
            "but SafeParseUint rejects non-integer input");
    }
}
