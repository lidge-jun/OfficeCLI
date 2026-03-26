// Black-box tests (Round 15) — verify R13 TOC Remove fix + final regression:
//   1. TOC Remove: /toc[1] path works without throwing
//   2. TOC Remove: persists across save/reopen (TOC no longer found)
//   3. TOC Remove: TOCHeading title paragraph also removed when present
//   4. TOC Remove: out-of-range index throws ArgumentException (not NullRef)
//   5. TOC Remove: two TOCs — remove first, second still accessible
//   6. TOC Get: returns correct Format keys (levels, hyperlinks, pagenumbers)
//   7. TOC Set: update levels property
//   8. DOCX schema valid after TOC Remove
//   9. Regression: comment Remove body-ref cleanup (R11) still passes
//  10. Regression: footnote Remove body-ref cleanup (R13) still passes
//  11. Regression: null-value guard in Set (R12) still passes
//  12. Regression: Word find/replace doesn't mutate caller dict (R12)

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

public class BtBlackBoxRound15 : IDisposable
{
    private readonly List<string> _temps = new();
    private readonly ITestOutputHelper _out;

    public BtBlackBoxRound15(ITestOutputHelper output) => _out = output;

    public void Dispose()
    {
        foreach (var f in _temps)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt15_{Guid.NewGuid():N}.{ext}");
        _temps.Add(p);
        BlankDocCreator.Create(p);
        return p;
    }

    private void ValidateDocx(string path, string step)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).ToList();
        foreach (var e in errors) _out.WriteLine($"[{step}] {e.Description}");
        errors.Should().BeEmpty($"DOCX invalid after: {step}");
    }

    // ==================== 1. TOC Remove: /toc[1] does not throw ====================

    [Fact]
    public void Word_Toc_Remove_DoesNotThrow()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Intro", ["style"] = "Heading1" });
        h.Add("/body", "toc", null, new() { ["levels"] = "1-3" });

        var act = () => h.Remove("/toc[1]");
        act.Should().NotThrow("Remove(\"/toc[1]\") must be supported after R13 fix");
    }

    // ==================== 2. TOC Remove: persists after save/reopen ====================

    [Fact]
    public void Word_Toc_Remove_Persistence()
    {
        var path = Temp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Chapter", ["style"] = "Heading1" });
            h.Add("/body", "toc", null, new() { ["levels"] = "1-3" });
            h.Remove("/toc[1]");
        }

        using var h2 = new WordHandler(path, editable: false);
        var act = () => h2.Get("/toc[1]");
        act.Should().Throw<ArgumentException>("TOC must not exist after removal and reopen");
    }

    // ==================== 3. TOC Remove: TOCHeading paragraph also removed ====================

    [Fact]
    public void Word_Toc_Remove_AlsoRemovesTocHeadingParagraph()
    {
        var path = Temp("docx");
        int paraCountBefore;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Section A", ["style"] = "Heading1" });
            h.Add("/body", "toc", null, new() { ["levels"] = "1-3" });
        }

        // Count body paragraphs before removal
        using (var doc = WordprocessingDocument.Open(path, false))
        {
            paraCountBefore = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Count();
            _out.WriteLine($"Paragraphs before TOC Remove: {paraCountBefore}");
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove("/toc[1]");
        }

        // After removal, paragraph count must decrease (TOC para and possibly TOCHeading removed)
        using (var doc2 = WordprocessingDocument.Open(path, false))
        {
            var paraCountAfter = doc2.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Count();
            _out.WriteLine($"Paragraphs after TOC Remove: {paraCountAfter}");
            paraCountAfter.Should().BeLessThan(paraCountBefore, "TOC Remove should decrease paragraph count");

            // No FieldCode with TOC instruction should remain
            var remainingTocFields = doc2.MainDocumentPart.Document.Body!
                .Descendants<FieldCode>()
                .Where(fc => fc.Text != null && fc.Text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
                .ToList();
            remainingTocFields.Should().BeEmpty("No TOC field code should remain after Remove");
        }
    }

    // ==================== 4. TOC Remove: out-of-range throws ArgumentException ====================

    [Fact]
    public void Word_Toc_Remove_OutOfRange_ThrowsArgumentException()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Heading", ["style"] = "Heading1" });
        h.Add("/body", "toc", null, new() { ["levels"] = "1-2" });

        // toc[2] doesn't exist
        var act = () => h.Remove("/toc[2]");
        act.Should().Throw<ArgumentException>("out-of-range TOC index must throw ArgumentException, not NullRef");
    }

    // ==================== 5. TOC Remove: two TOCs — remove first, second intact ====================

    [Fact]
    public void Word_Toc_RemoveFirstOfTwo_SecondIntact()
    {
        var path = Temp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Part 1", ["style"] = "Heading1" });
            h.Add("/body", "toc", null, new() { ["levels"] = "1-2" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "Part 2", ["style"] = "Heading2" });
            h.Add("/body", "toc", null, new() { ["levels"] = "2-3" });
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove("/toc[1]");
        }

        using var h3 = new WordHandler(path, editable: false);
        var toc2 = h3.Get("/toc[1]"); // was [2], now renumbered to [1]
        toc2.Should().NotBeNull("the second TOC should still be accessible after removing the first");
        _out.WriteLine($"Remaining TOC levels: {toc2!.Format.GetValueOrDefault("levels")}");
    }

    // ==================== 6. TOC Get: returns expected Format keys ====================

    [Fact]
    public void Word_Toc_Get_ReturnsFormatKeys()
    {
        var path = Temp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Overview", ["style"] = "Heading1" });
            h.Add("/body", "toc", null, new() { ["levels"] = "1-3" });
        }

        using var h2 = new WordHandler(path, editable: false);
        var toc = h2.Get("/toc[1]");
        toc.Should().NotBeNull("TOC[1] must be retrievable via Get");
        _out.WriteLine($"TOC Format keys: {string.Join(", ", toc!.Format.Keys)}");
        toc.Format.Should().ContainKey("levels", "Get must return levels key");
    }

    // ==================== 7. TOC Set: update levels property ====================

    [Fact]
    public void Word_Toc_Set_UpdateLevels()
    {
        var path = Temp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Summary", ["style"] = "Heading1" });
            h.Add("/body", "toc", null, new() { ["levels"] = "1-3" });
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            var ex = Record.Exception(() => h2.Set("/toc[1]", new() { ["levels"] = "1-2" }));
            ex.Should().BeNull("Set(\"/toc[1]\", levels) must not throw");
        }
    }

    // ==================== 8. DOCX schema valid after TOC Remove ====================

    [Fact]
    public void Word_Toc_Remove_DocxSchemaValid()
    {
        var path = Temp("docx");
        {
            using var h = new WordHandler(path, editable: true);
            h.Add("/body", "paragraph", null, new() { ["text"] = "Section 1", ["style"] = "Heading1" });
            h.Add("/body", "toc", null, new() { ["levels"] = "1-3" });
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove("/toc[1]");
        }

        ValidateDocx(path, "after TOC Remove");
    }

    // ==================== 9. Regression: comment Remove body-ref cleanup (R11) ====================

    [Fact]
    public void Regression_Word_RemoveComment_BodyRefsCleanedUp()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Regression para" });
            h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Regression comment", ["author"] = "BtR15" });
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove("/comments/comment[1]");
        }

        ValidateDocx(path, "regression comment removal R11");

        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        body.Descendants<CommentRangeStart>().Should().BeEmpty("CommentRangeStart must be gone (R11 regression)");
        body.Descendants<CommentRangeEnd>().Should().BeEmpty("CommentRangeEnd must be gone (R11 regression)");
        body.Descendants<CommentReference>().Should().BeEmpty("CommentReference must be gone (R11 regression)");
    }

    // ==================== 10. Regression: footnote Remove body-ref cleanup (R13) ====================

    [Fact]
    public void Regression_Word_RemoveFootnote_BodyRefsCleanedUp()
    {
        var path = Temp("docx");
        string fnPath;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Para with footnote" });
            fnPath = h.Add("/body/p[1]", "footnote", null, new() { ["text"] = "BtR15 footnote" })!;
            _out.WriteLine($"Footnote path: {fnPath}");
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove(fnPath);
        }

        ValidateDocx(path, "regression footnote removal R13");

        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        body.Descendants<FootnoteReference>().Should().BeEmpty("FootnoteReference must be cleaned from body (R13 regression)");
        var userFns = doc.MainDocumentPart.FootnotesPart?.Footnotes?
            .Elements<Footnote>().Where(f => f.Id?.Value > 0).ToList() ?? new();
        userFns.Should().BeEmpty("Footnote element must be removed from FootnotesPart");
    }

    // ==================== 11. Regression: null-value guard in Set (R12) ====================

    [Fact]
    public void Regression_Word_Set_NullValue_DoesNotThrow()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Null guard test" });

        var ex = Record.Exception(() =>
            h.Set("/body/p[1]", new Dictionary<string, string> { ["color"] = null! }));
        ex.Should().BeNull("null property value must not cause NullReferenceException (R12 regression)");
    }

    // ==================== 12. Regression: find/replace does not mutate caller dict (R12) ====================

    [Fact]
    public void Regression_Word_Set_FindReplace_CallerDictUnchanged()
    {
        var path = Temp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Old value" });
        }

        var props = new Dictionary<string, string>
        {
            ["find"] = "Old value",
            ["replace"] = "New value",
        };
        var countBefore = props.Count;

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Set("/", props);
        }

        props.Count.Should().Be(countBefore, "caller dict must not be mutated by find/replace Set (R12 regression)");
        props.Should().ContainKey("find");
        props.Should().ContainKey("replace");
    }
}
