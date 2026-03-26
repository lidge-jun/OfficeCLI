// Black-box tests (Round 14) — verify R12 fixes + regression:
//   1. Word font size half-point rounding (AwayFromZero fix) — 10.5pt, 10.25pt
//   2. Word Set find/replace does not mutate caller's dict
//   3. Word footnote Remove cleans up body references + schema valid
//   4. Word endnote Remove cleans up body references + schema valid
//   5. Word footnote/endnote persistence after reopen
//   6. Null-value guard: Set with null value does not throw NullReferenceException
//   7. Regression: comment body cleanup (R11) still passes under new code
//   8. Word Set find/replace + remaining props applied correctly

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

public class BtBlackBoxRound14 : IDisposable
{
    private readonly List<string> _temps = new();
    private readonly ITestOutputHelper _out;

    public BtBlackBoxRound14(ITestOutputHelper output) => _out = output;

    public void Dispose()
    {
        foreach (var f in _temps)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt14_{Guid.NewGuid():N}.{ext}");
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

    // ==================== 1. Font size rounding: 10.5pt round-trips correctly ====================

    [Fact]
    public void Word_FontSize_HalfPoint_RoundTrip_10_5pt()
    {
        var path = Temp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Half-point size", ["size"] = "10.5pt" });
        }

        ValidateDocx(path, "font size 10.5pt");

        using var h2 = new WordHandler(path, editable: false);
        var para = h2.Get("/body/p[1]");
        para.Should().NotBeNull();
        _out.WriteLine($"p[1] size: {para!.Format.GetValueOrDefault("size")}");
        // 10.5pt stored as 21 half-points; readback should be 10.5pt
        para.Format.GetValueOrDefault("size")?.ToString().Should().Be("10.5pt", "10.5pt must round-trip cleanly");
    }

    [Fact]
    public void Word_FontSize_HalfPoint_RoundTrip_10_25pt()
    {
        // AwayFromZero fix: 10.25 * 2 = 20.5 -> rounds to 21 (10.5pt), not 20 (10pt)
        var path = Temp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Near half-point", ["size"] = "10.25pt" });
        }

        ValidateDocx(path, "font size 10.25pt -> 10.5pt");

        using var h2 = new WordHandler(path, editable: false);
        var para = h2.Get("/body/p[1]");
        para.Should().NotBeNull();
        var sizeStr = para!.Format.GetValueOrDefault("size")?.ToString();
        _out.WriteLine($"p[1] size from 10.25pt: {sizeStr}");
        // 10.25 * 2 = 20.5 -> AwayFromZero -> 21 half-points -> 10.5pt
        sizeStr.Should().Be("10.5pt", "10.25pt should round up to nearest half-point 10.5pt");
    }

    [Fact]
    public void Word_FontSize_Set_HalfPoint_RoundTrip()
    {
        var path = Temp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Size via Set" });
            h.Set("/body/p[1]", new() { ["size"] = "13.5pt" });
            var para = h.Get("/body/p[1]");
            para.Should().NotBeNull();
            _out.WriteLine($"p[1] size after Set: {para!.Format.GetValueOrDefault("size")}");
            para.Format.GetValueOrDefault("size")?.ToString().Should().Be("13.5pt", "13.5pt Set should round-trip correctly");
        }

        ValidateDocx(path, "font size Set 13.5pt");
    }

    // ==================== 2. Find/replace does not mutate caller's dict ====================

    [Fact]
    public void Word_Set_FindReplace_DoesNotMutateCallerDict()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Hello world" });

        var props = new Dictionary<string, string>
        {
            ["find"] = "Hello",
            ["replace"] = "Hi",
            ["scope"] = "all"
        };
        var originalCount = props.Count;

        h.Set("/", props);

        props.Count.Should().Be(originalCount, "caller's dict must not be mutated by find/replace Set");
        props.Should().ContainKey("find", "find key must remain in caller's dict");
        props.Should().ContainKey("replace", "replace key must remain in caller's dict");
        props.Should().ContainKey("scope", "scope key must remain in caller's dict");
    }

    [Fact]
    public void Word_Set_FindReplace_ActuallyReplaces()
    {
        var path = Temp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Hello world" });
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Set("/", new() { ["find"] = "Hello", ["replace"] = "Greetings" });
        }

        ValidateDocx(path, "find/replace applied");

        using var h3 = new WordHandler(path, editable: false);
        var para = h3.Get("/body/p[1]");
        para.Should().NotBeNull();
        _out.WriteLine($"p[1] text after replace: {para!.Text}");
        para.Text.Should().Contain("Greetings", "find/replace should update text");
        para.Text.Should().NotContain("Hello", "old text should be gone");
    }

    // ==================== 3. Footnote Remove cleans body references ====================

    [Fact]
    public void Word_Footnote_AddRemove_BodyRefsCleanedUp()
    {
        var path = Temp("docx");
        string fnPath;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Main text" });
            fnPath = h.Add("/body/p[1]", "footnote", null, new() { ["text"] = "This is a footnote" })!;
            _out.WriteLine($"Footnote path: {fnPath}");
        }

        // Verify footnote ref in body before Remove
        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.Descendants<FootnoteReference>().Should().HaveCount(1, "FootnoteReference should exist before Remove");
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            var rmEx = Record.Exception(() => h2.Remove(fnPath));
            rmEx.Should().BeNull("footnote Remove should not throw");
        }

        ValidateDocx(path, "after footnote Remove");

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var userFns = doc.MainDocumentPart!.FootnotesPart?.Footnotes?
                .Elements<Footnote>().Where(f => f.Id?.Value > 0).ToList() ?? new();
            userFns.Should().BeEmpty("footnote entry should be removed from FootnotesPart");

            var body = doc.MainDocumentPart.Document!.Body!;
            body.Descendants<FootnoteReference>().Should().BeEmpty("FootnoteReference must be cleaned up from body");
        }
    }

    [Fact]
    public void Word_Footnote_AddTwoRemoveOne_OtherIntact()
    {
        var path = Temp("docx");
        string fn1Path, fn2Path;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Para one" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "Para two" });
            fn1Path = h.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Footnote 1" })!;
            fn2Path = h.Add("/body/p[2]", "footnote", null, new() { ["text"] = "Footnote 2" })!;
            _out.WriteLine($"fn1={fn1Path}, fn2={fn2Path}");
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove(fn1Path);
        }

        ValidateDocx(path, "after removing footnote[1] of 2");

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var remainingFns = doc.MainDocumentPart!.FootnotesPart!.Footnotes!
                .Elements<Footnote>().Where(f => f.Id?.Value > 0).ToList();
            remainingFns.Should().HaveCount(1, "one footnote should remain");

            // Extract fn1 id from path e.g. /footnote[1]
            var fn1Id = int.Parse(fn1Path.TrimStart('/').Replace("footnote[", "").TrimEnd(']'));
            var fn2Id = int.Parse(fn2Path.TrimStart('/').Replace("footnote[", "").TrimEnd(']'));

            var body = doc.MainDocumentPart.Document!.Body!;
            body.Descendants<FootnoteReference>().Where(r => r.Id?.Value == fn1Id)
                .Should().BeEmpty("removed footnote ref must be gone");
            body.Descendants<FootnoteReference>().Where(r => r.Id?.Value == fn2Id)
                .Should().HaveCount(1, "remaining footnote ref must persist");
        }
    }

    // ==================== 4. Endnote Remove cleans body references ====================

    [Fact]
    public void Word_Endnote_AddRemove_BodyRefsCleanedUp()
    {
        var path = Temp("docx");
        string enPath;

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Main content" });
            enPath = h.Add("/body/p[1]", "endnote", null, new() { ["text"] = "This is an endnote" })!;
            _out.WriteLine($"Endnote path: {enPath}");
        }

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.Descendants<EndnoteReference>().Should().HaveCount(1, "EndnoteReference should exist before Remove");
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            var rmEx = Record.Exception(() => h2.Remove(enPath));
            rmEx.Should().BeNull("endnote Remove should not throw");
        }

        ValidateDocx(path, "after endnote Remove");

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var userEns = doc.MainDocumentPart!.EndnotesPart?.Endnotes?
                .Elements<Endnote>().Where(e => e.Id?.Value > 0).ToList() ?? new();
            userEns.Should().BeEmpty("endnote entry should be removed from EndnotesPart");

            var body = doc.MainDocumentPart.Document!.Body!;
            body.Descendants<EndnoteReference>().Should().BeEmpty("EndnoteReference must be cleaned up from body");
        }
    }

    // ==================== 5. Footnote/endnote persistence ====================

    [Fact]
    public void Word_Footnote_Persistence_AfterReopen()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Documented claim" });
            var fnPath = h.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Source: official docs" });
            _out.WriteLine($"Added footnote: {fnPath}");
        }

        ValidateDocx(path, "footnote persists");

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var fns = doc.MainDocumentPart!.FootnotesPart?.Footnotes?
                .Elements<Footnote>().Where(f => f.Id?.Value > 0).ToList() ?? new();
            fns.Should().HaveCount(1, "footnote must persist after file save/reopen");
            fns[0].InnerText.Should().Contain("Source: official docs", "footnote text preserved");

            var body = doc.MainDocumentPart.Document!.Body!;
            body.Descendants<FootnoteReference>().Should().HaveCount(1, "FootnoteReference must persist in body");
        }
    }

    [Fact]
    public void Word_Endnote_Persistence_AfterReopen()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Article text" });
            var enPath = h.Add("/body/p[1]", "endnote", null, new() { ["text"] = "See bibliography" });
            _out.WriteLine($"Added endnote: {enPath}");
        }

        ValidateDocx(path, "endnote persists");

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var ens = doc.MainDocumentPart!.EndnotesPart?.Endnotes?
                .Elements<Endnote>().Where(e => e.Id?.Value > 0).ToList() ?? new();
            ens.Should().HaveCount(1, "endnote must persist after file save/reopen");
            ens[0].InnerText.Should().Contain("See bibliography", "endnote text preserved");

            var body = doc.MainDocumentPart.Document!.Body!;
            body.Descendants<EndnoteReference>().Should().HaveCount(1, "EndnoteReference must persist in body");
        }
    }

    // ==================== 6. Null value guard in Set ====================

    [Fact]
    public void Word_Set_NullValue_DoesNotThrow()
    {
        var path = Temp("docx");
        using var h = new WordHandler(path, editable: true);
        h.Add("/body", "paragraph", null, new() { ["text"] = "Test para" });

        // Passing null value for a known property should not throw NullReferenceException
        var ex = Record.Exception(() =>
            h.Set("/body/p[1]", new Dictionary<string, string> { ["color"] = null! }));
        ex.Should().BeNull("null property value must not cause NullReferenceException");
    }

    [Fact]
    public void Excel_Set_NullValue_DoesNotThrow()
    {
        var path = Temp("xlsx");
        using var h = new ExcelHandler(path, editable: true);
        h.Set("/Sheet1/A1", new() { ["value"] = "test" });

        var ex = Record.Exception(() =>
            h.Set("/Sheet1/A1", new Dictionary<string, string> { ["color"] = null! }));
        ex.Should().BeNull("null property value in Excel Set must not throw NullReferenceException");
    }

    [Fact]
    public void Pptx_Set_NullValue_DoesNotThrow()
    {
        var path = Temp("pptx");
        using var h = new PowerPointHandler(path, editable: true);
        h.Add("/", "slide", null, new() { ["title"] = "Null test" });
        h.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape" });

        var ex = Record.Exception(() =>
            h.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["fill"] = null! }));
        ex.Should().BeNull("null property value in PPTX Set must not throw NullReferenceException");
    }

    // ==================== 7. Regression: comment body cleanup (R11) still works ====================

    [Fact]
    public void Word_RemoveComment_Regression_BodyRefsRemoved()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Regression para" });
            h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Regression comment", ["author"] = "Tester" });
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove("/comments/comment[1]");
        }

        ValidateDocx(path, "regression comment removal");

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.Descendants<CommentRangeStart>().Should().BeEmpty("CommentRangeStart must be removed (R11 regression)");
            body.Descendants<CommentRangeEnd>().Should().BeEmpty("CommentRangeEnd must be removed (R11 regression)");
            body.Descendants<CommentReference>().Should().BeEmpty("CommentReference must be removed (R11 regression)");
        }
    }

    // ==================== 8. Find/replace with remaining props applied ====================

    [Fact]
    public void Word_Set_FindReplace_WithRemainingPropsDoesNotMutate()
    {
        var path = Temp("docx");
        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Old text here" });
        }

        var props = new Dictionary<string, string>
        {
            ["find"] = "Old text",
            ["replace"] = "New text",
        };
        var snapshot = props.ToDictionary(kv => kv.Key, kv => kv.Value);

        using (var h2 = new WordHandler(path, editable: true))
        {
            var ex = Record.Exception(() => h2.Set("/", props));
            ex.Should().BeNull("find/replace Set must not throw");
        }

        // Caller dict unchanged
        props.Count.Should().Be(snapshot.Count, "caller dict count unchanged");
        foreach (var kv in snapshot)
            props[kv.Key].Should().Be(kv.Value, $"key '{kv.Key}' must be unchanged in caller dict");

        // Replacement applied
        using var h3 = new WordHandler(path, editable: false);
        var para = h3.Get("/body/p[1]");
        para.Should().NotBeNull();
        para!.Text.Should().Contain("New text", "replacement should be applied");
    }
}
