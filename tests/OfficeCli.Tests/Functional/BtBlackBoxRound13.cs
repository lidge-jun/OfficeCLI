// Black-box tests (Round 13) — key objectives:
//   Focus: R11 fix — Word comment Remove cleans up dangling body references
//   (CommentRangeStart, CommentRangeEnd, CommentReference/Run all removed)

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

public class BtBlackBoxRound13 : IDisposable
{
    private readonly List<string> _temps = new();
    private readonly ITestOutputHelper _out;

    public BtBlackBoxRound13(ITestOutputHelper output) => _out = output;

    public void Dispose()
    {
        foreach (var f in _temps)
            if (File.Exists(f)) try { File.Delete(f); } catch { }
    }

    private string Temp(string ext)
    {
        var p = Path.Combine(Path.GetTempPath(), $"bt13_{Guid.NewGuid():N}.{ext}");
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

    // ==================== 1. Basic: Remove comment cleans up body references ====================

    [Fact]
    public void Word_RemoveComment_BodyRefsRemoved()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Hello world" });
            h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Review this", ["author"] = "Tester" });
        }

        // Verify comment + body markers exist
        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var comments = doc.MainDocumentPart!.WordprocessingCommentsPart?.Comments
                .Elements<Comment>().ToList();
            comments.Should().HaveCount(1);
            var body = doc.MainDocumentPart.Document!.Body!;
            body.Descendants<CommentRangeStart>().Should().HaveCount(1, "CommentRangeStart should be present before Remove");
            body.Descendants<CommentRangeEnd>().Should().HaveCount(1, "CommentRangeEnd should be present before Remove");
            body.Descendants<CommentReference>().Should().HaveCount(1, "CommentReference should be present before Remove");
        }

        // Remove the comment
        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove("/comments/comment[1]");
        }

        ValidateDocx(path, "after comment removal");

        // Verify all dangling references removed
        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var comments = doc.MainDocumentPart!.WordprocessingCommentsPart?.Comments
                .Elements<Comment>().ToList() ?? new();
            comments.Should().BeEmpty("comment element should be removed from Comments part");

            var body = doc.MainDocumentPart.Document!.Body!;
            body.Descendants<CommentRangeStart>().Should().BeEmpty("CommentRangeStart must be cleaned up");
            body.Descendants<CommentRangeEnd>().Should().BeEmpty("CommentRangeEnd must be cleaned up");
            body.Descendants<CommentReference>().Should().BeEmpty("CommentReference must be cleaned up");
        }
    }

    // ==================== 2. Multiple comments — only target's refs removed ====================

    [Fact]
    public void Word_RemoveOneOfTwoComments_OnlyTargetRefsRemoved()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "First para" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "Second para" });
            h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Comment A", ["author"] = "A" });
            h.Add("/body/p[2]", "comment", null, new() { ["text"] = "Comment B", ["author"] = "B" });
        }

        // Confirm 2 comments and 2 of each marker
        string commentAId, commentBId;
        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var allComments = doc.MainDocumentPart!.WordprocessingCommentsPart!.Comments
                .Elements<Comment>().ToList();
            allComments.Should().HaveCount(2);
            commentAId = allComments[0].Id!.Value!;
            commentBId = allComments[1].Id!.Value!;

            var body = doc.MainDocumentPart.Document!.Body!;
            body.Descendants<CommentRangeStart>().Should().HaveCount(2);
            body.Descendants<CommentRangeEnd>().Should().HaveCount(2);
            body.Descendants<CommentReference>().Should().HaveCount(2);
        }

        // Remove comment[1] (Comment A)
        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove("/comments/comment[1]");
        }

        ValidateDocx(path, "after removing comment[1] of 2");

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var remaining = doc.MainDocumentPart!.WordprocessingCommentsPart!.Comments
                .Elements<Comment>().ToList();
            remaining.Should().HaveCount(1, "one comment should remain");
            remaining[0].Id!.Value.Should().Be(commentBId, "Comment B should still exist");

            var body = doc.MainDocumentPart.Document!.Body!;

            // No refs for Comment A
            body.Descendants<CommentRangeStart>().Where(r => r.Id?.Value == commentAId)
                .Should().BeEmpty("CommentRangeStart for removed comment must be gone");
            body.Descendants<CommentRangeEnd>().Where(r => r.Id?.Value == commentAId)
                .Should().BeEmpty("CommentRangeEnd for removed comment must be gone");
            body.Descendants<CommentReference>().Where(r => r.Id?.Value == commentAId)
                .Should().BeEmpty("CommentReference for removed comment must be gone");

            // Comment B refs still present
            body.Descendants<CommentRangeStart>().Where(r => r.Id?.Value == commentBId)
                .Should().HaveCount(1, "CommentRangeStart for remaining comment must persist");
            body.Descendants<CommentRangeEnd>().Where(r => r.Id?.Value == commentBId)
                .Should().HaveCount(1, "CommentRangeEnd for remaining comment must persist");
            body.Descendants<CommentReference>().Where(r => r.Id?.Value == commentBId)
                .Should().HaveCount(1, "CommentReference for remaining comment must persist");
        }
    }

    // ==================== 3. Remove all comments — body is fully clean ====================

    [Fact]
    public void Word_RemoveAllComments_BodyFullyClean()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Para one" });
            h.Add("/body", "paragraph", null, new() { ["text"] = "Para two" });
            h.Add("/body/p[1]", "comment", null, new() { ["text"] = "C1" });
            h.Add("/body/p[1]", "comment", null, new() { ["text"] = "C2" });
            h.Add("/body/p[2]", "comment", null, new() { ["text"] = "C3" });
        }

        // Remove all 3 (remove from last to first to avoid index shifting issues)
        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove("/comments/comment[3]");
            h2.Remove("/comments/comment[2]");
            h2.Remove("/comments/comment[1]");
        }

        ValidateDocx(path, "after removing all comments");

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.Descendants<CommentRangeStart>().Should().BeEmpty("no CommentRangeStart should remain");
            body.Descendants<CommentRangeEnd>().Should().BeEmpty("no CommentRangeEnd should remain");
            body.Descendants<CommentReference>().Should().BeEmpty("no CommentReference should remain");

            // Body paragraphs still intact
            var paras = body.Elements<Paragraph>().ToList();
            paras.Should().HaveCountGreaterThanOrEqualTo(2, "paragraphs must survive comment removal");
        }
    }

    // ==================== 4. Persistence: reopened file has no dangling refs ====================

    [Fact]
    public void Word_RemoveComment_PersistsAfterReopen()
    {
        var path = Temp("docx");

        using (var h = new WordHandler(path, editable: true))
        {
            h.Add("/body", "paragraph", null, new() { ["text"] = "Check persistence" });
            h.Add("/body/p[1]", "comment", null, new() { ["text"] = "Persistent comment" });
        }

        using (var h2 = new WordHandler(path, editable: true))
        {
            h2.Remove("/comments/comment[1]");
        }

        // Reopen and verify
        using var h3 = new WordHandler(path, editable: false);
        var allComments = h3.Query("comment");
        allComments.Should().BeEmpty("comment must not appear in Query after removal");

        ValidateDocx(path, "persistence check after reopen");

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.Descendants<CommentRangeStart>().Should().BeEmpty();
            body.Descendants<CommentRangeEnd>().Should().BeEmpty();
            body.Descendants<CommentReference>().Should().BeEmpty();
        }
    }
}
