// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for Word Track Changes: accept-all, reject-all, and query revision.
/// Each test programmatically creates a .docx with tracked changes using the OpenXML SDK,
/// then exercises the handler's Set / Query API.
/// </summary>
public class TrackChangesTests : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public TrackChangesTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
    }

    public void Dispose()
    {
        _handler?.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler?.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    /// <summary>
    /// Helper: gets all text from the document body by reading each paragraph.
    /// </summary>
    private string GetAllBodyText()
    {
        var paras = _handler.Query("paragraph");
        return string.Join(" ", paras.Select(p => p.Text ?? ""));
    }

    /// <summary>
    /// Helper: creates a .docx file with tracked changes (insertions, deletions, format changes).
    /// </summary>
    private void CreateDocWithTrackedChanges()
    {
        using var doc = WordprocessingDocument.Create(_path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        mainPart.Document.AppendChild(body);

        // Paragraph 1: normal text "Hello " + tracked insertion "World"
        var para1 = new Paragraph();
        var run1 = new Run(new Text("Hello ") { Space = SpaceProcessingModeValues.Preserve });
        para1.AppendChild(run1);

        var insRun = new InsertedRun
        {
            Author = new StringValue("TestAuthor"),
            Date = new DateTimeValue(new DateTime(2025, 6, 15, 10, 0, 0, DateTimeKind.Utc))
        };
        insRun.AppendChild(new Run(new Text("World")));
        para1.AppendChild(insRun);
        body.AppendChild(para1);

        // Paragraph 2: tracked deletion
        var para2 = new Paragraph();
        var delRun = new DeletedRun
        {
            Author = new StringValue("TestAuthor"),
            Date = new DateTimeValue(new DateTime(2025, 6, 15, 11, 0, 0, DateTimeKind.Utc))
        };
        delRun.AppendChild(new Run(
            new RunProperties(),
            new DeletedText("removed text") { Space = SpaceProcessingModeValues.Preserve }
        ));
        para2.AppendChild(delRun);
        body.AppendChild(para2);

        // Paragraph 3: formatting change (rPrChange)
        var para3 = new Paragraph();
        var fmtRun = new Run();
        var rPr = new RunProperties(new Bold());
        rPr.AppendChild(new RunPropertiesChange
        {
            Author = new StringValue("TestAuthor"),
            Date = new DateTimeValue(new DateTime(2025, 6, 15, 12, 0, 0, DateTimeKind.Utc)),
            PreviousRunProperties = new PreviousRunProperties()
        });
        fmtRun.RunProperties = rPr;
        fmtRun.AppendChild(new Text("formatted text"));
        para3.AppendChild(fmtRun);
        body.AppendChild(para3);

        mainPart.Document.Save();
    }

    /// <summary>
    /// Helper: creates a .docx with only an insertion tracked change.
    /// </summary>
    private void CreateDocWithInsertion(string text = "inserted")
    {
        using var doc = WordprocessingDocument.Create(_path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        mainPart.Document.AppendChild(body);

        var para = new Paragraph();
        var normalRun = new Run(new Text("base ") { Space = SpaceProcessingModeValues.Preserve });
        para.AppendChild(normalRun);

        var insRun = new InsertedRun { Author = new StringValue("Author1") };
        insRun.AppendChild(new Run(new Text(text)));
        para.AppendChild(insRun);

        body.AppendChild(para);
        mainPart.Document.Save();
    }

    /// <summary>
    /// Helper: creates a .docx with only a deletion tracked change.
    /// </summary>
    private void CreateDocWithDeletion(string text = "deleted")
    {
        using var doc = WordprocessingDocument.Create(_path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        mainPart.Document.AppendChild(body);

        var para = new Paragraph();
        var normalRun = new Run(new Text("base ") { Space = SpaceProcessingModeValues.Preserve });
        para.AppendChild(normalRun);

        var delRun = new DeletedRun { Author = new StringValue("Author1") };
        delRun.AppendChild(new Run(
            new RunProperties(),
            new DeletedText(text) { Space = SpaceProcessingModeValues.Preserve }
        ));
        para.AppendChild(delRun);

        body.AppendChild(para);
        mainPart.Document.Save();
    }

    // ==================== Query revision ====================

    [Fact]
    public void QueryRevision_ReturnsAllTrackedChanges()
    {
        CreateDocWithTrackedChanges();
        _handler = new WordHandler(_path, editable: true);

        var revisions = _handler.Query("revision");

        revisions.Should().HaveCount(3);

        // Insertion
        revisions[0].Type.Should().Be("revision");
        revisions[0].Format["revisionType"].Should().Be("insertion");
        revisions[0].Text.Should().Be("World");
        revisions[0].Format["author"].Should().Be("TestAuthor");

        // Deletion
        revisions[1].Type.Should().Be("revision");
        revisions[1].Format["revisionType"].Should().Be("deletion");
        revisions[1].Text.Should().Be("removed text");

        // Format change
        revisions[2].Type.Should().Be("revision");
        revisions[2].Format["revisionType"].Should().Be("formatChange");
        revisions[2].Text.Should().Be("formatted text");
    }

    [Fact]
    public void QueryRevision_EmptyDocument_ReturnsEmpty()
    {
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);

        var revisions = _handler.Query("revision");
        revisions.Should().BeEmpty();
    }

    [Fact]
    public void QueryRevision_ContainsFilter()
    {
        CreateDocWithTrackedChanges();
        _handler = new WordHandler(_path, editable: true);

        var revisions = _handler.Query("revision:contains(World)");
        revisions.Should().HaveCount(1);
        revisions[0].Format["revisionType"].Should().Be("insertion");
        revisions[0].Text.Should().Be("World");
    }

    // ==================== Accept All Changes ====================

    [Fact]
    public void AcceptAllChanges_InsertionIsUnwrapped()
    {
        CreateDocWithInsertion("World");
        _handler = new WordHandler(_path, editable: true);

        // Verify revisions exist
        _handler.Query("revision").Should().HaveCount(1);

        // Accept all
        _handler.Set("/", new Dictionary<string, string> { ["acceptAllChanges"] = "true" });

        // No more revisions
        _handler.Query("revision").Should().BeEmpty();

        // Text is preserved — get paragraph text
        var p1 = _handler.Get("/body/p[1]");
        p1.Text.Should().Contain("base");
        p1.Text.Should().Contain("World");
    }

    [Fact]
    public void AcceptAllChanges_DeletionIsRemoved()
    {
        CreateDocWithDeletion("old text");
        _handler = new WordHandler(_path, editable: true);

        _handler.Query("revision").Should().HaveCount(1);

        _handler.Set("/", new Dictionary<string, string> { ["acceptAllChanges"] = "true" });

        _handler.Query("revision").Should().BeEmpty();

        // Deleted text should be gone
        var p1 = _handler.Get("/body/p[1]");
        p1.Text.Should().Contain("base");
        p1.Text.Should().NotContain("old text");
    }

    [Fact]
    public void AcceptAllChanges_FormatChangeMarkerRemoved()
    {
        CreateDocWithTrackedChanges();
        _handler = new WordHandler(_path, editable: true);

        _handler.Set("/", new Dictionary<string, string> { ["acceptAllChanges"] = "true" });

        _handler.Query("revision").Should().BeEmpty();

        // The formatted text should still be there (paragraph 3)
        var allText = GetAllBodyText();
        allText.Should().Contain("formatted text");
    }

    [Fact]
    public void AcceptAllChanges_PersistsAfterReopen()
    {
        CreateDocWithTrackedChanges();
        _handler = new WordHandler(_path, editable: true);

        _handler.Set("/", new Dictionary<string, string> { ["acceptAllChanges"] = "true" });

        Reopen();

        _handler.Query("revision").Should().BeEmpty();
        var allText = GetAllBodyText();
        allText.Should().Contain("Hello");
        allText.Should().Contain("World");
        allText.Should().NotContain("removed text");
    }

    // ==================== Reject All Changes ====================

    [Fact]
    public void RejectAllChanges_InsertionIsRemoved()
    {
        CreateDocWithInsertion("World");
        _handler = new WordHandler(_path, editable: true);

        _handler.Query("revision").Should().HaveCount(1);

        _handler.Set("/", new Dictionary<string, string> { ["rejectAllChanges"] = "true" });

        _handler.Query("revision").Should().BeEmpty();

        // Inserted text should be gone
        var p1 = _handler.Get("/body/p[1]");
        p1.Text.Should().Contain("base");
        p1.Text.Should().NotContain("World");
    }

    [Fact]
    public void RejectAllChanges_DeletionIsRestored()
    {
        CreateDocWithDeletion("restored text");
        _handler = new WordHandler(_path, editable: true);

        _handler.Query("revision").Should().HaveCount(1);

        _handler.Set("/", new Dictionary<string, string> { ["rejectAllChanges"] = "true" });

        _handler.Query("revision").Should().BeEmpty();

        // Deleted text should be restored (delText -> text)
        var p1 = _handler.Get("/body/p[1]");
        p1.Text.Should().Contain("base");
        p1.Text.Should().Contain("restored text");
    }

    [Fact]
    public void RejectAllChanges_FormatChangeRestoresOriginal()
    {
        // Create doc with a run that was changed from non-bold to bold
        using (var doc = WordprocessingDocument.Create(_path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.AppendChild(body);

            var para = new Paragraph();
            var run = new Run();
            // Current: bold
            var rPr = new RunProperties(new Bold());
            // Original: not bold (empty properties)
            rPr.AppendChild(new RunPropertiesChange
            {
                Author = new StringValue("Author1"),
                PreviousRunProperties = new PreviousRunProperties()
            });
            run.RunProperties = rPr;
            run.AppendChild(new Text("test text"));
            para.AppendChild(run);
            body.AppendChild(para);
            mainPart.Document.Save();
        }

        _handler = new WordHandler(_path, editable: true);

        _handler.Set("/", new Dictionary<string, string> { ["rejectAllChanges"] = "true" });

        _handler.Query("revision").Should().BeEmpty();

        // Text preserved
        var p1 = _handler.Get("/body/p[1]");
        p1.Text.Should().Contain("test text");
    }

    [Fact]
    public void RejectAllChanges_PersistsAfterReopen()
    {
        CreateDocWithTrackedChanges();
        _handler = new WordHandler(_path, editable: true);

        _handler.Set("/", new Dictionary<string, string> { ["rejectAllChanges"] = "true" });

        Reopen();

        _handler.Query("revision").Should().BeEmpty();
        var allText = GetAllBodyText();
        // Insertions removed
        allText.Should().NotContain("World");
        // Deletions restored
        allText.Should().Contain("removed text");
    }

    // ==================== Mixed scenarios ====================

    [Fact]
    public void AcceptAllChanges_MixedDocument_AllResolved()
    {
        CreateDocWithTrackedChanges();
        _handler = new WordHandler(_path, editable: true);

        // Should have 3 revisions (insert, delete, format)
        _handler.Query("revision").Should().HaveCount(3);

        _handler.Set("/", new Dictionary<string, string> { ["acceptAllChanges"] = "true" });

        _handler.Query("revision").Should().BeEmpty();

        var allText = GetAllBodyText();
        // Insertions kept
        allText.Should().Contain("World");
        // Deletions gone
        allText.Should().NotContain("removed text");
        // Format text kept
        allText.Should().Contain("formatted text");
    }

    [Fact]
    public void QueryRevision_ChangeAlias_Works()
    {
        CreateDocWithInsertion("test");
        _handler = new WordHandler(_path, editable: true);

        // "change" alias should work same as "revision"
        var revisions = _handler.Query("change");
        revisions.Should().HaveCount(1);
        revisions[0].Format["revisionType"].Should().Be("insertion");
    }

    [Fact]
    public void AcceptAllChanges_NoChanges_NoError()
    {
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "plain" });

        // Should not throw
        _handler.Set("/", new Dictionary<string, string> { ["acceptAllChanges"] = "true" });

        _handler.Query("revision").Should().BeEmpty();
    }

    [Fact]
    public void RejectAllChanges_NoChanges_NoError()
    {
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "plain" });

        _handler.Set("/", new Dictionary<string, string> { ["rejectAllChanges"] = "true" });

        _handler.Query("revision").Should().BeEmpty();
    }

    [Fact]
    public void QueryRevision_HasDateFormat()
    {
        CreateDocWithTrackedChanges();
        _handler = new WordHandler(_path, editable: true);

        var revisions = _handler.Query("revision");
        // Insertions and deletions have dates set
        revisions[0].Format.Should().ContainKey("date");
        ((string)revisions[0].Format["date"]!).Should().Contain("2025-06-15");
    }

    [Fact]
    public void AcceptAllChanges_ParagraphPropertiesChange_Removed()
    {
        // Create doc with paragraph property change
        using (var doc = WordprocessingDocument.Create(_path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.AppendChild(body);

            var para = new Paragraph();
            var pPr = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center }
            );
            var pPrChange = new ParagraphPropertiesChange
            {
                Author = new StringValue("Author1"),
            };
            pPrChange.AppendChild(new PreviousParagraphProperties(
                new Justification { Val = JustificationValues.Left }
            ));
            pPr.AppendChild(pPrChange);
            para.ParagraphProperties = pPr;
            para.AppendChild(new Run(new Text("centered text")));
            body.AppendChild(para);
            mainPart.Document.Save();
        }

        _handler = new WordHandler(_path, editable: true);

        var revisions = _handler.Query("revision");
        revisions.Should().HaveCount(1);
        revisions[0].Format["revisionType"].Should().Be("paragraphChange");

        _handler.Set("/", new Dictionary<string, string> { ["acceptAllChanges"] = "true" });
        _handler.Query("revision").Should().BeEmpty();
    }
}
