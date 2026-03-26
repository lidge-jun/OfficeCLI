// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for template merge: create template with {{key}} placeholders,
/// merge with JSON data, verify replacements and unresolved tracking.
/// </summary>
public class TemplateMergeTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string TempFile(string ext)
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
        {
            if (File.Exists(f)) File.Delete(f);
        }
    }

    // ==================== ParseMergeData ====================

    [Fact]
    public void ParseMergeData_InlineJson_ParsesCorrectly()
    {
        var data = TemplateMerger.ParseMergeData("{\"name\":\"Alice\",\"age\":\"30\"}");
        data.Should().ContainKey("name").WhoseValue.Should().Be("Alice");
        data.Should().ContainKey("age").WhoseValue.Should().Be("30");
    }

    [Fact]
    public void ParseMergeData_JsonFile_ReadsFromFile()
    {
        var jsonFile = TempFile(".json");
        File.WriteAllText(jsonFile, "{\"company\":\"Acme\",\"year\":\"2025\"}");

        var data = TemplateMerger.ParseMergeData(jsonFile);
        data.Should().ContainKey("company").WhoseValue.Should().Be("Acme");
        data.Should().ContainKey("year").WhoseValue.Should().Be("2025");
    }

    [Fact]
    public void ParseMergeData_InvalidJson_Throws()
    {
        var act = () => TemplateMerger.ParseMergeData("not valid json");
        act.Should().Throw<Exception>();
    }

    // ==================== DOCX Merge ====================

    [Fact]
    public void Docx_Merge_ReplacesPlaceholders()
    {
        // 1. Create template with placeholders
        var template = TempFile(".docx");
        BlankDocCreator.Create(template);
        using (var handler = new WordHandler(template, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello {{name}}, welcome to {{company}}!" });
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Your ID is {{id}}." });
        }

        // 2. Merge
        var output = TempFile(".docx");
        var result = TemplateMerger.Merge(template, output, new Dictionary<string, string>
        {
            ["name"] = "Alice",
            ["company"] = "Acme Corp",
            ["id"] = "12345"
        });

        // 3. Verify replacements
        using var reader = new WordHandler(output, editable: false);
        var p1 = reader.Get("/body/p[1]");
        p1.Text.Should().Contain("Hello Alice");
        p1.Text.Should().Contain("welcome to Acme Corp!");

        var p2 = reader.Get("/body/p[2]");
        p2.Text.Should().Contain("Your ID is 12345.");

        // 4. No unresolved
        result.UnresolvedPlaceholders.Should().BeEmpty();
    }

    [Fact]
    public void Docx_Merge_TracksUnresolvedPlaceholders()
    {
        var template = TempFile(".docx");
        BlankDocCreator.Create(template);
        using (var handler = new WordHandler(template, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Dear {{name}}, your order {{orderId}} is ready." });
        }

        var output = TempFile(".docx");
        var result = TemplateMerger.Merge(template, output, new Dictionary<string, string>
        {
            ["name"] = "Bob"
            // orderId intentionally missing
        });

        // name should be replaced
        using var reader = new WordHandler(output, editable: false);
        var p1 = reader.Get("/body/p[1]");
        p1.Text.Should().Contain("Dear Bob");

        // orderId should be unresolved
        result.UnresolvedPlaceholders.Should().Contain("orderId");
    }

    [Fact]
    public void Docx_Merge_PreservesTemplateUnchanged()
    {
        var template = TempFile(".docx");
        BlankDocCreator.Create(template);
        using (var handler = new WordHandler(template, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello {{name}}!" });
        }

        var output = TempFile(".docx");
        TemplateMerger.Merge(template, output, new Dictionary<string, string> { ["name"] = "Alice" });

        // Template should still have placeholder
        using var templateReader = new WordHandler(template, editable: false);
        var p = templateReader.Get("/body/p[1]");
        p.Text.Should().Contain("{{name}}");
    }

    // ==================== XLSX Merge ====================

    [Fact]
    public void Xlsx_Merge_ReplacesPlaceholders()
    {
        // 1. Create template with placeholders in cells
        var template = TempFile(".xlsx");
        BlankDocCreator.Create(template);
        using (var handler = new ExcelHandler(template, editable: true))
        {
            handler.Set("/Sheet1/A1", new() { ["value"] = "Name: {{name}}" });
            handler.Set("/Sheet1/B1", new() { ["value"] = "Company: {{company}}" });
            handler.Set("/Sheet1/A2", new() { ["value"] = "{{greeting}}" });
        }

        // 2. Merge
        var output = TempFile(".xlsx");
        var result = TemplateMerger.Merge(template, output, new Dictionary<string, string>
        {
            ["name"] = "Charlie",
            ["company"] = "TechCo",
            ["greeting"] = "Welcome aboard!"
        });

        // 3. Verify replacements
        using var reader = new ExcelHandler(output, editable: false);
        var a1 = reader.Get("/Sheet1/A1");
        a1.Text.Should().Be("Name: Charlie");

        var b1 = reader.Get("/Sheet1/B1");
        b1.Text.Should().Be("Company: TechCo");

        var a2 = reader.Get("/Sheet1/A2");
        a2.Text.Should().Be("Welcome aboard!");

        // 4. No unresolved
        result.UnresolvedPlaceholders.Should().BeEmpty();
    }

    [Fact]
    public void Xlsx_Merge_TracksUnresolvedPlaceholders()
    {
        var template = TempFile(".xlsx");
        BlankDocCreator.Create(template);
        using (var handler = new ExcelHandler(template, editable: true))
        {
            handler.Set("/Sheet1/A1", new() { ["value"] = "{{name}} - {{missing}}" });
        }

        var output = TempFile(".xlsx");
        var result = TemplateMerger.Merge(template, output, new Dictionary<string, string>
        {
            ["name"] = "Dave"
        });

        using var reader = new ExcelHandler(output, editable: false);
        var a1 = reader.Get("/Sheet1/A1");
        a1.Text.Should().Contain("Dave");
        a1.Text.Should().Contain("{{missing}}");

        result.UnresolvedPlaceholders.Should().Contain("missing");
    }

    // ==================== PPTX Merge ====================

    [Fact]
    public void Pptx_Merge_ReplacesShapeText()
    {
        // 1. Create template
        var template = TempFile(".pptx");
        BlankDocCreator.Create(template);
        using (var handler = new PowerPointHandler(template, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello {{name}}!" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Department: {{dept}}" });
        }

        // 2. Merge
        var output = TempFile(".pptx");
        var result = TemplateMerger.Merge(template, output, new Dictionary<string, string>
        {
            ["name"] = "Eve",
            ["dept"] = "Engineering"
        });

        // 3. Verify
        using var reader = new PowerPointHandler(output, editable: false);
        var shape1 = reader.Get("/slide[1]/shape[1]");
        shape1.Text.Should().Contain("Hello Eve!");

        var shape2 = reader.Get("/slide[1]/shape[2]");
        shape2.Text.Should().Contain("Department: Engineering");

        result.UnresolvedPlaceholders.Should().BeEmpty();
    }

    [Fact]
    public void Pptx_Merge_TracksUnresolvedPlaceholders()
    {
        var template = TempFile(".pptx");
        BlankDocCreator.Create(template);
        using (var handler = new PowerPointHandler(template, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "{{title}} by {{author}}" });
        }

        var output = TempFile(".pptx");
        var result = TemplateMerger.Merge(template, output, new Dictionary<string, string>
        {
            ["title"] = "My Presentation"
        });

        result.UnresolvedPlaceholders.Should().Contain("author");

        using var reader = new PowerPointHandler(output, editable: false);
        var shape = reader.Get("/slide[1]/shape[1]");
        shape.Text.Should().Contain("My Presentation");
        shape.Text.Should().Contain("{{author}}");
    }

    // ==================== Multiple placeholders in same text ====================

    [Fact]
    public void Docx_Merge_MultiplePlaceholdersInSameParagraph()
    {
        var template = TempFile(".docx");
        BlankDocCreator.Create(template);
        using (var handler = new WordHandler(template, editable: true))
        {
            handler.Add("/body", "paragraph", null, new()
            {
                ["text"] = "{{first}} {{middle}} {{last}}"
            });
        }

        var output = TempFile(".docx");
        var result = TemplateMerger.Merge(template, output, new Dictionary<string, string>
        {
            ["first"] = "John",
            ["middle"] = "Q",
            ["last"] = "Public"
        });

        using var reader = new WordHandler(output, editable: false);
        var p = reader.Get("/body/p[1]");
        p.Text.Should().Be("John Q Public");
        result.UnresolvedPlaceholders.Should().BeEmpty();
    }

    // ==================== Edge cases ====================

    [Fact]
    public void Merge_NoPlaceholders_NoErrors()
    {
        var template = TempFile(".docx");
        BlankDocCreator.Create(template);
        using (var handler = new WordHandler(template, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "No placeholders here." });
        }

        var output = TempFile(".docx");
        var result = TemplateMerger.Merge(template, output, new Dictionary<string, string>
        {
            ["unused"] = "value"
        });

        using var reader = new WordHandler(output, editable: false);
        reader.Get("/body/p[1]").Text.Should().Be("No placeholders here.");
        result.UnresolvedPlaceholders.Should().BeEmpty();
    }

    [Fact]
    public void Merge_EmptyData_LeavesPlaceholders()
    {
        var template = TempFile(".docx");
        BlankDocCreator.Create(template);
        using (var handler = new WordHandler(template, editable: true))
        {
            handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello {{name}}!" });
        }

        var output = TempFile(".docx");
        var result = TemplateMerger.Merge(template, output, new Dictionary<string, string>());

        result.UnresolvedPlaceholders.Should().Contain("name");

        using var reader = new WordHandler(output, editable: false);
        reader.Get("/body/p[1]").Text.Should().Contain("{{name}}");
    }

    [Fact]
    public void Merge_TemplateNotFound_Throws()
    {
        var output = TempFile(".docx");
        var act = () => TemplateMerger.Merge("/nonexistent/template.docx", output, new Dictionary<string, string>());
        act.Should().Throw<CliException>();
    }
}
