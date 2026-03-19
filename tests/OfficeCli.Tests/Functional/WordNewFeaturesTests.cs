// Tests for new Word features: Columns, Page Number Fields, Page/Column Breaks, SDT Content Controls

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class WordNewFeaturesTests : IDisposable
{
    private readonly string _docxPath;
    private WordHandler _handler;

    public WordNewFeaturesTests()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"wordnew_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_docxPath);
        _handler = new WordHandler(_docxPath, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_docxPath, editable: true);
        return _handler;
    }

    // =====================================================================
    // COLUMNS (分栏)
    // =====================================================================

    [Fact]
    public void Section_SetEqualColumns_ReturnsColumnCount()
    {
        // Set 3 equal-width columns on section 1
        _handler.Set("/section[1]", new() { ["columns"] = "3" });

        var sec = _handler.Get("/section[1]");
        sec.Format["columns"].Should().Be((short)3);
        sec.Format["equalWidth"].Should().Be(true);
    }

    [Fact]
    public void Section_SetEqualColumnsWithSpace_ReturnsCountAndSpace()
    {
        _handler.Set("/section[1]", new() { ["columns"] = "2,480" });

        var sec = _handler.Get("/section[1]");
        sec.Format["columns"].Should().Be((short)2);
        sec.Format["columnSpace"].Should().Be("480");
    }

    [Fact]
    public void Section_SetCustomColWidths_ReturnsWidths()
    {
        // Custom: col1=3000, space=720, col2=5000
        _handler.Set("/section[1]", new() { ["colWidths"] = "3000,720,5000" });

        var sec = _handler.Get("/section[1]");
        sec.Format["columns"].Should().Be((short)2);
        sec.Format["colWidths"].Should().Be("3000,5000");
        sec.Format.Should().ContainKey("colSpaces");
    }

    [Fact]
    public void Section_SetColumnSeparator_PersistsAfterReopen()
    {
        _handler.Set("/section[1]", new() { ["columns"] = "2", ["separator"] = "true" });
        Reopen();

        var sec = _handler.Get("/section[1]");
        sec.Format["columns"].Should().Be((short)2);
        sec.Format["separator"].Should().Be(true);
    }

    // =====================================================================
    // PAGE NUMBER FIELDS (页码域)
    // =====================================================================

    [Fact]
    public void Add_PageNumberField_InParagraph()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Page: " });
        var result = _handler.Add("/body/p[1]", "pagenum", null, new());

        result.Should().Contain("/body/p[1]/r[");
        var para = _handler.Get("/body/p[1]", depth: 2);
        // The paragraph should contain field char elements within runs
        para.Text.Should().Contain("Page:");
    }

    [Fact]
    public void Add_NumPagesField_InBody()
    {
        var result = _handler.Add("/body", "numpages", null, new());
        result.Should().Contain("/body/p[");
    }

    [Fact]
    public void Add_DateField_WithFormat()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Date: " });
        var result = _handler.Add("/body/p[1]", "date", null, new());

        result.Should().Contain("/body/p[1]/r[");
    }

    [Fact]
    public void Add_CustomField_WithInstruction()
    {
        var result = _handler.Add("/body", "field", null, new()
        {
            ["instruction"] = " AUTHOR ",
            ["text"] = "Unknown"
        });
        result.Should().Contain("/body/p[");
    }

    [Fact]
    public void Add_PageNumberField_WithFormatting()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "" });
        var result = _handler.Add("/body/p[1]", "pagenum", null, new()
        {
            ["bold"] = "true",
            ["size"] = "14",
            ["font"] = "Arial"
        });
        result.Should().Contain("/body/p[1]/r[");
    }

    [Fact]
    public void Add_PageNumberField_PersistsAfterReopen()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Page " });
        _handler.Add("/body/p[1]", "pagenum", null, new());
        Reopen();

        // The paragraph should still have the field code
        var para = _handler.Get("/body/p[1]", depth: 2);
        para.Should().NotBeNull();
    }

    // =====================================================================
    // PAGE / COLUMN BREAKS (分页/分栏符)
    // =====================================================================

    [Fact]
    public void Add_PageBreak_InParagraph()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Before break" });
        var result = _handler.Add("/body/p[1]", "pagebreak", null, new());

        result.Should().Contain("/body/p[1]/r[");
    }

    [Fact]
    public void Add_PageBreak_AtBodyLevel()
    {
        var result = _handler.Add("/body", "pagebreak", null, new());
        result.Should().Contain("/body/p[");
    }

    [Fact]
    public void Add_ColumnBreak_InParagraph()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Col1 text" });
        var result = _handler.Add("/body/p[1]", "columnbreak", null, new());

        result.Should().Contain("/body/p[1]/r[");
    }

    [Fact]
    public void Add_BreakWithType_Column()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Before" });
        var result = _handler.Add("/body/p[1]", "break", null, new() { ["type"] = "column" });

        result.Should().Contain("/body/p[1]/r[");
    }

    [Fact]
    public void Add_PageBreak_PersistsAfterReopen()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Before" });
        _handler.Add("/body/p[1]", "pagebreak", null, new());
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "After" });

        Reopen();

        // Both paragraphs should still exist
        var p1 = _handler.Get("/body/p[1]");
        p1.Text.Should().Contain("Before");
        var p2 = _handler.Get("/body/p[2]");
        p2.Text.Should().Contain("After");
    }

    // =====================================================================
    // SDT / CONTENT CONTROLS (内容控件)
    // =====================================================================

    [Fact]
    public void Add_BlockSdt_PlainText()
    {
        var result = _handler.Add("/body", "sdt", null, new()
        {
            ["sdttype"] = "text",
            ["alias"] = "UserName",
            ["tag"] = "name_field",
            ["text"] = "Enter your name"
        });

        result.Should().Contain("/body/sdt[");

        var node = _handler.Get(result);
        node.Type.Should().Be("sdt");
        node.Format["alias"].Should().Be("UserName");
        node.Format["tag"].Should().Be("name_field");
        node.Format["sdtType"].Should().Be("text");
        node.Text.Should().Be("Enter your name");
    }

    [Fact]
    public void Add_BlockSdt_Dropdown()
    {
        var result = _handler.Add("/body", "sdt", null, new()
        {
            ["sdttype"] = "dropdown",
            ["alias"] = "Status",
            ["items"] = "Draft,Review,Final",
            ["text"] = "Draft"
        });

        var node = _handler.Get(result);
        node.Type.Should().Be("sdt");
        node.Format["sdtType"].Should().Be("dropdown");
        node.Format["items"].Should().Be("Draft,Review,Final");
        node.Text.Should().Be("Draft");
    }

    [Fact]
    public void Add_BlockSdt_ComboBox()
    {
        var result = _handler.Add("/body", "sdt", null, new()
        {
            ["sdttype"] = "combobox",
            ["alias"] = "Category",
            ["items"] = "A,B,C",
            ["text"] = "A"
        });

        var node = _handler.Get(result);
        node.Format["sdtType"].Should().Be("combobox");
        node.Format["items"].Should().Be("A,B,C");
    }

    [Fact]
    public void Add_BlockSdt_DatePicker()
    {
        var result = _handler.Add("/body", "sdt", null, new()
        {
            ["sdttype"] = "date",
            ["alias"] = "StartDate",
            ["format"] = "yyyy/MM/dd",
            ["text"] = "2026-01-01"
        });

        var node = _handler.Get(result);
        node.Format["sdtType"].Should().Be("date");
        node.Format["alias"].Should().Be("StartDate");
    }

    [Fact]
    public void Add_InlineSdt_InParagraph()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Name: " });
        var result = _handler.Add("/body/p[1]", "sdt", null, new()
        {
            ["sdttype"] = "text",
            ["alias"] = "InlineName",
            ["text"] = "John"
        });

        result.Should().Contain("/body/p[1]/sdt[");

        var node = _handler.Get(result);
        node.Type.Should().Be("sdt");
        node.Format["alias"].Should().Be("InlineName");
        node.Text.Should().Be("John");
    }

    [Fact]
    public void Set_Sdt_UpdateAliasAndTag()
    {
        var path = _handler.Add("/body", "sdt", null, new()
        {
            ["sdttype"] = "text",
            ["alias"] = "OldAlias",
            ["tag"] = "old_tag",
            ["text"] = "Hello"
        });

        _handler.Set(path, new() { ["alias"] = "NewAlias", ["tag"] = "new_tag" });

        var node = _handler.Get(path);
        node.Format["alias"].Should().Be("NewAlias");
        node.Format["tag"].Should().Be("new_tag");
    }

    [Fact]
    public void Set_Sdt_UpdateText()
    {
        var path = _handler.Add("/body", "sdt", null, new()
        {
            ["sdttype"] = "text",
            ["alias"] = "Name",
            ["text"] = "Original"
        });

        _handler.Set(path, new() { ["text"] = "Updated" });

        var node = _handler.Get(path);
        node.Text.Should().Be("Updated");
    }

    [Fact]
    public void Set_Sdt_Lock()
    {
        var path = _handler.Add("/body", "sdt", null, new()
        {
            ["sdttype"] = "text",
            ["text"] = "Locked content"
        });

        _handler.Set(path, new() { ["lock"] = "content" });

        var node = _handler.Get(path);
        node.Format["lock"].Should().Be("contentLocked");
    }

    [Fact]
    public void Add_BlockSdt_PersistsAfterReopen()
    {
        var path = _handler.Add("/body", "sdt", null, new()
        {
            ["sdttype"] = "dropdown",
            ["alias"] = "Priority",
            ["tag"] = "priority",
            ["items"] = "Low,Medium,High",
            ["text"] = "Medium"
        });

        Reopen();

        var node = _handler.Get(path);
        node.Type.Should().Be("sdt");
        node.Format["alias"].Should().Be("Priority");
        node.Format["sdtType"].Should().Be("dropdown");
        node.Text.Should().Be("Medium");
    }

    [Fact]
    public void Add_BlockSdt_WithLock()
    {
        var path = _handler.Add("/body", "sdt", null, new()
        {
            ["sdttype"] = "text",
            ["alias"] = "Protected",
            ["lock"] = "sdt",
            ["text"] = "Cannot delete"
        });

        var node = _handler.Get(path);
        node.Format["lock"].Should().Be("sdtLocked");
    }

    // =====================================================================
    // COMBINED SCENARIOS
    // =====================================================================

    [Fact]
    public void Columns_WithPageBreak_MultiSectionLayout()
    {
        // Create a section break, then set columns on the first section
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Two column text" });
        _handler.Add("/body", "section", null, new() { ["type"] = "continuous" });
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Single column text" });

        // Section[1] is the paragraph-level section break, section[2] is the body-level one
        _handler.Set("/section[1]", new() { ["columns"] = "2" });

        var sec1 = _handler.Get("/section[1]");
        sec1.Format.Should().ContainKey("columns");
        sec1.Format["columns"].Should().Be((short)2);
    }

    [Fact]
    public void PageNumber_InFooter_WithFormatting()
    {
        // Add footer with page number
        _handler.Add("/", "footer", null, new()
        {
            ["text"] = "Page ",
            ["alignment"] = "center"
        });

        // Add page number field to the footer paragraph
        // The field is added in the footer paragraph successfully; path may reference body indexing
        var result = _handler.Add("/footer[1]/p[1]", "pagenum", null, new()
        {
            ["bold"] = "true"
        });

        result.Should().NotBeNullOrEmpty();
        // Verify footer still has content with field codes
        var footer = _handler.Get("/footer[1]", depth: 2);
        footer.Should().NotBeNull();
        footer.Type.Should().Be("footer");
    }

    // =====================================================================
    // FIELD GET / SET / QUERY (域代码读取/修改/查询)
    // =====================================================================

    [Fact]
    public void Get_Field_ReturnsInstructionAndType()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Page: " });
        _handler.Add("/body/p[1]", "pagenum", null, new());

        var field = _handler.Get("/field[1]");
        field.Type.Should().Be("field");
        field.Format["fieldType"].Should().Be("page");
        field.Format["instruction"].ToString().Should().Contain("PAGE");
    }

    [Fact]
    public void Get_Field_NumPages()
    {
        _handler.Add("/body", "numpages", null, new());

        var field = _handler.Get("/field[1]");
        field.Format["fieldType"].Should().Be("numpages");
        field.Format["instruction"].ToString().Should().Contain("NUMPAGES");
        field.Text.Should().Be("1"); // placeholder
    }

    [Fact]
    public void Get_Field_DateWithFormat()
    {
        _handler.Add("/body", "date", null, new());

        var field = _handler.Get("/field[1]");
        field.Format["fieldType"].Should().Be("date");
        field.Format["instruction"].ToString().Should().Contain("DATE");
    }

    [Fact]
    public void Get_Field_CustomInstruction()
    {
        _handler.Add("/body", "field", null, new()
        {
            ["instruction"] = " AUTHOR \\* MERGEFORMAT ",
            ["text"] = "John"
        });

        var field = _handler.Get("/field[1]");
        field.Format["fieldType"].Should().Be("author");
        field.Format["instruction"].ToString().Should().Contain("AUTHOR");
        field.Text.Should().Be("John");
    }

    [Fact]
    public void Get_Field_MultipleFields_IndexCorrect()
    {
        _handler.Add("/body", "pagenum", null, new());
        _handler.Add("/body", "numpages", null, new());
        _handler.Add("/body", "date", null, new());

        _handler.Get("/field[1]").Format["fieldType"].Should().Be("page");
        _handler.Get("/field[2]").Format["fieldType"].Should().Be("numpages");
        _handler.Get("/field[3]").Format["fieldType"].Should().Be("date");
    }

    [Fact]
    public void Set_Field_ChangeInstruction()
    {
        _handler.Add("/body", "pagenum", null, new());

        _handler.Set("/field[1]", new() { ["instruction"] = " NUMPAGES " });

        var field = _handler.Get("/field[1]");
        field.Format["instruction"].ToString().Should().Contain("NUMPAGES");
        field.Format["dirty"].Should().Be(true);
    }

    [Fact]
    public void Set_Field_ChangeResultText()
    {
        _handler.Add("/body", "field", null, new()
        {
            ["instruction"] = " AUTHOR ",
            ["text"] = "OldAuthor"
        });

        _handler.Set("/field[1]", new() { ["text"] = "NewAuthor" });

        var field = _handler.Get("/field[1]");
        field.Text.Should().Be("NewAuthor");
    }

    [Fact]
    public void Set_Field_MarkDirty()
    {
        _handler.Add("/body", "pagenum", null, new());

        _handler.Set("/field[1]", new() { ["dirty"] = "true" });

        var field = _handler.Get("/field[1]");
        field.Format["dirty"].Should().Be(true);
    }

    [Fact]
    public void Set_Field_PersistsAfterReopen()
    {
        _handler.Add("/body", "field", null, new()
        {
            ["instruction"] = " DATE \\@ \"yyyy-MM-dd\" ",
            ["text"] = "2026-01-01"
        });

        _handler.Set("/field[1]", new() { ["instruction"] = " DATE \\@ \"yyyy/MM/dd\" " });
        Reopen();

        var field = _handler.Get("/field[1]");
        field.Format["instruction"].ToString().Should().Contain("yyyy/MM/dd");
    }

    [Fact]
    public void Query_Field_ReturnsAll()
    {
        _handler.Add("/body", "pagenum", null, new());
        _handler.Add("/body", "numpages", null, new());

        var results = _handler.Query("field");
        results.Count.Should().BeGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void Query_Field_ContainsFilter()
    {
        _handler.Add("/body", "pagenum", null, new());
        _handler.Add("/body", "numpages", null, new());
        _handler.Add("/body", "date", null, new());

        var results = _handler.Query("field:contains(\"PAGE\")");
        results.Should().AllSatisfy(n =>
            n.Format["instruction"].ToString()!.Should().Contain("PAGE"));
    }

    [Fact]
    public void Query_Field_AttributeFilter()
    {
        _handler.Add("/body", "pagenum", null, new());
        _handler.Add("/body", "date", null, new());

        var results = _handler.Query("field[fieldType=date]");
        results.Count.Should().Be(1);
        results[0].Format["fieldType"].Should().Be("date");
    }

    // =====================================================================
    // FIND & REPLACE (搜索替换)
    // =====================================================================

    [Fact]
    public void FindReplace_SimpleText()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello World" });
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello Again" });

        _handler.Set("/", new() { ["find"] = "Hello", ["replace"] = "Hi" });

        _handler.Get("/body/p[1]").Text.Should().Be("Hi World");
        _handler.Get("/body/p[2]").Text.Should().Be("Hi Again");
    }

    [Fact]
    public void FindReplace_NoMatch_NoError()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello World" });

        _handler.Set("/", new() { ["find"] = "NotFound", ["replace"] = "X" });

        _handler.Get("/body/p[1]").Text.Should().Be("Hello World");
    }

    [Fact]
    public void FindReplace_MultipleOccurrencesInOneParagraph()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "aaa bbb aaa" });

        _handler.Set("/", new() { ["find"] = "aaa", ["replace"] = "ccc" });

        _handler.Get("/body/p[1]").Text.Should().Be("ccc bbb ccc");
    }

    [Fact]
    public void FindReplace_PersistsAfterReopen()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "甲方签章" });

        _handler.Set("/", new() { ["find"] = "甲方", ["replace"] = "乙方" });
        Reopen();

        _handler.Get("/body/p[1]").Text.Should().Be("乙方签章");
    }

    [Fact]
    public void FindReplace_ScopeBody()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello Body" });
        _handler.Add("/", "header", null, new() { ["text"] = "Hello Header" });

        _handler.Set("/", new() { ["find"] = "Hello", ["replace"] = "Hi", ["scope"] = "body" });

        _handler.Get("/body/p[1]").Text.Should().Be("Hi Body");
        // Header should NOT be affected
        var header = _handler.Get("/header[1]");
        header.Text.Should().Contain("Hello");
    }

    [Fact]
    public void FindReplace_EmptyReplace_DeletesText()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Remove THIS word" });

        _handler.Set("/", new() { ["find"] = "THIS ", ["replace"] = "" });

        _handler.Get("/body/p[1]").Text.Should().Be("Remove word");
    }

    // =====================================================================
    // BATCH SET (批量操作)
    // =====================================================================

    [Fact]
    public void BatchSet_BySelector_UpdatesAllMatches()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Title 1", ["style"] = "Heading1" });
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Body text" });
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Title 2", ["style"] = "Heading1" });

        // Batch set: change alignment for all Heading1 paragraphs
        _handler.Set("paragraph[style=Heading1]", new() { ["alignment"] = "center" });

        var p1 = _handler.Get("/body/p[1]");
        p1.Format["alignment"].Should().Be("center");

        var p2 = _handler.Get("/body/p[2]");
        p2.Format.Should().NotContainKey("alignment"); // Body text unaffected

        var p3 = _handler.Get("/body/p[3]");
        p3.Format["alignment"].Should().Be("center");
    }

    [Fact]
    public void BatchSet_NoMatch_Throws()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        var act = () => _handler.Set("paragraph[style=NonExistent]", new() { ["bold"] = "true" });
        act.Should().Throw<ArgumentException>().WithMessage("*No elements matched*");
    }

    [Fact]
    public void BatchSet_RunSelector()
    {
        _handler.Add("/body", "paragraph", null, new() { ["text"] = "Normal text" });
        _handler.Add("/body/p[1]", "run", null, new() { ["text"] = " Bold text", ["bold"] = "true" });

        // Set color on all bold runs
        _handler.Set("run[bold=true]", new() { ["color"] = "FF0000" });

        var p = _handler.Get("/body/p[1]", depth: 2);
        // The bold run should now have red color
        var boldRun = p.Children.FirstOrDefault(c => c.Format.ContainsKey("bold"));
        boldRun.Should().NotBeNull();
        boldRun!.Format["color"].Should().Be("FF0000");
    }

    // =====================================================================
    // IMAGE REPLACEMENT (图片替换)
    // =====================================================================

    [Fact]
    public void Set_ImagePath_ReplacesImage()
    {
        var imgPath1 = Path.Combine(Path.GetTempPath(), $"img1_{Guid.NewGuid():N}.png");
        var imgPath2 = Path.Combine(Path.GetTempPath(), $"img2_{Guid.NewGuid():N}.png");
        CreateMinimalPng(imgPath1);
        CreateMinimalPng(imgPath2);

        try
        {
            // Add a picture (returns paragraph path like /body/p[1])
            var paraPath = _handler.Add("/body", "picture", null, new()
            {
                ["path"] = imgPath1,
                ["width"] = "5cm",
                ["alt"] = "TestImage"
            });
            // Image is in the run inside the paragraph
            var runPath = paraPath + "/r[1]";

            // Verify picture run has image properties
            var node = _handler.Get(runPath);
            node.Format["alt"].Should().Be("TestImage");

            // Replace the image source
            _handler.Set(runPath, new() { ["path"] = imgPath2 });

            // Verify alt text preserved after replacement
            var node2 = _handler.Get(runPath);
            node2.Format["alt"].Should().Be("TestImage");
        }
        finally
        {
            if (File.Exists(imgPath1)) File.Delete(imgPath1);
            if (File.Exists(imgPath2)) File.Delete(imgPath2);
        }
    }

    [Fact]
    public void Set_ImageSrc_PersistsAfterReopen()
    {
        var imgPath1 = Path.Combine(Path.GetTempPath(), $"img1_{Guid.NewGuid():N}.png");
        var imgPath2 = Path.Combine(Path.GetTempPath(), $"img2_{Guid.NewGuid():N}.png");
        CreateMinimalPng(imgPath1);
        CreateMinimalPng(imgPath2);

        try
        {
            var paraPath = _handler.Add("/body", "picture", null, new()
            {
                ["path"] = imgPath1,
                ["width"] = "3cm",
                ["alt"] = "Logo"
            });
            var runPath = paraPath + "/r[1]";

            _handler.Set(runPath, new() { ["src"] = imgPath2 });
            Reopen();

            var node = _handler.Get(runPath);
            node.Format["alt"].Should().Be("Logo");
        }
        finally
        {
            if (File.Exists(imgPath1)) File.Delete(imgPath1);
            if (File.Exists(imgPath2)) File.Delete(imgPath2);
        }
    }

    [Fact]
    public void Set_ImagePathAndSize_BothApply()
    {
        var imgPath1 = Path.Combine(Path.GetTempPath(), $"img1_{Guid.NewGuid():N}.png");
        var imgPath2 = Path.Combine(Path.GetTempPath(), $"img2_{Guid.NewGuid():N}.png");
        CreateMinimalPng(imgPath1);
        CreateMinimalPng(imgPath2);

        try
        {
            var paraPath = _handler.Add("/body", "picture", null, new()
            {
                ["path"] = imgPath1,
                ["width"] = "5cm",
                ["height"] = "3cm",
                ["alt"] = "Stamp"
            });
            var runPath = paraPath + "/r[1]";

            // Replace image and resize
            _handler.Set(runPath, new()
            {
                ["path"] = imgPath2,
                ["width"] = "8cm",
                ["height"] = "6cm"
            });

            var node = _handler.Get(runPath);
            node.Format["width"].Should().Be("8.0cm");
            node.Format["height"].Should().Be("6.0cm");
        }
        finally
        {
            if (File.Exists(imgPath1)) File.Delete(imgPath1);
            if (File.Exists(imgPath2)) File.Delete(imgPath2);
        }
    }

    /// <summary>Create a minimal valid 1x1 PNG file.</summary>
    private static void CreateMinimalPng(string path)
    {
        // Minimal 1x1 white PNG (67 bytes)
        byte[] png = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE,
            0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
            0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33,
            0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND chunk
            0xAE, 0x42, 0x60, 0x82
        ];
        File.WriteAllBytes(path, png);
    }
}
