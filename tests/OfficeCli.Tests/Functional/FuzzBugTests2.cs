// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

// Bug confirmation tests for bugs found by Fuzz framework in Round 3.
// F41: Font size "1,000" (comma thousands separator) silently accepted as 1.0
// F42: PPTX Get shape[-1] returns a DocumentNode instead of null/error
// F43: Excel Get /Sheet1/A0, /Sheet1/ZZZ99999, /Sheet1/00 return DocumentNode instead of null/error
// F44: PPTX/DOCX Get "" (empty path) returns a DocumentNode instead of null/error

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class FuzzBugTests2 : IDisposable
{
    private readonly string _pptxPath;
    private readonly string _xlsxPath;
    private readonly string _docxPath;

    public FuzzBugTests2()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"fuzz2_{Guid.NewGuid():N}.pptx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"fuzz2_{Guid.NewGuid():N}.xlsx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"fuzz2_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_docxPath);

        using var pptx = new PowerPointHandler(_pptxPath, editable: true);
        pptx.Add("/", "slide", null, new());
        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello", ["x"] = "2cm", ["y"] = "2cm",
            ["width"] = "10cm", ["height"] = "3cm"
        });
    }

    public void Dispose()
    {
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    // ==================== F41: Font size "1,000" silently accepted ====================
    // Bug: double.TryParse("1,000", InvariantCulture) returns true (parses as 1).
    // Fix: reject strings containing ',' that don't represent a valid decimal number.

    [Fact]
    public void F41_Pptx_SetShapeSize_CommaThousandsSeparator_ThrowsArgumentException()
    {
        using var handler = new PowerPointHandler(_pptxPath, editable: true);
        var act = () => handler.Set("/slide[1]/shape[1]", new() { ["size"] = "1,000" });
        act.Should().Throw<ArgumentException>(
            "\"1,000\" with comma thousands separator is ambiguous and should be rejected with a clear error");
    }

    [Fact]
    public void F41_Docx_SetRunSize_CommaThousandsSeparator_ThrowsArgumentException()
    {
        using var handler = new WordHandler(_docxPath, editable: true);
        handler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });
        var act = () => handler.Set("/body/p[1]/r[1]", new() { ["size"] = "1,000" });
        act.Should().Throw<ArgumentException>(
            "\"1,000\" with comma thousands separator is ambiguous and should be rejected with a clear error");
    }

    // ==================== F42: PPTX Get shape[-1] returns DocumentNode ====================
    // Bug: negative index in path resolves to a real shape (Python-style negative indexing not intended).
    // Fix: negative indices should return null or throw ArgumentException.

    [Fact]
    public void F42_Pptx_Get_NegativeShapeIndex_ReturnsNullOrThrows()
    {
        using var handler = new PowerPointHandler(_pptxPath, editable: false);
        DocumentNode result = null;
        var threw = false;
        try
        {
            result = handler.Get("/slide[1]/shape[-1]");
        }
        catch (ArgumentException)
        {
            threw = true;
        }

        if (!threw)
            result.Should().BeNull(
                "path '/slide[1]/shape[-1]' uses a negative index — should return null, not resolve to a real shape");
    }

    // ==================== F43: Excel Get invalid cell refs return DocumentNode ====================
    // Bug: /Sheet1/A0 (row 0, invalid in Excel), /Sheet1/ZZZ99999, /Sheet1/00 all return DocumentNodes.
    // Fix: invalid cell references should return null or throw ArgumentException.

    [Theory]
    [InlineData("/Sheet1/A0")]       // row 0 is invalid in Excel (1-based)
    [InlineData("/Sheet1/ZZZ99999")] // very large column reference
    [InlineData("/Sheet1/00")]       // not a valid cell reference
    public void F43_Excel_Get_InvalidCellRef_ReturnsNullOrThrows(string path)
    {
        using var handler = new ExcelHandler(_xlsxPath, editable: false);
        DocumentNode result = null;
        var threw = false;
        try
        {
            result = handler.Get(path);
        }
        catch (ArgumentException)
        {
            threw = true;
        }

        if (!threw)
            result.Should().BeNull(
                $"path '{path}' is an invalid cell reference — should return null, not a DocumentNode");
    }

    // ==================== F44: Empty path returns DocumentNode ====================
    // Bug: Get("") returns a DocumentNode instead of null/error.
    // Fix: empty path is invalid and should return null or throw ArgumentException.

    [Fact]
    public void F44_Pptx_Get_EmptyPath_ReturnsNullOrThrows()
    {
        using var handler = new PowerPointHandler(_pptxPath, editable: false);
        DocumentNode result = null;
        var threw = false;
        try
        {
            result = handler.Get("");
        }
        catch (ArgumentException)
        {
            threw = true;
        }

        if (!threw)
            result.Should().BeNull("empty path \"\" is invalid — should return null, not a DocumentNode");
    }

    [Fact]
    public void F44_Docx_Get_EmptyPath_ReturnsNullOrThrows()
    {
        using var handler = new WordHandler(_docxPath, editable: false);
        DocumentNode result = null;
        var threw = false;
        try
        {
            result = handler.Get("");
        }
        catch (ArgumentException)
        {
            threw = true;
        }

        if (!threw)
            result.Should().BeNull("empty path \"\" is invalid — should return null, not a DocumentNode");
    }
}
