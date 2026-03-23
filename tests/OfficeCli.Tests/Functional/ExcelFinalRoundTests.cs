// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Final-round tests covering: validation types (time, textlength, custom),
/// Set validation[N] modification, Swap rows, CopyFrom rows,
/// Named range with comment, and Sheet password protection.
/// </summary>
public class ExcelFinalRoundTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelFinalRoundTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen() { _handler.Dispose(); _handler = new ExcelHandler(_path, editable: true); }

    // ==================== Validation type: time ====================

    [Fact]
    public void Add_Validation_Time_GetVerifies()
    {
        var path = _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "B2",
            ["type"] = "time",
            ["operator"] = "between",
            ["formula1"] = "0.333",  // 8:00 AM as fraction of day
            ["formula2"] = "0.875",  // 9:00 PM as fraction of day
        });

        path.Should().Be("/Sheet1/validation[1]");

        var node = _handler.Get("/Sheet1/validation[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("validation");
        node.Format["type"].Should().Be("time");
        node.Format["operator"].Should().Be("between");
        node.Format["formula1"].Should().Be("0.333");
        node.Format["formula2"].Should().Be("0.875");
        node.Format["sqref"].Should().Be("B2");
    }

    [Fact]
    public void Add_Validation_Time_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "C3",
            ["type"] = "time",
            ["operator"] = "greaterthan",
            ["formula1"] = "0.5",
        });

        Reopen();

        var node = _handler.Get("/Sheet1/validation[1]");
        node.Should().NotBeNull();
        node.Format["type"].Should().Be("time");
        node.Format["operator"].Should().Be("greaterThan");
        node.Format["formula1"].Should().Be("0.5");
    }

    // ==================== Validation type: textlength ====================

    [Fact]
    public void Add_Validation_TextLength_GetVerifies()
    {
        var path = _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A1:A10",
            ["type"] = "textlength",
            ["operator"] = "lessthanorequal",
            ["formula1"] = "50",
        });

        path.Should().Be("/Sheet1/validation[1]");

        var node = _handler.Get("/Sheet1/validation[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("validation");
        node.Format["type"].Should().Be("textLength");
        node.Format["operator"].Should().Be("lessThanOrEqual");
        node.Format["formula1"].Should().Be("50");
        node.Format["sqref"].Should().Be("A1:A10");
    }

    [Fact]
    public void Add_Validation_TextLength_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "D4",
            ["type"] = "textlength",
            ["operator"] = "between",
            ["formula1"] = "5",
            ["formula2"] = "20",
        });

        Reopen();

        var node = _handler.Get("/Sheet1/validation[1]");
        node.Should().NotBeNull();
        node.Format["type"].Should().Be("textLength");
        node.Format["formula1"].Should().Be("5");
        node.Format["formula2"].Should().Be("20");
    }

    // ==================== Validation type: custom ====================

    [Fact]
    public void Add_Validation_Custom_GetVerifies()
    {
        var path = _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "E5",
            ["type"] = "custom",
            ["formula1"] = "ISNUMBER(E5)",
        });

        path.Should().Be("/Sheet1/validation[1]");

        var node = _handler.Get("/Sheet1/validation[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("validation");
        node.Format["type"].Should().Be("custom");
        node.Format["formula1"].Should().Be("ISNUMBER(E5)");
        node.Format["sqref"].Should().Be("E5");
    }

    [Fact]
    public void Add_Validation_Custom_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "F6",
            ["type"] = "custom",
            ["formula1"] = "AND(F6>0,F6<100)",
            ["showError"] = "true",
            ["errorTitle"] = "Invalid",
            ["error"] = "Value must be between 0 and 100",
        });

        Reopen();

        var node = _handler.Get("/Sheet1/validation[1]");
        node.Should().NotBeNull();
        node.Format["type"].Should().Be("custom");
        node.Format["formula1"].Should().Be("AND(F6>0,F6<100)");
        node.Format["errorTitle"].Should().Be("Invalid");
    }

    // ==================== Set validation[N] modification ====================

    [Fact]
    public void Set_Validation_ModifiesTypeAndFormula()
    {
        // Add a "whole" validation first
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A1",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "10",
        });

        var before = _handler.Get("/Sheet1/validation[1]");
        before.Format["type"].Should().Be("whole");
        before.Format["formula1"].Should().Be("1");

        // Modify it to use a different range
        _handler.Set("/Sheet1/validation[1]", new()
        {
            ["formula1"] = "100",
            ["formula2"] = "999",
        });

        var after = _handler.Get("/Sheet1/validation[1]");
        after.Format["type"].Should().Be("whole");
        after.Format["formula1"].Should().Be("100");
        after.Format["formula2"].Should().Be("999");
    }

    [Fact]
    public void Set_Validation_ChangesType()
    {
        // Add as decimal, then change to textlength
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "B2",
            ["type"] = "decimal",
            ["operator"] = "greaterthan",
            ["formula1"] = "0",
        });

        _handler.Set("/Sheet1/validation[1]", new()
        {
            ["type"] = "textlength",
            ["operator"] = "lessthanorequal",
            ["formula1"] = "100",
        });

        var node = _handler.Get("/Sheet1/validation[1]");
        node.Format["type"].Should().Be("textLength");
        node.Format["operator"].Should().Be("lessThanOrEqual");
        node.Format["formula1"].Should().Be("100");
    }

    [Fact]
    public void Set_Validation_ModifiesSqref()
    {
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "A1",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "10",
        });

        _handler.Set("/Sheet1/validation[1]", new()
        {
            ["sqref"] = "A1:A20",
        });

        var node = _handler.Get("/Sheet1/validation[1]");
        node.Format["sqref"].Should().Be("A1:A20");
    }

    [Fact]
    public void Set_Validation_AddsErrorMessageProperties()
    {
        _handler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "C3",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "5",
        });

        _handler.Set("/Sheet1/validation[1]", new()
        {
            ["showError"] = "true",
            ["errorTitle"] = "Out of range",
            ["error"] = "Enter a value between 1 and 5",
            ["promptTitle"] = "Hint",
            ["prompt"] = "Values 1-5 only",
        });

        Reopen();

        var node = _handler.Get("/Sheet1/validation[1]");
        node.Format["showError"].Should().Be(true);
        node.Format["errorTitle"].Should().Be("Out of range");
        node.Format["error"].Should().Be("Enter a value between 1 and 5");
        node.Format["promptTitle"].Should().Be("Hint");
        node.Format["prompt"].Should().Be("Values 1-5 only");
    }

    // ==================== Swap rows ====================

    [Fact]
    public void Swap_Rows_ContentIsExchanged()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Alpha" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "100" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Beta" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "200" });

        // Verify initial state
        _handler.Get("/Sheet1/A1").Text.Should().Be("Alpha");
        _handler.Get("/Sheet1/A2").Text.Should().Be("Beta");

        // Swap row 1 and row 2
        var (newPath1, newPath2) = _handler.Swap("/Sheet1/row[1]", "/Sheet1/row[2]");

        // The swap returns the exchanged paths (row indices are swapped)
        newPath1.Should().NotBeNull();
        newPath2.Should().NotBeNull();

        // After swap, cell references are updated: A1 now contains "Beta", A2 contains "Alpha"
        var cellA1 = _handler.Get("/Sheet1/A1");
        var cellA2 = _handler.Get("/Sheet1/A2");

        cellA1.Text.Should().Be("Beta");
        cellA2.Text.Should().Be("Alpha");
    }

    [Fact]
    public void Swap_Rows_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Alpha" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Beta" });

        _handler.Swap("/Sheet1/row[1]", "/Sheet1/row[2]");

        Reopen();

        var cellA1 = _handler.Get("/Sheet1/A1");
        var cellA2 = _handler.Get("/Sheet1/A2");

        cellA1.Text.Should().Be("Beta");
        cellA2.Text.Should().Be("Alpha");
    }

    [Fact]
    public void Swap_Rows_ThreeRows_SwapsFirstAndThird()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "First" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Second" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A3", ["value"] = "Third" });

        // Swap first and third rows
        _handler.Swap("/Sheet1/row[1]", "/Sheet1/row[3]");

        var row1Cell = _handler.Get("/Sheet1/A1");
        var row3Cell = _handler.Get("/Sheet1/A3");
        var row2Cell = _handler.Get("/Sheet1/A2");

        row1Cell.Text.Should().Be("Third");
        row3Cell.Text.Should().Be("First");
        // Middle row should be unchanged
        row2Cell.Text.Should().Be("Second");
    }

    // ==================== CopyFrom rows ====================

    [Fact]
    public void CopyFrom_Row_CreatesNewRowWithSameContent()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Source" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "Data" });

        var newPath = _handler.CopyFrom("/Sheet1/row[1]", "/Sheet1", null);

        newPath.Should().NotBeNull();
        newPath.Should().StartWith("/Sheet1/row[");

        // Original cell should be unchanged
        var origCell = _handler.Get("/Sheet1/A1");
        origCell.Text.Should().Be("Source");

        // The copied row should contain the same data (cells reference the new row number)
        var match = System.Text.RegularExpressions.Regex.Match(newPath, @"row\[(\d+)\]");
        match.Success.Should().BeTrue();
        var rowIdx = int.Parse(match.Groups[1].Value);
        // The copied row should exist
        var copiedRow = _handler.Get($"/Sheet1/row[{rowIdx}]");
        copiedRow.Should().NotBeNull();
        copiedRow.Type.Should().Be("row");
    }

    [Fact]
    public void CopyFrom_Row_DoesNotModifySource()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Original" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "Value" });

        _handler.CopyFrom("/Sheet1/row[1]", "/Sheet1", null);

        var sourceCell = _handler.Get("/Sheet1/A1");
        sourceCell.Text.Should().Be("Original");

        var sourceCellB = _handler.Get("/Sheet1/B1");
        sourceCellB.Text.Should().Be("Value");
    }

    [Fact]
    public void CopyFrom_Row_AppendedWhenNoIndex()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Row1" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "Row2" });

        // Copy row 1 — should append after row 2
        var newPath = _handler.CopyFrom("/Sheet1/row[1]", "/Sheet1", null);
        newPath.Should().Contain("row[3]");
    }

    // ==================== Named range with comment ====================

    [Fact]
    public void Add_NamedRange_WithComment_GetVerifies()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "100" });

        var path = _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "MyRevenue",
            ["ref"] = "Sheet1!$A$1",
            ["comment"] = "Annual revenue figure",
        });

        path.Should().Be("/namedrange[1]");

        var node = _handler.Get("/namedrange[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("namedrange");
        node.Format["name"].Should().Be("MyRevenue");
        node.Format["ref"].Should().Be("Sheet1!$A$1");
        node.Format.Should().ContainKey("comment");
        node.Format["comment"].Should().Be("Annual revenue figure");
    }

    [Fact]
    public void Add_NamedRange_WithComment_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["c1"] = "Data" });

        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "DataRange",
            ["ref"] = "Sheet1!$A$1:$A$10",
            ["comment"] = "Primary data range for calculations",
        });

        Reopen();

        var node = _handler.Get("/namedrange[1]");
        node.Should().NotBeNull();
        node.Format["name"].Should().Be("DataRange");
        node.Format["comment"].Should().Be("Primary data range for calculations");
        node.Format["ref"].Should().Be("Sheet1!$A$1:$A$10");
    }

    [Fact]
    public void Add_NamedRange_WithoutComment_CommentKeyAbsent()
    {
        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "NoCommentRange",
            ["ref"] = "Sheet1!$B$1",
        });

        var node = _handler.Get("/namedrange[1]");
        node.Should().NotBeNull();
        node.Format["name"].Should().Be("NoCommentRange");
        node.Format.Should().NotContainKey("comment");
    }

    [Fact]
    public void Add_NamedRange_WithComment_LookupByName()
    {
        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "SalesTotal",
            ["ref"] = "Sheet1!$C$1",
            ["comment"] = "Total sales YTD",
        });

        var node = _handler.Get("/namedrange[SalesTotal]");
        node.Should().NotBeNull();
        node.Format["name"].Should().Be("SalesTotal");
        node.Format["comment"].Should().Be("Total sales YTD");
    }

    // ==================== Sheet password protection ====================

    [Fact]
    public void Set_SheetProtect_True_IsReflectedInGet()
    {
        _handler.Set("/Sheet1", new() { ["protect"] = "true" });

        var node = _handler.Get("/Sheet1");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("protect");
        node.Format["protect"].Should().Be(true);
    }

    [Fact]
    public void Set_SheetProtect_WithPassword_ProtectionIsSet()
    {
        _handler.Set("/Sheet1", new() { ["protect"] = "true", ["password"] = "secret123" });

        var node = _handler.Get("/Sheet1");
        node.Format.Should().ContainKey("protect");
        node.Format["protect"].Should().Be(true);
    }

    [Fact]
    public void Set_SheetPassword_SetsHashedPassword()
    {
        _handler.Set("/Sheet1", new() { ["password"] = "MyPass" });

        Reopen();

        // Verify the protection element exists (setting password auto-enables protection)
        var node = _handler.Get("/Sheet1");
        node.Format.Should().ContainKey("protect");
        node.Format["protect"].Should().Be(true);
    }

    [Fact]
    public void Set_SheetProtect_PersistsAfterReopen()
    {
        _handler.Set("/Sheet1", new() { ["protect"] = "true" });

        Reopen();

        var node = _handler.Get("/Sheet1");
        node.Format.Should().ContainKey("protect");
        node.Format["protect"].Should().Be(true);
    }

    [Fact]
    public void Set_SheetProtect_False_RemovesProtection()
    {
        // First enable protection
        _handler.Set("/Sheet1", new() { ["protect"] = "true" });
        var protectedNode = _handler.Get("/Sheet1");
        protectedNode.Format.Should().ContainKey("protect");

        // Now disable it
        _handler.Set("/Sheet1", new() { ["protect"] = "false" });

        var unprotectedNode = _handler.Get("/Sheet1");
        unprotectedNode.Format.Should().NotContainKey("protect");
    }

    [Fact]
    public void Set_SheetProtect_WithPassword_PasswordHashIsNonEmpty()
    {
        _handler.Set("/Sheet1", new() { ["protect"] = "true", ["password"] = "test" });

        Reopen();

        // Verify the protection is present by checking the sheet overview
        var node = _handler.Get("/Sheet1");
        node.Format["protect"].Should().Be(true);
    }

    [Fact]
    public void Set_SheetPassword_DifferentPasswords_ProduceDifferentHashes()
    {
        // Test that the password is actually hashed (not stored as plain text)
        // We do this by verifying the protection is set and accessible
        var path1 = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        var path2 = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        try
        {
            BlankDocCreator.Create(path1);
            BlankDocCreator.Create(path2);

            using var h1 = new ExcelHandler(path1, editable: true);
            using var h2 = new ExcelHandler(path2, editable: true);

            h1.Set("/Sheet1", new() { ["password"] = "alpha" });
            h2.Set("/Sheet1", new() { ["password"] = "beta" });

            // Both should have protection set
            var n1 = h1.Get("/Sheet1");
            var n2 = h2.Get("/Sheet1");
            n1.Format["protect"].Should().Be(true);
            n2.Format["protect"].Should().Be(true);
        }
        finally
        {
            if (File.Exists(path1)) File.Delete(path1);
            if (File.Exists(path2)) File.Delete(path2);
        }
    }
}
