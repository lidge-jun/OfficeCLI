// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug #86: Excel border color readback inconsistency.
/// CellToNode strips ARGB alpha prefix for font.color ("FFFF0000" → "FF0000")
/// but NOT for border.*.color, returning the raw 8-char ARGB value.
/// This means Set(border.left.color=FF0000) → Get reads back "FFFF0000", not "FF0000".
/// </summary>
public class ExcelBorderColorBugTests : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void Set_BorderColor_ShouldReadBackConsistentWithFontColor()
    {
        BlankDocCreator.Create(_path);
        using var handler = new ExcelHandler(_path, editable: true);

        handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Test" });

        // Set both font color and border color
        handler.Set("/Sheet1/A1", new()
        {
            ["font.color"] = "FF0000",
            ["border.all"] = "thin",
            ["border.color"] = "FF0000"
        });

        var node = handler.Get("/Sheet1/A1");

        // font.color should be 6-char (alpha stripped)
        var fontColor = node.Format.GetValueOrDefault("font.color")?.ToString();
        fontColor.Should().Be("#FF0000", "font.color strips the ARGB alpha prefix");

        // border.left.color should ALSO be 6-char for consistency
        var borderColor = node.Format.GetValueOrDefault("border.left.color")?.ToString();
        borderColor.Should().Be("#FF0000",
            "border.left.color should strip the ARGB alpha prefix like font.color does, " +
            "but it returns the raw 8-char ARGB value 'FFFF0000' instead");
    }

    public void Dispose()
    {
        try { File.Delete(_path); } catch { }
    }
}
