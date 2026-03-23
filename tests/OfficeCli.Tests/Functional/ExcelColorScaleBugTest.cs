// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug: ColorScale conditional formatting silently ignores user-specified
/// minColor/maxColor values when they are passed with mixed case (e.g. "minColor"
/// instead of "mincolor").
///
/// Root cause: ExcelHandler.Add.cs (case "colorscale") reads properties via
/// GetValueOrDefault("mincolor", ...) — lowercase keys only. The CLI CommandBuilder
/// builds the properties Dictionary with case-sensitive default StringComparer,
/// preserving the original casing the user typed ("minColor", "maxColor").
/// The GetValueOrDefault lookup therefore fails to match, falls through to the
/// default values ("F8696B"/"63BE7B"), and the custom colors are silently dropped.
///
/// This test should FAIL with current code: the colors read back will be the
/// defaults (#F8696B / #63BE7B) rather than the user-supplied values.
/// </summary>
public class ExcelColorScaleBugTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");

    public void Dispose()
    {
        try { File.Delete(_path); } catch { }
    }

    /// <summary>
    /// Simulates the CLI user typing --prop minColor=FFCCCC --prop maxColor=00CC00
    /// (mixed case, as documented in help examples). The colorscale handler must
    /// honour these values regardless of key casing.
    /// </summary>
    [Fact]
    public void AddColorScale_WithMixedCaseColorKeys_CustomColorsAreApplied()
    {
        BlankDocCreator.Create(_path);
        using var handler = new ExcelHandler(_path, editable: true);

        // User typed minColor / maxColor — mixed case, as they would from CLI --prop flags.
        // The CLI does NOT normalise key casing before passing to Add().
        var props = new Dictionary<string, string>
        {
            ["sqref"]    = "C2:D4",
            ["type"]     = "colorscale",
            ["minColor"] = "FFCCCC",   // mixed case: uppercase C — should be accepted
            ["maxColor"] = "00CC00",   // mixed case: uppercase C — should be accepted
        };

        var resultPath = handler.Add("/Sheet1", "cf", null, props);

        resultPath.Should().StartWith("/Sheet1/cf[");

        // Get the CF node back and verify that the custom colors were stored
        var node = handler.Get(resultPath);
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("mincolor");
        node.Format.Should().ContainKey("maxcolor");

        // These assertions FAIL with current code because the handler falls back to
        // the hardcoded defaults instead of reading "FFCCCC" / "00CC00"
        node.Format["mincolor"].ToString().Should().Be("#FFCCCC",
            "minColor=FFCCCC was specified; colorscale handler must accept mixed-case key");
        node.Format["maxcolor"].ToString().Should().Be("#00CC00",
            "maxColor=00CC00 was specified; colorscale handler must accept mixed-case key");
    }

    /// <summary>
    /// Lowercase keys should also work correctly (baseline sanity check).
    /// This test should PASS with current code.
    /// </summary>
    [Fact]
    public void AddColorScale_WithLowercaseColorKeys_CustomColorsAreApplied()
    {
        BlankDocCreator.Create(_path);
        using var handler = new ExcelHandler(_path, editable: true);

        var props = new Dictionary<string, string>
        {
            ["sqref"]    = "A1:A10",
            ["type"]     = "colorscale",
            ["mincolor"] = "FFCCCC",   // all lowercase — current code handles this
            ["maxcolor"] = "00CC00",
        };

        var resultPath = handler.Add("/Sheet1", "cf", null, props);
        var node = handler.Get(resultPath);

        node.Format["mincolor"].ToString().Should().Be("#FFCCCC");
        node.Format["maxcolor"].ToString().Should().Be("#00CC00");
    }

    /// <summary>
    /// Verifies that custom colors survive a save/reopen cycle (persistence test).
    /// Depends on the mixed-case bug being fixed first; written to catch regressions.
    /// </summary>
    [Fact]
    public void AddColorScale_WithMixedCaseColorKeys_PersistsAfterReopen()
    {
        BlankDocCreator.Create(_path);

        {
            using var handler = new ExcelHandler(_path, editable: true);
            handler.Add("/Sheet1", "cf", null, new Dictionary<string, string>
            {
                ["sqref"]    = "B2:B8",
                ["type"]     = "colorscale",
                ["minColor"] = "FF0000",
                ["maxColor"] = "0000FF",
            });
        }

        // Reopen
        using var reopened = new ExcelHandler(_path, editable: false);
        var node = reopened.Get("/Sheet1/cf[1]");

        node.Format["mincolor"].ToString().Should().Be("#FF0000",
            "minColor should persist after save and reopen");
        node.Format["maxcolor"].ToString().Should().Be("#0000FF",
            "maxColor should persist after save and reopen");
    }
}
