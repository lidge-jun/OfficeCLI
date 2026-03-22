// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

/// <summary>
/// Shared chart build/read/set logic used by PPTX, Excel, and Word handlers.
/// All methods operate on ChartPart / C.Chart / C.PlotArea — independent of host document type.
/// </summary>
internal static partial class ChartHelper
{
    // ==================== Parse Helpers ====================

    internal static (string kind, bool is3D, bool stacked, bool percentStacked) ParseChartType(string chartType)
    {
        var ct = chartType.ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");
        var is3D = ct.EndsWith("3d") || ct.Contains("3d");
        ct = ct.Replace("3d", "");

        var stacked = ct.Contains("stacked") && !ct.Contains("percent");
        var percentStacked = ct.Contains("percentstacked") || ct.Contains("pstacked");
        ct = ct.Replace("percentstacked", "").Replace("pstacked", "").Replace("stacked", "");

        var kind = ct switch
        {
            "bar" => "bar",
            "column" or "col" => "column",
            "line" => "line",
            "pie" => "pie",
            "doughnut" or "donut" => "doughnut",
            "area" => "area",
            "scatter" or "xy" => "scatter",
            "bubble" => "bubble",
            "radar" or "spider" => "radar",
            "stock" or "ohlc" => "stock",
            "combo" => "combo",
            _ => throw new ArgumentException(
                $"Unknown chart type: '{chartType}'. Supported types: " +
                "column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock, combo. " +
                "Modifiers: 3d (e.g. column3d), stacked (e.g. stackedColumn), percentStacked (e.g. percentStackedBar).")
        };

        return (kind, is3D, stacked, percentStacked);
    }

    internal static List<(string name, double[] values)> ParseSeriesData(Dictionary<string, string> properties)
    {
        var result = new List<(string name, double[] values)>();

        if (properties.TryGetValue("data", out var dataStr))
        {
            foreach (var seriesPart in dataStr.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var colonIdx = seriesPart.IndexOf(':');
                if (colonIdx < 0) continue;
                var name = seriesPart[..colonIdx].Trim();
                var valStr = seriesPart[(colonIdx + 1)..].Trim();
                if (string.IsNullOrEmpty(valStr))
                    throw new ArgumentException($"Series '{name}' has no data values. Expected format: 'Name:1,2,3'");
                var vals = ParseSeriesValues(valStr, name);
                result.Add((name, vals));
            }
            return result;
        }

        for (int i = 1; i <= 20; i++)
        {
            if (!properties.TryGetValue($"series{i}", out var seriesStr)) break;
            var colonIdx = seriesStr.IndexOf(':');
            if (colonIdx < 0)
            {
                var vals = ParseSeriesValues(seriesStr, $"series{i}");
                result.Add(($"Series {i}", vals));
            }
            else
            {
                var name = seriesStr[..colonIdx].Trim();
                var vals = ParseSeriesValues(seriesStr[(colonIdx + 1)..], name);
                result.Add((name, vals));
            }
        }

        return result;
    }

    private static double[] ParseSeriesValues(string valStr, string seriesName)
    {
        return valStr.Split(',').Select(v =>
        {
            var trimmed = v.Trim();
            if (!double.TryParse(trimmed, System.Globalization.CultureInfo.InvariantCulture, out var num))
                throw new ArgumentException($"Invalid data value '{trimmed}' in series '{seriesName}'. Expected comma-separated numbers (e.g. '1,2,3').");
            return num;
        }).ToArray();
    }

    internal static string[]? ParseCategories(Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("categories", out var catStr)) return null;
        return catStr.Split(',').Select(c => c.Trim()).ToArray();
    }

    internal static string[]? ParseSeriesColors(Dictionary<string, string> properties)
    {
        if (properties.TryGetValue("colors", out var colorsStr))
            return colorsStr.Split(',').Select(c => c.Trim()).ToArray();
        return null;
    }
}
