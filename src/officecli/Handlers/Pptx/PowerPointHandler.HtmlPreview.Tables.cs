// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Table Rendering ====================

    private static void RenderTable(StringBuilder sb, GraphicFrame gf, Dictionary<string, string> themeColors)
    {
        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
        if (table == null) return;

        var offset = gf.Transform?.Offset;
        var extents = gf.Transform?.Extents;
        if (offset == null || extents == null) return;

        var x = offset.X?.Value ?? 0;
        var y = offset.Y?.Value ?? 0;
        var cx = extents.Cx?.Value ?? 0;
        var cy = extents.Cy?.Value ?? 0;

        // Detect table style for style-based coloring
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        var tableStyleId = tblPr?.GetFirstChild<Drawing.TableStyleId>()?.InnerText;
        var tableStyleName = tableStyleId != null && _tableStyleGuidToName.TryGetValue(tableStyleId, out var sn) ? sn : null;
        bool hasFirstRow = tblPr?.FirstRow?.Value == true;
        bool hasBandRow = tblPr?.BandRow?.Value == true;

        sb.AppendLine($"    <div class=\"table-container\" style=\"left:{EmuToCm(x)}cm;top:{EmuToCm(y)}cm;width:{EmuToCm(cx)}cm;height:{EmuToCm(cy)}cm\">");
        sb.AppendLine("      <table class=\"slide-table\">");

        // Column widths
        var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();
        if (gridCols != null && gridCols.Count > 0)
        {
            sb.Append("        <colgroup>");
            long totalWidth = gridCols.Sum(gc => gc.Width?.Value ?? 0);
            foreach (var gc in gridCols)
            {
                var w = gc.Width?.Value ?? 0;
                var pct = totalWidth > 0 ? (w * 100.0 / totalWidth) : (100.0 / gridCols.Count);
                sb.Append($"<col style=\"width:{pct:0.##}%\">");
            }
            sb.AppendLine("</colgroup>");
        }

        int rowIndex = 0;
        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            sb.AppendLine("        <tr>");
            int skipCols = 0;
            bool isHeaderRow = hasFirstRow && rowIndex == 0;
            bool isBandedOdd = hasBandRow && (!hasFirstRow ? rowIndex % 2 == 0 : rowIndex > 0 && (rowIndex - 1) % 2 == 0);

            foreach (var cell in row.Elements<Drawing.TableCell>())
            {
                var cellStyles = new List<string>();

                // Cell fill
                var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                var cellSolid = tcPr?.GetFirstChild<Drawing.SolidFill>();
                var cellColor = ResolveFillColor(cellSolid, themeColors);
                bool hasExplicitFill = cellColor != null;
                if (cellColor != null)
                    cellStyles.Add($"background:{cellColor}");

                var cellGrad = tcPr?.GetFirstChild<Drawing.GradientFill>();
                if (cellGrad != null)
                {
                    cellStyles.Add($"background:{GradientToCss(cellGrad, themeColors)}");
                    hasExplicitFill = true;
                }

                // Apply table-style-based colors when no explicit cell fill
                if (!hasExplicitFill && tableStyleName != null)
                {
                    var (bg, fg) = GetTableStyleColors(tableStyleName, isHeaderRow, isBandedOdd);
                    if (bg != null) cellStyles.Add($"background:{bg}");
                    if (fg != null) cellStyles.Add($"color:{fg}");
                }

                // Vertical alignment
                if (tcPr?.Anchor?.HasValue == true)
                {
                    var va = tcPr.Anchor.InnerText switch
                    {
                        "ctr" => "middle",
                        "b" => "bottom",
                        _ => "top"
                    };
                    cellStyles.Add($"vertical-align:{va}");
                }

                // Cell text formatting
                var firstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
                if (firstRun?.RunProperties != null)
                {
                    var rp = firstRun.RunProperties;
                    if (rp.FontSize?.HasValue == true)
                        cellStyles.Add($"font-size:{rp.FontSize.Value / 100.0:0.##}pt");
                    if (rp.Bold?.Value == true)
                        cellStyles.Add("font-weight:bold");
                    var fontVal = rp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                        ?? rp.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                    if (fontVal != null && !fontVal.StartsWith("+", StringComparison.Ordinal))
                        cellStyles.Add(CssFontFamilyWithFallback(fontVal));
                    var runColor = ResolveFillColor(rp.GetFirstChild<Drawing.SolidFill>(), themeColors);
                    if (runColor != null)
                        cellStyles.Add($"color:{runColor}");
                }

                // Cell borders (per-edge)
                if (tcPr != null)
                {
                    var borderLeft = tcPr.GetFirstChild<Drawing.LeftBorderLineProperties>();
                    var borderRight = tcPr.GetFirstChild<Drawing.RightBorderLineProperties>();
                    var borderTop = tcPr.GetFirstChild<Drawing.TopBorderLineProperties>();
                    var borderBottom = tcPr.GetFirstChild<Drawing.BottomBorderLineProperties>();
                    var bl = TableBorderToCss(borderLeft, themeColors);
                    var br = TableBorderToCss(borderRight, themeColors);
                    var bt = TableBorderToCss(borderTop, themeColors);
                    var bb = TableBorderToCss(borderBottom, themeColors);
                    if (bl != null) cellStyles.Add($"border-left:{bl}");
                    if (br != null) cellStyles.Add($"border-right:{br}");
                    if (bt != null) cellStyles.Add($"border-top:{bt}");
                    if (bb != null) cellStyles.Add($"border-bottom:{bb}");
                    // If no explicit borders at all, render a thin default border
                    if (bl == null && br == null && bt == null && bb == null)
                        cellStyles.Add("border:1px solid rgba(0,0,0,0.2)");
                }

                // Cell margins/padding
                var marL = tcPr?.LeftMargin?.Value;
                var marR = tcPr?.RightMargin?.Value;
                var marT = tcPr?.TopMargin?.Value;
                var marB = tcPr?.BottomMargin?.Value;
                if (marL.HasValue || marR.HasValue || marT.HasValue || marB.HasValue)
                {
                    var pT = EmuToCm(marT ?? 45720);
                    var pR = EmuToCm(marR ?? 91440);
                    var pB = EmuToCm(marB ?? 45720);
                    var pL = EmuToCm(marL ?? 91440);
                    cellStyles.Add($"padding:{pT}cm {pR}cm {pB}cm {pL}cm");
                }

                // Paragraph alignment
                var firstPara = cell.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
                if (firstPara?.ParagraphProperties?.Alignment?.HasValue == true)
                {
                    var align = firstPara.ParagraphProperties.Alignment.InnerText switch
                    {
                        "ctr" => "center",
                        "r" => "right",
                        "just" => "justify",
                        _ => "left"
                    };
                    cellStyles.Add($"text-align:{align}");
                }

                var cellText = cell.TextBody?.InnerText ?? "";
                var styleStr = cellStyles.Count > 0 ? $" style=\"{string.Join(";", cellStyles)}\"" : "";

                // Column/row span (GridSpan and RowSpan are on the TableCell, not TableCellProperties)
                var gridSpan = cell.GridSpan?.Value;
                var rowSpan = cell.RowSpan?.Value;
                var spanAttrs = "";
                if (gridSpan > 1) spanAttrs += $" colspan=\"{gridSpan}\"";
                if (rowSpan > 1) spanAttrs += $" rowspan=\"{rowSpan}\"";

                // Skip merged continuation cells
                if (cell.HorizontalMerge?.Value == true || cell.VerticalMerge?.Value == true)
                    continue;

                // Skip cells covered by previous gridSpan
                if (skipCols > 0)
                {
                    skipCols--;
                    continue;
                }

                if (gridSpan > 1) skipCols = (int)gridSpan - 1;

                sb.AppendLine($"          <td{spanAttrs}{styleStr}>{HtmlEncode(cellText)}</td>");
            }
            sb.AppendLine("        </tr>");
            rowIndex++;
        }

        sb.AppendLine("      </table>");
        sb.AppendLine("    </div>");
    }

    /// <summary>
    /// Convert a table cell border line properties element to a CSS border value.
    /// Returns null if the border has NoFill or is absent.
    /// </summary>
    private static string? TableBorderToCss(OpenXmlCompositeElement? borderProps, Dictionary<string, string> themeColors)
    {
        if (borderProps == null) return null;
        if (borderProps.GetFirstChild<Drawing.NoFill>() != null) return "none";

        var solidFill = borderProps.GetFirstChild<Drawing.SolidFill>();
        var color = ResolveFillColor(solidFill, themeColors) ?? "#000000";

        // Width attribute is on the element itself (w attr in EMU)
        double widthPt = 1.0;
        if (borderProps is Drawing.LeftBorderLineProperties lb && lb.Width?.HasValue == true)
            widthPt = lb.Width.Value / 12700.0;
        else if (borderProps is Drawing.RightBorderLineProperties rb && rb.Width?.HasValue == true)
            widthPt = rb.Width.Value / 12700.0;
        else if (borderProps is Drawing.TopBorderLineProperties tb && tb.Width?.HasValue == true)
            widthPt = tb.Width.Value / 12700.0;
        else if (borderProps is Drawing.BottomBorderLineProperties bb && bb.Width?.HasValue == true)
            widthPt = bb.Width.Value / 12700.0;

        if (widthPt < 0.5) widthPt = 0.5;

        var dash = borderProps.GetFirstChild<Drawing.PresetDash>();
        var style = "solid";
        if (dash?.Val?.HasValue == true)
        {
            style = dash.Val.InnerText switch
            {
                "dash" or "lgDash" or "sysDash" => "dashed",
                "dot" or "sysDot" => "dotted",
                _ => "solid"
            };
        }

        return $"{widthPt:0.##}pt {style} {color}";
    }

    /// <summary>
    /// Returns (background, foreground) CSS colors for a table style based on row position.
    /// </summary>
    private static (string? bg, string? fg) GetTableStyleColors(string styleName, bool isHeader, bool isBandedOdd)
    {
        return styleName switch
        {
            "medium1" => isHeader ? ("#4472C4", "#FFFFFF")
                       : isBandedOdd ? ("#D6E4F0", null)
                       : (null, null),
            "medium2" => isHeader ? ("#404040", "#FFFFFF")
                       : isBandedOdd ? ("#E0E0E0", null)
                       : ("#F5F5F5", null),
            "medium3" => isHeader ? ("#4472C4", "#FFFFFF")
                       : isBandedOdd ? ("#B4C6E7", null)
                       : (null, null),
            "medium4" => isHeader ? ("#2F5496", "#FFFFFF")
                       : isBandedOdd ? ("#B4C6E7", null)
                       : ("#D6E4F0", null),
            "dark1" => isHeader ? ("#000000", "#FFFFFF")
                     : isBandedOdd ? ("#808080", "#FFFFFF")
                     : ("#595959", "#FFFFFF"),
            "dark2" => isHeader ? ("#4472C4", "#FFFFFF")
                     : isBandedOdd ? ("#2F5496", "#FFFFFF")
                     : ("#2A4A87", "#FFFFFF"),
            "light1" => (null, null), // minimal styling, no fills
            "light2" => isBandedOdd ? ("#D6E4F0", null) : (null, null),
            "light3" => isBandedOdd ? ("#D6E4F0", null) : (null, null),
            _ => (null, null),
        };
    }
}
