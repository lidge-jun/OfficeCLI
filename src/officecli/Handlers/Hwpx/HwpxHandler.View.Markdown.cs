using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Markdown Export (Plan 72) ====================

    /// <summary>Export document as GitHub Flavored Markdown.</summary>
    public string ViewAsMarkdown()
    {
        var sb = new StringBuilder();

        foreach (var (section, element, path) in _doc.AllContentInOrder())
        {
            var localName = element.Name.LocalName;
            if (localName == "p")
            {
                // Check if this paragraph is inside a table cell (skip — handled by table renderer)
                if (element.Ancestors().Any(a => a.Name.LocalName == "tc")) continue;

                var styleInfo = GetParagraphStyleInfo(element);
                var mdLine = ParagraphToMarkdown(element);
                if (string.IsNullOrWhiteSpace(mdLine)) { sb.AppendLine(); continue; }

                if (!string.IsNullOrEmpty(styleInfo.HeadingLevel))
                {
                    var level = Math.Clamp(int.Parse(styleInfo.HeadingLevel), 1, 6);
                    sb.AppendLine($"{new string('#', level)} {mdLine}");
                }
                else
                {
                    sb.AppendLine(mdLine);
                }
                sb.AppendLine();
            }
        }

        // Render tables
        foreach (var (sec, tbl, localTblIdx) in _doc.AllTables())
        {
            var (grid, cellList) = BuildTableGrid(tbl);
            if (cellList.Count == 0) continue;
            int maxRow = grid.GetLength(0), maxCol = grid.GetLength(1);

            // F5: Single-cell tables → emit as structured text instead of table
            if (maxRow == 1 && maxCol == 1 && cellList.Count == 1)
            {
                var cellText = ExtractCellText(cellList[0].Tc).Trim();
                if (!string.IsNullOrEmpty(cellText))
                {
                    var lines = cellText.Split('\n');
                    foreach (var line in lines)
                    {
                        var trimmed = line.Trim();
                        if (string.IsNullOrEmpty(trimmed)) { sb.AppendLine(); continue; }
                        var m = System.Text.RegularExpressions.Regex.Match(trimmed, @"^(\d+[.)]|[가-하][.]|[a-z][.)]) (.+)$");
                        if (m.Success)
                            sb.AppendLine($"**{m.Groups[1].Value}** {m.Groups[2].Value}");
                        else
                            sb.AppendLine(trimmed);
                    }
                    sb.AppendLine();
                }
                continue;
            }

            // F6: Pseudo-table demotion — skip tables with <=3 rows and >=30% empty cells
            if (maxRow <= 3)
            {
                int totalCells = maxRow * maxCol;
                int emptyCells = 0;
                for (int r = 0; r < maxRow; r++)
                    for (int c = 0; c < maxCol; c++)
                    {
                        var cell = grid[r, c];
                        if (cell == null || string.IsNullOrWhiteSpace(ExtractCellText(cell)))
                            emptyCells++;
                    }
                if (totalCells > 0 && (double)emptyCells / totalCells >= 0.3)
                {
                    for (int r = 0; r < maxRow; r++)
                        for (int c = 0; c < maxCol; c++)
                        {
                            var cell = grid[r, c];
                            if (cell == null) continue;
                            var (cr, cc, _, _) = GetCellAddr(cell);
                            if (cr != r || cc != c) continue;
                            var text = ExtractCellText(cell).Trim();
                            if (!string.IsNullOrEmpty(text))
                                sb.AppendLine(text);
                        }
                    sb.AppendLine();
                    continue;
                }
            }

            for (int r = 0; r < maxRow; r++)
            {
                sb.Append("| ");
                for (int c = 0; c < maxCol; c++)
                {
                    var cell = grid[r, c];
                    if (cell == null) { sb.Append("| "); continue; }
                    var (cr, cc, _, _) = GetCellAddr(cell);
                    if (cr != r || cc != c) { sb.Append("| "); continue; } // merged continuation
                    var text = ExtractCellText(cell).Trim().Replace("\n", " ").Replace("|", "\\|");
                    sb.Append($"{text} | ");
                }
                sb.AppendLine();

                // Separator after header row
                if (r == 0)
                {
                    sb.Append("| ");
                    for (int c = 0; c < maxCol; c++)
                        sb.Append("--- | ");
                    sb.AppendLine();
                }
            }
            sb.AppendLine();
        }

        return sb.ToString().Trim();
    }

    private string ParagraphToMarkdown(XElement p)
    {
        var sb = new StringBuilder();
        foreach (var run in p.Elements(HwpxNs.Hp + "run"))
            sb.Append(RunToMarkdown(run));
        return sb.ToString().Trim();
    }

    private string RunToMarkdown(XElement run)
    {
        var sb = new StringBuilder();
        var charPrId = run.Attribute("charPrIDRef")?.Value ?? "0";
        var charPr = FindCharPr(charPrId);
        var hasBold = charPr?.Element(HwpxNs.Hh + "bold") != null;
        var hasItalic = charPr?.Element(HwpxNs.Hh + "italic") != null;
        var soEl = charPr?.Element(HwpxNs.Hh + "strikeout");
        var hasStrikeout = soEl != null && soEl.Attribute("shape")?.Value != "NONE";

        var textParts = new StringBuilder();
        foreach (var child in run.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "t":
                    textParts.Append(child.Value);
                    break;
                case "lineBreak":
                    textParts.Append("  \n"); // MD hard line break
                    break;
                case "tab":
                    textParts.Append('\t');
                    break;
                case "equation":
                    var script = child.Element(HwpxNs.Hp + "script")?.Value
                        ?? child.Attribute("script")?.Value ?? child.Value;
                    textParts.Append($"`{script.Trim()}`");
                    break;
                case "img": case "picture":
                    var src = child.Attribute("binaryItemIDRef")?.Value ?? "image";
                    textParts.Append($"![{src}]({src})");
                    break;
            }
        }

        var text = textParts.ToString();
        if (string.IsNullOrEmpty(text)) return "";

        // F4: GFM tilde escape — prevent false strikethrough from literal tildes
        // Must happen BEFORE strikethrough wrapping
        if (!hasStrikeout)
            text = text.Replace("~", @"\~");

        if (hasStrikeout) text = $"~~{text}~~";
        if (hasBold && hasItalic) text = $"***{text}***";
        else if (hasBold) text = $"**{text}**";
        else if (hasItalic) text = $"*{text}*";

        sb.Append(text);
        return sb.ToString();
    }
}
