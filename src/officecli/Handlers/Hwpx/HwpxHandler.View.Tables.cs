using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Table Map (Plan 71) ====================

    /// <summary>
    /// Display all tables with grid structure, recognized labels, and cell paths.
    /// </summary>
    public string ViewAsTables()
    {
        var sb = new StringBuilder();
        int tblCount = 0;

        foreach (var (sec, tbl, localTblIdx) in _doc.AllTables())
        {
            tblCount++;
            var (grid, cellList) = BuildTableGrid(tbl);
            if (cellList.Count == 0) continue;

            int maxRow = grid.GetLength(0), maxCol = grid.GetLength(1);
            var basePath = $"/section[{sec.Index + 1}]/tbl[{localTblIdx + 1}]";
            sb.AppendLine($"Table {tblCount} ({basePath}, {maxRow}×{maxCol}):");

            // Grid visualization
            for (int r = 0; r < maxRow; r++)
            {
                sb.Append($"  [{r}] ");
                for (int c = 0; c < maxCol; c++)
                {
                    var cell = grid[r, c];
                    if (cell == null) { sb.Append("·  "); continue; }

                    // Skip duplicate merged cell refs (only show on first occurrence)
                    var (cr, cc, rs, cs) = GetCellAddr(cell);
                    if (cr != r || cc != c) { sb.Append("↕  "); continue; }

                    var text = ExtractCellText(cell).Trim();
                    var preview = text.Length > 12 ? text[..12] + "…" : text;
                    if (string.IsNullOrEmpty(preview)) preview = "(empty)";

                    var span = (rs > 1 || cs > 1) ? $"[{rs}×{cs}]" : "";
                    sb.Append($"{preview}{span}  ");
                }
                sb.AppendLine();
            }

            // Recognized fields for this table
            var fields = new List<RecognizedField>();
            var tableGrid = grid; // reuse
            var seen = new HashSet<XElement>();
            foreach (var (tc, row, col, rowSpan, colSpan) in cellList)
            {
                if (seen.Contains(tc)) continue;
                seen.Add(tc);
                var cellText = ExtractCellText(tc);
                if (!IsLabelCell(cellText)) continue;
                int targetCol = col + colSpan;
                if (targetCol < maxCol)
                {
                    var valueCell = grid[row, targetCol];
                    if (valueCell != null && valueCell != tc)
                    {
                        var value = ExtractCellText(valueCell).Trim();
                        if (!string.IsNullOrEmpty(value))
                            fields.Add(new RecognizedField(
                                NormalizeLabel(cellText), value, basePath, row, col, "adjacent"));
                    }
                }
            }
            if (fields.Count > 0)
            {
                sb.AppendLine($"  Labels: {fields.Count}");
                foreach (var f in fields)
                    sb.AppendLine($"    {f.Label}: {f.Value} (r{f.Row},c{f.Col})");
            }
            sb.AppendLine();
        }

        if (tblCount == 0)
            sb.AppendLine("(no tables)");
        else
            sb.Insert(0, $"Tables: {tblCount}\n\n");

        return sb.ToString().TrimEnd();
    }
    /// <summary>JSON output for table map view.</summary>
    public JsonNode ViewAsTablesJson()
    {
        var result = new JsonObject();
        var tablesArr = new JsonArray();

        foreach (var (sec, tbl, localTblIdx) in _doc.AllTables())
        {
            var (grid, cellList) = BuildTableGrid(tbl);
            if (cellList.Count == 0) continue;

            int maxRow = grid.GetLength(0), maxCol = grid.GetLength(1);
            var basePath = $"/section[{sec.Index + 1}]/tbl[{localTblIdx + 1}]";

            var tblObj = new JsonObject
            {
                ["path"] = basePath,
                ["rows"] = maxRow,
                ["cols"] = maxCol
            };

            // Cells grid
            var cellsArr = new JsonArray();
            for (int r = 0; r < maxRow; r++)
            {
                var rowArr = new JsonArray();
                for (int c = 0; c < maxCol; c++)
                {
                    var cell = grid[r, c];
                    if (cell == null) { rowArr.Add((JsonNode?)null); continue; }
                    var (cr, cc, rs, cs) = GetCellAddr(cell);
                    if (cr != r || cc != c) { rowArr.Add("↕"); continue; }
                    var text = ExtractCellText(cell).Trim();
                    rowArr.Add(new JsonObject
                    {
                        ["text"] = text,
                        ["path"] = $"{basePath}/tr[{r + 1}]/tc[{c + 1}]",
                        ["rowSpan"] = rs,
                        ["colSpan"] = cs
                    });
                }
                cellsArr.Add(rowArr);
            }
            tblObj["cells"] = cellsArr;

            tablesArr.Add(tblObj);
        }

        result["tables"] = tablesArr;
        return result;
    }
}
