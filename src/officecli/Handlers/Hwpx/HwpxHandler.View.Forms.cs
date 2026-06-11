using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    private sealed record FormFieldInfo(string Type, string Id, string Name, string Text, string? HelpText, bool IsDefault);
    // ==================== Forms ====================

    public string ViewAsForms(bool auto = true)
    {
        var sb = new StringBuilder();
        var fields = EnumerateInteractiveFormFields().ToList();
        foreach (var field in fields)
        {
            var nameSuffix = string.IsNullOrEmpty(field.Name) ? "" : $" {field.Name}";
            sb.AppendLine($"  [{field.Id}] {field.Type}{nameSuffix}: \"{field.Text}\"{(field.IsDefault ? " (default)" : "")}");
        }
        sb.Insert(0, $"Form fields: {fields.Count}\n");

        if (auto)
        {
            var recognized = RecognizeFormFields();
            if (recognized.Count > 0)
            {
                var adjacentCount = recognized.Count(f => f.Strategy == "adjacent");
                var headerDataCount = recognized.Count(f => f.Strategy == "header-data");
                var strategySummary = new List<string>();
                if (adjacentCount > 0) strategySummary.Add($"{adjacentCount} adjacent");
                if (headerDataCount > 0) strategySummary.Add($"{headerDataCount} header-data");
                var otherCount = recognized.Count - adjacentCount - headerDataCount;
                if (otherCount > 0) strategySummary.Add($"{otherCount} other");

                sb.AppendLine();
                sb.AppendLine($"Forms: {recognized.Count} fields recognized ({string.Join(", ", strategySummary)})");
                sb.AppendLine();

                // Compute column widths
                int labelW = Math.Max(5, recognized.Max(f => f.Label.Length));
                int valueW = Math.Max(5, recognized.Max(f => f.Value.Length));
                int pathW = Math.Max(4, recognized.Max(f => f.Path.Length));
                int stratW = Math.Max(8, recognized.Max(f => f.Strategy.Length));

                // Cap widths to keep output readable
                labelW = Math.Min(labelW, 20);
                valueW = Math.Min(valueW, 24);
                pathW = Math.Min(pathW, 44);

                sb.AppendLine($"  {"Label".PadRight(labelW)}  {"Value".PadRight(valueW)}  {"Path".PadRight(pathW)}  Strategy");
                sb.AppendLine($"  {new string('\u2500', labelW + 2 + valueW + 2 + pathW + 2 + stratW)}");

                foreach (var f in recognized)
                {
                    var label = f.Label.Length > labelW ? f.Label[..(labelW - 1)] + "\u2026" : f.Label.PadRight(labelW);
                    var value = f.Value.Length > valueW ? f.Value[..(valueW - 1)] + "\u2026" : f.Value.PadRight(valueW);
                    var path = f.Path.Length > pathW ? f.Path[..(pathW - 1)] + "\u2026" : f.Path.PadRight(pathW);
                    sb.AppendLine($"  {label}  {value}  {path}  [auto:{f.Strategy}]");
                }
            }

            // F8: Form confidence score
            int totalTables = _doc.Sections.Sum(s => s.Tables.Count);
            if (totalTables > 0)
            {
                var formTablePaths = recognized
                    .Select(f => System.Text.RegularExpressions.Regex.Match(f.Path, @"^/section\[\d+\]/tbl\[\d+\]").Value)
                    .Where(p => !string.IsNullOrEmpty(p))
                    .Distinct()
                    .Count();
                double confidence = (double)formTablePaths / totalTables;
                sb.AppendLine();
                sb.AppendLine($"Form confidence: {confidence:P0} ({formTablePaths}/{totalTables} tables are form-like)");
            }
        }

        return sb.ToString().TrimEnd();
    }

    /// <summary>JSON output for forms view. Supports CLICK_HERE + auto-recognized fields.</summary>
    public JsonNode ViewAsFormsJson(bool auto = true)
    {
        var result = new JsonObject();

        var clickFields = new JsonArray();
        var formFields = new JsonArray();
        foreach (var field in EnumerateInteractiveFormFields())
        {
            if (field.Type == "CLICK_HERE")
            {
                clickFields.Add(new JsonObject {
                    ["id"] = field.Id, ["text"] = field.Text,
                    ["helpText"] = field.HelpText, ["isDefault"] = field.IsDefault
                });
            }

            formFields.Add(new JsonObject {
                ["id"] = field.Id,
                ["type"] = field.Type,
                ["name"] = field.Name,
                ["text"] = field.Text,
                ["helpText"] = field.HelpText,
                ["isDefault"] = field.IsDefault
            });
        }
        result["clickHere"] = clickFields;
        result["formFields"] = formFields;

        if (auto)
        {
            var autoFields = new JsonArray();
            foreach (var f in RecognizeFormFields())
            {
                autoFields.Add(new JsonObject {
                    ["label"] = f.Label, ["value"] = f.Value,
                    ["path"] = f.Path, ["row"] = f.Row, ["col"] = f.Col,
                    ["strategy"] = f.Strategy
                });
            }
            result["autoRecognized"] = autoFields;
        }

        return result;
    }

    private IEnumerable<FormFieldInfo> EnumerateInteractiveFormFields()
    {
        foreach (var sec in _doc.Sections)
        {
            foreach (var run in sec.Root.Descendants(HwpxNs.Hp + "run"))
            {
                var ctrl = run.Element(HwpxNs.Hp + "ctrl");
                var fieldBegin = ctrl?.Element(HwpxNs.Hp + "fieldBegin");
                var fieldType = fieldBegin?.Attribute("type")?.Value;
                if (fieldType is not ("CLICK_HERE" or "CHECKBOX" or "DROPDOWN")) continue;

                var field = fieldBegin!;
                var id = field.Attribute("id")?.Value ?? "?";
                var name = field.Attribute("name")?.Value ?? "";
                var helpText = field.Descendants()
                    .FirstOrDefault(p => p.Attribute("name")?.Value is "Direction" or "Label")
                    ?.Value;
                var nextRun = run.ElementsAfterSelf(HwpxNs.Hp + "run").FirstOrDefault();
                var text = nextRun?.Elements(HwpxNs.Hp + "t").FirstOrDefault()?.Value ?? "";
                var isDefault = !string.IsNullOrEmpty(helpText) && text == helpText;

                yield return new FormFieldInfo(fieldType, id, name, text, helpText, isDefault);
            }
        }
    }
}
