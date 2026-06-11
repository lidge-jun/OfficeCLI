using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Object Finder (Plan 82) ====================

    private static readonly string[] DefaultObjectTypes = ["picture", "field", "bookmark", "equation"];

    /// <summary>List objects of specified type(s) with paths and previews.</summary>
    public string ViewAsObjects(string? objectType = null)
    {
        var types = objectType != null ? [objectType] : DefaultObjectTypes;
        var sb = new StringBuilder();
        int total = 0;

        foreach (var type in types)
        {
            List<System.Xml.Linq.XElement> elements;
            try { elements = ExecuteSelector(type); }
            catch { continue; }

            // formfield: list interactive form fields only
            if (type == "formfield")
                elements = elements.Where(e => e.Attribute("type")?.Value is "CLICK_HERE" or "CHECKBOX" or "DROPDOWN").ToList();

            if (elements.Count == 0) continue;
            total += elements.Count;
            sb.AppendLine($"{type}: {elements.Count}");
            foreach (var el in elements)
            {
                var path = BuildPath(el);
                var preview = GetElementText(el);
                if (preview.Length > 60) preview = preview[..60] + "…";
                if (string.IsNullOrWhiteSpace(preview)) preview = $"({el.Name.LocalName})";

                // Extra info per type
                var extra = type switch
                {
                    "picture" or "img" => el.Attribute("binaryItemIDRef")?.Value is { } r ? $" [{r}]" : "",
                    "field" or "formfield" => el.Attribute("type")?.Value is { } t ? $" [{t}]" : "",
                    "equation" => "",
                    "bookmark" => el.Attribute("name")?.Value is { } n ? $" [{n}]" : "",
                    _ => ""
                };
                sb.AppendLine($"  {path}{extra}: {preview}");
            }
            sb.AppendLine();
        }

        if (total == 0)
            return "(no objects found)";
        sb.Insert(0, $"Objects: {total}\n\n");
        return sb.ToString().TrimEnd();
    }

    /// <summary>JSON output for object finder.</summary>
    public JsonNode ViewAsObjectsJson(string? objectType = null)
    {
        var types = objectType != null ? [objectType] : DefaultObjectTypes;
        var result = new JsonObject();

        foreach (var type in types)
        {
            List<System.Xml.Linq.XElement> elements;
            try { elements = ExecuteSelector(type); }
            catch { continue; }

            if (type == "formfield")
                elements = elements.Where(e => e.Attribute("type")?.Value is "CLICK_HERE" or "CHECKBOX" or "DROPDOWN").ToList();

            if (elements.Count == 0) continue;

            var arr = new JsonArray();
            foreach (var el in elements)
            {
                var obj = new JsonObject
                {
                    ["path"] = BuildPath(el),
                    ["text"] = GetElementText(el)
                };
                // Type-specific attributes
                if (el.Attribute("binaryItemIDRef")?.Value is { } binRef) obj["binaryRef"] = binRef;
                if (el.Attribute("type")?.Value is { } ft) obj["fieldType"] = ft;
                if (el.Attribute("name")?.Value is { } bname) obj["name"] = bname;
                arr.Add(obj);
            }
            result[type] = arr;
        }

        return result;
    }
}
