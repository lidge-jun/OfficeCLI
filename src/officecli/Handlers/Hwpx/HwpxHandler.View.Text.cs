using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    public string ViewAsText(int? startLine = null, int? endLine = null,
                              int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int lineNum = 0;
        int emitted = 0;

        foreach (var (section, para, path) in _doc.AllContentInOrder())
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;

            var rawText = ExtractParagraphText(para);
            var text = HwpxKorean.Normalize(rawText);

            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                sb.AppendLine($"... (more lines)");
                break;
            }

            sb.AppendLine($"{lineNum}. {text}");
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null,
                                   int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int lineNum = 0;
        int emitted = 0;

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;
            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                int remaining = CountRemainingParagraphs(lineNum);
                if (remaining > 0)
                    sb.AppendLine($"... ({remaining} more lines)");
                break;
            }

            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";
            var styleInfo = GetParagraphStyleInfo(para);
            var runs = ExtractAnnotatedRuns(para);
            var text = string.Join("", runs.Select(r => r.Text));
            text = HwpxKorean.Normalize(text);

            // Build annotation prefix
            var annotations = new List<string>();
            if (!string.IsNullOrEmpty(styleInfo.HeadingLevel))
                annotations.Add($"h{styleInfo.HeadingLevel}");
            if (styleInfo.Alignment != "LEFT")
                annotations.Add(styleInfo.Alignment.ToLowerInvariant());

            var prefix = annotations.Count > 0 ? $"[{string.Join(",", annotations)}] " : "";
            sb.AppendLine($"{lineNum}. {path} {prefix}{text}");
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }
    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null,
                                    int? maxLines = null, HashSet<string>? cols = null)
    {
        var lines = new JsonArray();
        int lineNum = 0;
        int emitted = 0;

        foreach (var (section, para, path) in _doc.AllContentInOrder())
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;
            if (maxLines.HasValue && emitted >= maxLines.Value) break;

            var text = HwpxKorean.Normalize(ExtractParagraphText(para));

            lines.Add(new JsonObject
            {
                ["line"] = lineNum,
                ["path"] = path,
                ["text"] = text,
            });
            emitted++;
        }

        return new JsonObject
        {
            ["lines"] = lines,
            ["totalLines"] = lineNum,
        };
    }
    /// <summary>Extract runs with formatting annotations.</summary>
    private static List<(string Text, Dictionary<string, string> Format)> ExtractAnnotatedRuns(XElement para)
    {
        var result = new List<(string, Dictionary<string, string>)>();
        foreach (var run in para.Elements(HwpxNs.Hp + "run"))
        {
            var text = string.Join("", run.Elements(HwpxNs.Hp + "t").Select(t => t.Value));
            if (string.IsNullOrEmpty(text)) continue;

            var format = new Dictionary<string, string>();
            var charPrIdRef = run.Attribute("charPrIDRef")?.Value;
            if (charPrIdRef != null)
                format["charPrIDRef"] = charPrIdRef;

            result.Add((text, format));
        }
        return result;
    }
    private int CountRemainingParagraphs(int currentLine)
    {
        int total = _doc.AllParagraphs().Count();
        return Math.Max(0, total - currentLine);
    }
}
