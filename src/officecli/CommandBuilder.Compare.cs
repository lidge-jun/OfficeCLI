// Plan 84: Document Diff Workflow
using System.CommandLine;
using System.Text.Json.Nodes;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildCompareCommand(Option<bool> jsonOption)
    {
        var fileAArg = new Argument<FileInfo>("fileA") { Description = "First document" };
        var fileBArg = new Argument<FileInfo>("fileB") { Description = "Second document" };
        var modeOpt = new Option<string>("--mode") { Description = "Diff mode: text, outline, table, structural" };
        modeOpt.DefaultValueFactory = _ => "text";

        var cmd = new Command("compare", "Compare two HWPX documents and show differences");
        cmd.Add(fileAArg);
        cmd.Add(fileBArg);
        cmd.Add(modeOpt);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var fileA = result.GetValue(fileAArg)!;
            var fileB = result.GetValue(fileBArg)!;
            var mode = result.GetValue(modeOpt)!;

            using var handlerA = DocumentHandlerFactory.Open(fileA.FullName, editable: false);
            using var handlerB = DocumentHandlerFactory.Open(fileB.FullName, editable: false);

            if (handlerA is not HwpxHandler hwpxA || handlerB is not HwpxHandler hwpxB)
                throw new CliException("Compare is only supported for .hwpx files.")
                    { Code = "unsupported_type" };

            var diff = CompareHwpx(hwpxA, hwpxB, mode);

            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelope(diff.ToJsonString(OutputFormatter.PublicJsonOptions)));
            else
                Console.WriteLine(FormatDiffText(diff, mode));

            return 0;
        }, json); });

        return cmd;
    }

    private static JsonObject CompareHwpx(HwpxHandler a, HwpxHandler b, string mode)
    {
        var result = new JsonObject { ["mode"] = mode };

        switch (mode.ToLowerInvariant())
        {
            case "text":
            {
                var linesA = ExtractLines(a.ViewAsText());
                var linesB = ExtractLines(b.ViewAsText());
                result["diff"] = DiffLines(linesA, linesB);
                break;
            }
            case "outline":
            {
                var linesA = ExtractLines(a.ViewAsOutline());
                var linesB = ExtractLines(b.ViewAsOutline());
                result["diff"] = DiffLines(linesA, linesB);
                break;
            }
            case "table":
            {
                var tablesA = a.ViewAsTables();
                var tablesB = b.ViewAsTables();
                var linesA = ExtractLines(tablesA);
                var linesB = ExtractLines(tablesB);
                result["diff"] = DiffLines(linesA, linesB);
                break;
            }
            case "structural":
            case "struct":
            {
                var changes = new JsonArray();

                // Paragraph-level diff (from text view)
                var textDiff = a.CompareText(b);
                var textChanges = textDiff["changes"]?.AsArray();
                if (textChanges != null)
                {
                    foreach (var c in textChanges)
                    {
                        if (c is JsonObject obj && obj["type"]?.GetValue<string>() != "unchanged")
                        {
                            var entry = new JsonObject { ["level"] = "paragraph" };
                            foreach (var kv in obj) entry[kv.Key] = kv.Value?.DeepClone();
                            changes.Add(entry);
                        }
                    }
                }

                // Table-level diff
                var tabA = a.ViewAsTables();
                var tabB = b.ViewAsTables();
                var tabLinesA = ExtractLines(tabA);
                var tabLinesB = ExtractLines(tabB);
                var tableDiff = DiffLines(tabLinesA, tabLinesB);
                foreach (var item in tableDiff)
                {
                    if (item is JsonObject obj && !obj.ContainsKey("summary") && obj["status"]?.GetValue<string>() is string s && s != "unchanged")
                    {
                        changes.Add(new JsonObject { ["level"] = "table", ["type"] = s, ["text"] = obj["text"]?.DeepClone() });
                    }
                }

                // Outline-level diff (heading structure)
                var outA = ExtractLines(a.ViewAsOutline());
                var outB = ExtractLines(b.ViewAsOutline());
                var outDiff = DiffLines(outA, outB);
                foreach (var item in outDiff)
                {
                    if (item is JsonObject obj && !obj.ContainsKey("summary") && obj["status"]?.GetValue<string>() is string s && s != "unchanged")
                    {
                        changes.Add(new JsonObject { ["level"] = "outline", ["type"] = s, ["text"] = obj["text"]?.DeepClone() });
                    }
                }

                int paraChanges = changes.Count(c => c?["level"]?.GetValue<string>() == "paragraph");
                int tableChanges = changes.Count(c => c?["level"]?.GetValue<string>() == "table");
                int outlineChanges = changes.Count(c => c?["level"]?.GetValue<string>() == "outline");

                result["summary"] = new JsonObject
                {
                    ["paragraphChanges"] = paraChanges,
                    ["tableChanges"] = tableChanges,
                    ["outlineChanges"] = outlineChanges,
                    ["totalChanges"] = changes.Count,
                };
                result["changes"] = changes;
                break;
            }
            default:
                throw new CliException($"Unknown diff mode: {mode}. Available: text, outline, table, structural")
                    { Code = "invalid_value" };
        }

        return result;
    }

    private static string[] ExtractLines(string text)
        => text.Split('\n').Select(l =>
        {
            // Strip line numbers from ViewAsText output ("1. text" → "text")
            var dot = l.IndexOf(". ");
            if (dot > 0 && dot <= 5 && l[..dot].All(char.IsDigit))
                return l[(dot + 2)..];
            return l;
        }).ToArray();

    private static JsonArray DiffLines(string[] linesA, string[] linesB)
    {
        var diff = new JsonArray();
        int maxLen = Math.Max(linesA.Length, linesB.Length);

        // Simple line-by-line diff
        var setA = new HashSet<string>(linesA);
        var setB = new HashSet<string>(linesB);

        foreach (var line in linesA)
        {
            if (!setB.Contains(line) && !string.IsNullOrWhiteSpace(line))
                diff.Add(new JsonObject { ["status"] = "removed", ["text"] = line.Trim() });
        }
        foreach (var line in linesB)
        {
            if (!setA.Contains(line) && !string.IsNullOrWhiteSpace(line))
                diff.Add(new JsonObject { ["status"] = "added", ["text"] = line.Trim() });
        }

        // Summary
        var unchanged = linesA.Intersect(linesB).Count(l => !string.IsNullOrWhiteSpace(l));
        diff.Insert(0, new JsonObject {
            ["summary"] = $"added={diff.Count(d => d?["status"]?.GetValue<string>() == "added")}, removed={diff.Count(d => d?["status"]?.GetValue<string>() == "removed")}, unchanged={unchanged}"
        });

        return diff;
    }

    private static string FormatDiffText(JsonObject diff, string mode)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Diff mode: {mode}");
        sb.AppendLine();

        var diffArr = diff["diff"]?.AsArray();
        if (diffArr == null || diffArr.Count == 0)
        {
            sb.AppendLine("(no differences)");
            return sb.ToString().TrimEnd();
        }

        foreach (var item in diffArr)
        {
            if (item is not JsonObject obj) continue;
            if (obj.ContainsKey("summary"))
            {
                sb.AppendLine(obj["summary"]!.GetValue<string>());
                sb.AppendLine();
                continue;
            }
            var status = obj["status"]?.GetValue<string>() ?? "";
            var text = obj["text"]?.GetValue<string>() ?? "";
            var prefix = status switch { "added" => "+ ", "removed" => "- ", _ => "  " };
            sb.AppendLine($"{prefix}{text}");
        }

        return sb.ToString().TrimEnd();
    }
}
