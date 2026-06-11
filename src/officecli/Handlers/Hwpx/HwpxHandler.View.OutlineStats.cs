using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    public string ViewAsOutline()
    {
        var sb = new StringBuilder();

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            var styleInfo = GetParagraphStyleInfo(para);
            if (string.IsNullOrEmpty(styleInfo.HeadingLevel)) continue;

            var level = int.Parse(styleInfo.HeadingLevel);
            var indent = new string(' ', (level - 1) * 2);
            var text = HwpxKorean.Normalize(ExtractParagraphText(para));
            var preview = text.Length > 80 ? text[..80] + "…" : text;
            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";

            sb.AppendLine($"{indent}h{level}: {preview} ({path})");
        }

        return sb.Length > 0 ? sb.ToString().TrimEnd() : "(no headings found)";
    }

    public string ViewAsStats()
    {
        int totalParas = 0, totalTables = 0, totalChars = 0, totalWords = 0;
        int totalImages = 0;

        foreach (var sec in _doc.Sections)
        {
            totalParas += sec.Paragraphs.Count;
            totalTables += sec.Tables.Count;
            totalImages += sec.Root.Descendants(HwpxNs.Hp + "img").Count();

            foreach (var p in sec.Paragraphs)
            {
                var text = HwpxKorean.Normalize(ExtractParagraphText(p));
                totalChars += text.Length;
                totalWords += CountWords(text);
            }
        }

        var sb = new StringBuilder();
        sb.AppendLine($"Sections:   {_doc.Sections.Count}");
        sb.AppendLine($"Paragraphs: {totalParas}");
        sb.AppendLine($"Tables:     {totalTables}");
        sb.AppendLine($"Images:     {totalImages}");
        sb.AppendLine($"Characters: {totalChars}");
        sb.AppendLine($"Words:      {totalWords}");

        // Page info — iterate ALL sections for aggregate stats; use first secPr for page size reference
        foreach (var sec in _doc.Sections)
        {
            var secPr = sec.Root.Descendants(HwpxNs.Hp + "secPr").FirstOrDefault();
            var pagePr = secPr?.Element(HwpxNs.Hp + "pagePr");
            if (pagePr != null)
            {
                var width = (int?)pagePr.Attribute("width") ?? 0;
                var height = (int?)pagePr.Attribute("height") ?? 0;
                sb.AppendLine($"Page size:  {FormatHwpUnit(width)} × {FormatHwpUnit(height)}");
                break; // Report first section's page size; add per-section loop if needed
            }
        }

        // Metadata
        var meta = GetMetadata();
        if (meta.TryGetValue("title", out var mTitle) && !string.IsNullOrEmpty(mTitle))
            sb.AppendLine($"Title:      {mTitle}");
        if (meta.TryGetValue("creator", out var mCreator) && !string.IsNullOrEmpty(mCreator))
            sb.AppendLine($"Creator:    {mCreator}");

        return sb.ToString().TrimEnd();
    }

    public JsonNode ViewAsStatsJson()
    {
        int totalParas = 0, totalTables = 0, totalChars = 0, totalWords = 0;
        int totalImages = 0;

        foreach (var sec in _doc.Sections)
        {
            totalParas += sec.Paragraphs.Count;
            totalTables += sec.Tables.Count;
            totalImages += sec.Root.Descendants(HwpxNs.Hp + "img").Count();

            foreach (var p in sec.Paragraphs)
            {
                var text = HwpxKorean.Normalize(ExtractParagraphText(p));
                totalChars += text.Length;
                totalWords += CountWords(text);
            }
        }

        return new JsonObject
        {
            ["sections"] = _doc.Sections.Count,
            ["paragraphs"] = totalParas,
            ["tables"] = totalTables,
            ["images"] = totalImages,
            ["characters"] = totalChars,
            ["words"] = totalWords,
        };
    }

    public JsonNode ViewAsOutlineJson()
    {
        var items = new JsonArray();

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            var styleInfo = GetParagraphStyleInfo(para);
            if (string.IsNullOrEmpty(styleInfo.HeadingLevel)) continue;

            var level = int.Parse(styleInfo.HeadingLevel);
            var text = HwpxKorean.Normalize(ExtractParagraphText(para));
            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";

            items.Add(new JsonObject
            {
                ["level"] = level,
                ["text"] = text,
                ["path"] = path,
            });
        }

        return items;
    }
}
