using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== HTML Preview ====================

    public string ViewAsHtml(int? page = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html><html lang=\"ko\"><head><meta charset=\"UTF-8\">");
        sb.AppendLine("<title>HWPX Preview</title>");
        sb.AppendLine("<style>");
        sb.AppendLine(HwpxHtmlCss());
        sb.AppendLine("</style></head><body><div class=\"page\">");

        foreach (var (section, element, path) in _doc.AllContentInOrder())
        {
            switch (element.Name.LocalName)
            {
                case "p":
                    var wrappedTbl = element.Descendants(HwpxNs.Hp + "tbl").FirstOrDefault();
                    if (wrappedTbl != null)
                        sb.Append(TableToHtml(wrappedTbl));
                    else
                        sb.Append(ParagraphToHtml(element));
                    break;
            }
        }

        sb.AppendLine("</div></body></html>");
        return sb.ToString();
    }

    // ==================== HTML Helpers ====================

    private string ParagraphToHtml(XElement p)
    {
        var styleInfo = GetParagraphStyleInfo(p);
        var tag = "p";

        if (!string.IsNullOrEmpty(styleInfo.HeadingLevel))
        {
            var level = Math.Clamp(int.Parse(styleInfo.HeadingLevel), 1, 6);
            tag = $"h{level}";
        }

        var paraCss = GetParaPrCss(p.Attribute("paraPrIDRef")?.Value ?? "0");

        var sb = new StringBuilder();
        sb.Append($"<{tag}");
        if (!string.IsNullOrEmpty(paraCss)) sb.Append($" style=\"{paraCss}\"");
        sb.Append('>');

        foreach (var run in p.Elements(HwpxNs.Hp + "run"))
            sb.Append(RunToHtml(run));

        sb.Append($"</{tag}>");
        return sb.ToString();
    }

    private string RunToHtml(XElement run)
    {
        var sb = new StringBuilder();
        var charPrId = run.Attribute("charPrIDRef")?.Value ?? "0";
        var css = GetCharPrCss(charPrId);
        var charPr = FindCharPr(charPrId);
        var hasBold = charPr?.Element(HwpxNs.Hh + "bold") != null;
        var hasItalic = charPr?.Element(HwpxNs.Hh + "italic") != null;
        var ulEl = charPr?.Element(HwpxNs.Hh + "underline");
        var hasUnderline = ulEl != null && ulEl.Attribute("type")?.Value != "NONE";
        var soEl = charPr?.Element(HwpxNs.Hh + "strikeout");
        var hasStrikeout = soEl != null && soEl.Attribute("shape")?.Value != "NONE";
        var hasSup = charPr?.Element(HwpxNs.Hh + "supscript") != null;
        var hasSub = charPr?.Element(HwpxNs.Hh + "subscript") != null;

        if (!string.IsNullOrEmpty(css)) sb.Append($"<span style=\"{css}\">");
        if (hasBold) sb.Append("<b>");
        if (hasItalic) sb.Append("<i>");
        if (hasUnderline) sb.Append("<u>");
        if (hasStrikeout) sb.Append("<s>");
        if (hasSup) sb.Append("<sup>");
        if (hasSub) sb.Append("<sub>");

        foreach (var child in run.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "t":
                    sb.Append(TextWithMarkpenToHtml(child));
                    break;
                case "lineBreak":
                    sb.Append("<br/>");
                    break;
                case "tab":
                    sb.Append("&emsp;");
                    break;
                case "equation":
                    var script = child.Element(HwpxNs.Hp + "script")?.Value
                        ?? child.Attribute("script")?.Value ?? child.Value;
                    sb.Append($"<span class=\"hwpx-eq\" title=\"{EscapeHtml(script)}\">[{EscapeHtml(script.Trim())}]</span>");
                    break;
                case "pic":
                    sb.Append(PicToHtml(child));
                    break;
            }
        }

        if (hasSub) sb.Append("</sub>");
        if (hasSup) sb.Append("</sup>");
        if (hasStrikeout) sb.Append("</s>");
        if (hasUnderline) sb.Append("</u>");
        if (hasItalic) sb.Append("</i>");
        if (hasBold) sb.Append("</b>");
        if (!string.IsNullOrEmpty(css)) sb.Append("</span>");

        return sb.ToString();
    }

    private static string TextWithMarkpenToHtml(XElement t)
    {
        var sb = new StringBuilder();
        foreach (var node in t.Nodes())
        {
            if (node is System.Xml.Linq.XText text)
                sb.Append(EscapeHtml(text.Value));
            else if (node is XElement el)
            {
                if (el.Name.LocalName == "markpenBegin")
                {
                    var color = el.Attribute("color")?.Value ?? "#FFFF00";
                    sb.Append($"<mark style=\"background:{color}\">");
                }
                else if (el.Name.LocalName == "markpenEnd")
                    sb.Append("</mark>");
            }
        }
        return sb.ToString();
    }

    private string TableToHtml(XElement tbl)
    {
        var sb = new StringBuilder();
        sb.Append("<table>");
        foreach (var tr in tbl.Elements(HwpxNs.Hp + "tr"))
        {
            sb.Append("<tr>");
            foreach (var tc in tr.Elements(HwpxNs.Hp + "tc"))
            {
                var cellSpan = tc.Element(HwpxNs.Hp + "cellSpan");
                var colspan = (int?)cellSpan?.Attribute("colSpan") ?? 1;
                var rowspan = (int?)cellSpan?.Attribute("rowSpan") ?? 1;
                var subList = tc.Element(HwpxNs.Hp + "subList");
                var vAlign = subList?.Attribute("vertAlign")?.Value?.ToLowerInvariant() ?? "top";

                var bfId = tc.Attribute("borderFillIDRef")?.Value;
                var cellCss = $"vertical-align:{vAlign}";
                if (bfId != null)
                {
                    var bgColor = GetBorderFillBgColor(bfId);
                    if (bgColor != null) cellCss += $";background:{bgColor}";
                }

                sb.Append("<td");
                if (colspan > 1) sb.Append($" colspan=\"{colspan}\"");
                if (rowspan > 1) sb.Append($" rowspan=\"{rowspan}\"");
                sb.Append($" style=\"{cellCss}\">");

                if (subList != null)
                {
                    foreach (var cp in subList.Elements(HwpxNs.Hp + "p"))
                        sb.Append(ParagraphToHtml(cp));
                }
                sb.Append("</td>");
            }
            sb.Append("</tr>");
        }
        sb.Append("</table>");
        return sb.ToString();
    }

    private string PicToHtml(XElement pic)
    {
        var imgEl = pic.Descendants().FirstOrDefault(e => e.Name.LocalName == "img");
        var src = imgEl?.Attribute("src")?.Value ?? imgEl?.Attribute("binaryItemIDRef")?.Value;
        if (src != null)
        {
            var binData = _doc.GetBinData(src);
            if (binData != null)
            {
                var ext = Path.GetExtension(src).ToLowerInvariant();
                var mime = ext switch { ".png" => "image/png", ".gif" => "image/gif", ".bmp" => "image/bmp", _ => "image/jpeg" };
                return $"<img src=\"data:{mime};base64,{Convert.ToBase64String(binData)}\" style=\"max-width:100%\"/>";
            }
        }
        return "<span class=\"hwpx-img\">[image]</span>";
    }

    private string GetCharPrCss(string charPrId)
    {
        var charPr = FindCharPr(charPrId);
        if (charPr == null) return "";
        var parts = new List<string>();
        var height = (int?)charPr.Attribute("height") ?? 1000;
        parts.Add($"font-size:{height / 100.0:0.#}pt");
        var color = charPr.Attribute("textColor")?.Value;
        if (color != null && color != "#000000") parts.Add($"color:{color}");
        var fontRef = charPr.Element(HwpxNs.Hh + "fontRef");
        if (fontRef != null)
        {
            var hangulRef = fontRef.Attribute("hangul")?.Value ?? "0";
            var fontName = GetFontName("HANGUL", hangulRef);
            if (fontName != null) parts.Add($"font-family:'{fontName}',sans-serif");
        }
        return string.Join(";", parts);
    }

    private string GetParaPrCss(string paraPrId)
    {
        if (_doc.Header?.Root == null) return "";
        var paraPr = _doc.Header.Root.Descendants(HwpxNs.Hh + "paraPr")
            .FirstOrDefault(p => p.Attribute("id")?.Value == paraPrId);
        if (paraPr == null) return "";
        var parts = new List<string>();
        var align = paraPr.Element(HwpxNs.Hh + "align")?.Attribute("horizontal")?.Value;
        if (align != null && align != "JUSTIFY")
            parts.Add($"text-align:{align.ToLowerInvariant()}");
        else if (align == "JUSTIFY")
            parts.Add("text-align:justify");
        var margin = paraPr.Element(HwpxNs.Hh + "margin");
        if (margin != null)
        {
            var indent = (int?)margin.Attribute("indent") ?? 0;
            if (indent != 0) parts.Add($"text-indent:{indent / 283.46:0.#}mm");
            var left = (int?)margin.Attribute("left") ?? 0;
            if (left != 0) parts.Add($"margin-left:{left / 283.46:0.#}mm");
        }
        var ls = paraPr.Element(HwpxNs.Hh + "lineSpacing");
        if (ls != null)
        {
            var lsType = ls.Attribute("type")?.Value;
            var lsVal = (int?)ls.Attribute("value") ?? 160;
            if (lsType == "PERCENT") parts.Add($"line-height:{lsVal / 100.0:0.##}");
        }
        return string.Join(";", parts);
    }

    private string? GetBorderFillBgColor(string bfId)
    {
        var bf = _doc.Header?.Root?.Descendants(HwpxNs.Hh + "borderFill")
            .FirstOrDefault(b => b.Attribute("id")?.Value == bfId);
        var winBrush = bf?.Descendants(HwpxNs.Hc + "winBrush").FirstOrDefault();
        return winBrush?.Attribute("faceColor")?.Value;
    }

    private string? GetFontName(string lang, string fontRef)
    {
        var fontface = _doc.Header?.Root?.Descendants(HwpxNs.Hh + "fontface")
            .FirstOrDefault(f => f.Attribute("lang")?.Value == lang);
        var font = fontface?.Elements(HwpxNs.Hh + "font")
            .FirstOrDefault(f => f.Attribute("id")?.Value == fontRef);
        return font?.Attribute("face")?.Value;
    }

    private static string EscapeHtml(string text)
        => text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    private static string HwpxHtmlCss() => """
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { background: #e8e8e8; font-family: '함초롬돋움', 'Malgun Gothic', sans-serif; }
        .page { max-width: 210mm; margin: 20px auto; padding: 20mm 25mm; background: #fff;
                box-shadow: 0 2px 8px rgba(0,0,0,0.15); min-height: 297mm; }
        p { margin: 2px 0; font-size: 10pt; line-height: 1.6; }
        h1 { font-size: 16pt; margin: 12px 0 4px; }
        h2 { font-size: 14pt; margin: 10px 0 4px; }
        h3 { font-size: 12pt; margin: 8px 0 4px; }
        h4, h5, h6 { font-size: 11pt; margin: 6px 0 4px; }
        table { border-collapse: collapse; width: 100%; margin: 8px 0; }
        td, th { border: 1px solid #000; padding: 4px 8px; font-size: 10pt; }
        .hwpx-eq { font-family: 'HancomEQN', serif; color: #333; background: #f5f5f5;
                   padding: 2px 6px; border-radius: 3px; font-size: 0.9em; }
        .hwpx-img { color: #999; font-style: italic; }
        mark { padding: 1px 2px; }
        @media print { body { background: #fff; } .page { box-shadow: none; margin: 0; padding: 20mm; } }
        """;
}
