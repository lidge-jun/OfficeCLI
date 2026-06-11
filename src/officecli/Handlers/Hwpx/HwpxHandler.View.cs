using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== View Shared Helpers ====================

    /// <summary>Extract all text from a paragraph's hp:run/hp:t elements.</summary>
    private static string ExtractParagraphText(XElement para)
    {
        var runs = para.Elements(HwpxNs.Hp + "run");
        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            foreach (var t in run.Elements(HwpxNs.Hp + "t"))
            {
                sb.Append(t.Value);
            }
            // Handle equations — extract Hancom equation script text
            // Element name is hp:equation (confirmed by hwpxlib). hp:eqEdit is legacy HWP5 class name.
            var eqEl = run.Element(HwpxNs.Hp + "equation")
                ?? run.Element(HwpxNs.Hp + "eqEdit")
                ?? run.Descendants().FirstOrDefault(e =>
                    e.Name.LocalName == "equation" || e.Name.LocalName == "eqEdit");
            if (eqEl != null)
            {
                var script = eqEl.Element(HwpxNs.Hp + "script")?.Value
                    ?? eqEl.Attribute("script")?.Value
                    ?? eqEl.Value;
                if (!string.IsNullOrEmpty(script))
                    sb.Append($"[eq: {script}]");
            }
            // Handle line breaks
            if (run.Element(HwpxNs.Hp + "lineBreak") != null)
                sb.Append('\n');
            if (run.Element(HwpxNs.Hp + "tab") != null)
                sb.Append('\t');
        }
        return sb.ToString();
    }

    /// <summary>Get paragraph style info from attributes and header.xml lookup.</summary>
    private (string? HeadingLevel, string Alignment) GetParagraphStyleInfo(XElement para)
    {
        var styleIdRef = para.Attribute("styleIDRef")?.Value;
        var paraPrIdRef = para.Attribute("paraPrIDRef")?.Value;

        string? headingLevel = null;
        string alignment = "LEFT";

        // Look up style in header.xml
        if (_doc.Header != null && styleIdRef != null)
        {
            var style = _doc.Header.Root!.Descendants(HwpxNs.Hh + "style")
                .FirstOrDefault(s => s.Attribute("id")?.Value == styleIdRef);
            if (style != null)
            {
                var name = style.Attribute("name")?.Value ?? "";
                // Korean heading styles: "개요 1", "개요 2", etc.
                var headingMatch = System.Text.RegularExpressions.Regex.Match(name, @"개요\s*(\d+)");
                if (headingMatch.Success)
                    headingLevel = headingMatch.Groups[1].Value;
                // English heading styles
                var engMatch = System.Text.RegularExpressions.Regex.Match(name, @"(?i)heading\s*(\d+)");
                if (engMatch.Success)
                    headingLevel = engMatch.Groups[1].Value;
            }
        }

        // Look up paragraph properties for alignment and heading
        if (_doc.Header != null && paraPrIdRef != null)
        {
            var paraPr = _doc.Header.Root!.Descendants(HwpxNs.Hh + "paraPr")
                .FirstOrDefault(p => p.Attribute("id")?.Value == paraPrIdRef);
            if (paraPr != null)
            {
                // Real HWPX: alignment is a child element <hh:align horizontal="LEFT"/>
                var alignEl = paraPr.Element(HwpxNs.Hh + "align");
                alignment = alignEl?.Attribute("horizontal")?.Value ?? "LEFT";

                // Heading detection via paraPr > heading element (type="OUTLINE")
                if (headingLevel == null)
                {
                    var heading = paraPr.Element(HwpxNs.Hh + "heading");
                    if (heading?.Attribute("type")?.Value == "OUTLINE"
                        && int.TryParse(heading.Attribute("level")?.Value, out var hl) && hl >= 1)
                        headingLevel = hl.ToString();
                }
            }
        }

        // F3: Legal appendix heading detection (별표/별지/별첨, 제N조 관련)
        if (headingLevel == null)
        {
            var text = ExtractParagraphText(para);
            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\s*\[?별[표지첨]\s*(?:\d+\s*)?(?:의\s*\d+\s*)?(?:\]|$)"))
                headingLevel = "2";
            else if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\s*\(제\s*\d+\s*조\s*관련\)"))
                headingLevel = "3";
            // G3: Space-tolerant legal heading detection
            else
            {
                var compacted = System.Text.RegularExpressions.Regex.Replace(text.TrimStart(), @"\s+", "");
                if (System.Text.RegularExpressions.Regex.IsMatch(compacted, @"^제\d+[장편](?![에의은을로서와가는도])"))
                    headingLevel = "1";
                else if (System.Text.RegularExpressions.Regex.IsMatch(compacted, @"^제\d+[절관](?![에의은을로서와가는도])"))
                    headingLevel = "2";
            }
        }

        // Plan 99.9.I3: Font-size ratio heading detection (fallback when outline level not set)
        if (headingLevel == null && _doc.Header != null)
        {
            var charPrIdRef = para.Elements(HwpxNs.Hp + "run")
                .FirstOrDefault()?.Attribute("charPrIDRef")?.Value;
            if (charPrIdRef != null)
            {
                var charPr = FindCharPr(charPrIdRef);
                if (charPr != null)
                {
                    double fontSize = GetFontSizePt(charPr);
                    double baseFontSize = _baseFontSizePt ??= ComputeBaseFontSize();
                    if (baseFontSize > 0)
                    {
                        double ratio = fontSize / baseFontSize;
                        if (ratio >= 1.5) headingLevel = "1";       // H1: 150%+
                        else if (ratio >= 1.3) headingLevel = "2";  // H2: 130%+
                        else if (ratio >= 1.15) headingLevel = "3"; // H3: 115%+
                    }
                }
            }
        }

        return (headingLevel, alignment);
    }

    /// <summary>
    /// Plan 99.9.I3: Compute base (body) font size by finding the most frequent font size across all paragraphs.
    /// Used as denominator for heading ratio detection.
    /// </summary>
    private double ComputeBaseFontSize()
    {
        var sizeCounts = new Dictionary<double, int>();
        foreach (var (_, para, _) in _doc.AllParagraphs())
        {
            var charPrIdRef = para.Elements(HwpxNs.Hp + "run")
                .FirstOrDefault()?.Attribute("charPrIDRef")?.Value;
            if (charPrIdRef == null) continue;
            var charPr = FindCharPr(charPrIdRef);
            if (charPr == null) continue;
            double size = GetFontSizePt(charPr);
            sizeCounts[size] = sizeCounts.GetValueOrDefault(size) + 1;
        }
        return sizeCounts.Count > 0
            ? sizeCounts.MaxBy(kv => kv.Value).Key
            : 10.0; // default 10pt
    }

    private static int CountWords(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0;
        // Korean: each syllable cluster counts as a word boundary
        // Simple heuristic: split on whitespace, count non-empty
        return text.Split(Array.Empty<char>(), StringSplitOptions.RemoveEmptyEntries).Length;
    }

    private static string FormatHwpUnit(int hwpUnit)
    {
        var mm = hwpUnit / 283.46;
        return $"{mm:0.#}mm";
    }
}
