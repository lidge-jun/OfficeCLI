using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueId = 0;

        // Check for empty paragraphs
        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            var text = ExtractParagraphText(para);
            if (string.IsNullOrWhiteSpace(text))
            {
                // Skip — empty paragraphs are normal spacing
                continue;
            }

            // Check for PUA characters (corruption indicator)
            if (text.Any(c => c >= '\uE000' && c <= '\uF8FF'))
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"HWPX-{++issueId:D3}",
                    Type = IssueType.Content,
                    Severity = IssueSeverity.Warning,
                    Path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]",
                    Message = "Paragraph contains Private Use Area characters",
                    Context = text[..Math.Min(text.Length, 50)]
                });
            }
        }

        // Check for tables with inconsistent column counts
        foreach (var (section, tbl, tblIdx) in _doc.AllTables())
        {
            var rows = tbl.Elements(HwpxNs.Hp + "tr").ToList();
            if (rows.Count == 0) continue;

            var expectedCols = (int?)tbl.Attribute("colCnt") ?? -1;
            foreach (var (row, rowIdx) in rows.Select((r, i) => (r, i)))
            {
                // Sum colSpan values (handles merged cells); GetCellAddr is defined in this partial class
                var colSpanSum = row.Elements(HwpxNs.Hp + "tc")
                    .Sum(tc => (int?)GetCellAddr(tc).ColSpan ?? 1);
                if (expectedCols >= 0 && colSpanSum != expectedCols)
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"HWPX-{++issueId:D3}",
                        Type = IssueType.Structure,
                        Severity = IssueSeverity.Error,
                        Path = $"/section[{section.Index + 1}]/tbl[{tblIdx + 1}]/tr[{rowIdx + 1}]",
                        Message = $"Row colSpan sum {colSpanSum} != expected {expectedCols}",
                        Context = null
                    });
                }
            }
        }

        // Check for missing header.xml
        if (_doc.Header == null)
        {
            issues.Add(new DocumentIssue
            {
                Id = $"HWPX-{++issueId:D3}",
                Type = IssueType.Structure,
                Severity = IssueSeverity.Warning,
                Path = "/",
                Message = "Document missing header.xml (style definitions unavailable)",
                Context = null
            });
        }

        // Level 7: BinData integrity — orphan/missing binary references
        var referencedBinData = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sec in _doc.Sections)
        {
            foreach (var el in sec.Root.Descendants())
            {
                var binRef = el.Attribute("binaryItemIDRef")?.Value;
                if (binRef != null) referencedBinData.Add(binRef);
            }
        }
        var actualBinData = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in _doc.Archive.Entries)
        {
            if (entry.FullName.Contains("BinData/", StringComparison.OrdinalIgnoreCase))
                actualBinData.Add(System.IO.Path.GetFileNameWithoutExtension(entry.FullName));
        }
        foreach (var missing in referencedBinData.Except(actualBinData))
        {
            issues.Add(new DocumentIssue
            {
                Id = $"HWPX-{++issueId:D3}", Type = IssueType.Structure,
                Severity = IssueSeverity.Error, Path = "/BinData",
                Message = $"Referenced binary '{missing}' not found in archive",
                Context = null
            });
        }
        foreach (var orphan in actualBinData.Except(referencedBinData))
        {
            issues.Add(new DocumentIssue
            {
                Id = $"HWPX-{++issueId:D3}", Type = IssueType.Structure,
                Severity = IssueSeverity.Info, Path = "/BinData",
                Message = $"Orphan binary '{orphan}' not referenced by any element",
                Context = null
            });
        }

        // Level 8: Field pair validation — unclosed fieldBegin/fieldEnd
        foreach (var sec in _doc.Sections)
        {
            var fieldBegins = sec.Root.Descendants(HwpxNs.Hp + "fieldBegin").ToList();
            var fieldEnds = sec.Root.Descendants(HwpxNs.Hp + "fieldEnd").ToList();
            if (fieldBegins.Count != fieldEnds.Count)
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"HWPX-{++issueId:D3}", Type = IssueType.Structure,
                    Severity = IssueSeverity.Warning,
                    Path = $"/section[{sec.Index + 1}]",
                    Message = $"Field count mismatch: {fieldBegins.Count} opens vs {fieldEnds.Count} closes",
                    Context = null
                });
            }
        }

        // Level 9: Section count consistency — manifest vs actual
        if (_doc.ManifestDoc != null)
        {
            var manifestSections = _doc.ManifestDoc.Descendants()
                .Count(e => e.Attribute("media-type")?.Value == "application/xml"
                    && (e.Attribute("href")?.Value?.StartsWith("section") ?? false));
            if (manifestSections != _doc.Sections.Count)
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"HWPX-{++issueId:D3}", Type = IssueType.Structure,
                    Severity = IssueSeverity.Error, Path = "/content.hpf",
                    Message = $"Section count mismatch: manifest={manifestSections}, loaded={_doc.Sections.Count}",
                    Context = null
                });
            }
        }

        // Filter by type
        if (issueType != null)
        {
            var filterType = Enum.Parse<IssueType>(issueType, ignoreCase: true);
            issues = issues.Where(i => i.Type == filterType).ToList();
        }

        // Apply limit
        if (limit.HasValue)
            issues = issues.Take(limit.Value).ToList();

        return issues;
    }
}
