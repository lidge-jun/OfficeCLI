// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public class PowerPointHandler : IDocumentHandler
{
    private readonly PresentationDocument _doc;
    private readonly string _filePath;

    public PowerPointHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _doc = PresentationDocument.Open(filePath, editable);
    }

    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int slideNum = 0;
        int totalSlides = GetSlideParts().Count();

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            if (startLine.HasValue && slideNum < startLine.Value) continue;
            if (endLine.HasValue && slideNum > endLine.Value) break;

            if (maxLines.HasValue && slideNum - (startLine ?? 1) >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {maxLines.Value} of {totalSlides} slides, use --start/--end to see more)");
                break;
            }

            sb.AppendLine($"=== Slide {slideNum} ===");
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>() ?? Enumerable.Empty<Shape>();

            foreach (var shape in shapes)
            {
                var text = GetShapeText(shape);
                if (!string.IsNullOrWhiteSpace(text))
                    sb.AppendLine(text);
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int slideNum = 0;
        int totalSlides = GetSlideParts().Count();

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            if (startLine.HasValue && slideNum < startLine.Value) continue;
            if (endLine.HasValue && slideNum > endLine.Value) break;

            if (maxLines.HasValue && slideNum - (startLine ?? 1) >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {maxLines.Value} of {totalSlides} slides, use --start/--end to see more)");
                break;
            }

            sb.AppendLine($"[Slide {slideNum}]");
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.ChildElements ?? Enumerable.Empty<OpenXmlElement>();

            int shapeIdx = 0;
            foreach (var child in shapes)
            {
                if (child is Shape shape)
                {
                    // Check if shape contains equations
                    var mathElements = FindShapeMathElements(shape);
                    if (mathElements.Count > 0)
                    {
                        var latex = string.Concat(mathElements.Select(FormulaParser.ToLatex));
                        var text = GetShapeText(shape);
                        // Check for text runs NOT inside mc:Fallback
                        var hasOtherText = shape.TextBody?.Elements<Drawing.Paragraph>()
                            .SelectMany(p => p.Elements<Drawing.Run>())
                            .Any(r => !string.IsNullOrWhiteSpace(r.Text?.Text)) == true;
                        if (hasOtherText)
                            sb.AppendLine($"  [Text Box] \"{text}\" \u2190 contains equation: \"{latex}\"");
                        else
                            sb.AppendLine($"  [Equation] \"{latex}\"");
                    }
                    else
                    {
                        var text = GetShapeText(shape);
                        var type = IsTitle(shape) ? "Title" : "Text Box";

                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            var firstRun = shape.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
                            var font = firstRun?.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                                ?? firstRun?.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface
                                ?? "(default)";
                            var fontSize = firstRun?.RunProperties?.FontSize?.Value;
                            var sizeStr = fontSize.HasValue ? $"{fontSize.Value / 100}pt" : "";

                            sb.AppendLine($"  [{type}] \"{text}\" \u2190 {font} {sizeStr}");
                        }
                    }
                    shapeIdx++;
                }
                else if (child is GraphicFrame gf && gf.Descendants<Drawing.Table>().Any())
                {
                    var table = gf.Descendants<Drawing.Table>().First();
                    var tblRows = table.Elements<Drawing.TableRow>().Count();
                    var tblCols = table.Elements<Drawing.TableRow>().FirstOrDefault()?.Elements<Drawing.TableCell>().Count() ?? 0;
                    var tblName = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Table";
                    sb.AppendLine($"  [Table] \"{tblName}\" \u2190 {tblRows}x{tblCols}");
                }
                else if (child is Picture pic)
                {
                    var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture";
                    var altText = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
                    var altInfo = string.IsNullOrEmpty(altText) ? "\u26a0 no alt text" : $"alt=\"{altText}\"";
                    sb.AppendLine($"  [Picture] \"{name}\" \u2190 {altInfo}");
                }
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();
        var slideParts = GetSlideParts().ToList();

        sb.AppendLine($"File: {Path.GetFileName(_filePath)} | {slideParts.Count} slides");

        int slideNum = 0;
        foreach (var slidePart in slideParts)
        {
            slideNum++;
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>() ?? Enumerable.Empty<Shape>();

            var title = shapes.Where(IsTitle).Select(GetShapeText).FirstOrDefault(t => !string.IsNullOrWhiteSpace(t)) ?? "(untitled)";

            int textBoxes = shapes.Count(s => !IsTitle(s) && !string.IsNullOrWhiteSpace(GetShapeText(s)));
            int pictures = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Picture>().Count() ?? 0;

            var details = new List<string>();
            if (textBoxes > 0) details.Add($"{textBoxes} text box(es)");
            if (pictures > 0) details.Add($"{pictures} picture(s)");

            var detailStr = details.Count > 0 ? $" - {string.Join(", ", details)}" : "";
            sb.AppendLine($"\u251c\u2500\u2500 Slide {slideNum}: \"{title}\"{detailStr}");
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsStats()
    {
        var sb = new StringBuilder();
        var slideParts = GetSlideParts().ToList();

        int totalShapes = 0;
        int totalPictures = 0;
        int totalTextBoxes = 0;
        int slidesWithoutTitle = 0;
        int picturesWithoutAlt = 0;
        var fontCounts = new Dictionary<string, int>();

        foreach (var slidePart in slideParts)
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            var shapes = shapeTree.Elements<Shape>().ToList();
            var pictures = shapeTree.Elements<Picture>().ToList();
            totalShapes += shapes.Count;
            totalPictures += pictures.Count;
            totalTextBoxes += shapes.Count(s => !IsTitle(s));

            if (!shapes.Any(IsTitle))
                slidesWithoutTitle++;

            picturesWithoutAlt += pictures.Count(p =>
                string.IsNullOrEmpty(p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value));

            // Collect font usage
            foreach (var shape in shapes)
            {
                foreach (var run in shape.Descendants<Drawing.Run>())
                {
                    var font = run.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                        ?? run.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
                    if (font != null)
                        fontCounts[font!] = fontCounts.GetValueOrDefault(font!) + 1;
                }
            }
        }

        sb.AppendLine($"Slides: {slideParts.Count}");
        sb.AppendLine($"Total shapes: {totalShapes}");
        sb.AppendLine($"Text boxes: {totalTextBoxes}");
        sb.AppendLine($"Pictures: {totalPictures}");
        sb.AppendLine($"Slides without title: {slidesWithoutTitle}");
        sb.AppendLine($"Pictures without alt text: {picturesWithoutAlt}");

        if (fontCounts.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Font usage:");
            foreach (var (font, count) in fontCounts.OrderByDescending(kv => kv.Value))
                sb.AppendLine($"  {font}: {count} occurrence(s)");
        }

        return sb.ToString().TrimEnd();
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueNum = 0;
        int slideNum = 0;

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            var shapes = shapeTree.Elements<Shape>().ToList();
            if (!shapes.Any(IsTitle))
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"S{++issueNum}",
                    Type = IssueType.Structure,
                    Severity = IssueSeverity.Warning,
                    Path = $"/slide[{slideNum}]",
                    Message = "Slide has no title"
                });
            }

            // Check for font consistency issues
            int shapeIdx = 0;
            foreach (var shape in shapes)
            {
                var runs = shape.Descendants<Drawing.Run>().ToList();
                if (runs.Count <= 1) { shapeIdx++; continue; }

                var fonts = runs.Select(r =>
                    r.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                    ?? r.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface)
                    .Where(f => f != null).Distinct().ToList();

                if (fonts.Count > 1)
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"F{++issueNum}",
                        Type = IssueType.Format,
                        Severity = IssueSeverity.Info,
                        Path = $"/slide[{slideNum}]/shape[{shapeIdx + 1}]",
                        Message = $"Inconsistent fonts in text box: {string.Join(", ", fonts)}"
                    });
                }
                shapeIdx++;
            }

            foreach (var pic in shapeTree.Elements<Picture>())
            {
                var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
                if (string.IsNullOrEmpty(alt))
                {
                    var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "?";
                    issues.Add(new DocumentIssue
                    {
                        Id = $"F{++issueNum}",
                        Type = IssueType.Format,
                        Severity = IssueSeverity.Info,
                        Path = $"/slide[{slideNum}]",
                        Message = $"Picture \"{name}\" is missing alt text (accessibility issue)"
                    });
                }
            }

            if (limit.HasValue && issues.Count >= limit.Value) break;
        }

        return issues;
    }

    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (path == "/" || path == "")
        {
            var node = new DocumentNode { Path = "/", Type = "presentation" };
            int slideNum = 0;
            foreach (var slidePart in GetSlideParts())
            {
                slideNum++;
                var title = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>()
                    .Where(IsTitle).Select(GetShapeText).FirstOrDefault() ?? "(untitled)";

                var slideNode = new DocumentNode
                {
                    Path = $"/slide[{slideNum}]",
                    Type = "slide",
                    Preview = title
                };

                if (depth > 0)
                {
                    slideNode.Children = GetSlideChildNodes(slidePart, slideNum, depth - 1);
                    slideNode.ChildCount = slideNode.Children.Count;
                }
                else
                {
                    var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
                    slideNode.ChildCount = (shapeTree?.Elements<Shape>().Count() ?? 0)
                        + (shapeTree?.Elements<Picture>().Count() ?? 0);
                }

                node.Children.Add(slideNode);
            }
            node.ChildCount = node.Children.Count;
            return node;
        }

        // Try paragraph/run paths: /slide[N]/shape[M]/paragraph[P] or .../run[K] or .../paragraph[P]/run[K]
        var runPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/run\[(\d+)\]$");
        if (runPathMatch.Success)
        {
            var sIdx = int.Parse(runPathMatch.Groups[1].Value);
            var shIdx = int.Parse(runPathMatch.Groups[2].Value);
            var rIdx = int.Parse(runPathMatch.Groups[3].Value);
            var (_, shape) = ResolveShape(sIdx, shIdx);
            var allRuns = GetAllRuns(shape);
            if (rIdx < 1 || rIdx > allRuns.Count)
                throw new ArgumentException($"Run {rIdx} not found (shape has {allRuns.Count} runs)");
            return RunToNode(allRuns[rIdx - 1], path);
        }

        var paraPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\](?:/run\[(\d+)\])?$");
        if (paraPathMatch.Success)
        {
            var sIdx = int.Parse(paraPathMatch.Groups[1].Value);
            var shIdx = int.Parse(paraPathMatch.Groups[2].Value);
            var pIdx = int.Parse(paraPathMatch.Groups[3].Value);
            var (_, shape) = ResolveShape(sIdx, shIdx);
            var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
                ?? throw new ArgumentException("Shape has no text body");
            if (pIdx < 1 || pIdx > paragraphs.Count)
                throw new ArgumentException($"Paragraph {pIdx} not found (shape has {paragraphs.Count} paragraphs)");

            var para = paragraphs[pIdx - 1];

            if (paraPathMatch.Groups[4].Success)
            {
                // /slide[N]/shape[M]/paragraph[P]/run[K]
                var rIdx = int.Parse(paraPathMatch.Groups[4].Value);
                var paraRuns = para.Elements<Drawing.Run>().ToList();
                if (rIdx < 1 || rIdx > paraRuns.Count)
                    throw new ArgumentException($"Run {rIdx} not found (paragraph has {paraRuns.Count} runs)");
                return RunToNode(paraRuns[rIdx - 1],
                    $"/slide[{sIdx}]/shape[{shIdx}]/paragraph[{pIdx}]/run[{rIdx}]");
            }

            // /slide[N]/shape[M]/paragraph[P]
            var paraText = string.Join("", para.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? ""));
            var paraNode = new DocumentNode
            {
                Path = path,
                Type = "paragraph",
                Text = paraText
            };
            var align = para.ParagraphProperties?.Alignment;
            if (align != null && align.HasValue) paraNode.Format["align"] = align.InnerText;

            var runs = para.Elements<Drawing.Run>().ToList();
            paraNode.ChildCount = runs.Count;
            if (depth > 0)
            {
                int runIdx = 0;
                foreach (var run in runs)
                {
                    paraNode.Children.Add(RunToNode(run,
                        $"/slide[{sIdx}]/shape[{shIdx}]/paragraph[{pIdx}]/run[{runIdx + 1}]"));
                    runIdx++;
                }
            }
            return paraNode;
        }

        // Try table cell path: /slide[N]/table[M]/tr[R]/tc[C]
        var tblCellGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
        if (tblCellGetMatch.Success)
        {
            var sIdx = int.Parse(tblCellGetMatch.Groups[1].Value);
            var tIdx = int.Parse(tblCellGetMatch.Groups[2].Value);
            var rIdx = int.Parse(tblCellGetMatch.Groups[3].Value);
            var cIdx = int.Parse(tblCellGetMatch.Groups[4].Value);

            var (slidePart2, table) = ResolveTable(sIdx, tIdx);
            var tableRows = table.Elements<Drawing.TableRow>().ToList();
            if (rIdx < 1 || rIdx > tableRows.Count)
                throw new ArgumentException($"Row {rIdx} not found (table has {tableRows.Count} rows)");
            var cells = tableRows[rIdx - 1].Elements<Drawing.TableCell>().ToList();
            if (cIdx < 1 || cIdx > cells.Count)
                throw new ArgumentException($"Cell {cIdx} not found (row has {cells.Count} cells)");

            var cell = cells[cIdx - 1];
            var cellText = cell.TextBody?.InnerText ?? "";
            var cellNode = new DocumentNode
            {
                Path = path,
                Type = "tc",
                Text = cellText
            };

            // Cell fill
            var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
            var cellFillHex = tcPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (cellFillHex != null) cellNode.Format["fill"] = cellFillHex;

            // Font info from first run
            var firstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
            if (firstRun?.RunProperties != null)
            {
                var f = firstRun.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface
                    ?? firstRun.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
                if (f != null) cellNode.Format["font"] = f;
                var fs = firstRun.RunProperties.FontSize?.Value;
                if (fs.HasValue) cellNode.Format["size"] = $"{fs.Value / 100}pt";
                if (firstRun.RunProperties.Bold?.Value == true) cellNode.Format["bold"] = true;
                if (firstRun.RunProperties.Italic?.Value == true) cellNode.Format["italic"] = true;
                var colorHex = firstRun.RunProperties.GetFirstChild<Drawing.SolidFill>()
                    ?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (colorHex != null) cellNode.Format["color"] = colorHex;
            }

            return cellNode;
        }

        // Try placeholder path with type name: /slide[N]/placeholder[title]
        var phGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
        if (phGetMatch.Success && !Regex.IsMatch(path, @"^/slide\[\d+\](?:/\w+\[\d+\])?$"))
        {
            var phSlideIdx = int.Parse(phGetMatch.Groups[1].Value);
            var phId = phGetMatch.Groups[2].Value;

            var phSlideParts = GetSlideParts().ToList();
            if (phSlideIdx < 1 || phSlideIdx > phSlideParts.Count)
                throw new ArgumentException($"Slide {phSlideIdx} not found");

            var phSlidePart = phSlideParts[phSlideIdx - 1];

            // If numeric, delegate to GetPlaceholderNode
            if (int.TryParse(phId, out var phNumIdx))
                return GetPlaceholderNode(phSlidePart, phSlideIdx, phNumIdx, depth);

            // By type name: resolve the shape and return its node
            var phShape = ResolvePlaceholderShape(phSlidePart, phId);
            var ph = phShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            var shapeTree = GetSlide(phSlidePart).CommonSlideData?.ShapeTree;
            var shapeIdx = shapeTree?.Elements<Shape>().ToList().IndexOf(phShape) ?? 0;
            var node = ShapeToNode(phShape, phSlideIdx, shapeIdx + 1, depth);
            node.Path = path;
            node.Type = "placeholder";
            if (ph?.Type?.HasValue == true) node.Format["phType"] = ph.Type.InnerText;
            if (ph?.Index?.HasValue == true) node.Format["phIndex"] = ph.Index.Value;
            return node;
        }

        // Try resolving logical paths with deeper segments (e.g. /slide[1]/table[1]/tr[1])
        // Only for paths not handled by dedicated handlers above
        if (Regex.IsMatch(path, @"^/slide\[\d+\]/(table\[\d+\]/(tr|tc)|placeholder\[\w+\]/)"))
        {
            var logicalResolved = ResolveLogicalPath(path);
            if (logicalResolved.HasValue)
                return GenericXmlQuery.ElementToNode(logicalResolved.Value.element, path, depth);
        }

        // Parse /slide[N] or /slide[N]/shape[M]
        var match = Regex.Match(path, @"^/slide\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (!match.Success)
        {
            // Generic XML fallback: navigate by element localName
            var allSegments = GenericXmlQuery.ParsePathSegments(path);
            if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                throw new ArgumentException($"Path must start with /slide[N]: {path}");

            var fbSlideIdx = allSegments[0].Index!.Value;
            var fbSlideParts = GetSlideParts().ToList();
            if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                throw new ArgumentException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})");

            OpenXmlElement fbCurrent = GetSlide(fbSlideParts[fbSlideIdx - 1]);
            var remaining = allSegments.Skip(1).ToList();
            if (remaining.Count > 0)
            {
                var target = GenericXmlQuery.NavigateByPath(fbCurrent, remaining);
                if (target == null)
                    return new DocumentNode { Path = path, Type = "error", Text = $"Element not found: {path}" };
                return GenericXmlQuery.ElementToNode(target, path, depth);
            }
            return GenericXmlQuery.ElementToNode(fbCurrent, path, depth);
        }

        var slideIdx = int.Parse(match.Groups[1].Value);
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var targetSlidePart = slideParts[slideIdx - 1];

        if (!match.Groups[2].Success)
        {
            // Return slide node
            var slideNode = new DocumentNode
            {
                Path = path,
                Type = "slide",
                Preview = GetSlide(targetSlidePart).CommonSlideData?.ShapeTree?.Elements<Shape>()
                    .Where(IsTitle).Select(GetShapeText).FirstOrDefault() ?? "(untitled)"
            };
            slideNode.Children = GetSlideChildNodes(targetSlidePart, slideIdx, depth);
            slideNode.ChildCount = slideNode.Children.Count;
            return slideNode;
        }

        // Shape or picture
        var elementType = match.Groups[2].Value;
        var elementIdx = int.Parse(match.Groups[3].Value);
        var shapeTreeEl = GetSlide(targetSlidePart).CommonSlideData?.ShapeTree;
        if (shapeTreeEl == null)
            throw new ArgumentException($"Slide {slideIdx} has no shapes");

        if (elementType == "shape")
        {
            var shapes = shapeTreeEl.Elements<Shape>().ToList();
            if (elementIdx < 1 || elementIdx > shapes.Count)
                throw new ArgumentException($"Shape {elementIdx} not found (total: {shapes.Count})");
            return ShapeToNode(shapes[elementIdx - 1], slideIdx, elementIdx, depth);
        }
        else if (elementType == "table")
        {
            var tables = shapeTreeEl.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > tables.Count)
                throw new ArgumentException($"Table {elementIdx} not found (total: {tables.Count})");
            return TableToNode(tables[elementIdx - 1], slideIdx, elementIdx, depth);
        }
        else if (elementType == "placeholder")
        {
            return GetPlaceholderNode(targetSlidePart, slideIdx, elementIdx, depth);
        }
        else if (elementType == "picture" || elementType == "pic")
        {
            var pics = shapeTreeEl.Elements<Picture>().ToList();
            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"Picture {elementIdx} not found (total: {pics.Count})");
            return PictureToNode(pics[elementIdx - 1], slideIdx, elementIdx);
        }

        // Generic fallback for unknown element types
        {
            var shapes2 = shapeTreeEl.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase)).ToList();
            if (elementIdx < 1 || elementIdx > shapes2.Count)
                throw new ArgumentException($"{elementType} {elementIdx} not found (total: {shapes2.Count})");
            return GenericXmlQuery.ElementToNode(shapes2[elementIdx - 1], path, depth);
        }
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();
        var parsed = ParseShapeSelector(selector);
        bool isEquationSelector = parsed.ElementType is "equation" or "math" or "formula";

        // Scheme B: generic XML fallback for unrecognized element types
        // Check if selector has a type that ParseShapeSelector didn't recognize
        // Extract raw element type for generic XML fallback check
        // Strip pseudo-selectors (:contains, :empty, :no-alt) and attribute filters before checking
        var selectorForType = Regex.Replace(selector, @":(contains\([^)]*\)|empty|no-alt)", "");
        var typeMatch = Regex.Match(selectorForType.Contains(']') ? selectorForType.Split(']').Last() : selectorForType, @"^(?:slide\[\d+\]\s*>?\s*)?([\w:]+)");
        var rawType = typeMatch.Success ? typeMatch.Groups[1].Value.ToLowerInvariant() : "";
        bool isKnownType = string.IsNullOrEmpty(rawType)
            || rawType is "shape" or "textbox" or "title" or "picture" or "pic"
                or "equation" or "math" or "formula"
                or "table" or "placeholder";
        if (!isKnownType)
        {
            var genericParsed = GenericXmlQuery.ParseSelector(selector);
            foreach (var slidePart in GetSlideParts())
            {
                results.AddRange(GenericXmlQuery.Query(
                    GetSlide(slidePart), genericParsed.element, genericParsed.attrs, genericParsed.containsText));
            }
            return results;
        }

        int slideNum = 0;

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;

            // Slide filter
            if (parsed.SlideNum.HasValue && parsed.SlideNum.Value != slideNum)
                continue;

            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            int shapeIdx = 0;
            foreach (var shape in shapeTree.Elements<Shape>())
            {
                if (isEquationSelector)
                {
                    var mathElements = FindShapeMathElements(shape);
                    foreach (var mathElem in mathElements)
                    {
                        var latex = FormulaParser.ToLatex(mathElem);
                        if (parsed.TextContains == null || latex.Contains(parsed.TextContains))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/slide[{slideNum}]/shape[{shapeIdx + 1}]",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "display" }
                            });
                        }
                    }
                }
                else if (MatchesShapeSelector(shape, parsed))
                {
                    results.Add(ShapeToNode(shape, slideNum, shapeIdx + 1, 0));
                }
                shapeIdx++;
            }

            if (parsed.ElementType == "picture" || parsed.ElementType == "pic" || parsed.ElementType == null)
            {
                int picIdx = 0;
                foreach (var pic in shapeTree.Elements<Picture>())
                {
                    if (MatchesPictureSelector(pic, parsed))
                    {
                        results.Add(PictureToNode(pic, slideNum, picIdx + 1));
                    }
                    picIdx++;
                }
            }

            if (parsed.ElementType == "table" || (parsed.ElementType == null && !isEquationSelector))
            {
                int tblIdx = 0;
                foreach (var gf in shapeTree.Elements<GraphicFrame>())
                {
                    if (!gf.Descendants<Drawing.Table>().Any()) continue;
                    tblIdx++;
                    var tblNode = TableToNode(gf, slideNum, tblIdx, 0);
                    if (parsed.TextContains != null)
                    {
                        // GraphicData children may be opaque when loaded from disk,
                        // so extract text from all <a:t> elements via OuterXml
                        var xml = gf.OuterXml;
                        var textMatches = Regex.Matches(xml, @"<a:t[^>]*>([^<]*)</a:t>");
                        var allText = string.Concat(textMatches.Select(m => m.Groups[1].Value));
                        if (!allText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(tblNode);
                }
            }

            if (parsed.ElementType == "placeholder")
            {
                int phIdx = 0;
                foreach (var shape in shapeTree.Elements<Shape>())
                {
                    var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    if (ph == null) continue;
                    phIdx++;

                    if (parsed.TextContains != null)
                    {
                        var shapeText = GetShapeText(shape);
                        if (!shapeText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }

                    var node = ShapeToNode(shape, slideNum, phIdx, 0);
                    node.Path = $"/slide[{slideNum}]/placeholder[{phIdx}]";
                    node.Type = "placeholder";
                    if (ph.Type?.HasValue == true) node.Format["phType"] = ph.Type.InnerText;
                    if (ph.Index?.HasValue == true) node.Format["phIndex"] = ph.Index.Value;
                    results.Add(node);
                }
            }
        }

        return results;
    }

    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        // Try run-level path: /slide[N]/shape[M]/run[K]
        var runMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/run\[(\d+)\]$");
        if (runMatch.Success)
        {
            var slideIdx = int.Parse(runMatch.Groups[1].Value);
            var shapeIdx = int.Parse(runMatch.Groups[2].Value);
            var runIdx = int.Parse(runMatch.Groups[3].Value);

            var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
            var allRuns = GetAllRuns(shape);
            if (runIdx < 1 || runIdx > allRuns.Count)
                throw new ArgumentException($"Run {runIdx} not found (shape has {allRuns.Count} runs)");

            var targetRun = allRuns[runIdx - 1];
            var unsupported = SetRunOrShapeProperties(properties, new List<Drawing.Run> { targetRun }, shape);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try paragraph/run path: /slide[N]/shape[M]/paragraph[P]/run[K]
        var paraRunMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]/run\[(\d+)\]$");
        if (paraRunMatch.Success)
        {
            var slideIdx = int.Parse(paraRunMatch.Groups[1].Value);
            var shapeIdx = int.Parse(paraRunMatch.Groups[2].Value);
            var paraIdx = int.Parse(paraRunMatch.Groups[3].Value);
            var runIdx = int.Parse(paraRunMatch.Groups[4].Value);

            var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
            var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
                ?? throw new ArgumentException("Shape has no text body");
            if (paraIdx < 1 || paraIdx > paragraphs.Count)
                throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");

            var para = paragraphs[paraIdx - 1];
            var paraRuns = para.Elements<Drawing.Run>().ToList();
            if (runIdx < 1 || runIdx > paraRuns.Count)
                throw new ArgumentException($"Run {runIdx} not found (paragraph has {paraRuns.Count} runs)");

            var targetRun = paraRuns[runIdx - 1];
            var unsupported = SetRunOrShapeProperties(properties, new List<Drawing.Run> { targetRun }, shape);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try paragraph-level path: /slide[N]/shape[M]/paragraph[P]
        var paraMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]$");
        if (paraMatch.Success)
        {
            var slideIdx = int.Parse(paraMatch.Groups[1].Value);
            var shapeIdx = int.Parse(paraMatch.Groups[2].Value);
            var paraIdx = int.Parse(paraMatch.Groups[3].Value);

            var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
            var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
                ?? throw new ArgumentException("Shape has no text body");
            if (paraIdx < 1 || paraIdx > paragraphs.Count)
                throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");

            var para = paragraphs[paraIdx - 1];
            var paraRuns = para.Elements<Drawing.Run>().ToList();
            var unsupported = new List<string>();

            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "align":
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = ParseTextAlignment(value);
                        break;
                    default:
                        // Apply run-level properties to all runs in this paragraph
                        var runUnsup = SetRunOrShapeProperties(
                            new Dictionary<string, string> { { key, value } }, paraRuns, shape);
                        unsupported.AddRange(runUnsup);
                        break;
                }
            }

            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try table cell path: /slide[N]/table[M]/tr[R]/tc[C]
        var tblCellMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
        if (tblCellMatch.Success)
        {
            var slideIdx = int.Parse(tblCellMatch.Groups[1].Value);
            var tblIdx = int.Parse(tblCellMatch.Groups[2].Value);
            var rowIdx = int.Parse(tblCellMatch.Groups[3].Value);
            var cellIdx = int.Parse(tblCellMatch.Groups[4].Value);

            var (slidePart, table) = ResolveTable(slideIdx, tblIdx);
            var tableRows = table.Elements<Drawing.TableRow>().ToList();
            if (rowIdx < 1 || rowIdx > tableRows.Count)
                throw new ArgumentException($"Row {rowIdx} not found (table has {tableRows.Count} rows)");
            var cells = tableRows[rowIdx - 1].Elements<Drawing.TableCell>().ToList();
            if (cellIdx < 1 || cellIdx > cells.Count)
                throw new ArgumentException($"Cell {cellIdx} not found (row has {cells.Count} cells)");

            var cell = cells[cellIdx - 1];
            var unsupported = SetTableCellProperties(cell, properties);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try table-level path: /slide[N]/table[M]
        var tblMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]$");
        if (tblMatch.Success)
        {
            var slideIdx = int.Parse(tblMatch.Groups[1].Value);
            var tblIdx = int.Parse(tblMatch.Groups[2].Value);

            var slideParts2 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts2.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slidePart = slideParts2[slideIdx - 1];
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException("Slide has no shape tree");
            var graphicFrames = shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (tblIdx < 1 || tblIdx > graphicFrames.Count)
                throw new ArgumentException($"Table {tblIdx} not found (total: {graphicFrames.Count})");

            var gf = graphicFrames[tblIdx - 1];
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "x" or "y" or "width" or "height":
                    {
                        var xfrm = gf.Transform ?? (gf.Transform = new Transform());
                        var offset = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset());
                        var extents = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents());
                        var emu = ParseEmu(value);
                        switch (key.ToLowerInvariant())
                        {
                            case "x": offset.X = emu; break;
                            case "y": offset.Y = emu; break;
                            case "width": extents.Cx = emu; break;
                            case "height": extents.Cy = emu; break;
                        }
                        break;
                    }
                    case "name":
                        var nvPr = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                        if (nvPr != null) nvPr.Name = value;
                        break;
                    default:
                        if (!GenericXmlQuery.SetGenericAttribute(gf, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try table row path: /slide[N]/table[M]/tr[R]
        var tblRowMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
        if (tblRowMatch.Success)
        {
            var slideIdx = int.Parse(tblRowMatch.Groups[1].Value);
            var tblIdx = int.Parse(tblRowMatch.Groups[2].Value);
            var rowIdx = int.Parse(tblRowMatch.Groups[3].Value);

            var (slidePart, table) = ResolveTable(slideIdx, tblIdx);
            var tableRows = table.Elements<Drawing.TableRow>().ToList();
            if (rowIdx < 1 || rowIdx > tableRows.Count)
                throw new ArgumentException($"Row {rowIdx} not found (table has {tableRows.Count} rows)");

            var row = tableRows[rowIdx - 1];
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "height":
                        row.Height = ParseEmu(value);
                        break;
                    default:
                        // Apply to all cells in this row
                        var cellUnsup = new HashSet<string>();
                        foreach (var cell in row.Elements<Drawing.TableCell>())
                        {
                            var u = SetTableCellProperties(cell, new Dictionary<string, string> { { key, value } });
                            foreach (var k in u) cellUnsup.Add(k);
                        }
                        unsupported.AddRange(cellUnsup);
                        break;
                }
            }
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try placeholder path: /slide[N]/placeholder[M] or /slide[N]/placeholder[type]
        var phMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
        if (phMatch.Success)
        {
            var slideIdx = int.Parse(phMatch.Groups[1].Value);
            var phId = phMatch.Groups[2].Value;

            var slideParts2 = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts2.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");
            var slidePart = slideParts2[slideIdx - 1];
            var shape = ResolvePlaceholderShape(slidePart, phId);

            var allRuns = shape.Descendants<Drawing.Run>().ToList();
            var unsupported = SetRunOrShapeProperties(properties, allRuns, shape);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Try shape-level path: /slide[N]/shape[M]
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
        if (match.Success)
        {
            var slideIdx = int.Parse(match.Groups[1].Value);
            var shapeIdx = int.Parse(match.Groups[2].Value);

            var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
            var allRuns = shape.Descendants<Drawing.Run>().ToList();
            var unsupported = SetRunOrShapeProperties(properties, allRuns, shape);
            GetSlide(slidePart).Save();
            return unsupported;
        }

        // Generic XML fallback: navigate to element and set attributes
        {
            SlidePart fbSlidePart;
            OpenXmlElement target;

            // Try logical path resolution first (table/placeholder paths)
            var logicalResult = ResolveLogicalPath(path);
            if (logicalResult.HasValue)
            {
                fbSlidePart = logicalResult.Value.slidePart;
                target = logicalResult.Value.element;
            }
            else
            {
                var allSegments = GenericXmlQuery.ParsePathSegments(path);
                if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                    throw new ArgumentException($"Path must start with /slide[N]: {path}");

                var fbSlideIdx = allSegments[0].Index!.Value;
                var fbSlideParts = GetSlideParts().ToList();
                if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                    throw new ArgumentException($"Slide {fbSlideIdx} not found");

                fbSlidePart = fbSlideParts[fbSlideIdx - 1];
                var remaining = allSegments.Skip(1).ToList();
                target = GetSlide(fbSlidePart);
                if (remaining.Count > 0)
                {
                    target = GenericXmlQuery.NavigateByPath(target, remaining)
                        ?? throw new ArgumentException($"Element not found: {path}");
                }
            }

            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            GetSlide(fbSlidePart).Save();
            return unsup;
        }
    }

    private (SlidePart slidePart, Shape shape) ResolveShape(int slideIdx, int shapeIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException($"Slide {slideIdx} has no shapes");

        var shapes = shapeTree.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > shapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found");

        return (slidePart, shapes[shapeIdx - 1]);
    }

    private (SlidePart slidePart, Drawing.Table table) ResolveTable(int slideIdx, int tblIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException($"Slide {slideIdx} has no shapes");

        var tables = shapeTree.Elements<GraphicFrame>()
            .Select(gf => gf.Descendants<Drawing.Table>().FirstOrDefault())
            .Where(t => t != null).ToList();
        if (tblIdx < 1 || tblIdx > tables.Count)
            throw new ArgumentException($"Table {tblIdx} not found (total: {tables.Count})");

        return (slidePart, tables[tblIdx - 1]!);
    }

    private static List<string> SetTableCellProperties(Drawing.TableCell cell, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                {
                    var textBody = cell.TextBody;
                    var lines = value.Replace("\\n", "\n").Split('\n');
                    if (textBody == null)
                    {
                        textBody = new Drawing.TextBody(
                            new Drawing.BodyProperties(), new Drawing.ListStyle());
                        foreach (var line in lines)
                        {
                            textBody.AppendChild(new Drawing.Paragraph(new Drawing.Run(
                                new Drawing.RunProperties { Language = "zh-CN" },
                                new Drawing.Text(line))));
                        }
                        cell.PrependChild(textBody);
                    }
                    else
                    {
                        var firstRun = textBody.Descendants<Drawing.Run>().FirstOrDefault();
                        var runProps = firstRun?.RunProperties?.CloneNode(true) as Drawing.RunProperties;
                        textBody.RemoveAllChildren<Drawing.Paragraph>();
                        foreach (var line in lines)
                        {
                            var newRun = new Drawing.Run();
                            if (runProps != null) newRun.RunProperties = runProps.CloneNode(true) as Drawing.RunProperties;
                            else newRun.RunProperties = new Drawing.RunProperties { Language = "zh-CN" };
                            newRun.Text = new Drawing.Text(line);
                            textBody.Append(new Drawing.Paragraph(newRun));
                        }
                    }
                    break;
                }
                case "font":
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        rProps.Append(new Drawing.LatinFont { Typeface = value });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                    }
                    break;
                case "size":
                    var sz = int.Parse(value) * 100;
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sz;
                    }
                    break;
                case "bold":
                    var b = bool.Parse(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = b;
                    }
                    break;
                case "italic":
                    var it = bool.Parse(value);
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = it;
                    }
                    break;
                case "color":
                    foreach (var run in cell.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        var sf = new Drawing.SolidFill();
                        sf.Append(new Drawing.RgbColorModelHex { Val = value.ToUpperInvariant() });
                        rProps.AppendChild(sf);
                    }
                    break;
                case "fill":
                {
                    var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                    if (tcPr == null)
                    {
                        tcPr = new Drawing.TableCellProperties();
                        cell.Append(tcPr);
                    }
                    tcPr.RemoveAllChildren<Drawing.SolidFill>();
                    tcPr.RemoveAllChildren<Drawing.NoFill>();
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        tcPr.Append(new Drawing.NoFill());
                    }
                    else
                    {
                        var sf = new Drawing.SolidFill();
                        sf.Append(new Drawing.RgbColorModelHex { Val = value.TrimStart('#').ToUpperInvariant() });
                        tcPr.Append(sf);
                    }
                    break;
                }
                case "align":
                {
                    var para = cell.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
                    if (para != null)
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = value.ToLowerInvariant() switch
                        {
                            "left" or "l" => Drawing.TextAlignmentTypeValues.Left,
                            "center" or "c" => Drawing.TextAlignmentTypeValues.Center,
                            "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
                            "justify" or "j" => Drawing.TextAlignmentTypeValues.Justified,
                            _ => throw new ArgumentException($"Unknown alignment: {value}")
                        };
                    }
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cell, key, value))
                        unsupported.Add(key);
                    break;
            }
        }
        return unsupported;
    }

    /// <summary>
    /// Resolve a logical PPT path (e.g. /slide[1]/table[1]/tr[2]) to the actual OpenXML element.
    /// Returns null if the path doesn't contain logical segments that need resolving.
    /// </summary>
    private (SlidePart slidePart, OpenXmlElement element)? ResolveLogicalPath(string path)
    {
        // /slide[N]/table[M]...
        var tblPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\](.*)$");
        if (tblPathMatch.Success)
        {
            var slideIdx = int.Parse(tblPathMatch.Groups[1].Value);
            var tblIdx = int.Parse(tblPathMatch.Groups[2].Value);
            var rest = tblPathMatch.Groups[3].Value; // e.g. /tr[1]/tc[2]/txBody

            var (slidePart, table) = ResolveTable(slideIdx, tblIdx);
            OpenXmlElement current = table;

            if (!string.IsNullOrEmpty(rest))
            {
                var segments = GenericXmlQuery.ParsePathSegments(rest);
                var target = GenericXmlQuery.NavigateByPath(current, segments);
                if (target != null) current = target;
                else throw new ArgumentException($"Element not found: {path}");
            }
            return (slidePart, current);
        }

        // /slide[N]/placeholder[X]...
        var phPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\](.*)$");
        if (phPathMatch.Success)
        {
            var slideIdx = int.Parse(phPathMatch.Groups[1].Value);
            var phId = phPathMatch.Groups[2].Value;
            var rest = phPathMatch.Groups[3].Value;

            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");
            var slidePart = slideParts[slideIdx - 1];
            OpenXmlElement current = ResolvePlaceholderShape(slidePart, phId);

            if (!string.IsNullOrEmpty(rest))
            {
                var segments = GenericXmlQuery.ParsePathSegments(rest);
                var target = GenericXmlQuery.NavigateByPath(current, segments);
                if (target != null) current = target;
                else throw new ArgumentException($"Element not found: {path}");
            }
            return (slidePart, current);
        }

        return null;
    }

    private static PlaceholderValues? ParsePlaceholderType(string name)
    {
        return name.ToLowerInvariant() switch
        {
            "title" => PlaceholderValues.Title,
            "centertitle" or "centeredtitle" or "ctitle" => PlaceholderValues.CenteredTitle,
            "body" or "content" => PlaceholderValues.Body,
            "subtitle" or "sub" => PlaceholderValues.SubTitle,
            "date" or "datetime" or "dt" => PlaceholderValues.DateAndTime,
            "footer" => PlaceholderValues.Footer,
            "slidenum" or "slidenumber" or "sldnum" => PlaceholderValues.SlideNumber,
            "object" or "obj" => PlaceholderValues.Object,
            "chart" => PlaceholderValues.Chart,
            "table" => PlaceholderValues.Table,
            "clipart" => PlaceholderValues.ClipArt,
            "diagram" or "dgm" => PlaceholderValues.Diagram,
            "media" => PlaceholderValues.Media,
            "picture" or "pic" => PlaceholderValues.Picture,
            "header" => PlaceholderValues.Header,
            _ => null
        };
    }

    private Shape ResolvePlaceholderShape(SlidePart slidePart, string phId)
    {
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");

        // Try numeric index first
        if (int.TryParse(phId, out var numIdx))
        {
            // Match by placeholder index
            var byIndex = shapeTree.Elements<Shape>()
                .FirstOrDefault(s =>
                {
                    var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    return ph?.Index?.Value == (uint)numIdx;
                });
            if (byIndex != null) return byIndex;

            // Also try as 1-based ordinal of all placeholders
            var allPh = shapeTree.Elements<Shape>()
                .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>() != null).ToList();
            if (numIdx >= 1 && numIdx <= allPh.Count)
                return allPh[numIdx - 1];

            throw new ArgumentException($"Placeholder index {numIdx} not found");
        }

        // Try by type name
        var phType = ParsePlaceholderType(phId)
            ?? throw new ArgumentException($"Unknown placeholder type: '{phId}'. " +
                "Known types: title, body, subtitle, date, footer, slidenum, object, picture, centerTitle");

        var byType = shapeTree.Elements<Shape>()
            .FirstOrDefault(s =>
            {
                var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>();
                return ph?.Type?.Value == phType;
            });

        if (byType != null) return byType;

        // Check layout for inherited placeholders and create one on the slide
        var layoutPart = slidePart.SlideLayoutPart;
        if (layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree != null)
        {
            var layoutShape = layoutPart.SlideLayout.CommonSlideData.ShapeTree.Elements<Shape>()
                .FirstOrDefault(s =>
                {
                    var ph = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    return ph?.Type?.Value == phType;
                });

            if (layoutShape != null)
            {
                // Clone from layout and add to slide
                var newShape = (Shape)layoutShape.CloneNode(true);
                // Clear any text content from layout placeholder
                if (newShape.TextBody != null)
                {
                    newShape.TextBody.RemoveAllChildren<Drawing.Paragraph>();
                    newShape.TextBody.Append(new Drawing.Paragraph(
                        new Drawing.EndParagraphRunProperties { Language = "zh-CN" }));
                }
                shapeTree.AppendChild(newShape);
                return newShape;
            }
        }

        throw new ArgumentException($"Placeholder '{phId}' not found on slide or its layout");
    }

    private DocumentNode GetPlaceholderNode(SlidePart slidePart, int slideIdx, int phIdx, int depth)
    {
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");

        // Get all placeholders on slide
        var placeholders = shapeTree.Elements<Shape>()
            .Where(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>() != null).ToList();

        if (phIdx < 1 || phIdx > placeholders.Count)
            throw new ArgumentException($"Placeholder {phIdx} not found (total: {placeholders.Count})");

        var shape = placeholders[phIdx - 1];
        var ph = shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<PlaceholderShape>()!;

        var node = ShapeToNode(shape, slideIdx, phIdx, depth);
        node.Path = $"/slide[{slideIdx}]/placeholder[{phIdx}]";
        node.Type = "placeholder";
        if (ph.Type?.HasValue == true) node.Format["phType"] = ph.Type.InnerText;
        if (ph.Index?.HasValue == true) node.Format["phIndex"] = ph.Index.Value;
        return node;
    }



    private static List<Drawing.Run> GetAllRuns(Shape shape)
    {
        return shape.TextBody?.Elements<Drawing.Paragraph>()
            .SelectMany(p => p.Elements<Drawing.Run>()).ToList()
            ?? new List<Drawing.Run>();
    }

    private static List<string> SetRunOrShapeProperties(
        Dictionary<string, string> properties, List<Drawing.Run> runs, Shape shape)
    {
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                {
                    var textLines = value.Replace("\\n", "\n").Split('\n');
                    if (runs.Count == 1 && textLines.Length == 1)
                    {
                        // Single run, single line: just replace its text
                        runs[0].Text = new Drawing.Text(textLines[0]);
                    }
                    else
                    {
                        // Shape-level: replace all text, preserve first run formatting
                        var textBody = shape.TextBody;
                        if (textBody != null)
                        {
                            var firstRun = textBody.Descendants<Drawing.Run>().FirstOrDefault();
                            var runProps = firstRun?.RunProperties?.CloneNode(true) as Drawing.RunProperties;

                            textBody.RemoveAllChildren<Drawing.Paragraph>();

                            foreach (var textLine in textLines)
                            {
                                var newPara = new Drawing.Paragraph();
                                var newRun = new Drawing.Run();
                                if (runProps != null)
                                    newRun.RunProperties = runProps.CloneNode(true) as Drawing.RunProperties;
                                newRun.Text = new Drawing.Text(textLine);
                                newPara.Append(newRun);
                                textBody.Append(newPara);
                            }
                        }
                    }
                    break;
                }

                case "font":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                        rProps.Append(new Drawing.LatinFont { Typeface = value });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                    }
                    break;

                case "size":
                    var sizeVal = int.Parse(value) * 100;
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sizeVal;
                    }
                    break;

                case "bold":
                    var isBold = bool.Parse(value);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = isBold;
                    }
                    break;

                case "italic":
                    var isItalic = bool.Parse(value);
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = isItalic;
                    }
                    break;

                case "color":
                    foreach (var run in runs)
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        var solidFill = new Drawing.SolidFill();
                        solidFill.Append(new Drawing.RgbColorModelHex { Val = value.ToUpperInvariant() });
                        if (rProps is OpenXmlCompositeElement composite)
                        {
                            if (!composite.AddChild(solidFill, throwOnError: false))
                                rProps.AppendChild(solidFill);
                        }
                        else
                        {
                            rProps.AppendChild(solidFill);
                        }
                    }
                    break;

                case "fill":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    ApplyShapeFill(spPr, value);
                    break;
                }

                case "margin":
                {
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    ApplyTextMargin(bodyPr, value);
                    break;
                }

                case "align":
                {
                    var alignment = ParseTextAlignment(value);
                    foreach (var para in shape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = alignment;
                    }
                    break;
                }

                case "valign":
                {
                    var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null) { unsupported.Add(key); break; }
                    bodyPr.Anchor = value.ToLowerInvariant() switch
                    {
                        "top" or "t" => Drawing.TextAnchoringTypeValues.Top,
                        "center" or "middle" or "c" or "m" => Drawing.TextAnchoringTypeValues.Center,
                        "bottom" or "b" => Drawing.TextAnchoringTypeValues.Bottom,
                        _ => throw new ArgumentException($"Invalid valign: {value}. Use top/center/bottom")
                    };
                    break;
                }

                case "preset":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var existingGeom = spPr.GetFirstChild<Drawing.PresetGeometry>();
                    if (existingGeom != null)
                        existingGeom.Preset = ParsePresetShape(value);
                    else
                        spPr.AppendChild(new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = ParsePresetShape(value) });
                    break;
                }

                case "x" or "y" or "width" or "height":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) { unsupported.Add(key); break; }
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    var offset = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset());
                    var extents = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents());
                    var emu = ParseEmu(value);
                    switch (key.ToLowerInvariant())
                    {
                        case "x": offset.X = emu; break;
                        case "y": offset.Y = emu; break;
                        case "width": extents.Cx = emu; break;
                        case "height": extents.Cy = emu; break;
                    }
                    break;
                }

                default:
                    if (!GenericXmlQuery.SetGenericAttribute(shape, key, value))
                        unsupported.Add(key);
                    break;
            }
        }

        return unsupported;
    }

    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        switch (type.ToLowerInvariant())
        {
            case "slide":
                var presentationPart = _doc.PresentationPart
                    ?? throw new InvalidOperationException("Presentation not found");
                var presentation = presentationPart.Presentation
                    ?? throw new InvalidOperationException("No presentation");
                var slideIdList = presentation.GetFirstChild<SlideIdList>()
                    ?? presentation.AppendChild(new SlideIdList());

                var newSlidePart = presentationPart.AddNewPart<SlidePart>();

                // Link slide to slideLayout (required by PowerPoint)
                var slideMasterPart = presentationPart.SlideMasterParts.FirstOrDefault();
                if (slideMasterPart != null)
                {
                    var slideLayoutPart = slideMasterPart.SlideLayoutParts.FirstOrDefault();
                    if (slideLayoutPart != null)
                    {
                        newSlidePart.AddPart(slideLayoutPart);
                    }
                }

                newSlidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties { Id = 1, Name = "" },
                                new NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties()
                        )
                    )
                );

                // Add title shape if text provided
                if (properties.TryGetValue("title", out var titleText))
                {
                    var titleShape = CreateTextShape(1, "Title", titleText, true);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(titleShape);
                }

                // Add content text if provided
                if (properties.TryGetValue("text", out var contentText))
                {
                    var textShape = CreateTextShape(2, "Content", contentText, false);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(textShape);
                }

                newSlidePart.Slide.Save();

                var maxId = slideIdList.Elements<SlideId>().Any()
                    ? slideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255) + 1
                    : 256;
                var relId = presentationPart.GetIdOfPart(newSlidePart);

                if (index.HasValue && index.Value < slideIdList.Elements<SlideId>().Count())
                {
                    var refSlide = slideIdList.Elements<SlideId>().ElementAtOrDefault(index.Value);
                    if (refSlide != null)
                        slideIdList.InsertBefore(new SlideId { Id = maxId, RelationshipId = relId }, refSlide);
                    else
                        slideIdList.AppendChild(new SlideId { Id = maxId, RelationshipId = relId });
                }
                else
                {
                    slideIdList.AppendChild(new SlideId { Id = maxId, RelationshipId = relId });
                }

                presentation.Save();
                var slideCount = slideIdList.Elements<SlideId>().Count();
                return $"/slide[{slideCount}]";

            case "shape" or "textbox":
                var slideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException($"Shapes must be added to a slide: /slide[N]");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide {slideIdx} not found");

                var slidePart = slideParts[slideIdx - 1];
                var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var text = properties.GetValueOrDefault("text", "");
                var shapeName = properties.GetValueOrDefault("name", $"TextBox {shapeTree.Elements<Shape>().Count() + 1}");
                var shapeId = (uint)(shapeTree.Elements<Shape>().Count() + shapeTree.Elements<Picture>().Count() + 2);

                var newShape = CreateTextShape(shapeId, shapeName, text, false);

                if (properties.TryGetValue("font", out var font))
                {
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Append(new Drawing.LatinFont { Typeface = font });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = font });
                    }
                }
                if (properties.TryGetValue("size", out var sizeStr))
                {
                    var sizeVal = int.Parse(sizeStr) * 100;
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sizeVal;
                    }
                }
                if (properties.TryGetValue("bold", out var boldStr))
                {
                    var isBold = bool.Parse(boldStr);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = isBold;
                    }
                }
                if (properties.TryGetValue("italic", out var italicStr))
                {
                    var isItalic = bool.Parse(italicStr);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = isItalic;
                    }
                }
                if (properties.TryGetValue("color", out var colorVal))
                {
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        var solidFill = new Drawing.SolidFill();
                        solidFill.Append(new Drawing.RgbColorModelHex { Val = colorVal.ToUpperInvariant() });
                        if (rProps is OpenXmlCompositeElement composite)
                        {
                            if (!composite.AddChild(solidFill, throwOnError: false))
                                rProps.AppendChild(solidFill);
                        }
                        else
                        {
                            rProps.AppendChild(solidFill);
                        }
                    }
                }

                // Shape fill
                if (properties.TryGetValue("fill", out var fillVal))
                {
                    ApplyShapeFill(newShape.ShapeProperties!, fillVal);
                }

                // Text margin (padding inside shape)
                if (properties.TryGetValue("margin", out var marginVal))
                {
                    var bodyPr = newShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr != null)
                        ApplyTextMargin(bodyPr, marginVal);
                }

                // Text alignment (horizontal)
                if (properties.TryGetValue("align", out var alignVal))
                {
                    var alignment = ParseTextAlignment(alignVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = alignment;
                    }
                }

                // Vertical alignment
                if (properties.TryGetValue("valign", out var valignVal))
                {
                    var bodyPr = newShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr != null)
                    {
                        bodyPr.Anchor = valignVal.ToLowerInvariant() switch
                        {
                            "top" or "t" => Drawing.TextAnchoringTypeValues.Top,
                            "center" or "middle" or "c" or "m" => Drawing.TextAnchoringTypeValues.Center,
                            "bottom" or "b" => Drawing.TextAnchoringTypeValues.Bottom,
                            _ => throw new ArgumentException($"Invalid valign: {valignVal}. Use top/center/bottom")
                        };
                    }
                }

                // Position and size (in EMU, 1cm = 360000 EMU; or parse as cm/in)
                {
                    long xEmu = 0, yEmu = 0;
                    long cxEmu = 9144000, cyEmu = 742950; // default: ~25.4cm x ~2.06cm
                    if (properties.TryGetValue("x", out var xStr)) xEmu = ParseEmu(xStr);
                    if (properties.TryGetValue("y", out var yStr)) yEmu = ParseEmu(yStr);
                    if (properties.TryGetValue("width", out var wStr)) cxEmu = ParseEmu(wStr);
                    if (properties.TryGetValue("height", out var hStr)) cyEmu = ParseEmu(hStr);

                    newShape.ShapeProperties!.Transform2D = new Drawing.Transform2D
                    {
                        Offset = new Drawing.Offset { X = xEmu, Y = yEmu },
                        Extents = new Drawing.Extents { Cx = cxEmu, Cy = cyEmu }
                    };
                    var presetName = properties.GetValueOrDefault("preset", "rect");
                    newShape.ShapeProperties.AppendChild(
                        new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = ParsePresetShape(presetName) }
                    );
                }

                shapeTree.AppendChild(newShape);
                GetSlide(slidePart).Save();
                var shapeCount = shapeTree.Elements<Shape>().Count();
                return $"/slide[{slideIdx}]/shape[{shapeCount}]";

            case "picture" or "image" or "img":
            {
                if (!properties.TryGetValue("path", out var imgPath))
                    throw new ArgumentException("'path' property is required for picture type");
                if (!File.Exists(imgPath))
                    throw new FileNotFoundException($"Image file not found: {imgPath}");

                var imgSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!imgSlideMatch.Success)
                    throw new ArgumentException($"Pictures must be added to a slide: /slide[N]");

                var imgSlideIdx = int.Parse(imgSlideMatch.Groups[1].Value);
                var imgSlideParts = GetSlideParts().ToList();
                if (imgSlideIdx < 1 || imgSlideIdx > imgSlideParts.Count)
                    throw new ArgumentException($"Slide {imgSlideIdx} not found");

                var imgSlidePart = imgSlideParts[imgSlideIdx - 1];
                var imgShapeTree = GetSlide(imgSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                // Determine image type
                var imgExtension = Path.GetExtension(imgPath).ToLowerInvariant();
                var imgPartType = imgExtension switch
                {
                    ".png" => ImagePartType.Png,
                    ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                    ".gif" => ImagePartType.Gif,
                    ".bmp" => ImagePartType.Bmp,
                    ".tif" or ".tiff" => ImagePartType.Tiff,
                    ".emf" => ImagePartType.Emf,
                    ".wmf" => ImagePartType.Wmf,
                    ".svg" => ImagePartType.Svg,
                    _ => throw new ArgumentException($"Unsupported image format: {imgExtension}")
                };

                // Embed image into slide part
                var imagePart = imgSlidePart.AddImagePart(imgPartType);
                using (var imgStream = File.OpenRead(imgPath))
                    imagePart.FeedData(imgStream);
                var imgRelId = imgSlidePart.GetIdOfPart(imagePart);

                // Dimensions (default: 6in x 4in)
                long cxEmu = 5486400; // 6 inches in EMUs
                long cyEmu = 3657600; // 4 inches in EMUs
                if (properties.TryGetValue("width", out var widthStr))
                    cxEmu = ParseEmu(widthStr);
                if (properties.TryGetValue("height", out var heightStr))
                    cyEmu = ParseEmu(heightStr);

                // Position (default: centered on standard 10x7.5 inch slide)
                long xEmu = (9144000 - cxEmu) / 2;
                long yEmu = (6858000 - cyEmu) / 2;
                if (properties.TryGetValue("x", out var xStr))
                    xEmu = ParseEmu(xStr);
                if (properties.TryGetValue("y", out var yStr))
                    yEmu = ParseEmu(yStr);

                var imgShapeId = (uint)(imgShapeTree.Elements<Shape>().Count() + imgShapeTree.Elements<Picture>().Count() + 2);
                var imgName = properties.GetValueOrDefault("name", $"Picture {imgShapeId}");
                var altText = properties.GetValueOrDefault("alt", Path.GetFileName(imgPath));

                // Build Picture element following Open-XML-SDK conventions
                var picture = new Picture();

                picture.NonVisualPictureProperties = new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = imgShapeId, Name = imgName, Description = altText },
                    new NonVisualPictureDrawingProperties(
                        new Drawing.PictureLocks { NoChangeAspect = true }
                    ),
                    new ApplicationNonVisualDrawingProperties()
                );

                picture.BlipFill = new BlipFill();
                picture.BlipFill.Blip = new Drawing.Blip { Embed = imgRelId };
                picture.BlipFill.AppendChild(new Drawing.Stretch(new Drawing.FillRectangle()));

                picture.ShapeProperties = new ShapeProperties();
                picture.ShapeProperties.Transform2D = new Drawing.Transform2D();
                picture.ShapeProperties.Transform2D.Offset = new Drawing.Offset { X = xEmu, Y = yEmu };
                picture.ShapeProperties.Transform2D.Extents = new Drawing.Extents { Cx = cxEmu, Cy = cyEmu };
                picture.ShapeProperties.AppendChild(
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
                );

                imgShapeTree.AppendChild(picture);
                GetSlide(imgSlidePart).Save();

                var picCount = imgShapeTree.Elements<Picture>().Count();
                return $"/slide[{imgSlideIdx}]/picture[{picCount}]";
            }

            case "table":
            {
                var tblSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!tblSlideMatch.Success)
                    throw new ArgumentException("Tables must be added to a slide: /slide[N]");

                var tblSlideIdx = int.Parse(tblSlideMatch.Groups[1].Value);
                var tblSlideParts = GetSlideParts().ToList();
                if (tblSlideIdx < 1 || tblSlideIdx > tblSlideParts.Count)
                    throw new ArgumentException($"Slide {tblSlideIdx} not found");

                var tblSlidePart = tblSlideParts[tblSlideIdx - 1];
                var tblShapeTree = GetSlide(tblSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                int rows = int.Parse(properties.GetValueOrDefault("rows", "3"));
                int cols = int.Parse(properties.GetValueOrDefault("cols", "3"));
                if (rows < 1 || cols < 1)
                    throw new ArgumentException("rows and cols must be >= 1");

                // Position & size
                long tblX = properties.TryGetValue("x", out var txStr) ? ParseEmu(txStr) : 457200; // ~1.27cm
                long tblY = properties.TryGetValue("y", out var tyStr) ? ParseEmu(tyStr) : 1600200; // ~4.44cm
                long tblCx = properties.TryGetValue("width", out var twStr) ? ParseEmu(twStr) : 8229600; // ~22.86cm
                long tblCy = properties.TryGetValue("height", out var thStr) ? ParseEmu(thStr) : (long)(rows * 370840); // ~1.03cm per row
                long colWidth = tblCx / cols;
                long rowHeight = tblCy / rows;

                var tblId = (uint)(tblShapeTree.ChildElements.Count + 2);

                // Build GraphicFrame
                var graphicFrame = new GraphicFrame();
                graphicFrame.NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = tblId, Name = properties.GetValueOrDefault("name", $"Table {tblId}") },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                graphicFrame.Transform = new Transform(
                    new Drawing.Offset { X = tblX, Y = tblY },
                    new Drawing.Extents { Cx = tblCx, Cy = tblCy }
                );

                // Build table
                var table = new Drawing.Table();
                var tblProps = new Drawing.TableProperties { FirstRow = true, BandRow = true };
                table.Append(tblProps);

                var tableGrid = new Drawing.TableGrid();
                for (int c = 0; c < cols; c++)
                    tableGrid.Append(new Drawing.GridColumn { Width = colWidth });
                table.Append(tableGrid);

                for (int r = 0; r < rows; r++)
                {
                    var tableRow = new Drawing.TableRow { Height = rowHeight };
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = new Drawing.TableCell();
                        cell.Append(new Drawing.TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph(new Drawing.EndParagraphRunProperties { Language = "zh-CN" })
                        ));
                        cell.Append(new Drawing.TableCellProperties());
                        tableRow.Append(cell);
                    }
                    table.Append(tableRow);
                }

                var graphic = new Drawing.Graphic(
                    new Drawing.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }
                );
                graphicFrame.Append(graphic);
                tblShapeTree.AppendChild(graphicFrame);
                GetSlide(tblSlidePart).Save();

                var tblCount = tblShapeTree.Elements<GraphicFrame>()
                    .Count(gf => gf.Descendants<Drawing.Table>().Any());
                return $"/slide[{tblSlideIdx}]/table[{tblCount}]";
            }

            case "equation" or "formula" or "math":
            {
                if (!properties.TryGetValue("formula", out var eqFormula))
                    throw new ArgumentException("'formula' property is required for equation type");

                var eqSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!eqSlideMatch.Success)
                    throw new ArgumentException($"Equations must be added to a slide: /slide[N]");

                var eqSlideIdx = int.Parse(eqSlideMatch.Groups[1].Value);
                var eqSlideParts = GetSlideParts().ToList();
                if (eqSlideIdx < 1 || eqSlideIdx > eqSlideParts.Count)
                    throw new ArgumentException($"Slide {eqSlideIdx} not found");

                var eqSlidePart = eqSlideParts[eqSlideIdx - 1];
                var eqShapeTree = GetSlide(eqSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var eqShapeId = (uint)(eqShapeTree.Elements<Shape>().Count() + eqShapeTree.Elements<Picture>().Count() + 2);
                var eqShapeName = properties.GetValueOrDefault("name", $"Equation {eqShapeId}");

                // Parse formula to OMML
                var mathContent = FormulaParser.Parse(eqFormula);
                M.OfficeMath oMath;
                if (mathContent is M.OfficeMath directMath)
                    oMath = directMath;
                else
                    oMath = new M.OfficeMath(mathContent.CloneNode(true));

                // Build the a14:m wrapper element via raw XML
                // PPT equations are embedded as: a:p > a14:m > m:oMathPara > m:oMath
                var mathPara = new M.Paragraph(oMath);
                var a14mXml = $"<a14:m xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\">{mathPara.OuterXml}</a14:m>";

                // Create shape with equation paragraph
                var eqShape = new Shape();
                eqShape.NonVisualShapeProperties = new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = eqShapeId, Name = eqShapeName },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                var eqSpPr = new ShapeProperties();
                {
                    long eqX = 838200, eqY = 2743200;        // default: ~2.33cm, ~7.62cm
                    long eqCx = 10515600, eqCy = 2743200;    // default: ~29.21cm, ~7.62cm
                    if (properties.TryGetValue("x", out var exStr)) eqX = ParseEmu(exStr);
                    if (properties.TryGetValue("y", out var eyStr)) eqY = ParseEmu(eyStr);
                    if (properties.TryGetValue("width", out var ewStr)) eqCx = ParseEmu(ewStr);
                    if (properties.TryGetValue("height", out var ehStr)) eqCy = ParseEmu(ehStr);
                    eqSpPr.Transform2D = new Drawing.Transform2D
                    {
                        Offset = new Drawing.Offset { X = eqX, Y = eqY },
                        Extents = new Drawing.Extents { Cx = eqCx, Cy = eqCy }
                    };
                }
                eqShape.ShapeProperties = eqSpPr;

                // Create text body with math paragraph
                var bodyProps = new Drawing.BodyProperties();
                var listStyle = new Drawing.ListStyle();
                var drawingPara = new Drawing.Paragraph();

                // Build mc:AlternateContent > mc:Choice(Requires="a14") > a14:m > m:oMathPara
                var a14mElement = new OpenXmlUnknownElement("a14", "m", "http://schemas.microsoft.com/office/drawing/2010/main");
                a14mElement.AppendChild(mathPara.CloneNode(true));

                var choice = new AlternateContentChoice();
                choice.Requires = "a14";
                choice.AppendChild(a14mElement);

                // Fallback: readable text for older versions
                var fallback = new AlternateContentFallback();
                var fallbackRun = new Drawing.Run(
                    new Drawing.RunProperties { Language = "en-US" },
                    new Drawing.Text(FormulaParser.ToReadableText(mathPara))
                );
                fallback.AppendChild(fallbackRun);

                var altContent = new AlternateContent();
                altContent.AppendChild(choice);
                altContent.AppendChild(fallback);
                drawingPara.AppendChild(altContent);

                eqShape.TextBody = new TextBody(bodyProps, listStyle, drawingPara);
                eqShapeTree.AppendChild(eqShape);

                // Ensure slide root has xmlns:a14 and mc:Ignorable="a14" so PowerPoint accepts the equation
                var eqSlide = GetSlide(eqSlidePart);
                eqSlide.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                eqSlide.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                var currentIgnorable = eqSlide.MCAttributes?.Ignorable?.Value ?? "";
                if (!currentIgnorable.Contains("a14"))
                {
                    var newVal = string.IsNullOrEmpty(currentIgnorable) ? "a14" : $"{currentIgnorable} a14";
                    eqSlide.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = newVal };
                }
                eqSlide.Save();

                var eqShapeCount = eqShapeTree.Elements<Shape>().Count();
                return $"/slide[{eqSlideIdx}]/shape[{eqShapeCount}]";
            }

            default:
            {
                // Try resolving logical paths (table/placeholder) first
                var logicalResult = ResolveLogicalPath(parentPath);
                SlidePart fbSlidePart;
                OpenXmlElement fbParent;

                if (logicalResult.HasValue)
                {
                    fbSlidePart = logicalResult.Value.slidePart;
                    fbParent = logicalResult.Value.element;
                }
                else
                {
                    // Generic fallback: navigate by XML localName
                    var allSegments = GenericXmlQuery.ParsePathSegments(parentPath);
                    if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                        throw new ArgumentException($"Generic add requires a path starting with /slide[N]: {parentPath}");

                    var fbSlideIdx = allSegments[0].Index!.Value;
                    var fbSlideParts = GetSlideParts().ToList();
                    if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                        throw new ArgumentException($"Slide {fbSlideIdx} not found");

                    fbSlidePart = fbSlideParts[fbSlideIdx - 1];
                    fbParent = GetSlide(fbSlidePart);
                    var remaining = allSegments.Skip(1).ToList();
                    if (remaining.Count > 0)
                    {
                        fbParent = GenericXmlQuery.NavigateByPath(fbParent, remaining)
                            ?? throw new ArgumentException($"Parent element not found: {parentPath}");
                    }
                }

                var created = GenericXmlQuery.TryCreateTypedElement(fbParent, type, properties, index);
                if (created == null)
                    throw new ArgumentException($"Schema-invalid element type '{type}' for parent '{parentPath}'. " +
                        "Use raw-set --action append with explicit XML instead.");

                GetSlide(fbSlidePart).Save();

                // Build result path
                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
            }
        }
    }

    public void Remove(string path)
    {
        var slideMatch = Regex.Match(path, @"^/slide\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (!slideMatch.Success)
            throw new ArgumentException($"Invalid path: {path}");

        var slideIdx = int.Parse(slideMatch.Groups[1].Value);

        if (!slideMatch.Groups[2].Success)
        {
            // Remove entire slide
            var presentationPart = _doc.PresentationPart
                ?? throw new InvalidOperationException("Presentation not found");
            var presentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = presentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");

            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slideId = slideIds[slideIdx - 1];
            var relId = slideId.RelationshipId?.Value;
            slideId.Remove();
            if (relId != null)
                presentationPart.DeletePart(presentationPart.GetPartById(relId));
            presentation.Save();
            return;
        }

        // Remove shape or picture from slide
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shapes");

        var elementType = slideMatch.Groups[2].Value;
        var elementIdx = int.Parse(slideMatch.Groups[3].Value);

        if (elementType == "shape")
        {
            var shapes = shapeTree.Elements<Shape>().ToList();
            if (elementIdx < 1 || elementIdx > shapes.Count)
                throw new ArgumentException($"Shape {elementIdx} not found");
            shapes[elementIdx - 1].Remove();
        }
        else if (elementType is "picture" or "pic")
        {
            var pics = shapeTree.Elements<Picture>().ToList();
            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"Picture {elementIdx} not found");
            pics[elementIdx - 1].Remove();
        }
        else
        {
            throw new ArgumentException($"Unknown element type: {elementType}");
        }

        GetSlide(slidePart).Save();
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var slideParts = GetSlideParts().ToList();

        // Case 1: Move entire slide (reorder)
        var slideOnlyMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]$");
        if (slideOnlyMatch.Success)
        {
            var slideIdx = int.Parse(slideOnlyMatch.Groups[1].Value);
            var movePresentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = movePresentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");
            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slideId = slideIds[slideIdx - 1];
            slideId.Remove();

            if (index.HasValue)
            {
                var remaining = slideIdList.Elements<SlideId>().ToList();
                if (index.Value >= 0 && index.Value < remaining.Count)
                    remaining[index.Value].InsertBeforeSelf(slideId);
                else
                    slideIdList.AppendChild(slideId);
            }
            else
            {
                slideIdList.AppendChild(slideId);
            }

            movePresentation.Save();
            var newSlideIds = slideIdList.Elements<SlideId>().ToList();
            var newIdx = newSlideIds.IndexOf(slideId) + 1;
            return $"/slide[{newIdx}]";
        }

        // Case 2: Move element within/across slides
        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);

        // Determine target
        string effectiveParentPath;
        SlidePart tgtSlidePart;
        ShapeTree tgtShapeTree;

        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within same parent
            tgtSlidePart = srcSlidePart;
            tgtShapeTree = GetSlide(srcSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
            var srcSlideIdx = slideParts.IndexOf(srcSlidePart) + 1;
            effectiveParentPath = $"/slide[{srcSlideIdx}]";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
            if (!tgtSlideMatch.Success)
                throw new ArgumentException($"Target must be a slide: /slide[N]");
            var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
            if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {tgtSlideIdx} not found");
            tgtSlidePart = slideParts[tgtSlideIdx - 1];
            tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
        }

        srcElement.Remove();

        // Copy relationships if moving across slides
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(srcElement, srcSlidePart, tgtSlidePart);

        InsertAtPosition(tgtShapeTree, srcElement, index);

        GetSlide(srcSlidePart).Save();
        if (srcSlidePart != tgtSlidePart)
            GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(effectiveParentPath, srcElement, tgtShapeTree);
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var slideParts = GetSlideParts().ToList();

        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);
        var clone = srcElement.CloneNode(true);

        var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
        if (!tgtSlideMatch.Success)
            throw new ArgumentException($"Target must be a slide: /slide[N]");
        var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
        if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {tgtSlideIdx} not found");

        var tgtSlidePart = slideParts[tgtSlideIdx - 1];
        var tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        // Copy relationships if across slides
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(clone, srcSlidePart, tgtSlidePart);

        InsertAtPosition(tgtShapeTree, clone, index);
        GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(targetParentPath, clone, tgtShapeTree);
    }

    private (SlidePart slidePart, OpenXmlElement element) ResolveSlideElement(string path, List<SlidePart> slideParts)
    {
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/(\w+)\[(\d+)\]$");
        if (!match.Success)
            throw new ArgumentException($"Invalid element path: {path}. Expected /slide[N]/element[M]");

        var slideIdx = int.Parse(match.Groups[1].Value);
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        var elementType = match.Groups[2].Value;
        var elementIdx = int.Parse(match.Groups[3].Value);

        OpenXmlElement element = elementType switch
        {
            "shape" => shapeTree.Elements<Shape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Shape {elementIdx} not found"),
            "picture" or "pic" => shapeTree.Elements<Picture>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Picture {elementIdx} not found"),
            _ => shapeTree.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase))
                .ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"{elementType} {elementIdx} not found")
        };

        return (slidePart, element);
    }

    private static void CopyRelationships(OpenXmlElement element, SlidePart sourcePart, SlidePart targetPart)
    {
        var rNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var allElements = element.Descendants().Prepend(element);

        foreach (var el in allElements.ToList())
        {
            foreach (var attr in el.GetAttributes().ToList())
            {
                if (attr.NamespaceUri != rNsUri) continue;

                var oldRelId = attr.Value;
                if (string.IsNullOrEmpty(oldRelId)) continue;

                try
                {
                    var referencedPart = sourcePart.GetPartById(oldRelId);
                    string newRelId;
                    try
                    {
                        newRelId = targetPart.GetIdOfPart(referencedPart);
                    }
                    catch
                    {
                        newRelId = targetPart.CreateRelationshipToPart(referencedPart);
                    }

                    if (newRelId != oldRelId)
                    {
                        el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newRelId));
                    }
                }
                catch { /* Not a valid relationship, skip */ }
            }
        }
    }

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        if (index.HasValue && parent is ShapeTree)
        {
            // Skip structural elements (nvGrpSpPr, grpSpPr) that must stay at the beginning
            var contentChildren = parent.ChildElements
                .Where(e => e is not NonVisualGroupShapeProperties && e is not GroupShapeProperties)
                .ToList();
            if (index.Value >= 0 && index.Value < contentChildren.Count)
                contentChildren[index.Value].InsertBeforeSelf(element);
            else if (contentChildren.Count > 0)
                contentChildren.Last().InsertAfterSelf(element);
            else
                parent.AppendChild(element);
        }
        else if (index.HasValue)
        {
            var children = parent.ChildElements.ToList();
            if (index.Value >= 0 && index.Value < children.Count)
                children[index.Value].InsertBeforeSelf(element);
            else
                parent.AppendChild(element);
        }
        else
        {
            parent.AppendChild(element);
        }
    }

    private static string ComputeElementPath(string parentPath, OpenXmlElement element, ShapeTree shapeTree)
    {
        // Map back to semantic type names
        string typeName;
        int typeIdx;
        if (element is Shape)
        {
            typeName = "shape";
            typeIdx = shapeTree.Elements<Shape>().ToList().IndexOf((Shape)element) + 1;
        }
        else if (element is Picture)
        {
            typeName = "picture";
            typeIdx = shapeTree.Elements<Picture>().ToList().IndexOf((Picture)element) + 1;
        }
        else
        {
            typeName = element.LocalName;
            typeIdx = shapeTree.ChildElements
                .Where(e => e.LocalName == element.LocalName)
                .ToList().IndexOf(element) + 1;
        }
        return $"{parentPath}/{typeName}[{typeIdx}]";
    }

    private static Shape CreateTextShape(uint id, string name, string text, bool isTitle)
    {
        var shape = new Shape();
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = id, Name = name },
            new NonVisualShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties(
                isTitle ? new PlaceholderShape { Type = PlaceholderValues.Title } : new PlaceholderShape()
            )
        );
        var spPr = new ShapeProperties();
        if (isTitle)
        {
            // Default title position: top-center area of standard 16:9 slide
            spPr.Transform2D = new Drawing.Transform2D
            {
                Offset = new Drawing.Offset { X = 838200, Y = 365125 },    // ~2.33cm, ~1.01cm
                Extents = new Drawing.Extents { Cx = 10515600, Cy = 1325563 } // ~29.21cm, ~3.68cm
            };
        }
        else
        {
            // Default body/content position: below title
            spPr.Transform2D = new Drawing.Transform2D
            {
                Offset = new Drawing.Offset { X = 838200, Y = 1825625 },   // ~2.33cm, ~5.07cm
                Extents = new Drawing.Extents { Cx = 10515600, Cy = 4351338 } // ~29.21cm, ~12.09cm
            };
        }
        shape.ShapeProperties = spPr;
        var body = new TextBody(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle()
        );
        var lines = text.Replace("\\n", "\n").Split('\n');
        foreach (var line in lines)
        {
            body.AppendChild(new Drawing.Paragraph(
                new Drawing.Run(
                    new Drawing.RunProperties { Language = "zh-CN" },
                    new Drawing.Text(line)
                )
            ));
        }
        shape.TextBody = body;
        return shape;
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        if (partPath == "/" || partPath == "/presentation")
            return _doc.PresentationPart?.Presentation?.OuterXml ?? "(empty)";

        var match = Regex.Match(partPath, @"^/slide\[(\d+)\]$");
        if (match.Success)
        {
            var idx = int.Parse(match.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx >= 1 && idx <= slideParts.Count)
                return GetSlide(slideParts[idx - 1]).OuterXml;
            return $"(slide[{idx}] not found)";
        }

        return $"Unknown part: {partPath}. Available: /presentation, /slide[N]";
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        OpenXmlPartRootElement rootElement;

        if (partPath is "/" or "/presentation")
        {
            rootElement = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
        }
        else if (Regex.Match(partPath, @"^/slide\[(\d+)\]$") is { Success: true } slideMatch)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found");
            rootElement = GetSlide(slideParts[idx - 1]);
        }
        else if (Regex.Match(partPath, @"^/slideMaster\[(\d+)\]$") is { Success: true } masterMatch)
        {
            var idx = int.Parse(masterMatch.Groups[1].Value);
            var masters = presentationPart.SlideMasterParts.ToList();
            if (idx < 1 || idx > masters.Count)
                throw new ArgumentException($"SlideMaster {idx} not found");
            rootElement = masters[idx - 1].SlideMaster
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }
        else if (Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$") is { Success: true } layoutMatch)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (idx < 1 || idx > layouts.Count)
                throw new ArgumentException($"SlideLayout {idx} not found");
            rootElement = layouts[idx - 1].SlideLayout
                ?? throw new InvalidOperationException("Corrupt file: slide layout data missing");
        }
        else if (Regex.Match(partPath, @"^/noteSlide\[(\d+)\]$") is { Success: true } noteMatch)
        {
            var idx = int.Parse(noteMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found");
            var notesPart = slideParts[idx - 1].NotesSlidePart
                ?? throw new ArgumentException($"Slide {idx} has no notes");
            rootElement = notesPart.NotesSlide
                ?? throw new InvalidOperationException("Corrupt file: notes slide data missing");
        }
        else
        {
            throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N]");
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        Console.WriteLine($"raw-set: {affected} element(s) affected");
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a SlidePart
                var slideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException(
                        "Chart must be added under a slide: add-part <file> '/slide[N]' --type chart");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide index {slideIdx} out of range");

                var slidePart = slideParts[slideIdx - 1];
                var chartPart = slidePart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();
                var relId = slidePart.GetIdOfPart(chartPart);

                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = slidePart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/slide[{slideIdx}]/chart[{chartIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose() => _doc.Dispose();

    // ==================== Private Helpers ====================

    private static Slide GetSlide(SlidePart part) =>
        part.Slide ?? throw new InvalidOperationException("Corrupt file: slide data missing");

    private IEnumerable<SlidePart> GetSlideParts()
    {
        var presentation = _doc.PresentationPart?.Presentation;
        var slideIdList = presentation?.GetFirstChild<SlideIdList>();
        if (slideIdList == null) yield break;

        foreach (var slideId in slideIdList.Elements<SlideId>())
        {
            var relId = slideId.RelationshipId?.Value;
            if (relId == null) continue;
            yield return (SlidePart)_doc.PresentationPart!.GetPartById(relId);
        }
    }

    private static string GetShapeText(Shape shape)
    {
        var textBody = shape.TextBody;
        if (textBody == null) return "";

        var sb = new StringBuilder();
        var first = true;
        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            if (!first) sb.Append('\n');
            first = false;
            foreach (var child in para.ChildElements)
            {
                if (child is Drawing.Run run)
                    sb.Append(run.Text?.Text ?? "");
                else if (HasMathContent(child))
                    sb.Append(FormulaParser.ToReadableText(GetMathElement(child)));
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find all OMML math elements inside a shape's text body.
    /// </summary>
    private static List<OpenXmlElement> FindShapeMathElements(Shape shape)
    {
        var results = new List<OpenXmlElement>();
        var textBody = shape.TextBody;
        if (textBody == null) return results;

        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            foreach (var child in para.ChildElements)
            {
                if (HasMathContent(child))
                    results.Add(GetMathElement(child));
            }
        }
        return results;
    }

    /// <summary>
    /// Check if an element contains math content (a14:m or mc:AlternateContent with math).
    /// </summary>
    private static bool HasMathContent(OpenXmlElement element)
    {
        // Direct a14:m element
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
            return true;
        // mc:AlternateContent containing math (check both by type and localName)
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            // Check descendants for math, or check InnerXml
            if (element.Descendants().Any(e => e.LocalName == "oMath" || e.LocalName == "oMathPara"))
                return true;
            // Fallback: check raw XML for math namespace
            var innerXml = element.InnerXml;
            return innerXml.Contains("oMath");
        }
        return false;
    }

    /// <summary>
    /// Extract the OMML math element from an a14:m or mc:AlternateContent wrapper.
    /// </summary>
    private static OpenXmlElement GetMathElement(OpenXmlElement element)
    {
        // Direct a14:m → find oMath/oMathPara inside
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
        {
            // Try child elements first (works when element tree is properly parsed)
            var child = element.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (child != null) return child;

            // Try descendants
            var desc = element.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (desc != null) return desc;

            // Last resort: re-parse from InnerXml (handles case where InnerXml was set but not parsed into children)
            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;

            return element;
        }
        // mc:AlternateContent → find oMath inside Choice
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            // Find Choice element (by type or localName)
            var choice = element.ChildElements.FirstOrDefault(e => e is AlternateContentChoice || e.LocalName == "Choice");
            if (choice != null)
            {
                var a14m = choice.ChildElements.FirstOrDefault(e =>
                    e.LocalName == "m" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main");
                if (a14m != null)
                    return GetMathElement(a14m);

                // Try descendants directly
                var mathDesc = choice.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
                if (mathDesc != null)
                    return mathDesc;
            }

            // Fallback: try InnerXml parsing
            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;
        }
        return element;
    }

    /// <summary>
    /// Re-parse OMML XML string into an OpenXmlElement with navigable children.
    /// Uses OpenXmlUnknownElement which parses InnerXml into a proper child tree.
    /// </summary>
    private static OpenXmlElement? ReparseFromXml(string innerXml)
    {
        try
        {
            var xml = innerXml.Trim();
            // Find the outermost math element
            if (xml.Contains("oMathPara"))
            {
                // Extract the oMathPara element
                var startIdx = xml.IndexOf("<m:oMathPara", StringComparison.Ordinal);
                if (startIdx < 0) startIdx = xml.IndexOf("<oMathPara", StringComparison.Ordinal);
                if (startIdx >= 0)
                {
                    var endTag = xml.Contains("</m:oMathPara>") ? "</m:oMathPara>" : "</oMathPara>";
                    var endIdx = xml.IndexOf(endTag, StringComparison.Ordinal);
                    if (endIdx >= 0)
                    {
                        var oMathParaXml = xml[startIdx..(endIdx + endTag.Length)];
                        if (!oMathParaXml.Contains("xmlns:m="))
                            oMathParaXml = oMathParaXml.Replace("<m:oMathPara", "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"");
                        var wrapper = new OpenXmlUnknownElement("m", "oMathPara", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                        // Extract inner content of oMathPara
                        var innerStart = oMathParaXml.IndexOf('>') + 1;
                        var innerEnd = oMathParaXml.LastIndexOf('<');
                        if (innerStart > 0 && innerEnd > innerStart)
                            wrapper.InnerXml = oMathParaXml[innerStart..innerEnd];
                        return wrapper;
                    }
                }
            }
        }
        catch
        {
            // Ignore parse failures
        }
        return null;
    }

    private static bool IsTitle(Shape shape)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return false;

        var type = ph.Type?.Value;
        return type == PlaceholderValues.Title || type == PlaceholderValues.CenteredTitle;
    }

    private static string GetShapeName(Shape shape)
    {
        return shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "?";
    }

    // ==================== Node Builders ====================

    private List<DocumentNode> GetSlideChildNodes(SlidePart slidePart, int slideNum, int depth)
    {
        var children = new List<DocumentNode>();
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return children;

        int shapeIdx = 0;
        foreach (var shape in shapeTree.Elements<Shape>())
        {
            children.Add(ShapeToNode(shape, slideNum, shapeIdx + 1, depth));
            shapeIdx++;
        }

        int tblIdx = 0;
        foreach (var gf in shapeTree.Elements<GraphicFrame>())
        {
            if (gf.Descendants<Drawing.Table>().Any())
            {
                tblIdx++;
                children.Add(TableToNode(gf, slideNum, tblIdx, depth));
            }
        }

        int picIdx = 0;
        foreach (var pic in shapeTree.Elements<Picture>())
        {
            children.Add(PictureToNode(pic, slideNum, picIdx + 1));
            picIdx++;
        }

        return children;
    }

    private static DocumentNode TableToNode(GraphicFrame gf, int slideNum, int tblIdx, int depth)
    {
        var table = gf.Descendants<Drawing.Table>().First();
        var rows = table.Elements<Drawing.TableRow>().ToList();
        var cols = rows.FirstOrDefault()?.Elements<Drawing.TableCell>().Count() ?? 0;
        var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Table";

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/table[{tblIdx}]",
            Type = "table",
            Preview = $"{name} ({rows.Count}x{cols})",
            ChildCount = rows.Count
        };

        node.Format["name"] = name;
        node.Format["rows"] = rows.Count;
        node.Format["cols"] = cols;

        // Position
        var offset = gf.Transform?.Offset;
        if (offset != null)
        {
            if (offset.X is not null) node.Format["x"] = FormatEmu(offset.X!);
            if (offset.Y is not null) node.Format["y"] = FormatEmu(offset.Y!);
        }
        var extents = gf.Transform?.Extents;
        if (extents != null)
        {
            if (extents.Cx is not null) node.Format["width"] = FormatEmu(extents.Cx!);
            if (extents.Cy is not null) node.Format["height"] = FormatEmu(extents.Cy!);
        }

        if (depth > 0)
        {
            int rIdx = 0;
            foreach (var row in rows)
            {
                rIdx++;
                var rowNode = new DocumentNode
                {
                    Path = $"/slide[{slideNum}]/table[{tblIdx}]/tr[{rIdx}]",
                    Type = "tr",
                    ChildCount = row.Elements<Drawing.TableCell>().Count()
                };

                if (depth > 1)
                {
                    int cIdx = 0;
                    foreach (var cell in row.Elements<Drawing.TableCell>())
                    {
                        cIdx++;
                        var cellText = cell.TextBody?.InnerText ?? "";
                        var cellNode = new DocumentNode
                        {
                            Path = $"/slide[{slideNum}]/table[{tblIdx}]/tr[{rIdx}]/tc[{cIdx}]",
                            Type = "tc",
                            Text = cellText
                        };

                        // Cell fill
                        var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                        var cellFillHex = tcPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                        if (cellFillHex != null) cellNode.Format["fill"] = cellFillHex;

                        rowNode.Children.Add(cellNode);
                    }
                }
                node.Children.Add(rowNode);
            }
        }

        return node;
    }

    private static DocumentNode ShapeToNode(Shape shape, int slideNum, int shapeIdx, int depth)
    {
        var text = GetShapeText(shape);
        var name = GetShapeName(shape);
        var isTitle = IsTitle(shape);

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/shape[{shapeIdx}]",
            Type = isTitle ? "title" : "textbox",
            Text = text,
            Preview = string.IsNullOrEmpty(text) ? name : (text.Length > 50 ? text[..50] + "..." : text)
        };

        node.Format["name"] = name;
        if (isTitle) node.Format["isTitle"] = true;

        // Position and size
        var xfrm = shape.ShapeProperties?.Transform2D;
        if (xfrm != null)
        {
            if (xfrm.Offset != null)
            {
                if (xfrm.Offset.X is not null) node.Format["x"] = FormatEmu(xfrm.Offset.X!);
                if (xfrm.Offset.Y is not null) node.Format["y"] = FormatEmu(xfrm.Offset.Y!);
            }
            if (xfrm.Extents != null)
            {
                if (xfrm.Extents.Cx is not null) node.Format["width"] = FormatEmu(xfrm.Extents.Cx!);
                if (xfrm.Extents.Cy is not null) node.Format["height"] = FormatEmu(xfrm.Extents.Cy!);
            }
        }

        // Shape fill
        var shapeFill = shape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>();
        var shapeFillHex = shapeFill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (shapeFillHex != null) node.Format["fill"] = shapeFillHex;
        if (shape.ShapeProperties?.GetFirstChild<Drawing.NoFill>() != null) node.Format["fill"] = "none";

        // Shape preset
        var presetGeom = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
            node.Format["preset"] = presetGeom.Preset.InnerText;

        // Collect font info
        var firstRun = shape.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            var font = firstRun.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface
                ?? firstRun.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
            if (font != null) node.Format["font"] = font;

            var fontSize = firstRun.RunProperties.FontSize?.Value;
            if (fontSize.HasValue) node.Format["size"] = $"{fontSize.Value / 100}pt";

            if (firstRun.RunProperties.Bold?.Value == true) node.Format["bold"] = true;
            if (firstRun.RunProperties.Italic?.Value == true) node.Format["italic"] = true;
        }

        // Text margin
        var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
        if (bodyPr != null)
        {
            var lIns = bodyPr.LeftInset;
            var tIns = bodyPr.TopInset;
            var rIns = bodyPr.RightInset;
            var bIns = bodyPr.BottomInset;
            if (lIns != null || tIns != null || rIns != null || bIns != null)
            {
                // If all four are the same, show as single value
                if (lIns == tIns && tIns == rIns && rIns == bIns && lIns != null)
                    node.Format["margin"] = FormatEmu(lIns.Value);
                else
                    node.Format["margin"] = $"{FormatEmu(lIns ?? 91440)},{FormatEmu(tIns ?? 45720)},{FormatEmu(rIns ?? 91440)},{FormatEmu(bIns ?? 45720)}";
            }

            // Vertical alignment
            if (bodyPr.Anchor?.HasValue == true)
                node.Format["valign"] = bodyPr.Anchor.InnerText;
        }

        // Text alignment (from first paragraph)
        var firstPara = shape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.Alignment?.HasValue == true)
            node.Format["align"] = firstPara.ParagraphProperties.Alignment.InnerText;

        // Count paragraphs regardless of depth
        if (shape.TextBody != null)
        {
            var paragraphs = shape.TextBody.Elements<Drawing.Paragraph>().ToList();
            node.ChildCount = paragraphs.Count;

            // Include paragraph and run hierarchy at depth > 0
            if (depth > 0)
            {
                int paraIdx = 0;
                foreach (var para in paragraphs)
                {
                    var paraText = string.Join("", para.Elements<Drawing.Run>()
                        .Select(r => r.Text?.Text ?? ""));
                    var paraRuns = para.Elements<Drawing.Run>().ToList();

                    var paraNode = new DocumentNode
                    {
                        Path = $"/slide[{slideNum}]/shape[{shapeIdx}]/paragraph[{paraIdx + 1}]",
                        Type = "paragraph",
                        Text = paraText,
                        ChildCount = paraRuns.Count
                    };

                    // Add alignment info
                    var align = para.ParagraphProperties?.Alignment;
                    if (align != null && align.HasValue) paraNode.Format["align"] = align.InnerText;

                    // Include runs at depth > 1
                    if (depth > 1)
                    {
                        int runIdx = 0;
                        foreach (var run in paraRuns)
                        {
                            paraNode.Children.Add(RunToNode(run,
                                $"/slide[{slideNum}]/shape[{shapeIdx}]/paragraph[{paraIdx + 1}]/run[{runIdx + 1}]"));
                            runIdx++;
                        }
                    }

                    node.Children.Add(paraNode);
                    paraIdx++;
                }
            }
        }

        return node;
    }

    private static DocumentNode RunToNode(Drawing.Run run, string path)
    {
        var node = new DocumentNode
        {
            Path = path,
            Type = "run",
            Text = run.Text?.Text ?? ""
        };

        if (run.RunProperties != null)
        {
            var f = run.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface
                ?? run.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
            if (f != null) node.Format["font"] = f;
            var fs = run.RunProperties.FontSize?.Value;
            if (fs.HasValue) node.Format["size"] = $"{fs.Value / 100}pt";
            if (run.RunProperties.Bold?.Value == true) node.Format["bold"] = true;
            if (run.RunProperties.Italic?.Value == true) node.Format["italic"] = true;
            // Color
            var solidFill = run.RunProperties.GetFirstChild<Drawing.SolidFill>();
            var rgbHex = solidFill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (rgbHex != null) node.Format["color"] = rgbHex;
        }

        return node;
    }

    private static DocumentNode PictureToNode(Picture pic, int slideNum, int picIdx)
    {
        var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture";
        var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/picture[{picIdx}]",
            Type = "picture",
            Preview = name
        };

        node.Format["name"] = name;
        if (!string.IsNullOrEmpty(alt)) node.Format["alt"] = alt;
        else node.Format["alt"] = "(missing)";

        return node;
    }

    // ==================== Selector ====================

    private record ShapeSelector(string? ElementType, int? SlideNum, string? TextContains,
        string? FontEquals, string? FontNotEquals, bool? IsTitle, bool? HasAlt);

    private static ShapeSelector ParseShapeSelector(string selector)
    {
        string? elementType = null;
        int? slideNum = null;
        string? textContains = null;
        string? fontEquals = null;
        string? fontNotEquals = null;
        bool? isTitle = null;
        bool? hasAlt = null;

        // Check for slide prefix
        var slideMatch = Regex.Match(selector, @"slide\[(\d+)\]\s*(.*)");
        if (slideMatch.Success)
        {
            slideNum = int.Parse(slideMatch.Groups[1].Value);
            selector = slideMatch.Groups[2].Value.TrimStart('>', ' ');
        }

        // Element type
        var typeMatch = Regex.Match(selector, @"^(\w+)");
        if (typeMatch.Success)
        {
            var t = typeMatch.Groups[1].Value.ToLowerInvariant();
            if (t is "shape" or "textbox" or "title" or "picture" or "pic" or "equation" or "math" or "formula"
                or "table" or "placeholder")
                elementType = t;
        }

        // Attributes
        foreach (Match attrMatch in Regex.Matches(selector, @"\[(\w+)(!?=)([^\]]*)\]"))
        {
            var key = attrMatch.Groups[1].Value.ToLowerInvariant();
            var op = attrMatch.Groups[2].Value;
            var val = attrMatch.Groups[3].Value.Trim('\'', '"');

            switch (key)
            {
                case "font" when op == "=": fontEquals = val; break;
                case "font" when op == "!=": fontNotEquals = val; break;
                case "title": isTitle = val.ToLowerInvariant() != "false"; break;
                case "alt": hasAlt = !string.IsNullOrEmpty(val) && val.ToLowerInvariant() != "false"; break;
            }
        }

        // :contains()
        var containsMatch = Regex.Match(selector, @":contains\(['""]?(.+?)['""]?\)");
        if (containsMatch.Success) textContains = containsMatch.Groups[1].Value;

        // Element type shortcuts
        if (elementType == "title") isTitle = true;

        // :no-alt
        if (selector.Contains(":no-alt")) hasAlt = false;

        return new ShapeSelector(elementType, slideNum, textContains, fontEquals, fontNotEquals, isTitle, hasAlt);
    }

    private static bool MatchesShapeSelector(Shape shape, ShapeSelector selector)
    {
        // Element type filter
        if (selector.ElementType is "picture" or "pic" or "table" or "placeholder")
            return false;

        // Title filter
        if (selector.IsTitle.HasValue)
        {
            if (selector.IsTitle.Value != IsTitle(shape)) return false;
        }

        // Text contains
        if (selector.TextContains != null)
        {
            var text = GetShapeText(shape);
            if (!text.Contains(selector.TextContains, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        // Font filter
        var runs = shape.Descendants<Drawing.Run>().ToList();
        if (selector.FontEquals != null)
        {
            bool found = runs.Any(r =>
            {
                var font = r.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                    ?? r.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                return font != null && string.Equals(font, selector.FontEquals, StringComparison.OrdinalIgnoreCase);
            });
            if (!found) return false;
        }

        if (selector.FontNotEquals != null)
        {
            bool hasWrongFont = runs.Any(r =>
            {
                var font = r.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                    ?? r.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                return font != null && !string.Equals(font, selector.FontNotEquals, StringComparison.OrdinalIgnoreCase);
            });
            if (!hasWrongFont) return false;
        }

        return true;
    }

    private static bool MatchesPictureSelector(Picture pic, ShapeSelector selector)
    {
        // Only match if looking for pictures specifically or no type specified
        if (selector.ElementType != null && selector.ElementType != "picture" && selector.ElementType != "pic")
            return false;

        if (selector.IsTitle.HasValue) return false; // Pictures can't be titles

        // Alt text filter
        if (selector.HasAlt.HasValue)
        {
            var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
            bool hasAlt = !string.IsNullOrEmpty(alt);
            if (selector.HasAlt.Value != hasAlt) return false;
        }

        return true;
    }

    private static Drawing.ShapeTypeValues ParsePresetShape(string name)
    {
        return name.ToLowerInvariant() switch
        {
            "rect" or "rectangle" => Drawing.ShapeTypeValues.Rectangle,
            "roundrect" or "roundedrectangle" => Drawing.ShapeTypeValues.RoundRectangle,
            "ellipse" or "oval" => Drawing.ShapeTypeValues.Ellipse,
            "triangle" => Drawing.ShapeTypeValues.Triangle,
            "rtriangle" or "righttriangle" => Drawing.ShapeTypeValues.RightTriangle,
            "diamond" => Drawing.ShapeTypeValues.Diamond,
            "parallelogram" => Drawing.ShapeTypeValues.Parallelogram,
            "trapezoid" => Drawing.ShapeTypeValues.Trapezoid,
            "pentagon" => Drawing.ShapeTypeValues.Pentagon,
            "hexagon" => Drawing.ShapeTypeValues.Hexagon,
            "heptagon" => Drawing.ShapeTypeValues.Heptagon,
            "octagon" => Drawing.ShapeTypeValues.Octagon,
            "star4" => Drawing.ShapeTypeValues.Star4,
            "star5" => Drawing.ShapeTypeValues.Star5,
            "star6" => Drawing.ShapeTypeValues.Star6,
            "star8" => Drawing.ShapeTypeValues.Star8,
            "star10" => Drawing.ShapeTypeValues.Star10,
            "star12" => Drawing.ShapeTypeValues.Star12,
            "star16" => Drawing.ShapeTypeValues.Star16,
            "star24" => Drawing.ShapeTypeValues.Star24,
            "star32" => Drawing.ShapeTypeValues.Star32,
            "rightarrow" or "rarrow" => Drawing.ShapeTypeValues.RightArrow,
            "leftarrow" or "larrow" => Drawing.ShapeTypeValues.LeftArrow,
            "uparrow" => Drawing.ShapeTypeValues.UpArrow,
            "downarrow" => Drawing.ShapeTypeValues.DownArrow,
            "leftrightarrow" or "lrarrow" => Drawing.ShapeTypeValues.LeftRightArrow,
            "updownarrow" or "udarrow" => Drawing.ShapeTypeValues.UpDownArrow,
            "chevron" => Drawing.ShapeTypeValues.Chevron,
            "homeplat" or "homeplate" => Drawing.ShapeTypeValues.HomePlate,
            "plus" or "cross" => Drawing.ShapeTypeValues.Plus,
            "heart" => Drawing.ShapeTypeValues.Heart,
            "cloud" => Drawing.ShapeTypeValues.Cloud,
            "lightning" or "lightningbolt" => Drawing.ShapeTypeValues.LightningBolt,
            "sun" => Drawing.ShapeTypeValues.Sun,
            "moon" => Drawing.ShapeTypeValues.Moon,
            "arc" => Drawing.ShapeTypeValues.Arc,
            "donut" => Drawing.ShapeTypeValues.Donut,
            "nosmoking" or "blockarc" => Drawing.ShapeTypeValues.NoSmoking,
            "cube" => Drawing.ShapeTypeValues.Cube,
            "can" or "cylinder" => Drawing.ShapeTypeValues.Can,
            "line" => Drawing.ShapeTypeValues.Line,
            "decagon" => Drawing.ShapeTypeValues.Decagon,
            "dodecagon" => Drawing.ShapeTypeValues.Dodecagon,
            "ribbon" => Drawing.ShapeTypeValues.Ribbon,
            "ribbon2" => Drawing.ShapeTypeValues.Ribbon2,
            "callout1" => Drawing.ShapeTypeValues.Callout1,
            "callout2" => Drawing.ShapeTypeValues.Callout2,
            "callout3" => Drawing.ShapeTypeValues.Callout3,
            "wedgeroundrectcallout" or "callout" => Drawing.ShapeTypeValues.WedgeRoundRectangleCallout,
            "wedgeellipsecallout" => Drawing.ShapeTypeValues.WedgeEllipseCallout,
            "cloudcallout" => Drawing.ShapeTypeValues.CloudCallout,
            "flowchartprocess" or "process" => Drawing.ShapeTypeValues.FlowChartProcess,
            "flowchartdecision" or "decision" => Drawing.ShapeTypeValues.FlowChartDecision,
            "flowchartterminator" or "terminator" => Drawing.ShapeTypeValues.FlowChartTerminator,
            "flowchartdocument" => Drawing.ShapeTypeValues.FlowChartDocument,
            "flowchartinputoutput" or "io" => Drawing.ShapeTypeValues.FlowChartInputOutput,
            "brace" or "leftbrace" => Drawing.ShapeTypeValues.LeftBrace,
            "rightbrace" => Drawing.ShapeTypeValues.RightBrace,
            "leftbracket" => Drawing.ShapeTypeValues.LeftBracket,
            "rightbracket" => Drawing.ShapeTypeValues.RightBracket,
            "smileyface" or "smiley" => Drawing.ShapeTypeValues.SmileyFace,
            "foldedcorner" => Drawing.ShapeTypeValues.FoldedCorner,
            "frame" => Drawing.ShapeTypeValues.Frame,
            "gear6" => Drawing.ShapeTypeValues.Gear6,
            "gear9" => Drawing.ShapeTypeValues.Gear9,
            "notchedrightarrow" => Drawing.ShapeTypeValues.NotchedRightArrow,
            "bentuparrow" => Drawing.ShapeTypeValues.BentUpArrow,
            "curvedrightarrow" => Drawing.ShapeTypeValues.CurvedRightArrow,
            "stripedrightarrow" => Drawing.ShapeTypeValues.StripedRightArrow,
            "uturnArrow" => Drawing.ShapeTypeValues.UTurnArrow,
            "circularArrow" => Drawing.ShapeTypeValues.CircularArrow,
            _ => throw new ArgumentException(
                $"Unknown preset shape: '{name}'. Common presets: rect, roundRect, ellipse, triangle, diamond, " +
                "pentagon, hexagon, star5, rightArrow, leftArrow, chevron, plus, heart, cloud, cube, can, line, " +
                "callout, process, decision, smiley, frame, gear6")
        };
    }

    private static void ApplyShapeFill(ShapeProperties spPr, string value)
    {
        // Remove any existing fill
        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        spPr.RemoveAllChildren<Drawing.GradientFill>();
        spPr.RemoveAllChildren<Drawing.PatternFill>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            var noFill = new Drawing.NoFill();
            if (spPr is OpenXmlCompositeElement composite)
            {
                if (!composite.AddChild(noFill, throwOnError: false))
                    spPr.PrependChild(noFill);
            }
            else
                spPr.PrependChild(noFill);
        }
        else
        {
            var solidFill = new Drawing.SolidFill();
            solidFill.Append(new Drawing.RgbColorModelHex { Val = value.TrimStart('#').ToUpperInvariant() });
            if (spPr is OpenXmlCompositeElement composite)
            {
                if (!composite.AddChild(solidFill, throwOnError: false))
                    spPr.PrependChild(solidFill);
            }
            else
                spPr.PrependChild(solidFill);
        }
    }

    /// <summary>
    /// Apply text margin (padding) to a BodyProperties element.
    /// Supports: single value "0.5cm" (all sides), or "left,top,right,bottom" e.g. "0.5cm,0.3cm,0.5cm,0.3cm"
    /// </summary>
    private static void ApplyTextMargin(Drawing.BodyProperties bodyPr, string value)
    {
        var parts = value.Split(',');
        if (parts.Length == 1)
        {
            var emu = ParseEmu(parts[0]);
            bodyPr.LeftInset = (int)emu;
            bodyPr.TopInset = (int)emu;
            bodyPr.RightInset = (int)emu;
            bodyPr.BottomInset = (int)emu;
        }
        else if (parts.Length == 4)
        {
            bodyPr.LeftInset = (int)ParseEmu(parts[0].Trim());
            bodyPr.TopInset = (int)ParseEmu(parts[1].Trim());
            bodyPr.RightInset = (int)ParseEmu(parts[2].Trim());
            bodyPr.BottomInset = (int)ParseEmu(parts[3].Trim());
        }
        else
        {
            throw new ArgumentException("margin must be a single value or 4 comma-separated values (left,top,right,bottom)");
        }
    }

    private static Drawing.TextAlignmentTypeValues ParseTextAlignment(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "left" or "l" => Drawing.TextAlignmentTypeValues.Left,
            "center" or "c" => Drawing.TextAlignmentTypeValues.Center,
            "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
            "justify" or "j" => Drawing.TextAlignmentTypeValues.Justified,
            _ => throw new ArgumentException($"Invalid align: {value}. Use: left, center, right, justify")
        };
    }

    private static long ParseEmu(string value)
    {
        value = value.Trim();
        if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 360000);
        if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 914400);
        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 12700);
        if (value.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 9525);
        return long.Parse(value); // raw EMU
    }

    private static string FormatEmu(long emu)
    {
        var cm = emu / 360000.0;
        return $"{cm:0.##}cm";
    }
}
