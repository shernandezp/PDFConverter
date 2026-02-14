using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.DocumentObjectModel;

namespace PDFConverter;

internal static class WordHelpers
{
        public static byte[]? GetImageBytesFromWord(WordprocessingDocument doc, string relationshipId)
        {
            if (doc == null || string.IsNullOrEmpty(relationshipId)) return null;

            // Helper to try get image part from an OpenXmlPart by relationship id
            static byte[]? TryGetFromPart(OpenXmlPart? part, string relId)
            {
                try
                {
                    if (part == null) return null;
                    var p = part.GetPartById(relId);
                    if (p is ImagePart ip)
                    {
                        using var s = ip.GetStream();
                        using var ms = new MemoryStream();
                        s.CopyTo(ms);
                        return ms.ToArray();
                    }
                }
                catch { }
                return null;
            }

            // Try main document part
            var result = TryGetFromPart(doc.MainDocumentPart, relationshipId);
            if (result != null) return result;

            // Try headers
            if (doc.MainDocumentPart?.HeaderParts != null)
            {
                foreach (var hp in doc.MainDocumentPart.HeaderParts)
                {
                    result = TryGetFromPart(hp, relationshipId);
                    if (result != null) return result;
                }
            }

            // Try footers
            if (doc.MainDocumentPart?.FooterParts != null)
            {
                foreach (var fp in doc.MainDocumentPart.FooterParts)
                {
                    result = TryGetFromPart(fp, relationshipId);
                    if (result != null) return result;
                }
            }

            // Try footnotes and endnotes parts
            result = TryGetFromPart(doc.MainDocumentPart?.FootnotesPart, relationshipId);
            if (result != null) return result;
            result = TryGetFromPart(doc.MainDocumentPart?.EndnotesPart, relationshipId);
            if (result != null) return result;

            // Try searching child relationships of all parts reachable from the main document part
            try
            {
                var allParents = new List<OpenXmlPart> { doc.MainDocumentPart };
                if (doc.MainDocumentPart?.HeaderParts != null) allParents.AddRange(doc.MainDocumentPart.HeaderParts);
                if (doc.MainDocumentPart?.FooterParts != null) allParents.AddRange(doc.MainDocumentPart.FooterParts);
                if (doc.MainDocumentPart?.FootnotesPart != null) allParents.Add(doc.MainDocumentPart.FootnotesPart);
                if (doc.MainDocumentPart?.EndnotesPart != null) allParents.Add(doc.MainDocumentPart.EndnotesPart);

                // include other parts discovered under MainDocumentPart
                var discovered = doc.MainDocumentPart?.GetPartsOfType<OpenXmlPart>()?.ToList() ?? new List<OpenXmlPart>();
                foreach (var p in discovered) if (!allParents.Contains(p)) allParents.Add(p);

                foreach (var parent in allParents)
                {
                    try
                    {
                        foreach (var rel in parent.Parts)
                        {
                            try
                            {
                                if (rel.RelationshipId == relationshipId)
                                {
                                    var child = rel.OpenXmlPart;
                                    if (child is ImagePart ip2)
                                    {
                                        using var s = ip2.GetStream();
                                        using var ms = new MemoryStream();
                                        s.CopyTo(ms);
                                        return ms.ToArray();
                                    }
                                    // if child is a Drawing or other part, attempt to search its own child parts
                                    if (child != null)
                                    {
                                        var img = TryGetFromPart(child, relationshipId);
                                        if (img != null) return img;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // As a last resort, search any ImagePart in the package and return the first one (best-effort)
            try
            {
                var imageParts = doc.MainDocumentPart?.ImageParts?.ToList();
                if (imageParts != null && imageParts.Count > 0)
                {
                    using var s = imageParts[0].GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    return ms.ToArray();
                }
            }
            catch { }

            return null;
        }

        public static string? GetRunFontFamily(W.RunProperties? rPr)
        {
            var rf = rPr?.RunFonts;
            return rf == null ? null : rf.Ascii?.Value ?? rf.HighAnsi?.Value ?? rf.ComplexScript?.Value;
        }

        public static string? GetRunColor(W.RunProperties? rPr) => rPr?.Color?.Val?.Value;

        public static ParagraphFormat GetParagraphFormatting(W.ParagraphProperties? pPr)
        {
            var alignment = ParagraphAlignment.Left;
            double leftIndent = 0, rightIndent = 0, firstLine = 0, before = 0, after = 0;
            double? lineSpacing = null;
            string? lineRule = null;
            bool hasExplicitBefore = false, hasExplicitAfter = false;

            if (pPr == null)
                return new ParagraphFormat(alignment, leftIndent, rightIndent, firstLine, before, after, lineSpacing, lineRule, false, false);

            var justification = pPr.Justification?.Val?.Value;
            if (justification == W.JustificationValues.Center) alignment = ParagraphAlignment.Center;
            else if (justification == W.JustificationValues.Right) alignment = ParagraphAlignment.Right;
            else if (justification == W.JustificationValues.Both) alignment = ParagraphAlignment.Justify;

            if (pPr.Indentation != null)
            {
                var leftVal = pPr.Indentation.Left?.Value;
                if (!string.IsNullOrEmpty(leftVal) && double.TryParse(leftVal, out var li)) leftIndent = li / 20.0;

                var rightVal = pPr.Indentation.Right?.Value;
                if (!string.IsNullOrEmpty(rightVal) && double.TryParse(rightVal, out var ri)) rightIndent = ri / 20.0;

                var firstLineVal = pPr.Indentation.FirstLine?.Value;
                if (!string.IsNullOrEmpty(firstLineVal) && double.TryParse(firstLineVal, out var fi)) firstLine = fi / 20.0;

                var hangingVal = pPr.Indentation.Hanging?.Value;
                if (!string.IsNullOrEmpty(hangingVal) && double.TryParse(hangingVal, out var hg)) firstLine = -(hg / 20.0);
            }

            if (pPr.SpacingBetweenLines != null)
            {
                var beforeVal = pPr.SpacingBetweenLines.Before?.Value;
                if (!string.IsNullOrEmpty(beforeVal) && double.TryParse(beforeVal, out var b))
                {
                    before = b / 20.0;
                    hasExplicitBefore = true;
                }

                var afterVal = pPr.SpacingBetweenLines.After?.Value;
                if (!string.IsNullOrEmpty(afterVal) && double.TryParse(afterVal, out var a))
                {
                    after = a / 20.0;
                    hasExplicitAfter = true;
                }

                if (pPr.SpacingBetweenLines.LineRule?.Value != null)
                {
                    var ruleVal = pPr.SpacingBetweenLines.LineRule.Value;
                    if (ruleVal == W.LineSpacingRuleValues.Exact) lineRule = "Exact";
                    else if (ruleVal == W.LineSpacingRuleValues.AtLeast) lineRule = "AtLeast";
                    else lineRule = "Auto";
                }

                var lineVal = pPr.SpacingBetweenLines.Line?.Value;
                if (!string.IsNullOrEmpty(lineVal) && double.TryParse(lineVal, out var ln))
                {
                    if (lineRule is "Exact" or "AtLeast")
                    {
                        lineSpacing = ln / 20.0;
                    }
                    else
                    {
                        lineSpacing = ln / 240.0;
                        if (string.IsNullOrEmpty(lineRule)) lineRule = "Auto";
                    }
                }
            }

            return new ParagraphFormat(alignment, leftIndent, rightIndent, firstLine, before, after, lineSpacing, lineRule, hasExplicitBefore, hasExplicitAfter);
        }

        public static List<double> GetTableGridColumnWidths(W.Table table)
        {
            var widths = new List<double>();
            
            // Try to get widths from TableGrid first
            var grid = table.GetFirstChild<W.TableGrid>();
            if (grid != null)
            {
                foreach (var col in grid.Elements<W.GridColumn>())
                {
                    var wstr = col.Width?.Value;
                    if (!string.IsNullOrEmpty(wstr) && double.TryParse(wstr, out var w))
                        widths.Add(w / 20.0);
                    else
                        widths.Add(100); // default width if not specified
                }
                
                if (widths.Count > 0)
                    return widths;
            }

            // Fallback: try to get widths from the first row's cell properties
            var firstRow = table.Elements<W.TableRow>().FirstOrDefault();
            if (firstRow != null)
            {
                foreach (var cell in firstRow.Elements<W.TableCell>())
                {
                    var tcPr = cell.GetFirstChild<W.TableCellProperties>();
                    var tcW = tcPr?.GetFirstChild<W.TableCellWidth>();
                    
                    double cellWidth = 100; // default
                    var widthVal = tcW?.Width?.Value;
                    if (!string.IsNullOrEmpty(widthVal) && double.TryParse(widthVal, out var w))
                    {
                        // Width can be in different units based on Type
                        var widthType = tcW?.Type?.Value;
                        if (widthType == W.TableWidthUnitValues.Dxa)
                        {
                            cellWidth = w / 20.0;
                        }
                        else if (widthType == W.TableWidthUnitValues.Pct)
                        {
                            cellWidth = (w / 5000.0) * 500.0;
                        }
                        else
                        {
                            cellWidth = 100;
                        }
                    }
                    
                    // Handle gridSpan - the cell might span multiple columns
                    var gridSpanVal = tcPr?.GetFirstChild<W.GridSpan>()?.Val?.Value;
                    int gridSpan = gridSpanVal.HasValue ? (int)gridSpanVal.Value : 1;
                    double perColWidth = cellWidth / gridSpan;
                    
                    for (int i = 0; i < gridSpan; i++)
                    {
                        widths.Add(perColWidth);
                    }
                }
            }

            return widths;
        }

        /// <summary>
        /// Get the default paragraph properties from docDefaults/pPrDefault.
        /// </summary>
        public static W.ParagraphProperties? GetDocDefaultsParagraphProperties(MainDocumentPart mainPart)
        {
            var stylesPart = mainPart?.StyleDefinitionsPart;
            if (stylesPart?.Styles == null) return null;
            var docDefaults = stylesPart.Styles.GetFirstChild<W.DocDefaults>();
            var pPrDefault = docDefaults?.GetFirstChild<W.ParagraphPropertiesDefault>();
            return pPrDefault?.GetFirstChild<W.ParagraphProperties>();
        }

        /// <summary>
        /// Get the default run properties from docDefaults/rPrDefault.
        /// </summary>
        public static W.RunProperties? GetDocDefaultsRunProperties(MainDocumentPart mainPart)
        {
            var stylesPart = mainPart?.StyleDefinitionsPart;
            if (stylesPart?.Styles == null) return null;
            var docDefaults = stylesPart.Styles.GetFirstChild<W.DocDefaults>();
            var rPrDefault = docDefaults?.GetFirstChild<W.RunPropertiesDefault>();
            return rPrDefault?.GetFirstChild<W.RunProperties>();
        }

        public static W.RunProperties? GetStyleRunProperties(MainDocumentPart mainPart, string styleId)
        {
            var stylesPart = mainPart.StyleDefinitionsPart;
            if (stylesPart == null) return null;

            var styles = stylesPart.Styles;
            var style = styles?.Elements<W.Style>().FirstOrDefault(s => s.StyleId == styleId);

            while (style != null)
            {
                var srp = style.StyleRunProperties;
                if (srp != null)
                {
                    // StyleRunProperties is not RunProperties — build one from its children
                    var rPr = new W.RunProperties();
                    foreach (var child in srp.ChildElements)
                        rPr.AppendChild(child.CloneNode(true));
                    return rPr;
                }
                var basedOn = style.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOn)) break;
                style = styles.Elements<W.Style>().FirstOrDefault(s => s.StyleId == basedOn);
            }

            return null;
        }

        public static W.ParagraphProperties? GetStyleParagraphProperties(MainDocumentPart mainPart, string styleId)
        {
            var stylesPart = mainPart.StyleDefinitionsPart;
            if (stylesPart == null) return null;

            var styles = stylesPart.Styles;
            var style = styles?.Elements<W.Style>().FirstOrDefault(s => s.StyleId == styleId);

            while (style != null)
            {
                var spp = style.StyleParagraphProperties;
                if (spp != null)
                {
                    // StyleParagraphProperties is not ParagraphProperties — build one from its children
                    var pPr = new W.ParagraphProperties();
                    foreach (var child in spp.ChildElements)
                        pPr.AppendChild(child.CloneNode(true));
                    return pPr;
                }
                var basedOn = style.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOn)) break;
                style = styles.Elements<W.Style>().FirstOrDefault(s => s.StyleId == basedOn);
            }

            return null;
        }

        public static RunFormat ResolveRunFormatting(MainDocumentPart mainPart, W.Run run, W.Paragraph paragraph)
        {
            string? fontFamily = null;
            string? color = null;
            var bold = false;
            var italic = false;
            var underline = false;
            double? size = null;

            var rPr = run.RunProperties;
            if (rPr != null)
            {
                fontFamily ??= rPr.RunFonts?.Ascii?.Value ?? rPr.RunFonts?.HighAnsi?.Value ?? rPr.RunFonts?.ComplexScript?.Value;
                color ??= rPr.Color?.Val?.Value;
                bold = bold || (rPr.Bold != null && (rPr.Bold.Val == null || rPr.Bold.Val.Value));
                italic = italic || (rPr.Italic != null && (rPr.Italic.Val == null || rPr.Italic.Val.Value));
                underline = underline || (rPr.Underline != null && rPr.Underline.Val != null && rPr.Underline.Val.Value != W.UnderlineValues.None);
                if (size == null && rPr.FontSize?.Val != null)
                {
                    var szVal = rPr.FontSize.Val.Value;
                    if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var halfPoints))
                        size = halfPoints / 2.0;
                }
            }

            var runStyleId = rPr?.RunStyle?.Val?.Value;
            if (!string.IsNullOrEmpty(runStyleId))
            {
                var sr = GetStyleRunProperties(mainPart, runStyleId);
                if (sr != null)
                {
                    fontFamily ??= sr.RunFonts?.Ascii?.Value ?? sr.RunFonts?.HighAnsi?.Value ?? sr.RunFonts?.ComplexScript?.Value;
                    color ??= sr.Color?.Val?.Value;
                    bold = bold || (sr.Bold != null && (sr.Bold.Val == null || sr.Bold.Val.Value));
                    italic = italic || (sr.Italic != null && (sr.Italic.Val == null || sr.Italic.Val.Value));
                    underline = underline || (sr.Underline != null && sr.Underline.Val != null && sr.Underline.Val.Value != W.UnderlineValues.None);
                    if (size == null && sr.FontSize?.Val != null)
                    {
                        var szVal = sr.FontSize.Val.Value;
                        if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var sp))
                            size = sp / 2.0;
                    }
                }
            }

            var pStyleId = paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(pStyleId))
            {
                var psr = GetStyleRunProperties(mainPart, pStyleId);
                if (psr != null)
                {
                    fontFamily ??= psr.RunFonts?.Ascii?.Value ?? psr.RunFonts?.HighAnsi?.Value ?? psr.RunFonts?.ComplexScript?.Value;
                    color ??= psr.Color?.Val?.Value;
                    bold = bold || (psr.Bold != null && (psr.Bold.Val == null || psr.Bold.Val.Value));
                    italic = italic || (psr.Italic != null && (psr.Italic.Val == null || psr.Italic.Val.Value));
                    underline = underline || (psr.Underline != null && psr.Underline.Val != null && psr.Underline.Val.Value != W.UnderlineValues.None);
                    if (size == null && psr.FontSize?.Val != null)
                    {
                        var szVal = psr.FontSize.Val.Value;
                        if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var psp))
                            size = psp / 2.0;
                    }
                }
            }

            // Fallback to docDefaults rPrDefault for font family and size
            var docDefaultsRPr = GetDocDefaultsRunProperties(mainPart);
            if (docDefaultsRPr != null)
            {
                fontFamily ??= docDefaultsRPr.RunFonts?.Ascii?.Value ?? docDefaultsRPr.RunFonts?.HighAnsi?.Value ?? docDefaultsRPr.RunFonts?.ComplexScript?.Value;
                if (size == null && docDefaultsRPr.FontSize?.Val != null)
                {
                    var szVal = docDefaultsRPr.FontSize.Val.Value;
                    if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var dsp))
                        size = dsp / 2.0;
                }
            }

            return new RunFormat(fontFamily, color, bold, italic, underline, size);
        }

        public static (string numFmt, string lvlText, int? startAt) GetNumberingLevelFormat(MainDocumentPart mainPart, string? numId, int ilvl)
        {
            if (string.IsNullOrEmpty(numId)) return ("decimal", "{0}.", null);
            var numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart == null) return ("decimal", "{0}.", null);

            var xml = numberingPart.Numbering.OuterXml;
            var doc = XDocument.Parse(xml);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var numEl = doc.Root.Elements(w + "num").FirstOrDefault(n => (string?)n.Attribute(w + "numId") == numId);
            if (numEl == null) return ("decimal", "{0}.", null);

            var absId = (string?)numEl.Element(w + "abstractNumId")?.Attribute(w + "val");
            if (absId == null) return ("decimal", "{0}.", null);

            var absEl = doc.Root.Elements(w + "abstractNum").FirstOrDefault(a => (string?)a.Attribute(w + "abstractNumId") == absId);
            if (absEl == null) return ("decimal", "{0}.", null);

            var lvlEl = absEl.Elements(w + "lvl").FirstOrDefault(l => (string?)l.Attribute(w + "ilvl") == ilvl.ToString());
            if (lvlEl == null) return ("decimal", "{0}.", null);

            var fmt = (string?)lvlEl.Element(w + "numFmt")?.Attribute(w + "val") ?? "decimal";
            var text = (string?)lvlEl.Element(w + "lvlText")?.Attribute(w + "val") ?? "{0}.";

            int? startAt = null;
            var startEl = lvlEl.Element(w + "start");
            if (startEl != null && int.TryParse((string?)startEl.Attribute(w + "val"), out var s)) startAt = s;

            return (fmt, text, startAt);
        }

        public static BorderInfo GetWordCellBorders(W.TableCellProperties? tcPr)
        {
            return GetWordCellBorders(tcPr, null);
        }

        public static BorderInfo GetWordCellBorders(W.TableCellProperties? tcPr, W.TableProperties? tblPr)
        {
            double topW = 0, bottomW = 0, leftW = 0, rightW = 0;
            string? topColor = null, bottomColor = null, leftColor = null, rightColor = null;
            string? topStyle = null, bottomStyle = null, leftStyle = null, rightStyle = null;
            double padTop = 0, padBottom = 0, padLeft = 0, padRight = 0;

            // Start with table-level borders as defaults
            var tblBorders = tblPr?.GetFirstChild<W.TableBorders>();
            if (tblBorders != null)
            {
                ReadBorderEdge(tblBorders.TopBorder, out topW, out topColor, out topStyle);
                ReadBorderEdge(tblBorders.BottomBorder, out bottomW, out bottomColor, out bottomStyle);
                ReadBorderEdge(tblBorders.LeftBorder, out leftW, out leftColor, out leftStyle);
                ReadBorderEdge(tblBorders.RightBorder, out rightW, out rightColor, out rightStyle);

                // InsideH and InsideV apply to internal cell edges
                var insideH = tblBorders.InsideHorizontalBorder;
                var insideV = tblBorders.InsideVerticalBorder;
                if (insideH != null)
                {
                    ReadBorderEdge(insideH, out var ihW, out var ihC, out var ihS);
                    if (ihW > 0)
                    {
                        // Apply insideH as default top/bottom for cells (overridden per-cell below)
                        if (topW == 0) { topW = ihW; topColor = ihC; topStyle = ihS; }
                        if (bottomW == 0) { bottomW = ihW; bottomColor = ihC; bottomStyle = ihS; }
                    }
                }
                if (insideV != null)
                {
                    ReadBorderEdge(insideV, out var ivW, out var ivC, out var ivS);
                    if (ivW > 0)
                    {
                        if (leftW == 0) { leftW = ivW; leftColor = ivC; leftStyle = ivS; }
                        if (rightW == 0) { rightW = ivW; rightColor = ivC; rightStyle = ivS; }
                    }
                }
            }

            if (tcPr == null)
                return new BorderInfo(topW, topColor, topStyle, bottomW, bottomColor, bottomStyle,
                    leftW, leftColor, leftStyle, rightW, rightColor, rightStyle,
                    padTop, padBottom, padLeft, padRight);

            // Override with cell-level borders if present
            var cellBorders = tcPr.GetFirstChild<W.TableCellBorders>();
            if (cellBorders != null)
            {
                if (cellBorders.TopBorder != null)
                    ReadBorderEdge(cellBorders.TopBorder, out topW, out topColor, out topStyle);
                if (cellBorders.BottomBorder != null)
                    ReadBorderEdge(cellBorders.BottomBorder, out bottomW, out bottomColor, out bottomStyle);
                if (cellBorders.LeftBorder != null)
                    ReadBorderEdge(cellBorders.LeftBorder, out leftW, out leftColor, out leftStyle);
                if (cellBorders.RightBorder != null)
                    ReadBorderEdge(cellBorders.RightBorder, out rightW, out rightColor, out rightStyle);
            }

            var margins = tcPr.GetFirstChild<W.TableCellMargin>();
            if (margins != null)
            {
                if (margins.TopMargin != null && !string.IsNullOrEmpty(margins.TopMargin.Width) && double.TryParse(margins.TopMargin.Width, out var pt)) padTop = pt / 20.0;
                if (margins.BottomMargin != null && !string.IsNullOrEmpty(margins.BottomMargin.Width) && double.TryParse(margins.BottomMargin.Width, out var pb)) padBottom = pb / 20.0;
                if (margins.LeftMargin != null && !string.IsNullOrEmpty(margins.LeftMargin.Width) && double.TryParse(margins.LeftMargin.Width, out var pl)) padLeft = pl / 20.0;
                if (margins.RightMargin != null && !string.IsNullOrEmpty(margins.RightMargin.Width) && double.TryParse(margins.RightMargin.Width, out var pr)) padRight = pr / 20.0;
            }

            return new BorderInfo(topW, topColor, topStyle, bottomW, bottomColor, bottomStyle,
                leftW, leftColor, leftStyle, rightW, rightColor, rightStyle,
                padTop, padBottom, padLeft, padRight);
        }

        private static void ReadBorderEdge(W.BorderType? border, out double width, out string? color, out string? style)
        {
            width = 0;
            color = null;
            style = null;
            if (border == null) return;

            style = border.Val?.Value.ToString();

            // sz is in eighths of a point
            if (border.Size != null && border.Size.HasValue)
                width = border.Size.Value / 8.0;
            else
                width = 0.5; // fallback if style is set but no explicit size

            var col = border.Color?.Value;
            if (!string.IsNullOrEmpty(col) && !string.Equals(col, "auto", StringComparison.OrdinalIgnoreCase))
                color = "#" + col;
        }

        /// <summary>
        /// Resolve table-level borders by merging the referenced tblStyle with inline tblPr borders.
        /// </summary>
        public static W.TableBorders? ResolveTableBorders(MainDocumentPart mainPart, W.TableProperties? tblPr)
        {
            // Start with borders from the named table style (if any)
            W.TableBorders? styleBorders = null;
            var tblStyleVal = tblPr?.GetFirstChild<W.TableStyle>()?.Val?.Value;
            if (!string.IsNullOrEmpty(tblStyleVal))
            {
                var stylesPart = mainPart?.StyleDefinitionsPart;
                if (stylesPart?.Styles != null)
                {
                    var style = stylesPart.Styles.Elements<W.Style>()
                        .FirstOrDefault(s => s.StyleId == tblStyleVal && s.Type?.Value == W.StyleValues.Table);
                    while (style != null)
                    {
                        var stblPr = style.StyleTableProperties;
                        var borders = stblPr?.GetFirstChild<W.TableBorders>();
                        if (borders != null)
                        {
                            styleBorders = borders.CloneNode(true) as W.TableBorders;
                            break;
                        }
                        var basedOn = style.BasedOn?.Val?.Value;
                        if (string.IsNullOrEmpty(basedOn)) break;
                        style = stylesPart.Styles.Elements<W.Style>()
                            .FirstOrDefault(s => s.StyleId == basedOn);
                    }
                }
            }

            // Inline borders override style borders
            var inlineBorders = tblPr?.GetFirstChild<W.TableBorders>();
            return inlineBorders ?? styleBorders;
        }

        /// <summary>
        /// Gets conditional formatting (tblStylePr) for a given condition type
        /// from the table's named style.
        /// </summary>
        public static W.TableStyleProperties? GetTableStyleConditionalFormatting(
            MainDocumentPart? mainPart, W.TableProperties? tblPr, W.TableStyleOverrideValues conditionType)
        {
            var tblStyleVal = tblPr?.GetFirstChild<W.TableStyle>()?.Val?.Value;
            if (string.IsNullOrEmpty(tblStyleVal) || mainPart?.StyleDefinitionsPart?.Styles == null)
                return null;

            var style = mainPart.StyleDefinitionsPart.Styles.Elements<W.Style>()
                .FirstOrDefault(s => s.StyleId == tblStyleVal && s.Type?.Value == W.StyleValues.Table);

            while (style != null)
            {
                var tblStylePr = style.Elements<W.TableStyleProperties>()
                    .FirstOrDefault(p => p.Type?.Value == conditionType);
                if (tblStylePr != null) return tblStylePr;
                var basedOn = style.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOn)) break;
                style = mainPart.StyleDefinitionsPart.Styles.Elements<W.Style>()
                    .FirstOrDefault(s => s.StyleId == basedOn);
            }
            return null;
        }

        /// <summary>
        /// Determines if a row matches a conditional formatting type based on tblLook and cnfStyle.
        /// Handles both new-style attributes (firstRow, lastRow) and old-style hex val bitmask.
        /// </summary>
        public static bool IsConditionalRow(W.TableRow row, W.TableProperties? tblPr,
            int rowIndex, int totalRows, W.TableStyleOverrideValues conditionType)
        {
            // Check cnfStyle on the row itself (new-style attributes or old-style val string)
            var cnf = row.GetFirstChild<W.TableRowProperties>()?.GetFirstChild<W.ConditionalFormatStyle>();
            if (cnf != null)
            {
                if (conditionType == W.TableStyleOverrideValues.FirstRow
                    && (cnf.FirstRow?.Value == true || CnfStyleBit(cnf, 0))) return true;
                if (conditionType == W.TableStyleOverrideValues.LastRow
                    && (cnf.LastRow?.Value == true || CnfStyleBit(cnf, 1))) return true;
            }

            // Fall back to tblLook and position
            var tblLook = tblPr?.GetFirstChild<W.TableLook>();
            if (tblLook != null)
            {
                bool lookFirstRow = tblLook.FirstRow?.Value == true || TblLookBit(tblLook, 0x0020);
                bool lookLastRow = tblLook.LastRow?.Value == true || TblLookBit(tblLook, 0x0040);

                if (conditionType == W.TableStyleOverrideValues.FirstRow && rowIndex == 0 && lookFirstRow) return true;
                if (conditionType == W.TableStyleOverrideValues.LastRow && rowIndex == totalRows - 1 && lookLastRow) return true;
            }
            return false;
        }

        /// <summary>
        /// Determines if a cell matches a conditional column formatting type.
        /// Handles both new-style attributes and old-style bitmasks.
        /// </summary>
        public static bool IsConditionalColumn(W.TableCell cell, W.TableProperties? tblPr, int colIndex, int totalCols)
        {
            var cnf = cell.GetFirstChild<W.TableCellProperties>()?.GetFirstChild<W.ConditionalFormatStyle>();
            if (cnf?.FirstColumn?.Value == true || (cnf != null && CnfStyleBit(cnf, 2))) return true;

            var tblLook = tblPr?.GetFirstChild<W.TableLook>();
            if (tblLook != null && colIndex == 0
                && (tblLook.FirstColumn?.Value == true || TblLookBit(tblLook, 0x0080))) return true;
            return false;
        }

        /// <summary>
        /// Parses old-style tblLook hex val attribute for a specific bit flag.
        /// </summary>
        private static bool TblLookBit(W.TableLook tblLook, int bitMask)
        {
            var valAttr = tblLook.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
            if (string.IsNullOrEmpty(valAttr.Value)) return false;
            try
            {
                int val = Convert.ToInt32(valAttr.Value, 16);
                return (val & bitMask) != 0;
            }
            catch { return false; }
        }

        /// <summary>
        /// Parses old-style cnfStyle val attribute (12-digit binary string) for a specific bit position.
        /// Position 0=firstRow, 1=lastRow, 2=firstCol, 3=lastCol, etc.
        /// </summary>
        private static bool CnfStyleBit(W.ConditionalFormatStyle cnf, int position)
        {
            var valAttr = cnf.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
            if (string.IsNullOrEmpty(valAttr.Value) || position >= valAttr.Value.Length) return false;
            return valAttr.Value[position] == '1';
        }
}
