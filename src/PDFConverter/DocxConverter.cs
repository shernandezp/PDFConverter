using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Drawing;

namespace PDFConverter;

/// <summary>
/// Converter for DOCX files to PDF.
/// </summary>
public static class DocxConverter
{
    /// <summary>
    /// Convert a DOCX file to PDF at the specified path.
    /// </summary>
    public static void DocxToPdf(string docxPath, string pdfPath)
    {
        using var word = WordprocessingDocument.Open(docxPath, false);
        DocxToPdfInternal(word, pdfPath);
    }

    /// <summary>
    /// Convert a DOCX stream to PDF at the specified path.
    /// The input stream will not be closed by this method.
    /// </summary>
    public static void DocxToPdf(System.IO.Stream docxStream, string pdfPath)
    {
        using var ms = new System.IO.MemoryStream();
        docxStream.CopyTo(ms);
        ms.Position = 0;
        using var word = WordprocessingDocument.Open(ms, false);
        DocxToPdfInternal(word, pdfPath);
    }

    /// <summary>
    /// Convert a DOCX byte[] to PDF at the specified path.
    /// </summary>
    public static void DocxToPdf(byte[] docxBytes, string pdfPath)
    {
        using var ms = new System.IO.MemoryStream(docxBytes);
        using var word = WordprocessingDocument.Open(ms, false);
        DocxToPdfInternal(word, pdfPath);
    }

    /// <summary>
    /// Convert a DOCX byte[] to PDF and return the result as a byte array.
    /// </summary>
    public static byte[] DocxToPdfBytes(byte[] docxBytes)
    {
        using var ms = new System.IO.MemoryStream(docxBytes);
        using var word = WordprocessingDocument.Open(ms, false);
        return DocxToPdfToStream(word).ToArray();
    }

    /// <summary>
    /// Convert a DOCX stream to PDF and return the result as a byte array.
    /// The input stream is not closed by this method.
    /// </summary>
    public static byte[] DocxToPdfBytes(System.IO.Stream docxStream)
    {
        using var ms = new System.IO.MemoryStream();
        docxStream.CopyTo(ms);
        ms.Position = 0;
        using var word = WordprocessingDocument.Open(ms, false);
        return DocxToPdfToStream(word).ToArray();
    }

    internal static void DocxToPdfInternal(WordprocessingDocument word, string pdfPath)
    {
        var renderer = BuildRenderer(word, out var tempFiles);
        try
        {
            renderer.Save(pdfPath);
        }
        finally
        {
            CleanupTempFiles(tempFiles);
        }
    }

    internal static System.IO.MemoryStream DocxToPdfToStream(WordprocessingDocument word)
    {
        var renderer = BuildRenderer(word, out var tempFiles);
        try
        {
            var output = new System.IO.MemoryStream();
            renderer.Save(output, false);
            output.Position = 0;
            return output;
        }
        finally
        {
            CleanupTempFiles(tempFiles);
        }
    }

    static void CleanupTempFiles(List<string> tempFiles)
    {
        foreach (var tf in tempFiles)
        {
            try { ConverterExtensions.TryDeleteTempFile(tf); } catch { }
        }
    }

    static PdfDocumentRenderer BuildRenderer(WordprocessingDocument word, out List<string> tempFiles)
    {
        // Ensure fonts are available (system first, explicit mappings as fallback)
        OpenXmlHelpers.EnsureFontResolverInitialized();

        var body = word.MainDocumentPart?.Document.Body;
        if (body == null) throw new InvalidOperationException("Document body not found");

        var doc = new Document();
        var section = doc.AddSection();
        // default margins
        section.PageSetup.LeftMargin = Unit.FromCentimeter(2.54);
        section.PageSetup.RightMargin = Unit.FromCentimeter(2.54);

        // Try to use original DOCX section properties (page size, margins/header distance) if available
        try
        {
            var sectPr = word.MainDocumentPart?.Document.Body?.Elements<W.SectionProperties>()?.LastOrDefault();
            var pgSz = sectPr?.GetFirstChild<W.PageSize>();
            if (pgSz != null)
            {
                // pgSz values are in twentieths of a point; convert to points
                // Width and Height are UInt32Value
                if (pgSz.Width != null && pgSz.Width.HasValue)
                {
                    section.PageSetup.PageWidth = Unit.FromPoint(pgSz.Width.Value / 20.0);
                }
                if (pgSz.Height != null && pgSz.Height.HasValue)
                {
                    section.PageSetup.PageHeight = Unit.FromPoint(pgSz.Height.Value / 20.0);
                }
                OpenXmlHelpers.ImageLoadLogger?.Invoke($"Applied page size: width={section.PageSetup.PageWidth.Point}pt height={section.PageSetup.PageHeight.Point}pt");

                // Apply landscape orientation if specified
                if (pgSz.Orient?.Value == W.PageOrientationValues.Landscape)
                {
                    section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
                }
            }

            var pgMar = sectPr?.GetFirstChild<W.PageMargin>();
            if (pgMar != null)
            {
                OpenXmlHelpers.ImageLoadLogger?.Invoke($"Found pgMar: top={pgMar.Top} bottom={pgMar.Bottom} left={pgMar.Left} right={pgMar.Right} header={pgMar.Header}");
                
                // Top margin - note: Top is Int32Value in OpenXML
                if (pgMar.Top != null && pgMar.Top.HasValue)
                {
                    section.PageSetup.TopMargin = Unit.FromPoint(pgMar.Top.Value / 20.0);
                }
                // Bottom margin
                if (pgMar.Bottom != null && pgMar.Bottom.HasValue)
                {
                    section.PageSetup.BottomMargin = Unit.FromPoint(pgMar.Bottom.Value / 20.0);
                }
                // Left margin - note: Left is UInt32Value in OpenXML
                if (pgMar.Left != null && pgMar.Left.HasValue)
                {
                    section.PageSetup.LeftMargin = Unit.FromPoint(pgMar.Left.Value / 20.0);
                }
                // Right margin
                if (pgMar.Right != null && pgMar.Right.HasValue)
                {
                    section.PageSetup.RightMargin = Unit.FromPoint(pgMar.Right.Value / 20.0);
                }
                // Header distance
                if (pgMar.Header != null && pgMar.Header.HasValue)
                {
                    section.PageSetup.HeaderDistance = Unit.FromPoint(pgMar.Header.Value / 20.0);
                }
                // Footer distance
                if (pgMar.Footer != null && pgMar.Footer.HasValue)
                {
                    section.PageSetup.FooterDistance = Unit.FromPoint(pgMar.Footer.Value / 20.0);
                }
                OpenXmlHelpers.ImageLoadLogger?.Invoke($"Applied margins: top={section.PageSetup.TopMargin.Point}pt bottom={section.PageSetup.BottomMargin.Point}pt headerDistance={section.PageSetup.HeaderDistance.Point}pt footerDistance={section.PageSetup.FooterDistance.Point}pt");
            }
        }
        catch { }

        var usedFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Read docDefaults paragraph spacing for fallback
        ParagraphFormat? docDefaultsFmt = null;
        try
        {
            var docDefaultsPPr = OpenXmlHelpers.GetDocDefaultsParagraphProperties(word.MainDocumentPart);
            if (docDefaultsPPr != null)
                docDefaultsFmt = OpenXmlHelpers.GetParagraphFormatting(docDefaultsPPr);
        }
        catch { }

        // helper to convert EMU to points
        static double EmuToPoints(long emu) => emu / 12700.0;

        // collect temp files to delete after rendering
        tempFiles = new List<string>();
        var backgroundFiles = new List<string>();
        // collect hyperlinked images for post-render link annotation overlay
        var hyperlinkImages = new List<(string Url, string ImagePath, double WidthPt, double HeightPt)>();

        // Render header images (if any) at top of the first section so header graphics are preserved
        try
        {
            if (word.MainDocumentPart?.HeaderParts != null)
            {
                foreach (var headerPart in word.MainDocumentPart.HeaderParts)
                {
                    var header = headerPart.Header;
                    if (header == null) continue;
                    foreach (var hp in header.Elements<W.Paragraph>())
                    {
                        var infos = ConverterExtensions.GetImageInfosFromParagraph(word, hp).ToList();
                        if (infos.Count == 0) continue;

                        // Separate background images from inline images
                        var bgImages = infos.Where(i => i.IsBackground).ToList();
                        var inlineImages = infos.Where(i => !i.IsBackground).ToList();

                        // Process background images
                        foreach (var info in bgImages)
                        {
                            if (info.Bytes == null || info.Bytes.Length == 0) continue;
                            var imgPath = ConverterExtensions.SaveTempImage(info.Bytes);
                            tempFiles.Add(imgPath);
                            if (!backgroundFiles.Contains(imgPath))
                                backgroundFiles.Add(imgPath);
                            // reserve header space based on image height
                            if (info.ExtentCyEmu.HasValue)
                            {
                                try
                                {
                                    var hpt = EmuToPoints(info.ExtentCyEmu.Value);
                                    var pageH = section.PageSetup.PageHeight.Point;
                                    if (hpt / (pageH > 0 ? pageH : 1.0) < 0.6)
                                    {
                                        var avail = Math.Max(0.0, pageH - section.PageSetup.TopMargin.Point - section.PageSetup.BottomMargin.Point - 10.0);
                                        var headerPts = Math.Min(hpt, avail);
                                        if (section.PageSetup.HeaderDistance.Point < headerPts)
                                            section.PageSetup.HeaderDistance = Unit.FromPoint(headerPts);
                                    }
                                }
                                catch { }
                            }
                        }

                        // Process inline images — add all from same Word paragraph to ONE MigraDoc paragraph
                        if (inlineImages.Count > 0)
                        {
                            try
                            {
                                var headerContainer = section.Headers.Primary ?? section.Headers.FirstPage ?? section.Headers.EvenPage;
                                if (headerContainer == null)
                                {
                                    headerContainer = new HeaderFooter();
                                    section.Headers.Primary = headerContainer;
                                }
                                var hpHeader = headerContainer.AddParagraph();
                                double maxImgHeight = 0;

                                foreach (var info in inlineImages)
                                {
                                    if (info.Bytes == null || info.Bytes.Length == 0) continue;
                                    // Apply srcRect cropping if present
                                    var imgBytes = ConverterExtensions.ApplySrcRectCrop(
                                        info.Bytes, info.CropLeft, info.CropTop, info.CropRight, info.CropBottom);
                                    var imgPath = ConverterExtensions.SaveTempImage(imgBytes);
                                    tempFiles.Add(imgPath);
                                    try
                                    {
                                        var headerImg = hpHeader.AddImage(imgPath);
                                        headerImg.LockAspectRatio = true;
                                        if (info.ExtentCxEmu.HasValue)
                                            headerImg.Width = Unit.FromPoint(EmuToPoints(info.ExtentCxEmu.Value));
                                        else
                                            headerImg.Width = Unit.FromCentimeter(8);

                                        if (info.ExtentCyEmu.HasValue)
                                        {
                                            var hpt = EmuToPoints(info.ExtentCyEmu.Value);
                                            if (hpt > maxImgHeight) maxImgHeight = hpt;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed adding header inline image: {ex.Message}");
                                    }
                                }

                                // Ensure TopMargin is large enough to fit header content without overlapping body
                                // In MigraDoc: HeaderDistance = gap from page top to header content
                                //              TopMargin = gap from page top to body content
                                // Body starts at TopMargin, header starts at HeaderDistance
                                // So we need: TopMargin >= HeaderDistance + maxImgHeight + small gap
                                if (maxImgHeight > 0)
                                {
                                    var headerDist = section.PageSetup.HeaderDistance.Point;
                                    var neededTopMargin = headerDist + maxImgHeight + 4.0;
                                    if (section.PageSetup.TopMargin.Point < neededTopMargin)
                                        section.PageSetup.TopMargin = Unit.FromPoint(neededTopMargin);
                                }
                            }
                            catch (Exception ex)
                            {
                                OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed processing header images: {ex.Message}");
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Header extraction failed: {ex.Message}");
        }

        // Render footer content (text and images) into MigraDoc footer
        try
        {
            if (word.MainDocumentPart?.FooterParts != null)
            {
                foreach (var footerPart in word.MainDocumentPart.FooterParts)
                {
                    var footer = footerPart.Footer;
                    if (footer == null) continue;

                    var footerContainer = section.Footers.Primary;
                    foreach (var fp in footer.Elements<W.Paragraph>())
                    {
                        var fPPr = fp.ParagraphProperties;
                        var fParaFmt = OpenXmlHelpers.GetParagraphFormatting(fPPr);
                        var footerPara = footerContainer.AddParagraph();
                        footerPara.Format.Alignment = fParaFmt.Alignment;
                        if (fParaFmt.SpacingBefore > 0)
                            footerPara.Format.SpaceBefore = Unit.FromPoint(fParaFmt.SpacingBefore);
                        if (fParaFmt.SpacingAfter > 0)
                            footerPara.Format.SpaceAfter = Unit.FromPoint(fParaFmt.SpacingAfter);

                        foreach (var child in fp.ChildElements)
                        {
                            if (child is W.Run run)
                            {
                                var fmt = OpenXmlHelpers.ResolveRunFormatting(word.MainDocumentPart, run, fp);
                                if (!string.IsNullOrEmpty(fmt.FontFamily)) usedFonts.Add(fmt.FontFamily);
                                foreach (var textEl in run.Elements<W.Text>())
                                {
                                    var txt = textEl.Text;
                                    if (string.IsNullOrEmpty(txt)) continue;
                                    var formatted = footerPara.AddFormattedText(txt);
                                    fmt.ApplyTo(formatted);
                                }
                                foreach (var childEl in run.ChildElements)
                                {
                                    if (childEl is W.Break) footerPara.AddLineBreak();
                                    else if (childEl is W.TabChar) footerPara.AddTab();
                                }
                            }
                            else if (child is W.Hyperlink hyperlink)
                            {
                                foreach (var hRun in hyperlink.Elements<W.Run>())
                                {
                                    var fmt = OpenXmlHelpers.ResolveRunFormatting(word.MainDocumentPart, hRun, fp);
                                    string? txt = null;
                                    foreach (var textEl in hRun.Elements<W.Text>())
                                    {
                                        txt = textEl.Text;
                                        if (!string.IsNullOrEmpty(txt)) break;
                                    }
                                    txt ??= hRun.InnerText;
                                    if (!string.IsNullOrEmpty(txt))
                                    {
                                        var formatted = footerPara.AddFormattedText(txt);
                                        fmt.ApplyTo(formatted);
                                    }
                                }
                            }
                        }

                        // Handle images in footer paragraphs
                        try
                        {
                            var infos = ConverterExtensions.GetImageInfosFromParagraph(word, fp);
                            foreach (var info in infos)
                            {
                                if (info.Bytes == null || info.Bytes.Length == 0) continue;
                                var imgPath = ConverterExtensions.SaveTempImage(info.Bytes);
                                tempFiles.Add(imgPath);
                                try
                                {
                                    if (!System.IO.File.Exists(imgPath)) continue;
                                    var image = footerPara.AddImage(imgPath);
                                    image.LockAspectRatio = true;
                                    if (info.ExtentCxEmu.HasValue)
                                        image.Width = Unit.FromPoint(EmuToPoints(info.ExtentCxEmu.Value));
                                    else
                                        image.Width = Unit.FromCentimeter(6);
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Footer extraction failed: {ex.Message}");
        }

        // numbering counters per numId/level
        var numberingCounters = new Dictionary<string, int[]>();

        int _paraIndex = 0;
        foreach (var element in body.Elements())
        {
            if (element is W.Paragraph p)
            {
                _paraIndex++;
                var pPr = p.ParagraphProperties;

                // Merge paragraph style properties with direct paragraph properties.
                    var styleId = pPr?.ParagraphStyleId?.Val?.Value;
                    var stylePPr = !string.IsNullOrEmpty(styleId) ? OpenXmlHelpers.GetStyleParagraphProperties(word.MainDocumentPart, styleId) : null;
                    var styleFmt = OpenXmlHelpers.GetParagraphFormatting(stylePPr);
                    var paraFmt = OpenXmlHelpers.GetParagraphFormatting(pPr);

                    // Also resolve Normal style defaults for spacing fallback
                    ParagraphFormat? normalFmt = null;
                    if (string.IsNullOrEmpty(styleId) || styleId == "Normal")
                    {
                        var normalPPr = OpenXmlHelpers.GetStyleParagraphProperties(word.MainDocumentPart, "Normal");
                        normalFmt = OpenXmlHelpers.GetParagraphFormatting(normalPPr);
                    }

                    // Choose values: prefer paragraph-level settings when present, otherwise fall back to style
                    var alignment = styleFmt.Alignment;
                    if (pPr?.Justification != null) alignment = paraFmt.Alignment;

                    double leftIndent = styleFmt.LeftIndent;
                    double rightIndent = styleFmt.RightIndent;
                    double firstLine = styleFmt.FirstLineIndent;
                    if (pPr?.Indentation != null)
                    {
                        leftIndent = paraFmt.LeftIndent;
                        firstLine = paraFmt.FirstLineIndent;
                        // Only override right indent if explicitly set in inline pPr
                        if (pPr.Indentation.Right != null)
                            rightIndent = paraFmt.RightIndent;
                    }

                    // Spacing: paragraph explicit values override style; 0 is a valid explicit override
                    double before = paraFmt.HasExplicitSpacingBefore ? paraFmt.SpacingBefore : styleFmt.SpacingBefore;
                    double after = paraFmt.HasExplicitSpacingAfter ? paraFmt.SpacingAfter : styleFmt.SpacingAfter;
                    bool hasExplicitBefore = paraFmt.HasExplicitSpacingBefore || styleFmt.HasExplicitSpacingBefore;
                    bool hasExplicitAfter = paraFmt.HasExplicitSpacingAfter || styleFmt.HasExplicitSpacingAfter;

                    // If no spacing is set at paragraph or style level, use Normal style defaults
                    // Word's built-in default: after=8pt, line=1.08
                    if (!hasExplicitBefore && normalFmt != null && normalFmt.HasExplicitSpacingBefore)
                    {
                        before = normalFmt.SpacingBefore;
                        hasExplicitBefore = true;
                    }
                    if (!hasExplicitAfter && normalFmt != null && normalFmt.HasExplicitSpacingAfter)
                    {
                        after = normalFmt.SpacingAfter;
                        hasExplicitAfter = true;
                    }
                    if (!hasExplicitAfter)
                    {
                        // Fall back to docDefaults pPrDefault spacing, then Word default (160 twips = 8pt)
                        if (docDefaultsFmt != null && docDefaultsFmt.HasExplicitSpacingAfter)
                        {
                            after = docDefaultsFmt.SpacingAfter;
                        }
                        else
                        {
                            after = 8.0;
                        }
                        hasExplicitAfter = true;
                    }
                    if (!hasExplicitBefore && docDefaultsFmt != null && docDefaultsFmt.HasExplicitSpacingBefore)
                    {
                        before = docDefaultsFmt.SpacingBefore;
                        hasExplicitBefore = true;
                    }

                    double? lineSpacing = paraFmt.LineSpacing ?? styleFmt.LineSpacing;
                    string? lineRule = paraFmt.LineRule ?? styleFmt.LineRule;

                    // Fall back to Normal style line spacing
                    if (lineSpacing == null && normalFmt != null)
                    {
                        lineSpacing = normalFmt.LineSpacing;
                        lineRule ??= normalFmt.LineRule;
                    }
                    // Fall back to docDefaults line spacing
                    if (lineSpacing == null && docDefaultsFmt != null)
                    {
                        lineSpacing = docDefaultsFmt.LineSpacing;
                        lineRule ??= docDefaultsFmt.LineRule;
                    }

                // diagnostic preview of text and properties
                try
                {
                    var txtPreview = ConverterExtensions.GetParagraphText(p);
                    // limit preview length to avoid flooding the log
                    if (txtPreview.Length > 40) txtPreview = txtPreview[..37] + "...";
                    OpenXmlHelpers.ImageLoadLogger?.Invoke($"Paragraph {_paraIndex}: styleId={styleId} before={before}pt after={after}pt leftIndent={leftIndent}pt firstLine={firstLine}pt '{txtPreview}'");
                }
                catch { }

                // numbering resolution
                var numPrefix = string.Empty;
                if (pPr?.NumberingProperties != null)
                {
                    var numIdVal = pPr.NumberingProperties.NumberingId?.Val?.Value;
                    var ilvl = (int?)(pPr.NumberingProperties.NumberingLevelReference?.Val?.Value) ?? 0;
                    if (numIdVal.HasValue)
                    {
                        var numId = numIdVal.Value.ToString();
                        if (!numberingCounters.TryGetValue(numId, out var arr))
                        {
                            arr = new int[10];
                            numberingCounters[numId] = arr;
                        }
                        arr[ilvl]++;
                        // reset lower levels
                        for (int i = ilvl + 1; i < arr.Length; i++) arr[i] = 0;

                        var (fmt, text, startAt) = OpenXmlHelpers.GetNumberingLevelFormat(word.MainDocumentPart, numId, ilvl);
                        string label;
                        var n = arr[ilvl] + (startAt.HasValue ? startAt.Value - 1 : 0);
                        switch (fmt)
                        {
                            case "decimal": label = n.ToString() + "."; break;
                            case "lowerLetter": label = ((char)('a' + (n - 1) % 26)).ToString() + ")"; break;
                            case "upperLetter": label = ((char)('A' + (n - 1) % 26)).ToString() + ")"; break;
                            case "lowerRoman": label = ConverterExtensions.ToRoman(n).ToLowerInvariant() + "."; break;
                            case "upperRoman": label = ConverterExtensions.ToRoman(n).ToUpperInvariant() + "."; break;
                            default: label = n.ToString() + "."; break;
                        }
                        if (!string.IsNullOrEmpty(text) && text.Contains("{0}")) label = text.Replace("{0}", n.ToString());
                        numPrefix = label + " ";
                    }
                }

                var infos = ConverterExtensions.GetImageInfosFromParagraph(word, p).ToList();
                if (infos.Count > 0 && string.IsNullOrWhiteSpace(ConverterExtensions.GetParagraphText(p)))
                {
                    foreach (var info in infos)
                    {
                        if (info.Bytes == null || info.Bytes.Length == 0) continue;
                        // Skip very small behindDoc images (< 500 bytes) — likely transparent placeholders
                        if (info.Bytes.Length < 500 && info.IsBackground)
                        {
                            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Skipping tiny behindDoc image ({info.Bytes.Length} bytes) — likely placeholder");
                            continue;
                        }
                        var imgPath = ConverterExtensions.SaveTempImage(info.Bytes);
                        tempFiles.Add(imgPath);
                        try
                        {
                            // For images found in the document body we do NOT promote them to header/background
                            // unless the drawing explicitly marked them as behind the document (info.IsBackground).
                            if (info.IsBackground)
                            {
                                // Positioned anchors with behindDoc in body: render inline at their specified size
                                // since absolute page positioning is not feasible (paragraph Y is unknown)
                                if (info.IsAnchor && info.ExtentCxEmu.HasValue)
                                {
                                    var imgPara = section.AddParagraph();
                                    var image = imgPara.AddImage(imgPath);
                                    image.LockAspectRatio = true;
                                    image.Width = Unit.FromPoint(EmuToPoints(info.ExtentCxEmu.Value));
                                    imgPara.Format.Alignment = ParagraphAlignment.Left;
                                    if (info.OffsetXEmu.HasValue && info.OffsetXEmu.Value > 0)
                                        imgPara.Format.LeftIndent = Unit.FromPoint(EmuToPoints(info.OffsetXEmu.Value));
                                    OpenXmlHelpers.ImageLoadLogger?.Invoke($"Rendered positioned behindDoc image inline: {imgPath} width={EmuToPoints(info.ExtentCxEmu.Value):F1}pt");
                                }
                                else if (!backgroundFiles.Contains(imgPath))
                                {
                                    backgroundFiles.Add(imgPath);
                                    OpenXmlHelpers.ImageLoadLogger?.Invoke($"Collected background image (body): {imgPath}");
                                }
                            }
                            else
                            {
                                var imgPara = section.AddParagraph();
                                var image = imgPara.AddImage(imgPath);
                                image.LockAspectRatio = true;
                                double imgWidthPt;
                                double imgHeightPt = info.ExtentCyEmu.HasValue ? EmuToPoints(info.ExtentCyEmu.Value) : 113.4;
                                if (info.ExtentCxEmu.HasValue)
                                {
                                    imgWidthPt = EmuToPoints(info.ExtentCxEmu.Value);
                                    image.Width = Unit.FromPoint(imgWidthPt);
                                }
                                else
                                {
                                    imgWidthPt = Unit.FromCentimeter(16).Point;
                                    image.Width = Unit.FromCentimeter(16);
                                }

                                // Collect for post-render link annotation
                                if (!string.IsNullOrEmpty(info.HyperlinkUrl))
                                    hyperlinkImages.Add((info.HyperlinkUrl, imgPath, imgWidthPt, imgHeightPt));

                                // If image is anchored, prefer left alignment to match Word anchored placement
                                if (info.IsAnchor)
                                    imgPara.Format.Alignment = ParagraphAlignment.Left;
                                else
                                    imgPara.Format.Alignment = alignment;

                                // if wrapText is explicitly provided, it can override alignment (e.g., right)
                                if (!string.IsNullOrEmpty(info.WrapTextAttribute))
                                {
                                    var wt = info.WrapTextAttribute.ToLowerInvariant();
                                    if (wt.Contains("right")) imgPara.Format.Alignment = ParagraphAlignment.Right;
                                    else if (wt.Contains("bothSides") || wt.Contains("both")) imgPara.Format.Alignment = alignment;
                                }

                                section.AddParagraph();
                            }
                        }
                        catch (Exception ex)
                        {
                            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed processing body image {imgPath}: {ex.Message}");
                        }
                    }
                    continue;
                }

                // Pre-render positioned anchor images as floating shapes
                foreach (var run in p.Elements<W.Run>())
                {
                    var drawing = run.GetFirstChild<W.Drawing>();
                    if (drawing == null) continue;
                    try
                    {
                        var anchorImgInfos = ConverterExtensions.GetImageInfosFromParagraph(word,
                            new W.Paragraph(run.CloneNode(true))).Where(i => i.IsAnchor && i.OffsetXEmu.HasValue).ToList();
                        foreach (var info in anchorImgInfos)
                        {
                            if (info.Bytes == null || info.Bytes.Length == 0) continue;
                            if (info.Bytes.Length < 500 && info.IsBackground) continue;
                            var imgBytes = ConverterExtensions.ApplySrcRectCrop(
                                info.Bytes, info.CropLeft, info.CropTop, info.CropRight, info.CropBottom);
                            var imgPath = ConverterExtensions.SaveTempImage(imgBytes);
                            tempFiles.Add(imgPath);
                            var image = section.AddImage(imgPath);
                            image.LockAspectRatio = true;
                            if (info.ExtentCxEmu.HasValue)
                                image.Width = Unit.FromPoint(EmuToPoints(info.ExtentCxEmu.Value));
                            else
                                image.Width = Unit.FromCentimeter(4);
                            // Position as floating image
                            image.WrapFormat.Style = MigraDoc.DocumentObjectModel.Shapes.WrapStyle.Through;
                            image.RelativeHorizontal = MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Page;
                            image.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Page;
                            image.Left = Unit.FromPoint(EmuToPoints(info.OffsetXEmu.Value));
                            if (info.OffsetYEmu.HasValue)
                                image.Top = Unit.FromPoint(EmuToPoints(info.OffsetYEmu.Value) + section.PageSetup.TopMargin.Point);
                            else
                                image.Top = section.PageSetup.TopMargin;
                        }
                    }
                    catch { }
                }

                var para = section.AddParagraph();
                para.Format.Alignment = alignment;
                if (leftIndent > 0) para.Format.LeftIndent = Unit.FromPoint(leftIndent);
                if (rightIndent > 0) para.Format.RightIndent = Unit.FromPoint(rightIndent);
                if (firstLine != 0) para.Format.FirstLineIndent = Unit.FromPoint(firstLine);
                
                // Determine effective font size for paragraph (max run size or style/default)
                double effectiveFontSize = 0;
                try
                {
                    foreach (var run in p.Elements<W.Run>())
                    {
                        var rf = OpenXmlHelpers.ResolveRunFormatting(word.MainDocumentPart, run, p);
                        if (rf.Size.HasValue && rf.Size.Value > effectiveFontSize) effectiveFontSize = rf.Size.Value;
                    }
                    if (effectiveFontSize == 0)
                    {
                        var pRunPr = pPr?.GetFirstChild<W.RunProperties>();
                        var szVal = pRunPr?.FontSize?.Val?.Value;
                        if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var pHalf)) 
                            effectiveFontSize = pHalf / 2.0;
                    }
                    if (effectiveFontSize == 0 && !string.IsNullOrEmpty(styleId))
                    {
                        var sr = OpenXmlHelpers.GetStyleRunProperties(word.MainDocumentPart, styleId);
                        var szVal = sr?.FontSize?.Val?.Value;
                        if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var sHalf)) 
                            effectiveFontSize = sHalf / 2.0;
                    }
                    if (effectiveFontSize == 0)
                    {
                        var nsr = OpenXmlHelpers.GetStyleRunProperties(word.MainDocumentPart, "Normal");
                        var szVal = nsr?.FontSize?.Val?.Value;
                        if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var nHalf)) 
                            effectiveFontSize = nHalf / 2.0;
                    }
                    if (effectiveFontSize == 0) effectiveFontSize = 11.0; // fallback
                }
                catch { effectiveFontSize = 11.0; }

                // Line spacing
                try
                {
                    if (lineSpacing.HasValue && lineSpacing.Value > 0)
                    {
                        var rule = lineRule?.ToLowerInvariant() ?? "auto";
                        if (rule == "auto")
                        {
                            // Word "auto" = multiple of line height (e.g. 259/240 = 1.079×)
                            para.Format.LineSpacing = lineSpacing.Value;
                            para.Format.LineSpacingRule = LineSpacingRule.Multiple;
                        }
                        else if (rule == "exact")
                        {
                            para.Format.LineSpacing = Unit.FromPoint(lineSpacing.Value);
                            para.Format.LineSpacingRule = LineSpacingRule.Exactly;
                        }
                        else // atLeast
                        {
                            para.Format.LineSpacing = Unit.FromPoint(lineSpacing.Value);
                            para.Format.LineSpacingRule = LineSpacingRule.AtLeast;
                        }
                    }
                }
                catch { }

                // Apply SpaceBefore and SpaceAfter
                if (hasExplicitBefore)
                    para.Format.SpaceBefore = Unit.FromPoint(before);

                if (hasExplicitAfter)
                    para.Format.SpaceAfter = Unit.FromPoint(after);

                // Add tab stops from paragraph properties, or default tab stops
                try
                {
                    var tabs = pPr?.GetFirstChild<W.Tabs>();
                    if (tabs != null)
                    {
                        foreach (var tab in tabs.Elements<W.TabStop>())
                        {
                            if (tab.Position?.HasValue == true)
                            {
                                double tabPos = tab.Position.Value / 20.0; // twips to points
                                var tabAlign = TabAlignment.Left;
                                if (tab.Val?.Value == W.TabStopValues.Center) tabAlign = TabAlignment.Center;
                                else if (tab.Val?.Value == W.TabStopValues.Right) tabAlign = TabAlignment.Right;
                                para.Format.TabStops.AddTabStop(Unit.FromPoint(tabPos), tabAlign);
                            }
                        }
                    }
                    else
                    {
                        // Add default tab stops (Word default: every 36pt / 720 twips / 0.5 inch)
                        double pageContentWidth = section.PageSetup.PageWidth.Point - 
                            section.PageSetup.LeftMargin.Point - section.PageSetup.RightMargin.Point;
                        for (double ts = 36.0; ts < pageContentWidth; ts += 36.0)
                            para.Format.TabStops.AddTabStop(Unit.FromPoint(ts));
                    }
                }
                catch { }

                if (!string.IsNullOrEmpty(numPrefix)) para.AddText(numPrefix);

                var addedText = false;

                // Collect VML picts from this paragraph to render together as side-by-side layout
                var vmlPicts = new List<W.Picture>();
                
                // Process paragraph children (runs and hyperlinks)
                foreach (var child in p.ChildElements)
                {
                    if (child is W.Run run)
                    {
                        // Check if run contains w:pict with VML textbox content
                        var pict = run.GetFirstChild<W.Picture>();
                        if (pict != null)
                        {
                            vmlPicts.Add(pict);
                            continue;
                        }

                        // Check if run contains an inline drawing (image) — render it in the paragraph
                        var drawing = run.GetFirstChild<W.Drawing>();
                        if (drawing != null)
                        {
                            try
                            {
                                var runInfos = ConverterExtensions.GetImageInfosFromParagraph(word, 
                                    new W.Paragraph(run.CloneNode(true))).ToList();
                                foreach (var info in runInfos)
                                {
                                    // Skip anchor images — already pre-rendered above
                                    if (info.IsAnchor && info.OffsetXEmu.HasValue) continue;
                                    if (info.Bytes == null || info.Bytes.Length == 0) continue;
                                    if (info.Bytes.Length < 500 && info.IsBackground) continue;
                                    var imgBytes = ConverterExtensions.ApplySrcRectCrop(
                                        info.Bytes, info.CropLeft, info.CropTop, info.CropRight, info.CropBottom);
                                    var imgPath = ConverterExtensions.SaveTempImage(imgBytes);
                                    tempFiles.Add(imgPath);
                                    var image = para.AddImage(imgPath);
                                    image.LockAspectRatio = true;
                                    double imgWidthPt = info.ExtentCxEmu.HasValue ? EmuToPoints(info.ExtentCxEmu.Value) : 113.4;
                                    double imgHeightPt = info.ExtentCyEmu.HasValue ? EmuToPoints(info.ExtentCyEmu.Value) : 113.4;
                                    image.Width = Unit.FromPoint(imgWidthPt);
                                    addedText = true;
                                    // Collect for post-render link annotation
                                    if (!string.IsNullOrEmpty(info.HyperlinkUrl))
                                        hyperlinkImages.Add((info.HyperlinkUrl, imgPath, imgWidthPt, imgHeightPt));
                                }
                            }
                            catch { }
                        }

                        addedText |= ProcessRun(word.MainDocumentPart, run, p, para, usedFonts);
                    }
                    else if (child is W.Hyperlink hyperlink)
                    {
                        addedText |= ProcessHyperlink(word, hyperlink, p, para, usedFonts, tempFiles, hyperlinkImages);
                    }
                }

                // Render collected VML textbox content
                if (vmlPicts.Count > 0)
                {
                    RenderVmlPicts(word, vmlPicts, section, usedFonts, tempFiles);
                }

                // If paragraph contained no text, add a single space to preserve the paragraph
                if (!addedText)
                {
                    try
                    {
                        double? defSize = null;
                        // reuse styleId computed earlier when merging style/paragraph properties
                        if (!string.IsNullOrEmpty(styleId))
                        {
                            var sr = OpenXmlHelpers.GetStyleRunProperties(word.MainDocumentPart, styleId);
                            var szVal = sr?.FontSize?.Val?.Value;
                            if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var halfPts)) 
                                defSize = halfPts / 2.0;
                        }
                        // fallback: try paragraph-level run properties (pPr rPr)
                        if (defSize == null)
                        {
                            try
                            {
                                var pRunPr = pPr?.GetFirstChild<W.RunProperties>();
                                var szVal = pRunPr?.FontSize?.Val?.Value;
                                if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var pHalf)) 
                                    defSize = pHalf / 2.0;
                            }
                            catch { }
                        }
                        // fallback: try Normal style
                        if (defSize == null)
                        {
                            try
                            {
                                var nsr = OpenXmlHelpers.GetStyleRunProperties(word.MainDocumentPart, "Normal");
                                var szVal = nsr?.FontSize?.Val?.Value;
                                if (!string.IsNullOrEmpty(szVal) && double.TryParse(szVal, out var h)) 
                                    defSize = h / 2.0;
                            }
                            catch { }
                        }

                        // Use the determined font size or fallback to effective size
                        if (defSize == null) defSize = effectiveFontSize;

                        var f = para.AddFormattedText(" ");
                        if (defSize.HasValue) f.Size = defSize.Value;
                    }
                    catch { }
                }
            }
            else if (element is W.Table t)
            {
                WordTableRenderer.RenderTable(word, section, t, tempFiles);
            }
        }

        // Apply Normal style font from used fonts or default to Arial
        try
        {
            var normal = doc.Styles["Normal"];
            if (normal != null)
                normal.Font.Name = usedFonts.FirstOrDefault() ?? "Arial";
        }
        catch { }

        var renderer = new PdfDocumentRenderer()
        {
            Document = doc
        };
        renderer.RenderDocument();

        // Draw collected background images directly onto each PDF page so they appear behind content
        try
        {
            if (backgroundFiles.Count > 0)
            {
                var pdf = renderer.PdfDocument;
                OpenXmlHelpers.ImageLoadLogger?.Invoke($"Drawing {backgroundFiles.Count} background files onto PDF pages");
                OpenXmlHelpers.ImageLoadLogger?.Invoke($"Final HeaderDistance={section.PageSetup.HeaderDistance.Point}pt FooterDistance={section.PageSetup.FooterDistance.Point}pt TopMargin={section.PageSetup.TopMargin.Point} BottomMargin={section.PageSetup.BottomMargin.Point}");
                foreach (var page in pdf.Pages)
                {
                    using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend);
                    foreach (var b in backgroundFiles)
                    {
                        try
                        {
                            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Drawing background: {b} on page size {page.Width}x{page.Height}");
                            using var ximg = XImage.FromFile(b);
                            gfx.DrawImage(ximg, 0, 0, page.Width.Point, page.Height.Point);
                        }
                        catch (Exception ex)
                        {
                            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Failed drawing background {b}: {ex.Message}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Background drawing failed: {ex.Message}");
        }

        // Add clickable link annotations for hyperlinked images
        try
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Hyperlinked images collected: {hyperlinkImages.Count}");
            if (hyperlinkImages.Count > 0)
            {
                var pdf = renderer.PdfDocument;
                foreach (var page in pdf.Pages)
                {
                    AddImageHyperlinkAnnotations(page, hyperlinkImages);
                }
            }
        }
        catch (Exception ex)
        {
            OpenXmlHelpers.ImageLoadLogger?.Invoke($"Hyperlink annotation failed: {ex.Message}");
        }

        return renderer;
    }

    /// <summary>
    /// Scans a PDF page's content stream for image placements and adds web link annotations
    /// for images that match collected hyperlinked image info.
    /// </summary>
    private static void AddImageHyperlinkAnnotations(PdfSharp.Pdf.PdfPage page,
        List<(string Url, string ImagePath, double WidthPt, double HeightPt)> hyperlinkImages)
    {
        try
        {
            var content = PdfSharp.Pdf.Content.ContentReader.ReadContent(page);
            var usedImages = new HashSet<int>();
            // Walk the content stream looking for image placements (cm + Do pattern)
            // The cm operator sets CTM: a b c d e f where the image is drawn at (e,f) with size (a,d)
            double lastA = 0, lastD = 0, lastE = 0, lastF = 0;
            bool hasCm = false;

            WalkContentForImages(content, page, hyperlinkImages, usedImages, ref lastA, ref lastD, ref lastE, ref lastF, ref hasCm);
        }
        catch { }
    }

    private static void WalkContentForImages(PdfSharp.Pdf.Content.Objects.CSequence seq,
        PdfSharp.Pdf.PdfPage page,
        List<(string Url, string ImagePath, double WidthPt, double HeightPt)> hyperlinkImages,
        HashSet<int> usedImages,
        ref double lastA, ref double lastD, ref double lastE, ref double lastF, ref bool hasCm)
    {
        foreach (var item in seq)
        {
            if (item is PdfSharp.Pdf.Content.Objects.CSequence sub)
            {
                WalkContentForImages(sub, page, hyperlinkImages, usedImages,
                    ref lastA, ref lastD, ref lastE, ref lastF, ref hasCm);
                continue;
            }
            if (item is not PdfSharp.Pdf.Content.Objects.COperator op) continue;

            if (op.OpCode.Name == "cm" && op.Operands.Count >= 6)
            {
                lastA = OpVal(op.Operands[0]);
                lastD = OpVal(op.Operands[3]);
                lastE = OpVal(op.Operands[4]);
                lastF = OpVal(op.Operands[5]);
                hasCm = true;
            }
            else if (op.OpCode.Name == "Do" && hasCm)
            {
                // Image was placed — match by size against known hyperlinked images
                for (int i = 0; i < hyperlinkImages.Count; i++)
                {
                    if (usedImages.Contains(i)) continue;
                    var hi = hyperlinkImages[i];
                    // Match by width (within 2pt tolerance)
                    if (Math.Abs(lastA - hi.WidthPt) < 2.0)
                    {
                        // PDF coordinates: (lastE, lastF) is bottom-left, size is (lastA x lastD)
                        double x = lastE;
                        double y = lastF;
                        double w = lastA;
                        double h = Math.Abs(lastD);
                        // PdfRectangle uses PDF coordinates (origin at bottom-left)
                        var rect = new PdfSharp.Pdf.PdfRectangle(
                            new PdfSharp.Drawing.XPoint(x, y),
                            new PdfSharp.Drawing.XPoint(x + w, y + h));
                        page.AddWebLink(rect, hi.Url);
                        usedImages.Add(i);
                        break;
                    }
                }
                hasCm = false;
            }
        }
    }

    private static double OpVal(PdfSharp.Pdf.Content.Objects.CObject o) => o switch
    {
        PdfSharp.Pdf.Content.Objects.CReal r => r.Value,
        PdfSharp.Pdf.Content.Objects.CInteger ci => ci.Value,
        _ => 0
    };

    /// <summary>
    /// Process a run element and add its content to the paragraph
    /// </summary>
    private static bool ProcessRun(MainDocumentPart mainPart, W.Run run, W.Paragraph p, 
        MigraDoc.DocumentObjectModel.Paragraph para, HashSet<string> usedFonts)
    {
        var fmt = OpenXmlHelpers.ResolveRunFormatting(mainPart, run, p);
        if (!string.IsNullOrEmpty(fmt.FontFamily)) usedFonts.Add(fmt.FontFamily);

        bool addedText = false;

        // Process child elements in document order to preserve tab/text interleaving
        foreach (var child in run.ChildElements)
        {
            if (child is W.Text textEl)
            {
                var txt = textEl.Text;
                if (string.IsNullOrEmpty(txt)) continue;
                if (ConverterExtensions.ContainsEmoji(txt))
                {
                    foreach (var (seg, isEmoji) in ConverterExtensions.SplitEmojiSegments(txt))
                    {
                        var formatted = para.AddFormattedText(seg);
                        fmt.ApplyTo(formatted);
                        if (isEmoji) formatted.Font.Name = "Noto Emoji";
                    }
                }
                else
                {
                    var formatted = para.AddFormattedText(txt);
                    fmt.ApplyTo(formatted);
                }
                addedText = true;
            }
            else if (child is W.Break)
            {
                try { para.AddLineBreak(); } catch { }
            }
            else if (child is W.TabChar)
            {
                try { para.AddTab(); } catch { }
            }
        }

        return addedText;
    }

    /// <summary>
    /// Process a hyperlink element and add it to the paragraph
    /// </summary>
    private static bool ProcessHyperlink(WordprocessingDocument doc, W.Hyperlink hyperlink, 
        W.Paragraph p, MigraDoc.DocumentObjectModel.Paragraph para, HashSet<string> usedFonts,
        List<string> tempFiles, List<(string Url, string ImagePath, double WidthPt, double HeightPt)> hyperlinkImages)
    {
        var mainPart = doc.MainDocumentPart!;
        string? url = null;

        var relId = hyperlink.Id?.Value;
        if (!string.IsNullOrEmpty(relId))
        {
            try
            {
                var rel = mainPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
                if (rel != null)
                    url = rel.Uri?.ToString();
            }
            catch { }
        }

        if (string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(hyperlink.Anchor?.Value))
            url = "#" + hyperlink.Anchor.Value;

        bool addedText = false;

        foreach (var run in hyperlink.Elements<W.Run>())
        {
            // Handle images inside hyperlinks
            var drawing = run.GetFirstChild<W.Drawing>();
            if (drawing != null)
            {
                try
                {
                    var runInfos = ConverterExtensions.GetImageInfosFromParagraph(doc,
                        new W.Paragraph(run.CloneNode(true))).ToList();
                    foreach (var info in runInfos)
                    {
                        if (info.Bytes == null || info.Bytes.Length == 0) continue;
                        var imgBytes = ConverterExtensions.ApplySrcRectCrop(
                            info.Bytes, info.CropLeft, info.CropTop, info.CropRight, info.CropBottom);
                        var imgPath = ConverterExtensions.SaveTempImage(imgBytes);
                        tempFiles.Add(imgPath);
                        double widthPt = info.ExtentCxEmu.HasValue ? info.ExtentCxEmu.Value / 12700.0 : 113.4;
                        double heightPt = info.ExtentCyEmu.HasValue ? info.ExtentCyEmu.Value / 12700.0 : 113.4;
                        var image = para.AddImage(imgPath);
                        image.LockAspectRatio = true;
                        image.Width = Unit.FromPoint(widthPt);
                        addedText = true;
                        // Collect for post-render link annotation
                        if (!string.IsNullOrEmpty(url))
                            hyperlinkImages.Add((url, imgPath, widthPt, heightPt));
                    }
                }
                catch { }
                continue;
            }

            var fmt = OpenXmlHelpers.ResolveRunFormatting(mainPart, run, p);
            if (!string.IsNullOrEmpty(fmt.FontFamily)) usedFonts.Add(fmt.FontFamily);

            string? txt = null;
            foreach (var textEl in run.Elements<W.Text>())
            {
                txt = textEl.Text;
                if (!string.IsNullOrEmpty(txt)) break;
            }
            txt ??= run.InnerText;

            if (!string.IsNullOrEmpty(txt))
            {
                if (!string.IsNullOrEmpty(url))
                {
                    try
                    {
                        var link = para.AddHyperlink(url, HyperlinkType.Web);
                        var formatted = link.AddFormattedText(txt);
                        addedText = true;
                        fmt.ApplyTo(formatted);
                        if (string.IsNullOrEmpty(fmt.Color))
                            formatted.Color = Colors.Blue;
                        formatted.Underline = Underline.Single;
                    }
                    catch
                    {
                        var formatted = para.AddFormattedText(txt);
                        addedText = true;
                        formatted.Color = Colors.Blue;
                        formatted.Underline = Underline.Single;
                    }
                }
                else
                {
                    var formatted = para.AddFormattedText(txt);
                    addedText = true;
                    fmt.ApplyTo(formatted);
                }
            }

            foreach (var child in run.ChildElements)
            {
                if (child is W.Break)
                {
                    try { para.AddLineBreak(); } catch { }
                }
            }
        }

        return addedText;
    }

    /// <summary>
    /// Renders VML pict elements. When multiple picts exist, renders side-by-side in a table layout.
    /// </summary>
    private static void RenderVmlPicts(WordprocessingDocument doc, List<W.Picture> picts, 
        Section section, HashSet<string> usedFonts, List<string> tempFiles)
    {
        if (picts.Count == 0) return;

        if (picts.Count == 1)
        {
            RenderVmlTextBoxContent(doc, pict: picts[0], section, usedFonts, tempFiles);
            return;
        }

        // Multiple VML picts — parse widths from style attributes
        var pictInfos = new List<(W.Picture pict, double marginLeft, double width)>();
        foreach (var pict in picts)
        {
            double marginLeft = 0, width = 200;
            var topShape = pict.ChildElements.FirstOrDefault(c => c.LocalName == "group" || c.LocalName == "shape");
            if (topShape != null)
            {
                var styleVal = topShape.GetAttributes().Where(a => a.LocalName == "style").Select(a => a.Value).FirstOrDefault();
                if (!string.IsNullOrEmpty(styleVal))
                {
                    foreach (var part in styleVal.Split(';'))
                    {
                        var kv = part.Split(':');
                        if (kv.Length == 2)
                        {
                            var key = kv[0].Trim();
                            var val = kv[1].Trim().Replace("pt", "");
                            if (key == "margin-left" && double.TryParse(val, System.Globalization.NumberStyles.Any, 
                                System.Globalization.CultureInfo.InvariantCulture, out var ml))
                                marginLeft = ml;
                            else if (key == "width" && double.TryParse(val, System.Globalization.NumberStyles.Any, 
                                System.Globalization.CultureInfo.InvariantCulture, out var w))
                                width = w;
                        }
                    }
                }
            }
            pictInfos.Add((pict, marginLeft, width));
        }

        pictInfos.Sort((a, b) => a.marginLeft.CompareTo(b.marginLeft));

        // Render as a layout table with one column per pict
        try
        {
            double pageWidth = section.PageSetup.PageWidth.Point - 
                section.PageSetup.LeftMargin.Point - section.PageSetup.RightMargin.Point;

            var layoutTbl = section.AddTable();
            layoutTbl.Borders.Visible = false;

            foreach (var info in pictInfos)
                layoutTbl.AddColumn(Unit.FromPoint(Math.Min(info.width, pageWidth / pictInfos.Count)));

            var layoutRow = layoutTbl.AddRow();
            for (int i = 0; i < pictInfos.Count; i++)
            {
                var cell = layoutRow[i];
                cell.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Top;
                RenderVmlTextBoxIntoCell(doc, pictInfos[i].pict, cell, usedFonts);
            }
        }
        catch
        {
            foreach (var pict in picts)
                RenderVmlTextBoxContent(doc, pict, section, usedFonts, tempFiles);
        }
    }

    /// <summary>
    /// Renders VML textbox content into a table cell, including nested tables with proper formatting.
    /// </summary>
    private static void RenderVmlTextBoxIntoCell(WordprocessingDocument doc, W.Picture pict, 
        MigraDoc.DocumentObjectModel.Tables.Cell cell, HashSet<string> usedFonts)
    {
        try
        {
            var mainPart = doc.MainDocumentPart!;
            var txbxContents = pict.Descendants()
                .Where(d => d.LocalName == "txbxContent")
                .ToList();

            foreach (var txbx in txbxContents)
            {
                string? fillColor = null;
                var parentShape = txbx.Ancestors().FirstOrDefault(a => a.LocalName == "shape");
                if (parentShape != null)
                {
                    var fillAttr = parentShape.GetAttributes()
                        .Where(a => a.LocalName == "fillcolor").Select(a => a.Value).FirstOrDefault();
                    if (!string.IsNullOrEmpty(fillAttr))
                        fillColor = fillAttr.TrimStart('#');
                }

                foreach (var child in txbx.ChildElements)
                {
                    if (child is W.Paragraph wp)
                    {
                        var para = cell.AddParagraph();
                        para.Format.SpaceBefore = Unit.FromPoint(0);
                        para.Format.SpaceAfter = Unit.FromPoint(0);

                        var pPr = wp.ParagraphProperties;
                        var pFmt = WordHelpers.GetParagraphFormatting(pPr);
                        para.Format.Alignment = pFmt.Alignment;

                        if (!string.IsNullOrEmpty(fillColor))
                        {
                            try { para.Format.Shading.Color = Color.Parse("#" + fillColor); } catch { }
                        }

                        bool hasContent = false;
                        foreach (var run in wp.Elements<W.Run>())
                            hasContent |= ProcessRun(mainPart, run, wp, para, usedFonts);

                        if (!hasContent)
                            para.AddText(" ");
                    }
                    else if (child is W.Table innerTbl)
                    {
                        // Render the inner table as a nested table within the cell
                        RenderInnerTableInCell(mainPart, innerTbl, cell, usedFonts);
                    }
                }
            }
        }
        catch { }
    }

    /// <summary>
    /// Renders an inner Word table within a MigraDoc cell using proper table formatting.
    /// </summary>
    private static void RenderInnerTableInCell(MainDocumentPart mainPart, W.Table innerTbl,
        MigraDoc.DocumentObjectModel.Tables.Cell parentCell, HashSet<string> usedFonts)
    {
        try
        {
            var gridCols = innerTbl.GetFirstChild<W.TableGrid>()?.Elements<W.GridColumn>().ToList();
            if (gridCols == null || gridCols.Count == 0) return;

            var tblPr = innerTbl.GetFirstChild<W.TableProperties>();

            var mTbl = new MigraDoc.DocumentObjectModel.Tables.Table();

            foreach (var gc in gridCols)
            {
                var wVal = gc.Width?.Value;
                double colW = 80;
                if (!string.IsNullOrEmpty(wVal) && double.TryParse(wVal, out var tw)) colW = tw / 20.0;
                mTbl.AddColumn(Unit.FromPoint(colW));
            }

            // Apply table borders from tblPr/tblBorders
            var tblBorders = tblPr?.GetFirstChild<W.TableBorders>();
            if (tblBorders != null)
            {
                void ApplyBorder(MigraDoc.DocumentObjectModel.Border mBorder, W.BorderType? wBorder)
                {
                    if (wBorder == null) return;
                    var sz = wBorder.Size?.Value;
                    if (sz != null && sz > 0)
                        mBorder.Width = Unit.FromPoint(sz.Value / 8.0);
                    var col = wBorder.Color?.Value;
                    if (!string.IsNullOrEmpty(col) && col != "auto")
                    {
                        try { mBorder.Color = Color.Parse("#" + col); } catch { }
                    }
                }
                ApplyBorder(mTbl.Borders.Top, tblBorders.TopBorder);
                ApplyBorder(mTbl.Borders.Bottom, tblBorders.BottomBorder);
                ApplyBorder(mTbl.Borders.Left, tblBorders.LeftBorder);
                ApplyBorder(mTbl.Borders.Right, tblBorders.RightBorder);
                // insideH/insideV → apply to all cell borders
                var insideH = tblBorders.InsideHorizontalBorder;
                var insideV = tblBorders.InsideVerticalBorder;
                if (insideH != null)
                {
                    ApplyBorder(mTbl.Borders.Top, insideH);
                    ApplyBorder(mTbl.Borders.Bottom, insideH);
                }
                if (insideV != null)
                {
                    ApplyBorder(mTbl.Borders.Left, insideV);
                    ApplyBorder(mTbl.Borders.Right, insideV);
                }
            }

            int cols = gridCols.Count;
            foreach (var wRow in innerTbl.Elements<W.TableRow>())
            {
                var mRow = mTbl.AddRow();
                var cells = wRow.Elements<W.TableCell>().ToList();
                for (int c = 0; c < Math.Min(cells.Count, cols); c++)
                {
                    var wCell = cells[c];
                    var mCell = mRow[c];
                    var tcPr = wCell.GetFirstChild<W.TableCellProperties>();

                    // Cell shading
                    var shading = tcPr?.GetFirstChild<W.Shading>()?.Fill?.Value;
                    if (!string.IsNullOrEmpty(shading) && !string.Equals(shading, "auto", StringComparison.OrdinalIgnoreCase))
                    {
                        try { mCell.Shading.Color = Color.Parse("#" + shading); } catch { }
                    }

                    foreach (var cp in wCell.Elements<W.Paragraph>())
                    {
                        var para = mCell.AddParagraph();
                        para.Format.SpaceBefore = Unit.FromPoint(0);
                        para.Format.SpaceAfter = Unit.FromPoint(0);
                        foreach (var run in cp.Elements<W.Run>())
                            ProcessRun(mainPart, run, cp, para, usedFonts);
                    }
                }
            }

            parentCell.Elements.Add(mTbl);
        }
        catch { }
    }

    /// <summary>
    /// Extracts and renders text from VML text boxes (w:pict → v:shape → v:textbox → w:txbxContent).
    /// Each w:txbxContent paragraph is rendered as a paragraph in the section.
    /// </summary>
    private static void RenderVmlTextBoxContent(WordprocessingDocument doc, W.Picture pict, 
        Section section, HashSet<string> usedFonts, List<string> tempFiles)
    {
        try
        {
            var mainPart = doc.MainDocumentPart!;
            // Find all txbxContent elements (may be inside v:shape or v:group → v:shape)
            var txbxContents = pict.Descendants()
                .Where(d => d.LocalName == "txbxContent")
                .ToList();

            foreach (var txbx in txbxContents)
            {
                // Check parent shape for fill color (header box styling)
                string? fillColor = null;
                string? strokeColor = null;
                var parentShape = txbx.Ancestors().FirstOrDefault(a => a.LocalName == "shape");
                if (parentShape != null)
                {
                    var fillAttr = parentShape.GetAttributes()
                        .FirstOrDefault(a => a.LocalName == "fillcolor");
                    if (fillAttr != null && !string.IsNullOrEmpty(fillAttr.Value))
                        fillColor = fillAttr.Value.TrimStart('#');
                    var strokeAttr = parentShape.GetAttributes()
                        .FirstOrDefault(a => a.LocalName == "strokecolor");
                    if (strokeAttr != null && !string.IsNullOrEmpty(strokeAttr.Value))
                        strokeColor = strokeAttr.Value.TrimStart('#');
                }

                // Process children in document order (paragraphs and tables)
                foreach (var child in txbx.ChildElements)
                {
                    if (child is W.Paragraph wp)
                    {
                        var para = section.AddParagraph();
                        var pPr = wp.ParagraphProperties;
                        var pFmt = WordHelpers.GetParagraphFormatting(pPr);
                        para.Format.Alignment = pFmt.Alignment;

                        // VML textboxes are positioned absolutely in Word; use compact spacing
                        var spacing = pPr?.SpacingBetweenLines;
                        if (spacing?.Before?.Value != null && int.TryParse(spacing.Before.Value, out var bTwips))
                            para.Format.SpaceBefore = Unit.FromPoint(bTwips / 20.0);
                        else
                            para.Format.SpaceBefore = Unit.FromPoint(0);

                        if (spacing?.After?.Value != null && int.TryParse(spacing.After.Value, out var aTwips))
                            para.Format.SpaceAfter = Unit.FromPoint(aTwips / 20.0);
                        else
                            para.Format.SpaceAfter = Unit.FromPoint(0);

                        if (spacing?.Line?.Value != null && int.TryParse(spacing.Line.Value, out var lineVal))
                        {
                            var rule = spacing.LineRule?.Value;
                            if (rule == W.LineSpacingRuleValues.Exact)
                            {
                                para.Format.LineSpacingRule = LineSpacingRule.Exactly;
                                para.Format.LineSpacing = Unit.FromPoint(lineVal / 20.0);
                            }
                            else
                            {
                                para.Format.LineSpacingRule = LineSpacingRule.Multiple;
                                para.Format.LineSpacing = Unit.FromPoint(lineVal / 240.0 * 12);
                            }
                        }

                        // Apply fill color as paragraph shading if present
                        if (!string.IsNullOrEmpty(fillColor))
                        {
                            try { para.Format.Shading.Color = Color.Parse("#" + fillColor); } catch { }
                        }

                        bool hasContent = false;
                        foreach (var run in wp.Elements<W.Run>())
                        {
                            hasContent |= ProcessRun(mainPart, run, wp, para, usedFonts);
                        }

                        if (!hasContent)
                            para.AddText(" ");
                    }
                    else if (child is W.Table tbl)
                    {
                        WordTableRenderer.RenderTable(doc, section, tbl, tempFiles);
                    }
                }
            }
        }
        catch { }
    }
}