using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter
{
    internal static class ConverterExtensions
    {
        internal record ImageInfo(byte[] Bytes, long? ExtentCxEmu, long? ExtentCyEmu, string? WrapTextAttribute, bool IsBackground, bool IsAnchor, long? OffsetXEmu = null, long? OffsetYEmu = null,
            int CropLeft = 0, int CropTop = 0, int CropRight = 0, int CropBottom = 0, string? HyperlinkUrl = null);

        internal static string ToRoman(int number)
        {
            if (number < 1) return string.Empty;
            var map = new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") };
            var result = new System.Text.StringBuilder();
            foreach (var (val, sym) in map)
            {
                while (number >= val)
                {
                    result.Append(sym);
                    number -= val;
                }
            }
            return result.ToString();
        }

        internal static string GetParagraphText(W.Paragraph p)
        {
            var sb = new System.Text.StringBuilder();
            if (p == null) return string.Empty;
            foreach (var node in p.Descendants())
            {
                if (node is W.Text t)
                    sb.Append(t.Text);
                else if (node is W.TabChar)
                    sb.Append('\t');
                else if (node is W.Break)
                    sb.Append('\n');
            }
            return sb.ToString();
        }

        /// <summary>
        /// Checks if a string contains any emoji or supplementary Unicode characters (above U+FFFF).
        /// </summary>
        internal static bool ContainsEmoji(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsHighSurrogate(text[i])) return true;
                // Common emoji in BMP: miscellaneous symbols, dingbats, emoticons
                int cp = text[i];
                if (cp >= 0x2600 && cp <= 0x27BF) return true; // Misc symbols & dingbats
                if (cp >= 0x2700 && cp <= 0x27BF) return true;
                if (cp >= 0xFE00 && cp <= 0xFE0F) return true; // Variation selectors
                if (cp >= 0x200D && cp <= 0x200D) return true; // ZWJ
            }
            return false;
        }

        /// <summary>
        /// Splits text into segments of (text, isEmoji) pairs for rendering with appropriate fonts.
        /// </summary>
        internal static List<(string Text, bool IsEmoji)> SplitEmojiSegments(string text)
        {
            var segments = new List<(string, bool)>();
            if (string.IsNullOrEmpty(text)) return segments;

            var current = new System.Text.StringBuilder();
            bool currentIsEmoji = false;

            for (int i = 0; i < text.Length; i++)
            {
                bool charIsEmoji = false;
                if (char.IsHighSurrogate(text[i]))
                {
                    charIsEmoji = true;
                }
                else
                {
                    int cp = text[i];
                    if (cp >= 0x2600 && cp <= 0x27BF) charIsEmoji = true;
                    if (cp >= 0xFE00 && cp <= 0xFE0F) charIsEmoji = true;
                }

                if (charIsEmoji != currentIsEmoji && current.Length > 0)
                {
                    segments.Add((current.ToString(), currentIsEmoji));
                    current.Clear();
                }
                currentIsEmoji = charIsEmoji;

                current.Append(text[i]);
                // Include low surrogate with high surrogate
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    current.Append(text[++i]);
                }
            }

            if (current.Length > 0)
                segments.Add((current.ToString(), currentIsEmoji));

            return segments;
        }

        internal static IEnumerable<byte[]> GetImagesFromParagraph(WordprocessingDocument doc, W.Paragraph p)
        {
            // Backwards-compatible: return only bytes from new API
            return GetImageInfosFromParagraph(doc, p).Select(i => i.Bytes);
        }

        internal static IEnumerable<ImageInfo> GetImageInfosFromParagraph(WordprocessingDocument doc, W.Paragraph p)
        {
            var results = new List<ImageInfo>();
            if (p == null) return results;

            // DrawingML blips (modern WordprocessingML)
            var blipElements = p.Descendants<DocumentFormat.OpenXml.Drawing.Blip>();
            foreach (var blip in blipElements)
            {
                var rId = blip.Embed?.Value ?? blip.Link?.Value;
                if (string.IsNullOrEmpty(rId)) continue;

                // Try to find extent and wrap from ancestor wp:extent or wp:anchor/wp:wrapSquare
                long? cx = null, cy = null;
                string? wrapText = null;
                bool isBackground = false;
                bool isAnchor = false;
                long? offsetX = null, offsetY = null;
                var wpAncestor = blip.Ancestors().FirstOrDefault(a => a.LocalName == "anchor" || a.LocalName == "inline");
                if (wpAncestor != null)
                {
                    isAnchor = wpAncestor.LocalName == "anchor";
                    try
                    {
                        var extent = wpAncestor.Descendants().FirstOrDefault(d => d.LocalName == "extent");
                        if (extent != null)
                        {
                            // use safe attribute lookup to avoid schema exceptions
                            var cxAttr = extent.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, "cx", StringComparison.OrdinalIgnoreCase));
                            var cyAttr = extent.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, "cy", StringComparison.OrdinalIgnoreCase));
                            if (cxAttr != null && !string.IsNullOrEmpty(cxAttr.Value) && long.TryParse(cxAttr.Value, out var tmpCx)) cx = tmpCx;
                            if (cyAttr != null && !string.IsNullOrEmpty(cyAttr.Value) && long.TryParse(cyAttr.Value, out var tmpCy)) cy = tmpCy;
                        }

                        var wrap = wpAncestor.Descendants().FirstOrDefault(d => d.LocalName == "wrapSquare" || d.LocalName == "wrapTopAndBottom" || d.LocalName == "wrapNone");
                        if (wrap != null)
                        {
                            var wtAttr = wrap.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, "wrapText", StringComparison.OrdinalIgnoreCase));
                            if (wtAttr != null && !string.IsNullOrEmpty(wtAttr.Value)) wrapText = wtAttr.Value;
                        }

                        // detect behindDoc attribute on anchor (background image)
                        var behindAttr = wpAncestor.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, "behindDoc", StringComparison.OrdinalIgnoreCase));
                        if (behindAttr != null && !string.IsNullOrEmpty(behindAttr.Value))
                        {
                            if (behindAttr.Value == "1" || string.Equals(behindAttr.Value, "true", StringComparison.OrdinalIgnoreCase)) isBackground = true;
                        }

                        // extract position offsets from anchor's positionH/positionV
                        if (isAnchor)
                        {
                            var posH = wpAncestor.Descendants().FirstOrDefault(d => d.LocalName == "positionH");
                            if (posH != null)
                            {
                                var posOff = posH.Descendants().FirstOrDefault(d => d.LocalName == "posOffset");
                                if (posOff != null && long.TryParse(posOff.InnerText, out var hOff)) offsetX = hOff;
                            }
                            var posV = wpAncestor.Descendants().FirstOrDefault(d => d.LocalName == "positionV");
                            if (posV != null)
                            {
                                var posOff = posV.Descendants().FirstOrDefault(d => d.LocalName == "posOffset");
                                if (posOff != null && long.TryParse(posOff.InnerText, out var vOff)) offsetY = vOff;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        OpenXmlHelpers.ImageLoadLogger?.Invoke($"GetImageInfosFromParagraph: failed reading extent/wrap attributes: {ex.Message}");
                    }
                }

                var bytes = OpenXmlHelpers.GetImageBytesFromWord(doc, rId);

                // Extract srcRect crop info from blipFill
                int cropL = 0, cropT = 0, cropR = 0, cropB = 0;
                try
                {
                    var srcRect = blip.Parent?.Descendants().FirstOrDefault(d => d.LocalName == "srcRect");
                    if (srcRect != null)
                    {
                        var attrs = srcRect.GetAttributes();
                        foreach (var a in attrs)
                        {
                            if (a.LocalName == "l" && int.TryParse(a.Value, out var v)) cropL = v;
                            else if (a.LocalName == "t" && int.TryParse(a.Value, out var v2)) cropT = v2;
                            else if (a.LocalName == "r" && int.TryParse(a.Value, out var v3)) cropR = v3;
                            else if (a.LocalName == "b" && int.TryParse(a.Value, out var v4)) cropB = v4;
                        }
                    }
                }
                catch { }

                if (bytes != null)
                {
                    // Extract hyperlink from a:hlinkClick in wp:docPr
                    string? hyperlinkUrl = null;
                    try
                    {
                        // docPr is a direct child of wp:inline or wp:anchor
                        var docPrParent = wpAncestor ?? blip.Ancestors().FirstOrDefault(a => a.LocalName == "inline" || a.LocalName == "anchor");
                        var docPr = docPrParent?.ChildElements
                            .FirstOrDefault(e => e.LocalName == "docPr");
                        var hlinkClick = docPr?.ChildElements
                            .FirstOrDefault(e => e.LocalName == "hlinkClick");
                        if (hlinkClick != null)
                        {
                            const string rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                            var hlinkRId = hlinkClick.GetAttribute("id", rNs).Value;
                            if (!string.IsNullOrEmpty(hlinkRId))
                            {
                                var rel = doc.MainDocumentPart?.HyperlinkRelationships
                                    .FirstOrDefault(r => r.Id == hlinkRId);
                                hyperlinkUrl = rel?.Uri?.ToString();
                            }
                        }
                    }
                    catch { }

                    results.Add(new ImageInfo(bytes, cx, cy, wrapText, isBackground, isAnchor, offsetX, offsetY, cropL, cropT, cropR, cropB, hyperlinkUrl));
                }
            }

            // VML legacy images and picts: try image data ids
            var vmlImages = p.Descendants<DocumentFormat.OpenXml.Vml.ImageData>();
            foreach (var idata in vmlImages)
            {
                string? rId = null;
                try
                {
                    var attr = idata.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, "id", StringComparison.OrdinalIgnoreCase));
                    if (attr != null && !string.IsNullOrEmpty(attr.Value)) rId = attr.Value;
                    if (string.IsNullOrEmpty(rId))
                    {
                        var href = idata.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, "href", StringComparison.OrdinalIgnoreCase));
                        if (href != null && !string.IsNullOrEmpty(href.Value)) rId = href.Value;
                    }
                }
                catch { }

                if (string.IsNullOrEmpty(rId))
                {
                    var a2 = idata.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, "id", StringComparison.OrdinalIgnoreCase));
                    if (a2 != null && !string.IsNullOrEmpty(a2.Value)) rId = a2.Value;
                }

                var bytes = OpenXmlHelpers.GetImageBytesFromWord(doc, rId);
                if (bytes != null) results.Add(new ImageInfo(bytes, null, null, null, false, false));
            }

            // Legacy <w:pict> <v:imagedata/> inside Picture elements
            var pics = p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>();
            foreach (var pic in pics)
            {
                var imgData = pic.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().FirstOrDefault();
                if (imgData == null) continue;
                string? rId = null;
                try
                {
                    var attr = imgData.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, "id", StringComparison.OrdinalIgnoreCase));
                    if (attr != null && !string.IsNullOrEmpty(attr.Value)) rId = attr.Value;
                }
                catch { }

                if (string.IsNullOrEmpty(rId))
                {
                    var a2 = imgData.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, "id", StringComparison.OrdinalIgnoreCase));
                    if (a2 != null && !string.IsNullOrEmpty(a2.Value)) rId = a2.Value;
                }
                if (string.IsNullOrEmpty(rId)) continue;
                var bytes = OpenXmlHelpers.GetImageBytesFromWord(doc, rId);
                if (bytes != null) results.Add(new ImageInfo(bytes, null, null, null, false, false));
            }

            return results;
        }

        internal static string SaveTempImage(byte[] bytes)
        {
            string ext = DetectImageExtension(bytes);
            string tmp = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"copilot_img_{Guid.NewGuid():N}{ext}");
            System.IO.File.WriteAllBytes(tmp, bytes);
            return tmp;
        }

        internal static void TryDeleteTempFile(string path)
        {
            try { if (System.IO.File.Exists(path)) System.IO.File.Delete(path); } catch { }
        }

        internal static string DetectImageExtension(byte[] bytes)
        {
            if (bytes.Length >= 4)
            {
                // PNG
                if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47) return ".png";
                // JPG
                if (bytes[0] == 0xFF && bytes[1] == 0xD8) return ".jpg";
                // GIF
                if (bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46) return ".gif";
            }
            return ".png";
        }

        /// <summary>
        /// Applies srcRect cropping (values in 1000ths of a percent) to raw image bytes.
        /// Returns cropped image bytes, or original if cropping fails or is not needed.
        /// Uses System.Drawing which is Windows-only; silently returns original on other platforms.
        /// </summary>
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        internal static byte[] ApplySrcRectCrop(byte[] imageBytes, int cropL, int cropT, int cropR, int cropB)
        {
            if (cropL == 0 && cropT == 0 && cropR == 0 && cropB == 0) return imageBytes;
            try
            {
                using var ms = new System.IO.MemoryStream(imageBytes);
                using var bmp = new System.Drawing.Bitmap(ms);
                int w = bmp.Width;
                int h = bmp.Height;
                int x = (int)(w * cropL / 100000.0);
                int y = (int)(h * cropT / 100000.0);
                int x2 = w - (int)(w * cropR / 100000.0);
                int y2 = h - (int)(h * cropB / 100000.0);
                int cw = Math.Max(1, x2 - x);
                int ch = Math.Max(1, y2 - y);
                var rect = new System.Drawing.Rectangle(x, y, cw, ch);
                using var cropped = bmp.Clone(rect, bmp.PixelFormat);
                using var outMs = new System.IO.MemoryStream();
                cropped.Save(outMs, System.Drawing.Imaging.ImageFormat.Png);
                return outMs.ToArray();
            }
            catch
            {
                return imageBytes;
            }
        }
    }
}
