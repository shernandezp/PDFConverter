using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;

namespace PDFConverter;

internal sealed record WordImage(
    byte[] Bytes,
    long? ExtentCxEmu = null,
    long? ExtentCyEmu = null,
    string? WrapText = null,
    bool IsBackground = false,
    bool IsAnchor = false,
    long? OffsetXEmu = null,
    long? OffsetYEmu = null,
    int CropLeft = 0,
    int CropTop = 0,
    int CropRight = 0,
    int CropBottom = 0,
    string? HyperlinkUrl = null,
    string? HorizontalRelativeFrom = null,
    string? VerticalRelativeFrom = null,
    string? WrapKind = null)
{
    public double WidthPoints => ExtentCxEmu.HasValue
        ? Units.EmuToPoints(ExtentCxEmu.Value) : LayoutDefaults.ImageSizePoints;

    public double HeightPoints => ExtentCyEmu.HasValue
        ? Units.EmuToPoints(ExtentCyEmu.Value) : LayoutDefaults.ImageSizePoints;

    public bool IsPositionedAnchor => IsAnchor && OffsetXEmu.HasValue;
}

internal static class WordImageExtractor
{
    const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // Relationship ids are part-scoped: the same id names different images in the body, a header
    // and a footer, so the owning part has to come from the caller rather than be guessed.
    public static List<WordImage> Extract(OpenXmlPart? owner, OpenXmlElement? scope)
    {
        var images = new List<WordImage>();
        if (owner == null || scope == null) return images;

        var foundDrawingMl = false;
        foreach (var blip in scope.Descendants<A.Blip>())
        {
            foundDrawingMl = true;

            var relationshipId = blip.Embed?.Value ?? blip.Link?.Value;
            if (string.IsNullOrEmpty(relationshipId)) continue;

            var bytes = WordHelpers.GetImageBytes(owner, relationshipId);
            if (bytes == null) continue;

            var image = ReadDrawing(owner, blip, bytes);
            if (image != null) images.Add(image);
        }

        if (foundDrawingMl) return images;

        foreach (var imageData in scope.Descendants<V.ImageData>())
        {
            if (imageData.Ancestors().Any(a => a.LocalName == "group")) continue;
            var relationshipId = imageData.RelationshipId?.Value
                ?? Attribute(imageData, "id") ?? Attribute(imageData, "href");
            var bytes = WordHelpers.GetImageBytes(owner, relationshipId);
            if (bytes != null) images.Add(new WordImage(bytes));
        }

        return images;
    }

    public static OpenXmlPart OwnerOf(OpenXmlElement element, MainDocumentPart fallback)
    {
        var root = element as OpenXmlPartRootElement ?? element.Ancestors<OpenXmlPartRootElement>().FirstOrDefault();
        return root?.OpenXmlPart ?? fallback;
    }

    static WordImage? ReadDrawing(OpenXmlPart owner, A.Blip blip, byte[] bytes)
    {
        var frame = blip.Ancestors().FirstOrDefault(a => a.LocalName is "anchor" or "inline");
        if (frame == null) return new WordImage(bytes);

        var isAnchor = frame.LocalName == "anchor";

        var extent = frame.Descendants().FirstOrDefault(d => d.LocalName == "extent");
        var wrap = frame.Descendants().FirstOrDefault(d =>
            d.LocalName is "wrapSquare" or "wrapTopAndBottom" or "wrapNone"
                or "wrapTight" or "wrapThrough");
        var isBackground = Attribute(frame, "behindDoc") is "1" or "true";

        string? horizontalRelativeFrom = null, verticalRelativeFrom = null;
        long? offsetX = null, offsetY = null;
        if (isAnchor)
        {
            (horizontalRelativeFrom, offsetX) = ReadPosition(frame, "positionH");
            (verticalRelativeFrom, offsetY) = ReadPosition(frame, "positionV");
        }

        long? cx = ParseLong(Attribute(extent, "cx"));
        long? cy = ParseLong(Attribute(extent, "cy"));

        if (IsInsideGroupShape(blip))
        {
            var placement = ResolveGroupPlacement(blip);
            if (placement == null) return null;

            offsetX = (offsetX ?? 0) + placement.Value.X;
            offsetY = (offsetY ?? 0) + placement.Value.Y;
            cx = placement.Value.Cx;
            cy = placement.Value.Cy;
            isAnchor = true;
            isBackground = false;
        }

        var (cropLeft, cropTop, cropRight, cropBottom) = ReadCrop(blip);

        return new WordImage(bytes, cx, cy,
            Attribute(wrap, "wrapText"), isBackground, isAnchor, offsetX, offsetY,
            cropLeft, cropTop, cropRight, cropBottom,
            ReadHyperlink(owner, frame), horizontalRelativeFrom, verticalRelativeFrom, wrap?.LocalName);
    }

    // A shape inside a group is placed in the group's own coordinate space; each enclosing group
    // maps its child box (chOff/chExt) onto its own box (off/ext), so the transforms compose outward.
    static (long X, long Y, long Cx, long Cy)? ResolveGroupPlacement(A.Blip blip)
    {
        var shape = blip.Ancestors().FirstOrDefault(a => a.LocalName is "pic" or "wsp");
        var transform = shape?.Descendants().FirstOrDefault(d => d.LocalName == "xfrm");
        if (transform == null) return null;

        var (x, y, cx, cy) = Box(transform);
        if (cx <= 0 || cy <= 0) return null;

        foreach (var group in blip.Ancestors().Where(a => a.LocalName is "wgp" or "grpSp"))
        {
            var groupTransform = group.ChildElements
                .FirstOrDefault(c => c.LocalName == "grpSpPr")
                ?.ChildElements.FirstOrDefault(c => c.LocalName == "xfrm");
            if (groupTransform == null) continue;

            var (groupX, groupY, groupCx, groupCy) = Box(groupTransform);
            var child = Child(groupTransform);
            if (child.Cx <= 0 || child.Cy <= 0) continue;

            var scaleX = groupCx / child.Cx;
            var scaleY = groupCy / child.Cy;

            x = groupX + (x - child.X) * scaleX;
            y = groupY + (y - child.Y) * scaleY;
            cx *= scaleX;
            cy *= scaleY;
        }

        return ((long)x, (long)y, (long)cx, (long)cy);
    }

    static (double X, double Y, double Cx, double Cy) Box(OpenXmlElement transform) =>
        (Coordinate(transform, "off", "x"), Coordinate(transform, "off", "y"),
            Coordinate(transform, "ext", "cx"), Coordinate(transform, "ext", "cy"));

    static (double X, double Y, double Cx, double Cy) Child(OpenXmlElement transform) =>
        (Coordinate(transform, "chOff", "x"), Coordinate(transform, "chOff", "y"),
            Coordinate(transform, "chExt", "cx"), Coordinate(transform, "chExt", "cy"));

    static double Coordinate(OpenXmlElement transform, string localName, string attribute)
    {
        var element = transform.ChildElements.FirstOrDefault(c => c.LocalName == localName);
        return Units.TryParseDouble(Attribute(element, attribute), out var value) ? value : 0;
    }

    static (string? RelativeFrom, long? Offset) ReadPosition(OpenXmlElement frame, string localName)
    {
        var position = frame.Descendants().FirstOrDefault(d => d.LocalName == localName);
        if (position == null) return (null, null);
        var posOffset = position.Descendants().FirstOrDefault(d => d.LocalName == "posOffset");
        return (Attribute(position, "relativeFrom"), ParseLong(posOffset?.InnerText));
    }

    static (int Left, int Top, int Right, int Bottom) ReadCrop(A.Blip blip)
    {
        var srcRect = blip.Parent?.Descendants().FirstOrDefault(d => d.LocalName == "srcRect");
        if (srcRect == null) return (0, 0, 0, 0);
        return (ParseInt(Attribute(srcRect, "l")), ParseInt(Attribute(srcRect, "t")),
            ParseInt(Attribute(srcRect, "r")), ParseInt(Attribute(srcRect, "b")));
    }

    static string? ReadHyperlink(OpenXmlPart owner, OpenXmlElement frame)
    {
        var docPr = frame.ChildElements.FirstOrDefault(e => e.LocalName == "docPr");
        var hlinkClick = docPr?.ChildElements.FirstOrDefault(e => e.LocalName == "hlinkClick");
        if (hlinkClick == null) return null;

        var relationshipId = hlinkClick.GetAttributes()
            .FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri == RelationshipsNamespace).Value;
        if (string.IsNullOrEmpty(relationshipId)) return null;

        try
        {
            return owner.HyperlinkRelationships.FirstOrDefault(r => r.Id == relationshipId)?.Uri?.ToString();
        }
        catch
        {
            return null;
        }
    }

    static bool IsInsideGroupShape(OpenXmlElement element) =>
        element.Ancestors().Any(a => a.LocalName is "wgp" or "wGrp" or "grpSp");

    static string? Attribute(OpenXmlElement? element, string localName) =>
        element?.GetAttributes().FirstOrDefault(a =>
            string.Equals(a.LocalName, localName, StringComparison.OrdinalIgnoreCase)).Value;

    static long? ParseLong(string? text) => Units.TryParseLong(text, out var value) ? value : null;

    static int ParseInt(string? text) => Units.TryParseInt(text, out var value) ? value : 0;
}

internal sealed record WordLine(
    long OffsetXEmu,
    long OffsetYEmu,
    long LengthEmu,
    string? Color,
    string? HorizontalRelativeFrom,
    string? VerticalRelativeFrom)
{
    public double LeftPoints => Units.EmuToPoints(OffsetXEmu);
    public double TopPoints => Units.EmuToPoints(OffsetYEmu);
    public double LengthPoints => Units.EmuToPoints(LengthEmu);
}

internal static class WordLineExtractor
{
    // Word draws a signature rule as a "straight connector": a shape with line geometry and no
    // height. MigraDoc has no line primitive, so these are collected and drawn as frame borders.
    public static List<WordLine> Extract(OpenXmlElement? scope)
    {
        var lines = new List<WordLine>();
        if (scope == null) return lines;

        foreach (var shape in scope.Descendants().Where(d => d.LocalName == "wsp"))
        {
            var geometry = shape.Descendants().FirstOrDefault(d => d.LocalName == "prstGeom");
            if (Attribute(geometry, "prst") != "line") continue;

            var frame = shape.Ancestors().FirstOrDefault(a => a.LocalName == "anchor");
            if (frame == null) continue;

            var transform = shape.Descendants().FirstOrDefault(d => d.LocalName == "xfrm");
            var extent = transform?.ChildElements.FirstOrDefault(c => c.LocalName == "ext")
                ?? frame.Descendants().FirstOrDefault(d => d.LocalName == "extent");
            var length = ParseLong(Attribute(extent, "cx")) ?? 0;
            var height = ParseLong(Attribute(extent, "cy")) ?? 0;
            if (length <= 0 || Math.Abs(height) > length) continue;

            var (horizontalFrom, offsetX) = ReadPosition(frame, "positionH");
            var (verticalFrom, offsetY) = ReadPosition(frame, "positionV");

            lines.Add(new WordLine(offsetX ?? 0, offsetY ?? 0, length,
                ReadLineColor(shape), horizontalFrom, verticalFrom));
        }
        return lines;
    }

    static string? ReadLineColor(OpenXmlElement shape)
    {
        var outline = shape.Descendants().FirstOrDefault(d => d.LocalName == "ln");
        var rgb = outline?.Descendants().FirstOrDefault(d => d.LocalName == "srgbClr");
        return Attribute(rgb, "val");
    }

    static (string? RelativeFrom, long? Offset) ReadPosition(OpenXmlElement frame, string localName)
    {
        var position = frame.Descendants().FirstOrDefault(d => d.LocalName == localName);
        if (position == null) return (null, null);
        var posOffset = position.Descendants().FirstOrDefault(d => d.LocalName == "posOffset");
        return (Attribute(position, "relativeFrom"), ParseLong(posOffset?.InnerText));
    }

    static string? Attribute(OpenXmlElement? element, string localName) =>
        element?.GetAttributes().FirstOrDefault(a =>
            string.Equals(a.LocalName, localName, StringComparison.OrdinalIgnoreCase)).Value;

    static long? ParseLong(string? text) => Units.TryParseLong(text, out var value) ? value : null;
}
