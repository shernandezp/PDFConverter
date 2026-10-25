using DocumentFormat.OpenXml;
using PDFConverter.Layout;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PDFConverter;

internal static class VmlTextBoxRenderer
{
    public static void Render(WordRenderContext ctx, IReadOnlyList<W.Picture> pictures, Section section)
    {
        if (pictures.Count == 0) return;
        if (pictures.Count == 1)
        {
            RenderIntoSection(ctx, pictures[0], section);
            return;
        }

        // Several shapes in one paragraph are positioned side by side in Word; a borderless
        // one-row table is the closest equivalent.
        var boxes = pictures
            .Select(picture => (Picture: picture, Style: ReadShapeStyle(picture)))
            .OrderBy(box => box.Style.MarginLeft)
            .ToList();

        var contentWidth = section.ContentWidth;
        var table = section.AddTable();
        foreach (var box in boxes)
            table.ColumnWidths.Add(Math.Min(box.Style.Width, contentWidth / boxes.Count));

        var row = table.AddRow();
        for (var i = 0; i < boxes.Count; i++)
        {
            var cell = row.Cell(i);
            cell.VerticalAlignment = VerticalAlignment.Top;
            RenderIntoCell(ctx, boxes[i].Picture, cell, table.ColumnWidths[i]);
        }
    }

    static (double MarginLeft, double Width) ReadShapeStyle(W.Picture picture)
    {
        var shape = picture.ChildElements.FirstOrDefault(c => c.LocalName is "group" or "shape");
        var style = shape?.GetAttributes().FirstOrDefault(a => a.LocalName == "style").Value;
        double marginLeft = 0, width = 200;
        if (string.IsNullOrEmpty(style)) return (marginLeft, width);

        foreach (var declaration in style.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = declaration.Split(':', 2);
            if (parts.Length != 2) continue;
            if (!Units.TryParseDouble(parts[1].Trim().Replace("pt", string.Empty), out var value)) continue;

            if (parts[0].Trim() == "margin-left") marginLeft = value;
            else if (parts[0].Trim() == "width") width = value;
        }
        return (marginLeft, width);
    }

    static void RenderIntoSection(WordRenderContext ctx, W.Picture picture, Section section)
    {
        foreach (var (textBox, fill) in TextBoxes(picture))
        {
            foreach (var child in textBox.ChildElements)
            {
                if (child is W.Paragraph source)
                {
                    var target = ctx.Prepare(section.AddParagraph());
                    ApplyBoxParagraphFormat(target, source, fill);
                    AddListLabel(ctx, source, target);
                    WordContentRenderer.Render(ctx, source, source, target);
                }
                else if (child is W.Table table)
                {
                    WordTableRenderer.RenderTable(ctx, section, table);
                }
            }
        }
    }

    static void RenderIntoCell(WordRenderContext ctx, W.Picture picture, TableCell cell, double cellWidth)
    {
        foreach (var (textBox, fill) in TextBoxes(picture))
        {
            foreach (var child in textBox.ChildElements)
            {
                if (child is W.Paragraph source)
                {
                    var target = ctx.Prepare(cell.AddParagraph());
                    target.Alignment = WordHelpers.GetParagraphFormatting(source.ParagraphProperties).Alignment;
                    if (ColorUtils.TryParse(fill, out var color)) target.Shading = color;
                    AddListLabel(ctx, source, target);
                    WordContentRenderer.Render(ctx, source, source, target);
                }
                else if (child is W.Table table)
                {
                    WordTableRenderer.RenderNestedTable(ctx, table, cell, cellWidth);
                }
            }
        }
    }

    static void AddListLabel(WordRenderContext ctx, W.Paragraph source, Paragraph target)
    {
        var label = WordDocumentBuilder.ListLabel(ctx, source);
        if (label != null) target.Add(new TextRun(label, target.DefaultStyle));
    }

    static IEnumerable<(OpenXmlElement TextBox, string? FillColor)> TextBoxes(W.Picture picture)
    {
        foreach (var textBox in picture.Descendants().Where(d => d.LocalName == "txbxContent"))
        {
            var shape = textBox.Ancestors().FirstOrDefault(a => a.LocalName == "shape");
            var fill = shape?.GetAttributes().FirstOrDefault(a => a.LocalName == "fillcolor").Value;
            yield return (textBox, fill);
        }
    }

    static void ApplyBoxParagraphFormat(Paragraph target, W.Paragraph source, string? fillColor)
    {
        var format = WordHelpers.GetParagraphFormatting(source.ParagraphProperties);
        target.Alignment = format.Alignment;
        if (format.LeftIndent > 0) target.LeftIndent = format.LeftIndent;
        if (format.RightIndent > 0) target.RightIndent = format.RightIndent;
        if (format.FirstLineIndent != 0) target.FirstLineIndent = format.FirstLineIndent;
        target.SpaceBefore = format.HasExplicitSpacingBefore ? format.SpacingBefore : 0;
        target.SpaceAfter = format.HasExplicitSpacingAfter ? format.SpacingAfter : 0;

        if (format.LineSpacing is { } spacing && spacing > 0)
        {
            target.LineSpacing = spacing;
            target.LineSpacingRule = string.Equals(format.LineRule, "Exact", StringComparison.OrdinalIgnoreCase)
                ? LineSpacingRule.Exactly
                : LineSpacingRule.Multiple;
        }

        if (ColorUtils.TryParse(fillColor, out var color)) target.Shading = color;
    }
}
