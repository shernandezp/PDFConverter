using DocumentFormat.OpenXml;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
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
        // one-row table is the closest MigraDoc equivalent.
        var boxes = pictures
            .Select(picture => (Picture: picture, Style: ReadShapeStyle(picture)))
            .OrderBy(box => box.Style.MarginLeft)
            .ToList();

        var contentWidth = section.PageSetup.PageWidth.Point
            - section.PageSetup.LeftMargin.Point - section.PageSetup.RightMargin.Point;

        var table = section.AddTable();
        table.Borders.Visible = false;
        foreach (var box in boxes)
            table.AddColumn(Unit.FromPoint(Math.Min(box.Style.Width, contentWidth / boxes.Count)));

        var row = table.AddRow();
        for (var i = 0; i < boxes.Count; i++)
        {
            row[i].VerticalAlignment = VerticalAlignment.Top;
            RenderIntoCell(ctx, boxes[i].Picture, row[i]);
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
                    var target = section.AddParagraph();
                    ApplyBoxParagraphFormat(target, source, fill);
                    if (!WordContentRenderer.Render(ctx, source, source, target)) target.AddText(" ");
                }
                else if (child is W.Table table)
                {
                    WordTableRenderer.RenderTable(ctx, section, table);
                }
            }
        }
    }

    static void RenderIntoCell(WordRenderContext ctx, W.Picture picture, Cell cell)
    {
        foreach (var (textBox, fill) in TextBoxes(picture))
        {
            foreach (var child in textBox.ChildElements)
            {
                if (child is W.Paragraph source)
                {
                    var target = cell.AddParagraph();
                    target.Format.SpaceBefore = 0;
                    target.Format.SpaceAfter = 0;
                    target.Format.Alignment = WordHelpers.GetParagraphFormatting(source.ParagraphProperties).Alignment;
                    if (ColorUtils.TryParse(fill, out var color)) target.Format.Shading.Color = color;
                    if (!WordContentRenderer.Render(ctx, source, source, target)) target.AddText(" ");
                }
                else if (child is W.Table table)
                {
                    WordTableRenderer.RenderNestedTable(ctx, table, cell);
                }
            }
        }
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
        target.Format.Alignment = format.Alignment;
        target.Format.SpaceBefore = Unit.FromPoint(format.HasExplicitSpacingBefore ? format.SpacingBefore : 0);
        target.Format.SpaceAfter = Unit.FromPoint(format.HasExplicitSpacingAfter ? format.SpacingAfter : 0);

        if (format.LineSpacing is { } spacing && spacing > 0)
        {
            if (string.Equals(format.LineRule, "Exact", StringComparison.OrdinalIgnoreCase))
            {
                target.Format.LineSpacingRule = LineSpacingRule.Exactly;
                target.Format.LineSpacing = Unit.FromPoint(spacing);
            }
            else
            {
                target.Format.LineSpacingRule = LineSpacingRule.Multiple;
                target.Format.LineSpacing = spacing;
            }
        }

        if (ColorUtils.TryParse(fillColor, out var color)) target.Format.Shading.Color = color;
    }
}
