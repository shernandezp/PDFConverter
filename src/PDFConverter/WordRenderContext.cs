using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PDFConverter.Layout;

namespace PDFConverter;

internal sealed class WordRenderContext
{
    public WordRenderContext(WordprocessingDocument document, ImageStore images)
    {
        Document = document;
        MainPart = document.MainDocumentPart
            ?? throw new InvalidOperationException("The document has no main part.");
        Images = images;
        Styles = WordStyleCache.For(MainPart);
        Numbering = new WordNumbering(Styles);
        DefaultStyle = ResolveDefaultStyle();
    }

    public WordprocessingDocument Document { get; }
    public MainDocumentPart MainPart { get; }
    public ImageStore Images { get; }
    public WordStyleCache Styles { get; }
    public WordNumbering Numbering { get; }
    public TextStyle DefaultStyle { get; }

    public List<WordImage> ImagesIn(OpenXmlElement scope) =>
        WordImageExtractor.Extract(WordImageExtractor.OwnerOf(scope, MainPart), scope);

    public ImageSource? MaterializeImage(WordImage image) =>
        Images.Get(image.Bytes, image.CropLeft, image.CropTop, image.CropRight, image.CropBottom);

    public Paragraph Prepare(Paragraph paragraph)
    {
        paragraph.DefaultStyle = DefaultStyle;
        return paragraph;
    }

    TextStyle ResolveDefaultStyle()
    {
        var defaults = Styles.DocDefaultsRunProperties;
        var family = WordHelpers.GetRunFontFamily(defaults) ?? Styles.ThemeFont ?? "Arial";
        var size = Units.HalfPointsToPoints(defaults?.FontSize?.Val?.Value) ?? LayoutDefaults.FontSizePoints;
        return new TextStyle(family, size);
    }
}
