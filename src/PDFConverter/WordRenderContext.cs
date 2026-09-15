using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace PDFConverter;

internal readonly record struct HyperlinkedImage(string Url, double WidthPoints, double HeightPoints);

internal sealed class WordRenderContext
{
    public WordRenderContext(WordprocessingDocument document, TempImageStore images)
    {
        Document = document;
        MainPart = document.MainDocumentPart
            ?? throw new InvalidOperationException("The document has no main part.");
        Images = images;
        Styles = WordStyleCache.For(MainPart);
        Numbering = new WordNumbering(Styles);
    }

    public WordprocessingDocument Document { get; }
    public MainDocumentPart MainPart { get; }
    public TempImageStore Images { get; }
    public WordStyleCache Styles { get; }
    public WordNumbering Numbering { get; }
    public HashSet<string> UsedFonts { get; } = new(StringComparer.OrdinalIgnoreCase);
    public List<HyperlinkedImage> HyperlinkedImages { get; } = [];

    public List<WordImage> ImagesIn(OpenXmlElement scope) =>
        WordImageExtractor.Extract(WordImageExtractor.OwnerOf(scope, MainPart), scope);

    public string? MaterializeImage(WordImage image) =>
        Images.Save(image.Bytes, image.CropLeft, image.CropTop, image.CropRight, image.CropBottom);

    public void TrackFont(string? fontFamily)
    {
        if (!string.IsNullOrEmpty(fontFamily)) UsedFonts.Add(fontFamily);
    }

    public void TrackHyperlinkedImage(string? url, double widthPoints, double heightPoints)
    {
        if (!string.IsNullOrEmpty(url)) HyperlinkedImages.Add(new HyperlinkedImage(url, widthPoints, heightPoints));
    }
}
